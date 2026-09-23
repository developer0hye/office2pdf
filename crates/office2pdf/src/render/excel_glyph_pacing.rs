//! Excel's whole-point glyph pacing, applied to the completed sheet frames
//! (issue #1659).
//!
//! Excel for Mac advances every sheet glyph by its font advance rounded to a
//! whole sheet point. Codegen already reserves that rounding — see
//! `sheet_advance_grid_tracking_pt` — but `tracking` is uniform, so it can
//! only spread one average correction over the run's gaps; each glyph inside
//! a long string still drifts, up to 1.3pt by the end of a 34-glyph Malgun
//! Gothic 12pt line on the native export of
//! `tests/fixtures/xlsx/customers_overflow_strip.xlsx`.
//!
//! Typst offers no per-glyph advance in its source language, but the
//! completed frame carries one advance per glyph, so this pass seats each gap
//! on its own rounded advance.
//!
//! The reservation is what makes that safe. Codegen sets the tracking to
//! `(Σ rounded − Σ natural) / gaps` over exactly the advances that carry a
//! gap, so a run's placed width already equals the sum of its rounded
//! advances and redistributing them conserves it. Every geometry measured
//! from that width — the right-aligned trailing reserve of issue #1233, the
//! centred seat of issue #1600, wrapping, and the spill clip — is therefore
//! untouched, and only the glyph origins inside the run move.
//!
//! A run that carries no such reservation is left exactly as the layout
//! engine placed it. Codegen prices the correction on the face its own
//! metric lookup resolves, and that is not always the face the run was
//! shaped with: a host without the declared family substitutes one, and a
//! lookup that resolves nothing prices no correction at all. Re-seating
//! those gaps would move the run's last glyph off the width its frame was
//! already measured with, so the pass compares the two sums and declines
//! where they part.

use typst::foundations::Content;
use typst::introspection::Tag;
use typst::layout::{Abs, Em, Frame, FrameItem, Point};
use typst::model::TableElem;
use typst::text::TextItem;
use typst_layout::Page;

use super::excel_fill_paint::SHEET_TABLE_LABEL_PREFIX;

/// Excel's advance grid, in sheet points (issue #621's column model rounds on
/// the same one).
const SHEET_ADVANCE_GRID_PT: f64 = 1.0;

/// How far a glyph's placed advance may sit from `natural + tracking` and
/// still count as the uniform correction codegen emitted. A kerned pair or a
/// source-stated letter spacing breaks that uniformity, and such a run is
/// left as the layout engine placed it.
const UNIFORM_SPACING_TOLERANCE_PT: f64 = 1e-4;

/// How far the run's rounded gap advances may sum from the gaps the layout
/// engine placed and still count as the reservation this pass redistributes.
///
/// Codegen prices its `tracking` on the face its own metric lookup resolves,
/// which is the shaped face whenever the workbook's family is installed: the
/// two sums then agree to the last float, because `format_f64` round-trips
/// the correction exactly. They part where no correction was priced at all,
/// or where the host substituted a different face for the declared one, and
/// re-seating those gaps would move the run's last glyph off the width its
/// frame was already measured with — by up to 3.66pt on the Korean sheet of
/// `office2pdf_repository_workbook.xlsx`, where 856 of its 4,014 runs carry
/// no reservation.
const RESERVED_WIDTH_TOLERANCE_PT: f64 = 1e-7;

/// Seat every reserved sheet run on Excel's whole-point advance grid. A run
/// whose gaps carry no such reservation keeps the placement the layout engine
/// gave it.
pub(super) fn pace_sheet_glyphs_on_whole_points(pages: &mut [Page]) {
    for page in pages.iter_mut() {
        if !carries_sheet_table(&page.frame) {
            continue;
        }
        pace_frame(&mut page.frame, None);
    }
}

/// The print scale a sheet grid's table label states, or `None` for any
/// other table.
///
/// Pagination only ever scales a sheet *down* to fit — `fit_page_to_pages`
/// returns the page unchanged at 1.0 or above — so the label's `(0, 1)`
/// filter carries every scale a sheet can print at.
fn sheet_table_scale(content: &Content) -> Option<f64> {
    content.to_packed::<TableElem>()?;
    let label: String = content.label()?.resolve().to_string();
    let scale: f64 = label
        .strip_prefix(SHEET_TABLE_LABEL_PREFIX)?
        .rsplit('-')
        .next()?
        .parse()
        .ok()?;
    (scale.is_finite() && scale > 0.0).then_some(scale)
}

fn carries_sheet_table(frame: &Frame) -> bool {
    frame.items().any(|(_, item)| match item {
        FrameItem::Group(group) => carries_sheet_table(&group.frame),
        FrameItem::Tag(Tag::Start(content, _)) => sheet_table_scale(content).is_some(),
        _ => false,
    })
}

/// Re-pace every run inside a sheet grid. `inherited` is the scale of the
/// grid this frame already sits in, if any: Typst nests a cell's lines in
/// groups below the table's start tag, and may inline a single-item frame
/// into its parent, so the scope has to be carried down rather than looked
/// up per frame.
fn pace_frame(frame: &mut Frame, inherited: Option<f64>) {
    let mut items: Vec<(Point, FrameItem)> = frame.items().cloned().collect();
    // Sheet grids never nest, so one open table is all a frame can hold.
    let mut open: Option<typst::introspection::Location> = None;
    let mut active: Option<f64> = inherited;
    for (_, item) in &mut items {
        match item {
            FrameItem::Group(group) => pace_frame(&mut group.frame, active),
            FrameItem::Text(text) => {
                if let Some(scale) = active {
                    pace_run(text, scale);
                }
            }
            FrameItem::Tag(Tag::End(location, ..)) => {
                if open == Some(*location) {
                    open = None;
                    active = inherited;
                }
            }
            FrameItem::Tag(Tag::Start(content, _)) => {
                if let Some(scale) = sheet_table_scale(content) {
                    open = content.location();
                    active = Some(scale);
                }
            }
            _ => {}
        }
    }
    frame.clear();
    frame.push_multiple(items);
}

/// Seat `text`'s glyph origins on the grid without changing its width, or
/// leave the run exactly as placed when its gaps carry no reservation to
/// redistribute.
fn pace_run(text: &mut TextItem, scale: f64) {
    let size: Abs = text.size;
    if size <= Abs::zero() || text.glyphs.len() < 2 {
        return;
    }
    let mut natural_pt: Vec<f64> = Vec::with_capacity(text.glyphs.len());
    let mut placed_pt: Vec<f64> = Vec::with_capacity(text.glyphs.len());
    for glyph in &text.glyphs {
        // A vertical run advances down its line; its pacing is unmeasured.
        if glyph.y_advance != Em::zero() {
            return;
        }
        let Some(advance) = text.font.x_advance(glyph.id) else {
            return;
        };
        natural_pt.push(advance.at(size).to_pt());
        placed_pt.push(glyph.x_advance.at(size).to_pt());
    }

    // Typst drops the tracking after a shaped item's last glyph, and trims a
    // line's trailing space to nothing, so the last advance says nothing
    // about how the run is spaced and is read past here. What the gaps carry
    // must be one uniform correction: that is codegen's reservation, and a
    // run spaced by anything else — a kern pair, a source-stated letter
    // spacing — is left as the layout engine placed it.
    let (_, gap_natural) = natural_pt.split_last().expect("at least two glyphs");
    let (_, gap_placed) = placed_pt.split_last().expect("at least two glyphs");
    let tracking_pt: f64 = gap_placed[0] - gap_natural[0];
    let uniform: bool = gap_placed.iter().zip(gap_natural).all(|(placed, natural)| {
        ((placed - natural) - tracking_pt).abs() <= UNIFORM_SPACING_TOLERANCE_PT
    });
    if !uniform {
        return;
    }

    // Only the gaps are re-seated. The last advance places no glyph of this
    // run, and codegen's reservation deliberately stops short of it —
    // `sheet_trailing_advance_space_pt` carries that final rounding
    // separately, where a right-aligned line needs it (issue #1233). Leaving
    // it as placed is also what conserves a reserved run's width to the last
    // float: the tracking already made the placed gaps sum to the rounded
    // ones, so redistributing them changes nothing the line was measured
    // from.
    let paced_pt: Vec<f64> = gap_natural
        .iter()
        .map(|natural_pt: &f64| {
            round_half_up_to_grid(natural_pt / scale, SHEET_ADVANCE_GRID_PT) * scale
        })
        .collect();

    // Redistribution is free only where those rounded advances are what the
    // placed gaps already sum to. Where they part, the run was placed from a
    // width this grid never priced, so leave it as the layout engine placed
    // it rather than paying for the grid with the run's own geometry.
    let placed_sum_pt: f64 = gap_placed.iter().sum();
    let paced_sum_pt: f64 = paced_pt.iter().sum();
    if (paced_sum_pt - placed_sum_pt).abs() > RESERVED_WIDTH_TOLERANCE_PT {
        return;
    }

    for (glyph, advance_pt) in text.glyphs.iter_mut().zip(&paced_pt) {
        glyph.x_advance = Em::from_abs(Abs::pt(*advance_pt), size);
    }
}

/// `value` rounded to the nearest multiple of `grid`, halves away from zero —
/// the rule Excel's own metrics take (issue #621).
fn round_half_up_to_grid(value_pt: f64, grid_pt: f64) -> f64 {
    (value_pt / grid_pt).round() * grid_pt
}

// The tests read a completed frame through `pdf.rs`'s compile helpers, which
// are themselves native-only — the wasm build resolves neither the system
// font paths they search nor the readers built on them.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[path = "excel_glyph_pacing_tests.rs"]
mod tests;
