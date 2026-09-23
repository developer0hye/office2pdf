//! Excel positive-axis backgrounds in the original table fill layer.
//!
//! Expanding completed cell rectangles keeps automatic row sizes, merges, and
//! page fragments intact. Backgrounds remain below borders and cell content.
//!
//! Excel also clips each printed page's fills to that page's grid region
//! inset by one sheet point on the top and left edges (issue #1605): the
//! public #982 workbook's native trace clips the unscaled A1:B8 sheet's fills
//! to `[51, 55, 474, 418]` around raw rectangles starting at x=50/y=54, the
//! fitted 0.82 explicit-area sheet at `[55.76, 54.12]` for a grid origin of
//! `[54.94, 53.30]`, and that sheet's unscaled continuation page at x=473 for
//! a raw fill starting at x=472. The bottom and right edges keep the bleed,
//! and interior rows and columns keep their raw origin, so the clip only
//! trims the page's first row and first column. Text is unaffected: Excel
//! already seats cell content inside `[B+1, B_next]`.

use std::collections::HashMap;
use typst::introspection::Tag;
use typst::layout::{Abs, Frame, FrameItem, Point};
use typst::model::TableElem;
use typst::syntax::Span;
use typst::visualize::Geometry;
use typst_layout::Page;

/// The label codegen gives the worksheet grid's table, with the printed-to-
/// declared scale appended. It identifies the one table whose fills Excel
/// clips this way, and the one whose text paces on Excel's whole-point
/// advance grid (issue #1659).
pub(super) const SHEET_TABLE_LABEL_PREFIX: &str = "o2p-excel-fill-";

pub(super) fn adjust_cell_fills(pages: &mut [Page]) {
    let mut tables = HashMap::new();
    for page in pages.iter() {
        collect_tables(&page.frame, &mut tables);
    }
    if tables.is_empty() {
        return;
    }
    tracing::debug!(
        tables = tables.len(),
        "applying Excel cell background extents"
    );
    for page in pages.iter_mut() {
        adjust_frame(&mut page.frame, &tables);
    }
}

fn collect_tables(frame: &Frame, tables: &mut HashMap<Span, f64>) {
    for (_, item) in frame.items() {
        match item {
            FrameItem::Group(group) => collect_tables(&group.frame, tables),
            FrameItem::Tag(Tag::Start(content, _))
                if content.to_packed::<TableElem>().is_some() =>
            {
                let Some(label) = content.label() else {
                    continue;
                };
                let text = label.resolve();
                if !text.starts_with(SHEET_TABLE_LABEL_PREFIX) {
                    continue;
                }
                let Some(scale) = text.rsplit('-').next().and_then(|s| s.parse::<f64>().ok())
                else {
                    continue;
                };
                if scale.is_finite() && scale > 0.0 && scale <= 1.0 {
                    tables.insert(content.span(), scale);
                }
            }
            _ => {}
        }
    }
}

fn adjust_frame(frame: &mut Frame, tables: &HashMap<Span, f64>) {
    let mut items: Vec<(Point, FrameItem)> = frame.items().cloned().collect();
    // Where each tracked table's grid starts in this frame. Typst usually
    // nests the grid in its own group, so its fills sit at the group's
    // origin, but it inlines a single-item frame into its parent, and then
    // the fills carry the parent's coordinates. The table's start tag is
    // placed at the grid origin in whichever frame holds the fills.
    let mut origins: HashMap<Span, Point> = HashMap::new();
    for (position, item) in items.iter() {
        if let FrameItem::Tag(Tag::Start(content, _)) = item
            && content.to_packed::<TableElem>().is_some()
            && tables.contains_key(&content.span())
        {
            origins.insert(content.span(), *position);
        }
    }
    let mut table_slots: HashMap<Span, Vec<usize>> = HashMap::new();
    for (index, (position, item)) in items.iter_mut().enumerate() {
        match item {
            FrameItem::Group(group) => adjust_frame(&mut group.frame, tables),
            FrameItem::Shape(shape, span) if shape.fill.is_some() && shape.stroke.is_none() => {
                let Some(scale) = tables.get(span) else {
                    continue;
                };
                let Geometry::Rect(size) = &mut shape.geometry else {
                    continue;
                };
                let sheet_point = Abs::pt(*scale);
                size.x += sheet_point;
                size.y += sheet_point;
                // The page clip trims the first column's left strip and the
                // first row's top strip: exactly the fills whose track starts
                // on the grid origin. A fill in an unfilled gutter's shadow
                // keeps its raw origin, as native does (#1605).
                let origin: Point = origins.get(span).copied().unwrap_or_default();
                if position.x.approx_eq(origin.x) {
                    position.x += sheet_point;
                    size.x -= sheet_point;
                }
                if position.y.approx_eq(origin.y) {
                    position.y += sheet_point;
                    size.y -= sheet_point;
                }
                table_slots.entry(*span).or_default().push(index);
            }
            _ => {}
        }
    }
    // Typst groups fills by column. Excel paints row-first, so a later row
    // owns the shared lower-left corner even beside a horizontal merge.
    for slots in table_slots.values() {
        let mut fills: Vec<_> = slots.iter().map(|i| items[*i].clone()).collect();
        fills.sort_by(|a, b| {
            a.0.y
                .to_pt()
                .total_cmp(&b.0.y.to_pt())
                .then_with(|| a.0.x.to_pt().total_cmp(&b.0.x.to_pt()))
        });
        for (index, fill) in slots.iter().zip(fills) {
            items[*index] = fill;
        }
    }
    frame.clear();
    frame.push_multiple(items);
}
