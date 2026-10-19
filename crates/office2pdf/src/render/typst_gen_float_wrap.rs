//! Body text beside a DOCX floating table (issue #2229).
//!
//! Word treats a `w:tblpPr` table as an exclusion region: a line of body text
//! that falls within the table's vertical extent is laid out in the room
//! beside it, and moves below the table when neither side leaves room. Typst
//! places the table with `#place`, which excludes nothing, and has no text
//! wrap of its own. Which lines meet the table is only known once the page is
//! laid out, so each root paragraph that follows a floating table asks at
//! layout time, through the `-top`/`-bottom` labels the placed table carries,
//! whether its first line falls within the table's extent, and narrows or
//! moves itself accordingly. The room on each side is known from the frame
//! alone, so that choice is made here.
//!
//! The rule follows docxide-pdf's `FloatZone` (`src/pdf/mod.rs` at
//! `adb8c2b8`), which checks a paragraph at its first line only.
//!
//! TODO(per-line wrap beside a floating table): Word narrows each line in the
//! table's extent, so a paragraph that starts beside the table returns to the
//! full column below its bottom edge; here the whole paragraph stays narrowed,
//! and one that starts above the table runs over it.

use std::fmt::Write as _;

use crate::ir::{Block, Document, FloatingTable, Page};

use super::format_f64;

const FLOAT_WRAP_HELPER: &str = "o2p-float-wrap";

/// The narrowest strip beside a floating table that body text is laid out in.
///
/// FIXME(unprobed minimum wrap width): docxide-pdf's empirical
/// `min_wrap_w`. Its Word references at `adb8c2b8` only separate no room at
/// all (`fullwidth-float-empty-clear`) from 162pt
/// (`empty-float-float-indent-0-0`) and 201pt (`leading-table-clear-*`), so
/// the threshold itself still needs a Word probe.
const MIN_ROOM_BESIDE_PT: f64 = 72.0;

/// What a root paragraph whose first line meets a floating table does.
#[derive(Debug, Clone, Copy, PartialEq)]
enum FloatSide {
    /// Lies in the room right of the table, inset this far from the column's
    /// left edge.
    Right { inset_pt: f64 },
    /// Lies in the room left of the table, inset this far from the column's
    /// right edge.
    Left { inset_pt: f64 },
    /// Starts below the table.
    Below,
}

/// A floating table placed in the current flow, which the root paragraphs
/// after it wrap around.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct BodyFloat {
    id: usize,
    side: FloatSide,
    /// How much more flow the paragraphs after the table can advance before
    /// none of them can still be on its page.
    reach_pt: f64,
}

impl BodyFloat {
    /// The wrap a floating table imposes in a column `column_width_pt` wide
    /// and `column_height_pt` tall, with the table's frame already resolved
    /// against that column.
    pub(super) fn new(
        id: usize,
        table: &FloatingTable,
        column_width_pt: f64,
        column_height_pt: f64,
    ) -> Self {
        let frame = &table.frame;
        let x: f64 = frame.x.unwrap_or_default();
        let width: f64 = table.table.column_widths.iter().sum();
        let left_edge: f64 = x - frame.left_from_text;
        let right_edge: f64 = x + width + frame.right_from_text;
        let room_left: f64 = left_edge.min(column_width_pt);
        let room_right: f64 = column_width_pt - right_edge.max(0.0);
        // The wider side; the right one on a tie, a choice no Word reference
        // decides yet.
        let side: FloatSide = if room_right >= room_left && room_right >= MIN_ROOM_BESIDE_PT {
            FloatSide::Right {
                inset_pt: right_edge.max(0.0),
            }
        } else if room_left >= MIN_ROOM_BESIDE_PT {
            FloatSide::Left {
                inset_pt: (column_width_pt - left_edge).max(0.0),
            }
        } else {
            FloatSide::Below
        };
        tracing::debug!(
            id,
            table_width_pt = width,
            room_left_pt = room_left,
            room_right_pt = room_right,
            ?side,
            "resolved body text wrap beside a floating table"
        );
        Self {
            id,
            side,
            reach_pt: column_height_pt,
        }
    }

    pub(super) fn id(&self) -> usize {
        self.id
    }
}

/// Count a wrapped paragraph's least advance against each table's reach and
/// drop the tables no later paragraph can share a page with. Each wrapped
/// paragraph is a layout-time query, so a table at the top of a long section
/// would otherwise cost one for every paragraph after it.
pub(super) fn advance_past_floats(floats: &mut Vec<BodyFloat>, advance_pt: f64) {
    floats.retain_mut(|float| {
        float.reach_pt -= advance_pt.max(0.0);
        float.reach_pt > 0.0
    });
}

/// Whether any flow section places a body floating table at its root, which
/// is when the wrap helper has to be defined.
pub(super) fn document_has_body_floating_table(doc: &Document) -> bool {
    doc.pages.iter().any(|page| match page {
        Page::Flow(flow) | Page::FlowContinuous(flow) => flow
            .content
            .iter()
            .any(|block| matches!(block, Block::FloatingTable(_))),
        _ => false,
    })
}

/// Define the layout-time wrap: `floats` are the tables in force, `line` the
/// paragraph's first-line height and `before` its own `w:before`.
///
/// A paragraph meets a table when its first line overlaps the table's
/// vertical extent on the same page. Beside it, the paragraph is padded off
/// the table's side so its indents and `w:jc` resolve against the room left;
/// below it, the paragraph restarts at the table's bottom edge plus its own
/// before-gap, as Word restarts a line that has no room. The move is a top
/// pad rather than `v()`: Typst drops spacing that opens a block's content.
pub(super) fn write_float_wrap_helper(out: &mut String) {
    let _ = writeln!(
        out,
        "#let {FLOAT_WRAP_HELPER}(floats, line, before, body) = context {{ \
         let at = here().position(); \
         let push = 0pt; let left = 0pt; let right = 0pt; \
         for float in floats {{ \
         let tops = query(label(\"o2p-float-\" + str(float.id) + \"-top\")); \
         let bottoms = query(label(\"o2p-float-\" + str(float.id) + \"-bottom\")); \
         if tops.len() == 1 and bottoms.len() == 1 {{ \
         let top = tops.first().location().position(); \
         let bottom = bottoms.first().location().position(); \
         let line-top = at.y + push; \
         if top.page == at.page and line-top < bottom.y and line-top + line > top.y {{ \
         if float.side == \"below\" {{ push = calc.max(push, bottom.y + before - at.y) }} \
         else if float.side == \"left\" {{ right = calc.max(right, float.inset) }} \
         else {{ left = calc.max(left, float.inset) }} }} }} }}; \
         pad(top: push, left: left, right: right, body) }}"
    );
}

/// The labels a placed floating table carries at its top and bottom edges.
pub(super) fn write_float_top_label(out: &mut String, id: usize) {
    let _ = writeln!(out, "#metadata(none) <o2p-float-{id}-top>");
}

pub(super) fn write_float_bottom_label(out: &mut String, id: usize) {
    let _ = writeln!(out, "#metadata(none) <o2p-float-{id}-bottom>");
}

/// Open a root paragraph's wrap around `floats`; the caller closes it with
/// `]`. `first_line_pt` is the paragraph's first line height and
/// `space_before_pt` its own before-gap.
pub(super) fn write_float_wrap_open(
    out: &mut String,
    floats: &[BodyFloat],
    first_line_pt: f64,
    space_before_pt: f64,
) {
    let _ = write!(out, "#{FLOAT_WRAP_HELPER}((");
    for float in floats {
        let (side, inset_pt): (&str, f64) = match float.side {
            FloatSide::Right { inset_pt } => ("right", inset_pt),
            FloatSide::Left { inset_pt } => ("left", inset_pt),
            FloatSide::Below => ("below", 0.0),
        };
        let _ = write!(
            out,
            "(id: {}, side: \"{side}\", inset: {}pt),",
            float.id,
            format_f64(inset_pt)
        );
    }
    let _ = writeln!(
        out,
        "), {}pt, {}pt)[",
        format_f64(first_line_pt),
        format_f64(space_before_pt)
    );
}
