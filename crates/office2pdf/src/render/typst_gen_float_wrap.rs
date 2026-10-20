//! Body text beside a DOCX floating table (issue #2229) or text frame
//! (issue #2231).
//!
//! Word treats a `w:tblpPr` table as an exclusion region: a line of body text
//! that falls within the table's vertical extent is laid out in the room
//! beside it, and moves below the table when neither side leaves room. A
//! paragraph Word lifts into a `w:framePr` frame excludes text the same way.
//! Typst places the table with `#place`, which excludes nothing, and has no
//! text wrap of its own. Which lines meet the table is only known once the
//! page is laid out, so each root paragraph that follows a floating table asks
//! at layout time, through the `-top`/`-bottom` labels the placed table
//! carries, whether its first line falls within the table's extent, and
//! narrows or moves itself accordingly. The room on each side is known from
//! the frame alone, so that choice is made here.
//!
//! The rule follows docxide-pdf's `FloatZone` (`src/pdf/mod.rs` at
//! `adb8c2b8`), which checks a paragraph at its first line only.
//!
//! A line break with `w:clear` resumes the line after it below the tables it
//! clears (issue #2230). Typst cannot move one line of a paragraph, so the
//! paragraph is laid out as two at that break, and the part after it wraps
//! as though each table it clears left no room beside it.
//!
//! TODO(per-line wrap beside a float, #2355): Word narrows each line in a
//! float's extent, so a paragraph that starts beside the float returns to the
//! full column below its bottom edge; here the whole paragraph stays narrowed.
//! One that starts above a floating table runs over it, and one that starts
//! above a text frame moves below the frame (issue #2231).
//!
//! A frame Word keeps at a fixed place on its page and keeps text off the
//! sides of (`w:wrap` `notBeside` or `none`) leaves no room beside it: a
//! root paragraph that meets its band starts below it. A paragraph moved
//! below one band can meet another, so the bands are checked until none
//! moves it further (issue #2231).
//!
//! Word 2013 and later also keep the box of a group of empty framed
//! paragraphs at the top margin, though the paragraphs take no flow lines,
//! and start a full-width table on the box's page below it (issue #2231).
//! A table in the box's section asks at layout time, through the labels on
//! the box's edges, whether it starts within the box, which can be on its
//! page whether the group comes before or after it. The box of a group
//! that directly follows a table asks, through the labels on the table's
//! first and last pages, whether the table carries rows onto the box's page,
//! which start below the box too.

use std::fmt::Write as _;

use crate::ir::{
    Block, BreakClear, Document, FloatingTable, FrameAnchor, Page, Paragraph, Run, TextFrame,
};

use super::format_f64;

const FLOAT_WRAP_HELPER: &str = "o2p-float-wrap";
const FRAME_BOX_CLEARANCE_HELPER: &str = "o2p-frame-box-clearance";
const FRAME_BOX_CARRY_HELPER: &str = "o2p-frame-box-carry";
const FRAME_BOX_LABEL: &str = "o2p-frame-box";
const FRAME_BOX_TABLE_LABEL: &str = "o2p-frame-box-table";
const FRAME_BOX_CARRIED_LABEL: &str = "o2p-frame-box-carried";

/// The narrowest strip beside a floating table or text frame that body text
/// is laid out in.
///
/// FIXME(unprobed minimum wrap width): docxide-pdf's empirical
/// `min_wrap_w`, which its `FloatZone` applies to frames as well. Its Word
/// references at `adb8c2b8` only separate no room at all beside a table
/// (`fullwidth-float-empty-clear`) from 162pt (`empty-float-float-indent-0-0`)
/// and 201pt (`leading-table-clear-*`), and none probes a frame, so the
/// threshold itself still needs a Word probe.
const MIN_ROOM_BESIDE_PT: f64 = 72.0;

/// What a root paragraph whose first line meets a float does.
#[derive(Debug, Clone, Copy, PartialEq)]
enum FloatSide {
    /// Lies in the room right of the float, inset this far from the column's
    /// left edge.
    Right { inset_pt: f64 },
    /// Lies in the room left of the float, inset this far from the column's
    /// right edge.
    Left { inset_pt: f64 },
    /// Starts below the float.
    Below,
}

impl FloatSide {
    /// Where a root paragraph that meets a float of `extent` goes in a column
    /// `column_width_pt` wide.
    fn of(extent: FloatExtent, column_width_pt: f64) -> Self {
        let left_edge: f64 = extent.x_pt - extent.left_from_text_pt;
        let right_edge: f64 = extent.x_pt + extent.width_pt + extent.right_from_text_pt;
        let room_left: f64 = left_edge.min(column_width_pt);
        let room_right: f64 = column_width_pt - right_edge.max(0.0);
        // The wider side; the right one on a tie, a choice no Word reference
        // decides yet.
        if room_right >= room_left && room_right >= MIN_ROOM_BESIDE_PT {
            Self::Right {
                inset_pt: right_edge.max(0.0),
            }
        } else if room_left >= MIN_ROOM_BESIDE_PT {
            Self::Left {
                inset_pt: (column_width_pt - left_edge).max(0.0),
            }
        } else {
            Self::Below
        }
    }
}

/// Whether a float of `extent` leaves no room beside it for body text in a
/// column `column_width_pt` wide, so every line that meets it starts below
/// it.
pub(super) fn leaves_no_room_beside(extent: FloatExtent, column_width_pt: f64) -> bool {
    FloatSide::of(extent, column_width_pt) == FloatSide::Below
}

/// The horizontal extent of a float that body text keeps clear of, measured
/// from the left edge of the column it wraps in.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct FloatExtent {
    x_pt: f64,
    width_pt: f64,
    /// How far text keeps from the float's left edge.
    left_from_text_pt: f64,
    /// How far text keeps from the float's right edge.
    right_from_text_pt: f64,
}

impl FloatExtent {
    /// A floating table's extent, its frame already resolved against the
    /// column.
    pub(super) fn of_table(table: &FloatingTable) -> Self {
        Self {
            x_pt: table.frame.x.unwrap_or_default(),
            width_pt: table.table.column_widths.iter().sum(),
            left_from_text_pt: table.frame.left_from_text,
            right_from_text_pt: table.frame.right_from_text,
        }
    }

    /// A text frame's extent, its `x` already resolved against the column.
    pub(super) fn of_text_frame(frame: &TextFrame) -> Self {
        Self {
            x_pt: frame.x,
            width_pt: frame.width,
            left_from_text_pt: frame.horizontal_space,
            right_from_text_pt: frame.horizontal_space,
        }
    }
}

/// A floating table or text frame placed in the current flow, which the root
/// paragraphs after it wrap around.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct BodyFloat {
    id: usize,
    side: FloatSide,
    /// How much more flow the paragraphs after the float can advance before
    /// none of them can still be on its page.
    reach_pt: f64,
    /// Whether the float is a text frame. A text frame typically hangs into
    /// the paragraph after its own, which starts above it; that paragraph
    /// moves below the frame rather than run over it.
    text_frame: bool,
}

impl BodyFloat {
    /// The wrap a float of `extent` imposes in a column `column_width_pt`
    /// wide, on the paragraphs after it until they have advanced `reach_pt`.
    pub(super) fn new(
        id: usize,
        extent: FloatExtent,
        column_width_pt: f64,
        reach_pt: f64,
        text_frame: bool,
    ) -> Self {
        let side: FloatSide = FloatSide::of(extent, column_width_pt);
        tracing::debug!(
            id,
            float_width_pt = extent.width_pt,
            ?side,
            "resolved body text wrap beside a float"
        );
        Self {
            id,
            side,
            reach_pt,
            text_frame,
        }
    }

    /// The wrap of a frame that keeps text off its sides: a paragraph that
    /// meets it starts below it, until the paragraphs after it have advanced
    /// `reach_pt`.
    pub(super) fn below(id: usize, reach_pt: f64) -> Self {
        Self {
            id,
            side: FloatSide::Below,
            reach_pt,
            text_frame: true,
        }
    }

    pub(super) fn id(&self) -> usize {
        self.id
    }
}

/// Count a wrapped paragraph's least advance against each float's reach and
/// drop the floats no later paragraph can meet: a floating table once the
/// paragraphs have advanced a page, an exact text frame once they have
/// passed its bottom edge. Each wrapped paragraph is a layout-time query, so
/// a float at the top of a long section would otherwise cost one for every
/// paragraph after it.
pub(super) fn advance_past_floats(floats: &mut Vec<BodyFloat>, advance_pt: f64) {
    floats.retain_mut(|float| {
        float.reach_pt -= advance_pt.max(0.0);
        float.reach_pt > 0.0
    });
}

/// Whether any flow section places a body floating table, text-anchored text
/// frame or frame that keeps text off its sides at its root, which is when
/// the wrap helper has to be defined.
pub(super) fn document_has_body_float(doc: &Document) -> bool {
    doc.pages.iter().any(|page| match page {
        Page::Flow(flow) | Page::FlowContinuous(flow) => {
            flow.content.iter().any(|block| match block {
                Block::FloatingTable(_) => true,
                Block::TextFrame(frame) => {
                    frame.vertical_anchor == FrameAnchor::Text || frame.text_below
                }
                _ => false,
            })
        }
        _ => false,
    })
}

/// Define the layout-time wrap: `floats` are the floats in force, `line` the
/// paragraph's first-line height and `before` its own `w:before`.
///
/// A paragraph meets a float when its first line overlaps the float's
/// vertical extent on the same page. Beside it, the paragraph is padded off
/// the float's side so its indents and `w:jc` resolve against the room left;
/// below it, the paragraph restarts at the float's bottom edge plus its own
/// before-gap, as Word restarts a line that has no room. The move is a top
/// pad rather than `v()`: Typst drops spacing that opens a block's content.
/// The floats are checked again after a move, which can bring the first line
/// into a float checked before it or out of one, until a pass moves it no
/// further; the sides it lies beside are those of that last pass.
///
/// A paragraph whose first line ends above a text frame (`pushes`) but which
/// is tall enough to reach it also restarts below the frame. Its height is
/// measured at the width it is laid out in, which only a `layout` call knows,
/// so only a paragraph in reach of a text frame pays for one.
pub(super) fn write_float_wrap_helper(out: &mut String) {
    let _ = writeln!(
        out,
        "#let {FLOAT_WRAP_HELPER}(floats, line, before, body) = context {{ \
         let at = here().position(); \
         let wrap(width) = {{ \
         let push = 0pt; let left = 0pt; let right = 0pt; let height = none; \
         let moved = true; let passes = 0; \
         while moved and passes < floats.len() {{ \
         moved = false; passes += 1; let start = push; left = 0pt; right = 0pt; \
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
         else {{ left = calc.max(left, float.inset) }} }} \
         else if float.pushes and top.page == at.page and line-top < top.y {{ \
         if height == none {{ height = measure(body, width: width).height }}; \
         if line-top + height > top.y {{ push = calc.max(push, bottom.y + before - at.y) }} }} }} }}; \
         moved = push > start }}; \
         pad(top: push, left: left, right: right, body) }}; \
         if floats.any(float => float.pushes) {{ layout(size => wrap(size.width)) }} \
         else {{ wrap(auto) }} }}"
    );
}

/// Whether any flow section keeps an empty frame's box at the top margin,
/// which is when the table clearance helpers have to be defined.
pub(super) fn document_has_top_margin_frame_box(doc: &Document) -> bool {
    doc.pages.iter().any(|page| match page {
        Page::Flow(flow) | Page::FlowContinuous(flow) => has_top_margin_frame_box(&flow.content),
        _ => false,
    })
}

/// Whether a flow section's root `blocks` keep an empty frame's box at the
/// top margin.
pub(super) fn has_top_margin_frame_box(blocks: &[Block]) -> bool {
    blocks.iter().any(|block| {
        matches!(block, Block::TextFrame(frame) if frame.vertical_anchor != FrameAnchor::Text
            && !frame.text_below)
    })
}

/// Define the layout-time clearances a full-width table takes from the
/// top-margin frame boxes.
///
/// A table that starts within a box on its page starts at the box's bottom
/// edge instead. The clearance is a block of its own before the table, so the
/// gap above it stays as it was, and so do the rows the table carries on to
/// the next page.
///
/// The rows a table carries onto the page of the box of a group that directly
/// follows it start below the box as well, as in Word's export of
/// docxide-pdf's `empty-margin-frame-long-prefix`. Such a box is drawn by a
/// helper that floats it at the top of that page, which moves the carried
/// rows and everything after them down; nothing else on the page comes before
/// them. The helper floats the box only when the rows still end on that page
/// below it, measured from where they end without it: a float that moved them
/// on to the next page would move the group and its box there too, and
/// Typst's layout would never settle. It takes the bottom of the text area,
/// the box's offsets and its content. Typst does not settle either when the
/// float's height is read from the box's own edge labels, which the float
/// moves, so the helper measures the content instead.
///
/// TODO(rows carried past a box's page): when the box would push the carried
/// rows onto the next page, they keep their place here; no Word reference
/// shows that case.
pub(super) fn write_top_margin_frame_box_clearance_helper(out: &mut String) {
    let _ = writeln!(
        out,
        "#let {FRAME_BOX_CLEARANCE_HELPER}() = context {{ \
         let at = here().position(); \
         let push = 0pt; \
         for (top, bottom) in query(<{FRAME_BOX_LABEL}-top>).zip(query(<{FRAME_BOX_LABEL}-bottom>)) {{ \
         let top = top.location().position(); \
         let bottom = bottom.location().position(); \
         if top.page == at.page and at.y >= top.y and at.y < bottom.y {{ \
         push = calc.max(push, bottom.y - at.y) }} }}; \
         block(height: push, above: 0pt, below: 0pt) }}"
    );
    let _ = writeln!(
        out,
        "#let {FRAME_BOX_CARRY_HELPER}(text-bottom, dx, dy, body) = context {{ \
         let start = query(selector(<{FRAME_BOX_TABLE_LABEL}-start>).before(here())).last().location().position(); \
         let end = query(selector(<{FRAME_BOX_TABLE_LABEL}-end>).before(here())).last().location().position(); \
         let height = measure(body).height; \
         let carried = query(<{FRAME_BOX_CARRIED_LABEL}>).filter(carried => carried.location().page() == end.page).map(carried => carried.value).sum(default: 0pt); \
         if start.page < end.page and end.y - carried + height <= text-bottom {{ \
         place(top + left, float: true, clearance: 0pt, dx: dx, dy: dy, [#metadata(height) <{FRAME_BOX_CARRIED_LABEL}>#body]) }} \
         else {{ place(top + left, dx: dx, dy: dy, body) }} }}"
    );
}

/// Write the clearance a full-width table at the flow root takes from the
/// top-margin frame boxes, just before the table.
pub(super) fn write_top_margin_frame_box_clearance(out: &mut String) {
    let _ = writeln!(out, "#{FRAME_BOX_CLEARANCE_HELPER}()");
}

/// Open the box of a frame group that directly follows a full-width table at
/// the flow root, `x_pt` and `y_pt` from the top left of the text area,
/// through the helper that floats it above the rows the table carries onto
/// its page. `text_bottom_pt` is the bottom of the text area, from the top of
/// the page. The box's content follows; `]` closes it.
pub(super) fn write_top_margin_frame_box_carry_open(
    out: &mut String,
    text_bottom_pt: f64,
    x_pt: f64,
    y_pt: f64,
) {
    let _ = writeln!(
        out,
        "#{FRAME_BOX_CARRY_HELPER}({}pt, {}pt, {}pt)[",
        format_f64(text_bottom_pt),
        format_f64(x_pt),
        format_f64(y_pt)
    );
}

/// The label a top-margin frame box carries at its `edge`, `top` or
/// `bottom`.
pub(super) fn write_top_margin_frame_box_label(out: &mut String, edge: &str) {
    let _ = writeln!(out, "#metadata(none) <{FRAME_BOX_LABEL}-{edge}>");
}

/// The label a full-width table that a frame group directly follows carries
/// at its `edge`, `start` or `end`, inside its wrapper, so on its first or
/// last page.
pub(super) fn write_frame_box_table_label(out: &mut String, edge: &str) {
    let _ = writeln!(out, "#metadata(none) <{FRAME_BOX_TABLE_LABEL}-{edge}>");
}

/// The labels a placed floating table or text frame carries at its top and
/// bottom edges.
pub(super) fn write_float_top_label(out: &mut String, id: usize) {
    let _ = writeln!(out, "#metadata(none) <o2p-float-{id}-top>");
}

pub(super) fn write_float_bottom_label(out: &mut String, id: usize) {
    let _ = writeln!(out, "#metadata(none) <o2p-float-{id}-bottom>");
}

/// Whether a line break clearing `clear` resumes below a table whose text
/// lies on `side` of it: `left` clears a table on the line's left, which is
/// one with its room on the right.
///
/// FIXME(unprobed one-sided clear): the Word references at hand clear only
/// `all`; `left` and `right` follow ECMA-376's `ST_BrClear` wording.
fn clears(clear: BreakClear, side: FloatSide) -> bool {
    matches!(
        (clear, side),
        (BreakClear::All, _)
            | (_, FloatSide::Below)
            | (BreakClear::Left, FloatSide::Right { .. })
            | (BreakClear::Right, FloatSide::Left { .. })
    )
}

/// A paragraph split at its first clearing break: the part before the break,
/// the break's clear, and the part after it, or `None` when it has no
/// clearing break.
///
/// The break itself goes: the two parts are the lines on either side of it.
/// The paragraph's own spacing stays on its outer edges, and its first-line
/// indent and list number on the first part. A part with no text left keeps
/// the line it stands for, sized by the run the break is in.
///
/// TODO(bordered paragraph split at a clearing break): each part draws its
/// own `w:pBdr` box, where Word draws one around the paragraph.
pub(super) fn split_at_clearing_break(
    para: &Paragraph,
) -> Option<(Paragraph, BreakClear, Paragraph)> {
    let first = *para.style.clearing_breaks.first()?;
    let mut head_runs: Vec<Run> = Vec::new();
    let mut tail_runs: Vec<Run> = Vec::new();
    let mut break_run: Option<&Run> = None;
    let mut line_breaks_before: usize = 0;
    for run in &para.runs {
        if break_run.is_some() {
            tail_runs.push(run.clone());
            continue;
        }
        let line_breaks: usize = run.text.matches('\n').count();
        if line_breaks_before + line_breaks <= first.line_break {
            line_breaks_before += line_breaks;
            head_runs.push(run.clone());
            continue;
        }
        let (at, _) = run
            .text
            .match_indices('\n')
            .nth(first.line_break - line_breaks_before)?;
        for (part, runs) in [
            (&run.text[..at], &mut head_runs),
            (&run.text[at + 1..], &mut tail_runs),
        ] {
            if !part.is_empty() {
                runs.push(Run {
                    text: part.to_string(),
                    ..run.clone()
                });
            }
        }
        break_run = Some(run);
    }
    let break_run: &Run = break_run?;

    let mut head = Paragraph {
        style: para.style.clone(),
        runs: head_runs,
    };
    head.style.space_after = Some(0.0);
    head.style.clearing_breaks = Box::default();
    let mut tail = Paragraph {
        style: para.style.clone(),
        runs: tail_runs,
    };
    tail.style.space_before = Some(0.0);
    tail.style.decoration_before_spacing = None;
    tail.style.indent_first_line = None;
    tail.style.has_inline_list_number = false;
    tail.style.numbering_suffix = None;
    tail.style.starts_after_layout_break = false;
    tail.style.clearing_breaks = para.style.clearing_breaks[1..]
        .iter()
        .map(|clearing| crate::ir::ClearingBreak {
            line_break: clearing.line_break - first.line_break - 1,
            ..*clearing
        })
        .collect();
    for part in [&mut head, &mut tail] {
        if part.runs.is_empty() && part.style.paragraph_mark_text_style.is_none() {
            part.style.paragraph_mark_text_style = Some(Box::new(break_run.style.clone()));
        }
    }
    Some((head, first.clear, tail))
}

/// Open a root paragraph's wrap around `floats`; the caller closes it with
/// `]`. `first_line_pt` is the paragraph's first line height and
/// `space_before_pt` its own before-gap. A paragraph that follows a clearing
/// break of `clear` moves below the tables that break clears.
pub(super) fn write_float_wrap_open(
    out: &mut String,
    floats: &[BodyFloat],
    clear: Option<BreakClear>,
    first_line_pt: f64,
    space_before_pt: f64,
) {
    let _ = write!(out, "#{FLOAT_WRAP_HELPER}((");
    for float in floats {
        let side: FloatSide = match clear {
            Some(clear) if clears(clear, float.side) => FloatSide::Below,
            _ => float.side,
        };
        let (side, inset_pt): (&str, f64) = match side {
            FloatSide::Right { inset_pt } => ("right", inset_pt),
            FloatSide::Left { inset_pt } => ("left", inset_pt),
            FloatSide::Below => ("below", 0.0),
        };
        let _ = write!(
            out,
            "(id: {}, side: \"{side}\", inset: {}pt, pushes: {}),",
            float.id,
            format_f64(inset_pt),
            float.text_frame
        );
    }
    let _ = writeln!(
        out,
        "), {}pt, {}pt)[",
        format_f64(first_line_pt),
        format_f64(space_before_pt)
    );
}
