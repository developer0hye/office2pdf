//! Chart text a print tile's window hides, kept out of that tile's text layer
//! (issue #1736).
//!
//! Width pagination clips a chart that crosses a horizontal print tile to
//! each tile's window (issue #1598). The clip hides the runs outside it, but
//! the PDF still carries their glyphs under the clip, selectable and
//! searchable on a tile that does not show them. Excel's export writes each
//! tile only the chart text visible on it: on the public #982 workbook printed
//! unscaled, the first tile's category axis reads `Jan`…`Sep` and the second
//! tile's `Oct` `Nov` `Dec`, with no value-axis labels or legend entries on
//! the second.
//!
//! Codegen labels the window's clip box, so this pass removes, inside it, each
//! text run whose glyph ink lies wholly outside the window. A run the window
//! cuts keeps all its glyphs, as the clip keeps its visible part. Every other
//! clipped box keeps the text it hides.

use typst::layout::{Frame, FrameItem, Point, Size, Transform};
use typst::text::TextItem;
use typst_layout::Page;

/// The label codegen gives the clip box of a chart printed on a page-column
/// window.
pub(super) const SHEET_CHART_WINDOW_LABEL: &str = "o2p-sheet-chart-window";

pub(super) fn drop_text_outside_chart_windows(pages: &mut [Page]) {
    let mut dropped: usize = 0;
    for page in pages.iter_mut() {
        if contains_window(&page.frame) {
            dropped += drop_in_windows(&mut page.frame);
        }
    }
    if dropped > 0 {
        tracing::debug!(dropped, "removed chart text hidden by print-tile windows");
    }
}

fn is_window(item: &FrameItem) -> bool {
    matches!(
        item,
        FrameItem::Group(group)
            if group
                .label
                .is_some_and(|label| label.resolve().as_str() == SHEET_CHART_WINDOW_LABEL)
    )
}

/// Whether any window sits in `frame`, read before anything is copied: most
/// pages have none, and rebuilding their frames would cost a clone of every
/// item on them.
fn contains_window(frame: &Frame) -> bool {
    frame.items().any(|(_, item)| {
        is_window(item) || matches!(item, FrameItem::Group(group) if contains_window(&group.frame))
    })
}

/// Find each window under `frame` and drop the text it hides; returns the
/// number of runs dropped.
fn drop_in_windows(frame: &mut Frame) -> usize {
    let mut items: Vec<(Point, FrameItem)> = frame.items().cloned().collect();
    let mut dropped: usize = 0;
    for (_, item) in items.iter_mut() {
        let window: bool = is_window(item);
        let FrameItem::Group(group) = item else {
            continue;
        };
        // The clip box's frame is the window: Typst clips a box to its own
        // size, here the window's width by the page's height.
        dropped += if window {
            let size: Size = group.frame.size();
            drop_outside(&mut group.frame, Transform::identity(), size)
        } else if contains_window(&group.frame) {
            drop_in_windows(&mut group.frame)
        } else {
            0
        };
    }
    if dropped > 0 {
        frame.clear();
        frame.push_multiple(items);
    }
    dropped
}

/// Drop the runs under `frame` inked wholly outside the `window` rectangle
/// at the origin of the window's space; `transform` maps `frame` into it.
///
/// Transforms are accumulated in full because a fitted sheet scales the
/// chart inside the window (issue #1069).
fn drop_outside(frame: &mut Frame, transform: Transform, window: Size) -> usize {
    let mut items: Vec<(Point, FrameItem)> = frame.items().cloned().collect();
    let mut dropped: usize = 0;
    items.retain_mut(|(position, item)| {
        let at: Transform = transform.pre_concat(Transform::translate(position.x, position.y));
        match item {
            FrameItem::Group(group) => {
                dropped += drop_outside(&mut group.frame, at.pre_concat(group.transform), window);
                true
            }
            FrameItem::Text(text) if !inked_inside(text, at, window) => {
                dropped += 1;
                false
            }
            _ => true,
        }
    });
    if dropped > 0 {
        frame.clear();
        frame.push_multiple(items);
    }
    dropped
}

/// Whether the run's glyph ink, mapped by `transform`, overlaps the window.
/// A run with no ink, such as a space, is measured by its advance on the
/// baseline instead.
fn inked_inside(text: &TextItem, transform: Transform, window: Size) -> bool {
    let ink = text.bbox();
    let (min, max): (Point, Point) = if ink.min.x.to_pt().is_finite() {
        (ink.min, ink.max)
    } else {
        (Point::zero(), Point::with_x(text.width()))
    };
    let corners: [Point; 4] = [min, Point::new(max.x, min.y), Point::new(min.x, max.y), max]
        .map(|corner| corner.transform(transform));
    let span = |axis: fn(Point) -> f64| {
        corners
            .iter()
            .fold((f64::INFINITY, f64::NEG_INFINITY), |(low, high), corner| {
                (low.min(axis(*corner)), high.max(axis(*corner)))
            })
    };
    let (left, right) = span(|point| point.x.to_pt());
    let (top, bottom) = span(|point| point.y.to_pt());
    right > 0.0 && left < window.x.to_pt() && bottom >= 0.0 && top <= window.y.to_pt()
}
