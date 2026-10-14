//! Word automatic vertical borders sized from native zero-width table lines.
//!
//! Replace private line geometry with the resolved Word border paint without
//! changing cell fills, tagged models, text layout, or shared glyph/image data.

use std::collections::{HashMap, HashSet};
#[cfg(not(target_arch = "wasm32"))]
use std::time::Instant;
use typst::foundations::{FromValue, Resolve, StyleChain, Value};
use typst::introspection::{MetadataElem, Tag};
use typst::layout::{Abs, Frame, FrameItem, Point, Size};
use typst::model::TableElem;
use typst::syntax::Span;
use typst::visualize::{Color, FixedStroke, Geometry, Paint, Stroke};
use typst_layout::Page;
#[cfg(target_arch = "wasm32")]
use web_time::Instant;

pub(super) const MARKER: &str = "office2pdf-word-auto-cell-borders";
pub(super) const TABLE_LABEL_PREFIX: &str = "o2p-word-auto-borders-";

#[derive(Clone, Copy, PartialEq)]
enum Mode {
    Solid,
    Double,
    Patterned,
}

struct Edge {
    mode: Mode,
    stroke: FixedStroke,
}

impl Edge {
    fn read(value: &Value) -> Option<Self> {
        let Value::Array(fields) = value else {
            return None;
        };
        let [Value::Str(mode), stroke] = fields.as_slice() else {
            return None;
        };
        let mode = match mode.as_str() {
            "solid" => Mode::Solid,
            "double" => Mode::Double,
            "patterned" => Mode::Patterned,
            _ => return None,
        };
        let stroke = Stroke::from_value(stroke.clone())
            .ok()?
            .resolve(StyleChain::default())
            .unwrap_or_default();
        (stroke.thickness.to_pt().is_finite() && stroke.thickness > Abs::zero())
            .then(|| Self { mode, stroke })
    }

    fn paints(&self, size: Size, x: Abs) -> Vec<(Point, FrameItem)> {
        let width = self.stroke.thickness;
        let count = if self.mode == Mode::Double { 2 } else { 1 };
        (0..count)
            .map(|index| {
                let x = x + width * (2.0 * f64::from(index));
                let (position, shape) = if self.mode == Mode::Patterned {
                    (
                        Point::with_x(x + width / 2.0),
                        Geometry::Line(Point::with_y(size.y)).stroked(self.stroke.clone()),
                    )
                } else {
                    (
                        Point::with_x(x),
                        Geometry::Rect(Size::new(width, size.y)).filled(self.stroke.paint.clone()),
                    )
                };
                (position, FrameItem::Shape(shape, Span::detached()))
            })
            .collect()
    }
}

#[derive(Default)]
struct Tables {
    spans: HashSet<Span>,
    edges: HashMap<Color, Edge>,
}

pub(super) fn paint_auto_cell_edges(pages: &mut [Page]) {
    let started = Instant::now();
    let mut tables = Tables::default();
    for page in pages.iter() {
        collect(&page.frame, &mut tables);
    }
    if tables.edges.is_empty() || tables.spans.is_empty() {
        return;
    }
    for page in pages.iter_mut() {
        paint_frame(&mut page.frame, &tables);
    }
    tracing::debug!(
        tables = tables.spans.len(),
        elapsed_ms = started.elapsed().as_millis(),
        "painted completed Word automatic cell borders"
    );
}

fn collect(frame: &Frame, tables: &mut Tables) {
    for (_, item) in frame.items() {
        match item {
            FrameItem::Group(group) => collect(&group.frame, tables),
            FrameItem::Tag(Tag::Start(content, _)) => {
                if content.to_packed::<TableElem>().is_some()
                    && content
                        .label()
                        .is_some_and(|label| label.resolve().starts_with(TABLE_LABEL_PREFIX))
                {
                    tables.spans.insert(content.span());
                }
                if let Some(metadata) = content.to_packed::<MetadataElem>()
                    && let Value::Array(values) = &metadata.value
                    && let [Value::Str(name), Value::Array(entries)] = values.as_slice()
                    && name.as_str() == MARKER
                {
                    for entry in entries {
                        if let Value::Array(fields) = entry
                            && let [Value::Color(marker), spec] = fields.as_slice()
                            && let Some(edge) = Edge::read(spec)
                        {
                            tables.edges.insert(marker.clone(), edge);
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

fn marker_edge<'a>(item: &FrameItem, tables: &'a Tables) -> Option<(&'a Edge, Abs)> {
    let FrameItem::Shape(shape, span) = item else {
        return None;
    };
    if !tables.spans.contains(span) {
        return None;
    }
    let Geometry::Line(end) = shape.geometry else {
        return None;
    };
    if end.x != Abs::zero() || end.y <= Abs::zero() {
        return None;
    }
    let stroke = shape.stroke.as_ref()?;
    if stroke.thickness != Abs::zero() {
        return None;
    }
    let Paint::Solid(marker) = &stroke.paint else {
        return None;
    };
    tables.edges.get(marker).map(|edge| (edge, end.y))
}

fn paint_frame(frame: &mut Frame, tables: &Tables) -> bool {
    let mut groups = HashMap::new();
    let mut last_grid_paint = HashMap::new();
    let mut borders: HashMap<Span, Vec<(Point, FrameItem)>> = HashMap::new();
    for (index, (position, item)) in frame.items().enumerate() {
        if let FrameItem::Group(group) = item {
            // Keep unrelated glyph runs and images shared instead of cloning
            // whole frame subtrees.
            let mut changed = group.clone();
            if paint_frame(&mut changed.frame, tables) {
                groups.insert(index, changed);
            }
        }
        if let FrameItem::Shape(_, span) = item
            && tables.spans.contains(span)
        {
            last_grid_paint.insert(*span, index);
            if let Some((edge, height)) = marker_edge(item, tables) {
                let mut paints = edge.paints(Size::new(Abs::zero(), height), Abs::zero());
                for (at, _) in &mut paints {
                    *at += *position;
                }
                borders.entry(*span).or_default().extend(paints);
            }
        }
    }
    if borders.is_empty() && groups.is_empty() {
        return false;
    }
    let mut items =
        Vec::with_capacity(frame.layer() + borders.values().map(Vec::len).sum::<usize>());
    for (index, (position, item)) in frame.items().enumerate() {
        if let Some(group) = groups.remove(&index) {
            items.push((*position, FrameItem::Group(group)));
        } else if marker_edge(item, tables).is_none() {
            items.push((*position, item.clone()));
        }
        if let FrameItem::Shape(_, span) = item
            && last_grid_paint.get(span) == Some(&index)
            && let Some(paints) = borders.remove(span)
        {
            // Winning edges extend into their neighbours. Draw after native
            // backgrounds and before cell contents so those fills cannot hide ink.
            items.extend(paints);
        }
    }
    frame.clear();
    frame.push_multiple(items);
    true
}
