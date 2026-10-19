//! Paint diagonal cell rules in Typst's completed fill layer, before cell text.
//!
//! The private fill carries resolved cell bounds without contributing layout.
//! Restore the source fill and replace that marker with the requested rules.

use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::time::Instant;
use typst::foundations::{FromValue, Resolve, StyleChain, Value};
use typst::introspection::{MetadataElem, Tag};
use typst::layout::{Abs, Frame, FrameItem, Point, Size};
use typst::syntax::Span;
use typst::visualize::{Color, FixedStroke, Geometry, Paint, Stroke};
use typst_layout::Page;
#[cfg(target_arch = "wasm32")]
use web_time::Instant;

pub(super) const MARKER: &str = "office2pdf-cell-diagonal-borders";

struct Rule {
    double: bool,
    stroke: FixedStroke,
}

impl Rule {
    fn read(value: &Value) -> Option<Self> {
        let Value::Array(fields) = value else {
            return None;
        };
        let [Value::Bool(double), value] = fields.as_slice() else {
            return None;
        };
        let stroke = Stroke::from_value(value.clone())
            .ok()?
            .resolve(StyleChain::default())
            .unwrap_or_default();
        (stroke.thickness.to_pt().is_finite() && stroke.thickness > Abs::zero()).then(|| Self {
            double: *double,
            stroke,
        })
    }

    fn paints(&self, size: Size, upward: bool, origin: Point) -> Vec<(Point, FrameItem)> {
        let dy = if upward { -size.y } else { size.y };
        let start = if upward {
            Point::with_y(size.y)
        } else {
            Point::zero()
        };
        let end = Point::new(size.x, dy);
        let length = size.x.to_pt().hypot(size.y.to_pt());
        if length <= 0.0 || !length.is_finite() {
            return Vec::new();
        }
        let offsets: &[f64] = if self.double { &[-1.0, 1.0] } else { &[0.0] };
        offsets
            .iter()
            .map(|offset| {
                let distance = self.stroke.thickness * *offset;
                let shift = Point::new(
                    distance * (-dy.to_pt() / length),
                    distance * (size.x.to_pt() / length),
                );
                (
                    origin + start + shift,
                    FrameItem::Shape(
                        Geometry::Line(end).stroked(self.stroke.clone()),
                        Span::detached(),
                    ),
                )
            })
            .collect()
    }
}

struct CellPaint {
    fill: Option<Paint>,
    down: Option<Rule>,
    up: Option<Rule>,
}

pub(super) fn paint_cell_diagonals(pages: &mut [Page]) {
    let mut cells = HashMap::new();
    for page in pages.iter() {
        collect(&page.frame, &mut cells)
    }
    if cells.is_empty() {
        return;
    }
    let started = Instant::now();
    for page in pages.iter_mut() {
        paint_frame(&mut page.frame, &cells)
    }
    tracing::debug!(
        cells = cells.len(),
        elapsed_ms = started.elapsed().as_millis(),
        "painted completed table cell diagonals"
    );
}

fn collect(frame: &Frame, cells: &mut HashMap<Color, CellPaint>) {
    for (_, item) in frame.items() {
        match item {
            FrameItem::Group(group) => collect(&group.frame, cells),
            FrameItem::Tag(Tag::Start(content, _)) => {
                let Some(metadata) = content.to_packed::<MetadataElem>() else {
                    continue;
                };
                let Value::Array(values) = &metadata.value else {
                    continue;
                };
                let [Value::Str(name), Value::Array(entries)] = values.as_slice() else {
                    continue;
                };
                if name.as_str() != MARKER {
                    continue;
                }
                for entry in entries {
                    let Value::Array(fields) = entry else {
                        continue;
                    };
                    let [Value::Color(marker), fill, down, up] = fields.as_slice() else {
                        continue;
                    };
                    let fill = match fill {
                        Value::None => None,
                        Value::Color(color) => Some(Paint::Solid(color.clone())),
                        _ => continue,
                    };
                    cells.insert(
                        marker.clone(),
                        CellPaint {
                            fill,
                            down: Rule::read(down),
                            up: Rule::read(up),
                        },
                    );
                }
            }
            _ => {}
        }
    }
}

fn paint_frame(frame: &mut Frame, cells: &HashMap<Color, CellPaint>) {
    let mut output = Vec::new();
    for (position, item) in frame.items() {
        match item {
            FrameItem::Group(group) => {
                let mut group = group.clone();
                paint_frame(&mut group.frame, cells);
                output.push((*position, FrameItem::Group(group)));
            }
            FrameItem::Shape(shape, span) if shape.stroke.is_none() => {
                let (Geometry::Rect(size), Some(Paint::Solid(marker))) =
                    (&shape.geometry, &shape.fill)
                else {
                    output.push((*position, item.clone()));
                    continue;
                };
                let Some(cell) = cells.get(marker) else {
                    output.push((*position, item.clone()));
                    continue;
                };
                if let Some(fill) = &cell.fill {
                    output.push((
                        *position,
                        FrameItem::Shape(Geometry::Rect(*size).filled(fill.clone()), *span),
                    ));
                }
                if let Some(rule) = &cell.down {
                    output.extend(rule.paints(*size, false, *position))
                }
                if let Some(rule) = &cell.up {
                    output.extend(rule.paints(*size, true, *position))
                }
            }
            _ => output.push((*position, item.clone())),
        }
    }
    frame.clear();
    frame.push_multiple(output);
}
