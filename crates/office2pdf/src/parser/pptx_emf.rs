use std::collections::HashMap;
use std::fmt::Write;

const EMR_HEADER: u32 = 1;
const EMR_SETWINDOWEXTEX: u32 = 9;
const EMR_SETWINDOWORGEX: u32 = 10;
const EMR_SETVIEWPORTEXTEX: u32 = 11;
const EMR_SETVIEWPORTORGEX: u32 = 12;
const EMR_EOF: u32 = 14;
const EMR_SETPOLYFILLMODE: u32 = 19;
const EMR_MOVETOEX: u32 = 27;
const EMR_SELECTOBJECT: u32 = 37;
const EMR_CREATEBRUSHINDIRECT: u32 = 39;
const EMR_DELETEOBJECT: u32 = 40;
const EMR_BEGINPATH: u32 = 59;
const EMR_ENDPATH: u32 = 60;
const EMR_CLOSEFIGURE: u32 = 61;
const EMR_FILLPATH: u32 = 62;
const EMR_STROKEPATH: u32 = 64;
const EMR_POLYGON16: u32 = 86;
const EMR_POLYLINE16: u32 = 87;
const EMR_POLYBEZIERTO16: u32 = 88;
const EMR_POLYPOLYLINE16: u32 = 90;
const EMR_POLYPOLYGON16: u32 = 91;
const EMR_EXTCREATEPEN: u32 = 95;

const BS_SOLID: u32 = 0;
const BS_NULL: u32 = 1;
const PS_SOLID: u32 = 0;
const PS_NULL: u32 = 5;
const ALTERNATE_FILL_MODE: u32 = 1;
const NULL_BRUSH_STOCK_OBJECT: u32 = 0x8000_0005;
const NULL_PEN_STOCK_OBJECT: u32 = 0x8000_0008;

#[derive(Clone)]
enum EmfObject {
    Brush(BrushStyle),
    Pen(PenStyle),
    Font(EmfFont),
}

#[derive(Clone, Copy)]
enum BrushStyle {
    Solid(RgbColor),
    Null,
}

#[derive(Clone, Copy)]
enum PenStyle {
    Solid { color: RgbColor, width: i32 },
    Null,
}

#[derive(Clone, Copy, Default)]
enum FillRule {
    EvenOdd,
    #[default]
    NonZero,
}

#[derive(Clone, Copy)]
struct Point {
    x: i32,
    y: i32,
}

#[derive(Clone, Copy)]
struct RgbColor {
    red: u8,
    green: u8,
    blue: u8,
}

impl RgbColor {
    fn from_colorref(colorref: u32) -> Self {
        Self {
            red: (colorref & 0xFF) as u8,
            green: ((colorref >> 8) & 0xFF) as u8,
            blue: ((colorref >> 16) & 0xFF) as u8,
        }
    }

    fn as_svg_hex(self) -> String {
        format!("#{:02x}{:02x}{:02x}", self.red, self.green, self.blue)
    }
}

#[derive(Default)]
struct BoundingBox {
    min_x: i32,
    min_y: i32,
    max_x: i32,
    max_y: i32,
    is_empty: bool,
}

impl BoundingBox {
    fn new() -> Self {
        Self {
            min_x: 0,
            min_y: 0,
            max_x: 0,
            max_y: 0,
            is_empty: true,
        }
    }

    fn include(&mut self, point: Point) {
        if self.is_empty {
            self.min_x = point.x;
            self.min_y = point.y;
            self.max_x = point.x;
            self.max_y = point.y;
            self.is_empty = false;
            return;
        }

        self.min_x = self.min_x.min(point.x);
        self.min_y = self.min_y.min(point.y);
        self.max_x = self.max_x.max(point.x);
        self.max_y = self.max_y.max(point.y);
    }

    fn include_points(&mut self, points: &[Point]) {
        for point in points {
            self.include(*point);
        }
    }
}

#[path = "pptx_emf_text.rs"]
mod text;
use text::{EmfFont, EmfTextState};

struct SvgPathElement {
    data: String,
    fill: Option<RgbColor>,
    stroke: Option<RgbColor>,
    stroke_width: Option<i32>,
    fill_rule: FillRule,
}

#[derive(Default)]
struct EmfSvgConverter {
    objects: HashMap<u32, EmfObject>,
    current_brush: Option<BrushStyle>,
    current_pen: Option<PenStyle>,
    fill_rule: FillRule,
    current_path: Option<String>,
    current_path_points: Vec<Point>,
    current_point: Option<Point>,
    bounds: BoundingBox,
    elements: Vec<String>,
    text_state: EmfTextState,
    saved_states: Vec<EmfSavedState>,
    frame: Option<[f64; 4]>,
    window_origin: (i32, i32),
    window_extent: (i32, i32),
    viewport_origin: (i32, i32),
    viewport_extent: (i32, i32),
    clips: Vec<Vec<[i32; 4]>>,
}

#[derive(Clone)]
struct EmfSavedState {
    brush: Option<BrushStyle>,
    pen: Option<PenStyle>,
    fill_rule: FillRule,
    point: Option<Point>,
    text: EmfTextState,
    window_origin: (i32, i32),
    window_extent: (i32, i32),
    viewport_origin: (i32, i32),
    viewport_extent: (i32, i32),
    clips: Vec<Vec<[i32; 4]>>,
}

impl EmfSvgConverter {
    fn convert(data: &[u8]) -> Option<Vec<u8>> {
        let mut converter = Self {
            fill_rule: FillRule::NonZero,
            bounds: BoundingBox::new(),
            current_brush: Some(BrushStyle::Solid(RgbColor::from_colorref(0xffffff))),
            current_pen: Some(PenStyle::Solid {
                color: RgbColor::from_colorref(0),
                width: 1,
            }),
            window_extent: (1, 1),
            viewport_extent: (1, 1),
            ..Self::default()
        };
        converter.run(data)?;
        converter.finish()
    }

    fn run(&mut self, data: &[u8]) -> Option<()> {
        let mut offset: usize = 0;
        let mut saw_header: bool = false;

        while offset.checked_add(8)? <= data.len() {
            let record_type: u32 = read_u32(data, offset)?;
            let record_size: usize = read_u32(data, offset + 4)? as usize;
            if record_size < 8 || offset.checked_add(record_size)? > data.len() {
                return None;
            }
            let body: &[u8] = &data[offset + 8..offset + record_size];
            offset += record_size;

            if !saw_header {
                if record_type != EMR_HEADER {
                    return None;
                }
                saw_header = true;
                // A real header maps 0.01mm picture-frame coordinates into device pixels.
                // Legacy synthetic fixtures without a complete header retain geometry bounds.
                self.frame = text::header_frame(body);
                continue;
            }

            self.handle_record(record_type, body)?;
            if record_type == EMR_EOF {
                break;
            }
        }

        saw_header.then_some(())
    }

    fn handle_record(&mut self, record_type: u32, body: &[u8]) -> Option<()> {
        match record_type {
            EMR_SETWINDOWEXTEX => self.window_extent = read_point_pair(body)?,
            EMR_SETWINDOWORGEX => self.window_origin = read_point_pair(body)?,
            EMR_SETVIEWPORTEXTEX => self.viewport_extent = read_point_pair(body)?,
            EMR_SETVIEWPORTORGEX => self.viewport_origin = read_point_pair(body)?,
            18 => self.text_state.is_opaque = read_u32(body, 0)? == 2,
            22 => self.text_state.align = read_u32(body, 0)?,
            24 => self.text_state.color = RgbColor::from_colorref(read_u32(body, 0)?),
            25 => self.text_state.background = RgbColor::from_colorref(read_u32(body, 0)?),
            30 => self.clips.push(vec![read_rect(body, 0)?]),
            33 => self.save_state(),
            34 => self.restore_state(read_i32(body, 0)?)?,
            38 => {
                let handle = read_u32(body, 0)?;
                let pen = if read_u32(body, 4)? & 0xf == PS_NULL {
                    PenStyle::Null
                } else {
                    PenStyle::Solid {
                        color: RgbColor::from_colorref(read_u32(body, 16)?),
                        width: read_i32(body, 8)?.saturating_abs().max(1),
                    }
                };
                self.objects.insert(handle, EmfObject::Pen(pen));
            }
            43 => self.emit_rect(read_rect(body, 0)?),
            54 => {
                let end = Point {
                    x: read_i32(body, 0)?,
                    y: read_i32(body, 4)?,
                };
                if let Some(start) = self.current_point {
                    if let Some(path) = self.current_path.as_mut() {
                        let _ = write!(path, " L {} {}", end.x, end.y);
                        self.current_path_points.push(end);
                    } else {
                        self.emit_poly_path(&[start, end], false);
                    }
                }
                self.current_point = Some(end);
            }
            75 => self.select_clip_region(body)?,
            76 if read_u32(body, 88) == Some(0) => {
                let x = read_i32(body, 16)?;
                let y = read_i32(body, 20)?;
                let right = x.checked_add(read_i32(body, 24)?)?;
                let bottom = y.checked_add(read_i32(body, 28)?)?;
                let rop = read_u32(body, 32)?;
                let color = match rop {
                    0x00f00021 => self.current_fill(),                     // PATCOPY
                    0x00ff0062 => Some(RgbColor::from_colorref(0xffffff)), // WHITENESS
                    0x00000042 => Some(RgbColor::from_colorref(0)),        // BLACKNESS
                    _ => None,
                };
                if let Some(color) = color {
                    self.push_element(SvgPathElement {
                        data: format!("M {x} {y} H {right} V {bottom} H {x} Z"),
                        fill: Some(color),
                        stroke: None,
                        stroke_width: None,
                        fill_rule: self.fill_rule,
                    });
                    self.bounds.include_points(&[
                        Point { x, y },
                        Point {
                            x: right,
                            y: bottom,
                        },
                    ]);
                }
            }
            82 => {
                let (handle, font) = text::read_font(body)?;
                self.objects.insert(handle, EmfObject::Font(font));
            }
            84 => {
                let Some(font) = self.text_state.font.as_ref() else {
                    return Some(());
                };
                let Some(element) = text::write_text(body, &self.text_state, self.current_point)?
                else {
                    return Some(());
                };
                // The header frame controls placement; fallback bounds keep old minimal streams usable.
                let point = Point {
                    x: read_i32(body, 28)?,
                    y: read_i32(body, 32)?,
                };
                let height = font.height.saturating_abs().max(1);
                self.bounds.include_points(&[
                    point,
                    Point {
                        x: point.x.saturating_add(height),
                        y: point.y.saturating_add(height),
                    },
                ]);
                self.elements.push(self.wrap_element(element.svg));
                if self.text_state.align & 1 != 0 {
                    self.current_point = Some(element.next_point);
                }
            }
            EMR_SETPOLYFILLMODE => {
                let mode: u32 = read_u32(body, 0)?;
                self.fill_rule = if mode == ALTERNATE_FILL_MODE {
                    FillRule::EvenOdd
                } else {
                    FillRule::NonZero
                };
            }
            EMR_CREATEBRUSHINDIRECT => {
                let handle: u32 = read_u32(body, 0)?;
                let style: u32 = read_u32(body, 4)?;
                let colorref: u32 = read_u32(body, 8)?;
                let brush = match style {
                    BS_SOLID => BrushStyle::Solid(RgbColor::from_colorref(colorref)),
                    BS_NULL => BrushStyle::Null,
                    _ => BrushStyle::Null,
                };
                self.objects.insert(handle, EmfObject::Brush(brush));
            }
            EMR_EXTCREATEPEN => {
                let handle: u32 = read_u32(body, 0)?;
                let pen_style: u32 = read_u32(body, 20)? & 0xF;
                let width: i32 = read_u32(body, 24)? as i32;
                let colorref: u32 = read_u32(body, 32)?;
                let pen = match pen_style {
                    PS_SOLID => PenStyle::Solid {
                        color: RgbColor::from_colorref(colorref),
                        width: width.max(1),
                    },
                    PS_NULL => PenStyle::Null,
                    _ => PenStyle::Null,
                };
                self.objects.insert(handle, EmfObject::Pen(pen));
            }
            EMR_SELECTOBJECT => {
                let handle: u32 = read_u32(body, 0)?;
                match self.resolve_object(handle) {
                    Some(EmfObject::Brush(brush)) => self.current_brush = Some(brush),
                    Some(EmfObject::Pen(pen)) => self.current_pen = Some(pen),
                    Some(EmfObject::Font(font)) => self.text_state.font = Some(font),
                    None => {}
                }
            }
            EMR_DELETEOBJECT => {
                let handle: u32 = read_u32(body, 0)?;
                self.objects.remove(&handle);
            }
            EMR_BEGINPATH => {
                self.current_path = Some(String::new());
                self.current_path_points.clear();
                self.current_point = None;
            }
            EMR_MOVETOEX => {
                let point = Point {
                    x: read_i32(body, 0)?,
                    y: read_i32(body, 4)?,
                };
                self.current_point = Some(point);
                if let Some(path) = self.current_path.as_mut() {
                    append_move_to(path, point);
                    self.current_path_points.push(point);
                }
            }
            EMR_POLYBEZIERTO16 => self.handle_polybezier_to16(body)?,
            EMR_CLOSEFIGURE => {
                if let Some(path) = self.current_path.as_mut()
                    && !path.is_empty()
                {
                    path.push_str(" Z");
                }
            }
            EMR_ENDPATH => {}
            EMR_FILLPATH => self.flush_path(false),
            EMR_STROKEPATH => self.flush_path(true),
            EMR_POLYGON16 => self.handle_polygon16(body)?,
            EMR_POLYLINE16 => self.handle_polyline16(body)?,
            EMR_POLYPOLYGON16 => self.handle_poly_shape16(body, true)?,
            EMR_POLYPOLYLINE16 => self.handle_poly_shape16(body, false)?,
            EMR_EOF => {}
            _ => {}
        }
        Some(())
    }

    fn resolve_object(&self, handle: u32) -> Option<EmfObject> {
        if let Some(object) = self.objects.get(&handle).cloned() {
            return Some(object);
        }

        match handle {
            0x8000_0000 => Some(EmfObject::Brush(BrushStyle::Solid(
                RgbColor::from_colorref(0xffffff),
            ))),
            0x8000_0004 => Some(EmfObject::Brush(BrushStyle::Solid(
                RgbColor::from_colorref(0),
            ))),
            0x8000_0006 => Some(EmfObject::Pen(PenStyle::Solid {
                color: RgbColor::from_colorref(0xffffff),
                width: 1,
            })),
            0x8000_0007 => Some(EmfObject::Pen(PenStyle::Solid {
                color: RgbColor::from_colorref(0),
                width: 1,
            })),
            NULL_BRUSH_STOCK_OBJECT => Some(EmfObject::Brush(BrushStyle::Null)),
            NULL_PEN_STOCK_OBJECT => Some(EmfObject::Pen(PenStyle::Null)),
            _ => None,
        }
    }

    fn handle_polybezier_to16(&mut self, body: &[u8]) -> Option<()> {
        let point_count: usize = read_u32(body, 16)? as usize;
        let points: Vec<Point> = parse_points16(body, 20, point_count)?;
        if points.is_empty() {
            return Some(());
        }

        let path = self.current_path.as_mut()?;
        if self.current_point.is_none() {
            append_move_to(path, points[0]);
            self.current_path_points.push(points[0]);
            self.current_point = Some(points[0]);
        }

        let mut chunk_start: usize = 0;
        while chunk_start + 2 < points.len() {
            let control1 = points[chunk_start];
            let control2 = points[chunk_start + 1];
            let end_point = points[chunk_start + 2];
            let _ = write!(
                path,
                " C {} {} {} {} {} {}",
                control1.x, control1.y, control2.x, control2.y, end_point.x, end_point.y
            );
            self.current_path_points
                .extend_from_slice(&[control1, control2, end_point]);
            self.current_point = Some(end_point);
            chunk_start += 3;
        }

        Some(())
    }

    fn handle_polygon16(&mut self, body: &[u8]) -> Option<()> {
        let points: Vec<Point> = parse_points16(body, 20, read_u32(body, 16)? as usize)?;
        self.emit_poly_path(&points, true);
        Some(())
    }

    fn handle_polyline16(&mut self, body: &[u8]) -> Option<()> {
        let points: Vec<Point> = parse_points16(body, 20, read_u32(body, 16)? as usize)?;
        self.emit_poly_path(&points, false);
        Some(())
    }

    fn handle_poly_shape16(&mut self, body: &[u8], is_polygon: bool) -> Option<()> {
        let polygon_count: usize = read_u32(body, 16)? as usize;
        let point_count: usize = read_u32(body, 20)? as usize;
        let counts_offset: usize = 24;
        let mut counts = Vec::with_capacity(polygon_count);
        for index in 0..polygon_count {
            counts.push(read_u32(body, counts_offset + index * 4)? as usize);
        }

        let points_offset: usize = counts_offset + polygon_count * 4;
        let points: Vec<Point> = parse_points16(body, points_offset, point_count)?;
        let mut point_index: usize = 0;

        for count in counts {
            let next_index: usize = point_index.checked_add(count)?;
            let polygon_points: &[Point] = points.get(point_index..next_index)?;
            self.emit_poly_path(polygon_points, is_polygon);
            point_index = next_index;
        }

        Some(())
    }

    fn emit_poly_path(&mut self, points: &[Point], close_path: bool) {
        if points.is_empty() {
            return;
        }

        let fill: Option<RgbColor> = if close_path {
            self.current_fill()
        } else {
            None
        };
        let stroke: Option<RgbColor> = self.current_stroke_color();
        let stroke_width: Option<i32> = self.current_stroke_width();
        if fill.is_none() && stroke.is_none() {
            return;
        }

        let mut data = String::new();
        append_move_to(&mut data, points[0]);
        for point in &points[1..] {
            let _ = write!(data, " L {} {}", point.x, point.y);
        }
        if close_path {
            data.push_str(" Z");
        }

        self.bounds.include_points(points);
        self.push_element(SvgPathElement {
            data,
            fill,
            stroke,
            stroke_width,
            fill_rule: self.fill_rule,
        });
    }

    fn flush_path(&mut self, stroke_only: bool) {
        let Some(path) = self.current_path.take() else {
            return;
        };
        if path.is_empty() {
            self.current_path_points.clear();
            self.current_point = None;
            return;
        }

        let fill: Option<RgbColor> = if stroke_only {
            None
        } else {
            self.current_fill()
        };
        let stroke: Option<RgbColor> = if stroke_only {
            self.current_stroke_color()
        } else {
            None
        };
        let stroke_width: Option<i32> = if stroke_only {
            self.current_stroke_width()
        } else {
            None
        };
        if fill.is_none() && stroke.is_none() {
            self.current_path_points.clear();
            self.current_point = None;
            return;
        }

        self.bounds.include_points(&self.current_path_points);
        self.push_element(SvgPathElement {
            data: path,
            fill,
            stroke,
            stroke_width,
            fill_rule: self.fill_rule,
        });
        self.current_path_points.clear();
        self.current_point = None;
    }

    fn current_fill(&self) -> Option<RgbColor> {
        match self.current_brush {
            Some(BrushStyle::Solid(color)) => Some(color),
            _ => None,
        }
    }

    fn current_stroke_color(&self) -> Option<RgbColor> {
        match self.current_pen {
            Some(PenStyle::Solid { color, .. }) => Some(color),
            _ => None,
        }
    }

    fn current_stroke_width(&self) -> Option<i32> {
        match self.current_pen {
            Some(PenStyle::Solid { width, .. }) => Some(width.max(1)),
            _ => None,
        }
    }

    fn save_state(&mut self) {
        self.saved_states.push(EmfSavedState {
            brush: self.current_brush,
            pen: self.current_pen,
            fill_rule: self.fill_rule,
            point: self.current_point,
            text: self.text_state.clone(),
            window_origin: self.window_origin,
            window_extent: self.window_extent,
            viewport_origin: self.viewport_origin,
            viewport_extent: self.viewport_extent,
            clips: self.clips.clone(),
        });
    }

    fn restore_state(&mut self, level: i32) -> Option<()> {
        let index = if level < 0 {
            self.saved_states
                .len()
                .checked_sub(level.unsigned_abs() as usize)?
        } else {
            (level as usize).checked_sub(1)?
        };
        let state = self.saved_states.get(index)?.clone();
        self.saved_states.truncate(index);
        self.current_brush = state.brush;
        self.current_pen = state.pen;
        self.fill_rule = state.fill_rule;
        self.current_point = state.point;
        self.text_state = state.text;
        self.window_origin = state.window_origin;
        self.window_extent = state.window_extent;
        self.viewport_origin = state.viewport_origin;
        self.viewport_extent = state.viewport_extent;
        self.clips = state.clips;
        Some(())
    }

    fn select_clip_region(&mut self, body: &[u8]) -> Option<()> {
        let size = read_u32(body, 0)? as usize;
        let mode = read_u32(body, 4)?;
        if size == 0 {
            if mode == 5 {
                self.clips.clear();
            }
            return Some(());
        }
        let region = body.get(8..8usize.checked_add(size)?)?;
        let count = read_u32(region, 8)? as usize;
        let offset = read_u32(region, 0)? as usize;
        let _ = region.get(offset..offset.checked_add(count.checked_mul(16)?)?)?;
        let rects = (0..count)
            .map(|i| read_rect(region, offset + i * 16))
            .collect::<Option<Vec<_>>>()?;
        match mode {
            5 => {
                self.clips.clear();
                self.clips.push(rects);
            }
            1 => self.clips.push(rects),
            _ => {
                tracing::warn!(mode, "Unsupported EMF clipping region combination");
            }
        }
        Some(())
    }

    fn emit_rect(&mut self, rect: [i32; 4]) {
        let [left, top, right, bottom] = rect;
        self.emit_poly_path(
            &[
                Point { x: left, y: top },
                Point { x: right, y: top },
                Point {
                    x: right,
                    y: bottom,
                },
                Point { x: left, y: bottom },
            ],
            true,
        );
    }

    fn wrap_element(&self, element: String) -> String {
        let sx = f64::from(self.viewport_extent.0) / f64::from(self.window_extent.0);
        let sy = f64::from(self.viewport_extent.1) / f64::from(self.window_extent.1);
        if !sx.is_finite() || !sy.is_finite() {
            return String::new();
        }
        let dx = f64::from(self.viewport_origin.0) - f64::from(self.window_origin.0) * sx;
        let dy = f64::from(self.viewport_origin.1) - f64::from(self.window_origin.1) * sy;
        let mut svg = format!("<g transform=\"matrix({sx} 0 0 {sy} {dx} {dy})\">");
        // Nested SVGs intersect the saved clipping regions; rectangles within one region form a union.
        for (index, rects) in self.clips.iter().enumerate() {
            let _ = write!(
                svg,
                "<defs><clipPath id=\"clip-{}-{index}\">",
                self.elements.len()
            );
            for &[l, t, r, b] in rects {
                let width = i64::from(r) - i64::from(l);
                let height = i64::from(b) - i64::from(t);
                let _ = write!(
                    svg,
                    "<rect x=\"{l}\" y=\"{t}\" width=\"{}\" height=\"{}\"/>",
                    width.max(0),
                    height.max(0)
                );
            }
            let _ = write!(
                svg,
                "</clipPath></defs><g clip-path=\"url(#clip-{}-{index})\">",
                self.elements.len()
            );
        }
        svg.push_str(&element);
        for _ in &self.clips {
            svg.push_str("</g>");
        }
        svg.push_str("</g>");
        svg
    }

    fn push_element(&mut self, element: SvgPathElement) {
        let mut svg = format!("<path d=\"{}\"", element.data);
        match element.fill {
            Some(fill) => {
                let _ = write!(svg, " fill=\"{}\"", fill.as_svg_hex());
                if matches!(element.fill_rule, FillRule::EvenOdd) {
                    svg.push_str(" fill-rule=\"evenodd\"");
                }
            }
            None => svg.push_str(" fill=\"none\""),
        }
        if let Some(stroke) = element.stroke {
            let _ = write!(
                svg,
                " stroke=\"{}\" stroke-width=\"{}\"",
                stroke.as_svg_hex(),
                element.stroke_width.unwrap_or(1)
            );
        }
        svg.push_str("/>");
        self.elements.push(self.wrap_element(svg));
    }

    fn finish(self) -> Option<Vec<u8>> {
        if self.elements.is_empty() || (self.frame.is_none() && self.bounds.is_empty) {
            return None;
        }
        let [left, top, right, bottom] = self.frame.unwrap_or([
            f64::from(self.bounds.min_x),
            f64::from(self.bounds.min_y),
            f64::from(self.bounds.max_x),
            f64::from(self.bounds.max_y),
        ]);
        let width = (right - left).max(1.0);
        let height = (bottom - top).max(1.0);
        let mut svg = format!(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{left} {top} {width} {height}\">\n"
        );
        for element in self.elements {
            svg.push_str(&element);
            svg.push('\n');
        }
        svg.push_str("</svg>\n");
        Some(svg.into_bytes())
    }
}

fn append_move_to(out: &mut String, point: Point) {
    if !out.is_empty() {
        out.push(' ');
    }
    let _ = write!(out, "M {} {}", point.x, point.y);
}

fn parse_points16(data: &[u8], offset: usize, count: usize) -> Option<Vec<Point>> {
    let mut points = Vec::with_capacity(count);
    for index in 0..count {
        let point_offset: usize = offset.checked_add(index.checked_mul(4)?)?;
        points.push(Point {
            x: read_i16(data, point_offset)? as i32,
            y: read_i16(data, point_offset + 2)? as i32,
        });
    }
    Some(points)
}

fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    let bytes: [u8; 4] = data.get(offset..offset + 4)?.try_into().ok()?;
    Some(u32::from_le_bytes(bytes))
}

fn read_i32(data: &[u8], offset: usize) -> Option<i32> {
    let bytes: [u8; 4] = data.get(offset..offset + 4)?.try_into().ok()?;
    Some(i32::from_le_bytes(bytes))
}

fn read_i16(data: &[u8], offset: usize) -> Option<i16> {
    let bytes: [u8; 2] = data.get(offset..offset + 2)?.try_into().ok()?;
    Some(i16::from_le_bytes(bytes))
}

pub(super) fn convert_emf_to_svg(data: &[u8]) -> Option<Vec<u8>> {
    EmfSvgConverter::convert(data)
}

fn read_point_pair(data: &[u8]) -> Option<(i32, i32)> {
    Some((read_i32(data, 0)?, read_i32(data, 4)?))
}
fn read_rect(data: &[u8], offset: usize) -> Option<[i32; 4]> {
    Some([
        read_i32(data, offset)?,
        read_i32(data, offset + 4)?,
        read_i32(data, offset + 8)?,
        read_i32(data, offset + 12)?,
    ])
}

#[cfg(test)]
#[path = "pptx_emf_tests.rs"]
mod tests;
