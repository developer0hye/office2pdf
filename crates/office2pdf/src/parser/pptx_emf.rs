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

#[derive(Clone, Copy)]
enum EmfObject {
    Brush(BrushStyle),
    Pen(PenStyle),
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
    elements: Vec<SvgPathElement>,
}

impl EmfSvgConverter {
    fn convert(data: &[u8]) -> Option<Vec<u8>> {
        let mut converter = Self {
            fill_rule: FillRule::NonZero,
            bounds: BoundingBox::new(),
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
            EMR_SETWINDOWEXTEX | EMR_SETWINDOWORGEX | EMR_SETVIEWPORTEXTEX
            | EMR_SETVIEWPORTORGEX => {
                // The converter derives its SVG viewBox from actual drawn geometry instead
                // of these logical extents because many Office EMFs use a wider drawing space.
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
        if let Some(object) = self.objects.get(&handle).copied() {
            return Some(object);
        }

        match handle {
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
        self.elements.push(SvgPathElement {
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
        self.elements.push(SvgPathElement {
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

    fn finish(self) -> Option<Vec<u8>> {
        if self.bounds.is_empty || self.elements.is_empty() {
            return None;
        }

        let width: i32 = (self.bounds.max_x - self.bounds.min_x).max(1);
        let height: i32 = (self.bounds.max_y - self.bounds.min_y).max(1);
        let mut svg = String::new();
        let _ = writeln!(
            svg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{} {} {} {}\">",
            self.bounds.min_x, self.bounds.min_y, width, height
        );
        for element in self.elements {
            svg.push_str("  <path d=\"");
            svg.push_str(&element.data);
            svg.push('"');
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
                let _ = write!(svg, " stroke=\"{}\"", stroke.as_svg_hex());
                let _ = write!(
                    svg,
                    " stroke-width=\"{}\"",
                    element.stroke_width.unwrap_or(1)
                );
            }
            svg.push_str("/>\n");
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
