//! Raw-XML side-channel for floating DrawingML shapes and grouped drawings.
//!
//! docx-rs (the upstream DOCX parser) only models `<w:drawing>` data as either a
//! picture (`Pic`) or a text box (`TextBox`). A DrawingML word-processing shape
//! that carries geometry but no text box — e.g. a `<a:prstGeom prst="rect">`
//! rectangle or a `prst="line"` connector/arrow authored by LibreOffice — parses
//! into a `Drawing` with `data == None`, so its geometry, fill and stroke are
//! lost entirely (issue #176).
//!
//! This module scans the raw `word/document.xml` in document order for those
//! geometry-only shapes, WordprocessingGroup (`wpg`) shape children, and
//! canvas picture offsets. Cursors keep that metadata aligned with the main
//! docx-rs drawing walk.

use crate::parser::xml_util::OOXML_XML_VERSION;
use std::cell::Cell;
use std::collections::HashMap;

use docx_rs::FromXML;
use quick_xml::events::{BytesStart, Event};

use crate::ir::{
    ArrowHead, BorderLineStyle, BorderSide, Color, FloatingShape, GradientFill, Insets, LineCap,
    LineJoin, Shape, ShapeKind, Subpath, TextBoxVerticalAlign, WrapMode,
};
use crate::parser::drawingml::{
    ParsedColor, SchemeColors, line_cap, parse_color_from_empty, parse_color_from_start,
    parse_theme_color_scheme,
};
use crate::parser::pptx::custom_geometry::parse_custom_geometry;
use crate::parser::pptx::geometry_guides::ShapeExtent;
use crate::parser::pptx::{self, ThemeLineStyle, drawingml_line_join};
use crate::parser::units::emu_to_pt;
use crate::parser::xml_util::parse_hex_color;

use super::super::parse_docx_shape_gradient;

#[derive(Debug, Clone)]
pub(in super::super) struct WpgShapeInfo {
    pub(in super::super) shape: Option<FloatingShape>,
    pub(in super::super) content: Vec<docx_rs::DocumentChild>,
    pub(in super::super) width: f64,
    pub(in super::super) height: f64,
    pub(in super::super) rotation_deg: Option<f64>,
    pub(in super::super) padding: Insets,
    pub(in super::super) vertical_align: TextBoxVerticalAlign,
    pub(in super::super) text_color: Option<Color>,
    pub(in super::super) offset_x: f64,
    pub(in super::super) offset_y: f64,
    pub(in super::super) wrap_mode: WrapMode,
}

#[derive(Debug, Clone, Default)]
pub(in super::super) struct WpgDrawingInfo {
    pub(in super::super) children: Vec<WpgShapeInfo>,
}

/// Default stroke width (pt) when a shape's `<a:ln w="0">` requests the
/// renderer's hairline default. Word/LibreOffice treat `w="0"` as "thin but
/// visible"; 0 pt would make the outline disappear.
const DEFAULT_STROKE_WIDTH_PT: f64 = 0.75;

/// EMU per point (914400 EMU/inch ÷ 72 pt/inch).
const EMU_PER_POINT: f64 = 12700.0;

/// `wps:bodyPr` content insets OOXML applies to every side the element leaves
/// out: `lIns`/`rIns` 91440 EMU and `tIns`/`bIns` 45720 EMU. A box with no
/// `wps:bodyPr` at all still seats its text by these.
const DEFAULT_TEXT_BOX_INSET_HORIZONTAL_PT: f64 = 91440.0 / EMU_PER_POINT;
const DEFAULT_TEXT_BOX_INSET_VERTICAL_PT: f64 = 45720.0 / EMU_PER_POINT;

/// Raw drawing metadata scanned from `word/document.xml`, consumed in document
/// order alongside the docx-rs element walk.
#[derive(Debug, Clone)]
pub(in super::super) struct DrawingShapeContext {
    shapes: Vec<FloatingShape>,
    cursor: Cell<usize>,
    wpg_drawings: Vec<Option<WpgDrawingInfo>>,
    wpg_cursor: Cell<usize>,
    canvas_image_offsets: Vec<Option<(f64, f64)>>,
    canvas_cursor: Cell<usize>,
}

impl DrawingShapeContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self::from_xml_with_theme(xml, None)
    }

    pub(in super::super) fn from_xml_with_theme(
        xml: Option<&str>,
        theme_xml: Option<&str>,
    ) -> Self {
        let theme_line_styles: Vec<ThemeLineStyle> = theme_xml
            .map(|xml| pptx::parse_theme_line_styles(xml))
            .unwrap_or_default();
        Self {
            shapes: xml
                .map(|xml| scan_drawing_shapes_with_theme(xml, &theme_line_styles))
                .unwrap_or_default(),
            cursor: Cell::new(0),
            wpg_drawings: xml
                .map(|xml| scan_wpg_drawings_with_theme(xml, theme_xml, &theme_line_styles))
                .unwrap_or_default(),
            wpg_cursor: Cell::new(0),
            canvas_image_offsets: xml.map(scan_canvas_image_offsets).unwrap_or_default(),
            canvas_cursor: Cell::new(0),
        }
    }

    /// Return the next scanned shape, advancing the cursor. Returns `None` once
    /// the scanned shapes are exhausted so a mismatched walk degrades to "no
    /// shape" rather than panicking.
    pub(in super::super) fn consume_next(&self) -> Option<FloatingShape> {
        let index: usize = self.cursor.get();
        self.cursor.set(index + 1);
        self.shapes.get(index).cloned()
    }

    /// Advance once for every docx-rs `Drawing`, returning WPG children only
    /// when the matching raw drawing is a WordprocessingGroup.
    pub(in super::super) fn consume_wpg_drawing(&self) -> Option<WpgDrawingInfo> {
        let index: usize = self.wpg_cursor.get();
        self.wpg_cursor.set(index + 1);
        self.wpg_drawings.get(index).cloned().flatten()
    }

    /// Return the selected picture's offset inside a WordprocessingCanvas.
    /// The cursor advances for every docx-rs drawing so AlternateContent
    /// fallbacks remain metadata-only and are never rendered a second time.
    pub(in super::super) fn consume_canvas_image_offset(&self) -> Option<(f64, f64)> {
        let index: usize = self.canvas_cursor.get();
        self.canvas_cursor.set(index + 1);
        self.canvas_image_offsets.get(index).copied().flatten()
    }
}

/// Which `<wp:positionH>` / `<wp:positionV>` axis the current `<wp:posOffset>`
/// text belongs to.
#[derive(Clone, Copy, PartialEq, Eq)]
enum PositionAxis {
    None,
    Horizontal,
    Vertical,
}

/// Mutable accumulator for a single `<w:drawing>` while scanning.
#[derive(Default)]
pub(super) struct ShapeBuilder {
    has_wsp: bool,
    has_wpg: bool,
    has_text_box: bool,
    preset: Option<String>,
    box_width_pt: Option<f64>,
    box_height_pt: Option<f64>,
    offset_x_pt: f64,
    offset_y_pt: f64,
    flip_h: bool,
    flip_v: bool,
    rotation_deg: Option<f64>,
    fill_color: Option<Color>,
    fill_opacity: Option<f64>,
    fill_none: bool,
    line_color: Option<Color>,
    line_width_pt: Option<f64>,
    line_none: bool,
    has_line: bool,
    line_cap: Option<LineCap>,
    line_join: Option<LineJoin>,
    line_reference_index: Option<usize>,
    head_arrow: bool,
    tail_arrow: bool,
    custom_subpaths: Vec<Subpath>,
    gradient_fill: Option<GradientFill>,
    body_inset_left_pt: Option<f64>,
    body_inset_top_pt: Option<f64>,
    body_inset_right_pt: Option<f64>,
    body_inset_bottom_pt: Option<f64>,
}

impl ShapeBuilder {
    /// Build a [`FloatingShape`] from the accumulated geometry, or `None` when
    /// this drawing is not a geometry-only shape (it is a picture or a text box,
    /// both handled by docx-rs).
    fn finish(self, theme_line_styles: &[ThemeLineStyle]) -> Option<FloatingShape> {
        self.finish_with_text_box(false, theme_line_styles)
    }

    fn finish_wpg(self, theme_line_styles: &[ThemeLineStyle]) -> Option<FloatingShape> {
        self.finish_with_text_box(true, theme_line_styles)
    }

    fn finish_with_text_box(
        self,
        allow_text_box: bool,
        theme_line_styles: &[ThemeLineStyle],
    ) -> Option<FloatingShape> {
        if !self.has_wsp || self.has_wpg || (self.has_text_box && !allow_text_box) {
            return None;
        }

        let width: f64 = self.box_width_pt.unwrap_or(0.0);
        let height: f64 = self.box_height_pt.unwrap_or(0.0);
        let mut kind: ShapeKind = self.resolve_kind(width, height);
        mirror_shape_kind(&mut kind, self.flip_h, self.flip_v);
        let stroke: Option<BorderSide> = self.resolve_stroke(theme_line_styles);

        let mut gradient_fill: Option<GradientFill> = self.gradient_fill;
        if let Some(gradient) = gradient_fill.as_mut() {
            gradient.angle = mirrored_angle(gradient.angle, self.flip_h, self.flip_v);
        }

        let fill: Option<Color> = if self.fill_none {
            None
        } else {
            gradient_fill
                .as_ref()
                .and_then(|gradient| gradient.stops.first().map(|stop| stop.color))
                .or(self.fill_color)
        };
        // A shape with neither fill, stroke nor a line geometry would render as
        // nothing — skip it so we stay in sync with the renderer.
        let is_line: bool = matches!(kind, ShapeKind::Line { .. });
        if fill.is_none() && gradient_fill.is_none() && stroke.is_none() && !is_line {
            return None;
        }

        Some(FloatingShape {
            shape: Shape {
                kind,
                fill,
                gradient_fill,
                pattern_fill: None,
                stroke,
                rotation_deg: self.rotation_deg,
                opacity: self.fill_opacity,
                shadow: None,
                top_bevel: None,
            },
            width,
            height,
            offset_x: self.offset_x_pt,
            offset_y: self.offset_y_pt,
            wrap_mode: WrapMode::None,
        })
    }

    fn resolve_kind(&self, width: f64, height: f64) -> ShapeKind {
        if !self.custom_subpaths.is_empty() {
            return ShapeKind::Path {
                subpaths: self.custom_subpaths.clone(),
            };
        }

        match self.preset.as_deref() {
            Some("line") | Some("straightConnector1") => {
                // Endpoints run corner-to-corner of the bounding box; flips swap
                // the diagonal direction (no-op for axis-aligned lines).
                let (x1, x2) = if self.flip_h {
                    (width, 0.0)
                } else {
                    (0.0, width)
                };
                let (y1, y2) = if self.flip_v {
                    (height, 0.0)
                } else {
                    (0.0, height)
                };
                ShapeKind::Line {
                    x1,
                    y1,
                    x2,
                    y2,
                    head_end: arrow(self.head_arrow),
                    tail_end: arrow(self.tail_arrow),
                }
            }
            Some("ellipse") | Some("oval") => ShapeKind::Ellipse,
            Some("roundRect") => ShapeKind::RoundedRectangle {
                radius_fraction: 0.1,
            },
            Some("triangle") => ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
            },
            Some("diamond") => ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
            },
            // "rect" and any unsupported preset fall back to a rectangle so the
            // shape's area, fill and outline are still conveyed.
            _ => ShapeKind::Rectangle,
        }
    }

    /// Whether this drawing is a `wps:wsp` text box rather than a picture,
    /// a group, or a geometry-only shape.
    pub(super) fn is_text_box(&self) -> bool {
        self.has_wsp && self.has_text_box && !self.has_wpg
    }

    pub(super) fn box_size_pt(&self) -> (Option<f64>, Option<f64>) {
        (self.box_width_pt, self.box_height_pt)
    }

    /// The outline and background a text box shape paints around its text.
    /// `finish_with_text_box` discards a text box drawing, so the inline box of
    /// issue #1690 resolves its frame from the same accumulated `a:ln` and
    /// `a:solidFill` rather than parsing them a second way.
    pub(super) fn text_box_frame(
        &self,
        theme_line_styles: &[ThemeLineStyle],
    ) -> (Option<BorderSide>, Option<Color>) {
        let fill: Option<Color> = if self.fill_none {
            None
        } else {
            self.fill_color
        };
        (self.resolve_stroke(theme_line_styles), fill)
    }

    /// Where the box's text starts inside it, from `wps:bodyPr`.
    pub(super) fn text_box_insets(&self) -> Insets {
        Insets {
            top: self
                .body_inset_top_pt
                .unwrap_or(DEFAULT_TEXT_BOX_INSET_VERTICAL_PT),
            right: self
                .body_inset_right_pt
                .unwrap_or(DEFAULT_TEXT_BOX_INSET_HORIZONTAL_PT),
            bottom: self
                .body_inset_bottom_pt
                .unwrap_or(DEFAULT_TEXT_BOX_INSET_VERTICAL_PT),
            left: self
                .body_inset_left_pt
                .unwrap_or(DEFAULT_TEXT_BOX_INSET_HORIZONTAL_PT),
        }
    }

    fn resolve_stroke(&self, theme_line_styles: &[ThemeLineStyle]) -> Option<BorderSide> {
        if self.line_none || !self.has_line {
            return None;
        }
        let width: f64 = match self.line_width_pt {
            Some(width) if width > 0.0 => width,
            _ => DEFAULT_STROKE_WIDTH_PT,
        };
        let theme_style: Option<&ThemeLineStyle> = self.line_reference_index.and_then(|index| {
            index
                .checked_sub(1)
                .and_then(|index| theme_line_styles.get(index))
        });
        Some(BorderSide {
            width,
            color: self.line_color.unwrap_or(Color { r: 0, g: 0, b: 0 }),
            style: BorderLineStyle::Solid,
            join: self
                .line_join
                .or_else(|| theme_style.and_then(|style| style.join))
                .unwrap_or(LineJoin::Round),
            cap: self
                .line_cap
                .or_else(|| theme_style.and_then(|style| style.cap))
                .unwrap_or(LineCap::Flat),
        })
    }
}

fn mirror_shape_kind(kind: &mut ShapeKind, flip_h: bool, flip_v: bool) {
    if !flip_h && !flip_v {
        return;
    }

    let mirror = |(x, y): (f64, f64)| -> (f64, f64) {
        (
            if flip_h { 1.0 - x } else { x },
            if flip_v { 1.0 - y } else { y },
        )
    };
    match kind {
        ShapeKind::Polygon { vertices } => {
            for point in vertices {
                *point = mirror(*point);
            }
        }
        ShapeKind::Path { subpaths } => {
            for subpath in subpaths {
                for point in &mut subpath.vertices {
                    *point = mirror(*point);
                }
            }
        }
        // Line flips are already expressed by its endpoint direction in
        // `resolve_kind`; the remaining preset geometries are symmetric.
        _ => {}
    }
}

fn mirrored_angle(angle: f64, flip_h: bool, flip_v: bool) -> f64 {
    let horizontal: f64 = if flip_h { 180.0 - angle } else { angle };
    let vertical: f64 = if flip_v { -horizontal } else { horizontal };
    vertical.rem_euclid(360.0)
}

fn arrow(present: bool) -> ArrowHead {
    if present {
        ArrowHead::Triangle
    } else {
        ArrowHead::None
    }
}

fn attribute_value(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

fn emu_attr_to_pt(element: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attribute_value(element, name)
        .and_then(|value| value.parse::<i64>().ok())
        .map(emu_to_pt)
}

fn bool_attr(element: &BytesStart<'_>, name: &[u8]) -> bool {
    matches!(
        attribute_value(element, name).as_deref(),
        Some("1") | Some("true")
    )
}

fn rotation_attr_degrees(element: &BytesStart<'_>) -> Option<f64> {
    attribute_value(element, b"rot")
        .and_then(|value| value.parse::<f64>().ok())
        .map(|raw| (raw / 60_000.0).rem_euclid(360.0))
        .filter(|degrees| degrees.abs() > f64::EPSILON)
}

/// Scan `word/document.xml`, returning one [`FloatingShape`] per geometry-only
/// `wps:wsp` drawing, in document order.
#[cfg(test)]
fn scan_drawing_shapes(xml: &str) -> Vec<FloatingShape> {
    scan_drawing_shapes_with_theme(xml, &[])
}

fn scan_drawing_shapes_with_theme(
    xml: &str,
    theme_line_styles: &[ThemeLineStyle],
) -> Vec<FloatingShape> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<FloatingShape> = Vec::new();

    let mut drawing_depth: usize = 0;
    let mut state: ShapeScanState = ShapeScanState::default();
    let mut axis: PositionAxis = PositionAxis::None;
    let mut in_position_offset: bool = false;
    let mut builder: Option<ShapeBuilder> = None;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element)) => match element.local_name().as_ref() {
                b"drawing" => {
                    drawing_depth += 1;
                    if drawing_depth == 1 {
                        builder = Some(ShapeBuilder::default());
                        state.reset();
                        axis = PositionAxis::None;
                        in_position_offset = false;
                    }
                }
                b"positionH" => axis = PositionAxis::Horizontal,
                b"positionV" => axis = PositionAxis::Vertical,
                b"posOffset" => in_position_offset = true,
                _ => state.start(builder.as_mut(), element),
            },
            Ok(Event::Empty(ref element)) => {
                state.empty(builder.as_mut(), element);
            }
            Ok(Event::Text(ref text)) => {
                if in_position_offset
                    && let Some(builder) = builder.as_mut()
                    && let Ok(raw) = text.xml_content(OOXML_XML_VERSION)
                    && let Ok(emu) = raw.trim().parse::<i64>()
                {
                    let pt: f64 = (emu as f64) / EMU_PER_POINT;
                    match axis {
                        PositionAxis::Horizontal => builder.offset_x_pt = pt,
                        PositionAxis::Vertical => builder.offset_y_pt = pt,
                        PositionAxis::None => {}
                    }
                }
            }
            Ok(Event::End(ref element)) => match element.local_name().as_ref() {
                b"posOffset" => in_position_offset = false,
                b"positionH" | b"positionV" => axis = PositionAxis::None,
                b"drawing" if drawing_depth > 0 => {
                    drawing_depth -= 1;
                    if drawing_depth == 0
                        && let Some(shape) = builder
                            .take()
                            .and_then(|builder| builder.finish(theme_line_styles))
                    {
                        result.push(shape);
                    }
                }
                other => state.end(other),
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    result
}

/// The `<a:ln>` / `<wps:spPr>` nesting a `wps:wsp` scan has to track to tell a
/// line colour from a fill colour.
///
/// Two scanners walk `word/document.xml` for `wps:wsp` drawings: this module's
/// geometry-only one and the text box one in `docx_context_drawing.rs`. Sharing
/// the state and the dispatch keeps one reading of the OOXML rather than letting
/// the two drift apart (issue #1690).
#[derive(Default)]
pub(super) struct ShapeScanState {
    sppr_depth: usize,
    line_depth: usize,
}

impl ShapeScanState {
    /// Feed a start element to the builder, tracking `spPr`/`ln` nesting.
    pub(super) fn start(&mut self, builder: Option<&mut ShapeBuilder>, element: &BytesStart<'_>) {
        match element.local_name().as_ref() {
            b"spPr" if builder.is_some() => self.sppr_depth += 1,
            b"ln" if builder.is_some() => {
                self.line_depth += 1;
                if let Some(builder) = builder {
                    builder.has_line = true;
                    builder.line_width_pt = emu_attr_to_pt(element, b"w").or(builder.line_width_pt);
                    builder.line_cap = line_cap(element).or(builder.line_cap);
                }
            }
            b"lnRef" if builder.is_some() => {
                if let Some(index) = attribute_value(element, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|index| *index > 0)
                    && let Some(builder) = builder
                {
                    builder.line_reference_index = Some(index);
                    builder.has_line = true;
                }
            }
            other => {
                handle_geometry_element(builder, other, element, self.sppr_depth, self.line_depth)
            }
        }
    }

    /// Feed an empty element to the builder.
    pub(super) fn empty(&mut self, builder: Option<&mut ShapeBuilder>, element: &BytesStart<'_>) {
        if element.local_name().as_ref() == b"ln" {
            if let Some(builder) = builder {
                builder.has_line = true;
                builder.line_width_pt = emu_attr_to_pt(element, b"w").or(builder.line_width_pt);
                builder.line_cap = line_cap(element).or(builder.line_cap);
            }
            return;
        }
        if element.local_name().as_ref() == b"lnRef" {
            if let Some(index) = attribute_value(element, b"idx")
                .and_then(|value| value.parse::<usize>().ok())
                .filter(|index| *index > 0)
                && let Some(builder) = builder
            {
                builder.line_reference_index = Some(index);
                builder.has_line = true;
            }
            return;
        }
        handle_geometry_element(
            builder,
            element.local_name().as_ref(),
            element,
            self.sppr_depth,
            self.line_depth,
        );
    }

    /// Close the `spPr`/`ln` nesting for an end element.
    pub(super) fn end(&mut self, local_name: &[u8]) {
        match local_name {
            b"spPr" if self.sppr_depth > 0 => self.sppr_depth -= 1,
            b"ln" if self.line_depth > 0 => self.line_depth -= 1,
            _ => {}
        }
    }

    /// Reset at the start of a new top-level `<w:drawing>`.
    pub(super) fn reset(&mut self) {
        self.sppr_depth = 0;
        self.line_depth = 0;
    }
}

/// Apply a geometry/fill/stroke element (`wsp`, `txbx`, `extent`, `prstGeom`,
/// `xfrm`, `srgbClr`, `noFill`, `tailEnd`, `headEnd`) to the current builder.
pub(super) fn handle_geometry_element(
    builder: Option<&mut ShapeBuilder>,
    local_name: &[u8],
    element: &BytesStart<'_>,
    sppr_depth: usize,
    line_depth: usize,
) {
    let Some(builder) = builder else {
        return;
    };

    match local_name {
        b"wsp" => builder.has_wsp = true,
        b"wgp" => builder.has_wpg = true,
        b"txbx" => builder.has_text_box = true,
        // The anchor extent gives the on-page bounding box.
        b"extent" => {
            if let Some(width) = emu_attr_to_pt(element, b"cx") {
                builder.box_width_pt = Some(width);
            }
            if let Some(height) = emu_attr_to_pt(element, b"cy") {
                builder.box_height_pt = Some(height);
            }
        }
        b"prstGeom" => {
            if let Some(preset) = attribute_value(element, b"prst") {
                builder.preset = Some(preset);
            }
        }
        b"xfrm" => {
            builder.flip_h = bool_attr(element, b"flipH");
            builder.flip_v = bool_attr(element, b"flipV");
            builder.rotation_deg = rotation_attr_degrees(element);
        }
        b"srgbClr" if sppr_depth > 0 => {
            if let Some(color) = attribute_value(element, b"val").and_then(|v| parse_hex_color(&v))
            {
                if line_depth > 0 {
                    builder.line_color = builder.line_color.or(Some(color));
                } else {
                    builder.fill_color = builder.fill_color.or(Some(color));
                }
            }
        }
        b"noFill" if sppr_depth > 0 => {
            if line_depth > 0 {
                builder.line_none = true;
            } else {
                builder.fill_none = true;
            }
        }
        // A geometry-only shape has no text to inset, but the same `wps:wsp`
        // element states where an inline text box's paragraphs start (#1690).
        b"bodyPr" => {
            builder.body_inset_left_pt = emu_attr_to_pt(element, b"lIns");
            builder.body_inset_top_pt = emu_attr_to_pt(element, b"tIns");
            builder.body_inset_right_pt = emu_attr_to_pt(element, b"rIns");
            builder.body_inset_bottom_pt = emu_attr_to_pt(element, b"bIns");
        }
        b"tailEnd" if line_depth > 0 => builder.tail_arrow = arrow_type_present(element),
        b"headEnd" if line_depth > 0 => builder.head_arrow = arrow_type_present(element),
        name if line_depth > 0 => {
            builder.line_join = drawingml_line_join(name).or(builder.line_join);
        }
        _ => {}
    }
}

fn arrow_type_present(element: &BytesStart<'_>) -> bool {
    !matches!(attribute_value(element, b"type").as_deref(), Some("none"))
}

#[derive(Default)]
struct CanvasDrawingBuilder {
    record_index: usize,
    is_canvas: bool,
    picture_depth: usize,
    shape_properties_depth: usize,
    transform_depth: usize,
    offset: Option<(f64, f64)>,
}

fn scan_canvas_image_offsets(xml: &str) -> Vec<Option<(f64, f64)>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut records: Vec<Option<(f64, f64)>> = Vec::new();
    let mut drawings: Vec<CanvasDrawingBuilder> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"drawing" => {
                let record_index: usize = records.len();
                records.push(None);
                drawings.push(CanvasDrawingBuilder {
                    record_index,
                    ..CanvasDrawingBuilder::default()
                });
            }
            Ok(Event::Start(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    handle_canvas_start(drawing, element);
                }
            }
            Ok(Event::Empty(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    handle_canvas_element(drawing, element);
                }
            }
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"drawing" => {
                if let Some(drawing) = drawings.pop() {
                    records[drawing.record_index] = drawing.offset;
                }
            }
            Ok(Event::End(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    handle_canvas_end(drawing, element.local_name().as_ref());
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    records
}

fn handle_canvas_start(drawing: &mut CanvasDrawingBuilder, element: &BytesStart<'_>) {
    match element.local_name().as_ref() {
        b"graphicData"
            if attribute_value(element, b"uri")
                .is_some_and(|uri| uri.ends_with("/wordprocessingCanvas")) =>
        {
            drawing.is_canvas = true;
        }
        b"pic" if drawing.is_canvas => drawing.picture_depth += 1,
        b"spPr" if drawing.picture_depth > 0 => drawing.shape_properties_depth += 1,
        b"xfrm" if drawing.shape_properties_depth > 0 => drawing.transform_depth += 1,
        _ => handle_canvas_element(drawing, element),
    }
}

fn handle_canvas_element(drawing: &mut CanvasDrawingBuilder, element: &BytesStart<'_>) {
    if drawing.transform_depth == 0 || element.local_name().as_ref() != b"off" {
        return;
    }
    let Some(x) = numeric_attr(element, b"x") else {
        return;
    };
    let Some(y) = numeric_attr(element, b"y") else {
        return;
    };
    drawing.offset = Some((x / EMU_PER_POINT, y / EMU_PER_POINT));
}

fn handle_canvas_end(drawing: &mut CanvasDrawingBuilder, local_name: &[u8]) {
    match local_name {
        b"xfrm" if drawing.transform_depth > 0 => drawing.transform_depth -= 1,
        b"spPr" if drawing.shape_properties_depth > 0 => drawing.shape_properties_depth -= 1,
        b"pic" if drawing.picture_depth > 0 => drawing.picture_depth -= 1,
        _ => {}
    }
}

#[derive(Debug, Clone, Copy)]
struct AffineTransform {
    scale_x: f64,
    scale_y: f64,
    translate_x: f64,
    translate_y: f64,
}

impl Default for AffineTransform {
    fn default() -> Self {
        Self {
            scale_x: 1.0,
            scale_y: 1.0,
            translate_x: 0.0,
            translate_y: 0.0,
        }
    }
}

impl AffineTransform {
    fn compose(self, child: Self) -> Self {
        Self {
            scale_x: self.scale_x * child.scale_x,
            scale_y: self.scale_y * child.scale_y,
            translate_x: self.translate_x + child.translate_x * self.scale_x,
            translate_y: self.translate_y + child.translate_y * self.scale_y,
        }
    }

    fn point(self, x: f64, y: f64) -> (f64, f64) {
        (
            self.translate_x + x * self.scale_x,
            self.translate_y + y * self.scale_y,
        )
    }
}

#[derive(Default)]
struct GroupTransformBuilder {
    offset_x: f64,
    offset_y: f64,
    extent_x: f64,
    extent_y: f64,
    child_offset_x: f64,
    child_offset_y: f64,
    child_extent_x: f64,
    child_extent_y: f64,
}

impl GroupTransformBuilder {
    fn finish(&self) -> AffineTransform {
        let scale_x: f64 = if self.child_extent_x.abs() > f64::EPSILON {
            self.extent_x / self.child_extent_x
        } else {
            1.0
        };
        let scale_y: f64 = if self.child_extent_y.abs() > f64::EPSILON {
            self.extent_y / self.child_extent_y
        } else {
            1.0
        };
        AffineTransform {
            scale_x,
            scale_y,
            translate_x: self.offset_x - self.child_offset_x * scale_x,
            translate_y: self.offset_y - self.child_offset_y * scale_y,
        }
    }
}

#[derive(Default)]
struct WpgChildBuilder {
    shape: ShapeBuilder,
    parent_transform: AffineTransform,
    offset_x: f64,
    offset_y: f64,
    extent_x: f64,
    extent_y: f64,
    content: Vec<docx_rs::DocumentChild>,
    padding: Insets,
    vertical_align: TextBoxVerticalAlign,
    text_color: Option<Color>,
    shape_properties_depth: usize,
    shape_transform_depth: usize,
    line_depth: usize,
    fill_reference_depth: usize,
    line_reference_depth: usize,
    font_reference_depth: usize,
}

struct WpgDrawingBuilder {
    record_index: usize,
    is_wpg: bool,
    anchor_offset_x: f64,
    anchor_offset_y: f64,
    position_axis: PositionAxis,
    in_position_offset: bool,
    wrap_mode: WrapMode,
    group_transforms: Vec<AffineTransform>,
    group_transform_builder: Option<GroupTransformBuilder>,
    group_properties_depth: usize,
    group_transform_depth: usize,
    child: Option<WpgChildBuilder>,
    children: Vec<WpgShapeInfo>,
}

impl WpgDrawingBuilder {
    fn new(record_index: usize) -> Self {
        Self {
            record_index,
            is_wpg: false,
            anchor_offset_x: 0.0,
            anchor_offset_y: 0.0,
            position_axis: PositionAxis::None,
            in_position_offset: false,
            wrap_mode: WrapMode::None,
            group_transforms: Vec::new(),
            group_transform_builder: None,
            group_properties_depth: 0,
            group_transform_depth: 0,
            child: None,
            children: Vec::new(),
        }
    }

    fn finish_child(&mut self, theme_line_styles: &[ThemeLineStyle]) {
        let Some(mut child) = self.child.take() else {
            return;
        };
        let rotation_deg: Option<f64> = child.shape.rotation_deg;
        let (local_x, local_y) = child.parent_transform.point(child.offset_x, child.offset_y);
        let width_emu: f64 = child.extent_x * child.parent_transform.scale_x.abs();
        let height_emu: f64 = child.extent_y * child.parent_transform.scale_y.abs();
        let width: f64 = width_emu / EMU_PER_POINT;
        let height: f64 = height_emu / EMU_PER_POINT;
        let offset_x: f64 = self.anchor_offset_x + local_x / EMU_PER_POINT;
        let offset_y: f64 = self.anchor_offset_y + local_y / EMU_PER_POINT;

        child.shape.box_width_pt = Some(width);
        child.shape.box_height_pt = Some(height);
        child.shape.offset_x_pt = offset_x;
        child.shape.offset_y_pt = offset_y;
        let shape: Option<FloatingShape> = child.shape.finish_wpg(theme_line_styles);
        let padding: Insets = shape
            .as_ref()
            .map(|shape| shape_text_padding(&shape.shape.kind, width, height, child.padding))
            .unwrap_or(child.padding);

        if shape.is_some() || !child.content.is_empty() {
            self.children.push(WpgShapeInfo {
                shape,
                content: child.content,
                width,
                height,
                rotation_deg,
                padding,
                vertical_align: child.vertical_align,
                text_color: child.text_color,
                offset_x,
                offset_y,
                wrap_mode: self.wrap_mode,
            });
        }
    }
}

fn shape_text_padding(kind: &ShapeKind, width: f64, height: f64, body: Insets) -> Insets {
    let (horizontal_fraction, top_fraction, bottom_fraction): (f64, f64, f64) = match kind {
        ShapeKind::Ellipse => {
            let circle_inset: f64 = (1.0 - std::f64::consts::FRAC_1_SQRT_2) / 2.0;
            (circle_inset, circle_inset, circle_inset)
        }
        ShapeKind::Polygon { vertices }
            if vertices.as_slice() == [(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)] =>
        {
            (0.25, 0.5, 0.0)
        }
        ShapeKind::Polygon { vertices }
            if vertices.as_slice() == [(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)] =>
        {
            (0.25, 0.25, 0.25)
        }
        _ => (0.0, 0.0, 0.0),
    };

    Insets {
        left: body.left + width * horizontal_fraction,
        right: body.right + width * horizontal_fraction,
        top: body.top + height * top_fraction,
        bottom: body.bottom + height * bottom_fraction,
    }
}

fn numeric_attr(element: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attribute_value(element, name).and_then(|value| value.parse::<f64>().ok())
}

#[cfg(test)]
fn scan_wpg_drawings(xml: &str, theme_xml: Option<&str>) -> Vec<Option<WpgDrawingInfo>> {
    scan_wpg_drawings_with_theme(xml, theme_xml, &[])
}

fn scan_wpg_drawings_with_theme(
    xml: &str,
    theme_xml: Option<&str>,
    theme_line_styles: &[ThemeLineStyle],
) -> Vec<Option<WpgDrawingInfo>> {
    let text_box_contents: Vec<Vec<docx_rs::DocumentChild>> = scan_wpg_text_box_contents(xml);
    let mut text_box_cursor: usize = 0;
    let theme_colors: HashMap<String, Color> =
        parse_theme_color_scheme(theme_xml.unwrap_or_default());
    let theme_aliases: HashMap<String, String> = HashMap::new();
    let scheme = SchemeColors {
        colors: &theme_colors,
        aliases: &theme_aliases,
    };
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut records: Vec<Option<WpgDrawingInfo>> = Vec::new();
    let mut drawings: Vec<WpgDrawingBuilder> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"drawing" => {
                let record_index: usize = records.len();
                records.push(None);
                drawings.push(WpgDrawingBuilder::new(record_index));
            }
            Ok(Event::Start(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    let consumed = match element.local_name().as_ref() {
                        b"custGeom" => consume_wpg_custom_geometry(drawing, &mut reader),
                        b"gradFill" => {
                            consume_wpg_gradient_fill(drawing, &mut reader, &theme_colors)
                        }
                        b"srgbClr" | b"schemeClr" | b"sysClr" => {
                            consume_wpg_color(drawing, &mut reader, element, &scheme)
                        }
                        _ => false,
                    };
                    if !consumed {
                        handle_wpg_start(
                            drawing,
                            element,
                            &text_box_contents,
                            &mut text_box_cursor,
                            theme_line_styles,
                        );
                    }
                }
            }
            Ok(Event::Empty(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    if matches!(
                        element.local_name().as_ref(),
                        b"srgbClr" | b"schemeClr" | b"sysClr"
                    ) && wpg_color_context(drawing)
                    {
                        apply_wpg_color(drawing, parse_color_from_empty(element, &scheme));
                    } else {
                        handle_wpg_empty(drawing, element);
                    }
                }
            }
            Ok(Event::Text(ref text)) => {
                if let Some(drawing) = drawings.last_mut()
                    && drawing.in_position_offset
                    && let Ok(raw) = text.xml_content(OOXML_XML_VERSION)
                    && let Ok(emu) = raw.trim().parse::<f64>()
                {
                    match drawing.position_axis {
                        PositionAxis::Horizontal => drawing.anchor_offset_x = emu / EMU_PER_POINT,
                        PositionAxis::Vertical => drawing.anchor_offset_y = emu / EMU_PER_POINT,
                        PositionAxis::None => {}
                    }
                }
            }
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"drawing" => {
                if let Some(mut drawing) = drawings.pop() {
                    drawing.finish_child(theme_line_styles);
                    if drawing.is_wpg {
                        records[drawing.record_index] = Some(WpgDrawingInfo {
                            children: drawing.children,
                        });
                    }
                }
            }
            Ok(Event::End(ref element)) => {
                if let Some(drawing) = drawings.last_mut() {
                    handle_wpg_end(drawing, element.local_name().as_ref(), theme_line_styles);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    records
}

fn consume_wpg_custom_geometry(
    drawing: &mut WpgDrawingBuilder,
    reader: &mut quick_xml::Reader<&[u8]>,
) -> bool {
    let Some(child) = drawing.child.as_mut() else {
        return false;
    };
    if child.shape_properties_depth == 0 {
        return false;
    }

    let extent = ShapeExtent::new(child.extent_x, child.extent_y);
    let subpaths = parse_custom_geometry(reader, extent);
    if !subpaths.is_empty() {
        child.shape.custom_subpaths = subpaths;
    }
    true
}

fn consume_wpg_gradient_fill(
    drawing: &mut WpgDrawingBuilder,
    reader: &mut quick_xml::Reader<&[u8]>,
    theme_colors: &HashMap<String, Color>,
) -> bool {
    let Some(child) = drawing.child.as_mut() else {
        return false;
    };
    if child.shape_properties_depth == 0 || child.line_depth > 0 {
        return false;
    }

    child.shape.gradient_fill = parse_docx_shape_gradient(reader, theme_colors);
    true
}

fn consume_wpg_color(
    drawing: &mut WpgDrawingBuilder,
    reader: &mut quick_xml::Reader<&[u8]>,
    element: &BytesStart<'_>,
    scheme: &SchemeColors<'_>,
) -> bool {
    if !wpg_color_context(drawing) {
        return false;
    }
    let parsed = parse_color_from_start(reader, element, scheme);
    apply_wpg_color(drawing, parsed);
    true
}

fn wpg_color_context(drawing: &WpgDrawingBuilder) -> bool {
    drawing.child.as_ref().is_some_and(|child| {
        child.font_reference_depth > 0
            || child.line_reference_depth > 0
            || child.line_depth > 0
            || child.fill_reference_depth > 0
            || child.shape_properties_depth > 0
    })
}

fn apply_wpg_color(drawing: &mut WpgDrawingBuilder, parsed: ParsedColor) {
    let Some(child) = drawing.child.as_mut() else {
        return;
    };
    let Some(color) = parsed.color else {
        return;
    };

    if child.font_reference_depth > 0 {
        child.text_color = Some(color);
    } else if child.line_reference_depth > 0 || child.line_depth > 0 {
        child.shape.line_color.get_or_insert(color);
        child.shape.has_line = true;
    } else if (child.fill_reference_depth > 0 || child.shape_properties_depth > 0)
        && child.shape.fill_color.is_none()
    {
        child.shape.fill_color = Some(color);
        child.shape.fill_opacity = parsed.alpha;
    }
}

fn handle_wpg_start(
    drawing: &mut WpgDrawingBuilder,
    element: &BytesStart<'_>,
    text_box_contents: &[Vec<docx_rs::DocumentChild>],
    text_box_cursor: &mut usize,
    theme_line_styles: &[ThemeLineStyle],
) {
    match element.local_name().as_ref() {
        b"positionH" => drawing.position_axis = PositionAxis::Horizontal,
        b"positionV" => drawing.position_axis = PositionAxis::Vertical,
        b"posOffset" => drawing.in_position_offset = true,
        b"wrapSquare" => drawing.wrap_mode = WrapMode::Square,
        b"wrapTight" => drawing.wrap_mode = WrapMode::Tight,
        b"wrapTopAndBottom" => drawing.wrap_mode = WrapMode::TopAndBottom,
        b"wgp" => {
            drawing.is_wpg = true;
            drawing.group_transforms.push(AffineTransform::default());
        }
        b"grpSp" if drawing.is_wpg => {
            let inherited: AffineTransform =
                drawing.group_transforms.last().copied().unwrap_or_default();
            drawing.group_transforms.push(inherited);
        }
        b"grpSpPr" if drawing.is_wpg => {
            drawing.group_properties_depth += 1;
            drawing.group_transform_builder = Some(GroupTransformBuilder::default());
        }
        b"wsp" if drawing.is_wpg => {
            drawing.finish_child(theme_line_styles);
            let mut child = WpgChildBuilder {
                parent_transform: drawing.group_transforms.last().copied().unwrap_or_default(),
                ..WpgChildBuilder::default()
            };
            child.shape.has_wsp = true;
            drawing.child = Some(child);
        }
        b"spPr" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.shape_properties_depth += 1;
            }
        }
        b"xfrm" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut()
                && child.shape_properties_depth > 0
            {
                child.shape_transform_depth += 1;
                child.shape.flip_h = bool_attr(element, b"flipH");
                child.shape.flip_v = bool_attr(element, b"flipV");
                child.shape.rotation_deg = rotation_attr_degrees(element);
            }
        }
        b"xfrm" if drawing.group_properties_depth > 0 => {
            drawing.group_transform_depth += 1;
        }
        b"ln" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.line_depth += 1;
                child.shape.has_line = true;
                child.shape.line_width_pt =
                    emu_attr_to_pt(element, b"w").or(child.shape.line_width_pt);
                child.shape.line_cap = line_cap(element).or(child.shape.line_cap);
            }
        }
        b"fillRef" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.fill_reference_depth += 1;
            }
        }
        b"lnRef" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.line_reference_depth += 1;
                if numeric_attr(element, b"idx").unwrap_or_default() > 0.0 {
                    child.shape.has_line = true;
                    child.shape.line_reference_index = attribute_value(element, b"idx")
                        .and_then(|value| value.parse::<usize>().ok())
                        .filter(|index| *index > 0);
                }
            }
        }
        b"fontRef" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.font_reference_depth += 1;
            }
        }
        b"txbx" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.shape.has_text_box = true;
                child.content = text_box_contents
                    .get(*text_box_cursor)
                    .cloned()
                    .unwrap_or_default();
                *text_box_cursor += 1;
            }
        }
        b"bodyPr" if drawing.child.is_some() => {
            if let Some(child) = drawing.child.as_mut() {
                child.vertical_align = match attribute_value(element, b"anchor").as_deref() {
                    Some("ctr") => TextBoxVerticalAlign::Center,
                    Some("b") => TextBoxVerticalAlign::Bottom,
                    _ => TextBoxVerticalAlign::Top,
                };
                child.padding = Insets {
                    left: emu_attr_to_pt(element, b"lIns").unwrap_or_default(),
                    top: emu_attr_to_pt(element, b"tIns").unwrap_or_default(),
                    right: emu_attr_to_pt(element, b"rIns").unwrap_or_default(),
                    bottom: emu_attr_to_pt(element, b"bIns").unwrap_or_default(),
                };
            }
        }
        _ => handle_wpg_geometry_element(drawing, element),
    }
}

fn handle_wpg_empty(drawing: &mut WpgDrawingBuilder, element: &BytesStart<'_>) {
    match element.local_name().as_ref() {
        b"ln" => {
            if let Some(child) = drawing.child.as_mut() {
                child.shape.has_line = true;
                child.shape.line_width_pt =
                    emu_attr_to_pt(element, b"w").or(child.shape.line_width_pt);
                child.shape.line_cap = line_cap(element).or(child.shape.line_cap);
            }
            return;
        }
        b"lnRef" => {
            if let Some(child) = drawing.child.as_mut()
                && let Some(index) = attribute_value(element, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|index| *index > 0)
            {
                child.shape.has_line = true;
                child.shape.line_reference_index = Some(index);
            }
            return;
        }
        _ => {}
    }
    handle_wpg_geometry_element(drawing, element);
}

fn handle_wpg_geometry_element(drawing: &mut WpgDrawingBuilder, element: &BytesStart<'_>) {
    let element_name = element.local_name();
    let local_name: &[u8] = element_name.as_ref();
    if drawing.group_transform_depth > 0 {
        if let Some(group) = drawing.group_transform_builder.as_mut() {
            match local_name {
                b"off" => {
                    group.offset_x = numeric_attr(element, b"x").unwrap_or_default();
                    group.offset_y = numeric_attr(element, b"y").unwrap_or_default();
                }
                b"ext" => {
                    group.extent_x = numeric_attr(element, b"cx").unwrap_or_default();
                    group.extent_y = numeric_attr(element, b"cy").unwrap_or_default();
                }
                b"chOff" => {
                    group.child_offset_x = numeric_attr(element, b"x").unwrap_or_default();
                    group.child_offset_y = numeric_attr(element, b"y").unwrap_or_default();
                }
                b"chExt" => {
                    group.child_extent_x = numeric_attr(element, b"cx").unwrap_or_default();
                    group.child_extent_y = numeric_attr(element, b"cy").unwrap_or_default();
                }
                _ => {}
            }
        }
        return;
    }

    let Some(child) = drawing.child.as_mut() else {
        return;
    };
    if child.shape_transform_depth > 0 {
        match local_name {
            b"off" => {
                child.offset_x = numeric_attr(element, b"x").unwrap_or_default();
                child.offset_y = numeric_attr(element, b"y").unwrap_or_default();
            }
            b"ext" => {
                child.extent_x = numeric_attr(element, b"cx").unwrap_or_default();
                child.extent_y = numeric_attr(element, b"cy").unwrap_or_default();
            }
            _ => {}
        }
    }

    handle_geometry_element(
        Some(&mut child.shape),
        local_name,
        element,
        child.shape_properties_depth,
        child.line_depth,
    );
}

fn handle_wpg_end(
    drawing: &mut WpgDrawingBuilder,
    local_name: &[u8],
    theme_line_styles: &[ThemeLineStyle],
) {
    match local_name {
        b"posOffset" => drawing.in_position_offset = false,
        b"positionH" | b"positionV" => drawing.position_axis = PositionAxis::None,
        b"xfrm"
            if drawing
                .child
                .as_ref()
                .is_some_and(|child| child.shape_transform_depth > 0) =>
        {
            if let Some(child) = drawing.child.as_mut() {
                child.shape_transform_depth -= 1;
            }
        }
        b"xfrm" if drawing.group_transform_depth > 0 => {
            drawing.group_transform_depth -= 1;
            if drawing.group_transform_depth == 0
                && let Some(group) = drawing.group_transform_builder.take()
            {
                let parent: AffineTransform = if drawing.group_transforms.len() > 1 {
                    drawing.group_transforms[drawing.group_transforms.len() - 2]
                } else {
                    AffineTransform::default()
                };
                if let Some(current) = drawing.group_transforms.last_mut() {
                    *current = parent.compose(group.finish());
                }
            }
        }
        b"spPr" => {
            if let Some(child) = drawing.child.as_mut()
                && child.shape_properties_depth > 0
            {
                child.shape_properties_depth -= 1;
            }
        }
        b"ln" => {
            if let Some(child) = drawing.child.as_mut()
                && child.line_depth > 0
            {
                child.line_depth -= 1;
            }
        }
        b"fillRef" => {
            if let Some(child) = drawing.child.as_mut()
                && child.fill_reference_depth > 0
            {
                child.fill_reference_depth -= 1;
            }
        }
        b"lnRef" => {
            if let Some(child) = drawing.child.as_mut()
                && child.line_reference_depth > 0
            {
                child.line_reference_depth -= 1;
            }
        }
        b"fontRef" => {
            if let Some(child) = drawing.child.as_mut()
                && child.font_reference_depth > 0
            {
                child.font_reference_depth -= 1;
            }
        }
        b"wsp" => drawing.finish_child(theme_line_styles),
        b"grpSpPr" if drawing.group_properties_depth > 0 => {
            drawing.group_properties_depth -= 1;
        }
        b"grpSp" | b"wgp" if drawing.is_wpg => {
            drawing.group_transforms.pop();
        }
        _ => {}
    }
}

fn scan_wpg_text_box_contents(xml: &str) -> Vec<Vec<docx_rs::DocumentChild>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<Vec<docx_rs::DocumentChild>> = Vec::new();
    let mut wpg_depth: usize = 0;
    let mut capture: Option<(quick_xml::Writer<Vec<u8>>, usize)> = None;

    while let Ok(event) = reader.read_event_into(&mut buffer) {
        if let Some((writer, depth)) = capture.as_mut() {
            match &event {
                Event::Start(element) => {
                    *depth += usize::from(element.local_name().as_ref() == b"txbxContent");
                    let _ = writer.write_event(event.into_owned());
                }
                Event::End(element)
                    if element.local_name().as_ref() == b"txbxContent" && *depth == 1 =>
                {
                    let (writer, _) = capture.take().expect("capture should exist");
                    result.push(parse_wpg_text_box_document(writer.into_inner()));
                }
                Event::End(element) => {
                    if element.local_name().as_ref() == b"txbxContent" {
                        *depth -= 1;
                    }
                    let _ = writer.write_event(event.into_owned());
                }
                Event::Eof => break,
                _ => {
                    let _ = writer.write_event(event.into_owned());
                }
            }
            buffer.clear();
            continue;
        }

        match &event {
            Event::Start(element) if element.local_name().as_ref() == b"wgp" => wpg_depth += 1,
            Event::End(element) if element.local_name().as_ref() == b"wgp" && wpg_depth > 0 => {
                wpg_depth -= 1;
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"txbxContent" && wpg_depth > 0 =>
            {
                capture = Some((quick_xml::Writer::new(Vec::new()), 1));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    result
}

fn parse_wpg_text_box_document(inner_xml: Vec<u8>) -> Vec<docx_rs::DocumentChild> {
    let Ok(inner) = String::from_utf8(inner_xml) else {
        return Vec::new();
    };
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{inner}</w:body></w:document>"#
    );
    docx_rs::Document::from_xml(xml.as_bytes())
        .map(|document| {
            document
                .children
                .into_iter()
                .filter(|child| {
                    matches!(
                        child,
                        docx_rs::DocumentChild::Paragraph(_) | docx_rs::DocumentChild::Table(_)
                    )
                })
                .collect()
        })
        .unwrap_or_default()
}

#[cfg(test)]
#[path = "docx_context_shape_tests.rs"]
mod tests;
