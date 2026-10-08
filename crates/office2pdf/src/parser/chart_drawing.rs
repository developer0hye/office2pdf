//! The drawing part a chart's `<c:userShapes>` names.
//!
//! A chart carries a drawing layer of its own: `c:chartSpace/c:userShapes`
//! points at a part whose `cdr:` anchors each hold an ordinary DrawingML
//! shape, and Office draws those over the finished chart. Nothing followed the
//! relationship, so the `CASH FLOW` caption of
//! `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx` never reached the page
//! (issue #1186).
//!
//! The anchors are chart-relative: `cdr:from`/`cdr:to` are fractions of the
//! chart area rather than the cell coordinates a worksheet drawing anchors by,
//! which is why this is its own reader rather than a mode of
//! `xlsx_drawing::parse_drawing_text_boxes`. Everything below the anchor is
//! plain `a:` DrawingML and reuses that module's readers.
//!
//! TODO(no corpus file anchors one): an anchor may also hold `cdr:pic`,
//! `cdr:grpSp`, `cdr:graphicFrame` or `cdr:cxnSp`. Only `cdr:sp` is read, so
//! one of the others yields no shape rather than a wrong one.

use crate::parser::xml_util::OOXML_XML_VERSION;
use std::io::{Read, Seek};

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use super::drawingml::{self, SchemeColors, ThemeFontScheme};
use super::pptx::ThemeLineStyle;
use super::xlsx::xlsx_drawing::{apply_run_properties, parse_rels_targets, resolved_or_legacy};
use super::xml_util;
use crate::ir::{
    Alignment, BorderLineStyle, BorderSide, ChartUserShape, ChartUserShapeExtent, Color, Insets,
    LineCap, LineJoin, Paragraph, ParagraphStyle, Run, TextStyle,
};

/// EMU per point.
const EMU_PER_POINT: f64 = 12_700.0;

/// `<a:bodyPr>`'s default insets, in points: ECMA-376 gives `lIns`/`rIns`
/// 91440 EMU (0.1in) and `tIns`/`bIns` 45720 EMU (0.05in). A native Excel for
/// Mac 16 export of the reported workbook seats the caption's pen exactly
/// 7.2pt right of its box, and rewriting `lIns="0"` moves it left by exactly
/// that, so the horizontal default is the one the schema states rather than
/// the flat inset a worksheet text box is drawn with.
const DEFAULT_SIDE_INSET_PT: f64 = 7.2;
const DEFAULT_VERTICAL_INSET_PT: f64 = 3.6;

/// The relationship id `<c:userShapes>` names, if the chart part names one.
///
/// Read from the chart rather than from its `.rels`, because a package may
/// declare a `chartUserShapes` relationship the chart itself never references,
/// and Office draws only what the part points at — the mirror of issue #1158,
/// where a drawing relationship no sheet element named still printed.
pub(crate) fn user_shapes_rid(chart_xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(chart_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"userShapes" =>
            {
                return xml_util::get_attr_str(e, b"id");
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// Read the shapes a chart lays over itself, from whichever package holds it.
///
/// `chart_path` is the chart part's own path in the archive, which the
/// relationship target is resolved against; a chart naming no user shapes, or
/// one whose part is missing, yields none.
pub(crate) fn load_chart_user_shapes<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    chart_path: &str,
    chart_xml: &str,
    scheme: &SchemeColors<'_>,
    theme_fonts: &ThemeFontScheme,
    theme_line_styles: &[ThemeLineStyle],
) -> Vec<ChartUserShape> {
    let Some(rid) = user_shapes_rid(chart_xml) else {
        return Vec::new();
    };
    let (chart_dir, chart_file) = match chart_path.rsplit_once('/') {
        Some((dir, file)) => (dir, file),
        None => ("", chart_path),
    };
    let rels_path: String = format!("{chart_dir}/_rels/{chart_file}.rels");
    let Some(rels_xml) = read_part(archive, &rels_path) else {
        return Vec::new();
    };
    let Some(target) = parse_rels_targets(&rels_xml).get(&rid).cloned() else {
        return Vec::new();
    };
    let drawing_path: String = resolve_relative_part(chart_dir, &target);
    let Some(drawing_xml) = read_part(archive, &drawing_path) else {
        return Vec::new();
    };
    parse_chart_user_shapes_with_theme_lines(&drawing_xml, scheme, theme_fonts, theme_line_styles)
}

fn read_part<R: Read + Seek>(archive: &mut ZipArchive<R>, path: &str) -> Option<String> {
    let mut entry = archive.by_name(path).ok()?;
    let mut text = String::new();
    entry.read_to_string(&mut text).ok()?;
    Some(text)
}

/// Resolve a relationship target against the part directory that declared it,
/// honouring both the `../` hops Office writes and an absolute package path.
fn resolve_relative_part(base_dir: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut segments: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            name => segments.push(name),
        }
    }
    segments.join("/")
}

/// Parse a chart drawing part into the shapes it anchors over the chart area.
#[cfg(test)]
pub(crate) fn parse_chart_user_shapes(
    xml: &str,
    scheme: &SchemeColors<'_>,
    theme_fonts: &ThemeFontScheme,
) -> Vec<ChartUserShape> {
    parse_chart_user_shapes_with_theme_lines(xml, scheme, theme_fonts, &[])
}

fn parse_chart_user_shapes_with_theme_lines(
    xml: &str,
    scheme: &SchemeColors<'_>,
    theme_fonts: &ThemeFontScheme,
    theme_line_styles: &[ThemeLineStyle],
) -> Vec<ChartUserShape> {
    let mut shapes: Vec<ChartUserShape> = Vec::new();
    let mut reader = Reader::from_str(xml);
    let mut anchor: Option<AnchorState> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"relSizeAnchor" | b"absSizeAnchor" => anchor = Some(AnchorState::default()),
                    _ => {
                        if let Some(state) = anchor.as_mut() {
                            state.start(local.as_ref(), e, &mut reader, scheme);
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(state) = anchor.as_mut() {
                    state.empty(e.local_name().as_ref(), e, scheme, theme_fonts);
                }
            }
            Ok(Event::Text(ref t)) => {
                if let Some(state) = anchor.as_mut()
                    && let Ok(text) = t.xml_content(OOXML_XML_VERSION)
                {
                    state.text(&text, theme_fonts);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"relSizeAnchor" | b"absSizeAnchor" => {
                        if let Some(state) = anchor.take()
                            && let Some(shape) = state.finish(theme_line_styles)
                        {
                            shapes.push(shape);
                        }
                    }
                    _ => {
                        if let Some(state) = anchor.as_mut() {
                            state.end(local.as_ref());
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    shapes
}

/// Which of the anchor's own coordinate elements the reader is inside.
#[derive(Clone, Copy, PartialEq)]
enum Corner {
    From,
    To,
}

/// One `cdr:` anchor as it is read.
#[derive(Default)]
struct AnchorState {
    corner: Option<Corner>,
    axis: Option<&'static str>,
    from: (f64, f64),
    to: Option<(f64, f64)>,
    ext_pt: Option<(f64, f64)>,
    saw_shape: bool,
    in_tx_body: bool,
    in_shape_fill: bool,
    in_line: bool,
    in_run: bool,
    in_text: bool,
    insets: Option<Insets>,
    no_wrap: bool,
    fill: Option<Color>,
    border_color: Option<Color>,
    border_width_pt: f64,
    border_cap: Option<LineCap>,
    border_join: Option<LineJoin>,
    line_reference_index: Option<usize>,
    paragraphs: Vec<Paragraph>,
    current_paragraph: Option<Paragraph>,
    current_style: TextStyle,
}

impl AnchorState {
    fn start(
        &mut self,
        tag: &[u8],
        element: &quick_xml::events::BytesStart<'_>,
        reader: &mut Reader<&[u8]>,
        scheme: &SchemeColors<'_>,
    ) {
        match tag {
            b"from" => self.corner = Some(Corner::From),
            b"to" => {
                self.corner = Some(Corner::To);
                self.to = Some((0.0, 0.0));
            }
            b"x" if self.corner.is_some() => self.axis = Some("x"),
            b"y" if self.corner.is_some() => self.axis = Some("y"),
            b"sp" => {
                self.saw_shape = true;
                self.border_width_pt = DEFAULT_BORDER_WIDTH_PT;
            }
            b"txBody" => self.in_tx_body = true,
            b"bodyPr" => self.read_body_insets(element),
            b"solidFill" if !self.in_tx_body && !self.in_line => self.in_shape_fill = true,
            b"ln" if !self.in_tx_body => {
                self.in_line = true;
                self.border_cap = drawingml::line_cap(element).or(self.border_cap);
                if let Some(width) = xml_util::get_attr_str(element, b"w")
                    .and_then(|value| value.parse::<f64>().ok())
                {
                    self.border_width_pt = width / EMU_PER_POINT;
                }
            }
            b"lnRef" => {
                self.line_reference_index = xml_util::get_attr_str(element, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|index| *index > 0);
            }
            b"round" | b"bevel" | b"miter" if self.in_line => {
                self.border_join = super::pptx::drawingml_line_join(tag).or(self.border_join);
            }
            b"p" if self.in_tx_body => {
                self.current_paragraph = Some(Paragraph {
                    // As in a worksheet text box: DrawingML paragraphs stack
                    // with no gap of their own unless the body asks for one,
                    // and leaving these unset lets the renderer's default
                    // block spacing in on top of the line height (issue #656).
                    style: ParagraphStyle {
                        space_before: Some(0.0),
                        space_after: Some(0.0),
                        ..ParagraphStyle::default()
                    },
                    runs: Vec::new(),
                });
            }
            b"pPr" if self.current_paragraph.is_some() => self.read_alignment(element),
            b"r" if self.current_paragraph.is_some() => {
                self.in_run = true;
                self.current_style = TextStyle::default();
            }
            b"rPr" if self.in_run => apply_run_properties(&mut self.current_style, element),
            b"t" if self.in_run => self.in_text = true,
            b"srgbClr" | b"schemeClr" => {
                let parsed = drawingml::parse_color_from_start(reader, element, scheme).color;
                self.apply_color(resolved_or_legacy(parsed, tag, element));
            }
            _ => {}
        }
    }

    fn empty(
        &mut self,
        tag: &[u8],
        element: &quick_xml::events::BytesStart<'_>,
        scheme: &SchemeColors<'_>,
        theme_fonts: &ThemeFontScheme,
    ) {
        match tag {
            b"ln" if !self.in_tx_body => {
                self.border_cap = drawingml::line_cap(element).or(self.border_cap);
                if let Some(width) = xml_util::get_attr_str(element, b"w")
                    .and_then(|value| value.parse::<f64>().ok())
                {
                    self.border_width_pt = width / EMU_PER_POINT;
                }
            }
            b"lnRef" => {
                self.line_reference_index = xml_util::get_attr_str(element, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|index| *index > 0);
            }
            b"round" | b"bevel" | b"miter" if self.in_line => {
                self.border_join = super::pptx::drawingml_line_join(tag).or(self.border_join);
            }
            b"bodyPr" => self.read_body_insets(element),
            b"pPr" if self.current_paragraph.is_some() => self.read_alignment(element),
            b"rPr" if self.in_run => apply_run_properties(&mut self.current_style, element),
            b"latin" if self.in_run => {
                if let Some(typeface) = xml_util::get_attr_str(element, b"typeface")
                    && let Some(family) = theme_fonts.resolve_typeface(&typeface)
                {
                    self.current_style.font_family = Some(family);
                }
            }
            b"srgbClr" | b"schemeClr" => {
                let parsed = drawingml::parse_color_from_empty(element, scheme).color;
                self.apply_color(resolved_or_legacy(parsed, tag, element));
            }
            b"ext" if !self.saw_shape => {
                let emu = |name: &[u8]| -> f64 {
                    xml_util::get_attr_str(element, name)
                        .and_then(|value| value.parse::<f64>().ok())
                        .unwrap_or(0.0)
                };
                self.ext_pt = Some((emu(b"cx") / EMU_PER_POINT, emu(b"cy") / EMU_PER_POINT));
            }
            _ => {}
        }
    }

    fn text(&mut self, text: &str, theme_fonts: &ThemeFontScheme) {
        if self.in_text {
            if let Some(paragraph) = self.current_paragraph.as_mut() {
                let mut style: TextStyle = self.current_style.clone();
                // An `<a:rPr>` naming no `<a:latin>` resolves to the theme's
                // minor Latin font in DrawingML, not to a renderer default
                // (issue #461).
                if style.font_family.is_none() {
                    style.font_family = theme_fonts.minor_latin.clone();
                }
                paragraph.runs.push(Run {
                    text: text.to_string(),
                    style,
                    href: None,
                    footnote: None,
                    inline_box: None,
                });
            }
            return;
        }
        let (Some(corner), Some(axis)) = (self.corner, self.axis) else {
            return;
        };
        let Ok(fraction) = text.trim().parse::<f64>() else {
            return;
        };
        let target: &mut (f64, f64) = match corner {
            Corner::From => &mut self.from,
            Corner::To => self.to.get_or_insert((0.0, 0.0)),
        };
        match axis {
            "x" => target.0 = fraction,
            _ => target.1 = fraction,
        }
    }

    fn end(&mut self, tag: &[u8]) {
        match tag {
            b"from" | b"to" => self.corner = None,
            b"x" | b"y" => self.axis = None,
            b"txBody" => self.in_tx_body = false,
            b"solidFill" => self.in_shape_fill = false,
            b"ln" => self.in_line = false,
            b"p" => {
                if let Some(paragraph) = self.current_paragraph.take() {
                    self.paragraphs.push(paragraph);
                }
            }
            b"r" => self.in_run = false,
            b"t" => self.in_text = false,
            _ => {}
        }
    }

    fn read_body_insets(&mut self, element: &quick_xml::events::BytesStart<'_>) {
        let inset = |name: &[u8], default: f64| -> f64 {
            xml_util::get_attr_str(element, name)
                .and_then(|value| value.parse::<f64>().ok())
                .map(|emu| emu / EMU_PER_POINT)
                .unwrap_or(default)
        };
        self.no_wrap = xml_util::get_attr_str(element, b"wrap").as_deref() == Some("none");
        self.insets = Some(Insets {
            top: inset(b"tIns", DEFAULT_VERTICAL_INSET_PT),
            right: inset(b"rIns", DEFAULT_SIDE_INSET_PT),
            bottom: inset(b"bIns", DEFAULT_VERTICAL_INSET_PT),
            left: inset(b"lIns", DEFAULT_SIDE_INSET_PT),
        });
    }

    fn read_alignment(&mut self, element: &quick_xml::events::BytesStart<'_>) {
        let Some(paragraph) = self.current_paragraph.as_mut() else {
            return;
        };
        if let Some(algn) = xml_util::get_attr_str(element, b"algn") {
            paragraph.style.alignment = match algn.as_str() {
                "ctr" => Some(Alignment::Center),
                "r" => Some(Alignment::Right),
                "just" => Some(Alignment::Justify),
                _ => None,
            };
        }
    }

    fn apply_color(&mut self, color: Option<Color>) {
        if self.in_run {
            if self.current_style.color.is_none() {
                self.current_style.color = color;
            }
        } else if self.in_line {
            if self.border_color.is_none() {
                self.border_color = color;
            }
        } else if self.in_shape_fill && self.fill.is_none() {
            self.fill = color;
        }
    }

    fn finish(self, theme_line_styles: &[ThemeLineStyle]) -> Option<ChartUserShape> {
        if !self.saw_shape {
            return None;
        }
        let extent: ChartUserShapeExtent = match (self.to, self.ext_pt) {
            (Some((x, y)), _) => ChartUserShapeExtent::Corner { x, y },
            (None, Some((width, height))) => ChartUserShapeExtent::Size { width, height },
            // An anchor stating neither corner nor extent gives the shape no
            // size to be drawn at.
            (None, None) => return None,
        };
        Some(ChartUserShape {
            from: self.from,
            extent,
            paragraphs: self
                .paragraphs
                .into_iter()
                .filter(|paragraph| !paragraph.runs.is_empty())
                .collect(),
            text_insets: self.insets.unwrap_or(Insets {
                top: DEFAULT_VERTICAL_INSET_PT,
                right: DEFAULT_SIDE_INSET_PT,
                bottom: DEFAULT_VERTICAL_INSET_PT,
                left: DEFAULT_SIDE_INSET_PT,
            }),
            fill: self.fill,
            no_wrap: self.no_wrap,
            border: self.border_color.map(|color| BorderSide {
                width: self.border_width_pt,
                color,
                style: BorderLineStyle::Solid,
                join: self
                    .border_join
                    .or_else(|| {
                        self.line_reference_index
                            .and_then(|index| index.checked_sub(1))
                            .and_then(|index| theme_line_styles.get(index))
                            .and_then(|style| style.join)
                    })
                    .unwrap_or(LineJoin::Round),
                cap: self
                    .border_cap
                    .or_else(|| {
                        self.line_reference_index
                            .and_then(|index| index.checked_sub(1))
                            .and_then(|index| theme_line_styles.get(index))
                            .and_then(|style| style.cap)
                    })
                    .unwrap_or(LineCap::Flat),
            }),
        })
    }
}

/// Weight a shape outline is stroked at when `<a:ln>` states none, matching
/// the worksheet text boxes' own default.
const DEFAULT_BORDER_WIDTH_PT: f64 = 0.75;

#[cfg(test)]
#[path = "chart_drawing_tests.rs"]
mod chart_drawing_tests;
