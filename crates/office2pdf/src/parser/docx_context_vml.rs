use crate::parser::xml_util::{OOXML_XML_VERSION, parse_hex_color};
use std::cell::Cell;

use quick_xml::Reader;
use quick_xml::events::BytesStart;

use crate::ir::{
    Alignment, Block, Color, HorizontalRule, HorizontalRulePaint, HorizontalRuleWidth, Paragraph,
    ParagraphStyle, Run, TextStyle, WrapMode,
};

use super::super::media::{extract_vml_style_dimension, extract_vml_style_value};
use super::paragraph_cursor::attribute_value;
use super::wrap::extract_vml_wrap_mode_from_element;

/// Word's *Insert > Horizontal Line*, a `v:rect` carrying `o:hr="t"`, or
/// `None` for any other `v:rect` and for a rule with nothing to paint.
///
/// The width follows Word rather than the VML specification, as LibreOffice's
/// importer records (`oox/source/vml/vmlshapecontext.cxx`): `o:hrpct` is in
/// tenths of a percent, an absent one is the paragraph's whole width when the
/// `style` width is absent or Word's unitless `0` (any other written width is
/// not read; see the TODO below), and the `style` width counts only with
/// `o:hrpct="0"`. Word's exports agree for a header rule, a body rule and a
/// 60.9% one (issue #2251).
///
/// TODO(#2251 explicit default width): those whole-width rules are written as
/// Word writes them, `width:0` with no unit. Word drew a body rule a Google
/// Docs export writes as `width:0.0pt`, also without `o:hrpct`, as a one-twip
/// sliver (docxide-pdf `mandated_reporter_child_abuse`). Until a native probe
/// settles which width Word gives such a rule, one without `o:hrpct` whose
/// width is written any other way is not read.
pub(in super::super) fn horizontal_rule(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Option<HorizontalRule> {
    let is_true = |name: &[u8]| {
        attribute_value(reader, element, name)
            .is_some_and(|value| matches!(value.trim(), "t" | "true"))
    };
    if !is_true(b"hr") {
        return None;
    }
    let style: Option<String> = attribute_value(reader, element, b"style");
    let height: f64 =
        extract_vml_style_dimension(style.as_deref(), "height").filter(|height| *height > 0.0)?;
    let percent: Option<f64> = attribute_value(reader, element, b"hrpct")
        .and_then(|value| value.trim().parse::<f64>().ok());
    let width: HorizontalRuleWidth = match percent {
        Some(0.0) => HorizontalRuleWidth::Points(
            extract_vml_style_dimension(style.as_deref(), "width").filter(|width| *width > 0.0)?,
        ),
        Some(percent) if percent > 0.0 => HorizontalRuleWidth::Fraction(percent / 1000.0),
        _ => {
            let written: Option<String> = style
                .as_deref()
                .and_then(|style| extract_vml_style_value(style, "width"));
            if written.as_deref().is_some_and(|width| width != "0") {
                tracing::debug!(
                    width = ?written,
                    "horizontal rule without o:hrpct has an explicit width; drawing none"
                );
                return None;
            }
            HorizontalRuleWidth::Fraction(1.0)
        }
    };
    let alignment: Alignment = match attribute_value(reader, element, b"hralign").as_deref() {
        Some("center") => Alignment::Center,
        Some("right") => Alignment::Right,
        _ => Alignment::Left,
    };
    let paint: HorizontalRulePaint = if is_true(b"hrnoshade") {
        HorizontalRulePaint::Solid(
            attribute_value(reader, element, b"fillcolor").and_then(|value| vml_color(&value))?,
        )
    } else {
        HorizontalRulePaint::Shaded
    };
    Some(HorizontalRule {
        width,
        height,
        alignment,
        paint,
    })
}

/// A VML colour: `#rrggbb`, `#rgb` or one of the sixteen named colours, with
/// any palette index Word appends (`#a0a0a0 [3212]`) ignored.
fn vml_color(value: &str) -> Option<Color> {
    let color: &str = value.split_whitespace().next()?;
    if let Some(hex) = color.strip_prefix('#') {
        return match hex.len() {
            6 => parse_hex_color(hex),
            3 => parse_hex_color(
                &hex.chars()
                    .flat_map(|digit| [digit, digit])
                    .collect::<String>(),
            ),
            _ => None,
        };
    }
    let hex: &str = match color.to_ascii_lowercase().as_str() {
        "black" => "000000",
        "silver" => "C0C0C0",
        "gray" => "808080",
        "white" => "FFFFFF",
        "maroon" => "800000",
        "red" => "FF0000",
        "purple" => "800080",
        "fuchsia" => "FF00FF",
        "green" => "008000",
        "lime" => "00FF00",
        "olive" => "808000",
        "yellow" => "FFFF00",
        "navy" => "000080",
        "blue" => "0000FF",
        "teal" => "008080",
        "aqua" => "00FFFF",
        _ => return None,
    };
    parse_hex_color(hex)
}

#[derive(Debug, Clone, Default)]
pub(in super::super) struct VmlTextBoxInfo {
    pub(in super::super) paragraphs: Vec<String>,
    pub(in super::super) wrap_mode: Option<WrapMode>,
}

impl VmlTextBoxInfo {
    pub(in super::super) fn into_blocks(self) -> Vec<Block> {
        self.paragraphs
            .into_iter()
            .filter(|text| !text.is_empty())
            .map(|text| {
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text,
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                        inline_box: None,
                    }],
                })
            })
            .collect()
    }
}

pub(in super::super) struct VmlTextBoxContext {
    text_boxes: Vec<VmlTextBoxInfo>,
    cursor: Cell<usize>,
}

impl VmlTextBoxContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_vml_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn consume_next(&self) -> VmlTextBoxInfo {
        let index: usize = self.cursor.get();
        self.cursor.set(index + 1);
        self.text_boxes.get(index).cloned().unwrap_or_default()
    }
}

fn scan_vml_text_boxes(xml: &str) -> Vec<VmlTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<VmlTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut pict_depth: usize = 0;
    let mut shape_depth: usize = 0;
    let mut in_text_box_content: bool = false;
    let mut in_paragraph: bool = false;
    let mut current_picture_shapes: Vec<VmlTextBoxInfo> = Vec::new();
    let mut current_picture_wrap: Option<WrapMode> = None;
    let mut current_shape_paragraphs: Vec<String> = Vec::new();
    let mut current_paragraph_text: String = String::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"body" => in_body = true,
                b"pict" if in_body => {
                    if pict_depth == 0 {
                        current_picture_shapes.clear();
                        current_picture_wrap = None;
                    }
                    pict_depth += 1;
                }
                b"shape" if pict_depth > 0 => {
                    if shape_depth == 0 {
                        current_shape_paragraphs.clear();
                    }
                    shape_depth += 1;
                }
                b"txbxContent" if shape_depth > 0 => in_text_box_content = true,
                b"p" if in_text_box_content => {
                    in_paragraph = true;
                    current_paragraph_text.clear();
                }
                b"wrap" if pict_depth > 0 => {
                    current_picture_wrap = extract_vml_wrap_mode_from_element(element);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"tab" if in_paragraph => current_paragraph_text.push('\t'),
                b"br" if in_paragraph => current_paragraph_text.push('\n'),
                b"wrap" if pict_depth > 0 => {
                    current_picture_wrap = extract_vml_wrap_mode_from_element(element);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref element)) => {
                if in_paragraph && let Ok(text) = element.xml_content(OOXML_XML_VERSION) {
                    current_paragraph_text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"body" => in_body = false,
                b"p" if in_paragraph => {
                    current_shape_paragraphs.push(std::mem::take(&mut current_paragraph_text));
                    in_paragraph = false;
                }
                b"txbxContent" if in_text_box_content => in_text_box_content = false,
                b"shape" if shape_depth > 0 => {
                    shape_depth -= 1;
                    if shape_depth == 0 {
                        current_picture_shapes.push(VmlTextBoxInfo {
                            paragraphs: std::mem::take(&mut current_shape_paragraphs),
                            wrap_mode: None,
                        });
                        in_text_box_content = false;
                        in_paragraph = false;
                        current_paragraph_text.clear();
                    }
                }
                b"pict" if pict_depth > 0 => {
                    pict_depth -= 1;
                    if pict_depth == 0 {
                        for mut text_box in current_picture_shapes.drain(..) {
                            text_box.wrap_mode = current_picture_wrap;
                            result.push(text_box);
                        }
                        current_picture_wrap = None;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    result
}
