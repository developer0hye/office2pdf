use std::cell::Cell;

use crate::ir::{BorderSide, Color, FrameAlign, FrameAnchor, Insets};
use crate::parser::xml_util::OOXML_XML_VERSION;
use quick_xml::events::BytesStart;

use super::docx_context_shape::{ShapeBuilder, ShapeScanState};

#[derive(Debug, Clone, Default)]
pub(in super::super) struct DrawingTextBoxInfo {
    pub(in super::super) width_pt: Option<f64>,
    pub(in super::super) height_pt: Option<f64>,
    /// Outline from the shape's `a:ln`, `None` when it paints none.
    pub(in super::super) stroke: Option<BorderSide>,
    /// Background from the shape's `a:solidFill`.
    pub(in super::super) fill: Option<Color>,
    /// Where the box's text starts inside it, from `wps:bodyPr`.
    pub(in super::super) padding: Insets,
    pub(in super::super) horizontal_anchor: FrameAnchor,
    pub(in super::super) horizontal_align: Option<FrameAlign>,
}

pub(in super::super) struct DrawingTextBoxContext {
    text_boxes: Vec<DrawingTextBoxInfo>,
    cursor: Cell<usize>,
}

impl DrawingTextBoxContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_drawing_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn consume_next(&self) -> DrawingTextBoxInfo {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.text_boxes.get(index).cloned().unwrap_or_default()
    }
}

/// Scan `word/document.xml` for every `wps:wsp` text box drawing, in document
/// order, retaining its extent, frame styles, and horizontal position metadata.
///
/// The shared [`ShapeScanState`] reads `wp:extent`, `a:ln`, `a:solidFill` and
/// `wps:bodyPr`; this scan also preserves `wp:positionH` alignment/reference.
fn scan_drawing_text_boxes(xml: &str) -> Vec<DrawingTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<DrawingTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut drawing_depth: usize = 0;
    let mut horizontal_anchor: FrameAnchor = FrameAnchor::Text;
    let mut horizontal_align: Option<FrameAlign> = None;
    let mut in_position_h: bool = false;
    let mut in_horizontal_align: bool = false;
    let mut state: ShapeScanState = ShapeScanState::default();
    let mut builder: Option<ShapeBuilder> = None;

    // docx-rs reduces `<wp:align>` to a zero offset; keep its text and
    // reference frame alongside the drawing metadata it already omits.
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => {
                match element.local_name().as_ref() {
                    b"body" => in_body = true,
                    b"drawing" if in_body => {
                        if drawing_depth == 0 {
                            builder = Some(ShapeBuilder::default());
                            state.reset();
                            horizontal_anchor = FrameAnchor::Text;
                            horizontal_align = None;
                            in_position_h = false;
                            in_horizontal_align = false;
                        }
                        drawing_depth += 1;
                    }
                    b"positionH" if drawing_depth == 1 => {
                        horizontal_anchor = drawing_horizontal_anchor(element);
                        horizontal_align = None;
                        in_position_h = true;
                    }
                    b"align" if drawing_depth == 1 && in_position_h => {
                        in_horizontal_align = true;
                    }
                    // Only the box's own `wps:wsp` describes the box. A nested
                    // `w:drawing` inside `w:txbxContent` — a picture in one of
                    // the box's paragraphs — carries its own `wp:extent`,
                    // `a:ln` and `a:solidFill`, and would otherwise overwrite
                    // the box's size and frame with the picture's.
                    _ if drawing_depth == 1 => state.start(builder.as_mut(), element),
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref element)) => {
                if drawing_depth == 1 {
                    if element.local_name().as_ref() == b"positionH" {
                        horizontal_anchor = drawing_horizontal_anchor(element);
                        horizontal_align = None;
                    } else {
                        state.empty(builder.as_mut(), element);
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref text)) if in_horizontal_align => {
                if let Ok(value) = text.xml_content(OOXML_XML_VERSION) {
                    horizontal_align = drawing_horizontal_alignment(&value);
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"body" => in_body = false,
                b"align" if in_horizontal_align => in_horizontal_align = false,
                b"positionH" if in_position_h => {
                    in_position_h = false;
                    in_horizontal_align = false;
                }
                b"drawing" if drawing_depth > 0 => {
                    drawing_depth -= 1;
                    if drawing_depth == 0
                        && let Some(builder) = builder.take()
                        && builder.is_text_box()
                    {
                        result.push(text_box_info(&builder, horizontal_anchor, horizontal_align));
                    }
                }
                other if drawing_depth == 1 => state.end(other),
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

fn drawing_horizontal_alignment(value: &str) -> Option<FrameAlign> {
    match value.trim() {
        "left" | "start" => Some(FrameAlign::Start),
        "center" => Some(FrameAlign::Center),
        "right" | "end" => Some(FrameAlign::End),
        _ => None,
    }
}

fn drawing_horizontal_anchor(element: &BytesStart<'_>) -> FrameAnchor {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == b"relativeFrom")
        .map(|attribute| match attribute.value.as_ref() {
            b"page" => FrameAnchor::Page,
            b"margin" => FrameAnchor::Margin,
            _ => FrameAnchor::Text,
        })
        .unwrap_or(FrameAnchor::Text)
}

fn text_box_info(
    builder: &ShapeBuilder,
    horizontal_anchor: FrameAnchor,
    horizontal_align: Option<FrameAlign>,
) -> DrawingTextBoxInfo {
    let (width_pt, height_pt) = builder.box_size_pt();
    let (stroke, fill) = builder.text_box_frame();
    DrawingTextBoxInfo {
        width_pt,
        height_pt,
        stroke,
        fill,
        padding: builder.text_box_insets(),
        horizontal_anchor,
        horizontal_align,
    }
}
