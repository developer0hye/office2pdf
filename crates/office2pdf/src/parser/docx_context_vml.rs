use std::cell::Cell;

use crate::ir::{Block, Paragraph, ParagraphStyle, Run, TextStyle, WrapMode};

use super::wrap::extract_vml_wrap_mode_from_element;

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
                if in_paragraph && let Ok(text) = element.xml_content() {
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
