use std::cell::Cell;

use crate::parser::units::emu_to_pt;

#[derive(Debug, Clone, Copy, Default)]
pub(in super::super) struct DrawingTextBoxInfo {
    pub(in super::super) width_pt: Option<f64>,
    pub(in super::super) height_pt: Option<f64>,
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
        self.text_boxes.get(index).copied().unwrap_or_default()
    }
}

fn scan_drawing_text_boxes(xml: &str) -> Vec<DrawingTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<DrawingTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut drawing_depth: usize = 0;
    let mut current_info: DrawingTextBoxInfo = DrawingTextBoxInfo::default();
    let mut saw_text_box: bool = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"body" => in_body = true,
                b"drawing" if in_body => {
                    if drawing_depth == 0 {
                        current_info = DrawingTextBoxInfo::default();
                        saw_text_box = false;
                    }
                    drawing_depth += 1;
                }
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, element);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, element);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"body" => in_body = false,
                b"drawing" if drawing_depth > 0 => {
                    drawing_depth -= 1;
                    if drawing_depth == 0 && saw_text_box {
                        result.push(current_info);
                        current_info = DrawingTextBoxInfo::default();
                        saw_text_box = false;
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

fn update_drawing_text_box_extent(
    info: &mut DrawingTextBoxInfo,
    element: &quick_xml::events::BytesStart<'_>,
) {
    if info.width_pt.is_some() && info.height_pt.is_some() {
        return;
    }

    let mut width_emu: Option<u32> = None;
    let mut height_emu: Option<u32> = None;

    for attribute in element.attributes().flatten() {
        match attribute.key.local_name().as_ref() {
            b"cx" => {
                width_emu = std::str::from_utf8(attribute.value.as_ref())
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok());
            }
            b"cy" => {
                height_emu = std::str::from_utf8(attribute.value.as_ref())
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok());
            }
            _ => {}
        }
    }

    if let Some(width_emu) = width_emu {
        info.width_pt = Some(emu_to_pt(width_emu));
    }
    if let Some(height_emu) = height_emu {
        info.height_pt = Some(emu_to_pt(height_emu));
    }
}
