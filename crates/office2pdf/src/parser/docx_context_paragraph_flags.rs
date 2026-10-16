//! Paragraph flags recovered from raw XML for published docx-rs compatibility.
//!
//! Published docx-rs omits `wordWrap` (issue #1041) and drops explicit false
//! `pageBreakBefore` values (upstream docx-rs #939/#940). Combining both flags in one
//! cursor avoids a second body scan and another paragraph/style-id inventory.
//! The common paragraph sequence skips revision history and fallback text boxes
//! whose paragraphs never reach conversion (issue #1689).

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::paragraph_cursor::{
    ParagraphCursor, ScannedParagraph, attribute_value, scan_body_paragraphs,
};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(in super::super) struct RawParagraphFlags {
    pub(in super::super) word_wrap: Option<bool>,
    pub(in super::super) page_break_before: Option<bool>,
}

/// ST_OnOff's absent attribute means on; omission of the whole element is
/// represented separately by the optional fields above.
fn on_off_value(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> bool {
    !matches!(
        attribute_value(reader, element, b"val").as_deref(),
        Some("0" | "false" | "off")
    )
}

impl RawParagraphFlags {
    fn read(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        match element.local_name().as_ref() {
            b"wordWrap" => self.word_wrap = Some(on_off_value(reader, element)),
            b"pageBreakBefore" => self.page_break_before = Some(on_off_value(reader, element)),
            _ => {}
        }
    }
}

pub(in super::super) struct ParagraphFlagsContext {
    values: ParagraphCursor<RawParagraphFlags>,
}

impl ParagraphFlagsContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            values: ParagraphCursor::new(
                "paragraph flags",
                xml.map(Self::scan).unwrap_or_default(),
            ),
        }
    }

    /// Advance exactly once per converted `w:p`, checking its `w:pStyle`.
    pub(in super::super) fn next_flags(&self, style_id: Option<&str>) -> RawParagraphFlags {
        self.values
            .next(style_id)
            .map(|(_, flags)| *flags)
            .unwrap_or_default()
    }

    fn scan(xml: &str) -> Vec<ScannedParagraph<RawParagraphFlags>> {
        scan_body_paragraphs::<RawParagraphFlags>(xml, |reader, element, flags| {
            flags.read(reader, element)
        })
    }
}

/// Paragraph-style flags, plus document-default pageBreakBefore under the resolver's
/// reserved default key. Only current, direct paragraph properties count:
/// character styles and `pPrChange` history cannot override the live style.
pub(in super::super) fn scan_style_paragraph_flags(
    xml: Option<&str>,
) -> HashMap<String, RawParagraphFlags> {
    let Some(xml) = xml else {
        return HashMap::new();
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut map: HashMap<String, RawParagraphFlags> = HashMap::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut current_style: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_start = matches!(&event, Event::Start(_));
                let (Event::Start(element) | Event::Empty(element)) = event else {
                    unreachable!()
                };
                let local = element.local_name();
                let name = local.as_ref();
                let named_style = stack.len() == 1 && stack[0] == b"styles" && name == b"style";
                if named_style {
                    current_style = attribute_value(&reader, &element, b"styleId").filter(|_| {
                        attribute_value(&reader, &element, b"type").as_deref() == Some("paragraph")
                    });
                }
                let direct_style_properties = stack.len() == 3
                    && stack[0] == b"styles"
                    && stack[1] == b"style"
                    && stack[2] == b"pPr";
                let default_properties = stack.len() == 4
                    && stack[0] == b"styles"
                    && stack[1] == b"docDefaults"
                    && stack[2] == b"pPrDefault"
                    && stack[3] == b"pPr";
                let style_id = if direct_style_properties {
                    current_style.as_deref()
                } else if default_properties && name == b"pageBreakBefore" {
                    Some(super::super::styles::DOC_DEFAULT_STYLE_ID)
                } else {
                    None
                };
                if matches!(name, b"wordWrap" | b"pageBreakBefore")
                    && let Some(style_id) = style_id
                {
                    map.entry(style_id.to_owned())
                        .or_default()
                        .read(&reader, &element);
                }
                if is_start {
                    stack.push(name.to_vec());
                } else if named_style {
                    current_style = None;
                }
            }
            Ok(Event::End(_)) => {
                if stack.len() == 2 && stack.last().is_some_and(|name| name == b"style") {
                    current_style = None;
                }
                stack.pop();
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    map
}
