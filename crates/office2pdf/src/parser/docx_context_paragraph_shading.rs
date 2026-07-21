use std::cell::Cell;
use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::ir::Color;
use crate::parser::xml_util;

fn attr_value(reader: &Reader<&[u8]>, element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            attribute
                .decode_and_unescape_value(reader.decoder())
                .ok()
                .map(|value| value.into_owned())
        })
}

fn shading_fill(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> Option<Color> {
    attr_value(reader, element, b"fill").and_then(|fill| xml_util::parse_hex_color(&fill))
}

pub(in super::super) struct ParagraphShadingContext {
    backgrounds: Vec<Option<Color>>,
    cursor: Cell<usize>,
}

impl ParagraphShadingContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            backgrounds: xml.map(Self::scan).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn next_background(&self) -> Option<Color> {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.backgrounds.get(index).copied().flatten()
    }

    fn scan(xml: &str) -> Vec<Option<Color>> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut backgrounds = Vec::new();
        let mut paragraph_stack = Vec::new();
        let mut in_body = false;
        let mut in_paragraph_properties = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) => match element.local_name().as_ref() {
                    b"body" => in_body = true,
                    b"p" if in_body => {
                        backgrounds.push(None);
                        paragraph_stack.push(backgrounds.len() - 1);
                    }
                    b"pPr" if !paragraph_stack.is_empty() => in_paragraph_properties = true,
                    b"shd" if in_paragraph_properties => {
                        if let Some(index) = paragraph_stack.last().copied() {
                            backgrounds[index] = shading_fill(&reader, &element);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                    b"p" if in_body => backgrounds.push(None),
                    b"shd" if in_paragraph_properties => {
                        if let Some(index) = paragraph_stack.last().copied() {
                            backgrounds[index] = shading_fill(&reader, &element);
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(element)) => match element.local_name().as_ref() {
                    b"body" => in_body = false,
                    b"p" if in_body => {
                        paragraph_stack.pop();
                        in_paragraph_properties = false;
                    }
                    b"pPr" => in_paragraph_properties = false,
                    _ => {}
                },
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }

        backgrounds
    }
}

pub(in super::super) fn scan_style_paragraph_shading(xml: Option<&str>) -> HashMap<String, Color> {
    let Some(xml) = xml else {
        return HashMap::new();
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut backgrounds = HashMap::new();
    let mut paragraph_style_id = None;
    let mut in_paragraph_properties = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"style" => {
                    let style_type = attr_value(&reader, &element, b"type");
                    paragraph_style_id = (style_type.as_deref() == Some("paragraph"))
                        .then(|| attr_value(&reader, &element, b"styleId"))
                        .flatten();
                }
                b"pPr" if paragraph_style_id.is_some() => in_paragraph_properties = true,
                b"shd" if in_paragraph_properties => {
                    if let (Some(style_id), Some(color)) =
                        (paragraph_style_id.as_ref(), shading_fill(&reader, &element))
                    {
                        backgrounds.insert(style_id.clone(), color);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"shd" => {
                if in_paragraph_properties
                    && let (Some(style_id), Some(color)) =
                        (paragraph_style_id.as_ref(), shading_fill(&reader, &element))
                {
                    backgrounds.insert(style_id.clone(), color);
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"style" => {
                    paragraph_style_id = None;
                    in_paragraph_properties = false;
                }
                b"pPr" => in_paragraph_properties = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    backgrounds
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scans_direct_paragraph_shading_in_document_order() {
        let xml = r#"<w:document xmlns:w="urn:w"><w:body>
          <w:p><w:pPr><w:shd w:fill="F4F4F4"/></w:pPr></w:p>
          <w:p><w:pPr><w:shd w:fill="auto"/></w:pPr></w:p>
        </w:body></w:document>"#;
        let context = ParagraphShadingContext::from_xml(Some(xml));

        assert_eq!(
            context.next_background(),
            Some(Color::new(0xF4, 0xF4, 0xF4))
        );
        assert_eq!(context.next_background(), None);
    }

    #[test]
    fn scans_only_paragraph_style_shading() {
        let xml = r#"<w:styles xmlns:w="urn:w">
          <w:style w:type="character" w:styleId="CodeChar"><w:rPr><w:shd w:fill="111111"/></w:rPr></w:style>
          <w:style w:type="paragraph" w:styleId="Code"><w:pPr><w:shd w:fill="E7E7E7"/></w:pPr></w:style>
        </w:styles>"#;
        let backgrounds = scan_style_paragraph_shading(Some(xml));

        assert_eq!(backgrounds.get("Code"), Some(&Color::new(0xE7, 0xE7, 0xE7)));
        assert!(!backgrounds.contains_key("CodeChar"));
    }
}
