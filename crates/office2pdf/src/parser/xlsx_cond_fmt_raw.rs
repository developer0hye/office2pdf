use std::collections::HashMap;
use std::io::Read;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct RawCondFmtHint {
    pub(crate) min_length: Option<u32>,
    pub(crate) max_length: Option<u32>,
    pub(crate) icon_set_type: Option<String>,
}

pub(crate) type RawCondFmtHints = HashMap<i32, RawCondFmtHint>;
pub(crate) type SheetCondFmtHints = HashMap<String, RawCondFmtHints>;

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

fn read_zip_text(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Option<String> {
    let mut file = archive.by_name(path).ok()?;
    let mut text = String::new();
    file.read_to_string(&mut text).ok()?;
    Some(text)
}

fn parse_relationships(xml: &str) -> HashMap<String, String> {
    let mut relationships = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"Relationship" =>
            {
                if let (Some(id), Some(target)) = (
                    attr_value(&reader, &element, b"Id"),
                    attr_value(&reader, &element, b"Target"),
                ) {
                    relationships.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    relationships
}

fn parse_sheet_relationships(xml: &str) -> Vec<(String, String)> {
    let mut sheets = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"sheet" =>
            {
                if let (Some(name), Some(relationship_id)) = (
                    attr_value(&reader, &element, b"name"),
                    attr_value(&reader, &element, b"id"),
                ) {
                    sheets.push((name, relationship_id));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    sheets
}

fn worksheet_path(target: &str) -> String {
    let target = target.trim_start_matches('/');
    if target.starts_with("xl/") {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
}

pub(crate) fn parse_worksheet_hints(xml: &str) -> RawCondFmtHints {
    let mut hints = HashMap::new();
    let mut current_priority = None;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"cfRule" => {
                current_priority = attr_value(&reader, &element, b"priority")
                    .and_then(|value| value.parse::<i32>().ok());
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"dataBar" =>
            {
                if let Some(priority) = current_priority {
                    let hint = hints
                        .entry(priority)
                        .or_insert_with(RawCondFmtHint::default);
                    hint.min_length = attr_value(&reader, &element, b"minLength")
                        .and_then(|value| value.parse::<u32>().ok());
                    hint.max_length = attr_value(&reader, &element, b"maxLength")
                        .and_then(|value| value.parse::<u32>().ok());
                }
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"iconSet" =>
            {
                if let Some(priority) = current_priority {
                    let hint = hints
                        .entry(priority)
                        .or_insert_with(RawCondFmtHint::default);
                    hint.icon_set_type = attr_value(&reader, &element, b"iconSet");
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"cfRule" => {
                current_priority = None;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    hints
}

/// Preserve conditional-format attributes that umya-spreadsheet does not
/// expose in its registry release. The result is keyed by worksheet name and
/// rule priority so it can be joined with umya's parsed rule collection.
pub(crate) fn extract_cond_fmt_hints(data: &[u8]) -> SheetCondFmtHints {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };
    let Some(workbook_xml) = read_zip_text(&mut archive, "xl/workbook.xml") else {
        return HashMap::new();
    };
    let Some(relationships_xml) = read_zip_text(&mut archive, "xl/_rels/workbook.xml.rels") else {
        return HashMap::new();
    };

    let relationships = parse_relationships(&relationships_xml);
    let mut result = HashMap::new();
    for (sheet_name, relationship_id) in parse_sheet_relationships(&workbook_xml) {
        let Some(target) = relationships.get(&relationship_id) else {
            continue;
        };
        let Some(worksheet_xml) = read_zip_text(&mut archive, &worksheet_path(target)) else {
            continue;
        };
        let hints = parse_worksheet_hints(&worksheet_xml);
        if !hints.is_empty() {
            result.insert(sheet_name, hints);
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn worksheet_hints_preserve_data_bar_lengths_and_icon_set_type() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1:A3">
    <cfRule type="dataBar" priority="1">
      <dataBar minLength="15" maxLength="85"/>
    </cfRule>
    <cfRule type="iconSet" priority="2">
      <iconSet iconSet="3Arrows"/>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

        let hints = parse_worksheet_hints(xml);
        assert_eq!(hints.get(&1).and_then(|hint| hint.min_length), Some(15));
        assert_eq!(hints.get(&1).and_then(|hint| hint.max_length), Some(85));
        assert_eq!(
            hints.get(&2).and_then(|hint| hint.icon_set_type.as_deref()),
            Some("3Arrows")
        );
    }
}
