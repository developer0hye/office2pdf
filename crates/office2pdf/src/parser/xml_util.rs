//! Shared XML parsing utility functions.
//!
//! These free functions eliminate duplication across pptx, docx, xlsx, chart,
//! and omml parsers.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::ir::Color;

/// Get a string attribute value from an XML element.
/// Matches on full qualified name first (e.g. `r:id`), then local name.
pub(crate) fn get_attr_str(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key || attr.key.local_name().as_ref() == key {
            return attr.unescape_value().ok().map(|v| v.to_string());
        }
    }
    None
}

/// Get an i64 attribute value from an XML element.
pub(crate) fn get_attr_i64(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<i64> {
    get_attr_str(e, key).and_then(|v| v.parse().ok())
}

/// Skip an XML element and all its children, consuming events until the
/// matching end tag is found. `end_tag` is the local name of the element.
pub(crate) fn skip_element(reader: &mut Reader<&[u8]>, end_tag: &[u8]) {
    let mut depth = 1u32;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth -= 1;
                    if depth == 0 {
                        return;
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => return,
            _ => {}
        }
    }
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
pub(crate) fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Parse an RGB or ARGB hex string into an IR Color.
///
/// OOXML producers use both six-digit `RRGGBB` and eight-digit `AARRGGBB`
/// encodings. Alpha is currently ignored because the IR color is opaque.
pub(crate) fn parse_argb_color(argb: &str) -> Option<Color> {
    match argb.len() {
        6 => parse_hex_color(argb),
        8 => {
            let r = u8::from_str_radix(&argb[2..4], 16).ok()?;
            let g = u8::from_str_radix(&argb[4..6], 16).ok()?;
            let b = u8::from_str_radix(&argb[6..8], 16).ok()?;
            Some(Color::new(r, g, b))
        }
        _ => None,
    }
}

/// Resolve a relative path (like `../drawings/drawing1.xml`) against a base directory.
pub(crate) fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    if relative.starts_with('/') {
        return relative.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for segment in relative.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// One `<Relationship>` entry from an OPC `.rels` part.
pub(crate) struct RelationshipEntry {
    pub id: String,
    pub target: String,
    pub rel_type: Option<String>,
}

/// Parse an OPC `.rels` part into its `<Relationship>` entries in document
/// order. Entries without an `Id` or `Target` are dropped; a missing `Type`
/// is kept as `None` (some producers omit it).
pub(crate) fn parse_relationships(xml: &str) -> Vec<RelationshipEntry> {
    let mut entries: Vec<RelationshipEntry> = Vec::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                {
                    entries.push(RelationshipEntry {
                        id,
                        target,
                        rel_type: get_attr_str(e, b"Type"),
                    });
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    entries
}

/// Parse a .rels XML file and return a map of Id → Target.
pub(crate) fn parse_rels_id_target(xml: &str) -> std::collections::HashMap<String, String> {
    parse_relationships(xml)
        .into_iter()
        .map(|entry| (entry.id, entry.target))
        .collect()
}

#[cfg(test)]
#[path = "xml_util_tests.rs"]
mod tests;
