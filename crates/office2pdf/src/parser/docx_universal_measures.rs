//! Write each table width as a number docx-rs reads before the raw-XML scans
//! and docx-rs read the package.
//!
//! A `CT_TblWidth` (`w:tcW`, `w:tblW`, `w:tblInd`, `w:tblCellSpacing`,
//! `w:wBefore`, `w:wAfter` and the sides of `w:tblCellMar` and `w:tcMar`)
//! states its `w:w` as an `ST_MeasurementOrPercent`: twips, a universal
//! measure such as `81.85pt`, or a percentage such as `50%`. A Strict package
//! writes every length as a universal measure. docx-rs reads a `w:tcW` only as
//! an integer, so any other form dropped the whole `w:tcPr` with the cell's
//! shading, merge and alignment, and it reads the other widths' leading number
//! as twips whatever unit follows (issue #2245). Each value is rewritten in
//! the form a Transitional package writes, as LibreOffice reads it
//! (`OOXMLValue::createMeasurementOrPercent`, see
//! `references/libreoffice-ooxml-measures.md`): a percentage as fiftieths of a
//! percent, truncated, and anything else as twips, rounded.
use quick_xml::events::{BytesStart, Event, attributes::Attribute};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer};
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use std::time::Instant;

/// Elements whose `w:w` is a `CT_TblWidth` wherever they appear.
const WIDTH_ELEMENTS: [&[u8]; 6] = [
    b"tcW",
    b"tblW",
    b"tblInd",
    b"tblCellSpacing",
    b"wBefore",
    b"wAfter",
];

/// Elements whose children's `w:w` are `CT_TblWidth` cell margins.
const MARGIN_ELEMENTS: [&[u8]; 2] = [b"tblCellMar", b"tcMar"];

/// The package with every table width in its Transitional form, or `data`
/// itself when each already is. A package that cannot be rewritten is passed
/// on unchanged, so docx-rs reads it as before.
pub(super) fn normalize(data: &[u8]) -> Cow<'_, [u8]> {
    let started = Instant::now();
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return Cow::Borrowed(data);
    };
    let mut rewritten_parts: BTreeMap<usize, Vec<u8>> = BTreeMap::new();
    for index in 0..archive.len() {
        let Ok(mut entry) = archive.by_index(index) else {
            continue;
        };
        if !entry.name().is_ok_and(|name| name.ends_with(".xml")) {
            continue;
        }
        let mut xml: Vec<u8> = Vec::new();
        if entry.read_to_end(&mut xml).is_err() {
            continue;
        }
        if let Some(rewritten) = rewrite_widths(&xml) {
            rewritten_parts.insert(index, rewritten);
        }
    }
    if rewritten_parts.is_empty() {
        return Cow::Borrowed(data);
    }
    match repack(&mut archive, &rewritten_parts) {
        Ok(package) => {
            tracing::debug!(
                format = "DOCX",
                parts = rewritten_parts.len(),
                elapsed_ms = started.elapsed().as_secs_f64() * 1000.0,
                "Rewrote table widths in their Transitional form"
            );
            Cow::Owned(package)
        }
        Err(error) => {
            tracing::warn!(format = "DOCX", %error, "Table widths left as written");
            Cow::Borrowed(data)
        }
    }
}

/// The part with each table width rewritten, or `None` when none needs it
/// (or the part is not well-formed, which docx-rs handles itself).
fn rewrite_widths(xml: &[u8]) -> Option<Vec<u8>> {
    if !states_a_non_integer_w(xml) {
        return None;
    }
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut in_margins = false;
    let mut changed = false;
    loop {
        let event = match reader.read_event().ok()? {
            Event::Eof => break,
            Event::Start(element) => {
                let local_name = element.local_name();
                let is_width = is_width_element(local_name.as_ref(), in_margins);
                in_margins |= MARGIN_ELEMENTS.contains(&local_name.as_ref());
                Event::Start(rewrite_element(element, is_width, &mut changed)?)
            }
            Event::Empty(element) => {
                let is_width = is_width_element(element.local_name().as_ref(), in_margins);
                Event::Empty(rewrite_element(element, is_width, &mut changed)?)
            }
            Event::End(element) => {
                if MARGIN_ELEMENTS.contains(&element.local_name().as_ref()) {
                    in_margins = false;
                }
                Event::End(element)
            }
            event => event,
        };
        writer.write_event(event).ok()?;
    }
    changed.then(|| writer.into_inner())
}

/// Matched by local name, as docx-rs matches them.
fn is_width_element(local_name: &[u8], in_margins: bool) -> bool {
    WIDTH_ELEMENTS.contains(&local_name) || in_margins
}

/// Whether any attribute named `w` holds something other than an integer.
/// Cheap enough to run over every part: most packages state none, and only
/// those parts go through the XML reader.
fn states_a_non_integer_w(xml: &[u8]) -> bool {
    xml.iter().enumerate().any(|(at, &byte)| {
        if byte != b'=' {
            return false;
        }
        let name_end: usize = xml[..at]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .map_or(0, |end| end + 1);
        let names_w = name_end >= 2
            && xml[name_end - 1] == b'w'
            && (xml[name_end - 2] == b':' || xml[name_end - 2].is_ascii_whitespace());
        if !names_w {
            return false;
        }
        let rest: &[u8] = &xml[at + 1..];
        let Some(open) = rest.iter().position(|byte| !byte.is_ascii_whitespace()) else {
            return false;
        };
        let quote: u8 = rest[open];
        if quote != b'"' && quote != b'\'' {
            return false;
        }
        let value: &[u8] = &rest[open + 1..];
        let value: &[u8] = &value[..value
            .iter()
            .position(|&byte| byte == quote)
            .unwrap_or(value.len())];
        match std::str::from_utf8(value) {
            Ok(value) => transitional_width(value).is_some(),
            Err(_) => true,
        }
    })
}

/// `element` with its `w` attribute in the Transitional form when it is a
/// table width.
fn rewrite_element<'a>(
    element: BytesStart<'a>,
    is_width: bool,
    changed: &mut bool,
) -> Option<BytesStart<'a>> {
    if !is_width {
        return Some(element);
    }
    let mut attributes: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut rewritten = false;
    for attribute in element.attributes() {
        let attribute = attribute.ok()?;
        let mut value: Vec<u8> = attribute.value.into_owned();
        if attribute.key.local_name().as_ref() == b"w"
            && let Some(width) = std::str::from_utf8(&value)
                .ok()
                .and_then(transitional_width)
        {
            value = width.to_string().into_bytes();
            rewritten = true;
        }
        attributes.push((attribute.key.as_ref().to_vec(), value));
    }
    if !rewritten {
        return Some(element);
    }
    *changed = true;
    let mut element = element;
    element.clear_attributes();
    for (key, value) in &attributes {
        // Every attribute is written back in double quotes, so a `"` that
        // stood inside single quotes is escaped.
        let value: Cow<'_, [u8]> = if value.contains(&b'"') {
            Cow::Owned(
                value
                    .iter()
                    .flat_map(|&byte| match byte {
                        b'"' => b"&quot;".to_vec(),
                        byte => vec![byte],
                    })
                    .collect(),
            )
        } else {
            Cow::Borrowed(value)
        };
        element.push_attribute(Attribute {
            key: QName(key),
            value,
        });
    }
    Some(element)
}

/// An `ST_MeasurementOrPercent` as the integer a Transitional package writes
/// for it, or `None` when `value` already is that integer. The number is the
/// value's leading decimal, 0 when it has none, as LibreOffice reads it.
fn transitional_width(value: &str) -> Option<i64> {
    let number: f64 = leading_decimal(value.trim_ascii_start());
    let value: &str = value.trim_ascii_end();
    let width: f64 = if value.ends_with('%') {
        (number * 50.0).trunc()
    } else {
        let twips_per_unit: f64 = match value.get(value.len().saturating_sub(2)..) {
            Some("pt") => 20.0,
            Some("in") => 1440.0,
            Some("cm") => 1440.0 / 2.54,
            Some("mm") => 144.0 / 2.54,
            Some("pc" | "pi") => 240.0,
            _ => 1.0,
        };
        (number * twips_per_unit).round()
    };
    let width = width as i64;
    (width.to_string() != value).then_some(width)
}

/// The decimal number `value` starts with: an optional sign, digits and an
/// optional fraction.
fn leading_decimal(value: &str) -> f64 {
    let bytes: &[u8] = value.as_bytes();
    let mut end: usize = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let digits = |from: usize| -> usize {
        from + bytes[from..]
            .iter()
            .take_while(|byte| byte.is_ascii_digit())
            .count()
    };
    end = digits(end);
    if bytes.get(end) == Some(&b'.') {
        end = digits(end + 1);
    }
    value[..end].parse().unwrap_or(0.0)
}

/// Copy the package, writing each rewritten part in place of its original and
/// every other entry byte for byte.
fn repack(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    rewritten_parts: &BTreeMap<usize, Vec<u8>>,
) -> Result<Vec<u8>, zip::result::ZipError> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let entry = archive.by_index(index)?;
        if let Some(part) = rewritten_parts.get(&index) {
            let options =
                zip::write::SimpleFileOptions::default().compression_method(entry.compression());
            zip.start_file(entry.name()?.into_owned(), options)?;
            zip.write_all(part)?;
        } else {
            zip.raw_copy_file(entry)?;
        }
    }
    Ok(zip.finish()?.into_inner())
}
