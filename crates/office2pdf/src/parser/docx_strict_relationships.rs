//! Give a Strict-conformance package's relationship types their Transitional
//! names before the raw-XML scans and docx-rs read the package.
//!
//! ISO/IEC 29500 Strict names the officeDocument relationship types under
//! `http://purl.oclc.org/ooxml/officeDocument/relationships/`. docx-rs looks
//! parts up by the Transitional types only, so a Strict package lost its
//! styles, numbering and the other parts docx-rs resolves (issue #2243).
//! LibreOffice resolves the same type name in either namespace
//! (`Relations::getFragmentPathFromFirstTypeFromOfficeDoc`, see
//! `references/libreoffice-oox-relations.md`). This swaps the Strict prefix for
//! the Transitional one in each `Type`. A type whose name also differs
//! (Strict `extendedProperties`, Transitional `extended-properties`) keeps its
//! Strict name; nothing here resolves the document-properties parts by type.
use quick_xml::events::{BytesStart, Event, attributes::Attribute};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer};
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

const STRICT_PREFIX: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_PREFIX: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/// The package with every Strict relationship type renamed, or `data` itself
/// when no relationship part names one. A package that cannot be rewritten is
/// passed on unchanged, so docx-rs reports it as before.
pub(super) fn normalize(data: &[u8]) -> Cow<'_, [u8]> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return Cow::Borrowed(data);
    };
    let mut renamed_parts: BTreeMap<usize, Vec<u8>> = BTreeMap::new();
    for index in 0..archive.len() {
        let Ok(mut entry) = archive.by_index(index) else {
            continue;
        };
        if !entry.name().is_ok_and(|name| name.ends_with(".rels")) {
            continue;
        }
        let mut xml: Vec<u8> = Vec::new();
        if entry.read_to_end(&mut xml).is_err() {
            continue;
        }
        if let Some(renamed) = rename_strict_types(&xml) {
            renamed_parts.insert(index, renamed);
        }
    }
    if renamed_parts.is_empty() {
        return Cow::Borrowed(data);
    }
    match repack(&mut archive, &renamed_parts) {
        Ok(package) => {
            tracing::debug!(
                format = "DOCX",
                relationship_parts = renamed_parts.len(),
                "Renamed Strict relationship types to their Transitional names"
            );
            Cow::Owned(package)
        }
        Err(error) => {
            tracing::warn!(
                format = "DOCX",
                %error,
                "Strict relationship types left as written"
            );
            Cow::Borrowed(data)
        }
    }
}

/// The relationship part with its Strict `Type` prefixes swapped, or `None`
/// when it names none (or is not well-formed, which docx-rs handles itself).
fn rename_strict_types(xml: &[u8]) -> Option<Vec<u8>> {
    if !xml
        .windows(STRICT_PREFIX.len())
        .any(|window| window == STRICT_PREFIX)
    {
        return None;
    }
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut changed = false;
    loop {
        let event = match reader.read_event().ok()? {
            Event::Eof => break,
            Event::Start(element) => Event::Start(rename_type(element, &mut changed)?),
            Event::Empty(element) => Event::Empty(rename_type(element, &mut changed)?),
            event => event,
        };
        writer.write_event(event).ok()?;
    }
    changed.then(|| writer.into_inner())
}

/// Swap the Strict prefix of a `Relationship` element's `Type`. Only `Type`
/// names a relationship type: an external `Target` that happens to spell one
/// stays as written.
fn rename_type<'a>(element: BytesStart<'a>, changed: &mut bool) -> Option<BytesStart<'a>> {
    if element.local_name().as_ref() != b"Relationship" {
        return Some(element);
    }
    let mut attributes: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut renamed = false;
    for attribute in element.attributes() {
        let attribute = attribute.ok()?;
        let mut value: Vec<u8> = attribute.value.into_owned();
        if attribute.key.as_ref() == b"Type"
            && let Some(name) = value.strip_prefix(STRICT_PREFIX)
        {
            value = [TRANSITIONAL_PREFIX, name].concat();
            renamed = true;
        }
        attributes.push((attribute.key.as_ref().to_vec(), value));
    }
    if !renamed {
        return Some(element);
    }
    *changed = true;
    let mut element = element;
    element.clear_attributes();
    for (key, value) in &attributes {
        element.push_attribute(Attribute {
            key: QName(key),
            value: Cow::Borrowed(value),
        });
    }
    Some(element)
}

/// Copy the package, writing each renamed relationship part in place of its
/// original and every other entry byte for byte.
fn repack(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    renamed_parts: &BTreeMap<usize, Vec<u8>>,
) -> Result<Vec<u8>, zip::result::ZipError> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let entry = archive.by_index(index)?;
        if let Some(part) = renamed_parts.get(&index) {
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
