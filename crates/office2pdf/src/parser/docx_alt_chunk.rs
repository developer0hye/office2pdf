//! Expand HTML imports before both raw-XML scans and docx-rs see the document.
//! Keeping the two inputs identical preserves paragraph/section cursor alignment.
use super::docx_html_import::html_to_wordml;
use crate::error::{ConvertError, ConvertWarning};
use quick_xml::{Writer, events::Event, name::ResolveResult, reader::NsReader};
use std::{
    borrow::Cow,
    collections::HashMap,
    io::{Cursor, Read, Write},
};

const WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE: &[u8] = b"http://schemas.openxmlformats.org/package/2006/relationships";

fn bound(namespace: ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(n) if n.as_ref() == expected
        || expected == WORD && n.as_ref() == b"http://purl.oclc.org/ooxml/wordprocessingml/main"
        || expected == REL && n.as_ref() == b"http://purl.oclc.org/ooxml/officeDocument/relationships")
}

#[derive(Default)]
struct Relationship {
    target: String,
    kind: String,
    external: bool,
}

fn attrs(element: &quick_xml::events::BytesStart<'_>) -> HashMap<String, String> {
    element
        .attributes()
        .flatten()
        .filter_map(|a| {
            Some((
                String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                a.normalized_value(super::xml_util::OOXML_XML_VERSION)
                    .ok()?
                    .into_owned(),
            ))
        })
        .collect()
}

fn read_part(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, path: &str) -> Result<Vec<u8>, String> {
    let mut bytes = Vec::new();
    archive
        .by_name(path)
        .map_err(|_| "import part is missing".to_string())?
        .read_to_end(&mut bytes)
        .map_err(|_| "import part could not be read".to_string())?;
    Ok(bytes)
}

fn relationships(xml: &[u8]) -> HashMap<String, Relationship> {
    let mut reader = NsReader::from_reader(xml);
    let mut result = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => {
                let (ns, name) = reader.resolver().resolve_element(e.name());
                if bound(ns, PACKAGE) && name.as_ref() == b"Relationship" {
                    let mut a = attrs(&e);
                    if let Some(id) = a.remove("Id") {
                        result.insert(
                            id,
                            Relationship {
                                target: a.remove("Target").unwrap_or_default(),
                                kind: a.remove("Type").unwrap_or_default(),
                                external: a.get("TargetMode").is_some_and(|v| v == "External"),
                            },
                        );
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn part_path(target: &str) -> Result<String, String> {
    if target.contains(['\\', ':', '?', '#']) {
        return Err("invalid import part URI".into());
    }
    let mut parts: Vec<&str> = if target.starts_with('/') {
        vec![]
    } else {
        vec!["word"]
    };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop().ok_or("import part URI escapes the package")?;
            }
            _ => parts.push(part),
        }
    }
    Ok(parts.join("/"))
}

fn content_type(xml: &[u8], path: &str) -> Option<String> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut extension_type = None;
    loop {
        match reader.read_event() {
            Ok(Event::Empty(e) | Event::Start(e)) => {
                let a = attrs(&e);
                if e.local_name().as_ref() == b"Override"
                    && a.get("PartName")
                        .is_some_and(|n| n.trim_start_matches('/') == path)
                {
                    return a.get("ContentType").cloned();
                }
                if e.local_name().as_ref() == b"Default"
                    && a.get("Extension")
                        .is_some_and(|ext| path.rsplit('.').next() == Some(ext.as_str()))
                {
                    extension_type = a.get("ContentType").cloned();
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    extension_type
}

fn mime_text(message: &mailparse::ParsedMail<'_>) -> Result<String, String> {
    if let mailparse::body::Body::QuotedPrintable(body) = message.get_body_encoded()
        && body.get_raw().iter().any(|byte| *byte >= 0x80)
    {
        // Word-generated MHTML sometimes mixes literal UTF-8 with =XX escapes.
        // The MIME decoder's robust mode strips those non-conforming bytes.
        // Escape them before transfer decoding; keep charset decoding in MIME.
        let mut escaped = Vec::with_capacity(body.get_raw().len());
        const HEX: &[u8] = b"0123456789ABCDEF";
        for byte in body.get_raw() {
            if *byte >= 0x80 {
                escaped.extend([b'=', HEX[(byte >> 4) as usize], HEX[(byte & 15) as usize]]);
            } else {
                escaped.push(*byte);
            }
        }
        if let mailparse::body::Body::QuotedPrintable(preserved) =
            mailparse::body::Body::new(&escaped, &message.ctype, &Some("quoted-printable".into()))
        {
            return preserved
                .get_decoded_as_string()
                .map_err(|_| "MHTML text decoding failed".into());
        }
    }
    message
        .get_body()
        .map_err(|_| "MHTML text decoding failed".into())
}

fn mime_html(message: &mailparse::ParsedMail<'_>) -> Option<Result<String, String>> {
    if message.ctype.mimetype.eq_ignore_ascii_case("text/html") {
        return Some(mime_text(message));
    }
    message.subparts.iter().find_map(mime_html)
}

fn import(
    rel: &Relationship,
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    types: &[u8],
) -> Result<(String, Vec<String>), String> {
    if rel.external {
        return Err("external altChunk imports are not fetched".into());
    }
    if !rel
        .kind
        .rsplit('/')
        .next()
        .is_some_and(|kind| kind.eq_ignore_ascii_case("afChunk"))
    {
        return Err("altChunk relationship has the wrong type".into());
    }
    let path = part_path(&rel.target)?;
    let mime = content_type(types, &path).ok_or("import content type is missing")?;
    let bytes = read_part(archive, &path)?;
    let html = match mime
        .split(';')
        .next()
        .unwrap_or("")
        .trim()
        .to_ascii_lowercase()
        .as_str()
    {
        "text/html" | "application/xhtml+xml" => {
            // Reuse MIME's charset handling; an absent charset uses UTF-8 for
            // HTML here rather than MIME's ASCII default.
            let mut message = format!(
                "Content-Type: {mime}{}\r\n\r\n",
                if mime.to_ascii_lowercase().contains("charset=") {
                    ""
                } else {
                    "; charset=utf-8"
                }
            )
            .into_bytes();
            message.extend(bytes);
            mailparse::parse_mail(&message)
                .map_err(|_| "HTML charset header is invalid")?
                .get_body()
                .map_err(|_| "HTML text decoding failed")?
        }
        "message/rfc822" | "multipart/related" => {
            let message = mailparse::parse_mail(&bytes).map_err(|_| "MHTML message is invalid")?;
            mime_html(&message).ok_or("MHTML has no HTML body")??
        }
        _ => return Err(format!("unsupported altChunk content type: {mime}")),
    };
    html_to_wordml(&html)
}

/// Word suppresses HTML auto-before spacing at the document/cell opening.
/// Explicit spacing stays unchanged; all scanners then consume this same XML.
fn suppress_opening_auto_spacing(xml: &[u8]) -> Result<Vec<u8>, ConvertError> {
    let mut reader = NsReader::from_reader(xml);
    let mut writer = Writer::new(Vec::new());
    let mut stack: Vec<String> = Vec::new();
    let mut scopes: Vec<(usize, bool)> = Vec::new();
    let mut first_paragraph: Option<usize> = None;
    loop {
        let event = reader.read_event().map_err(crate::parser::parse_err)?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let (ns, local) = reader.resolver().resolve_element(element.name());
                let word = bound(ns, WORD);
                let name = String::from_utf8_lossy(local.as_ref()).into_owned();
                let depth = stack.len();
                if word
                    && (name == "p" || name == "tbl")
                    && let Some((scope, occupied)) = scopes.last_mut()
                    && depth == *scope + 1
                {
                    first_paragraph = (name == "p" && !*occupied).then_some(depth);
                    *occupied = true;
                }
                if word && name == "spacing" && first_paragraph.is_some_and(|p| depth == p + 2) {
                    let auto = element.attributes().flatten().any(|a| {
                        let (ns, name) = reader.resolver().resolve_attribute(a.key);
                        bound(ns, WORD)
                            && name.as_ref() == b"beforeAutospacing"
                            && a.normalized_value(super::xml_util::OOXML_XML_VERSION)
                                .is_ok_and(|value| matches!(value.as_ref(), "1" | "true" | "on"))
                    });
                    if auto {
                        let mut fixed = element.to_owned();
                        fixed.clear_attributes();
                        for a in element.attributes().flatten() {
                            let (ns, name) = reader.resolver().resolve_attribute(a.key);
                            if bound(ns, WORD) && name.as_ref() == b"before" {
                                fixed.push_attribute((a.key.as_ref(), b"0".as_slice()));
                            } else {
                                fixed.push_attribute(a);
                            }
                        }
                        writer
                            .write_event(if matches!(event, Event::Start(_)) {
                                Event::Start(fixed)
                            } else {
                                Event::Empty(fixed)
                            })
                            .map_err(crate::parser::parse_err)?;
                        if matches!(event, Event::Start(_)) {
                            stack.push(name);
                        }
                        continue;
                    }
                }
                if matches!(event, Event::Start(_)) {
                    if word && matches!(name.as_str(), "body" | "tc") {
                        scopes.push((depth, false));
                    }
                    stack.push(if word { name } else { String::new() });
                }
            }
            Event::End(_) => {
                stack.pop();
                let depth = stack.len();
                if scopes.last().is_some_and(|(scope, _)| *scope == depth) {
                    scopes.pop();
                }
                if first_paragraph == Some(depth) {
                    first_paragraph = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        writer
            .write_event(event)
            .map_err(crate::parser::parse_err)?;
    }
    Ok(writer.into_inner())
}

pub(super) fn expand(data: &[u8]) -> Result<(Cow<'_, [u8]>, Vec<ConvertWarning>), ConvertError> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return Ok((Cow::Borrowed(data), Vec::new()));
    };
    let Ok(xml) = read_part(&mut archive, "word/document.xml") else {
        return Ok((Cow::Borrowed(data), Vec::new()));
    };
    if !xml
        .windows(b"altChunk".len())
        .any(|window| window == b"altChunk")
    {
        return Ok((Cow::Borrowed(data), Vec::new()));
    }
    let mut reader = NsReader::from_reader(xml.as_slice());
    let mut writer = Writer::new(Vec::new());
    let rels =
        relationships(&read_part(&mut archive, "word/_rels/document.xml.rels").unwrap_or_default());
    let types = read_part(&mut archive, "[Content_Types].xml").unwrap_or_default();
    let mut warnings = Vec::new();
    let mut changed = false;
    let mut skip_depth: usize = 0;
    loop {
        let event = reader
            .read_event()
            .map_err(|e| crate::parser::parse_err(format!("DOCX import XML: {e}")))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                _ => {}
            };
            continue;
        }
        if let Event::Start(e) | Event::Empty(e) = &event {
            let (ns, name) = reader.resolver().resolve_element(e.name());
            if bound(ns, WORD) && name.as_ref() == b"altChunk" {
                changed = true;
                let id = e.attributes().flatten().find_map(|a| {
                    let (ns, name) = reader.resolver().resolve_attribute(a.key);
                    (bound(ns, REL) && name.as_ref() == b"id")
                        .then(|| {
                            a.normalized_value(super::xml_util::OOXML_XML_VERSION)
                                .ok()
                                .map(Cow::into_owned)
                        })
                        .flatten()
                });
                let result = id
                    .as_ref()
                    .and_then(|id| rels.get(id))
                    .ok_or_else(|| "altChunk relationship is missing".to_string())
                    .and_then(|rel| import(rel, &mut archive, &types));
                match result {
                    Ok((fragment, findings)) => {
                        warnings.extend(findings.into_iter().map(|reason| {
                            ConvertWarning::ParseSkipped {
                                format: "DOCX".into(),
                                reason,
                            }
                        }));
                        writer
                            .get_mut()
                            .write_all(fragment.as_bytes())
                            .map_err(crate::parser::parse_err)?;
                    }
                    Err(reason) => {
                        tracing::warn!(
                            relationship_id = id.as_deref().unwrap_or(""),
                            reason,
                            "DOCX imported content skipped"
                        );
                        warnings.push(ConvertWarning::ParseSkipped {
                            format: "DOCX".into(),
                            reason,
                        });
                    }
                }
                if matches!(event, Event::Start(_)) {
                    skip_depth = 1;
                }
                continue;
            }
        }
        if matches!(event, Event::Eof) {
            break;
        }
        writer
            .write_event(event)
            .map_err(crate::parser::parse_err)?;
    }
    if !changed {
        return Ok((Cow::Borrowed(data), warnings));
    }
    let document = suppress_opening_auto_spacing(&writer.into_inner())?;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let entry = archive.by_index(index).map_err(crate::parser::parse_err)?;
        let entry_name = entry.name().map_err(crate::parser::parse_err)?.into_owned();
        if entry_name == "word/document.xml" {
            zip.start_file(
                entry_name,
                zip::write::SimpleFileOptions::default().compression_method(entry.compression()),
            )
            .map_err(crate::parser::parse_err)?;
            zip.write_all(&document).map_err(crate::parser::parse_err)?;
        } else {
            zip.raw_copy_file(entry).map_err(crate::parser::parse_err)?;
        }
    }
    tracing::debug!(
        skipped_imports = warnings.len(),
        "Expanded DOCX imported content"
    );
    Ok((
        Cow::Owned(zip.finish().map_err(crate::parser::parse_err)?.into_inner()),
        warnings,
    ))
}
