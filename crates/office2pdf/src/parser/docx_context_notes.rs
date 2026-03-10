use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use super::super::extract_run_text;

// ── Footnote / Endnote support ──────────────────────────────────────────

#[derive(Debug, Clone, Copy)]
enum NoteKind {
    Footnote,
    Endnote,
}

/// Context for resolving footnote/endnote references during parsing.
/// The `cursor` is advanced each time a note reference run is encountered.
pub(in super::super) struct NoteContext {
    footnote_content: HashMap<usize, String>,
    endnote_content: HashMap<usize, String>,
    note_refs: Vec<(NoteKind, usize)>,
    cursor: Cell<usize>,
    note_style_ids: HashSet<String>,
}

impl NoteContext {
    pub(in super::super) fn empty() -> Self {
        let note_style_ids: HashSet<String> = ["FootnoteReference", "EndnoteReference"]
            .iter()
            .map(|style_id| (*style_id).to_string())
            .collect();
        Self {
            footnote_content: HashMap::new(),
            endnote_content: HashMap::new(),
            note_refs: Vec::new(),
            cursor: Cell::new(0),
            note_style_ids,
        }
    }

    pub(in super::super) fn consume_next(&self) -> Option<String> {
        let index = self.cursor.get();
        if index >= self.note_refs.len() {
            return None;
        }
        let (kind, id) = self.note_refs[index];
        self.cursor.set(index + 1);
        match kind {
            NoteKind::Footnote => self.footnote_content.get(&id).cloned(),
            NoteKind::Endnote => self.endnote_content.get(&id).cloned(),
        }
    }

    pub(in super::super) fn populate_style_ids(&mut self, styles: &docx_rs::Styles) {
        for style in &styles.styles {
            if let Ok(name_value) = serde_json::to_value(&style.name)
                && let Some(name_str) = name_value.as_str()
            {
                let lower = name_str.to_lowercase();
                if lower == "footnote reference" || lower == "endnote reference" {
                    self.note_style_ids.insert(style.style_id.clone());
                }
            }
        }
    }
}

pub(in super::super) fn build_note_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> NoteContext {
    let mut note_context = NoteContext::empty();

    if let Some(xml) = read_zip_text(archive, "word/footnotes.xml") {
        note_context.footnote_content = parse_notes_xml(&xml);
    }
    if let Some(xml) = read_zip_text(archive, "word/endnotes.xml") {
        note_context.endnote_content = parse_notes_xml(&xml);
    }
    note_context.note_refs = doc_xml.map(scan_note_refs).unwrap_or_default();

    note_context
}

pub(in super::super) fn read_zip_text(
    archive: &mut zip::ZipArchive<impl Read + Seek>,
    name: &str,
) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut contents = String::new();
    file.read_to_string(&mut contents).ok()?;
    Some(contents)
}

fn parse_notes_xml(xml: &str) -> HashMap<usize, String> {
    let mut map: HashMap<usize, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut current_id: Option<usize> = None;
    let mut current_text = String::new();
    let mut in_text = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"footnote" | b"endnote" => {
                        if let Some(id) = current_id.take() {
                            let text = current_text.trim().to_string();
                            if !text.is_empty() {
                                map.insert(id, text);
                            }
                        }
                        current_text.clear();
                        for attribute in element.attributes().flatten() {
                            if attribute.key.local_name().as_ref() == b"id"
                                && let Ok(value) = attribute.unescape_value()
                            {
                                current_id = value.parse::<usize>().ok();
                            }
                        }
                    }
                    b"t" => in_text = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"t" => in_text = false,
                b"footnote" | b"endnote" => {
                    if let Some(id) = current_id.take() {
                        let text = current_text.trim().to_string();
                        if !text.is_empty() {
                            map.insert(id, text);
                        }
                    }
                    current_text.clear();
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref element)) => {
                if in_text && let Ok(text) = element.xml_content() {
                    if !current_text.is_empty() {
                        current_text.push(' ');
                    }
                    current_text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

fn scan_note_refs(xml: &str) -> Vec<(NoteKind, usize)> {
    let mut refs: Vec<(NoteKind, usize)> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                let kind = match element.local_name().as_ref() {
                    b"footnoteReference" => Some(NoteKind::Footnote),
                    b"endnoteReference" => Some(NoteKind::Endnote),
                    _ => None,
                };
                if let Some(kind) = kind {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"id"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(id) = value.parse::<usize>()
                        {
                            refs.push((kind, id));
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    refs
}

pub(in super::super) fn is_note_reference_run(run: &docx_rs::Run, notes: &NoteContext) -> bool {
    if let Some(ref style) = run.run_property.style
        && notes.note_style_ids.contains(&style.val)
    {
        return extract_run_text(run).is_empty();
    }
    false
}
