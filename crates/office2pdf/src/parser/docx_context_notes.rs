use crate::parser::xml_util::OOXML_XML_VERSION;
use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use super::super::extract_run_text;
#[cfg(test)]
use super::super::text::parse_theme_fonts;
use super::super::text::{ThemeFonts, resolve_latin_theme_font_family};
use crate::ir::TextStyle;
use crate::parser::units::half_points_to_pt;
use crate::parser::xml_util::parse_hex_color;

// ── Footnote / Endnote support ──────────────────────────────────────────

#[derive(Debug, Clone, Copy)]
enum NoteKind {
    Footnote,
    Endnote,
}

/// One run of a note's text, carrying only what its own `w:rPr` states.
///
/// The style it resolves against is not known here: notes are read from the
/// archive before the stylesheet is, so the cascade runs at reference time.
#[derive(Debug, Clone, Default)]
pub(in super::super) struct NoteRun {
    pub(in super::super) text: String,
    pub(in super::super) explicit: TextStyle,
}

/// A note's content: the paragraph style it names, and its runs.
#[derive(Debug, Clone, Default)]
pub(in super::super) struct NoteContent {
    pub(in super::super) style_id: Option<String>,
    pub(in super::super) runs: Vec<NoteRun>,
}

/// Context for resolving footnote/endnote references during parsing.
/// The `cursor` is advanced each time a note reference run is encountered.
pub(in super::super) struct NoteContext {
    footnote_content: HashMap<usize, NoteContent>,
    endnote_content: HashMap<usize, NoteContent>,
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

    pub(in super::super) fn consume_next(&self) -> Option<NoteContent> {
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
    theme_fonts: &ThemeFonts,
) -> NoteContext {
    let mut note_context = NoteContext::empty();

    if let Some(xml) = read_zip_text(archive, "word/footnotes.xml") {
        note_context.footnote_content = parse_notes_xml(&xml, theme_fonts);
    }
    if let Some(xml) = read_zip_text(archive, "word/endnotes.xml") {
        note_context.endnote_content = parse_notes_xml(&xml, theme_fonts);
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

/// Read `footnotes.xml` or `endnotes.xml` into per-note styled runs.
///
/// The note's `w:pStyle` and each run's `w:rPr` are what Word styles the text
/// with; flattening the part to one unstyled string left every note at the
/// rendering engine's own footnote size and face — 9.35pt Libertinus Serif
/// against the 8pt Calibri the document asks for (issue #580).
///
/// docx-rs does not read these parts, so this stays a scan; it reads the run
/// properties the cascade needs — weight, slant, size, colour, and family —
/// and leaves the rest to the style the note names.
fn parse_notes_xml(xml: &str, theme_fonts: &ThemeFonts) -> HashMap<usize, NoteContent> {
    let mut map: HashMap<usize, NoteContent> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut current_id: Option<usize> = None;
    let mut current: NoteContent = NoteContent::default();
    let mut run: NoteRun = NoteRun::default();
    let mut in_text = false;
    let mut in_run_property = false;

    let attribute_value =
        |element: &quick_xml::events::BytesStart, name: &[u8]| -> Option<String> {
            element.attributes().flatten().find_map(|attribute| {
                (attribute.key.local_name().as_ref() == name)
                    .then(|| {
                        attribute
                            .normalized_value(OOXML_XML_VERSION)
                            .ok()
                            .map(|value| value.to_string())
                    })
                    .flatten()
            })
        };
    // `w:b`, `w:i`, and the rest are on when present unless they say `w:val="0"`.
    let toggle_on = |element: &quick_xml::events::BytesStart| -> bool {
        !matches!(
            attribute_value(element, b"val").as_deref(),
            Some("0") | Some("false")
        )
    };

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"footnote" | b"endnote" => {
                        if let Some(id) = current_id.take() {
                            finish_note(&mut map, id, std::mem::take(&mut current));
                        }
                        current = NoteContent::default();
                        run = NoteRun::default();
                        current_id = attribute_value(element, b"id")
                            .and_then(|value| value.parse::<usize>().ok());
                    }
                    b"pStyle" => {
                        if current.style_id.is_none() {
                            current.style_id = attribute_value(element, b"val");
                        }
                    }
                    b"rPr" => in_run_property = true,
                    b"b" if in_run_property => run.explicit.bold = Some(toggle_on(element)),
                    b"i" if in_run_property => run.explicit.italic = Some(toggle_on(element)),
                    b"sz" if in_run_property => {
                        run.explicit.font_size = attribute_value(element, b"val")
                            .and_then(|value| value.parse::<f64>().ok())
                            .map(half_points_to_pt);
                    }
                    b"color" if in_run_property => {
                        run.explicit.color = attribute_value(element, b"val")
                            .as_deref()
                            .and_then(parse_hex_color);
                    }
                    b"rFonts" if in_run_property => {
                        run.explicit.font_family = attribute_value(element, b"ascii")
                            .or_else(|| attribute_value(element, b"hAnsi"))
                            .or_else(|| attribute_value(element, b"eastAsia"));
                        if run.explicit.font_family.is_none() {
                            let run_property = serde_json::json!({
                                "fonts": {
                                    "asciiTheme": attribute_value(element, b"asciiTheme"),
                                    "hiAnsiTheme": attribute_value(element, b"hAnsiTheme"),
                                }
                            });
                            run.explicit.font_family =
                                resolve_latin_theme_font_family(&run_property, theme_fonts);
                        }
                    }
                    b"t" => in_text = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"t" => in_text = false,
                b"rPr" => in_run_property = false,
                b"r" => {
                    let finished = std::mem::take(&mut run);
                    if !finished.text.is_empty() {
                        current.runs.push(finished);
                    }
                }
                b"footnote" | b"endnote" => {
                    if let Some(id) = current_id.take() {
                        finish_note(&mut map, id, std::mem::take(&mut current));
                    }
                    current = NoteContent::default();
                    run = NoteRun::default();
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref element)) => {
                if in_text && let Ok(text) = element.xml_content(OOXML_XML_VERSION) {
                    run.text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Record a note, dropping one whose runs carry no text at all.
///
/// The leading run of a Word note holds only the reference mark
/// (`w:footnoteRef`), so a note that separates cleanly and a note that is
/// genuinely empty both arrive here with nothing but that run.
fn finish_note(map: &mut HashMap<usize, NoteContent>, id: usize, mut content: NoteContent) {
    if let Some(last) = content.runs.last_mut() {
        last.text = last.text.trim_end().to_string();
    }
    if let Some(first) = content.runs.first_mut() {
        first.text = first.text.trim_start().to_string();
    }
    content.runs.retain(|run| !run.text.is_empty());
    if !content.runs.is_empty() {
        map.insert(id, content);
    }
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
                            && let Ok(value) = attribute.normalized_value(OOXML_XML_VERSION)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn note_run_latin_theme_font_resolves_from_document_theme() {
        let theme_fonts: ThemeFonts = parse_theme_fonts(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Office"><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
        );
        let notes = parse_notes_xml(
            r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr><w:t>note text</w:t></w:r></w:p></w:footnote></w:footnotes>"#,
            &theme_fonts,
        );

        assert_eq!(
            notes[&2].runs[0].explicit.font_family.as_deref(),
            Some("Calibri")
        );
    }
}
