//! What docx-rs drops from each run, recovered from the raw document XML.
//!
//! - The run's own `w14:ligatures` (issue #2205). docx-rs drops the element
//!   from `RunProperty`, so `LigatureRules` can only read the levels above the
//!   run. Word writes it on a run when a user picks ligatures in the Font
//!   dialog, and keeps it on runs pasted from a document whose defaults turn
//!   them on.
//! - Each line break's `w:clear` (issue #2230). docx-rs keeps only a `w:br`'s
//!   `w:type`, so a break that resumes text below a floating table reached the
//!   converter as a plain one.
//!
//! TODO(docx-rs drops w:clear): read it from `docx_rs::Break` once the
//! published fork carries it.
//!
//! The scan counts the paragraph sequence of [`super::paragraph_cursor`] and,
//! inside each paragraph, the runs docx-rs hands the converter in document
//! order: every `w:r`, except those in `w:del` and `w:moveFrom`, which
//! `flatten_tracked_changes` drops, and math. A positional run cursor still
//! drifts wherever the two sides count runs differently, which is why
//! `PairKerningRules` leaves the run-level `w:kern` unread. Each scanned run
//! therefore keeps its `w:t` text, and the converter takes a value only from
//! the entry whose text matches its own run. The first mismatch stops the rest
//! of that paragraph, which leaves those runs with their styles' answer
//! rather than another run's.

use std::cell::Cell;

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::ir::{BreakClear, Ligatures};
use crate::parser::xml_util::{OOXML_XML_VERSION, decode_general_ref};

use super::super::styles::parse_ligatures_value;
use super::super::text::is_line_break;
use super::paragraph_cursor::{
    ParagraphCursor, ScannedParagraph, attribute_value, is_skipped_subtree,
};

/// A `w:r` as the scan saw it.
#[derive(Debug, Default, Clone, PartialEq)]
pub(in super::super) struct ScannedRun {
    /// The run's `w:t` content with whitespace removed, to check the converter
    /// is on the same run.
    text_key: String,
    /// The run's own `w:rPr/w14:ligatures`, or `None` when it states none.
    ligatures: Option<Ligatures>,
    /// The `w:clear` of each of the run's line breaks, in order: every `w:br`
    /// but a page or column break, which never clears.
    line_break_clears: Vec<Option<BreakClear>>,
}

/// What the scan read for one converted run.
#[derive(Debug, Default, Clone, Copy, PartialEq)]
pub(in super::super) struct RunScan<'a> {
    pub(in super::super) ligatures: Option<Ligatures>,
    /// One entry per line break of the run, in order, or none at all when
    /// the scan cannot vouch for them.
    pub(in super::super) line_break_clears: &'a [Option<BreakClear>],
}

/// A `w:br`'s `w:clear`; `none`, an absent value and an unknown one clear
/// nothing.
fn parse_break_clear(value: Option<&str>) -> Option<BreakClear> {
    match value? {
        "left" => Some(BreakClear::Left),
        "right" => Some(BreakClear::Right),
        "all" => Some(BreakClear::All),
        _ => None,
    }
}

/// Subtrees whose runs never reach run conversion. The paragraphs inside them
/// still count, so the paragraph sequence stays the shared one.
fn drops_runs(local_name: &[u8]) -> bool {
    matches!(local_name, b"del" | b"moveFrom" | b"oMath" | b"oMathPara")
}

/// The text a run is matched on: its `w:t` content without whitespace, which
/// docx-rs and the raw part do not always preserve alike.
fn text_key(text: &str) -> impl Iterator<Item = char> + '_ {
    text.chars().filter(|ch| !ch.is_whitespace())
}

pub(in super::super) struct RunScanContext {
    paragraphs: ParagraphCursor<Vec<ScannedRun>>,
}

impl RunScanContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            paragraphs: ParagraphCursor::new("run scan", xml.map(Self::scan).unwrap_or_default()),
        }
    }

    /// The runs of the next paragraph. Must be called exactly once per
    /// converted `w:p`, with that paragraph's `w:pStyle`.
    pub(in super::super) fn next_paragraph(&self, style_id: Option<&str>) -> ParagraphRunScan<'_> {
        ParagraphRunScan {
            runs: self
                .paragraphs
                .next(style_id)
                .map(|(_, runs)| runs.as_slice())
                .unwrap_or_default(),
            cursor: Cell::new(0),
            has_drifted: Cell::new(false),
        }
    }

    fn scan(xml: &str) -> Vec<ScannedParagraph<Vec<ScannedRun>>> {
        let mut reader = Reader::from_str(xml);
        let mut paragraphs: Vec<ScannedParagraph<Vec<ScannedRun>>> = Vec::new();
        // Each open paragraph with the run-dropping depth it opened at: a text
        // box inside a deleted run is still a paragraph of its own whose runs
        // convert.
        let mut open_paragraphs: Vec<(usize, usize)> = Vec::new();
        // Each open `w:r`, as the paragraph and run its content belongs to, or
        // `None` for one conversion never reaches.
        let mut open_runs: Vec<Option<(usize, usize)>> = Vec::new();
        let mut in_body = false;
        let mut in_paragraph_properties = false;
        let mut in_run_properties = false;
        let mut in_text = false;
        let mut skipped_depth: usize = 0;
        let mut dropped_depth: usize = 0;

        // The run whose own content is being read: the innermost open `w:r`,
        // when it belongs to the innermost open paragraph.
        let current_run = |open_paragraphs: &[(usize, usize)],
                           open_runs: &[Option<(usize, usize)>]|
         -> Option<(usize, usize)> {
            let (paragraph, run) = (*open_runs.last()?)?;
            (open_paragraphs.last()?.0 == paragraph).then_some((paragraph, run))
        };

        loop {
            let (element, is_empty) = match reader.read_event() {
                Ok(Event::Start(element)) => (element, false),
                Ok(Event::Empty(element)) => (element, true),
                Ok(Event::End(element)) => {
                    let name = element.local_name();
                    if skipped_depth > 0 {
                        skipped_depth -= usize::from(is_skipped_subtree(name.as_ref()));
                        continue;
                    }
                    dropped_depth -= usize::from(dropped_depth > 0 && drops_runs(name.as_ref()));
                    match name.as_ref() {
                        b"body" => in_body = false,
                        b"p" if in_body => {
                            open_paragraphs.pop();
                            in_paragraph_properties = false;
                        }
                        b"pPr" => in_paragraph_properties = false,
                        b"r" if in_body => {
                            open_runs.pop();
                            in_run_properties = false;
                        }
                        b"rPr" => in_run_properties = false,
                        b"t" => in_text = false,
                        _ => {}
                    }
                    continue;
                }
                Ok(Event::Text(text)) => {
                    if in_text
                        && let Some((paragraph, run)) = current_run(&open_paragraphs, &open_runs)
                        && let Ok(content) = text.xml_content(OOXML_XML_VERSION)
                    {
                        paragraphs[paragraph].value[run]
                            .text_key
                            .extend(text_key(&content));
                    }
                    continue;
                }
                Ok(Event::GeneralRef(reference)) => {
                    if in_text
                        && let Some((paragraph, run)) = current_run(&open_paragraphs, &open_runs)
                        && let Some(content) = decode_general_ref(&reference)
                    {
                        paragraphs[paragraph].value[run]
                            .text_key
                            .extend(text_key(&content));
                    }
                    continue;
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => continue,
            };
            let name = element.local_name();
            if skipped_depth > 0 || is_skipped_subtree(name.as_ref()) {
                skipped_depth += usize::from(!is_empty && is_skipped_subtree(name.as_ref()));
                continue;
            }
            if drops_runs(name.as_ref()) && !is_empty {
                dropped_depth += 1;
                continue;
            }
            match name.as_ref() {
                b"body" if !is_empty => in_body = true,
                b"p" if in_body => {
                    paragraphs.push(ScannedParagraph {
                        style_id: None,
                        value: Vec::new(),
                    });
                    if !is_empty {
                        open_paragraphs.push((paragraphs.len() - 1, dropped_depth));
                    }
                }
                b"pPr" if !is_empty && !open_paragraphs.is_empty() => {
                    in_paragraph_properties = true;
                }
                b"pStyle" if in_paragraph_properties => {
                    if let Some(&(paragraph, _)) = open_paragraphs.last() {
                        paragraphs[paragraph].style_id = attribute_value(&reader, &element, b"val");
                    }
                }
                b"r" if in_body => {
                    let converted_in: Option<usize> = open_paragraphs
                        .last()
                        .filter(|(_, opened_at)| *opened_at == dropped_depth)
                        .map(|(paragraph, _)| *paragraph);
                    let entry: Option<(usize, usize)> = match (converted_in, open_runs.last()) {
                        (None, _) => None,
                        // A `w:r` nested in its own paragraph's run is part of
                        // that run, as docx-rs reads it.
                        (Some(paragraph), Some(Some((open_paragraph, run))))
                            if *open_paragraph == paragraph =>
                        {
                            Some((paragraph, *run))
                        }
                        (Some(paragraph), _) => {
                            paragraphs[paragraph].value.push(ScannedRun::default());
                            Some((paragraph, paragraphs[paragraph].value.len() - 1))
                        }
                    };
                    if !is_empty {
                        open_runs.push(entry);
                    }
                }
                b"rPr"
                    if !is_empty
                        && !in_paragraph_properties
                        && current_run(&open_paragraphs, &open_runs).is_some() =>
                {
                    in_run_properties = true;
                }
                b"ligatures" if in_run_properties => {
                    if let Some((paragraph, run)) = current_run(&open_paragraphs, &open_runs) {
                        paragraphs[paragraph].value[run].ligatures =
                            attribute_value(&reader, &element, b"val")
                                .and_then(|value| parse_ligatures_value(&value));
                    }
                }
                b"br" => {
                    if let Some((paragraph, run)) = current_run(&open_paragraphs, &open_runs)
                        && !matches!(
                            attribute_value(&reader, &element, b"type").as_deref(),
                            Some("page" | "column")
                        )
                    {
                        paragraphs[paragraph].value[run]
                            .line_break_clears
                            .push(parse_break_clear(
                                attribute_value(&reader, &element, b"clear").as_deref(),
                            ));
                    }
                }
                b"t" if !is_empty && current_run(&open_paragraphs, &open_runs).is_some() => {
                    in_text = true;
                }
                _ => {}
            }
        }

        paragraphs
    }
}

/// One paragraph's scanned runs, handed out one per converted docx-rs run.
pub(in super::super) struct ParagraphRunScan<'a> {
    runs: &'a [ScannedRun],
    cursor: Cell<usize>,
    has_drifted: Cell<bool>,
}

impl<'a> ParagraphRunScan<'a> {
    /// What the scan read for `run`, or nothing where it cannot vouch for it.
    /// Must be called once per docx-rs run of the paragraph, in order, before
    /// the run is split.
    pub(in super::super) fn next_run(&self, run: &docx_rs::Run) -> RunScan<'a> {
        let index: usize = self.cursor.get();
        self.cursor.set(index + 1);
        if self.has_drifted.get() {
            return RunScan::default();
        }
        let Some(scanned) = self.runs.get(index) else {
            return RunScan::default();
        };
        let matches_run: bool = run
            .children
            .iter()
            .filter_map(|child| match child {
                docx_rs::RunChild::Text(text) => Some(text.text.as_str()),
                _ => None,
            })
            .flat_map(text_key)
            .eq(scanned.text_key.chars());
        if !matches_run {
            self.has_drifted.set(true);
            tracing::debug!(
                run_index = index,
                "run scan drifted from the converted runs; leaving the paragraph's \
                 remaining runs to their styles' ligatures and plain line breaks"
            );
            return RunScan::default();
        }
        // A run whose breaks the two sides count differently keeps plain
        // ones: a clear handed to the wrong break would move the wrong line.
        let line_break_count: usize = run
            .children
            .iter()
            .filter(|child| matches!(child, docx_rs::RunChild::Break(br) if is_line_break(br)))
            .count();
        let line_break_clears: &[Option<BreakClear>] =
            if scanned.line_break_clears.len() == line_break_count {
                &scanned.line_break_clears
            } else {
                tracing::debug!(
                    run_index = index,
                    scanned = scanned.line_break_clears.len(),
                    converted = line_break_count,
                    "run scan counts the run's line breaks differently; keeping them plain"
                );
                &[]
            };
        RunScan {
            ligatures: scanned.ligatures,
            line_break_clears,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const STANDARD: Ligatures = Ligatures {
        standard: true,
        contextual: false,
        historical: false,
        discretional: false,
    };

    fn run(text: &str) -> docx_rs::Run {
        docx_rs::Run::new().add_text(text)
    }

    fn scanned(xml: &str) -> Vec<Vec<ScannedRun>> {
        RunScanContext::scan(xml)
            .into_iter()
            .map(|paragraph| paragraph.value)
            .collect()
    }

    fn scanned_run(text_key: &str, ligatures: Option<Ligatures>) -> ScannedRun {
        ScannedRun {
            text_key: text_key.to_string(),
            ligatures,
            line_break_clears: Vec::new(),
        }
    }

    #[test]
    fn reads_each_runs_own_value_in_order() {
        let xml = r#"<w:document xmlns:w="urn:w" xmlns:w14="urn:w14"><w:body>
          <w:p>
            <w:pPr><w:rPr><w14:ligatures w14:val="all"/></w:rPr></w:pPr>
            <w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>fine fit</w:t></w:r>
            <w:r><w:t xml:space="preserve"> office </w:t></w:r>
            <w:hyperlink><w:r><w:rPr><w14:ligatures w14:val="none"/></w:rPr><w:t>R&amp;D</w:t></w:r></w:hyperlink>
          </w:p>
          <w:p/>
        </w:body></w:document>"#;

        assert_eq!(
            scanned(xml),
            vec![
                vec![
                    scanned_run("finefit", Some(STANDARD)),
                    scanned_run("office", None),
                    scanned_run("R&D", Some(Ligatures::NONE)),
                ],
                vec![],
            ],
            "the paragraph mark's own w:rPr is not a run's"
        );
    }

    /// `flatten_tracked_changes` drops deleted and moved-away runs, and a
    /// text box's `mc:Fallback` copy never converts, so none of them may take
    /// a slot from the runs after them.
    #[test]
    fn skips_the_runs_conversion_never_reaches() {
        let xml = r#"<w:document xmlns:w="urn:w" xmlns:w14="urn:w14" xmlns:mc="urn:mc" xmlns:m="urn:m"><w:body>
          <w:p>
            <w:del><w:r><w:rPr><w14:ligatures w14:val="none"/></w:rPr><w:delText>gone</w:delText></w:r></w:del>
            <w:moveFrom><w:r><w:t>moved</w:t></w:r></w:moveFrom>
            <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
            <w:ins><w:r><w:t>kept</w:t></w:r></w:ins>
            <w:r><mc:AlternateContent>
              <mc:Choice><w:drawing><w:txbxContent>
                <w:p><w:r><w:rPr><w14:ligatures w14:val="none"/></w:rPr><w:t>box</w:t></w:r></w:p>
              </w:txbxContent></w:drawing></mc:Choice>
              <mc:Fallback><w:pict><w:txbxContent>
                <w:p><w:r><w:t>box</w:t></w:r></w:p>
              </w:txbxContent></w:pict></mc:Fallback>
            </mc:AlternateContent></w:r>
            <w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>after</w:t></w:r>
          </w:p>
        </w:body></w:document>"#;

        assert_eq!(
            scanned(xml),
            vec![
                vec![
                    scanned_run("kept", None),
                    scanned_run("", None),
                    scanned_run("after", Some(STANDARD)),
                ],
                vec![scanned_run("box", Some(Ligatures::NONE))],
            ]
        );
    }

    #[test]
    fn hands_each_run_its_own_value() {
        let context = RunScanContext::from_xml(Some(
            r#"<w:document xmlns:w="urn:w" xmlns:w14="urn:w14"><w:body>
              <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
                <w:r><w:t>plain</w:t></w:r>
                <w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>fluffy</w:t></w:r>
              </w:p>
            </w:body></w:document>"#,
        ));
        let paragraph = context.next_paragraph(Some("Body"));

        assert_eq!(paragraph.next_run(&run("plain")).ligatures, None);
        assert_eq!(paragraph.next_run(&run("fluffy")).ligatures, Some(STANDARD));
        assert_eq!(
            paragraph.next_run(&run("extra")).ligatures,
            None,
            "past the scan"
        );
    }

    /// A run the scan and docx-rs count differently must not hand its value to
    /// another run: the rest of the paragraph keeps its styles' answer.
    #[test]
    fn stops_at_the_first_run_whose_text_differs() {
        let context = RunScanContext::from_xml(Some(
            r#"<w:document xmlns:w="urn:w" xmlns:w14="urn:w14"><w:body>
              <w:p>
                <w:r><w:rPr><w14:ligatures w14:val="none"/></w:rPr><w:t>first</w:t></w:r>
                <w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>second</w:t></w:r>
              </w:p>
            </w:body></w:document>"#,
        ));
        let paragraph = context.next_paragraph(None);

        assert_eq!(
            paragraph.next_run(&run("second")).ligatures,
            None,
            "out of step"
        );
        assert_eq!(
            paragraph.next_run(&run("second")).ligatures,
            None,
            "a drifted paragraph hands nothing out, even to a run that matches"
        );
    }

    #[test]
    fn a_paragraph_whose_style_disagrees_hands_nothing_out() {
        let context = RunScanContext::from_xml(Some(
            r#"<w:document xmlns:w="urn:w" xmlns:w14="urn:w14"><w:body>
              <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>
                <w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>fit</w:t></w:r>
              </w:p>
            </w:body></w:document>"#,
        ));

        assert_eq!(
            context.next_paragraph(None).next_run(&run("fit")).ligatures,
            None
        );
    }

    /// Every line break of a run takes an entry, so the converter can pair
    /// them in order; a page or column break is no line break and takes none.
    #[test]
    fn reads_each_line_breaks_clear_in_order() {
        let xml = r#"<w:document xmlns:w="urn:w"><w:body>
          <w:p>
            <w:r><w:t>One</w:t><w:br/><w:t>Two</w:t><w:br w:type="textWrapping" w:clear="all"/></w:r>
            <w:r><w:br w:type="page" w:clear="all"/><w:br w:clear="none"/><w:br w:clear="left"/><w:br w:clear="right"/></w:r>
            <w:del><w:r><w:br w:clear="all"/></w:r></w:del>
          </w:p>
        </w:body></w:document>"#;

        let runs: Vec<Vec<Option<BreakClear>>> = scanned(xml)
            .remove(0)
            .into_iter()
            .map(|run| run.line_break_clears)
            .collect();
        assert_eq!(
            runs,
            vec![
                vec![None, Some(BreakClear::All)],
                vec![None, Some(BreakClear::Left), Some(BreakClear::Right)],
            ]
        );
    }

    #[test]
    fn hands_out_a_runs_clears_only_when_both_sides_count_its_breaks_alike() {
        let context = RunScanContext::from_xml(Some(
            r#"<w:document xmlns:w="urn:w"><w:body>
              <w:p>
                <w:r><w:t>Before</w:t><w:br w:clear="all"/></w:r>
                <w:r><w:t>After</w:t><w:br w:clear="all"/></w:r>
              </w:p>
            </w:body></w:document>"#,
        ));
        let paragraph = context.next_paragraph(None);
        let with_break = |text: &str, breaks: usize| {
            let mut run = run(text);
            for _ in 0..breaks {
                run = run.add_break(docx_rs::BreakType::TextWrapping);
            }
            run
        };

        assert_eq!(
            paragraph
                .next_run(&with_break("Before", 1))
                .line_break_clears,
            &[Some(BreakClear::All)]
        );
        assert!(
            paragraph
                .next_run(&with_break("After", 2))
                .line_break_clears
                .is_empty(),
            "a break the scan did not see"
        );
    }
}
