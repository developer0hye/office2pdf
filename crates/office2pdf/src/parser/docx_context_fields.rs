use std::cell::RefCell;
use std::collections::HashMap;

use crate::ir::{PAGE_NUMBER_FIELD_CHAR, TOTAL_PAGES_FIELD_CHAR};

/// Running values for the document's `SEQ` fields, one counter per
/// identifier, and whether a story's `PAGE` and `NUMPAGES` fields are
/// computed per page, which only a header or footer story's are.
///
/// Word numbers a caption with `SEQ Figure` or `SEQ Table`, so the number
/// lives in the field rather than in the text. Dropping the field left every
/// caption in the technical brief reading `표` where Word reads `표 26`, and
/// left the prose that cross-references those numbers pointing at nothing
/// (issue #577).
#[derive(Debug, Default)]
pub(in super::super) struct FieldContext {
    sequences: RefCell<HashMap<String, u32>>,
    /// Whether a `PAGE` or `NUMPAGES` field shows the page it lands on rather
    /// than the result Word cached. Set for a header or footer story, whose
    /// tables carry the field in ordinary runs (issue #2338).
    ///
    /// FIXME(body page fields): Word computes a body `PAGE` field per page as
    /// well; the body still prints the cached result.
    evaluates_page_fields: bool,
}

impl FieldContext {
    /// The fields of a header or footer story, which Word computes on every
    /// page the story is drawn on.
    pub(in super::super) fn for_header_footer_story() -> Self {
        Self {
            evaluates_page_fields: true,
            ..Self::default()
        }
    }

    /// A scan of one paragraph's `w:fldChar` fields.
    pub(in super::super) fn page_field_scan(&self) -> PageFieldScan {
        PageFieldScan {
            active: self.evaluates_page_fields,
            open: Vec::new(),
        }
    }

    /// Advance `identifier`'s counter and return its new value.
    ///
    /// Word counts a `SEQ` in document order, so the paragraphs must be
    /// visited in order — which is how the body is converted.
    pub(in super::super) fn next_in_sequence(&self, identifier: &str) -> u32 {
        let mut sequences = self.sequences.borrow_mut();
        let value = sequences.entry(identifier.to_string()).or_insert(0);
        *value += 1;
        *value
    }
}

/// The `w:fldChar` fields open at the current point of a paragraph, read run
/// child by run child, so that a `PAGE` or `NUMPAGES` field's cached result is
/// replaced by the marker the renderer draws the page counter for.
///
/// Word writes a field across several runs — begin, the instruction,
/// separate, the cached result, end — and a field may sit inside another's
/// instruction or result, as a header's `IF` holding a `DOCPROPERTY` does. A
/// stack keeps each `end` closing its own field, so a field this does not
/// model keeps its cached text.
#[derive(Debug, Default)]
pub(in super::super) struct PageFieldScan {
    /// False outside a header or footer story: every child is shown.
    active: bool,
    open: Vec<OpenField>,
}

#[derive(Debug)]
struct OpenField {
    instruction: FieldInstruction,
    past_separate: bool,
    /// Whether the marker standing for the result has been written.
    marked: bool,
}

impl OpenField {
    fn marker(&self) -> Option<char> {
        match self.instruction {
            FieldInstruction::Page => Some(PAGE_NUMBER_FIELD_CHAR),
            FieldInstruction::TotalPages => Some(TOTAL_PAGES_FIELD_CHAR),
            FieldInstruction::Unread | FieldInstruction::Other => None,
        }
    }

    /// Whether this is a page field whose cached result is being read.
    fn hides_result(&self) -> bool {
        self.past_separate && self.marker().is_some()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldInstruction {
    /// No instruction word read yet.
    Unread,
    Page,
    TotalPages,
    Other,
}

/// What a text, tab or break child shows where [`PageFieldScan`] stands.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in super::super) enum FieldContent {
    Shown,
    /// Part of a page field's cached result.
    Hidden,
    /// The first child of a page field's cached result, which the field's
    /// marker replaces so the result keeps that run's properties.
    Marker(char),
}

impl PageFieldScan {
    /// Open a field. One inside another field's instruction is part of that
    /// field's code, not of the page: an `IF` that holds `PAGE` in its text
    /// arguments shows only its own result. Such a field is never modelled.
    pub(in super::super) fn begin(&mut self) {
        if self.active {
            let in_instruction: bool = self.open.iter().any(|field| !field.past_separate);
            self.open.push(OpenField {
                instruction: if in_instruction {
                    FieldInstruction::Other
                } else {
                    FieldInstruction::Unread
                },
                past_separate: false,
                marked: false,
            });
        }
    }

    /// Read the innermost field's instruction from its first word, as Word
    /// does: `PAGE   \* MERGEFORMAT` is a `PAGE` field, `PAGEREF` is not.
    pub(in super::super) fn instruction(&mut self, instruction: &str) {
        let Some(field) = self.open.last_mut() else {
            return;
        };
        if field.instruction != FieldInstruction::Unread || field.past_separate {
            return;
        }
        let Some(word) = instruction.split_whitespace().next() else {
            return;
        };
        field.instruction = if word.eq_ignore_ascii_case("PAGE") {
            FieldInstruction::Page
        } else if word.eq_ignore_ascii_case("NUMPAGES") {
            FieldInstruction::TotalPages
        } else {
            FieldInstruction::Other
        };
    }

    pub(in super::super) fn separate(&mut self) {
        if let Some(field) = self.open.last_mut() {
            field.past_separate = true;
        }
    }

    /// Close the innermost field. A page field that showed no cached result —
    /// a dirty field, or one with an empty result — returns its marker, to be
    /// written where it ends.
    pub(in super::super) fn end(&mut self) -> Option<char> {
        let field: OpenField = self.open.pop()?;
        if field.marked || self.open.iter().any(OpenField::hides_result) {
            return None;
        }
        field.marker()
    }

    /// What the next text, tab or break child shows.
    pub(in super::super) fn content(&mut self) -> FieldContent {
        let Some(field) = self.open.iter_mut().find(|field| field.hides_result()) else {
            return FieldContent::Shown;
        };
        if field.marked {
            return FieldContent::Hidden;
        }
        field.marked = true;
        field
            .marker()
            .map_or(FieldContent::Hidden, FieldContent::Marker)
    }
}

/// The identifier a `SEQ` field instruction counts, if it is one.
///
/// The instruction is `SEQ <identifier>` followed by optional switches —
/// `\* ARABIC` for the numeral format, `\n`/`\c`/`\r` for what to do with the
/// counter. Only the identifier is read here; a switch that resets or repeats
/// the counter is rare in a caption and would need the whole field grammar.
pub(in super::super) fn seq_identifier(instruction: &str) -> Option<&str> {
    let rest = instruction.trim().strip_prefix("SEQ")?;
    // `SEQ` has to be a whole word: `SEQUENCE Table` is not a sequence field.
    if !rest.starts_with(char::is_whitespace) {
        return None;
    }
    let identifier = rest.split_whitespace().next()?;
    (!identifier.starts_with('\\')).then_some(identifier)
}

/// The outline depth a `TOC` field collects, if it is one.
///
/// `TOC \o "1-3"` takes headings of outline level 1 through 3. The upper bound
/// is what Typst's outline depth means, and Word's own default when the range
/// is absent is the whole outline. A `\a` list collects captions rather than
/// headings, so it reads as `None` here.
pub(in super::super) fn toc_heading_depth(instruction: &str) -> Option<u8> {
    let rest = instruction.trim().strip_prefix("TOC")?;
    if !rest.is_empty() && !rest.starts_with(char::is_whitespace) {
        return None;
    }
    if rest.contains("\\a") {
        return None;
    }
    let Some(after) = rest.split("\\o").nth(1) else {
        return Some(FULL_OUTLINE_DEPTH);
    };
    let range = after
        .trim()
        .trim_start_matches('"')
        .split('"')
        .next()
        .unwrap_or_default();
    let upper = range.split('-').nth(1).unwrap_or(range).trim();
    Some(
        upper
            .parse::<u8>()
            .unwrap_or(FULL_OUTLINE_DEPTH)
            .clamp(1, FULL_OUTLINE_DEPTH),
    )
}

/// Word's outline runs from level 1 to level 9.
const FULL_OUTLINE_DEPTH: u8 = 9;

/// The `SEQ` identifier a `TOC \a` field lists, if it is one.
///
/// `TOC \a "Figure"` collects the paragraphs `SEQ Figure` numbers — the
/// document's figure contents — rather than its headings.
pub(in super::super) fn toc_caption_identifier(instruction: &str) -> Option<String> {
    let rest = instruction.trim().strip_prefix("TOC")?;
    if !rest.is_empty() && !rest.starts_with(char::is_whitespace) {
        return None;
    }
    let after = rest.split("\\a").nth(1)?;
    let identifier = after
        .trim_start()
        .trim_start_matches('"')
        .split('"')
        .next()
        .unwrap_or_default()
        .split_whitespace()
        .next()
        .unwrap_or_default();
    (!identifier.is_empty()).then(|| identifier.to_string())
}

#[cfg(test)]
#[path = "docx_context_fields_tests.rs"]
mod tests;
