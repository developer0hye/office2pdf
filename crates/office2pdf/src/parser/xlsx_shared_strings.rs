//! Excel's per-cell character limit, enforced on the shared string table.
//!
//! A shared string is stored once and referenced by every cell that shows it,
//! so one over-long item amplifies: the reader copies the item's text into
//! each referencing cell, and `poc-shared-strings.xlsx` points 12,000 cells at
//! a single 1 MiB string. Reading it alone needed about 12 GiB before layout
//! began (issue #1702). Excel itself stores at most 32,767 characters in a
//! cell, so an item past that is longer than any cell Excel writes — reject
//! the package while only the table itself has been read. What Excel's reader
//! does with such an item is untested here; the limit is the documented
//! storage maximum, not an observed repair behaviour.

use std::io::BufRead;

use quick_xml::events::Event;

use crate::error::ConvertError;
use crate::parser::xml_util::{self, OOXML_XML_VERSION};

/// The most characters Excel stores in one cell.
/// <https://support.microsoft.com/en-gb/excel/excel-specifications-and-limits>
const EXCEL_MAX_CELL_CHARACTERS: usize = 32_767;

/// Where a workbook keeps its shared string table.
const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

/// Reject a workbook whose shared string table holds an item longer than a
/// cell can be, before the reader multiplies it across referencing cells.
///
/// A package with no readable table is left to the reader, which reports its
/// own error for a malformed archive.
pub(crate) fn validate_shared_string_lengths(data: &[u8]) -> Result<(), ConvertError> {
    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return Ok(());
    };
    let Some(entry) = shared_strings_entry_name(&mut archive) else {
        return Ok(());
    };
    let Ok(part) = archive.by_name(&entry) else {
        return Ok(());
    };
    // One character costs at least one byte, so a table whose whole part is no
    // larger than the limit cannot hold an item past it. This bounds the size
    // of the table, not the length of any item, so a workbook with many short
    // strings still falls through to the scan and comes back clean — across
    // the `fixtures-v1` corpus the widest item outside this PoC is 6,817
    // characters, a fifth of the limit.
    if part.size() <= EXCEL_MAX_CELL_CHARACTERS as u64 {
        return Ok(());
    }

    let Some(item) = first_item_past_cell_limit(std::io::BufReader::new(part)) else {
        return Ok(());
    };
    tracing::warn!(
        item,
        limit = EXCEL_MAX_CELL_CHARACTERS,
        "XLSX shared string is longer than an Excel cell"
    );
    Err(crate::parser::parse_err(format!(
        "XLSX shared string {item} is longer than Excel's {EXCEL_MAX_CELL_CHARACTERS}-character cell limit"
    )))
}

/// The archive entry holding the shared string table, resolved the way the
/// reader resolves it: the exact name, then the backslash-separated and
/// case-folded spellings some Windows generators write (umya's `zip_by_name`).
/// Accepting only the exact name would leave a package the reader still reads
/// unguarded, and two workbooks in the `fixtures-v1` corpus spell the part
/// `xl\sharedstrings.xml`.
fn shared_strings_entry_name<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Option<String> {
    if archive.by_name(SHARED_STRINGS_PART).is_ok() {
        return Some(SHARED_STRINGS_PART.to_string());
    }
    let backslash_spelling: String = SHARED_STRINGS_PART.replace('/', "\\");
    if archive.by_name(&backslash_spelling).is_ok() {
        return Some(backslash_spelling);
    }
    let folded: String = SHARED_STRINGS_PART.to_ascii_lowercase();
    for index in 0..archive.len() {
        let Ok(entry) = archive.by_index(index) else {
            continue;
        };
        let Ok(name) = entry.name() else {
            continue;
        };
        let name: String = name.into_owned();
        if name.replace('\\', "/").to_ascii_lowercase() == folded {
            return Some(name);
        }
    }
    None
}

/// The 1-based position of the first `<si>` whose text passes the limit.
///
/// Counts what a referencing cell receives: the item's own `<t>` plus the
/// `<t>` of every rich-text `<r>`, since Excel concatenates the runs into one
/// cell value. `<rPh>` phonetic guides are excluded — they annotate the
/// reading of the base text rather than adding to it.
///
/// Inside a `<t>` everything counts, not only character data, and no nested
/// element is read as structure. The reader takes the element's whole raw span
/// and then strips CDATA delimiters and resolves entity references, so a
/// payload hidden in `<![CDATA[…]]>`, padded with nested markup, or split
/// around a nested `<si>` still reaches every referencing cell. Markup is
/// counted by its own bytes without its angle brackets, which under-counts
/// rather than over-counts — no valid item can be rejected for markup it does
/// not have.
///
/// Returns at the first offending item, so an amplifying table costs one item
/// to reject rather than a full pass. A malformed part yields `None` and is
/// left to the reader's own error.
fn first_item_past_cell_limit<R: BufRead>(part: R) -> Option<usize> {
    let mut reader: quick_xml::Reader<R> = quick_xml::Reader::from_reader(part);
    let mut buffer: Vec<u8> = Vec::new();
    let mut item: usize = 0;
    let mut characters: usize = 0;
    let mut in_text: bool = false;
    let mut phonetic_depth: usize = 0;

    loop {
        buffer.clear();
        let Ok(event) = reader.read_event_into(&mut buffer) else {
            return None;
        };
        // Structure is only read outside a `<t>`. A `<t>` holds no `<si>` or
        // `<rPh>` child in a valid package, and honouring one inside it would
        // let an item reset its own count part way through its text.
        match &event {
            Event::Start(e) if !in_text => match e.local_name().as_ref() {
                b"si" => {
                    item += 1;
                    characters = 0;
                    phonetic_depth = 0;
                }
                b"rPh" => phonetic_depth += 1,
                // The opening tag itself is not part of the cell value.
                b"t" if phonetic_depth == 0 => {
                    in_text = true;
                    continue;
                }
                _ => {}
            },
            Event::End(e) => match e.local_name().as_ref() {
                b"t" if in_text => {
                    in_text = false;
                    continue;
                }
                b"rPh" if !in_text => phonetic_depth = phonetic_depth.saturating_sub(1),
                _ => {}
            },
            Event::Eof => return None,
            _ => {}
        }
        if !in_text {
            continue;
        }

        characters += match &event {
            // The reader splits a text node at every entity or character
            // reference, so both event kinds contribute to one item's count.
            Event::Text(text) => text
                .xml_content(OOXML_XML_VERSION)
                .map(|decoded| decoded.chars().count())
                .unwrap_or(0),
            Event::GeneralRef(reference) => xml_util::decode_general_ref(reference)
                .map(|decoded| decoded.chars().count())
                .unwrap_or(0),
            // CDATA reaches the cell with its delimiters removed.
            Event::CData(cdata) => String::from_utf8_lossy(cdata.as_ref()).chars().count(),
            // Anything else nested in a `<t>` is markup the reader keeps
            // verbatim; `buffer` holds this event's own bytes.
            _ => String::from_utf8_lossy(&buffer).chars().count(),
        };
        if characters > EXCEL_MAX_CELL_CHARACTERS {
            return Some(item);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        EXCEL_MAX_CELL_CHARACTERS, first_item_past_cell_limit, validate_shared_string_lengths,
    };

    fn shared_string_table(items: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{items}</sst>"#
        )
    }

    fn first_offender(items: &str) -> Option<usize> {
        first_item_past_cell_limit(shared_string_table(items).as_bytes())
    }

    #[test]
    fn the_limit_itself_is_a_length_a_cell_holds() {
        let item: String = format!("<si><t>{}</t></si>", "x".repeat(EXCEL_MAX_CELL_CHARACTERS));
        assert_eq!(first_offender(&item), None);
    }

    #[test]
    fn one_character_past_the_limit_is_reported() {
        let item: String = format!(
            "<si><t>{}</t></si>",
            "x".repeat(EXCEL_MAX_CELL_CHARACTERS + 1)
        );
        assert_eq!(first_offender(&item), Some(1));
    }

    #[test]
    fn a_later_item_reports_its_own_position() {
        let items: String = format!(
            "<si><t>short</t></si><si><t>also short</t></si><si><t>{}</t></si>",
            "x".repeat(EXCEL_MAX_CELL_CHARACTERS + 1)
        );
        assert_eq!(first_offender(&items), Some(3));
    }

    #[test]
    fn rich_text_runs_add_up_into_one_cell_value() {
        // Excel concatenates an item's runs, so two halves that each fit still
        // overflow the cell together.
        let half: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS / 2 + 1);
        let item: String = format!("<si><r><t>{half}</t></r><r><t>{half}</t></r></si>");
        assert_eq!(first_offender(&item), Some(1));

        let fitting: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS / 2);
        let within: String = format!("<si><r><t>{fitting}</t></r><r><t>{fitting}</t></r></si>");
        assert_eq!(first_offender(&within), None);
    }

    #[test]
    fn a_phonetic_guide_does_not_count_toward_the_cell() {
        // `<rPh>` carries the reading of the base text, which Excel stores
        // outside the cell's own character budget.
        let base: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS);
        let reading: String = "ホ".repeat(EXCEL_MAX_CELL_CHARACTERS);
        let item: String = format!(
            r#"<si><t>{base}</t><rPh sb="0" eb="1"><t>{reading}</t></rPh><phoneticPr fontId="1"/></si>"#
        );
        assert_eq!(first_offender(&item), None);
    }

    #[test]
    fn the_budget_counts_characters_rather_than_bytes() {
        // A three-byte character is still one character to Excel, so a table
        // far past the limit in bytes can be a workbook that opens.
        let item: String = format!("<si><t>{}</t></si>", "한".repeat(EXCEL_MAX_CELL_CHARACTERS));
        assert!(item.len() > EXCEL_MAX_CELL_CHARACTERS * 2);
        assert_eq!(first_offender(&item), None);
    }

    #[test]
    fn an_entity_reference_contributes_the_character_it_encodes() {
        // The reader hands back `&amp;` as its own event, so a count that read
        // only text events would let an item past the limit through.
        let filler: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS - 1);
        let item: String = format!("<si><t>{filler}&amp;&amp;</t></si>");
        assert_eq!(first_offender(&item), Some(1));

        let exact: String = format!("<si><t>{filler}&amp;</t></si>");
        assert_eq!(first_offender(&exact), None);
    }

    #[test]
    fn a_malformed_table_is_left_to_the_reader() {
        assert_eq!(first_offender("<si><t>unterminated"), None);
    }

    #[test]
    fn an_empty_table_reports_nothing() {
        assert_eq!(first_offender(""), None);
    }

    #[test]
    fn a_cdata_payload_counts_as_the_text_it_carries() {
        // The reader takes the element's raw span and strips the delimiters,
        // so a count that read only text events would let a CDATA-wrapped
        // megabyte reach every referencing cell.
        let payload: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS + 1);
        let item: String = format!("<si><t><![CDATA[{payload}]]></t></si>");
        assert_eq!(first_offender(&item), Some(1));

        let fitting: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS);
        let within: String = format!("<si><t><![CDATA[{fitting}]]></t></si>");
        assert_eq!(first_offender(&within), None);
    }

    #[test]
    fn a_nested_item_tag_cannot_reset_a_running_count() {
        // Reading `<si>` as structure while inside a `<t>` would zero the
        // count part way through the text and hide the rest of the payload.
        let half: String = "x".repeat(EXCEL_MAX_CELL_CHARACTERS / 2 + 1);
        let item: String = format!("<si><t>{half}<si>{half}</t></si>");
        assert_eq!(first_offender(&item), Some(1));
    }

    #[test]
    fn markup_nested_in_a_text_element_counts_too() {
        // `<t>` holds character data in a valid package, but the reader keeps
        // whatever raw span it finds, so padding with elements would otherwise
        // carry an unbounded payload past the guard.
        let padding: String = "<b/>".repeat(EXCEL_MAX_CELL_CHARACTERS);
        let item: String = format!("<si><t>{padding}</t></si>");
        assert_eq!(first_offender(&item), Some(1));
    }

    /// A package holding nothing but a shared string table under `part`.
    /// `validate_shared_string_lengths` reads no other part, and the reader
    /// resolves this one by several spellings.
    fn package_with_shared_strings(part: &str, items: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        zip.start_file(part, zip::write::SimpleFileOptions::default())
            .expect("writable entry");
        std::io::Write::write_all(&mut zip, shared_string_table(items).as_bytes())
            .expect("writable entry body");
        zip.finish().expect("finished zip").into_inner()
    }

    #[test]
    fn every_spelling_the_reader_resolves_is_guarded() {
        // umya's `zip_by_name` tolerates a backslash separator and a folded
        // case, so a guard that accepted only the canonical name would leave
        // those packages amplifying.
        let item: String = format!(
            "<si><t>{}</t></si>",
            "x".repeat(EXCEL_MAX_CELL_CHARACTERS + 1)
        );
        for part in [
            "xl/sharedStrings.xml",
            "xl\\sharedStrings.xml",
            "xl\\sharedstrings.xml",
            "XL/SHAREDSTRINGS.XML",
        ] {
            let package = package_with_shared_strings(part, &item);
            assert!(
                validate_shared_string_lengths(&package).is_err(),
                "{part} must be guarded"
            );
        }
    }

    #[test]
    fn a_package_without_a_shared_string_table_is_left_alone() {
        let package = package_with_shared_strings("xl/styles.xml", "");
        assert!(validate_shared_string_lengths(&package).is_ok());
        assert!(validate_shared_string_lengths(b"not a zip archive").is_ok());
    }

    #[test]
    fn the_error_names_the_item_and_the_limit() {
        let item: String = format!(
            "<si><t>short</t></si><si><t>{}</t></si>",
            "x".repeat(EXCEL_MAX_CELL_CHARACTERS + 1)
        );
        let package = package_with_shared_strings("xl/sharedStrings.xml", &item);
        let message: String = validate_shared_string_lengths(&package)
            .expect_err("the second item is past the limit")
            .to_string();
        assert!(
            message.contains("shared string 2")
                && message.contains("Excel's 32767-character cell limit"),
            "{message}"
        );
    }
}
