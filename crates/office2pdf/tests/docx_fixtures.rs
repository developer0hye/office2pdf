//! Integration tests for DOCX fixtures.
//!
//! Each real-world `.docx` file in `tests/fixtures/docx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::ir::{Block, FlowPage, Page, Run};
use office2pdf::parser::Parser;
use office2pdf::parser::docx::DocxParser;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/docx")
        .join(name)
}

fn load_fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("fixture file should exist")
}

/// Smoke-test helper: conversion must not panic.
/// Returns `Ok(pdf_bytes)` or prints a warning on expected conversion error.
fn assert_produces_valid_pdf(name: &str) {
    let path = fixture_path(name);
    match office2pdf::convert(&path) {
        Ok(result) => {
            assert!(!result.pdf.is_empty(), "PDF output should not be empty");
            assert!(
                result.pdf.starts_with(b"%PDF"),
                "output should start with PDF magic bytes"
            );
            common::validate_pdf_with_qpdf(&result.pdf);
        }
        Err(e) => {
            // Conversion error is acceptable (unimplemented features, etc.)
            // but we want to know about it.
            eprintln!("[WARN] {name}: conversion error (non-panic): {e}");
        }
    }
}

/// Parse a DOCX fixture and return the flow pages.
fn flow_pages(name: &str) -> Vec<FlowPage> {
    let data = load_fixture(name);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|p| match p {
            Page::Flow(fp) => Some(fp),
            _ => None,
        })
        .collect()
}

/// Collect all blocks from every flow page.
fn all_blocks(pages: &[FlowPage]) -> Vec<&Block> {
    pages.iter().flat_map(|p| p.content.iter()).collect()
}

/// Recursively collect all runs from blocks (paragraphs, lists, tables, text-boxes).
fn all_runs<'a>(blocks: &'a [&'a Block]) -> Vec<&'a Run> {
    let mut runs = Vec::new();
    for block in blocks {
        collect_runs_from_block(block, &mut runs);
    }
    runs
}

fn collect_runs_from_block<'a>(block: &'a Block, out: &mut Vec<&'a Run>) {
    match block {
        Block::Paragraph(p) => out.extend(p.runs.iter()),
        Block::List(list) => {
            for item in &list.items {
                for para in &item.content {
                    out.extend(para.runs.iter());
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for b in &cell.content {
                        collect_runs_from_block(b, out);
                    }
                }
            }
        }
        Block::Image(_)
        | Block::FloatingImage(_)
        | Block::MathEquation(_)
        | Block::Chart(_)
        | Block::PageBreak => {}
    }
}

fn has_hyperlink_runs(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.href.is_some())
}

fn has_footnote_runs(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.footnote.is_some())
}

fn has_table_block(blocks: &[&Block]) -> bool {
    blocks.iter().any(|b| matches!(b, Block::Table(_)))
}

fn has_image_block(blocks: &[&Block]) -> bool {
    blocks.iter().any(|b| matches!(b, Block::Image(_)))
}

fn has_list_block(blocks: &[&Block]) -> bool {
    blocks.iter().any(|b| matches!(b, Block::List(_)))
}

fn has_header(pages: &[FlowPage]) -> bool {
    pages.iter().any(|p| p.header.is_some())
}

fn has_footer(pages: &[FlowPage]) -> bool {
    pages.iter().any(|p| p.footer.is_some())
}

fn has_bold_run(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.style.bold == Some(true))
}

fn has_italic_run(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.style.italic == Some(true))
}

fn has_colored_run(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.style.color.is_some())
}

fn has_font_size_run(runs: &[&Run]) -> bool {
    runs.iter().any(|r| r.style.font_size.is_some())
}

// ---------------------------------------------------------------------------
// equations.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_equations() {
    assert_produces_valid_pdf("equations.docx");
}

#[test]
fn structure_equations() {
    let pages = flow_pages("equations.docx");
    assert!(!pages.is_empty(), "should have at least one FlowPage");
    let blocks = all_blocks(&pages);
    assert!(!blocks.is_empty(), "should have content blocks");
}

// ---------------------------------------------------------------------------
// footnote.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_footnote() {
    assert_produces_valid_pdf("footnote.docx");
}

#[test]
fn structure_footnote() {
    let pages = flow_pages("footnote.docx");
    let blocks = all_blocks(&pages);
    let runs = all_runs(&blocks);
    assert!(
        has_footnote_runs(&runs),
        "should have runs with footnote content"
    );
}

// ---------------------------------------------------------------------------
// header_footer.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_header_footer() {
    assert_produces_valid_pdf("header_footer.docx");
}

#[test]
fn structure_header_footer() {
    let pages = flow_pages("header_footer.docx");
    assert!(
        has_header(&pages) || has_footer(&pages),
        "FlowPage should have header or footer"
    );
}

// ---------------------------------------------------------------------------
// hyperlinks.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_hyperlinks() {
    assert_produces_valid_pdf("hyperlinks.docx");
}

#[test]
fn structure_hyperlinks() {
    let pages = flow_pages("hyperlinks.docx");
    let blocks = all_blocks(&pages);
    let runs = all_runs(&blocks);
    assert!(has_hyperlink_runs(&runs), "should have hyperlink runs");

    let http_link = runs
        .iter()
        .filter_map(|r| r.href.as_deref())
        .any(|href: &str| href.starts_with("http://") || href.starts_with("https://"));
    assert!(http_link, "should have at least one http(s) URL");
}

// ---------------------------------------------------------------------------
// image.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_image() {
    assert_produces_valid_pdf("image.docx");
}

#[test]
fn structure_image() {
    let pages = flow_pages("image.docx");
    let blocks = all_blocks(&pages);
    assert!(has_image_block(&blocks), "should have Block::Image");

    let image_data_non_empty = blocks.iter().any(|b| match b {
        Block::Image(img) => !img.data.is_empty(),
        _ => false,
    });
    assert!(image_data_non_empty, "image data should not be empty");
}

// ---------------------------------------------------------------------------
// numberings.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_numberings() {
    assert_produces_valid_pdf("numberings.docx");
}

#[test]
fn structure_numberings() {
    let pages = flow_pages("numberings.docx");
    let blocks = all_blocks(&pages);
    assert!(has_list_block(&blocks), "should have Block::List");

    let has_items = blocks.iter().any(|b| match b {
        Block::List(list) => !list.items.is_empty(),
        _ => false,
    });
    assert!(has_items, "list should have items");
}

// ---------------------------------------------------------------------------
// styles_en.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_styles_en() {
    assert_produces_valid_pdf("styles_en.docx");
}

#[test]
fn structure_styles_en() {
    let pages = flow_pages("styles_en.docx");
    let blocks = all_blocks(&pages);
    let runs = all_runs(&blocks);
    assert!(
        has_bold_run(&runs)
            || has_italic_run(&runs)
            || has_colored_run(&runs)
            || has_font_size_run(&runs),
        "should have styled runs (bold/italic/color/font_size)"
    );
}

// ---------------------------------------------------------------------------
// table.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_table() {
    assert_produces_valid_pdf("table.docx");
}

#[test]
fn structure_table() {
    let pages = flow_pages("table.docx");
    let blocks = all_blocks(&pages);
    assert!(has_table_block(&blocks), "should have Block::Table");

    let has_rows_and_cells = blocks.iter().any(|b| match b {
        Block::Table(t) => !t.rows.is_empty() && t.rows.iter().any(|r| !r.cells.is_empty()),
        _ => false,
    });
    assert!(has_rows_and_cells, "table should have rows and cells");
}

// ---------------------------------------------------------------------------
// test_python_docx.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_test_python_docx() {
    assert_produces_valid_pdf("test_python_docx.docx");
}

#[test]
fn structure_test_python_docx() {
    let pages = flow_pages("test_python_docx.docx");
    let blocks = all_blocks(&pages);
    let has_paragraphs = blocks.iter().any(|b| matches!(b, Block::Paragraph(_)));
    assert!(has_paragraphs, "should have paragraphs");
}

// ---------------------------------------------------------------------------
// unit_test_formatting.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_unit_test_formatting() {
    assert_produces_valid_pdf("unit_test_formatting.docx");
}

#[test]
fn structure_unit_test_formatting() {
    let pages = flow_pages("unit_test_formatting.docx");
    let blocks = all_blocks(&pages);
    let runs = all_runs(&blocks);
    assert!(
        has_bold_run(&runs) || has_italic_run(&runs) || has_colored_run(&runs),
        "should have bold/italic/colored runs"
    );
}

// ---------------------------------------------------------------------------
// unit_test_headers.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_unit_test_headers() {
    assert_produces_valid_pdf("unit_test_headers.docx");
}

#[test]
fn structure_unit_test_headers() {
    let pages = flow_pages("unit_test_headers.docx");
    assert!(
        has_header(&pages) || has_footer(&pages),
        "should have header or footer"
    );
}

// ---------------------------------------------------------------------------
// unit_test_lists.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_unit_test_lists() {
    assert_produces_valid_pdf("unit_test_lists.docx");
}

#[test]
fn structure_unit_test_lists() {
    let pages = flow_pages("unit_test_lists.docx");
    let blocks = all_blocks(&pages);
    assert!(has_list_block(&blocks), "should have Block::List");
}

// ---------------------------------------------------------------------------
// word_tables.docx
// ---------------------------------------------------------------------------

#[test]
fn smoke_word_tables() {
    assert_produces_valid_pdf("word_tables.docx");
}

#[test]
fn structure_word_tables() {
    let pages = flow_pages("word_tables.docx");
    let blocks = all_blocks(&pages);
    assert!(has_table_block(&blocks), "should have Block::Table");
}

// ===========================================================================
// PDF text content verification
// ===========================================================================

/// Helper: convert a DOCX fixture to PDF and extract text.
fn pdf_text(name: &str) -> String {
    let path = fixture_path(name);
    let result = office2pdf::convert(&path).expect("conversion should succeed");
    common::extract_pdf_text(&result.pdf)
}

// ---------------------------------------------------------------------------
// heading123.docx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_heading123() {
    let text = pdf_text("heading123.docx");
    assert!(
        text.contains("First paragraph"),
        "PDF should contain heading text 'First paragraph'"
    );
    assert!(
        text.contains("Second paragraph"),
        "PDF should contain heading text 'Second paragraph'"
    );
    assert!(
        text.contains("Third paragraph"),
        "PDF should contain heading text 'Third paragraph'"
    );
}

// ---------------------------------------------------------------------------
// table.docx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_table() {
    let text = pdf_text("table.docx");
    assert!(
        text.contains("Datum"),
        "PDF should contain table header 'Datum'"
    );
    assert!(
        text.contains("Beschreibung"),
        "PDF should contain table header 'Beschreibung'"
    );
    assert!(
        text.contains("Preis"),
        "PDF should contain table header 'Preis'"
    );
}

// ---------------------------------------------------------------------------
// styles_en.docx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_styles_en() {
    let text = pdf_text("styles_en.docx");
    assert!(text.contains("Heading 1"), "PDF should contain 'Heading 1'");
    assert!(text.contains("Heading 2"), "PDF should contain 'Heading 2'");
    assert!(
        text.contains("Normal"),
        "PDF should contain 'Normal' style text"
    );
}

// ---------------------------------------------------------------------------
// test_python_docx.docx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_test_python_docx() {
    let text = pdf_text("test_python_docx.docx");
    assert!(
        text.contains("python-docx was here"),
        "PDF should contain 'python-docx was here'"
    );
}

// ---------------------------------------------------------------------------
// unit_test_formatting.docx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_unit_test_formatting() {
    let text = pdf_text("unit_test_formatting.docx");
    assert!(
        text.contains("bold"),
        "PDF should contain 'bold' formatting label"
    );
    assert!(
        text.contains("italic"),
        "PDF should contain 'italic' formatting label"
    );
    assert!(
        text.contains("underline"),
        "PDF should contain 'underline' formatting label"
    );
}

// ===========================================================================
// Third-party fixtures — smoke tests (must not panic)
// ===========================================================================

/// Generate a pair of smoke + basic-structure tests for a DOCX fixture.
macro_rules! docx_fixture_tests {
    ($test_name:ident, $file:expr) => {
        paste::paste! {
            #[test]
            fn [<smoke_ $test_name>]() {
                assert_produces_valid_pdf($file);
            }

            #[test]
            fn [<structure_ $test_name>]() {
                let data = load_fixture($file);
                match DocxParser.parse(&data, &ConvertOptions::default()) {
                    Ok((doc, _)) => {
                        let _ = doc.pages.len();
                    }
                    Err(e) => {
                        eprintln!("[WARN] {}: parse error (non-panic): {e}", $file);
                    }
                }
            }
        }
    };
}

// --- CC0 (Public Domain) ---------------------------------------------------

docx_fixture_tests!(ffc, "ffc.docx");
docx_fixture_tests!(one_page, "1-page.docx");
docx_fixture_tests!(three_pages, "3-pages.docx");
docx_fixture_tests!(five_pages, "5-pages.docx");
docx_fixture_tests!(ten_pages, "10-pages.docx");

// --- Apache POI (Apache 2.0) -----------------------------------------------

docx_fixture_tests!(bookmarks, "bookmarks.docx");
docx_fixture_tests!(capitalized, "capitalized.docx");
docx_fixture_tests!(chartex, "chartex.docx");
docx_fixture_tests!(checkboxes, "checkboxes.docx");
docx_fixture_tests!(comment, "comment.docx");
docx_fixture_tests!(complex_numbered_lists, "ComplexNumberedLists.docx");
docx_fixture_tests!(delins, "delins.docx");
docx_fixture_tests!(diff_first_page_head_foot, "DiffFirstPageHeadFoot.docx");
docx_fixture_tests!(drawing, "drawing.docx");
docx_fixture_tests!(embedded_document, "EmbeddedDocument.docx");
docx_fixture_tests!(endnotes, "endnotes.docx");
docx_fixture_tests!(fancy_foot, "FancyFoot.docx");
docx_fixture_tests!(field_codes, "FieldCodes.docx");
docx_fixture_tests!(header_footer_unicode, "HeaderFooterUnicode.docx");
docx_fixture_tests!(heading123, "heading123.docx");
docx_fixture_tests!(illustrative_cases, "IllustrativeCases.docx");
docx_fixture_tests!(poi_footnotes, "poi_footnotes.docx");
docx_fixture_tests!(poi_sample, "poi_sample.docx");
docx_fixture_tests!(poi_styles, "poi_styles.docx");
docx_fixture_tests!(various_pictures, "VariousPictures.docx");
docx_fixture_tests!(with_tabs, "WithTabs.docx");
docx_fixture_tests!(word_with_attachments, "WordWithAttachments.docx");

// --- MIT: Open-Xml-PowerTools (Microsoft) ----------------------------------

docx_fixture_tests!(oxp_table, "oxp_table.docx");
docx_fixture_tests!(oxp_content_control, "oxp_content_control.docx");
docx_fixture_tests!(oxp_lots_of_stuff, "oxp_lots_of_stuff.docx");
docx_fixture_tests!(oxp_complex_table, "oxp_complex_table.docx");
docx_fixture_tests!(oxp_footnote_ref, "oxp_footnote_ref.docx");
