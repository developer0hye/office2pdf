#![cfg(not(target_arch = "wasm32"))] // native-only integration tests (fs, qpdf, criterion)
//! Integration tests for DOCX fixtures.
//!
//! Each real-world `.docx` file in `tests/fixtures/docx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::ir::{
    ArrowHead, Block, BorderLineStyle, Color, FlowPage, FrameAnchor, HFInline, Insets, ListKind,
    Page, Paragraph, PositionedTabAlignment, PositionedTabRelativeTo, Run, ShapeKind,
    TextBoxVerticalAlign,
};
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

/// Recursively collect all runs from blocks (paragraphs, lists, tables, floating text boxes).
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
        Block::FloatingTextBox(text_box) => {
            for block in &text_box.content {
                collect_runs_from_block(block, out);
            }
        }
        Block::Image(_)
        | Block::InlineImages(_)
        | Block::FloatingImage(_)
        | Block::FloatingShape(_)
        | Block::MathEquation(_)
        | Block::Chart(_)
        | Block::PageBreak
        | Block::ColumnBreak => {}
    }
}

fn paragraph_text(paragraph: &Paragraph) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

fn block_text(block: &Block) -> String {
    match block {
        Block::Paragraph(paragraph) => paragraph_text(paragraph),
        Block::List(list) => list
            .items
            .iter()
            .flat_map(|item| item.content.iter())
            .map(paragraph_text)
            .collect::<Vec<String>>()
            .join("\n"),
        Block::Table(table) => table
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .flat_map(|cell| cell.content.iter())
            .map(block_text)
            .collect::<Vec<String>>()
            .join("\n"),
        Block::FloatingTextBox(text_box) => text_box
            .content
            .iter()
            .map(block_text)
            .collect::<Vec<String>>()
            .join("\n"),
        Block::Image(_)
        | Block::InlineImages(_)
        | Block::FloatingImage(_)
        | Block::FloatingShape(_)
        | Block::MathEquation(_)
        | Block::Chart(_)
        | Block::PageBreak
        | Block::ColumnBreak => String::new(),
    }
}

// ---------------------------------------------------------------------------
// PR #187 contributor acceptance fixture
// ---------------------------------------------------------------------------

const PR_187_FIXTURE: &str = "pr_187_contributor_acceptance.docx";

#[test]
fn smoke_pr_187_contributor_acceptance_fixture() {
    assert_produces_valid_pdf(PR_187_FIXTURE);
}

fn pr_187_table(pages: &[FlowPage]) -> &office2pdf::ir::Table {
    pages
        .iter()
        .flat_map(|page| page.content.iter())
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("PR #187 fixture should contain a table")
}

#[test]
fn structure_pr_187_contributor_acceptance_supported_behavior() {
    let pages = flow_pages(PR_187_FIXTURE);
    let table = pr_187_table(&pages);

    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 3);
    assert_eq!(
        table.default_cell_padding,
        Some(Insets {
            top: 5.0,
            right: 20.0,
            bottom: 15.0,
            left: 10.0,
        })
    );
    assert_eq!(
        table.rows[0].cells[0].padding,
        Some(Insets {
            top: 10.0,
            right: 20.0,
            bottom: 15.0,
            left: 15.0,
        })
    );
    assert_eq!(
        table.rows[0].cells[2].background,
        Some(Color::new(0x22, 0x22, 0x22)),
        "a cell without explicit shading should inherit the dark table style"
    );
}

#[test]
#[ignore = "pending PR #187 adaptation: retain explicit white cell shading"]
fn acceptance_pr_187_contributor_acceptance_explicit_white_shading() {
    let pages = flow_pages(PR_187_FIXTURE);
    let table = pr_187_table(&pages);

    assert_eq!(
        table.rows[0].cells[1].background,
        Some(Color::new(0xFF, 0xFF, 0xFF))
    );
}

#[test]
#[ignore = "pending PR #187 adaptation: resolve paragraph numbering from its style"]
fn acceptance_pr_187_contributor_acceptance_style_inherited_numbering() {
    let pages = flow_pages(PR_187_FIXTURE);
    let lists = pages
        .iter()
        .flat_map(|page| page.content.iter())
        .filter_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .collect::<Vec<_>>();

    assert_eq!(lists.len(), 2);
    assert_eq!(lists[0].kind, ListKind::Unordered);
    assert_eq!(lists[1].kind, ListKind::Ordered);
    let item_text = lists
        .iter()
        .flat_map(|list| list.items.iter())
        .flat_map(|item| item.content.iter())
        .map(paragraph_text)
        .collect::<Vec<_>>();
    assert_eq!(
        item_text,
        vec!["Bullet alpha", "Bullet beta", "Number one", "Number two"]
    );
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

fn image_block_count(blocks: &[&Block]) -> usize {
    blocks
        .iter()
        .map(|block| match block {
            Block::Image(_) | Block::FloatingImage(_) => 1,
            Block::InlineImages(images) => images.len(),
            _ => 0,
        })
        .sum()
}

fn image_flow_container_count(blocks: &[&Block]) -> usize {
    blocks
        .iter()
        .filter(|block| {
            matches!(
                block,
                Block::Image(_) | Block::InlineImages(_) | Block::FloatingImage(_)
            )
        })
        .count()
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

fn footer_elements(pages: &[FlowPage]) -> Vec<&HFInline> {
    pages
        .iter()
        .filter_map(|page| page.footer.as_ref())
        .flat_map(|footer| footer.paragraphs.iter())
        .flat_map(|paragraph| paragraph.elements.iter())
        .collect()
}

fn footer_text(pages: &[FlowPage]) -> String {
    footer_elements(pages)
        .iter()
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect()
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

// ---------------------------------------------------------------------------
// issue_176_office2pdf_test.docx
// ---------------------------------------------------------------------------

const ISSUE_176_FIXTURE: &str = "issue_176_office2pdf_test.docx";

#[test]
fn smoke_issue_176_office2pdf_test() {
    assert_produces_valid_pdf(ISSUE_176_FIXTURE);
}

#[test]
fn structure_issue_176_office2pdf_test() {
    let pages = flow_pages(ISSUE_176_FIXTURE);
    let blocks = all_blocks(&pages);
    let runs = all_runs(&blocks);

    let floating_shape_count = blocks
        .iter()
        .filter(|block| matches!(block, Block::FloatingShape(_)))
        .count();
    assert_eq!(
        floating_shape_count, 3,
        "issue #176 should preserve two rectangles and one arrow shape"
    );

    let rectangle_count = blocks
        .iter()
        .filter(|block| {
            matches!(
                block,
                Block::FloatingShape(shape)
                    if matches!(shape.shape.kind, ShapeKind::Rectangle)
                        && shape.shape.fill.is_some()
                        && shape.shape.stroke.is_some()
            )
        })
        .count();
    assert_eq!(
        rectangle_count, 2,
        "blue filled rectangle shapes should survive parsing"
    );

    assert!(
        blocks.iter().any(|block| {
            matches!(
                block,
                Block::FloatingShape(shape)
                    if matches!(
                        shape.shape.kind,
                        ShapeKind::Line {
                            tail_end: ArrowHead::Triangle,
                            ..
                        }
                    )
            )
        }),
        "the connector arrow should survive parsing with its arrowhead"
    );

    let floating_text_box_texts: Vec<String> = blocks
        .iter()
        .filter_map(|block| match block {
            Block::FloatingTextBox(text_box) => Some(
                text_box
                    .content
                    .iter()
                    .map(block_text)
                    .collect::<Vec<String>>()
                    .join("\n"),
            ),
            _ => None,
        })
        .collect();
    assert_eq!(
        floating_text_box_texts.len(),
        2,
        "issue #176 should preserve both floating text boxes"
    );
    assert!(
        floating_text_box_texts
            .iter()
            .any(|text| text.contains("Very important drawing")),
        "left text box content should be preserved"
    );
    assert!(
        floating_text_box_texts
            .iter()
            .any(|text| text.contains("Very important text inside a box")),
        "right text box content should be preserved"
    );

    let list = blocks
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("issue #176 should contain one logical list");
    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(
        list.items
            .iter()
            .map(|item| item.level)
            .collect::<Vec<u32>>(),
        vec![0, 0, 1, 1],
        "ordered items should continue while sub-items stay nested"
    );
    assert_eq!(
        list.items[1].start_at, None,
        "the second ordered item should continue numbering instead of restarting"
    );

    let table = blocks
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("issue #176 should contain the final data table");
    assert_eq!(table.rows.len(), 4);
    assert!(
        table.rows.iter().all(|row| row.cells.len() == 2),
        "the data table should remain two columns wide"
    );
    assert_eq!(table.header_row_count, 1);

    let document_text = runs.iter().map(|run| run.text.as_str()).collect::<String>();
    assert!(
        document_text.contains("$TERM\nprintf"),
        "hard line breaks in the code block should be preserved"
    );
    assert!(
        runs.iter()
            .any(|run| run.text == "echo" && run.style.color.is_some()),
        "syntax-highlight character styles should apply to code tokens"
    );
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

#[test]
fn structure_complex_numbered_lists_preserves_restarts_and_continuations() {
    let pages = flow_pages("ComplexNumberedLists.docx");
    let blocks = all_blocks(&pages);
    let lists = blocks
        .iter()
        .filter_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .collect::<Vec<_>>();

    assert_eq!(lists.len(), 2, "normal text should split the list blocks");
    let items = lists
        .iter()
        .flat_map(|list| list.items.iter())
        .map(|item| {
            let text = item
                .content
                .iter()
                .flat_map(|paragraph| paragraph.runs.iter())
                .map(|run| run.text.as_str())
                .collect::<String>();
            (text, item.level, item.start_at)
        })
        .collect::<Vec<_>>();

    assert_eq!(
        items,
        vec![
            ("Entry #1".to_string(), 0, Some(1)),
            ("Entry #2, with children".to_string(), 0, None),
            ("2-a".to_string(), 1, Some(1)),
            ("2-b".to_string(), 1, None),
            ("2-c".to_string(), 1, None),
            ("Entry #3".to_string(), 0, None),
            ("Entry #4".to_string(), 0, None),
            ("Restarted to 1 from 5".to_string(), 0, Some(1)),
            ("Restarted @ 2".to_string(), 0, None),
            ("Restarted @ 3".to_string(), 0, None),
            ("Jump to new list at 10".to_string(), 0, Some(10)),
            ("Now 11".to_string(), 0, None),
            ("Carrying on @ 12".to_string(), 0, Some(12)),
            ("Carrying on @ 13".to_string(), 0, None),
        ]
    );
}
docx_fixture_tests!(delins, "delins.docx");
docx_fixture_tests!(diff_first_page_head_foot, "DiffFirstPageHeadFoot.docx");

#[test]
fn structure_diff_first_page_head_foot_preserves_default_footer_columns() {
    let text = footer_text(&flow_pages("DiffFirstPageHeadFoot.docx"));

    assert!(text.contains("Footer Left"));
    assert!(text.contains("Footer Middle"));
    assert!(text.contains("Footer Right"));
}
docx_fixture_tests!(drawing, "drawing.docx");
docx_fixture_tests!(embedded_document, "EmbeddedDocument.docx");
docx_fixture_tests!(endnotes, "endnotes.docx");
docx_fixture_tests!(fancy_foot, "FancyFoot.docx");

#[test]
fn structure_fancy_foot_preserves_text_and_simple_page_field() {
    let pages = flow_pages("FancyFoot.docx");
    let footer = pages[0].footer.as_ref().expect("default footer");
    let first_paragraph = &footer.paragraphs[0];
    let elements = footer_elements(&pages);
    let text = footer_text(&pages);

    assert!(text.contains("This is a fancy alphabet footer, with page number and everything"));
    assert!(text.contains("Page "));
    assert!(
        !elements
            .iter()
            .any(|element| matches!(element, HFInline::Run(run) if run.text.trim() == "2")),
        "the cached fldSimple result must not follow the dynamic page number"
    );
    assert!(
        elements
            .iter()
            .any(|element| matches!(element, HFInline::PageNumber)),
        "the real-world fldSimple PAGE field should survive parsing"
    );
    let top_border = first_paragraph
        .border
        .as_ref()
        .and_then(|border| border.top.as_ref())
        .expect("the compound top border should survive parsing");
    assert_eq!(top_border.style, BorderLineStyle::Double);
    assert!((top_border.width - 3.0).abs() < f64::EPSILON);
    assert_eq!(top_border.color, Color::new(0x62, 0x24, 0x23));
    assert!(first_paragraph.elements.iter().any(|element| matches!(
        element,
        HFInline::PositionedTab(tab)
            if tab.alignment == PositionedTabAlignment::Right
                && tab.relative_to == PositionedTabRelativeTo::Margin
    )));
    assert!((footer.distance_from_edge.expect("footer distance") - 35.4).abs() < 0.01);
}
docx_fixture_tests!(field_codes, "FieldCodes.docx");
docx_fixture_tests!(header_footer_unicode, "HeaderFooterUnicode.docx");

#[test]
fn structure_header_footer_unicode_preserves_accented_text() {
    let text = footer_text(&flow_pages("HeaderFooterUnicode.docx"));

    assert!(text.contains("The footer, with Molière, has Unicode in it."));
}
docx_fixture_tests!(heading123, "heading123.docx");
docx_fixture_tests!(illustrative_cases, "IllustrativeCases.docx");
docx_fixture_tests!(poi_footnotes, "poi_footnotes.docx");
docx_fixture_tests!(poi_sample, "poi_sample.docx");
docx_fixture_tests!(poi_styles, "poi_styles.docx");
docx_fixture_tests!(various_pictures, "VariousPictures.docx");

#[test]
fn structure_various_pictures_preserves_raster_emf_and_wmf_images() {
    let pages = flow_pages("VariousPictures.docx");
    let blocks = all_blocks(&pages);

    assert_eq!(
        image_block_count(&blocks),
        5,
        "PNG, JPEG, both EMF images, and the WMF image should survive parsing"
    );
}

#[test]
fn structure_various_pictures_keeps_inline_images_in_one_flow_container() {
    let pages = flow_pages("VariousPictures.docx");
    let blocks = all_blocks(&pages);

    assert_eq!(
        image_flow_container_count(&blocks),
        1,
        "images from one Word paragraph should wrap within one flow container"
    );
}
docx_fixture_tests!(with_tabs, "WithTabs.docx");
docx_fixture_tests!(word_with_attachments, "WordWithAttachments.docx");

// --- MIT: Open-Xml-PowerTools (Microsoft) ----------------------------------

docx_fixture_tests!(oxp_table, "oxp_table.docx");
docx_fixture_tests!(oxp_content_control, "oxp_content_control.docx");
docx_fixture_tests!(oxp_lots_of_stuff, "oxp_lots_of_stuff.docx");
docx_fixture_tests!(oxp_complex_table, "oxp_complex_table.docx");

#[test]
fn structure_oxp_complex_table_preserves_conditional_shading_and_merges() {
    let pages = flow_pages("oxp_complex_table.docx");
    let blocks = all_blocks(&pages);
    let table = blocks
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("fixture should contain a table");

    assert_eq!(table.rows.len(), 8);
    assert_eq!(table.rows[0].cells.len(), 9);
    assert_eq!(table.rows[0].cells[1].col_span, 2);
    assert_eq!(table.rows[1].cells[0].row_span, 3);
    assert_eq!(table.rows[2].cells.len(), 9);
    assert_eq!(table.rows[3].cells.len(), 9);

    let cell_text = |row: usize, cell: usize| {
        table.rows[row].cells[cell]
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph_text(paragraph)),
                _ => None,
            })
            .collect::<String>()
    };
    assert_eq!(cell_text(0, 1), "Hort merged");
    assert_eq!(cell_text(1, 0), "Vert merged");
    assert_eq!(cell_text(3, 0), "bb");
    assert_eq!(cell_text(7, 1), "A44");

    let black = Some(Color::new(0x00, 0x00, 0x00));
    let dark_gray = Some(Color::new(0x99, 0x99, 0x99));
    let light_gray = Some(Color::new(0xCC, 0xCC, 0xCC));
    assert!(
        table.rows[0]
            .cells
            .iter()
            .all(|cell| cell.background == black),
        "first-row conditional style should use a black fill"
    );
    let first_header_run = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        _ => panic!("first header cell should contain a paragraph"),
    };
    assert_eq!(
        first_header_run.style.color,
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(first_header_run.style.bold, Some(true));
    assert_eq!(table.rows[1].cells[0].background, black);
    assert_eq!(table.rows[1].cells[1].background, dark_gray);
    assert!(
        table.rows[2]
            .cells
            .iter()
            .all(|cell| cell.background == light_gray)
    );
    assert!(
        table.rows[3]
            .cells
            .iter()
            .all(|cell| cell.background == dark_gray)
    );
    for (row_index, expected_body_fill) in [
        (4, light_gray),
        (5, dark_gray),
        (6, light_gray),
        (7, dark_gray),
    ] {
        assert_eq!(table.rows[row_index].cells[0].background, black);
        assert!(
            table.rows[row_index].cells[1..]
                .iter()
                .all(|cell| cell.background == expected_body_fill)
        );
    }
}
docx_fixture_tests!(oxp_footnote_ref, "oxp_footnote_ref.docx");

// --- Encrypted / password-protected fixtures --------------------------------
// These files are OLE2 containers (not ZIP); conversion must return
// ConvertError::UnsupportedEncryption instead of a misleading parse error.

/// Returns `true` if the file is a Git LFS pointer (not the actual content).
fn is_lfs_pointer(path: &std::path::Path) -> bool {
    std::fs::read(path)
        .map(|data| data.starts_with(b"version https://git-lfs"))
        .unwrap_or(false)
}

macro_rules! encrypted_docx_tests {
    ($name:ident, $fixture:expr) => {
        mod $name {
            use super::*;

            #[test]
            fn returns_unsupported_encryption() {
                let path = fixture_path($fixture);
                if is_lfs_pointer(&path) {
                    eprintln!("Skipping {}: Git LFS pointer (not fetched)", $fixture);
                    return;
                }
                let err = office2pdf::convert(&path).unwrap_err();
                assert!(
                    matches!(err, office2pdf::error::ConvertError::UnsupportedEncryption),
                    "Expected UnsupportedEncryption for {}, got: {err:?}",
                    $fixture
                );
            }
        }
    };
}

encrypted_docx_tests!(
    encrypted_lo_standard,
    "libreoffice/Encrypted_LO_Standard_abc.docx"
);
encrypted_docx_tests!(encrypted_mso2007, "libreoffice/Encrypted_MSO2007_abc.docx");
encrypted_docx_tests!(encrypted_mso2010, "libreoffice/Encrypted_MSO2010_abc.docx");
encrypted_docx_tests!(encrypted_mso2013, "libreoffice/Encrypted_MSO2013_abc.docx");
encrypted_docx_tests!(password_is_pass, "poi/bug53475-password-is-pass.docx");
encrypted_docx_tests!(
    password_is_solrcell,
    "poi/bug53475-password-is-solrcell.docx"
);

// --- LibreOffice DOCX fixtures (previously failing due to docx-rs limitations) ---
// Fixed by patched docx-rs fork (developer0hye/docx-rs, branch fix/parse-tolerance).
// See: https://github.com/developer0hye/office2pdf/issues/84

// Previously panicked — Strict OOXML dxa unit suffix in width values
docx_fixture_tests!(tdf79272_strict_dxa, "libreoffice/tdf79272_strictDxa.docx");

// Previously "Failed to read from zip" — minimal DOCX without document rels
docx_fixture_tests!(tdf108350, "libreoffice/tdf108350.docx");
docx_fixture_tests!(tdf108408, "libreoffice/tdf108408.docx");
docx_fixture_tests!(tdf108714, "libreoffice/tdf108714.docx");
docx_fixture_tests!(tdf108806, "libreoffice/tdf108806.docx");
docx_fixture_tests!(tdf108849, "libreoffice/tdf108849.docx");
docx_fixture_tests!(tdf109306, "libreoffice/tdf109306.docx");
docx_fixture_tests!(tdf109524, "libreoffice/tdf109524.docx");
docx_fixture_tests!(tdf111550, "libreoffice/tdf111550.docx");
docx_fixture_tests!(tdf111964, "libreoffice/tdf111964.docx");
docx_fixture_tests!(tdf124670, "libreoffice/tdf124670.docx");
docx_fixture_tests!(tdf129659, "libreoffice/tdf129659.docx");
docx_fixture_tests!(
    tdf159207_footer_frame_border,
    "libreoffice/tdf159207_footerFramePrBorder.docx"
);

#[test]
fn structure_tdf159207_preserves_page_anchored_footer_frame_position() {
    let pages = flow_pages("libreoffice/tdf159207_footerFramePrBorder.docx");
    let footer = pages[0].footer.as_ref().expect("default footer");
    let framed_paragraphs = footer
        .paragraphs
        .iter()
        .filter_map(|paragraph| paragraph.frame.as_ref())
        .collect::<Vec<_>>();

    assert_eq!(framed_paragraphs.len(), 6);
    for frame in framed_paragraphs {
        assert_eq!(frame.horizontal_anchor, FrameAnchor::Page);
        assert_eq!(frame.vertical_anchor, FrameAnchor::Page);
        assert!((frame.x.expect("absolute x") - 71.8).abs() < 0.01);
        assert!((frame.y.expect("absolute y") - 198.5).abs() < 0.01);
    }
}

#[test]
fn conversion_tdf159207_renders_positioned_footer_frame() {
    let result = office2pdf::convert(fixture_path(
        "libreoffice/tdf159207_footerFramePrBorder.docx",
    ))
    .expect("the positioned footer frame should render");

    assert!(result.pdf.starts_with(b"%PDF"));
    common::validate_pdf_with_qpdf(&result.pdf);
}
docx_fixture_tests!(table_rtl, "libreoffice/table-rtl.docx");

docx_fixture_tests!(wpg_only, "libreoffice/wpg-only.docx");

#[test]
fn structure_wpg_only_preserves_grouped_shapes() {
    let pages = flow_pages("libreoffice/wpg-only.docx");
    let blocks = all_blocks(&pages);
    let shapes = blocks
        .iter()
        .filter_map(|block| match block {
            Block::FloatingShape(shape) => Some(shape),
            _ => None,
        })
        .collect::<Vec<_>>();

    assert_eq!(shapes.len(), 2, "both WPG child shapes should survive");
    assert!(matches!(shapes[0].shape.kind, ShapeKind::Ellipse));
    assert!(matches!(shapes[1].shape.kind, ShapeKind::Polygon { .. }));
    assert!((shapes[0].offset_x - 43.15).abs() < 0.01);
    assert!((shapes[0].offset_y - 27.40).abs() < 0.01);
    assert!((shapes[0].width - 42.75).abs() < 0.01);
    assert!((shapes[0].height - 39.75).abs() < 0.01);
    assert!((shapes[1].offset_x - 112.15).abs() < 0.01);
    assert!((shapes[1].offset_y - 38.65).abs() < 0.01);
    assert!((shapes[1].width - 57.0).abs() < 0.01);
    assert!((shapes[1].height - 45.0).abs() < 0.01);
}

docx_fixture_tests!(wpg_textboxes, "libreoffice/testWPGtextboxes.docx");

#[test]
fn structure_wpg_textboxes_preserves_grouped_text_content() {
    let pages = flow_pages("libreoffice/testWPGtextboxes.docx");
    let blocks = all_blocks(&pages);
    let text = blocks
        .iter()
        .map(|block| block_text(block))
        .collect::<Vec<String>>()
        .join("\n");
    let text_box_count = blocks
        .iter()
        .filter(|block| matches!(block, Block::FloatingTextBox(_)))
        .count();

    assert_eq!(text_box_count, 3, "all WPG text boxes should survive");

    let text_boxes = blocks
        .iter()
        .filter_map(|block| match block {
            Block::FloatingTextBox(text_box) => Some(text_box),
            _ => None,
        })
        .collect::<Vec<_>>();
    for text_box in &text_boxes {
        assert_eq!(text_box.vertical_align, TextBoxVerticalAlign::Center);
    }
    assert!(text_boxes[0].padding.left > 50.0);
    assert!(text_boxes[0].padding.top > 100.0);
    assert!(text_boxes[1].padding.left > 25.0);
    assert!(text_boxes[1].padding.top > 20.0);
    assert!(text_boxes[2].padding.left > 30.0);
    assert!(text_boxes[2].padding.top > 20.0);

    let text_box_blocks = text_boxes
        .iter()
        .flat_map(|text_box| text_box.content.iter())
        .collect::<Vec<_>>();
    let text_box_runs = all_runs(&text_box_blocks);
    assert!(
        text_box_runs
            .iter()
            .filter(|run| !run.text.trim().is_empty())
            .all(|run| run.style.color == Some(Color::new(255, 255, 255))),
        "WPG fontRef color should be applied to unstyled text"
    );
    assert!(
        text.contains("This is a triangle having a table inside:"),
        "missing triangle text in {text:?}"
    );
    assert!(
        text.contains("This is a circle, having a picture inside:"),
        "missing circle text in {text:?}"
    );
    assert!(
        text.contains("This is a diamond"),
        "missing diamond text in {text:?}"
    );
    for cell in ["A", "B", "C", "D"] {
        assert!(
            text.lines().any(|line| line == cell),
            "missing table cell {cell}"
        );
    }
}

docx_fixture_tests!(groupshape_picture, "libreoffice/groupshape-picture.docx");

#[test]
fn structure_groupshape_picture_preserves_one_positioned_choice_image() {
    let pages = flow_pages("libreoffice/groupshape-picture.docx");
    let blocks = all_blocks(&pages);
    let floating_images = blocks
        .iter()
        .filter_map(|block| match block {
            Block::FloatingImage(image) => Some(image),
            _ => None,
        })
        .collect::<Vec<_>>();
    let other_image_count = blocks
        .iter()
        .map(|block| match block {
            Block::Image(_) => 1,
            Block::InlineImages(images) => images.len(),
            _ => 0,
        })
        .sum::<usize>();

    assert_eq!(
        floating_images.len(),
        1,
        "selected choice image should survive"
    );
    assert_eq!(other_image_count, 0, "VML fallback must not be duplicated");
    assert!((floating_images[0].offset_x - 31.423).abs() < 0.01);
    assert!((floating_images[0].offset_y - 74.573).abs() < 0.01);
}

docx_fixture_tests!(n592908_picture, "libreoffice/n592908-picture.docx");

#[test]
fn structure_n592908_picture_preserves_one_legacy_vml_image() {
    let pages = flow_pages("libreoffice/n592908-picture.docx");
    let blocks = all_blocks(&pages);
    let image_count = blocks
        .iter()
        .map(|block| match block {
            Block::Image(_) | Block::FloatingImage(_) => 1,
            Block::InlineImages(images) => images.len(),
            _ => 0,
        })
        .sum::<usize>();

    assert_eq!(
        image_count, 1,
        "legacy w:pict image should survive exactly once"
    );
}

#[test]
fn structure_table_rtl_uses_visual_right_to_left_cell_order() {
    let pages = flow_pages("libreoffice/table-rtl.docx");
    let blocks = all_blocks(&pages);
    let table = blocks
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("the RTL fixture should contain a table");
    let row_texts: Vec<Vec<String>> = table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| cell.content.iter().map(block_text).collect::<String>())
                .collect()
        })
        .collect();

    assert_eq!(row_texts, vec![vec!["B1", "A1"], vec!["B2", "A2"]]);
}
docx_fixture_tests!(cloud, "libreoffice/cloud.docx");
docx_fixture_tests!(xml_space, "libreoffice/xml_space.docx");
docx_fixture_tests!(
    sdt_after_section_break,
    "libreoffice/sdt_after_section_break.docx"
);

// ODT files with .docx extension — clean parse error (not panic)
docx_fixture_tests!(tdf171025_page_after, "libreoffice/tdf171025_pageAfter.docx");
docx_fixture_tests!(tdf171038_page_after, "libreoffice/tdf171038_pageAfter.docx");

// Intentionally malformed XML — clean parse error (not panic)
docx_fixture_tests!(math_malformed_xml, "libreoffice/math-malformed_xml.docx");

// XML external entity references — docx-rs correctly rejects for security
docx_fixture_tests!(external_entity_in_text, "poi/ExternalEntityInText.docx");

// Deeply nested tables (5000+ levels) — clean error after depth-limit fix (not stack overflow)
docx_fixture_tests!(deep_table_cell, "poi/deep-table-cell.docx");
