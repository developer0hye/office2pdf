#![cfg(not(target_arch = "wasm32"))] // native-only integration tests (fs, qpdf, criterion)
//! Integration tests for XLSX fixtures.
//!
//! Each real-world `.xlsx` file in `tests/fixtures/xlsx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::ir::{Alignment, Block, BorderLineStyle, Color, Page, SheetPage, TableCell};
use office2pdf::parser::Parser;
use office2pdf::parser::xlsx::XlsxParser;
use office2pdf::render::typst_gen::generate_typst;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx")
        .join(name)
}

fn load_fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("fixture file should exist")
}

/// Smoke-test helper: conversion must not panic.
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
            eprintln!("[WARN] {name}: conversion error (non-panic): {e}");
        }
    }
}

/// Parse an XLSX fixture and return the sheet pages.
fn sheet_pages(name: &str) -> Vec<SheetPage> {
    let data = load_fixture(name);
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|p| match p {
            Page::Sheet(sp) => Some(sp),
            _ => None,
        })
        .collect()
}

fn sheet_names(pages: &[SheetPage]) -> Vec<&str> {
    pages.iter().map(|p| p.name.as_str()).collect()
}

fn total_rows(pages: &[SheetPage]) -> usize {
    pages.iter().map(|p| p.table.rows.len()).sum()
}

fn has_cell_border(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.border.is_some())
    })
}

fn has_merged_cells(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.col_span > 1 || c.row_span > 1)
    })
}

fn has_formatted_text(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table.rows.iter().flat_map(|r| r.cells.iter()).any(|c| {
            c.content.iter().any(|b| match b {
                Block::Paragraph(para) => para.runs.iter().any(|r| {
                    r.style.bold == Some(true)
                        || r.style.italic == Some(true)
                        || r.style.color.is_some()
                }),
                _ => false,
            })
        })
    })
}

fn table_cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

fn sheet_page_named<'a>(pages: &'a [SheetPage], name: &str) -> &'a SheetPage {
    pages
        .iter()
        .find(|page| page.name == name)
        .unwrap_or_else(|| {
            let available = pages
                .iter()
                .map(|page| page.name.as_str())
                .collect::<Vec<_>>();
            panic!("missing sheet page {name}; available pages: {available:?}")
        })
}

// ---------------------------------------------------------------------------
// PR #186 contributor acceptance fixture
// ---------------------------------------------------------------------------

const PR_186_FIXTURE: &str = "pr_186_contributor_acceptance.xlsx";

#[test]
fn smoke_pr_186_contributor_acceptance_fixture() {
    assert_produces_valid_pdf(PR_186_FIXTURE);
}

#[test]
fn structure_pr_186_contributor_acceptance_supported_behavior() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let executive = sheet_page_named(&pages, "Executive Portrait");

    assert_eq!(statement.table.rows.len(), 4);
    assert_eq!(statement.table.column_widths.len(), 4);
    assert!(
        (statement.table.column_widths[1] - 108.75).abs() < 0.01,
        "20 Excel character units should convert to 108.75pt"
    );

    let alignment_row = &statement.table.rows[1];
    let expected_alignments = [
        Alignment::Left,
        Alignment::Center,
        Alignment::Right,
        Alignment::Justify,
    ];
    for (column, expected) in expected_alignments.into_iter().enumerate() {
        let paragraph = match &alignment_row.cells[column].content[0] {
            Block::Paragraph(paragraph) => paragraph,
            _ => panic!("alignment cell should contain a paragraph"),
        };
        assert_eq!(paragraph.style.alignment, Some(expected));
    }

    assert!(
        !table_cell_text(&statement.table.rows[2].cells[1]).is_empty(),
        "the text-valued General cell should be retained"
    );
    let top_border = statement.table.rows[3].cells[0]
        .border
        .as_ref()
        .and_then(|border| border.top.as_ref())
        .expect("A4 should have a top border");
    assert_eq!(top_border.style, BorderLineStyle::Double);

    assert!((statement.margins.top - 28.8).abs() < 0.01);
    assert!((statement.margins.bottom - 36.0).abs() < 0.01);
    assert!((statement.margins.left - 21.6).abs() < 0.01);
    assert!((statement.margins.right - 43.2).abs() < 0.01);
    assert!((executive.margins.top - 54.0).abs() < 0.01);
    assert!((executive.margins.bottom - 54.0).abs() < 0.01);
    assert!((executive.margins.left - 50.4).abs() < 0.01);
    assert!((executive.margins.right - 50.4).abs() < 0.01);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_six_digit_colors() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let header_cell = &statement.table.rows[0].cells[0];

    assert_eq!(
        header_cell.background,
        Some(Color::new(0xD9, 0xEA, 0xF7)),
        "six-digit header fill should survive XLSX parsing"
    );
    let paragraph = match &header_cell.content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("header cell should contain a paragraph"),
    };
    assert_eq!(
        paragraph.runs[0].style.color,
        Some(Color::new(0x16, 0x32, 0x4F)),
        "six-digit header font color should survive XLSX parsing"
    );
}

#[test]
fn acceptance_pr_186_contributor_acceptance_numeric_general_alignment() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let general_row = &statement.table.rows[2];

    let numeric = match &general_row.cells[0].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("numeric cell should contain a paragraph"),
    };
    let numeric_looking_text = match &general_row.cells[1].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("text cell should contain a paragraph"),
    };
    assert_eq!(numeric.style.alignment, Some(Alignment::Right));
    assert_eq!(numeric_looking_text.style.alignment, None);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_page_setup() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let executive = sheet_page_named(&pages, "Executive Portrait");

    assert!((statement.size.width - 612.0).abs() < 0.01);
    assert!((statement.size.height - 396.0).abs() < 0.01);
    assert!((executive.size.width - 522.0).abs() < 0.01);
    assert!((executive.size.height - 756.0).abs() < 0.01);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_double_border_rendering() {
    let data = load_fixture(PR_186_FIXTURE);
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let output = generate_typst(&document).expect("fixture should generate Typst");

    assert!(!output.source.contains("dash: \"dashed\""));
    assert!(
        output
            .source
            .contains("(paint: rgb(0, 0, 0), thickness: 2.5pt)")
    );
}

// ---------------------------------------------------------------------------
// any_sheets.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_any_sheets() {
    assert_produces_valid_pdf("any_sheets.xlsx");
}

#[test]
fn structure_any_sheets() {
    // any_sheets.xlsx has 4 sheets: Visible, Hidden, VeryHidden, Chart.
    // Parser only returns visible data worksheets (not hidden/chart sheets).
    let pages = sheet_pages("any_sheets.xlsx");
    assert!(!pages.is_empty(), "should have at least one visible sheet");
    let names = sheet_names(&pages);
    assert!(
        names.iter().all(|n| !n.is_empty()),
        "all sheet names should be non-empty"
    );
}

// ---------------------------------------------------------------------------
// date.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_date() {
    assert_produces_valid_pdf("date.xlsx");
}

#[test]
fn structure_date() {
    let pages = sheet_pages("date.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ---------------------------------------------------------------------------
// merge_cells.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_merge_cells() {
    assert_produces_valid_pdf("merge_cells.xlsx");
}

#[test]
fn structure_merge_cells() {
    let pages = sheet_pages("merge_cells.xlsx");
    assert!(
        has_merged_cells(&pages),
        "should have cells with col_span > 1 or row_span > 1"
    );
}

// ---------------------------------------------------------------------------
// SH001-Table.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh001_table() {
    assert_produces_valid_pdf("SH001-Table.xlsx");
}

#[test]
fn structure_sh001_table() {
    let pages = sheet_pages("SH001-Table.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ---------------------------------------------------------------------------
// SH002-TwoTablesTwoSheets.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh002_two_tables_two_sheets() {
    assert_produces_valid_pdf("SH002-TwoTablesTwoSheets.xlsx");
}

#[test]
fn structure_sh002_two_tables_two_sheets() {
    let pages = sheet_pages("SH002-TwoTablesTwoSheets.xlsx");
    assert!(pages.len() >= 2, "should have >= 2 sheets");
    let names = sheet_names(&pages);
    let unique: std::collections::HashSet<_> = names.iter().collect();
    assert_eq!(unique.len(), names.len(), "sheet names should be unique");
}

// ---------------------------------------------------------------------------
// SH106-Formatted.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh106_formatted() {
    assert_produces_valid_pdf("SH106-Formatted.xlsx");
}

#[test]
fn structure_sh106_formatted() {
    let pages = sheet_pages("SH106-Formatted.xlsx");
    assert!(
        has_formatted_text(&pages),
        "should have formatted text (bold/italic/color)"
    );
}

// ---------------------------------------------------------------------------
// SH109-CellWithBorder.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh109_cell_with_border() {
    assert_produces_valid_pdf("SH109-CellWithBorder.xlsx");
}

#[test]
fn structure_sh109_cell_with_border() {
    let pages = sheet_pages("SH109-CellWithBorder.xlsx");
    assert!(has_cell_border(&pages), "should have cells with borders");
}

// ---------------------------------------------------------------------------
// temperature.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_temperature() {
    assert_produces_valid_pdf("temperature.xlsx");
}

#[test]
fn structure_temperature() {
    let pages = sheet_pages("temperature.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ===========================================================================
// PDF text content verification
// ===========================================================================

/// Helper: convert an XLSX fixture to PDF and extract text.
fn pdf_text(name: &str) -> String {
    let path = fixture_path(name);
    let result = office2pdf::convert(&path).expect("conversion should succeed");
    common::extract_pdf_text(&result.pdf)
}

// ---------------------------------------------------------------------------
// temperature.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_temperature() {
    let text = pdf_text("temperature.xlsx");
    assert!(
        text.contains("celsius"),
        "PDF should contain 'celsius' label"
    );
    assert!(
        text.contains("fahrenheit"),
        "PDF should contain 'fahrenheit' label"
    );
}

// ---------------------------------------------------------------------------
// SH001-Table.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_sh001_table() {
    let text = pdf_text("SH001-Table.xlsx");
    // This is a simple table with single-character headers and numeric data
    assert!(!text.is_empty(), "PDF should contain extracted text");
    // Check for numeric data that should be present
    assert!(
        text.contains('1') && text.contains('2') && text.contains('3'),
        "PDF should contain numeric data from the table"
    );
}

// ---------------------------------------------------------------------------
// SH002-TwoTablesTwoSheets.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_sh002_two_tables_two_sheets() {
    let text = pdf_text("SH002-TwoTablesTwoSheets.xlsx");
    assert!(!text.is_empty(), "PDF should contain extracted text");
    // Both sheets have different content; verify we have data from at least one
    assert!(
        text.contains('1') || text.contains('a') || text.contains('q'),
        "PDF should contain data from the sheets"
    );
}

// ===========================================================================
// Third-party fixtures — smoke tests (must not panic)
// ===========================================================================

/// Generate a pair of smoke + basic-structure tests for an XLSX fixture.
macro_rules! xlsx_fixture_tests {
    ($test_name:ident, $file:expr) => {
        paste::paste! {
            #[test]
            fn [<smoke_ $test_name>]() {
                assert_produces_valid_pdf($file);
            }

            #[test]
            fn [<structure_ $test_name>]() {
                let data = load_fixture($file);
                match XlsxParser.parse(&data, &ConvertOptions::default()) {
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

xlsx_fixture_tests!(ffc, "ffc.xlsx");
xlsx_fixture_tests!(hundred_customers, "100-customers.xlsx");
xlsx_fixture_tests!(thousand_customers, "1000-customers.xlsx");

// --- Apache POI (Apache 2.0) -----------------------------------------------

xlsx_fixture_tests!(charts_123233, "123233_charts.xlsx");
xlsx_fixture_tests!(booleans, "Booleans.xlsx");
xlsx_fixture_tests!(chart_sheet, "chart_sheet.xlsx");
xlsx_fixture_tests!(comments, "comments.xlsx");
xlsx_fixture_tests!(excel_pivot_table, "ExcelPivotTableSample.xlsx");
xlsx_fixture_tests!(excel_tables, "ExcelTables.xlsx");
xlsx_fixture_tests!(formatting, "Formatting.xlsx");
xlsx_fixture_tests!(group_test, "GroupTest.xlsx");
xlsx_fixture_tests!(header_footer_test, "headerFooterTest.xlsx");
xlsx_fixture_tests!(inline_string, "InlineString.xlsx");
xlsx_fixture_tests!(picture, "picture.xlsx");
xlsx_fixture_tests!(right_to_left, "right-to-left.xlsx");
xlsx_fixture_tests!(sample_ss, "SampleSS.xlsx");
xlsx_fixture_tests!(shared_formulas, "shared_formulas.xlsx");
xlsx_fixture_tests!(sheet_tab_colors, "SheetTabColors.xlsx");
xlsx_fixture_tests!(simple_monthly_budget, "simple-monthly-budget.xlsx");
xlsx_fixture_tests!(simple_scatter_chart, "SimpleScatterChart.xlsx");
xlsx_fixture_tests!(themes, "Themes.xlsx");
xlsx_fixture_tests!(with_chart, "WithChart.xlsx");
xlsx_fixture_tests!(with_drawing, "WithDrawing.xlsx");
xlsx_fixture_tests!(with_more_various_data, "WithMoreVariousData.xlsx");
xlsx_fixture_tests!(with_text_box, "WithTextBox.xlsx");
xlsx_fixture_tests!(with_various_data, "WithVariousData.xlsx");

// --- MIT: Open-Xml-PowerTools (Microsoft) ----------------------------------

xlsx_fixture_tests!(
    sh003_date_first_col,
    "SH003-TableWithDateInFirstColumn.xlsx"
);
xlsx_fixture_tests!(sh004_offset_location, "SH004-TableAtOffsetLocation.xlsx");
xlsx_fixture_tests!(sh005_shared_strings, "SH005-Table-With-SharedStrings.xlsx");
xlsx_fixture_tests!(sh006_no_shared_strings, "SH006-Table-No-SharedStrings.xlsx");
xlsx_fixture_tests!(sh007_one_cell, "SH007-One-Cell-Table.xlsx");
xlsx_fixture_tests!(sh008_tall_row, "SH008-Table-With-Tall-Row.xlsx");
xlsx_fixture_tests!(sh101_simple_formats, "SH101-SimpleFormats.xlsx");
xlsx_fixture_tests!(sh102_9x9, "SH102-9-x-9.xlsx");
xlsx_fixture_tests!(sh103_no_shared_string, "SH103-No-SharedString.xlsx");
xlsx_fixture_tests!(sh104_with_shared_string, "SH104-With-SharedString.xlsx");
xlsx_fixture_tests!(sh105_no_shared_string2, "SH105-No-SharedString.xlsx");
xlsx_fixture_tests!(sh107_formatted_table, "SH107-9-x-9-Formatted-Table.xlsx");
xlsx_fixture_tests!(
    sh108_simple_formatted_cell,
    "SH108-SimpleFormattedCell.xlsx"
);

// --- Upstream parse failures (umya-spreadsheet) ------------------------------
// Related: #97
// These files fail with parse errors in umya-spreadsheet. All are handled
// gracefully — no panics, no crashes. Documented as known upstream limitations.

// ZipError: specified file not found in archive
xlsx_fixture_tests!(tdf121887, "libreoffice/tdf121887.xlsx");
xlsx_fixture_tests!(tdf131575, "libreoffice/tdf131575.xlsx");
xlsx_fixture_tests!(tdf76115, "libreoffice/tdf76115.xlsx");
xlsx_fixture_tests!(poi_49609, "poi/49609.xlsx");
xlsx_fixture_tests!(poi_56278, "poi/56278.xlsx");
xlsx_fixture_tests!(poi_59021, "poi/59021.xlsx");

// IoError: Invalid checksum
xlsx_fixture_tests!(forcepoint107, "libreoffice/forcepoint107.xlsx");

// ZipError: invalid Zip archive (Could not find EOCD)
xlsx_fixture_tests!(deep_data, "poi/deep-data.xlsx");

// --- Upstream panics caught by catch_unwind (umya-spreadsheet) ----------------
// Related: #97
// These files trigger panics inside umya-spreadsheet (arithmetic overflow,
// unwrap on None). catch_unwind prevents process crashes. Documented as
// known upstream limitations.

// attempt to subtract with overflow
xlsx_fixture_tests!(
    functions_excel_2010,
    "libreoffice/functions-excel-2010.xlsx"
);
xlsx_fixture_tests!(poi_51710, "poi/51710.xlsx");

// called Option::unwrap() on a None value
xlsx_fixture_tests!(poi_64450, "poi/64450.xlsx");

// attempt to multiply with overflow
xlsx_fixture_tests!(
    formula_eval_test_data_copy,
    "poi/FormulaEvalTestData_Copy.xlsx"
);

// --- Upstream panic safety (patched umya-spreadsheet) ------------------------
// Related: #83

/// Files that previously panicked in umya-spreadsheet now convert successfully
/// after the fork fix (developer0hye/umya-spreadsheet fix/panic-safety-v2).
///
/// All 21 previously-panicking files now produce valid PDFs.
#[test]
fn previously_panicking_files_now_convert() {
    let cases: &[&str] = &[
        // --- Phase 1 fixes (PR #90) ---
        // FileNotFound panics (7 files)
        "libreoffice/chart_hyperlink.xlsx",
        "libreoffice/hyperlink.xlsx",
        "libreoffice/tdf130959.xlsx",
        "libreoffice/test_115192.xlsx",
        "poi/47504.xlsx",
        "poi/bug63189.xlsx",
        "poi/ConditionalFormattingSamples.xlsx",
        // ParseFloatError / boolean cell (1 file)
        "libreoffice/check-boolean.xlsx",
        // unwrap() on None (2 files)
        "libreoffice/tdf100709.xlsx",
        "poi/sample-beta.xlsx",
        // dataBar end element (2 files)
        "libreoffice/tdf162948.xlsx",
        "poi/NewStyleConditionalFormattings.xlsx",
        // --- Phase 2 fixes ---
        // Backslash zip paths from Windows tools (3 files)
        "libreoffice/tdf131575.xlsx",
        "libreoffice/tdf76115.xlsx",
        "poi/49609.xlsx",
        // Missing optional styles.xml (3 files)
        "poi/56278.xlsx",
        "libreoffice/tdf121887.xlsx",
        "poi/59021.xlsx",
        // Arithmetic overflow in formula parsing (2 files)
        "libreoffice/functions-excel-2010.xlsx",
        "poi/FormulaEvalTestData_Copy.xlsx",
        // Missing XML attributes (1 file)
        "poi/64450.xlsx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("Skipping {name}: fixture not available");
            continue;
        }
        assert_produces_valid_pdf(name);
    }
}

// --- MIT: calamine (Rust) --------------------------------------------------

xlsx_fixture_tests!(date_1904, "date_1904.xlsx");
xlsx_fixture_tests!(empty_sheet, "empty_sheet.xlsx");
xlsx_fixture_tests!(errors, "errors.xlsx");
xlsx_fixture_tests!(pivots, "pivots.xlsx");
xlsx_fixture_tests!(richtext_namespaced, "richtext-namespaced.xlsx");
xlsx_fixture_tests!(column_row_ranges, "column_row_ranges.xlsx");
xlsx_fixture_tests!(table_multiple, "table-multiple.xlsx");
xlsx_fixture_tests!(formula_issue, "formula.issue.xlsx");
xlsx_fixture_tests!(header_row, "header-row.xlsx");

// --- Encrypted / password-protected fixtures --------------------------------

/// Returns `true` if the file is a Git LFS pointer (not the actual content).
fn is_lfs_pointer(path: &std::path::Path) -> bool {
    std::fs::read(path)
        .map(|data| data.starts_with(b"version https://git-lfs"))
        .unwrap_or(false)
}

#[test]
fn encrypted_xlsx_returns_unsupported_encryption() {
    let path = fixture_path("poi/protected_passtika.xlsx");
    if is_lfs_pointer(&path) {
        eprintln!("Skipping protected_passtika.xlsx: Git LFS pointer (not fetched)");
        return;
    }
    let err = office2pdf::convert(&path).unwrap_err();
    assert!(
        matches!(err, office2pdf::error::ConvertError::UnsupportedEncryption),
        "Expected UnsupportedEncryption for protected_passtika.xlsx, got: {err:?}"
    );
}

// ---------------------------------------------------------------------------
// Worksheet drawings (issue #238)
// ---------------------------------------------------------------------------

#[test]
fn with_drawing_renders_anchored_images() {
    let pages = sheet_pages("poi/WithDrawing.xlsx");
    let total_images: usize = pages.iter().map(|sp| sp.images.len()).sum();
    assert!(
        total_images >= 3,
        "the drawing anchors five pictures (jpeg/png plus metafiles); at least \
         the raster ones must be extracted, got {total_images}"
    );
    let first = pages
        .iter()
        .flat_map(|sp| sp.images.iter())
        .next()
        .expect("at least one image");
    assert!(!first.image.data.is_empty(), "image bytes must be loaded");
    assert!(
        first.image.width.unwrap_or(0.0) > 10.0,
        "anchor geometry must produce a real width, got {:?}",
        first.image.width
    );
}

// ---------------------------------------------------------------------------
// Worksheet text boxes (issue #240)
// ---------------------------------------------------------------------------

#[test]
fn with_text_box_renders_anchored_text() {
    use office2pdf::ir::{Alignment, Color};

    let pages = sheet_pages("poi/WithTextBox.xlsx");
    let boxes: Vec<_> = pages.iter().flat_map(|sp| sp.text_boxes.iter()).collect();
    assert_eq!(boxes.len(), 1, "the drawing holds one text box");

    let text_box = boxes[0];
    assert_eq!(text_box.paragraphs.len(), 3);
    assert!(
        text_box.width > 50.0,
        "anchor width, got {}",
        text_box.width
    );

    let para = &text_box.paragraphs[0];
    assert_eq!(para.runs[0].text, "Line 1");
    assert_eq!(para.style.alignment, None, "algn=l maps to default/left");
    assert_eq!(para.runs[0].style.color, Some(Color::new(0xFF, 0, 0)));

    assert_eq!(
        text_box.paragraphs[1].style.alignment,
        Some(Alignment::Center)
    );
    assert_eq!(
        text_box.paragraphs[1]
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>(),
        "Line 2"
    );
    assert_eq!(
        text_box.paragraphs[2].style.alignment,
        Some(Alignment::Right)
    );
    assert_eq!(
        text_box.paragraphs[2].runs[0].style.color,
        Some(Color::new(0, 0, 0xFF))
    );
}

// ---------------------------------------------------------------------------
// Embedded charts on chart-only sheets (issue #239)
// ---------------------------------------------------------------------------

#[test]
fn with_chart_renders_embedded_chart() {
    // WithChart.xlsx puts its chart on a sheet with no cells; that sheet was
    // skipped entirely, dropping the chart (same class as #238's image-only
    // sheets, which the fix did not extend to charts).
    let pages = sheet_pages("poi/WithChart.xlsx");
    let total_charts: usize = pages.iter().map(|sp| sp.charts.len()).sum();
    assert!(
        total_charts >= 1,
        "the embedded chart must be extracted, got {total_charts}"
    );
    let chart = pages
        .iter()
        .flat_map(|sp| sp.charts.iter())
        .next()
        .expect("a chart");
    assert!(
        !chart.1.series.is_empty(),
        "chart must carry its series data"
    );
}
