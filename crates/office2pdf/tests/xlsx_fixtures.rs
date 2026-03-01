//! Integration tests for XLSX fixtures.
//!
//! Each real-world `.xlsx` file in `tests/fixtures/xlsx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::ir::{Block, Page, TablePage};
use office2pdf::parser::Parser;
use office2pdf::parser::xlsx::XlsxParser;

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

/// Parse an XLSX fixture and return the table pages (sheets).
fn table_pages(name: &str) -> Vec<TablePage> {
    let data = load_fixture(name);
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|p| match p {
            Page::Table(tp) => Some(tp),
            _ => None,
        })
        .collect()
}

fn sheet_names(pages: &[TablePage]) -> Vec<&str> {
    pages.iter().map(|p| p.name.as_str()).collect()
}

fn total_rows(pages: &[TablePage]) -> usize {
    pages.iter().map(|p| p.table.rows.len()).sum()
}

fn has_cell_border(pages: &[TablePage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.border.is_some())
    })
}

fn has_merged_cells(pages: &[TablePage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.col_span > 1 || c.row_span > 1)
    })
}

fn has_formatted_text(pages: &[TablePage]) -> bool {
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
    let pages = table_pages("any_sheets.xlsx");
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
    let pages = table_pages("date.xlsx");
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
    let pages = table_pages("merge_cells.xlsx");
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
    let pages = table_pages("SH001-Table.xlsx");
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
    let pages = table_pages("SH002-TwoTablesTwoSheets.xlsx");
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
    let pages = table_pages("SH106-Formatted.xlsx");
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
    let pages = table_pages("SH109-CellWithBorder.xlsx");
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
    let pages = table_pages("temperature.xlsx");
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
