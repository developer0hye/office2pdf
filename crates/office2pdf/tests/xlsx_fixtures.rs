//! Integration tests for XLSX fixtures.
//!
//! Each real-world `.xlsx` file in `tests/fixtures/xlsx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

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
