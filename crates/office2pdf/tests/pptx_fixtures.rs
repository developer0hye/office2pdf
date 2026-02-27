//! Integration tests for PPTX fixtures.
//!
//! Each real-world `.pptx` file in `tests/fixtures/pptx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::ir::{Block, FixedElementKind, FixedPage, Page};
use office2pdf::parser::Parser;
use office2pdf::parser::pptx::PptxParser;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx")
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

/// Parse a PPTX fixture and return the fixed pages (slides).
fn fixed_pages(name: &str) -> Vec<FixedPage> {
    let data = load_fixture(name);
    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|p| match p {
            Page::Fixed(fp) => Some(fp),
            _ => None,
        })
        .collect()
}

fn has_fixed_image(pages: &[FixedPage]) -> bool {
    pages
        .iter()
        .flat_map(|p| p.elements.iter())
        .any(|e| matches!(e.kind, FixedElementKind::Image(_)))
}

fn has_textbox_with_content(pages: &[FixedPage]) -> bool {
    pages
        .iter()
        .flat_map(|p| p.elements.iter())
        .any(|e| match &e.kind {
            FixedElementKind::TextBox(blocks) => blocks.iter().any(|b| match b {
                Block::Paragraph(para) => para.runs.iter().any(|r| !r.text.is_empty()),
                _ => false,
            }),
            _ => false,
        })
}

// ---------------------------------------------------------------------------
// minimal.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_minimal() {
    assert_produces_valid_pdf("minimal.pptx");
}

#[test]
fn structure_minimal() {
    // minimal.pptx contains only slide layouts/masters but no actual slides
    let data = load_fixture("minimal.pptx");
    let result = PptxParser.parse(&data, &ConvertOptions::default());
    match result {
        Ok((doc, _)) => {
            let slides: Vec<_> = doc
                .pages
                .iter()
                .filter(|p| matches!(p, Page::Fixed(_)))
                .collect();
            // 0 slides is the expected result for this fixture
            assert!(
                slides.is_empty(),
                "minimal.pptx has no actual slides, expected 0 pages"
            );
        }
        Err(_) => {
            // Parse error is also acceptable for a file with no slides
        }
    }
}

// ---------------------------------------------------------------------------
// no-slides.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_no_slides() {
    // Must not panic — either empty result or parse error is fine.
    let path = fixture_path("no-slides.pptx");
    let _ = office2pdf::convert(&path);
}

#[test]
fn structure_no_slides() {
    let data = load_fixture("no-slides.pptx");
    match PptxParser.parse(&data, &ConvertOptions::default()) {
        Ok((doc, _)) => {
            // 0 pages is acceptable for a file with no slides
            let slide_count = doc
                .pages
                .iter()
                .filter(|p| matches!(p, Page::Fixed(_)))
                .count();
            assert_eq!(slide_count, 0, "no-slides file should produce 0 pages");
        }
        Err(_) => {
            // Parse error is also acceptable
        }
    }
}

// ---------------------------------------------------------------------------
// powerpoint_sample.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_powerpoint_sample() {
    assert_produces_valid_pdf("powerpoint_sample.pptx");
}

#[test]
fn structure_powerpoint_sample() {
    let pages = fixed_pages("powerpoint_sample.pptx");
    assert!(pages.len() >= 2, "should have >= 2 slides");
    assert!(has_textbox_with_content(&pages), "should have text content");
}

// ---------------------------------------------------------------------------
// powerpoint_with_image.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_powerpoint_with_image() {
    assert_produces_valid_pdf("powerpoint_with_image.pptx");
}

#[test]
fn structure_powerpoint_with_image() {
    let pages = fixed_pages("powerpoint_with_image.pptx");
    assert!(
        has_fixed_image(&pages),
        "should have FixedElementKind::Image"
    );
}

// ---------------------------------------------------------------------------
// test_slides.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_test_slides() {
    assert_produces_valid_pdf("test_slides.pptx");
}

#[test]
fn structure_test_slides() {
    let pages = fixed_pages("test_slides.pptx");
    assert!(!pages.is_empty(), "should have at least 1 slide");
}

// ---------------------------------------------------------------------------
// test.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_test() {
    assert_produces_valid_pdf("test.pptx");
}

#[test]
fn structure_test() {
    let pages = fixed_pages("test.pptx");
    assert!(!pages.is_empty(), "should have at least one slide");
    assert!(has_textbox_with_content(&pages), "should have text content");
}
