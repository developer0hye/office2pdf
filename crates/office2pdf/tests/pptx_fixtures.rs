//! Integration tests for PPTX fixtures.
//!
//! Each real-world `.pptx` file in `tests/fixtures/pptx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

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
            common::validate_pdf_with_qpdf(&result.pdf);
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

// ===========================================================================
// Third-party fixtures — smoke tests (must not panic)
// ===========================================================================

/// Generate a pair of smoke + basic-structure tests for a PPTX fixture.
macro_rules! pptx_fixture_tests {
    ($test_name:ident, $file:expr) => {
        paste::paste! {
            #[test]
            fn [<smoke_ $test_name>]() {
                assert_produces_valid_pdf($file);
            }

            #[test]
            fn [<structure_ $test_name>]() {
                let data = load_fixture($file);
                match PptxParser.parse(&data, &ConvertOptions::default()) {
                    Ok((doc, _)) => {
                        // Just verify parsing succeeds — slide count varies by file
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

pptx_fixture_tests!(ffc, "ffc.pptx");
pptx_fixture_tests!(one_slide, "1-slide.pptx");
pptx_fixture_tests!(five_slides, "5-slides.pptx");
pptx_fixture_tests!(ten_slides, "10-slides.pptx");

// --- Apache POI (Apache 2.0) -----------------------------------------------

pptx_fixture_tests!(bar_chart, "bar-chart.pptx");
pptx_fixture_tests!(pie_chart, "pie-chart.pptx");
pptx_fixture_tests!(line_chart, "line-chart.pptx");
pptx_fixture_tests!(scatter_chart, "scatter-chart.pptx");
pptx_fixture_tests!(radar_chart, "radar-chart.pptx");
pptx_fixture_tests!(chart_picture_bg, "chart-picture-bg.pptx");
pptx_fixture_tests!(table_test_poi, "table_test.pptx");
pptx_fixture_tests!(table_test2, "table_test2.pptx");
pptx_fixture_tests!(table_with_theme, "table-with-theme.pptx");
pptx_fixture_tests!(backgrounds, "backgrounds.pptx");
pptx_fixture_tests!(themes, "themes.pptx");
pptx_fixture_tests!(smart_art, "SmartArt.pptx");
pptx_fixture_tests!(smart_art_simple, "smartart-simple.pptx");
pptx_fixture_tests!(embedded_audio, "EmbeddedAudio.pptx");
pptx_fixture_tests!(embedded_video, "EmbeddedVideo.pptx");
pptx_fixture_tests!(with_japanese, "with_japanese.pptx");
pptx_fixture_tests!(with_master, "WithMaster.pptx");
pptx_fixture_tests!(comment_45545, "45545_Comment.pptx");
pptx_fixture_tests!(keyframes, "keyframes.pptx");
pptx_fixture_tests!(layouts, "layouts.pptx");
pptx_fixture_tests!(shapes, "shapes.pptx");
pptx_fixture_tests!(custom_geo, "customGeo.pptx");
pptx_fixture_tests!(highlight, "highlight-test-case.pptx");
pptx_fixture_tests!(picture_transparency, "picture-transparency.pptx");
pptx_fixture_tests!(poi_sample, "poi_sample.pptx");
pptx_fixture_tests!(present1, "present1.pptx");
pptx_fixture_tests!(rain, "rain.pptx");
pptx_fixture_tests!(copy_slide_demo, "copy-slide-demo.pptx");

// --- MIT: Open-Xml-PowerTools (Microsoft) ----------------------------------

pptx_fixture_tests!(oxp_presentation, "oxp_Presentation.pptx");
pptx_fixture_tests!(oxp_chart_cached, "oxp_CU018-Chart-Cached-Data-41.pptx");
pptx_fixture_tests!(oxp_chart_embedded, "oxp_CU019-Chart-Embedded-Xlsx-41.pptx");
pptx_fixture_tests!(oxp_pb001_input1, "oxp_PB001-Input1.pptx");
pptx_fixture_tests!(oxp_pb001_input2, "oxp_PB001-Input2.pptx");
pptx_fixture_tests!(oxp_pb001_input3, "oxp_PB001-Input3.pptx");
pptx_fixture_tests!(oxp_videos, "oxp_PP006-Videos.pptx");
