#![cfg(not(target_arch = "wasm32"))] // native-only integration tests (fs, qpdf, criterion)
//! Integration tests for PPTX fixtures.
//!
//! Each real-world `.pptx` file in `tests/fixtures/pptx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::internal::Parser;
use office2pdf::internal::PptxParser;
use office2pdf::internal::generate_typst;
use office2pdf::ir::{
    Block, Color, FixedElementKind, FixedPage, LineSpacing, Page, PatternPreset, ShapeKind,
};

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
            FixedElementKind::TextBox(text_box) => text_box.content.iter().any(|b| match b {
                Block::Paragraph(para) => para.runs.iter().any(|r| !r.text.is_empty()),
                _ => false,
            }),
            _ => false,
        })
}

// ---------------------------------------------------------------------------
// PR #188 contributor acceptance fixtures
// ---------------------------------------------------------------------------

const PR_188_PAGE_FILL_FIXTURE: &str = "pr_188_page_fill_reset.pptx";
const PR_188_LAYOUT_GRADIENT_FIXTURE: &str = "pr_188_layout_gradient.pptx";
const PR_188_MASTER_BG_REF_FIXTURE: &str = "pr_188_master_bg_ref.pptx";

#[test]
fn structure_pr_188_contributor_acceptance_supported_behavior() {
    let reset_pages = fixed_pages(PR_188_PAGE_FILL_FIXTURE);
    assert_eq!(reset_pages.len(), 2);
    assert_eq!(
        reset_pages[0].background_color,
        Some(Color::new(0xC0, 0x00, 0x00))
    );
    assert_eq!(reset_pages[1].background_color, None);
    assert!(reset_pages[1].background_gradient.is_none());

    let gradient_pages = fixed_pages(PR_188_LAYOUT_GRADIENT_FIXTURE);
    let gradient = gradient_pages[0]
        .background_gradient
        .as_ref()
        .expect("slide should inherit the layout gradient");
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].color, Color::new(0x11, 0x22, 0x33));
    assert_eq!(gradient.stops[1].color, Color::new(0x44, 0x55, 0x66));

    let bg_ref_pages = fixed_pages(PR_188_MASTER_BG_REF_FIXTURE);
    assert_eq!(
        bg_ref_pages[0].background_color,
        Some(Color::new(0x44, 0x72, 0xC4)),
        "the master's bgRef should resolve the first theme background fill with accent1"
    );
}

#[test]
fn smoke_pr_188_contributor_acceptance_fixtures() {
    for fixture in [
        PR_188_PAGE_FILL_FIXTURE,
        PR_188_LAYOUT_GRADIENT_FIXTURE,
        PR_188_MASTER_BG_REF_FIXTURE,
    ] {
        assert_produces_valid_pdf(fixture);
    }
}

#[test]
fn acceptance_pr_188_contributor_acceptance_page_fill_reset() {
    let data = load_fixture(PR_188_PAGE_FILL_FIXTURE);
    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let output = generate_typst(&document).expect("fixture should generate Typst");
    let page_settings = output
        .source
        .lines()
        .filter(|line| line.starts_with("#set page("))
        .collect::<Vec<_>>();

    assert_eq!(page_settings.len(), 2);
    assert!(page_settings[0].contains("fill: rgb(192, 0, 0)"));
    assert!(page_settings[1].contains("fill: white"));
}

// ---------------------------------------------------------------------------
// pattern-fill.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_pattern_fill() {
    assert_produces_valid_pdf("pattern-fill.pptx");
}

#[test]
fn structure_pattern_fill_preserves_hatch_and_omits_outline() {
    let pages = fixed_pages("pattern-fill.pptx");
    assert_eq!(pages.len(), 1);

    let patterned_shape = pages[0]
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::Shape(shape) => Some(shape),
            _ => None,
        })
        .find(|shape| shape.pattern_fill.is_some())
        .expect("fixture should contain a patterned shape");
    let pattern = patterned_shape.pattern_fill.as_ref().unwrap();
    assert_eq!(pattern.preset, PatternPreset::LightUpwardDiagonal);
    assert_eq!(pattern.foreground, Color::new(0, 0, 255));
    assert_eq!(pattern.background, Color::new(255, 255, 255));
    assert!(patterned_shape.stroke.is_none());
}

// ---------------------------------------------------------------------------
// paragraph-boundary-spacing.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_paragraph_boundary_spacing() {
    assert_produces_valid_pdf("paragraph-boundary-spacing.pptx");
}

#[test]
fn structure_paragraph_boundary_spacing_keeps_three_unstyled_paragraphs() {
    let pages = fixed_pages("paragraph-boundary-spacing.pptx");
    assert_eq!(pages.len(), 1);
    let text_box = pages[0]
        .elements
        .iter()
        .find_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box),
            _ => None,
        })
        .expect("fixture should contain its text frame");
    let paragraphs: Vec<_> = text_box
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();

    assert_eq!(paragraphs.len(), 3);
    for (index, paragraph) in paragraphs.iter().enumerate() {
        assert!(
            matches!(
                paragraph.style.line_spacing,
                Some(LineSpacing::Proportional(factor)) if (factor - 1.0).abs() < f64::EPSILON
            ),
            "paragraph should retain its 100% line spacing: {:?}",
            paragraph.style.line_spacing
        );
        assert!(paragraph.runs[0].style.font_family.is_none());
        assert_eq!(paragraph.runs[0].style.font_size, Some(18.0));
        assert!(
            paragraph.runs[0]
                .text
                .starts_with(&format!("Paragraph {}", ["one", "two", "three"][index]))
        );
    }
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
// PDF text content verification
// ===========================================================================

/// Helper: convert a PPTX fixture to PDF and extract text.
fn pdf_text(name: &str) -> String {
    let path = fixture_path(name);
    let result = office2pdf::convert(&path).expect("conversion should succeed");
    common::extract_pdf_text(&result.pdf)
}

// ---------------------------------------------------------------------------
// powerpoint_sample.pptx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_powerpoint_sample() {
    let text = pdf_text("powerpoint_sample.pptx");
    assert!(
        text.contains("slide title") || text.contains("Slide Title") || text.contains("Test"),
        "PDF should contain slide title text"
    );
}

// ---------------------------------------------------------------------------
// test.pptx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_test() {
    let text = pdf_text("test.pptx");
    assert!(
        text.contains("Presentation Title") || text.contains("Title"),
        "PDF should contain presentation title"
    );
}

// ---------------------------------------------------------------------------
// test_slides.pptx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_test_slides() {
    let text = pdf_text("test_slides.pptx");
    assert!(
        text.contains("Test text") || text.contains("Box"),
        "PDF should contain slide text content"
    );
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

/// `customGeo.pptx` slide 22 repeats `startAt="6"` on every paragraph in
/// one authored ordered list. PowerPoint uses that repeated metadata to seed
/// the sequence once, then renders 6 through 10 (issue #1347).
#[test]
fn custom_geo_slide_22_repeated_start_at_continues_one_sequence() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[21];
    let matching_lists: Vec<_> = page
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(&text_box.content),
            _ => None,
        })
        .flat_map(|content| content.iter())
        .filter_map(|block| match block {
            Block::List(list)
                if list.items.iter().any(|item| {
                    item.content.iter().any(|paragraph| {
                        paragraph
                            .runs
                            .iter()
                            .any(|run| run.text.starts_with("Provide extensive research"))
                    })
                }) =>
            {
                Some(list)
            }
            _ => None,
        })
        .collect();

    assert_eq!(
        matching_lists.len(),
        1,
        "the five consecutive principles must remain one ordered list"
    );
    let list = matching_lists[0];
    assert_eq!(list.kind, office2pdf::ir::ListKind::Ordered);
    assert_eq!(list.items.len(), 5);
    assert_eq!(
        list.items
            .iter()
            .map(|item| item.start_at)
            .collect::<Vec<_>>(),
        vec![Some(6), None, None, None, None]
    );
}

#[test]
fn custom_geo_slide_5_keeps_a_latin_typeface_scoped_to_its_declaring_run() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[4];
    let paragraph = page
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(&text_box.content),
            _ => None,
        })
        .flat_map(|content| content.iter())
        .find_map(|block| {
            let Block::Paragraph(paragraph) = block else {
                return None;
            };
            paragraph
                .runs
                .iter()
                .any(|run| run.text.starts_with("the state board of education"))
                .then_some(paragraph)
        })
        .expect("slide 5 standards body paragraph");

    assert_eq!(paragraph.runs.len(), 4);
    for run in &paragraph.runs[..3] {
        assert_eq!(
            run.style.font_family.as_deref(),
            Some("Calibri"),
            "a complex-script typeface must not replace the inherited Latin family for {:?}",
            run.text
        );
    }
    assert_eq!(paragraph.runs[3].text, ".");
    assert_eq!(
        paragraph.runs[3].style.font_family.as_deref(),
        Some("Times New Roman"),
        "the final punctuation run keeps the Latin family it explicitly declares"
    );
}

#[test]
fn custom_geo_slide_6_title_uses_its_saved_normal_autofit_scale() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[5];
    let (text_box, run) = page
        .elements
        .iter()
        .find_map(|element| {
            let FixedElementKind::TextBox(text_box) = &element.kind else {
                return None;
            };
            text_box.content.iter().find_map(|block| {
                let Block::Paragraph(paragraph) = block else {
                    return None;
                };
                paragraph
                    .runs
                    .iter()
                    .find(|run| run.text == "ELA Standards Framework")
                    .map(|run| (text_box, run))
            })
        })
        .expect("slide 6 title run");

    assert_eq!(
        run.style.font_size,
        Some(32.0),
        "the native export paints 36pt x 0.9 = 32.4pt at a whole 32pt"
    );
    assert!(
        !text_box.auto_fit,
        "fontScale=90000 already states PowerPoint's fitted size"
    );
}

#[test]
fn custom_geo_slide_6_callout_uses_its_geometry_text_rectangle() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[5];
    let overlay = page
        .elements
        .iter()
        .find(|element| {
            let FixedElementKind::TextBox(text_box) = &element.kind else {
                return false;
            };
            text_box.content.iter().any(|block| {
                let Block::Paragraph(paragraph) = block else {
                    return false;
                };
                paragraph.runs.iter().any(|run| {
                    run.text
                        .starts_with("\u{201c}What\u{201d} students should know")
                })
            })
        })
        .expect("slide 6 rounded callout text overlay");

    let radius = 114.0 * 16_667.0 / 100_000.0;
    let inset = radius * 29_289.0 / 100_000.0;
    assert!((overlay.x - (66.0 + inset)).abs() < 1e-9);
    assert!((overlay.y - (237.5 + inset)).abs() < 1e-9);
    assert!((overlay.width - (150.0 - 2.0 * inset)).abs() < 1e-9);
    assert!((overlay.height - (114.0 - 2.0 * inset)).abs() < 1e-9);
}

#[test]
fn custom_geo_slide_6_callouts_keep_their_rounded_wedge_outlines() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[5];
    let expected_callouts = [
        (150.0, 114.0, (0.912_42, 1.422_45)),
        (180.0, 84.0, (-0.393_65, 0.634_68)),
        (180.0, 72.0, (-0.374_57, 0.679_05)),
    ];

    for (width, height, expected_tip) in expected_callouts {
        let shape = page
            .elements
            .iter()
            .find_map(|element| {
                if (element.width - width).abs() >= 0.01 || (element.height - height).abs() >= 0.01
                {
                    return None;
                }
                let FixedElementKind::Shape(shape) = &element.kind else {
                    return None;
                };
                Some(shape)
            })
            .unwrap_or_else(|| panic!("slide 6 callout {width}x{height}pt"));
        let ShapeKind::Path { subpaths } = &shape.kind else {
            panic!("slide 6 callout {width}x{height}pt must retain its preset path");
        };
        assert_eq!(subpaths.len(), 1);
        assert!(subpaths[0].closed);
        assert!(subpaths[0].vertices.iter().any(|vertex| {
            (vertex.0 - expected_tip.0).abs() < 1e-6 && (vertex.1 - expected_tip.1).abs() < 1e-6
        }));
    }
}

#[test]
fn custom_geo_open_connectors_cast_their_declared_stroke_shadows() {
    let data = load_fixture("customGeo.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = generate_typst(&document).unwrap().source;
    let connector_shadows: Vec<&str> = source
        .lines()
        .filter(|line| {
            line.contains("#curve(stroke:")
                && line.contains("paint: rgb(0, 0, 0, 89)")
                && line.contains("curve.move(")
        })
        .collect();

    assert!(
        connector_shadows.len() >= 4,
        "slide 6's four open connectors must cast their declared 35%-opaque shadows: {connector_shadows:?}"
    );
    assert!(
        connector_shadows
            .iter()
            .all(|line| line.contains("thickness: 2.0001574803149604pt")
                && line.contains("join: \"round\"")
                && !line.contains("fill:")
                && !line.contains("curve.close()")),
        "each connector shadow must retain the source stroke without closing or filling it: {connector_shadows:?}"
    );
}

#[test]
fn structure_custom_geo_page_31_preserves_theme_hyperlink_runs() {
    let pages = fixed_pages("customGeo.pptx");
    let page = &pages[30];
    let mut runs = Vec::new();
    let mut marker_styles = Vec::new();

    for element in &page.elements {
        let FixedElementKind::TextBox(text_box) = &element.kind else {
            continue;
        };
        for block in &text_box.content {
            match block {
                Block::Paragraph(paragraph) => runs.extend(paragraph.runs.iter()),
                Block::List(list) => {
                    marker_styles.extend(
                        list.level_styles
                            .values()
                            .filter_map(|level| level.marker_style.as_ref()),
                    );
                    for item in &list.items {
                        for paragraph in &item.content {
                            runs.extend(paragraph.runs.iter());
                        }
                    }
                }
                _ => {}
            }
        }
    }

    let expected_links = [
        "www.corestandards.org",
        "www.commoncore.org",
        "www.education.ohio.gov",
        "www.achievethecore.org",
    ];
    for expected in expected_links {
        let run = runs
            .iter()
            .find(|run| run.text.contains(expected))
            .unwrap_or_else(|| panic!("missing hyperlink run: {expected}"));
        assert_eq!(run.style.color, Some(Color::new(0x00, 0x00, 0xFF)));
        assert_eq!(run.style.underline, Some(true));
    }

    let control = runs
        .iter()
        .find(|run| run.text.contains("Curriculum maps linked to Common Core"))
        .expect("control text should remain present");
    assert_ne!(control.style.color, Some(Color::new(0x00, 0x00, 0xFF)));
    assert_ne!(control.style.underline, Some(true));
    assert!(
        marker_styles
            .iter()
            .all(|style| style.color != Some(Color::new(0x00, 0x00, 0xFF)))
    );
}
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

#[test]
fn smart_art_renders_cached_drawing_shapes() {
    // The SmartArt drawing cache holds five shapes; PowerPoint renders them
    // as blue blocks, but office2pdf produced a blank slide (issue #223).
    let pages = fixed_pages("SmartArt.pptx");
    let shape_count: usize = pages
        .iter()
        .flat_map(|p| p.elements.iter())
        .filter(|e| matches!(e.kind, FixedElementKind::Shape(_)))
        .count();
    assert!(
        shape_count >= 5,
        "SmartArt drawing cache must render its shapes, got {shape_count}"
    );
    // The shapes carry a fill (the accent color), not a blank slide.
    let filled: bool = pages
        .iter()
        .flat_map(|p| p.elements.iter())
        .any(|e| matches!(&e.kind, FixedElementKind::Shape(s) if s.fill.is_some()));
    assert!(filled, "SmartArt shapes must carry their fill color");
}

// ---------------------------------------------------------------------------
// hangul_kinsoku_terminal_punct.pptx — Hangul + trailing punctuation at a
// line boundary (issue #438). Authored and exported by Windows PowerPoint:
// each box holds ten 18pt Malgun Gothic syllables (180pt) in a 183pt-usable
// box, plus one trailing mark, so the mark alone overflows the line.
// ---------------------------------------------------------------------------

/// Runs only where Malgun Gothic exists (the Windows CI runner): without
/// a Hangul-capable font the glyphs never reach the PDF text layer at
/// all, which would vacuously pass the ZWSP check and fail the content
/// check.
#[cfg(target_os = "windows")]
#[test]
fn kinsoku_fixture_text_layer_has_no_zero_width_space() {
    // The break opportunity is carried as U+200B inside the IR, and must
    // never surface in the PDF text layer.
    let text = pdf_text("hangul_kinsoku_terminal_punct.pptx");
    assert!(
        !text.contains('\u{200B}'),
        "zero-width space leaked into the PDF text layer"
    );
    assert!(
        text.contains("가나다라마바사아자차"),
        "fixture text lost during conversion"
    );
}

#[test]
fn kinsoku_break_marker_becomes_inline_box() {
    let data = load_fixture("hangul_kinsoku_terminal_punct.pptx");
    let parser = PptxParser;
    let (document, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = generate_typst(&document).unwrap();
    // Ten syllables, then the boxed mark: an inline box is a Contingent
    // Break in UAX #14, so the mark may start the next line instead of
    // dragging the syllable down with it (LB13). Verified against native
    // PowerPoint exports in #438.
    assert!(
        output.source.contains("자차]#box[]"),
        "no break frame before the terminal mark"
    );
    // '%' is the counter-case: PowerPoint keeps it glued to the syllable,
    // so it must never be boxed away from it.
    assert!(
        output.source.contains("차%]"),
        "'%' must stay glued to the syllable, with no break frame between"
    );
    assert!(
        !output.source.contains('\u{200B}'),
        "kinsoku marker leaked into the Typst source"
    );
    // The 09_lecture_ko card replica: the break lands before '?', not
    // between the last two syllables.
    assert!(
        output.source.contains("뜻인가]#box[]"),
        "no break frame in the #438 replica"
    );
}

// ---------------------------------------------------------------------------
// shadow_blur_radii.pptx — outer shadows at blurRad 6/12/24pt plus the #390
// reproduction (9pt), authored and exported by Windows PowerPoint.
// ---------------------------------------------------------------------------

#[test]
fn shadow_blur_uses_one_gaussian_asset_per_shape() {
    let data = load_fixture("shadow_blur_radii.pptx");
    let parser = PptxParser;
    let (document, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = generate_typst(&document).unwrap();
    let shadows: Vec<&str> = output
        .images
        .iter()
        .filter_map(|asset| std::str::from_utf8(&asset.data).ok())
        .filter(|svg| svg.contains("<feGaussianBlur"))
        .collect();
    assert_eq!(
        shadows.len(),
        4,
        "each of the four blurred rectangles should emit one filtered asset"
    );
    assert!(
        shadows.iter().all(|svg| svg.contains("<rect")),
        "each filter must blur the rectangle silhouette itself: {shadows:?}"
    );
    // blur 24pt: sigma 8pt (blurRad/3, #784), followed to 2.6 sigma on
    // either side, so the 220x130pt frame becomes a 261.6x171.6pt asset.
    assert!(
        shadows.iter().any(|svg| {
            svg.contains("stdDeviation=\"8\"")
                && svg.contains("width=\"261.6\"")
                && svg.contains("height=\"171.6\"")
        }),
        "24pt blur must span the declared Gaussian extent: {shadows:?}"
    );
    assert_eq!(
        output.source.matches("#pdf.artifact(image(\"").count(),
        4,
        "decorative shadows should be marked as artifacts"
    );
}

// ---------------------------------------------------------------------------
// Repository introduction deck — thirty-slide Korean presentation
// ---------------------------------------------------------------------------

/// Thirty-slide 16:9 Korean deck introducing this repository. The focused
/// fixtures above each isolate one feature; this one puts the features a real
/// deck combines on the same slide: a gradient title slide, sectioned body
/// layouts, native tables, 71 raster icons, monospaced code blocks,
/// `a:fld` slide numbers, character tracking, per-paragraph line spacing, and
/// stacked bar, radar, and doughnut charts.
const INTRODUCTION_DECK_FIXTURE: &str = "office2pdf_introduction_ko.pptx";

/// Converting all thirty slides costs ~80s in a debug build, too much for the
/// default suite to pay on every platform. The fast path renders the three
/// slides that carry the deck's hardest content — the gradient title slide, the
/// stacked bar chart, and the monospaced code blocks — and the whole deck is
/// covered by the `#[ignore]`d test below.
#[test]
fn smoke_introduction_deck_representative_slides() {
    let path = fixture_path(INTRODUCTION_DECK_FIXTURE);
    for slide in [1u32, 17, 20] {
        let options = ConvertOptions {
            slide_range: Some(office2pdf::config::SlideRange::new(slide, slide)),
            ..ConvertOptions::default()
        };
        let result = office2pdf::convert_with_options(&path, &options)
            .unwrap_or_else(|error| panic!("slide {slide} must convert: {error}"));
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "slide {slide} output should start with PDF magic bytes"
        );
        common::validate_pdf_with_qpdf(&result.pdf);
    }
}

// Converting all 30 slides costs ~80s in a debug build, so the whole-deck run
// is opt-in via `--ignored` rather than part of the default suite.
#[test]
#[ignore]
fn smoke_introduction_deck_full_fixture() {
    assert_produces_valid_pdf(INTRODUCTION_DECK_FIXTURE);
}

#[test]
fn structure_introduction_deck_keeps_one_page_per_slide() {
    let pages = fixed_pages(INTRODUCTION_DECK_FIXTURE);

    assert_eq!(pages.len(), 30, "PowerPoint prints one page per slide");
    for (index, page) in pages.iter().enumerate() {
        assert!(
            (page.size.width - 960.0).abs() < 0.5 && (page.size.height - 540.0).abs() < 0.5,
            "slide {} must keep the deck's 16:9 stage, got {}x{}",
            index + 1,
            page.size.width,
            page.size.height
        );
    }
}

#[test]
fn structure_introduction_deck_carries_slide_text_in_order() {
    let pages = fixed_pages(INTRODUCTION_DECK_FIXTURE);

    let slide_text = |index: usize| -> String {
        pages[index]
            .elements
            .iter()
            .filter_map(|element| match &element.kind {
                FixedElementKind::TextBox(text_box) => Some(text_box),
                _ => None,
            })
            .flat_map(|text_box| text_box.content.iter())
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .flat_map(|paragraph| paragraph.runs.iter())
            .map(|run| run.text.as_str())
            .collect::<String>()
    };

    assert!(
        slide_text(0).contains("office2pdf"),
        "the title slide keeps its wordmark"
    );
    assert!(
        slide_text(0).contains("DOCX"),
        "the title slide keeps its format list"
    );
    assert!(
        slide_text(19).contains("ConvertOptions"),
        "slide 20 keeps its code block"
    );
}

#[test]
fn structure_introduction_deck_places_raster_icons_on_slides() {
    let pages = fixed_pages(INTRODUCTION_DECK_FIXTURE);

    let images: usize = pages
        .iter()
        .flat_map(|page| page.elements.iter())
        .filter(|element| matches!(element.kind, FixedElementKind::Image(_)))
        .count();

    // The slides carry 71 `<p:pic>` shapes between them, and slide layout 3
    // contributes one more wherever it is applied. Asserting the slide-level
    // floor keeps the test from breaking when a slide is re-authored, while
    // still failing if image extraction regresses.
    assert!(
        images >= 71,
        "the deck's raster icons must survive parsing, got {images}"
    );
}

/// `p:defaultTextStyle` supplies the size for table cell text that declares
/// none (issue #675).
///
/// `table-with-theme.pptx` has no `sz` on any run; `ppt/presentation.xml`
/// declares `<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800"/>`. Without
/// consulting it the cells carried no size at all and fell through to Typst's
/// own 11pt default.
#[test]
fn structure_table_with_theme_takes_size_from_default_text_style() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/poi/table-with-theme.pptx");
    let data = std::fs::read(&path).expect("fixture");
    let (doc, _w) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("parses");

    let mut sizes: Vec<Option<f64>> = Vec::new();
    for page in &doc.pages {
        let Page::Fixed(fixed) = page else { continue };
        for element in &fixed.elements {
            let FixedElementKind::Table(table) = &element.kind else {
                continue;
            };
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.content {
                        if let Block::Paragraph(paragraph) = block {
                            for run in &paragraph.runs {
                                if !run.text.trim().is_empty() {
                                    sizes.push(run.style.font_size);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    assert!(!sizes.is_empty(), "the fixture has table cell runs");
    assert!(
        sizes.iter().all(|size| *size == Some(18.0)),
        "every cell run takes 18pt from p:defaultTextStyle, got {sizes:?}"
    );
}

/// `a:xfrm/@rot` on a `p:pic` reaches the IR (issue #682).
///
/// The shape path read `rot` from its own `a:xfrm`; the picture path set only
/// the nesting flags and dropped it, so a rotated picture drew upright. Slide
/// 1's decorative picture declares `rot="360000"` (6 degrees) and slide 7's
/// declares `rot="480000"` (8 degrees).
#[test]
fn structure_introduction_ko_keeps_picture_rotation() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/office2pdf_introduction_ko.pptx");
    let data = std::fs::read(&path).expect("fixture");
    let (doc, _w) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("parses");

    let rotation_on = |page_index: usize| -> Vec<Option<f64>> {
        let Some(Page::Fixed(fixed)) = doc.pages.get(page_index) else {
            return Vec::new();
        };
        fixed
            .elements
            .iter()
            .filter_map(|element| match &element.kind {
                FixedElementKind::Image(image) => Some(image.rotation_deg),
                _ => None,
            })
            .collect()
    };

    assert!(
        rotation_on(0).contains(&Some(6.0)),
        "slide 1's picture declares rot=360000 (6 deg), got {:?}",
        rotation_on(0)
    );
    assert!(
        rotation_on(6).contains(&Some(8.0)),
        "slide 7's picture declares rot=480000 (8 deg), got {:?}",
        rotation_on(6)
    );
}

/// A doughnut chart plots as a ring rather than degrading to a data table
/// (issue #679).
///
/// Page 29's `c:doughnutChart` previously rendered as a bordered rectangle
/// holding an italic caption and a `Category | <series>` table, which collided
/// with the separate `99.5%` overlay the deck draws over the real doughnut.
#[test]
fn structure_introduction_ko_plots_the_doughnut_chart() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/office2pdf_introduction_ko.pptx");
    let data = std::fs::read(&path).expect("fixture");
    let (doc, _w) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("parses");

    let charts: Vec<&office2pdf::ir::Chart> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Fixed(fixed) => Some(fixed),
            _ => None,
        })
        .flat_map(|fixed| fixed.elements.iter())
        .filter_map(|element| match &element.kind {
            FixedElementKind::Chart(chart) => Some(chart.as_ref()),
            _ => None,
        })
        .collect();

    let doughnut = charts
        .iter()
        .find(|chart| matches!(chart.chart_type, office2pdf::ir::ChartType::Doughnut))
        .expect("the deck's doughnut is a plotted type, not a caption fallback");

    // `<c:holeSize val="62"/>` in chart3.xml.
    assert_eq!(doughnut.hole_size_percent, Some(62));
    assert!(
        !doughnut.series.is_empty() && doughnut.series[0].values.iter().any(|v| *v > 0.0),
        "the doughnut carries its plotted values"
    );
}

// ---------------------------------------------------------------------------
// rotated_text_box.pptx
// ---------------------------------------------------------------------------

#[test]
fn smoke_rotated_text_box() {
    assert_produces_valid_pdf("rotated_text_box.pptx");
}

/// The fixture's three rotated rails keep the angles their `a:xfrm rot`
/// declares, and its unrotated heading keeps none (issue #894).
#[test]
fn structure_rotated_text_box_keeps_each_declared_angle() {
    let pages = fixed_pages("rotated_text_box.pptx");
    assert_eq!(pages.len(), 1);

    let mut angles: Vec<Option<i64>> = pages[0]
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => {
                Some(text_box.shape_rotation_deg.map(|deg| deg.round() as i64))
            }
            _ => None,
        })
        .collect();
    angles.sort();
    assert_eq!(
        angles,
        vec![None, Some(45), Some(90), Some(270)],
        "one unrotated heading and the 45/90/270 degree rails"
    );
}

// ---------------------------------------------------------------------------
// hard_break_line_advance.pptx — one `wrap="none"` caption column per size
// (6, 8, 9, 10, 11, 12 and 14pt Arial), each holding four single-word lines
// separated by `<a:br/>` (issue #1115). A `<a:br/>` reaches the IR as a run
// with no run properties, so the paragraph states no size every run agrees on;
// the line box then has to carry the size it was derived from itself.
//
// A native PowerPoint 16 export of this deck advances each column at
// 1.16-1.21em — its per-baseline dither around 1.2em — while every column under
// 11pt used to advance a flat 13.20pt.
// ---------------------------------------------------------------------------

/// Every caption column states PowerPoint's `1.2 x size` line, whatever the
/// size its runs declare.
#[test]
fn hard_break_columns_state_their_own_line_advance() {
    let data = load_fixture("hard_break_line_advance.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = generate_typst(&document).unwrap().source;

    let advances: Vec<f64> = source
        .match_indices("#set text(top-edge: ")
        .filter_map(|(offset, needle)| {
            let rest: &str = &source[offset + needle.len()..];
            let (top, rest) = rest.split_once("pt, bottom-edge: -")?;
            let (bottom, _) = rest.split_once("pt)")?;
            Some(top.parse::<f64>().ok()? + bottom.parse::<f64>().ok()?)
        })
        .collect();

    // Seven caption columns plus their seven 14pt labels.
    let mut expected: Vec<f64> = [6.0, 8.0, 9.0, 10.0, 11.0, 12.0, 14.0]
        .into_iter()
        .chain(std::iter::repeat_n(14.0, 7))
        .map(|size: f64| 1.2 * size)
        .collect();
    expected.sort_by(f64::total_cmp);
    let mut got: Vec<f64> = advances;
    got.sort_by(f64::total_cmp);

    assert_eq!(got.len(), expected.len(), "line boxes emitted: {got:?}");
    for (got, want) in got.iter().zip(&expected) {
        assert!(
            (got - want).abs() < 0.001,
            "a slide line spans 1.2 x its size: expected {expected:?}, got {got:?}"
        );
    }
}

#[test]
fn hard_break_line_advance_smoke() {
    assert_produces_valid_pdf("hard_break_line_advance.pptx");
}

// ---------------------------------------------------------------------------
// percentage_line_spacing_expanded.pptx — a minimal native PowerPoint probe
// with a top-anchored 10pt Arial paragraph at 150% line spacing. Its four hard
// broken lines isolate the above-100% baseline regime from wrapping and other
// slide layout behavior (issue #1254).
// ---------------------------------------------------------------------------

#[test]
fn percentage_line_spacing_expanded_keeps_the_source_ratio_and_seat() {
    let data = load_fixture("percentage_line_spacing_expanded.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let paragraph = document
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Fixed(page) => Some(page),
            _ => None,
        })
        .flat_map(|page| page.elements.iter())
        .filter_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box),
            _ => None,
        })
        .flat_map(|text_box| text_box.content.iter())
        .find_map(|block| match block {
            Block::Paragraph(paragraph)
                if paragraph.runs.iter().any(|run| run.text == "Hxg10a") =>
            {
                Some(paragraph)
            }
            _ => None,
        })
        .expect("fixture should contain its 10pt / 150% body paragraph");

    assert!(
        matches!(
            paragraph.style.line_spacing,
            Some(LineSpacing::Proportional(factor)) if (factor - 1.5).abs() < f64::EPSILON
        ),
        "the fixture must preserve its 150% a:spcPct value: {:?}",
        paragraph.style.line_spacing
    );

    let source = generate_typst(&document).unwrap().source;
    assert!(
        source.contains("#set text(top-edge: 14pt, bottom-edge: -4pt)"),
        "the 10pt / 150% line must use PowerPoint's measured 14pt seat and 18pt advance: {source}"
    );
}

#[test]
fn percentage_line_spacing_expanded_smoke() {
    assert_produces_valid_pdf("percentage_line_spacing_expanded.pptx");
}

// ---------------------------------------------------------------------------
// run-fill-alpha.pptx
//
// Three 32pt Arial lines over one `32D6A6` backdrop, declaring the same black
// `a:solidFill` at `a:alpha` 100%, 50% and 25% (issue #1121). A native
// PowerPoint 16.112 export composites the ink at exactly those fractions: the
// darkest pixel of each line reads (0,0,0), (24,106,82) and (37,160,124)
// against the backdrop's (50,214,166). We used to drop the alpha and print all
// three as solid black.
// ---------------------------------------------------------------------------

#[test]
fn smoke_run_fill_alpha() {
    assert_produces_valid_pdf("run-fill-alpha.pptx");
}

/// Each line's declared opacity reaches the IR.
#[test]
fn structure_run_fill_alpha_keeps_each_declared_opacity() {
    let pages = fixed_pages("run-fill-alpha.pptx");
    assert_eq!(pages.len(), 1);
    let text_box = pages[0]
        .elements
        .iter()
        .find_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box),
            _ => None,
        })
        .expect("fixture should contain its label frame");
    let opacities: Vec<Option<f64>> = text_box
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .map(|paragraph| paragraph.runs[0].style.color_alpha)
        .collect();

    assert_eq!(opacities, vec![None, Some(0.5), Some(0.25)]);
}

/// The opacity survives codegen, so the ink composites against the backdrop
/// instead of printing as solid black.
#[test]
fn run_fill_alpha_composites_in_the_generated_source() {
    let data = load_fixture("run-fill-alpha.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = generate_typst(&document).unwrap().source;

    assert!(
        source.contains("fill: rgb(0, 0, 0, 128)"),
        "the 50% line should composite at 128/255: {source}"
    );
    assert!(
        source.contains("fill: rgb(0, 0, 0, 64)"),
        "the 25% line should composite at 64/255: {source}"
    );
}

/// A paragraph that writes no `<a:endParaRPr>` at all puts no face of its own
/// on PowerPoint's shared line box, and still seats where the native export
/// of this fixture seats it.
///
/// This test asserted `Calisto MT` — the theme's minor Latin font, which
/// `presentation.xml`'s `<a:defaultTextStyle>` names — until the native
/// one-factor probes of issue #1645 measured that an *absent* mark inherits
/// nothing at all, list style or theme. #1179's own resolution records that
/// its frame cannot tell the two apart on the first baseline: at 32pt the
/// shared Arial/Calisto MT share (0.963654em) and Arial's own (0.972378em)
/// both round to the 31pt seat its export reads, so the 202.56pt baseline
/// that issue measured is unchanged.
///
/// The deck's own second and third paragraphs do separate them. They carry
/// `a:alpha`, so a native PowerPoint 16 export rasterises both to identical
/// 306.48x66.72pt sprites, placed at y 200.88 and 239.04 — a 38.16pt step,
/// and two identical sprites step by exactly their baseline gap. Rounding
/// each absolute story position (#1259) off Arial's own share gives seats 31,
/// 70 and 108 below the content top, a 38pt step; off the shared share it
/// gives 31, 69 and 108, a 39pt step, which the export's 0.24pt position grid
/// cannot reach. The `seat` assertion below pins the first baseline; the mark
/// family is the model that also puts the other two where the sprites are.
#[test]
fn structure_run_fill_alpha_marks_add_no_face_to_their_line() {
    let data = load_fixture("run-fill-alpha.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Fixed(slide) = &document.pages[0] else {
        panic!("the fixture's only page is a slide");
    };
    let text_box = slide
        .elements
        .iter()
        .find_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box),
            _ => None,
        })
        .expect("fixture should contain its label frame");
    let mark_families: Vec<Option<&str>> = text_box
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .map(|paragraph| paragraph.style.paragraph_mark_font_family.as_deref())
        .collect();

    assert_eq!(
        mark_families,
        vec![Some("Arial"); 3],
        "an absent mark stays on the face of the run it follows"
    );

    let source = generate_typst(&document).unwrap().source;
    let seat_pt: f64 = source
        .split("top-edge: ")
        .nth(1)
        .and_then(|rest| rest.split_once("pt"))
        .and_then(|(seat, _)| seat.parse().ok())
        .unwrap_or_else(|| panic!("the label frame emits a line box: {source}"));
    assert!(
        (seat_pt - 31.0).abs() < 0.001,
        "the 32pt label must keep the 31pt seat PowerPoint's export reads          (issue #1179), got {seat_pt}pt"
    );
}

// ---------------------------------------------------------------------------
// hard_break_wrapped_line_advance.pptx — `hard_break_line_advance.pptx` with
// `<a:bodyPr wrap="none">` changed to `wrap="square"` (issue #1172). A wrapping
// box paces its hard-broken lines through a measured `#stack` of per-line
// boxes, a different path from the `#set text` edges the `wrap="none"` deck
// takes, and that path floored every line box at 10pt: each column under 10pt
// advanced a flat 12.00pt where the native export paces it at `1.2 x size`.
// ---------------------------------------------------------------------------

/// Every stacked line box spans `1.2 x` the size of the line it holds.
#[test]
fn wrapped_hard_break_columns_stack_their_own_line_box() {
    let data = load_fixture("hard_break_wrapped_line_advance.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = generate_typst(&document).unwrap().source;

    let mut heights: Vec<f64> = source
        .match_indices("pt, height: ")
        .filter_map(|(offset, needle)| {
            let rest: &str = &source[offset + needle.len()..];
            let (height, tail) = rest.split_once("pt)[#place(")?;
            let _ = tail;
            height.parse::<f64>().ok()
        })
        .collect();
    heights.sort_by(f64::total_cmp);

    // Four stacked lines per caption column; the 14pt labels carry no break.
    let mut expected: Vec<f64> = [6.0_f64, 8.0, 9.0, 10.0, 11.0, 12.0, 14.0]
        .into_iter()
        .flat_map(|size| std::iter::repeat_n(1.2 * size, 4))
        .collect();
    expected.sort_by(f64::total_cmp);

    assert_eq!(
        heights.len(),
        expected.len(),
        "stacked line boxes emitted: {heights:?}"
    );
    for (got, want) in heights.iter().zip(&expected) {
        assert!(
            (got - want).abs() < 0.001,
            "a wrapping box's hard-broken line spans 1.2 x its size: \
             expected {expected:?}, got {heights:?}"
        );
    }
}

#[test]
fn hard_break_wrapped_line_advance_smoke() {
    assert_produces_valid_pdf("hard_break_wrapped_line_advance.pptx");
}

// ---------------------------------------------------------------------------
// hard_break_mixed_size_seat.pptx — a dedicated minimal probe containing a
// caption and nine alternating 20pt/10pt Arial lines with no `<a:lnSpc>`
// (issue #1252). PowerPoint rounds the seat of each physical line at that
// line's own size. The paragraph-wide split used to round at 20pt once and
// rescale the ratio, emitting a 9.5pt seat for every 10pt line instead of the
// independently rounded 10pt seat.
// ---------------------------------------------------------------------------

#[test]
fn hard_break_mixed_size_seat_smoke() {
    assert_produces_valid_pdf("hard_break_mixed_size_seat.pptx");
}

// ---------------------------------------------------------------------------
// hard_break_table_row_min_height.pptx — one 20pt-high PowerPoint table row
// containing nine hard-broken 6pt Arial lines with zero cell margins. The
// declared `a:tr/@h` is a floor: native PowerPoint grows the row and keeps all
// nine lines in the PDF text layer (issue #1253).
// ---------------------------------------------------------------------------

#[test]
fn hard_break_table_row_grows_past_its_declared_height() {
    let text = pdf_text("hard_break_table_row_min_height.pptx");
    let compact: String = text.split_whitespace().collect();
    let mut previous_position: Option<usize> = None;

    for number in 1..=9 {
        let line = format!("Line-{number}");
        let position = compact.find(&line).unwrap_or_else(|| {
            panic!("a content-grown PowerPoint table row must retain {line}; got:\n{text}")
        });
        if let Some(previous) = previous_position {
            assert!(
                position > previous,
                "content-grown rows must paint hard-broken lines in source order; got:\n{text}"
            );
        }
        previous_position = Some(position);
    }
}

// ---------------------------------------------------------------------------
// Paragraph mark face (issue #1176)
// ---------------------------------------------------------------------------

/// Parse a business golden-mock deck and return its slides.
fn golden_mock_pages(name: &str) -> Vec<FixedPage> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/golden_mocks/business/sources/pptx")
        .join(name);
    let data = std::fs::read(path).expect("golden mock deck should exist");
    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|page| match page {
            Page::Fixed(fixed) => Some(fixed),
            _ => None,
        })
        .collect()
}

/// The Korean golden mocks set every run in Malgun Gothic and leave every
/// `<a:endParaRPr>` without a typeface, so each paragraph mark falls to the
/// theme's minor Latin font.
///
/// PowerPoint shares one 1.2em line box across every font on the line, the mark
/// included, and that Calibri mark is what seats these titles 0.94573em into
/// the box rather than at Malgun's own 0.98194em share — a whole point higher
/// at every size the four decks use (issue #1176). The face has to survive
/// parsing for the renderer to be able to share the box with it.
#[test]
fn structure_korean_golden_mock_marks_take_the_theme_minor_latin_font() {
    let pages = golden_mock_pages("02_quarterly_review_ko.pptx");
    let mut checked: usize = 0;
    for page in &pages {
        for element in &page.elements {
            let FixedElementKind::TextBox(text_box) = &element.kind else {
                continue;
            };
            for block in &text_box.content {
                let Block::Paragraph(para) = block else {
                    continue;
                };
                if !para
                    .runs
                    .iter()
                    .any(|run| run.style.font_family.as_deref() == Some("Malgun Gothic"))
                {
                    continue;
                }
                assert_eq!(
                    para.style.paragraph_mark_font_family.as_deref(),
                    Some("Calibri"),
                    "a Malgun Gothic paragraph whose mark declares no typeface \
                     must carry the theme's minor Latin font"
                );
                checked += 1;
            }
        }
    }
    assert!(
        checked >= 3,
        "the deck should contribute several Malgun Gothic paragraphs, got {checked}"
    );
}

// ---------------------------------------------------------------------------
// polygon_shadow_offset.pptx — six preset polygons (rect control, a tall
// triangle, a diamond, a right arrow, a five-pointed star and a chevron),
// each filled `C00000` with no outline and a black `a:outerShdw` at
// `blurRad` 254000 EMU (20pt) and `dist` 0, so every shadow's silhouette is
// the fill path itself (issue #1206).
//
// A native macOS PowerPoint 16 export flattens each shadow to its own bitmap
// whose alpha mask is a plain Gaussian blur of the polygon at sigma =
// blurRad/3: sampled against an exact convolution the residual is 0.33-0.53
// alpha levels rms over the whole mask, never past 2.8 of 255. The filtered
// asset must therefore carry the polygon itself as its source, not a box or a
// sequence of scaled copies.
// ---------------------------------------------------------------------------

fn svg_polygon_points(svg: &str) -> Vec<(f64, f64)> {
    let marker = "<polygon points=\"";
    let rest: &str = &svg[svg.find(marker).expect("a polygon source") + marker.len()..];
    let points: &str = &rest[..rest.find('"').expect("a closed points attribute")];
    points
        .split_whitespace()
        .map(|pair| {
            let (x, y) = pair.split_once(',').expect("an SVG coordinate pair");
            (
                x.parse::<f64>().expect("an x coordinate"),
                y.parse::<f64>().expect("a y coordinate"),
            )
        })
        .collect()
}

fn svg_number_attr(svg: &str, name: &str) -> f64 {
    let marker = format!("{name}=\"");
    let rest: &str = &svg[svg.find(&marker).expect("the SVG numeric attribute") + marker.len()..];
    rest[..rest.find('"').expect("a closed numeric attribute")]
        .parse::<f64>()
        .expect("a numeric SVG attribute")
}

/// The tall triangle itself feeds the Gaussian filter at its exact 120x240pt
/// geometry, so the renderer rather than a scaled-ring approximation decides
/// the blur around its slanted edges and sharp apex.
#[test]
fn polygon_shadow_uses_its_exact_outline_as_the_gaussian_source() {
    let data = load_fixture("polygon_shadow_offset.pptx");
    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = generate_typst(&document).unwrap();
    let shadows: Vec<&str> = output
        .images
        .iter()
        .filter_map(|asset| std::str::from_utf8(&asset.data).ok())
        .filter(|svg| svg.contains("<feGaussianBlur"))
        .collect();
    assert_eq!(
        shadows.len(),
        6,
        "one continuous shadow asset per fixture shape"
    );
    let polygon_shadows: Vec<&&str> = shadows
        .iter()
        .filter(|svg| svg.contains("<polygon points=\""))
        .collect();
    assert_eq!(polygon_shadows.len(), 5, "five polygons beside one rect");
    assert!(
        shadows.iter().all(|svg| {
            (svg_number_attr(svg, "stdDeviation") - 20.0 / 3.0).abs() < 1e-9
                && svg.contains("opacity=\"1\"")
        }),
        "every 20pt, fully opaque shadow must carry the declared filter: {shadows:?}"
    );

    let triangle: Vec<(f64, f64)> = polygon_shadows
        .iter()
        .map(|svg| svg_polygon_points(svg))
        .find(|points| {
            let (low, high): (f64, f64) = points
                .iter()
                .fold((f64::MAX, f64::MIN), |(low, high), point| {
                    (low.min(point.1), high.max(point.1))
                });
            (high - low - 240.0).abs() < 0.01
        })
        .expect("the fixture's 120x240pt triangle source");
    assert_eq!(
        triangle.len(),
        3,
        "the triangle should keep its three vertices"
    );
    let (left, right): (f64, f64) = triangle
        .iter()
        .fold((f64::MAX, f64::MIN), |(left, right), point| {
            (left.min(point.0), right.max(point.0))
        });
    let (top, bottom): (f64, f64) = triangle
        .iter()
        .fold((f64::MAX, f64::MIN), |(top, bottom), point| {
            (top.min(point.1), bottom.max(point.1))
        });
    assert!(
        (right - left - 120.0).abs() < 0.01 && (bottom - top - 240.0).abs() < 0.01,
        "the Gaussian source must preserve the 120x240pt frame: {triangle:?}"
    );
    let apex: &(f64, f64) = triangle
        .iter()
        .min_by(|left, right| left.1.total_cmp(&right.1))
        .unwrap();
    assert!(
        (apex.0 - (left + 60.0)).abs() < 0.01 && (apex.1 - top).abs() < 0.01,
        "the apex must remain centred over the full-width base: {triangle:?}"
    );
}

#[test]
fn polygon_shadow_offset_smoke() {
    assert_produces_valid_pdf("polygon_shadow_offset.pptx");
}
