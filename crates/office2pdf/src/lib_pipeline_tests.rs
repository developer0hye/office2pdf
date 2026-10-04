#![cfg(not(target_arch = "wasm32"))] // native-only unit tests (filesystem, system fonts)
use super::test_support::{
    build_docx_with_title, build_test_docx, make_simple_document, make_test_docx_bytes,
};
use super::*;
use crate::ir::*;
use crate::pipeline::{extend_document_fonts, load_additional_fonts};

#[test]
fn test_convert_unsupported_format() {
    let result = convert("test.txt");
    assert!(result.is_err());
    let err = result.unwrap_err();
    assert!(matches!(err, ConvertError::UnsupportedFormat(_)));
}

#[test]
fn test_convert_no_extension() {
    let result = convert("test");
    assert!(result.is_err());
    assert!(matches!(
        result.unwrap_err(),
        ConvertError::UnsupportedFormat(_)
    ));
}

#[test]
fn test_format_detection_all_supported_extensions() {
    assert!(convert_bytes(b"fake", Format::Docx, &ConvertOptions::default()).is_err());
    assert!(convert_bytes(b"fake", Format::Pptx, &ConvertOptions::default()).is_err());
    assert!(convert_bytes(b"fake", Format::Xlsx, &ConvertOptions::default()).is_err());
}

#[test]
fn test_convert_bytes_propagates_parse_error() {
    for format in [Format::Docx, Format::Pptx, Format::Xlsx] {
        let result = convert_bytes(b"fake", format, &ConvertOptions::default());
        assert!(result.is_err());
        assert!(
            matches!(result.unwrap_err(), ConvertError::Parse(_)),
            "Expected Parse error for {format:?}"
        );
    }
}

#[test]
fn test_convert_nonexistent_file_returns_io_error() {
    let result = convert("nonexistent_file.docx");
    assert!(result.is_err());
    assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_should_resolve_font_context_false_for_default_document_without_user_paths() {
    let doc = make_simple_document("Plain text");

    assert!(!should_resolve_font_context(
        &doc,
        &ConvertOptions::default()
    ));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_should_resolve_font_context_true_when_user_font_paths_are_provided() {
    let doc = make_simple_document("Plain text");
    let options = ConvertOptions {
        font_paths: vec![std::env::temp_dir()],
        ..ConvertOptions::default()
    };

    assert!(should_resolve_font_context(&doc, &options));
}

#[test]
fn test_should_resolve_font_context_true_for_in_memory_font_options() {
    let doc = make_simple_document("Plain text");
    let font_options = ConvertOptions {
        font_bytes: vec![vec![0, 1, 2]],
        ..Default::default()
    };
    let fallback_options = ConvertOptions {
        last_resort_font_family: Some("Noto Sans SC".to_string()),
        ..Default::default()
    };

    assert!(should_resolve_font_context(&doc, &font_options));
    assert!(should_resolve_font_context(&doc, &fallback_options));
}

#[test]
fn test_convert_bytes_rejects_invalid_registered_font() {
    let options = ConvertOptions {
        font_bytes: vec![b"not a font".to_vec()],
        ..Default::default()
    };

    let error = convert_bytes(b"not an office file", Format::Docx, &options)
        .expect_err("invalid caller-provided font bytes must not be ignored");
    assert!(
        error
            .to_string()
            .contains("registered font at index 0 contains no usable font faces")
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_should_resolve_font_context_true_when_document_requests_font_family() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Styled text".to_string(),
                    style: TextStyle {
                        font_family: Some("Pretendard".to_string()),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: StyleSheet::default(),
    };

    assert!(should_resolve_font_context(
        &doc,
        &ConvertOptions::default()
    ));
}

#[test]
fn bundled_noto_serif_is_loaded_only_for_the_poster_families() {
    let document = |family: &str| Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Poster".to_string(),
                    style: TextStyle {
                        font_family: Some(family.to_string()),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: StyleSheet::default(),
    };

    let options = ConvertOptions::default();
    let mut unrelated = load_additional_fonts(&options).expect("load default fonts");
    extend_document_fonts(&mut unrelated, &document("Calibri"));
    assert!(
        unrelated
            .iter()
            .all(|font| { font.info().family != crate::bundled_fonts::NOTO_SERIF_FAMILY }),
        "unrelated documents must keep the existing fallback book"
    );

    let mut poster = load_additional_fonts(&options).expect("load default fonts");
    extend_document_fonts(&mut poster, &document("The Hand Black"));
    let poster_faces: Vec<_> = poster
        .iter()
        .filter(|font| font.info().family == crate::bundled_fonts::NOTO_SERIF_FAMILY)
        .collect();
    assert_eq!(poster_faces.len(), 2);

    extend_document_fonts(&mut poster, &document("Avenir Next LT Pro"));
    assert_eq!(
        poster
            .iter()
            .filter(|font| font.info().family == crate::bundled_fonts::NOTO_SERIF_FAMILY)
            .count(),
        2,
        "several poster declarations must not duplicate the bundled faces"
    );
}

#[test]
fn bundled_sans_families_are_loaded_only_for_their_document_requests() {
    let document = |family: &str| Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Footer".to_string(),
                    style: TextStyle {
                        font_family: Some(family.to_string()),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: StyleSheet::default(),
    };

    let options = ConvertOptions::default();
    let mut unrelated = load_additional_fonts(&options).expect("load default fonts");
    extend_document_fonts(&mut unrelated, &document("Calibri"));
    assert!(unrelated.iter().all(|font| {
        !matches!(
            font.info().family.as_str(),
            crate::bundled_fonts::NOTO_SANS_FAMILY | crate::bundled_fonts::SELAWIK_FAMILY
        )
    }));

    let mut aptos = load_additional_fonts(&options).expect("load default fonts");
    extend_document_fonts(&mut aptos, &document("Aptos"));
    assert_eq!(
        aptos
            .iter()
            .filter(|font| font.info().family == crate::bundled_fonts::NOTO_SANS_FAMILY)
            .count(),
        1
    );

    extend_document_fonts(&mut aptos, &document("Aptos"));
    assert_eq!(
        aptos
            .iter()
            .filter(|font| font.info().family == crate::bundled_fonts::NOTO_SANS_FAMILY)
            .count(),
        1,
        "several Aptos declarations must not duplicate the bundled face"
    );

    let mut segoe = load_additional_fonts(&options).expect("load default fonts");
    extend_document_fonts(&mut segoe, &document("Segoe UI"));
    assert_eq!(
        segoe
            .iter()
            .filter(|font| font.info().family == crate::bundled_fonts::SELAWIK_FAMILY)
            .count(),
        2,
        "a Segoe UI request must materialize regular and bold Selawik"
    );
    assert!(
        segoe
            .iter()
            .all(|font| font.info().family != crate::bundled_fonts::NOTO_SANS_FAMILY),
        "Segoe UI must not pull in the metrically different Aptos fallback"
    );
    extend_document_fonts(&mut segoe, &document("Segoe UI"));
    assert_eq!(
        segoe
            .iter()
            .filter(|font| font.info().family == crate::bundled_fonts::SELAWIK_FAMILY)
            .count(),
        2,
        "several Segoe UI declarations must not duplicate bundled faces"
    );
}

#[test]
fn test_convert_with_options_delegates_to_convert_bytes() {
    let result = convert_with_options("nonexistent.docx", &ConvertOptions::default());
    assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
}

#[test]
fn test_convert_delegates_to_convert_with_options() {
    let result = convert("nonexistent.docx");
    assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
}

#[test]
fn test_convert_result_has_pdf_and_warnings() {
    let docx_bytes = build_test_docx();
    let result = convert_bytes(&docx_bytes, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    let _warnings: &Vec<crate::error::ConvertWarning> = &result.warnings;
}

#[test]
fn test_convert_bytes_with_pdfa_option() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("PDF/A test")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let options = ConvertOptions {
        pdf_standard: Some(config::PdfStandard::PdfA2b),
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result.pdf);
    assert!(
        pdf_str.contains("pdfaid") || pdf_str.contains("PDF/A"),
        "PDF/A conversion should include PDF/A metadata"
    );
}

#[test]
fn test_render_document_default_no_pdfa() {
    let doc = make_simple_document("No PDF/A");
    let pdf = render_document(&doc).unwrap();
    let pdf_str = String::from_utf8_lossy(&pdf);
    assert!(
        !pdf_str.contains("pdfaid:conformance"),
        "Default render_document should not produce PDF/A"
    );
}

#[test]
fn test_convert_bytes_with_paper_size_override() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Paper size test")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let options = ConvertOptions {
        paper_size: Some(config::PaperSize::Letter),
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "DOCX with Letter paper override should produce valid PDF"
    );
}

#[test]
fn test_convert_bytes_with_landscape_override() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape override test")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let options = ConvertOptions {
        landscape: Some(true),
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "DOCX with landscape override should produce valid PDF"
    );
}

#[test]
fn test_convert_bytes_returns_populated_metrics() {
    let data = make_test_docx_bytes();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let metrics = result.metrics.expect("convert_bytes should return metrics");
    assert!(
        metrics.parse_duration.as_nanos() > 0,
        "parse_duration should be non-zero"
    );
    assert!(
        metrics.codegen_duration.as_nanos() > 0,
        "codegen_duration should be non-zero"
    );
    assert!(
        metrics.compile_duration.as_nanos() > 0,
        "compile_duration should be non-zero"
    );
    assert!(
        metrics.total_duration.as_nanos() > 0,
        "total_duration should be non-zero"
    );
    assert_eq!(metrics.input_size_bytes, data.len() as u64);
    assert_eq!(metrics.output_size_bytes, result.pdf.len() as u64);
    assert!(metrics.page_count >= 1, "should have at least 1 page");
}

/// `ConvertMetrics::page_count` answers for the produced PDF, not for the IR.
/// This fixture has one flow section however long it runs, so reporting the IR
/// length told every caller "1 page" after Typst broke it across many.
#[test]
#[cfg(feature = "pdf-ops")]
fn test_metrics_page_count_counts_printed_pages() {
    let mut docx = docx_rs::Docx::new();
    for index in 0..400 {
        docx = docx.add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text(format!("Paragraph {index}"))),
        );
    }
    let mut cursor = std::io::Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let metrics = result.metrics.expect("convert_bytes should return metrics");
    let printed = crate::pdf_ops::page_count(&result.pdf).expect("the output PDF should load");

    assert!(printed > 1, "400 paragraphs should not fit one page");
    assert_eq!(
        metrics.page_count, printed,
        "reported page_count should match the pages in the PDF"
    );
}

#[test]
fn continuous_section_fixtures_keep_their_native_page_counts() {
    let fixtures: [(&str, &[u8], u32); 2] = [
        (
            "continuous multi-column transitions",
            include_bytes!(
                "../../../tests/fixtures/docx/libreoffice/tdf46940_dontEquallyDistributeColumns.docx"
            ),
            3,
        ),
        (
            "continuous transition after a table",
            include_bytes!("../../../tests/fixtures/docx/libreoffice/tdf105688.docx"),
            2,
        ),
    ];
    let mut mismatches: Vec<String> = Vec::new();

    for (fixture_name, data, expected_page_count) in fixtures {
        let result = convert_bytes(data, Format::Docx, &ConvertOptions::default())
            .unwrap_or_else(|error| panic!("{fixture_name} should convert: {error}"));
        let metrics = result
            .metrics
            .unwrap_or_else(|| panic!("{fixture_name} should report conversion metrics"));

        if metrics.page_count != expected_page_count {
            mismatches.push(format!(
                "{fixture_name}: expected {expected_page_count} pages, got {}",
                metrics.page_count
            ));
        }
    }

    assert!(mismatches.is_empty(), "{}", mismatches.join("; "));
}

#[test]
fn continuous_section_marker_does_not_add_an_empty_body_line() {
    let (document, _) = crate::parser::Parser::parse(
        &crate::parser::docx::DocxParser,
        include_bytes!(
            "../../../tests/fixtures/docx/libreoffice/tdf46940_dontEquallyDistributeColumns.docx"
        ),
        &ConvertOptions::default(),
    )
    .expect("section-break fixture should parse");
    let Some(crate::ir::Page::Flow(first_section)) = document.pages.first() else {
        panic!("the fixture should begin with its ordinary first section");
    };

    assert!(
        !matches!(
            first_section.content.last(),
            Some(crate::ir::Block::Paragraph(paragraph)) if paragraph.runs.is_empty()
        ),
        "an empty paragraph carrying section properties must not create a blank line"
    );

    for expected_heading in [
        "Continuous break section after a regular page break",
        "New page-break-section. Docx does not balance the last section.",
    ] {
        let paragraph = document
            .pages
            .iter()
            .filter_map(|page| match page {
                Page::Flow(flow) | Page::FlowContinuous(flow) => Some(&flow.content),
                _ => None,
            })
            .flatten()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .find(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .any(|run| run.text.contains(expected_heading))
            })
            .unwrap_or_else(|| panic!("missing shaded heading {expected_heading:?}"));

        assert_eq!(
            paragraph.style.background,
            Some(Color::new(217, 217, 217)),
            "dropping a section marker must keep the later paragraph's shading"
        );
    }
}

#[test]
fn test_metrics_total_ge_sum_of_stages() {
    let data = make_test_docx_bytes();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let metrics = result.metrics.expect("should have metrics");
    let sum = metrics.parse_duration + metrics.codegen_duration + metrics.compile_duration;
    assert!(
        metrics.total_duration >= sum,
        "total ({:?}) should be >= sum of stages ({:?})",
        metrics.total_duration,
        sum
    );
}

#[test]
fn test_metrics_output_size_matches_pdf() {
    let data = make_test_docx_bytes();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let metrics = result.metrics.expect("should have metrics");
    assert_eq!(
        metrics.output_size_bytes,
        result.pdf.len() as u64,
        "output_size_bytes should match actual PDF size"
    );
}

#[test]
fn test_convert_bytes_with_tagged_option() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tagged test")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let options = ConvertOptions {
        tagged: true,
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result.pdf);
    assert!(
        pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
        "Tagged conversion should include structure tree"
    );
}

#[test]
fn test_convert_bytes_with_pdf_ua_option() {
    let data = build_docx_with_title("PDF/UA Test Document");

    let options = ConvertOptions {
        pdf_ua: true,
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result.pdf);
    assert!(
        pdf_str.contains("pdfuaid"),
        "PDF/UA conversion should include pdfuaid metadata"
    );
}

#[test]
fn test_convert_bytes_with_pdf_a_and_pdf_ua_together() {
    // typst 0.14 took a single PDF substandard, so asking for both failed with
    // "Typst currently only supports one PDF substandard at a time"; typst
    // 0.15 targets compatible standards together.
    let data = build_docx_with_title("PDF/A and PDF/UA Test Document");

    let options = ConvertOptions {
        pdf_standard: Some(crate::config::PdfStandard::PdfA2b),
        pdf_ua: true,
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    let pdf_str = String::from_utf8_lossy(&result.pdf);
    assert!(
        pdf_str.contains("pdfaid"),
        "the output should declare PDF/A conformance"
    );
    assert!(
        pdf_str.contains("pdfuaid"),
        "the output should declare PDF/UA conformance"
    );
}

#[test]
fn test_convert_bytes_tagged_pdf_with_heading() {
    use std::io::Cursor;

    let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);

    let docx = docx_rs::Docx::new()
        .add_style(h1_style)
        .add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("My Title"))
                .style("Heading1"),
        )
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
        );

    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let options = ConvertOptions {
        tagged: true,
        ..Default::default()
    };
    let result = convert_bytes(&data, Format::Docx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result.pdf);
    assert!(
        pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
        "Tagged PDF with headings should contain structure tags"
    );
}
