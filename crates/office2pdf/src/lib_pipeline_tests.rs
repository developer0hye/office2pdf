use super::test_support::{
    build_docx_with_title, build_test_docx, make_simple_document, make_test_docx_bytes,
};
use super::*;
use crate::ir::*;

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

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_should_resolve_font_context_true_when_document_requests_font_family() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
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
                }],
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    assert!(should_resolve_font_context(
        &doc,
        &ConvertOptions::default()
    ));
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
