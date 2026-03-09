use super::*;

#[test]
fn test_paragraph_alignment_center() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Centered"))
            .align(docx_rs::AlignmentType::Center),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_paragraph_alignment_right() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Right"))
            .align(docx_rs::AlignmentType::Right),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.alignment, Some(Alignment::Right));
}

#[test]
fn test_paragraph_alignment_left() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Left"))
            .align(docx_rs::AlignmentType::Left),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.alignment, Some(Alignment::Left));
}

#[test]
fn test_paragraph_alignment_justify() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Justified"))
            .align(docx_rs::AlignmentType::Both),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.alignment, Some(Alignment::Justify));
}

#[test]
fn test_paragraph_indent_left() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Indented"))
            .indent(Some(720), None, None, None),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.indent_left, Some(36.0));
}

#[test]
fn test_paragraph_indent_right() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Indented"))
            .indent(None, None, Some(360), None),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.indent_right, Some(18.0));
}

#[test]
fn test_paragraph_indent_first_line() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("First line indented"))
            .indent(
                None,
                Some(docx_rs::SpecialIndentType::FirstLine(480)),
                None,
                None,
            ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.indent_first_line, Some(24.0));
}

#[test]
fn test_paragraph_indent_hanging() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Hanging indent"))
            .indent(
                Some(720),
                Some(docx_rs::SpecialIndentType::Hanging(360)),
                None,
                None,
            ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.indent_left, Some(36.0));
    assert_eq!(para.style.indent_first_line, Some(-18.0));
}

#[test]
fn test_paragraph_line_spacing_auto() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Double spaced"))
            .line_spacing(
                docx_rs::LineSpacing::new()
                    .line_rule(docx_rs::LineSpacingType::Auto)
                    .line(480),
            ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    match para.style.line_spacing {
        Some(LineSpacing::Proportional(factor)) => {
            assert!(
                (factor - 2.0).abs() < 0.01,
                "Expected 2.0 (double spacing), got {factor}"
            );
        }
        other => panic!("Expected Proportional line spacing, got {other:?}"),
    }
}

#[test]
fn test_paragraph_line_spacing_exact() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Exact spaced"))
            .line_spacing(
                docx_rs::LineSpacing::new()
                    .line_rule(docx_rs::LineSpacingType::Exact)
                    .line(240),
            ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    match para.style.line_spacing {
        Some(LineSpacing::Exact(pts)) => {
            assert!((pts - 12.0).abs() < 0.01, "Expected 12pt, got {pts}");
        }
        other => panic!("Expected Exact line spacing, got {other:?}"),
    }
}

#[test]
fn test_paragraph_space_before_after() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Spaced paragraph"))
            .line_spacing(docx_rs::LineSpacing::new().before(240).after(120)),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.space_before, Some(12.0));
    assert_eq!(para.style.space_after, Some(6.0));
}

#[test]
fn test_paragraph_page_break_before() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before break")),
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("After break"))
            .page_break_before(true),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let blocks = all_blocks(&doc);
    assert_eq!(blocks.len(), 3, "Expected 3 blocks, got {}", blocks.len());
    assert!(matches!(&blocks[0], Block::Paragraph(_)));
    assert!(matches!(&blocks[1], Block::PageBreak));
    assert!(matches!(&blocks[2], Block::Paragraph(_)));
}

#[test]
fn test_paragraph_combined_formatting() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Styled paragraph"))
            .align(docx_rs::AlignmentType::Center)
            .indent(
                Some(720),
                Some(docx_rs::SpecialIndentType::FirstLine(360)),
                None,
                None,
            )
            .line_spacing(
                docx_rs::LineSpacing::new()
                    .line_rule(docx_rs::LineSpacingType::Auto)
                    .line(360)
                    .before(120)
                    .after(60),
            ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(para.style.alignment, Some(Alignment::Center));
    assert_eq!(para.style.indent_left, Some(36.0));
    assert_eq!(para.style.indent_first_line, Some(18.0));
    assert_eq!(para.style.space_before, Some(6.0));
    assert_eq!(para.style.space_after, Some(3.0));
    match para.style.line_spacing {
        Some(LineSpacing::Proportional(factor)) => {
            assert!(
                (factor - 1.5).abs() < 0.01,
                "Expected 1.5 spacing, got {factor}"
            );
        }
        other => panic!("Expected Proportional line spacing, got {other:?}"),
    }
}
