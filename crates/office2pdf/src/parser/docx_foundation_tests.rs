use super::*;

// ----- Basic parsing tests -----

#[test]
fn test_parse_empty_docx() {
    let data = build_docx_bytes(vec![]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.pages.len(), 1);
    match &doc.pages[0] {
        Page::Flow(page) => {
            assert!(page.content.is_empty());
        }
        _ => panic!("Expected FlowPage"),
    }
}

#[test]
fn test_parse_single_paragraph() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello, world!")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert_eq!(page.content.len(), 1);
    match &page.content[0] {
        Block::Paragraph(para) => {
            assert_eq!(para.runs.len(), 1);
            assert_eq!(para.runs[0].text, "Hello, world!");
        }
        _ => panic!("Expected Paragraph block"),
    }
}

#[test]
fn test_parse_multiple_paragraphs() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First paragraph")),
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second paragraph")),
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Third paragraph")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert_eq!(page.content.len(), 3);

    let texts: Vec<&str> = page
        .content
        .iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => paragraph.runs[0].text.as_str(),
            _ => panic!("Expected Paragraph"),
        })
        .collect();
    assert_eq!(
        texts,
        vec!["First paragraph", "Second paragraph", "Third paragraph"]
    );
}

#[test]
fn test_parse_paragraph_with_multiple_runs() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Hello, "))
            .add_run(docx_rs::Run::new().add_text("beautiful "))
            .add_run(docx_rs::Run::new().add_text("world!")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 3);
    assert_eq!(para.runs[0].text, "Hello, ");
    assert_eq!(para.runs[1].text, "beautiful ");
    assert_eq!(para.runs[2].text, "world!");
}

#[test]
fn test_parse_empty_paragraph() {
    let data = build_docx_bytes(vec![docx_rs::Paragraph::new()]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert_eq!(page.content.len(), 1);
    match &page.content[0] {
        Block::Paragraph(para) => {
            assert!(para.runs.is_empty());
        }
        _ => panic!("Expected Paragraph block"),
    }
}

// ----- Page setup tests -----

#[test]
fn test_default_page_size_is_used() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(page.size.width > 0.0);
    assert!(page.size.height > 0.0);
}

#[test]
fn test_custom_page_size_extracted() {
    let width_twips: u32 = 8392;
    let height_twips: u32 = 11907;
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
        width_twips,
        height_twips,
        1440,
        1440,
        1440,
        1440,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let expected_width = width_twips as f64 / 20.0;
    let expected_height = height_twips as f64 / 20.0;
    assert!(
        (page.size.width - expected_width).abs() < 1.0,
        "Expected width ~{expected_width}, got {}",
        page.size.width
    );
    assert!(
        (page.size.height - expected_height).abs() < 1.0,
        "Expected height ~{expected_height}, got {}",
        page.size.height
    );
}

#[test]
fn test_custom_margins_extracted() {
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
        12240,
        15840,
        720,
        720,
        720,
        720,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let expected_margin = 720.0 / 20.0;
    assert!(
        (page.margins.top - expected_margin).abs() < 1.0,
        "Expected top margin ~{expected_margin}, got {}",
        page.margins.top
    );
    assert!((page.margins.bottom - expected_margin).abs() < 1.0);
    assert!((page.margins.left - expected_margin).abs() < 1.0);
    assert!((page.margins.right - expected_margin).abs() < 1.0);
}

// ----- Error handling tests -----

#[test]
fn test_parse_invalid_data_returns_error() {
    let parser = DocxParser;
    let result = parser.parse(b"not a valid docx file", &ConvertOptions::default());
    assert!(result.is_err());
    match result.unwrap_err() {
        ConvertError::Parse(_) => {}
        other => panic!("Expected Parse error, got: {other:?}"),
    }
}

#[test]
fn test_parse_error_includes_library_name() {
    let parser = DocxParser;
    let result = parser.parse(b"not a valid docx file", &ConvertOptions::default());
    let err = result.unwrap_err();
    let msg = err.to_string();
    assert!(
        msg.contains("docx-rs"),
        "Parse error should include upstream library name 'docx-rs', got: {msg}"
    );
}

// ----- Text style defaults -----

#[test]
fn test_parsed_runs_have_default_text_style() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert!(run.style.bold.is_none() || run.style.bold == Some(false));
    assert!(run.style.italic.is_none() || run.style.italic == Some(false));
    assert!(run.style.underline.is_none() || run.style.underline == Some(false));
}

#[test]
fn test_parsed_paragraphs_have_default_style() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert!(para.style.alignment.is_none());
}

// ----- Inline formatting tests (US-004) -----

#[test]
fn test_bold_formatting_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bold text").bold()),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_italic_formatting_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Italic text").italic()),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.italic, Some(true));
}

#[test]
fn test_underline_formatting_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Underlined text")
                .underline("single"),
        ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.underline, Some(true));
}

#[test]
fn test_strikethrough_formatting_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Struck text").strike()),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.strikethrough, Some(true));
}

#[test]
fn test_font_size_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Sized text").size(24)),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.font_size, Some(12.0));
}

#[test]
fn test_letter_spacing_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Tracked text")
                .character_spacing(40),
        ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.letter_spacing, Some(2.0));
}

#[test]
fn test_font_color_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Red text").color("FF0000")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_font_family_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Arial text")
                .fonts(docx_rs::RunFonts::new().ascii("Arial")),
        ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.font_family, Some("Arial".to_string()));
}

#[test]
fn test_combined_formatting_extracted() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Styled text")
                .bold()
                .italic()
                .underline("single")
                .strike()
                .size(28)
                .color("0000FF")
                .fonts(docx_rs::RunFonts::new().ascii("Courier")),
        ),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.italic, Some(true));
    assert_eq!(run.style.underline, Some(true));
    assert_eq!(run.style.strikethrough, Some(true));
    assert_eq!(run.style.font_size, Some(14.0));
    assert_eq!(run.style.color, Some(Color::new(0, 0, 255)));
    assert_eq!(run.style.font_family, Some("Courier".to_string()));
}

#[test]
fn test_plain_text_has_no_formatting() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);
    assert!(run.style.bold.is_none());
    assert!(run.style.italic.is_none());
    assert!(run.style.underline.is_none());
    assert!(run.style.strikethrough.is_none());
    assert!(run.style.font_size.is_none());
    assert!(run.style.letter_spacing.is_none());
    assert!(run.style.color.is_none());
    assert!(run.style.font_family.is_none());
}

// ----- Paragraph formatting tests (US-005) -----

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

#[test]
fn test_multiple_runs_with_different_formatting() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Bold ").bold())
            .add_run(docx_rs::Run::new().add_text("Italic ").italic())
            .add_run(docx_rs::Run::new().add_text("Plain")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 3);
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert!(para.runs[0].style.italic.is_none());
    assert!(para.runs[1].style.bold.is_none());
    assert_eq!(para.runs[1].style.italic, Some(true));
    assert!(para.runs[2].style.bold.is_none());
    assert!(para.runs[2].style.italic.is_none());
}
