use super::*;
use crate::ir::*;
use std::io::Cursor;

/// Helper: build a minimal DOCX as bytes using docx-rs builder.
fn build_docx_bytes(paragraphs: Vec<docx_rs::Paragraph>) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new();
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with custom page size and margins.
fn build_docx_bytes_with_page_setup(
    paragraphs: Vec<docx_rs::Paragraph>,
    width_twips: u32,
    height_twips: u32,
    margin_top: i32,
    margin_bottom: i32,
    margin_left: i32,
    margin_right: i32,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new()
        .page_size(width_twips, height_twips)
        .page_margin(
            docx_rs::PageMargin::new()
                .top(margin_top)
                .bottom(margin_bottom)
                .left(margin_left)
                .right(margin_right),
        );
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

// ----- Basic parsing tests -----

#[test]
fn test_parse_empty_docx() {
    let data = build_docx_bytes(vec![]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    // An empty DOCX should produce a document with one FlowPage and no content blocks
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
        .map(|b| match b {
            Block::Paragraph(p) => p.runs[0].text.as_str(),
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
    // An empty paragraph should still be present (it may have no runs)
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
    // docx-rs default: 11906 x 16838 twips (A4)
    // = 595.3 x 841.9 pt
    assert!(page.size.width > 0.0);
    assert!(page.size.height > 0.0);
}

#[test]
fn test_custom_page_size_extracted() {
    // A5 page: 148mm x 210mm
    // In twips: 8392 x 11907 (1 pt = 20 twips)
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
    // Margins: 0.5 inch = 720 twips = 36pt
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
    let expected_margin = 720.0 / 20.0; // 36pt
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
    // Plain text should have default style (all None)
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
    // Default paragraph style should have no explicit alignment
    assert!(para.style.alignment.is_none());
}

// ----- Inline formatting tests (US-004) -----

/// Helper: extract the first run from the first paragraph of a parsed document.
fn first_run(doc: &Document) -> &Run {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    &para.runs[0]
}

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
    // docx-rs size is in half-points: 24 half-points = 12pt
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
    // docx-rs character spacing is in twips: 40 twips = 2pt
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
                .size(28) // 14pt
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

/// Helper: extract the first paragraph from a parsed document.
fn first_paragraph(doc: &Document) -> &Paragraph {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph block"),
    }
}

/// Helper: get all blocks from the first page.
fn all_blocks(doc: &Document) -> &[Block] {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    &page.content
}

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
    // 720 twips = 36pt
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
    // 360 twips = 18pt
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
    // first line indent: 480 twips = 24pt
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
    // hanging indent: 360 twips = 18pt (negative first-line indent)
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
    // Auto line spacing: line=480 means 480/240 = 2.0 (double spacing)
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
    // Exact line spacing: line=240 twips = 12pt
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
    // before=240 twips = 12pt, after=120 twips = 6pt
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
    // Should have: Paragraph("Before break"), PageBreak, Paragraph("After break")
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

// ----- Table parsing tests (US-007) -----

/// Helper: build a DOCX with a table using docx-rs builder.
fn build_docx_with_table(table: docx_rs::Table) -> Vec<u8> {
    let docx = docx_rs::Docx::new().add_table(table);
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: extract the first table block from a parsed document.
fn first_table(doc: &Document) -> &crate::ir::Table {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    for block in &page.content {
        if let Block::Table(t) = block {
            return t;
        }
    }
    panic!("No Table block found");
}

#[test]
fn test_table_simple_2x2() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A1")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 3000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 2);
    assert_eq!(t.rows[0].cells.len(), 2);
    assert_eq!(t.rows[1].cells.len(), 2);

    // Check cell content
    let cell_text = |row: usize, col: usize| -> String {
        t.rows[row].cells[col]
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect::<String>()
    };
    assert_eq!(cell_text(0, 0), "A1");
    assert_eq!(cell_text(0, 1), "B1");
    assert_eq!(cell_text(1, 0), "A2");
    assert_eq!(cell_text(1, 1), "B2");
}

#[test]
fn test_table_column_widths_from_grid() {
    // Grid widths in twips: 2000, 3000 → 100pt, 150pt
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A"))),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B"))),
    ])])
    .set_grid(vec![2000, 3000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 150.0).abs() < 0.1,
        "Expected 150pt, got {}",
        t.column_widths[1]
    );
}

#[test]
fn test_table_tiny_autofit_grid_is_ignored() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Left column text")),
        ),
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Right column text")),
        ),
    ])])
    .set_grid(vec![100, 100]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert!(
        t.column_widths.is_empty(),
        "Expected tiny autofit grid widths to be ignored, got {:?}",
        t.column_widths
    );
}

#[test]
fn test_table_tiny_non_auto_grid_is_normalized_to_table_width() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A"))),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B"))),
    ])])
    .set_grid(vec![100, 100])
    .width(400, docx_rs::WidthType::Dxa);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!((t.column_widths[0] - 10.0).abs() < 0.1);
    assert!((t.column_widths[1] - 10.0).abs() < 0.1);
}

#[test]
fn test_table_single_column_placeholder_grid_uses_table_width() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Long single cell")),
        ),
    ])])
    .set_grid(vec![100])
    .width(9360, docx_rs::WidthType::Dxa);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 1);
    assert!(
        (t.column_widths[0] - 468.0).abs() < 0.2,
        "Expected table width normalization to 468pt, got {}",
        t.column_widths[0]
    );
}

#[test]
fn test_table_column_widths_from_cell_widths_without_grid() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A")))
            .width(2000, docx_rs::WidthType::Dxa),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B")))
            .width(3000, docx_rs::WidthType::Dxa),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 150.0).abs() < 0.1,
        "Expected 150pt, got {}",
        t.column_widths[1]
    );
}

#[test]
fn test_table_column_widths_from_spanned_cell_widths_without_grid() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
            )
            .grid_span(2)
            .width(4000, docx_rs::WidthType::Dxa),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C")))
            .width(2000, docx_rs::WidthType::Dxa),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 3);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected first merged column to be 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 100.0).abs() < 0.1,
        "Expected second merged column to be 100pt, got {}",
        t.column_widths[1]
    );
    assert!(
        (t.column_widths[2] - 100.0).abs() < 0.1,
        "Expected final column to be 100pt, got {}",
        t.column_widths[2]
    );
}

#[test]
fn test_scan_table_headers_counts_only_leading_rows() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>D1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Ignored</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Only body</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
</w:document>"#;

    let headers = scan_table_headers(document_xml);

    assert_eq!(headers.len(), 2);
    assert_eq!(headers[0].repeat_rows, 2);
    assert_eq!(headers[1].repeat_rows, 0);
}

#[test]
fn test_table_header_rows_from_raw_docx_xml() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tblPr/>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Body A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Body B</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.header_row_count, 1);
}

#[test]
fn test_table_default_cell_margins_from_table_property() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell"))),
    ])])
    .margins(docx_rs::TableCellMargins::new().margin(40, 60, 20, 80));

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(
        t.default_cell_padding,
        Some(Insets {
            top: 2.0,
            right: 3.0,
            bottom: 1.0,
            left: 4.0,
        })
    );
    assert!(t.rows[0].cells[0].padding.is_none());
}

#[test]
fn test_table_cell_margins_override_table_defaults() {
    let mut cell = docx_rs::TableCell::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")));
    cell.property = docx_rs::TableCellProperty::new()
        .margin_top(100, docx_rs::WidthType::Dxa)
        .margin_left(120, docx_rs::WidthType::Dxa);

    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![cell])])
        .margins(docx_rs::TableCellMargins::new().margin(20, 40, 60, 80));

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(
        t.default_cell_padding,
        Some(Insets {
            top: 1.0,
            right: 2.0,
            bottom: 3.0,
            left: 4.0,
        })
    );
    assert_eq!(
        t.rows[0].cells[0].padding,
        Some(Insets {
            top: 5.0,
            right: 2.0,
            bottom: 3.0,
            left: 6.0,
        })
    );
}

#[test]
fn test_table_alignment_from_table_property() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
        ),
    ])])
    .align(docx_rs::TableAlignmentType::Center);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.alignment, Some(Alignment::Center));
}

#[test]
fn test_table_cell_with_formatted_text() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bold").bold())
                .add_run(docx_rs::Run::new().add_text(" and italic").italic()),
        ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let para = match &cell.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph in cell"),
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Bold");
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert_eq!(para.runs[1].text, " and italic");
    assert_eq!(para.runs[1].style.italic, Some(true));
}

#[test]
fn test_table_colspan_via_grid_span() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
                )
                .grid_span(2),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    // First row: one merged cell with colspan=2
    assert_eq!(t.rows[0].cells.len(), 1);
    assert_eq!(t.rows[0].cells[0].col_span, 2);

    // Second row: two normal cells
    assert_eq!(t.rows[1].cells.len(), 2);
    assert_eq!(t.rows[1].cells[0].col_span, 1);
}

#[test]
fn test_table_rowspan_via_vmerge() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tall")),
                )
                .vertical_merge(docx_rs::VMergeType::Restart),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 3);

    // First row: the restart cell should have rowspan=3
    let tall_cell = &t.rows[0].cells[0];
    assert_eq!(tall_cell.row_span, 3);

    // Second and third rows: continue cells should be skipped
    // so rows[1] and rows[2] should have only 1 cell each (B2, B3)
    assert_eq!(t.rows[1].cells.len(), 1);
    assert_eq!(t.rows[2].cells.len(), 1);
}

#[test]
fn test_table_exact_row_height_and_cell_vertical_align() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
                )
                .vertical_align(docx_rs::VAlignType::Center),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Peer")),
            ),
        ])
        .row_height(36.0)
        .height_rule(docx_rs::HeightRule::Exact),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows[0].height, Some(36.0));
    assert_eq!(
        t.rows[0].cells[0].vertical_align,
        Some(CellVerticalAlign::Center)
    );
}

#[test]
fn test_table_cell_background_color() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Red bg")),
            )
            .shading(docx_rs::Shading::new().fill("FF0000")),
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("No bg")),
        ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows[0].cells[0].background, Some(Color::new(255, 0, 0)));
    assert!(t.rows[0].cells[1].background.is_none());
}

#[test]
fn test_table_cell_borders() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bordered")),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("FF0000"),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                    .size(8)
                    .color("0000FF"),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");

    // Top: size=16 eighths → 2pt, color=FF0000
    let top = border.top.as_ref().expect("Expected top border");
    assert!(
        (top.width - 2.0).abs() < 0.01,
        "Expected 2pt, got {}",
        top.width
    );
    assert_eq!(top.color, Color::new(255, 0, 0));

    // Bottom: size=8 eighths → 1pt, color=0000FF
    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert!(
        (bottom.width - 1.0).abs() < 0.01,
        "Expected 1pt, got {}",
        bottom.width
    );
    assert_eq!(bottom.color, Color::new(0, 0, 255));
}

#[test]
fn test_table_cell_border_styles() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Styled borders")),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("000000")
                    .border_type(docx_rs::BorderType::Dashed),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                    .size(8)
                    .color("0000FF")
                    .border_type(docx_rs::BorderType::Dotted),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Left)
                    .size(12)
                    .color("FF0000")
                    .border_type(docx_rs::BorderType::DotDash),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Right)
                    .size(16)
                    .color("00FF00")
                    .border_type(docx_rs::BorderType::Double),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");

    // Top: dashed
    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Dashed, "Top should be dashed");

    // Bottom: dotted
    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert_eq!(
        bottom.style,
        BorderLineStyle::Dotted,
        "Bottom should be dotted"
    );

    // Left: dashDot
    let left = border.left.as_ref().expect("Expected left border");
    assert_eq!(
        left.style,
        BorderLineStyle::DashDot,
        "Left should be dashDot"
    );

    // Right: double
    let right = border.right.as_ref().expect("Expected right border");
    assert_eq!(
        right.style,
        BorderLineStyle::Double,
        "Right should be double"
    );
}

#[test]
fn test_table_cell_solid_border_default_style() {
    // Single (default) border type should map to Solid
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Solid")))
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("000000"),
                // Default border_type is Single → should map to Solid
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);
    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");
    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Solid, "Single → Solid");
}

#[test]
fn test_table_cell_with_multiple_paragraphs() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 1")),
            )
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 2")),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let paras: Vec<&str> = cell
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) if !p.runs.is_empty() => Some(p.runs[0].text.as_str()),
            _ => None,
        })
        .collect();
    assert!(paras.contains(&"Para 1"), "Expected 'Para 1' in cell");
    assert!(paras.contains(&"Para 2"), "Expected 'Para 2' in cell");
}

#[test]
fn test_table_with_paragraph_before_and_after() {
    let data = {
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before")),
            )
            .add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")),
                ),
            ])]))
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After")),
            );
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    };

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let blocks = all_blocks(&doc);

    // Should have: Paragraph("Before"), Table, Paragraph("After")
    assert!(
        blocks.len() >= 3,
        "Expected at least 3 blocks, got {}",
        blocks.len()
    );
    assert!(matches!(&blocks[0], Block::Paragraph(_)));
    let has_table = blocks.iter().any(|b| matches!(b, Block::Table(_)));
    assert!(has_table, "Expected a Table block");
}

#[test]
fn test_table_colspan_and_rowspan_combined() {
    // 3x3 table with top-left 2x2 merged
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Big")),
                )
                .grid_span(2)
                .vertical_merge(docx_rs::VMergeType::Restart),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .grid_span(2)
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C2")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A3")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C3")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    // First row: "Big" (colspan=2, rowspan=2) + "C1"
    let big_cell = &t.rows[0].cells[0];
    assert_eq!(big_cell.col_span, 2, "Expected colspan=2");
    assert_eq!(big_cell.row_span, 2, "Expected rowspan=2");

    // Second row: continue cell skipped, so only "C2"
    assert_eq!(t.rows[1].cells.len(), 1);

    // Third row: three normal cells
    assert_eq!(t.rows[2].cells.len(), 3);
}

#[test]
fn test_table_empty_cells() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
        docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 1);
    assert_eq!(t.rows[0].cells.len(), 2);
    // Empty cells should still have content (possibly empty paragraphs)
    for cell in &t.rows[0].cells {
        assert_eq!(cell.col_span, 1);
        assert_eq!(cell.row_span, 1);
    }
}

// ── Image extraction tests ──────────────────────────────────────────

/// Build a minimal valid 1×1 red pixel BMP image.
/// BMP is trivially decodable by the `image` crate (no compression).
fn make_test_bmp() -> Vec<u8> {
    let mut bmp = Vec::new();
    // BMP file header (14 bytes)
    bmp.extend_from_slice(b"BM"); // magic
    bmp.extend_from_slice(&58u32.to_le_bytes()); // file size
    bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
    bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset

    // BITMAPINFOHEADER (40 bytes)
    bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
    bmp.extend_from_slice(&1i32.to_le_bytes()); // width
    bmp.extend_from_slice(&1i32.to_le_bytes()); // height
    bmp.extend_from_slice(&1u16.to_le_bytes()); // color planes
    bmp.extend_from_slice(&24u16.to_le_bytes()); // bits per pixel (RGB)
    bmp.extend_from_slice(&0u32.to_le_bytes()); // compression (none)
    bmp.extend_from_slice(&4u32.to_le_bytes()); // image size (row aligned to 4 bytes)
    bmp.extend_from_slice(&0u32.to_le_bytes()); // x pixels/meter
    bmp.extend_from_slice(&0u32.to_le_bytes()); // y pixels/meter
    bmp.extend_from_slice(&0u32.to_le_bytes()); // total colors
    bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors

    // Pixel data: 1 pixel BGR (red = 00 00 FF) + 1 byte padding
    bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);

    bmp
}

/// Build a DOCX containing an inline image using Pic::new() which processes
/// the image through the `image` crate, ensuring valid PNG output in the ZIP.
fn build_docx_with_image(width_px: u32, height_px: u32) -> Vec<u8> {
    let bmp_data = make_test_bmp();
    // Use Pic::new() which decodes the BMP and re-encodes as PNG internally
    let pic = docx_rs::Pic::new(&bmp_data).size(width_px * 9525, height_px * 9525);
    let para_with_image = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic));
    let docx = docx_rs::Docx::new().add_paragraph(para_with_image);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Build a minimal DOCX with a custom `document.xml` and one image relationship.
fn build_docx_with_custom_image_document(document_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bmp"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.start_file("word/media/image1.bmp", options).unwrap();
    std::io::Write::write_all(&mut zip, &make_test_bmp()).unwrap();

    zip.finish().unwrap().into_inner()
}

/// Helper: find all Image blocks in a FlowPage.
fn find_images(doc: &Document) -> Vec<&ImageData> {
    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|b| match b {
            Block::Image(img) => Some(img),
            _ => None,
        })
        .collect()
}

#[test]
fn test_docx_image_inline_basic() {
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(images.len(), 1, "Expected exactly one image block");
    assert!(!images[0].data.is_empty(), "Image data should not be empty");
}

#[test]
fn test_docx_image_format_is_png() {
    let data = build_docx_with_image(50, 50);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(
        images[0].format,
        ImageFormat::Png,
        "Image format should be PNG"
    );
}

#[test]
fn test_docx_vml_shape_image_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="VMLImage1" style="width:72pt;height:36pt">
                        <v:imagedata r:id="rIdImage1"/>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_custom_image_document(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(images.len(), 1, "Expected one VML image");
    assert_eq!(images[0].format, ImageFormat::Png);
    assert_eq!(images[0].width, Some(72.0));
    assert_eq!(images[0].height, Some(36.0));
}

#[test]
fn test_docx_image_dimensions() {
    // 100px × 80px → EMU: 100*9525=952500, 80*9525=762000
    // EMU to points: 952500/12700=75.0, 762000/12700=60.0
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    let img = images[0];

    let width = img.width.expect("Expected width");
    let height = img.height.expect("Expected height");

    assert!(
        (width - 75.0).abs() < 0.1,
        "Expected width ~75pt, got {width}"
    );
    assert!(
        (height - 60.0).abs() < 0.1,
        "Expected height ~60pt, got {height}"
    );
}

#[test]
fn test_docx_image_with_text_paragraphs() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data);
    let docx = docx_rs::Docx::new()
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before image")),
        )
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)))
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After image")),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };

    let has_image = page.content.iter().any(|b| matches!(b, Block::Image(_)));
    assert!(has_image, "Expected an image block in the content");

    let has_before = page.content.iter().any(|b| match b {
        Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("Before")),
        _ => false,
    });
    assert!(has_before, "Expected 'Before image' text");

    let has_after = page.content.iter().any(|b| match b {
        Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("After")),
        _ => false,
    });
    assert!(has_after, "Expected 'After image' text");
}

#[test]
fn test_docx_multiple_images() {
    let bmp_data = make_test_bmp();
    let pic1 = docx_rs::Pic::new(&bmp_data).size(100 * 9525, 100 * 9525);
    let pic2 = docx_rs::Pic::new(&bmp_data).size(200 * 9525, 150 * 9525);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic1)))
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic2)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert!(
        images.len() >= 2,
        "Expected at least 2 images, got {}",
        images.len()
    );
}

#[test]
fn test_docx_image_data_contains_png_header() {
    let data = build_docx_with_image(50, 50);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    let img_data = &images[0].data;

    // docx-rs converts all images to PNG; verify PNG magic bytes
    assert!(
        img_data.len() >= 8 && img_data[0..4] == [0x89, 0x50, 0x4E, 0x47],
        "Image data should start with PNG magic bytes"
    );
}

#[test]
fn test_docx_no_images_produces_no_image_blocks() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };

    let image_count = page
        .content
        .iter()
        .filter(|b| matches!(b, Block::Image(_)))
        .count();
    assert_eq!(image_count, 0, "Expected no image blocks");
}

#[test]
fn test_docx_image_with_custom_emu_size() {
    // Create image with specific EMU size via .size() override
    // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data).size(2_540_000, 1_270_000);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert!(!images.is_empty(), "Expected at least one image");
    let img = images[0];

    let width = img.width.expect("Expected width");
    let height = img.height.expect("Expected height");

    assert!(
        (width - 200.0).abs() < 0.1,
        "Expected width ~200pt, got {width}"
    );
    assert!(
        (height - 100.0).abs() < 0.1,
        "Expected height ~100pt, got {height}"
    );
}

// ----- List parsing tests -----

/// Helper: build a DOCX with numbering definitions and list paragraphs.
fn build_docx_with_numbering(
    abstract_nums: Vec<docx_rs::AbstractNumbering>,
    numberings: Vec<docx_rs::Numbering>,
    paragraphs: Vec<docx_rs::Paragraph>,
) -> Vec<u8> {
    let mut nums = docx_rs::Numberings::new();
    for an in abstract_nums {
        nums = nums.add_abstract_numbering(an);
    }
    for n in numberings {
        nums = nums.add_numbering(n);
    }

    let mut docx = docx_rs::Docx::new().numberings(nums);
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_parse_simple_bulleted_list() {
    // Create a bullet list: abstractNum with format "bullet", numId=1, ilvl=0
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("bullet"),
        docx_rs::LevelText::new("•"),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item A"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item B"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item C"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    // Should produce a single List block with 3 items
    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].kind, ListKind::Unordered);
    assert_eq!(lists[0].items.len(), 3);
    assert_eq!(lists[0].items[0].level, 0);
    assert_eq!(
        lists[0].level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Unordered,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );

    // Verify item content
    let text0: String = lists[0].items[0]
        .content
        .iter()
        .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
        .collect();
    assert_eq!(text0, "Item A");
}

#[test]
fn test_parse_simple_numbered_list() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Second"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].kind, ListKind::Ordered);
    assert_eq!(lists[0].items.len(), 2);
    assert_eq!(lists[0].items[0].start_at, Some(1));
    assert_eq!(
        lists[0].level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Ordered,
            numbering_pattern: Some("1.".to_string()),
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_nested_multi_level_list() {
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("◦"),
            docx_rs::LevelJc::new("left"),
        ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Top level"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Nested item"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Back to top"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].items.len(), 3);
    assert_eq!(lists[0].items[0].level, 0);
    assert_eq!(lists[0].items[1].level, 1);
    assert_eq!(lists[0].items[2].level, 0);
    assert_eq!(
        lists[0].level_styles.get(&1),
        Some(&ListLevelStyle {
            kind: ListKind::Unordered,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_numbered_list_start_override() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering =
        docx_rs::Numbering::new(1, 0).add_override(docx_rs::LevelOverride::new(0).start(3));

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Third"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Fourth"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("Expected list block");

    assert_eq!(list.items[0].start_at, Some(3));
    assert_eq!(list.items[1].start_at, None);
    assert_eq!(
        list.level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Ordered,
            numbering_pattern: Some("1.".to_string()),
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_mixed_ordered_and_bulleted_levels() {
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Step"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bullet child"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("Expected list block");

    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(
        list.level_styles,
        std::collections::BTreeMap::from([
            (
                0,
                ListLevelStyle {
                    kind: ListKind::Ordered,
                    numbering_pattern: Some("1.".to_string()),
                    full_numbering: false,
                    marker_text: None,
                    marker_style: None,
                },
            ),
            (
                1,
                ListLevelStyle {
                    kind: ListKind::Unordered,
                    numbering_pattern: None,
                    full_numbering: false,
                    marker_text: None,
                    marker_style: None,
                },
            ),
        ])
    );
}

#[test]
fn test_parse_mixed_levels_with_different_num_ids_preserves_parent_sequence() {
    let ordered_abstract = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let bullet_abstract = docx_rs::AbstractNumbering::new(1)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ));

    let data = build_docx_with_numbering(
        vec![ordered_abstract, bullet_abstract],
        vec![
            docx_rs::Numbering::new(1, 1),
            docx_rs::Numbering::new(2, 0),
        ],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Parent 1"))
                .numbering(docx_rs::NumberingId::new(2), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Child A"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Child B"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Parent 2"))
                .numbering(docx_rs::NumberingId::new(2), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(
        lists.len(),
        1,
        "Expected one merged mixed list when nested level uses a different numId"
    );

    let list = lists[0];
    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(list.items.len(), 4);
    assert_eq!(list.items[0].level, 0);
    assert_eq!(list.items[1].level, 1);
    assert_eq!(list.items[2].level, 1);
    assert_eq!(list.items[3].level, 0);
    assert_eq!(list.items[0].start_at, Some(1));
    assert_eq!(
        list.items[3].start_at,
        None,
        "Top-level numbering should continue after nested bullet items"
    );
    assert!(list.level_styles.contains_key(&1));
}

#[test]
fn test_parse_mixed_list_and_paragraphs() {
    // A list followed by a regular paragraph should produce two separate blocks
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item 1"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item 2"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Regular paragraph")),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    // Should have at least a List block and a Paragraph block
    let list_count = page
        .content
        .iter()
        .filter(|b| matches!(b, Block::List(_)))
        .count();
    let para_count = page
        .content
        .iter()
        .filter(|b| matches!(b, Block::Paragraph(_)))
        .count();
    assert!(list_count >= 1, "Expected at least 1 list block");
    assert!(para_count >= 1, "Expected at least 1 paragraph block");
}

// ----- US-020: Header/footer parsing tests -----

/// Helper: build a DOCX with a text header.
fn build_docx_with_header(header_text: &str) -> Vec<u8> {
    let header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(header_text)),
    );
    let docx = docx_rs::Docx::new().header(header).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with a text footer.
fn build_docx_with_footer(footer_text: &str) -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(footer_text)),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with a page number field in footer.
fn build_docx_with_page_number_footer() -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Page ")
                .add_field_char(docx_rs::FieldCharType::Begin, false)
                .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                .add_field_char(docx_rs::FieldCharType::Separate, false)
                .add_text("1")
                .add_field_char(docx_rs::FieldCharType::End, false),
        ),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_parse_docx_with_text_header() {
    let data = build_docx_with_header("My Document Header");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    // Should have a header
    assert!(page.header.is_some(), "FlowPage should have a header");
    let header = page.header.as_ref().unwrap();
    assert!(
        !header.paragraphs.is_empty(),
        "Header should have paragraphs"
    );

    // Find the text run in header
    let has_text = header.paragraphs.iter().any(|p| {
        p.elements.iter().any(
            |e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("My Document Header")),
        )
    });
    assert!(
        has_text,
        "Header should contain the text 'My Document Header'"
    );
}

#[test]
fn test_parse_docx_with_text_footer() {
    let data = build_docx_with_footer("Footer Text");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.footer.is_some(), "FlowPage should have a footer");
    let footer = page.footer.as_ref().unwrap();

    let has_text = footer.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("Footer Text")))
    });
    assert!(has_text, "Footer should contain 'Footer Text'");
}

#[test]
fn test_parse_docx_with_page_number_in_footer() {
    let data = build_docx_with_page_number_footer();
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.footer.is_some(), "Should have footer");
    let footer = page.footer.as_ref().unwrap();

    // Footer should contain a PageNumber element
    let has_page_num = footer.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, crate::ir::HFInline::PageNumber))
    });
    assert!(has_page_num, "Footer should contain a PageNumber field");

    // Footer should also contain the "Page " text
    let has_text = footer.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("Page ")))
    });
    assert!(
        has_text,
        "Footer should contain 'Page ' text before page number"
    );
}

/// Helper: build a DOCX with a total page count field in footer.
fn build_docx_with_total_pages_footer() -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Total "))
            .add_run(
                docx_rs::Run::new()
                    .add_field_char(docx_rs::FieldCharType::Begin, false)
                    .add_instr_text(docx_rs::InstrText::NUMPAGES(docx_rs::InstrNUMPAGES::new()))
                    .add_field_char(docx_rs::FieldCharType::Separate, false)
                    .add_text("1")
                    .add_field_char(docx_rs::FieldCharType::End, false),
            ),
    );
    let docx = docx_rs::Docx::new()
        .footer(footer)
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_parse_docx_with_total_pages_in_footer() {
    let data = build_docx_with_total_pages_footer();
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let footer = page.footer.as_ref().expect("Should have footer");
    let has_total_pages = footer.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, crate::ir::HFInline::TotalPages))
    });
    assert!(has_total_pages, "Footer should contain a TotalPages field");
}

#[test]
fn test_parse_docx_multiple_sections_with_distinct_page_setup_and_headers() {
    let first_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One Header")),
    );
    let second_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two Header")),
    );

    let first_section = docx_rs::Section::new()
        .page_size(docx_rs::PageSize::new().size(12240, 15840))
        .header(first_header)
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One")),
        );

    let docx = docx_rs::Docx::new()
        .add_section(first_section)
        .header(second_header)
        .page_size(15840, 12240)
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two")),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2, "Expected one FlowPage per DOCX section");

    let first_page = match &doc.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected first page to be FlowPage"),
    };
    let second_page = match &doc.pages[1] {
        Page::Flow(page) => page,
        _ => panic!("Expected second page to be FlowPage"),
    };

    assert!(
        (first_page.size.width - 612.0).abs() < 0.1,
        "first page width should come from first section"
    );
    assert!(
        (first_page.size.height - 792.0).abs() < 0.1,
        "first page height should come from first section"
    );
    assert!(
        (second_page.size.width - 792.0).abs() < 0.1,
        "second page width should come from final section"
    );
    assert!(
        (second_page.size.height - 612.0).abs() < 0.1,
        "second page height should come from final section"
    );

    let first_header_text = first_page
        .header
        .as_ref()
        .and_then(|hf| {
            hf.paragraphs
                .iter()
                .flat_map(|p| p.elements.iter())
                .find_map(|e| match e {
                    crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
        })
        .unwrap_or("");
    assert_eq!(first_header_text, "Section One Header");

    let second_header_text = second_page
        .header
        .as_ref()
        .and_then(|hf| {
            hf.paragraphs
                .iter()
                .flat_map(|p| p.elements.iter())
                .find_map(|e| match e {
                    crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
        })
        .unwrap_or("");
    assert_eq!(second_header_text, "Section Two Header");
}

#[test]
fn test_parse_docx_with_header_and_footer() {
    let header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Header Text")),
    );
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Footer Text")),
    );
    let docx = docx_rs::Docx::new()
        .header(header)
        .footer(footer)
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.header.is_some(), "Should have header");
    assert!(page.footer.is_some(), "Should have footer");
}

#[test]
fn test_parse_docx_without_header_footer() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.header.is_none(), "No header expected");
    assert!(page.footer.is_none(), "No footer expected");
}

// ----- Page orientation tests -----

#[test]
fn test_portrait_document_width_less_than_height() {
    // Standard A4 portrait: 11906 x 16838 twips
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Portrait"))],
        11906,
        16838,
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
    assert!(
        page.size.width < page.size.height,
        "Portrait: width ({}) should be < height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_landscape_document_width_greater_than_height() {
    // Landscape A4: width and height swapped → 16838 x 11906 twips
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape"))],
        16838,
        11906,
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
    assert!(
        page.size.width > page.size.height,
        "Landscape: width ({}) should be > height ({})",
        page.size.width,
        page.size.height
    );
    // Verify approximate values: 16838/20 = 841.9pt, 11906/20 = 595.3pt
    assert!(
        (page.size.width - 841.9).abs() < 1.0,
        "Expected width ~841.9, got {}",
        page.size.width
    );
    assert!(
        (page.size.height - 595.3).abs() < 1.0,
        "Expected height ~595.3, got {}",
        page.size.height
    );
}

#[test]
fn test_default_document_is_portrait() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Default")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    // Default docx-rs page is A4 portrait
    assert!(
        page.size.width < page.size.height,
        "Default should be portrait: width ({}) < height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_landscape_with_orient_attribute() {
    // Build a landscape DOCX using page_orient + swapped dimensions
    let mut docx = docx_rs::Docx::new()
        .page_size(16838, 11906)
        .page_orient(docx_rs::PageOrientationType::Landscape)
        .page_margin(
            docx_rs::PageMargin::new()
                .top(1440)
                .bottom(1440)
                .left(1440)
                .right(1440),
        );
    docx = docx.add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape with orient")),
    );
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        page.size.width > page.size.height,
        "Landscape with orient: width ({}) should be > height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_extract_page_size_orient_landscape_swaps_dimensions() {
    // Edge case: orient=landscape but dimensions are portrait-style (w < h).
    // The parser should detect orient and swap width/height.
    let page_size = docx_rs::PageSize::new()
        .width(11906) // portrait w
        .height(16838) // portrait h
        .orient(docx_rs::PageOrientationType::Landscape);

    let result = extract_page_size(&page_size);
    assert!(
        result.width > result.height,
        "orient=landscape should ensure width ({}) > height ({})",
        result.width,
        result.height
    );
}

#[test]
fn test_extract_page_size_no_orient_keeps_dimensions() {
    // No orient attribute: dimensions should be used as-is
    let page_size = docx_rs::PageSize::new().width(11906).height(16838);

    let result = extract_page_size(&page_size);
    // 11906/20 = 595.3, 16838/20 = 841.9
    assert!(
        result.width < result.height,
        "No orient: width ({}) should be < height ({})",
        result.width,
        result.height
    );
}

// ----- Document styles tests (US-022) -----

/// Helper: build a DOCX with custom styles and paragraphs.
fn build_docx_bytes_with_styles(
    paragraphs: Vec<docx_rs::Paragraph>,
    styles: Vec<docx_rs::Style>,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new();
    for s in styles {
        docx = docx.add_style(s);
    }
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with an explicit stylesheet and paragraphs.
fn build_docx_bytes_with_stylesheet(
    paragraphs: Vec<docx_rs::Paragraph>,
    styles: docx_rs::Styles,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new().styles(styles);
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_heading1_style_applies_defaults() {
    // Create a Heading 1 style with outline level 0 (no explicit size/bold)
    let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Title"))
                .style("Heading1"),
        ],
        vec![h1_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    // Heading 1 default: 24pt bold
    assert_eq!(run.style.font_size, Some(24.0));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_heading2_style_applies_defaults() {
    let h2_style = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
        .name("Heading 2")
        .outline_lvl(1);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Subtitle"))
                .style("Heading2"),
        ],
        vec![h2_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    // Heading 2 default: 20pt bold
    assert_eq!(run.style.font_size, Some(20.0));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_heading3_through_6_defaults() {
    // Test heading levels 3-6 with their expected default sizes
    let expected: Vec<(usize, &str, f64)> = vec![
        (2, "Heading3", 16.0), // H3
        (3, "Heading4", 14.0), // H4
        (4, "Heading5", 12.0), // H5
        (5, "Heading6", 11.0), // H6
    ];

    for (outline_lvl, style_id, expected_size) in expected {
        let style = docx_rs::Style::new(style_id, docx_rs::StyleType::Paragraph)
            .name(format!("Heading {}", outline_lvl + 1))
            .outline_lvl(outline_lvl);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Heading text"))
                    .style(style_id),
            ],
            vec![style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        assert_eq!(
            run.style.font_size,
            Some(expected_size),
            "Heading {} should have size {expected_size}pt",
            outline_lvl + 1
        );
        assert_eq!(
            run.style.bold,
            Some(true),
            "Heading {} should be bold",
            outline_lvl + 1
        );
    }
}

#[test]
fn test_style_with_explicit_formatting() {
    // Style defines size=36 (half-points = 18pt) and bold
    let custom = docx_rs::Style::new("CustomStyle", docx_rs::StyleType::Paragraph)
        .name("Custom Style")
        .size(36) // 18pt in half-points
        .bold();

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Custom styled"))
                .style("CustomStyle"),
        ],
        vec![custom],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_explicit_run_formatting_overrides_style() {
    // Style says bold + 24pt (via heading defaults), but run explicitly sets size=20 (10pt)
    let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Small heading").size(20)) // 10pt
                .style("Heading1"),
        ],
        vec![h1_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    // Explicit size (10pt) overrides heading default (24pt)
    assert_eq!(run.style.font_size, Some(10.0));
    // Bold still comes from heading defaults since not explicitly overridden
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_style_alignment_applied_to_paragraph() {
    let centered = docx_rs::Style::new("CenteredStyle", docx_rs::StyleType::Paragraph)
        .name("Centered")
        .align(docx_rs::AlignmentType::Center);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Centered paragraph"))
                .style("CenteredStyle"),
        ],
        vec![centered],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    assert_eq!(para.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_style_inherits_spacing_from_based_on_parent() {
    let normal = docx_rs::Style::new("Normal", docx_rs::StyleType::Paragraph)
        .name("Normal")
        .line_spacing(docx_rs::LineSpacing::new().before(200).after(40));
    let heading4 = docx_rs::Style::new("Heading4", docx_rs::StyleType::Paragraph)
        .name("Heading 4")
        .based_on("Normal")
        .outline_lvl(3);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Inherited spacing heading"))
                .style("Heading4"),
        ],
        vec![normal, heading4],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    assert_eq!(para.style.space_before, Some(10.0));
    assert_eq!(para.style.space_after, Some(2.0));
    assert_eq!(para.style.heading_level, Some(4));
}

#[test]
fn test_normal_style_no_heading_defaults() {
    // Normal paragraphs (no heading) should not get heading defaults
    let normal = docx_rs::Style::new("Normal", docx_rs::StyleType::Paragraph).name("Normal");

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Normal text"))
                .style("Normal"),
        ],
        vec![normal],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    // Normal style should NOT have heading defaults
    assert!(run.style.font_size.is_none());
    assert!(run.style.bold.is_none());
}

#[test]
fn test_heading_with_mixed_paragraphs() {
    // Document with Heading 1, Normal, Heading 2 paragraphs
    let h1 = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);
    let h2 = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
        .name("Heading 2")
        .outline_lvl(1);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Title"))
                .style("Heading1"),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Subtitle"))
                .style("Heading2"),
        ],
        vec![h1, h2],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let blocks = all_blocks(&doc);

    // First paragraph: Heading 1
    if let Block::Paragraph(p) = &blocks[0] {
        assert_eq!(p.runs[0].style.font_size, Some(24.0));
        assert_eq!(p.runs[0].style.bold, Some(true));
    } else {
        panic!("Expected Paragraph");
    }

    // Second paragraph: Normal (no style)
    if let Block::Paragraph(p) = &blocks[1] {
        assert!(p.runs[0].style.font_size.is_none());
        assert!(p.runs[0].style.bold.is_none());
    } else {
        panic!("Expected Paragraph");
    }

    // Third paragraph: Heading 2
    if let Block::Paragraph(p) = &blocks[2] {
        assert_eq!(p.runs[0].style.font_size, Some(20.0));
        assert_eq!(p.runs[0].style.bold, Some(true));
    } else {
        panic!("Expected Paragraph");
    }
}

#[test]
fn test_style_with_color_and_font() {
    let custom = docx_rs::Style::new("Fancy", docx_rs::StyleType::Paragraph)
        .name("Fancy Style")
        .color("FF0000")
        .fonts(docx_rs::RunFonts::new().ascii("Georgia"));

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Fancy text"))
                .style("Fancy"),
        ],
        vec![custom],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    assert_eq!(run.style.font_family, Some("Georgia".to_string()));
}

#[test]
fn test_runs_inherit_document_default_font() {
    let styles = docx_rs::Styles::new()
        .default_fonts(docx_rs::RunFonts::new().ascii("Raleway"))
        .default_size(18);

    let link = docx_rs::Hyperlink::new("https://example.com", docx_rs::HyperlinkType::External)
        .add_run(
            docx_rs::Run::new()
                .color("1155cc")
                .underline("single")
                .add_text("Linked text"),
        );
    let paragraph = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Plain text "))
        .add_hyperlink(link);
    let data = build_docx_bytes_with_stylesheet(vec![paragraph], styles);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].style.font_family.as_deref(), Some("Raleway"));
    assert_eq!(para.runs[0].style.font_size, Some(9.0));
    assert_eq!(para.runs[1].href.as_deref(), Some("https://example.com"));
    assert_eq!(para.runs[1].style.font_family.as_deref(), Some("Raleway"));
    assert_eq!(para.runs[1].style.font_size, Some(9.0));
    assert_eq!(para.runs[1].style.color, Some(Color::new(17, 85, 204)));
    assert_eq!(para.runs[1].style.underline, Some(true));
}

// ----- Hyperlink tests (US-030) -----

#[test]
fn test_hyperlink_single_link_in_paragraph() {
    let link = docx_rs::Hyperlink::new("https://example.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Click here"));
    let para = docx_rs::Paragraph::new().add_hyperlink(link);
    let data = build_docx_bytes(vec![para]);

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

    assert_eq!(para.runs.len(), 1);
    assert_eq!(para.runs[0].text, "Click here");
    assert_eq!(para.runs[0].href, Some("https://example.com".to_string()));
}

#[test]
fn test_hyperlink_mixed_text_and_link() {
    let link = docx_rs::Hyperlink::new("https://rust-lang.org", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Rust"));
    let para = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Visit "))
        .add_hyperlink(link)
        .add_run(docx_rs::Run::new().add_text(" for more."));
    let data = build_docx_bytes(vec![para]);

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

    // Should have 3 runs: "Visit ", hyperlink "Rust", " for more."
    assert_eq!(para.runs.len(), 3);

    assert_eq!(para.runs[0].text, "Visit ");
    assert_eq!(para.runs[0].href, None);

    assert_eq!(para.runs[1].text, "Rust");
    assert_eq!(para.runs[1].href, Some("https://rust-lang.org".to_string()));

    assert_eq!(para.runs[2].text, " for more.");
    assert_eq!(para.runs[2].href, None);
}

#[test]
fn test_hyperlink_multiple_links_in_paragraph() {
    let link1 = docx_rs::Hyperlink::new("https://first.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("First"));
    let link2 = docx_rs::Hyperlink::new("https://second.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Second"));
    let para = docx_rs::Paragraph::new()
        .add_hyperlink(link1)
        .add_run(docx_rs::Run::new().add_text(" and "))
        .add_hyperlink(link2);
    let data = build_docx_bytes(vec![para]);

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

    assert_eq!(para.runs[0].text, "First");
    assert_eq!(para.runs[0].href, Some("https://first.com".to_string()));

    assert_eq!(para.runs[1].text, " and ");
    assert_eq!(para.runs[1].href, None);

    assert_eq!(para.runs[2].text, "Second");
    assert_eq!(para.runs[2].href, Some("https://second.com".to_string()));
}

// ── Footnotes and endnotes ──────────────────────────────────────────

#[test]
fn test_footnote_single_in_paragraph() {
    let footnote = docx_rs::Footnote::new().add_content(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("This is a footnote.")),
    );

    let para = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Some text"))
        .add_run(docx_rs::Run::new().add_footnote_reference(footnote))
        .add_run(docx_rs::Run::new().add_text(" after note."));

    let data = build_docx_bytes(vec![para]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected flow page"),
    };

    // Find the paragraph
    let para = match &flow.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected paragraph"),
    };

    // Should have runs including a footnote reference
    let note_run = para.runs.iter().find(|r| r.footnote.is_some());
    assert!(note_run.is_some(), "Expected a run with footnote content");
    assert_eq!(
        note_run.unwrap().footnote.as_deref(),
        Some("This is a footnote.")
    );
}

#[test]
fn test_footnote_multiple_in_paragraph() {
    let fn1 = docx_rs::Footnote::new().add_content(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First note.")),
    );
    let fn2 = docx_rs::Footnote::new().add_content(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second note.")),
    );

    let para = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("A"))
        .add_run(docx_rs::Run::new().add_footnote_reference(fn1))
        .add_run(docx_rs::Run::new().add_text(" B"))
        .add_run(docx_rs::Run::new().add_footnote_reference(fn2));

    let data = build_docx_bytes(vec![para]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected flow page"),
    };

    let para = match &flow.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected paragraph"),
    };

    let note_runs: Vec<_> = para.runs.iter().filter(|r| r.footnote.is_some()).collect();
    assert_eq!(note_runs.len(), 2);
    assert_eq!(note_runs[0].footnote.as_deref(), Some("First note."));
    assert_eq!(note_runs[1].footnote.as_deref(), Some("Second note."));
}

#[test]
fn test_endnote_parsed_as_footnote() {
    // docx-rs doesn't support endnotes, so we build a minimal DOCX ZIP manually
    // with word/endnotes.xml and w:endnoteReference in document.xml
    let data = build_docx_with_endnote("Text before endnote", 1, "This is an endnote.");

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected flow page"),
    };

    let para = match &flow.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected paragraph"),
    };

    let note_run = para.runs.iter().find(|r| r.footnote.is_some());
    assert!(note_run.is_some(), "Expected a run with endnote content");
    assert_eq!(
        note_run.unwrap().footnote.as_deref(),
        Some("This is an endnote.")
    );
}

/// Build a minimal DOCX ZIP with an endnote reference in the document body
/// and endnote content in word/endnotes.xml.
fn build_docx_with_endnote(text: &str, endnote_id: usize, endnote_text: &str) -> Vec<u8> {
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    let buf = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(buf));
    let opts = FileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"#).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", opts)
        .unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"#).unwrap();

    // word/document.xml - with endnoteReference
    zip.start_file("word/document.xml", opts).unwrap();
    let doc_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">{text}</w:t></w:r>
      <w:r>
        <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
        <w:endnoteReference w:id="{endnote_id}"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#
    );
    zip.write_all(doc_xml.as_bytes()).unwrap();

    // word/endnotes.xml
    zip.start_file("word/endnotes.xml", opts).unwrap();
    let endnotes_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="{endnote_id}">
    <w:p>
      <w:r><w:t xml:space="preserve">{endnote_text}</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>"#
    );
    zip.write_all(endnotes_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

// ----- Table of Contents (TOC) parsing tests -----

/// Helper: build a DOCX with a table of contents containing items.
fn build_docx_with_toc(items: Vec<docx_rs::TableOfContentsItem>) -> Vec<u8> {
    let toc = items.into_iter().fold(
        docx_rs::TableOfContents::new()
            .heading_styles_range(1, 3)
            .alias("Table of contents"),
        |toc, item| toc.add_item(item),
    );

    let style1 = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph).name("Heading 1");
    let style2 = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph).name("Heading 2");

    let p1 = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Introduction"))
        .style("Heading1");
    let p2 = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Details"))
        .style("Heading2");

    let docx = docx_rs::Docx::new()
        .add_style(style1)
        .add_style(style2)
        .add_table_of_contents(toc)
        .add_paragraph(p1)
        .add_paragraph(p2);

    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_docx_toc_with_entries() {
    let items = vec![
        docx_rs::TableOfContentsItem::new()
            .text("Introduction")
            .toc_key("_Toc00000000")
            .level(1)
            .page_ref("2"),
        docx_rs::TableOfContentsItem::new()
            .text("Details")
            .toc_key("_Toc00000001")
            .level(2)
            .page_ref("3"),
    ];

    let data = build_docx_with_toc(items);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    // The document should have content, and the TOC entry texts should be present
    let page = &doc.pages[0];
    let content = match page {
        Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };

    // Collect all text from all paragraphs
    let all_text: Vec<String> = content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => {
                let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                if t.is_empty() { None } else { Some(t) }
            }
            _ => None,
        })
        .collect();

    // TOC entries "Introduction" and "Details" should appear in the output
    // (along with the heading paragraphs themselves)
    let has_introduction = all_text.iter().any(|t| t.contains("Introduction"));
    let has_details = all_text.iter().any(|t| t.contains("Details"));
    assert!(
        has_introduction,
        "Expected 'Introduction' in TOC output, got: {all_text:?}"
    );
    assert!(
        has_details,
        "Expected 'Details' in TOC output, got: {all_text:?}"
    );
}

#[test]
fn test_docx_toc_multiple_entries_in_paragraph_list() {
    let items = vec![
        docx_rs::TableOfContentsItem::new()
            .text("Chapter One")
            .toc_key("_Toc10000001")
            .level(1)
            .page_ref("1"),
        docx_rs::TableOfContentsItem::new()
            .text("Chapter Two")
            .toc_key("_Toc10000002")
            .level(1)
            .page_ref("5"),
        docx_rs::TableOfContentsItem::new()
            .text("Section A")
            .toc_key("_Toc10000003")
            .level(2)
            .page_ref("10"),
    ];

    let data = build_docx_with_toc(items);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = &doc.pages[0];
    let content = match page {
        Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };

    // All three TOC entry texts should appear
    let all_text: Vec<String> = content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => {
                let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                if t.is_empty() { None } else { Some(t) }
            }
            _ => None,
        })
        .collect();

    assert!(
        all_text.iter().any(|t| t.contains("Chapter One")),
        "Expected 'Chapter One' in output, got: {all_text:?}"
    );
    assert!(
        all_text.iter().any(|t| t.contains("Chapter Two")),
        "Expected 'Chapter Two' in output, got: {all_text:?}"
    );
    assert!(
        all_text.iter().any(|t| t.contains("Section A")),
        "Expected 'Section A' in output, got: {all_text:?}"
    );
}

#[test]
fn test_docx_sdt_with_paragraphs() {
    // Test that generic SDTs with paragraph content are also parsed.
    // Build a DOCX manually using docx-rs StructuredDataTag.
    let sdt = docx_rs::StructuredDataTag::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("SDT Content")),
    );

    let docx = docx_rs::Docx::new().add_structured_data_tag(sdt);

    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = &doc.pages[0];
    let content = match page {
        Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };

    let all_text: Vec<String> = content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => {
                let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                if t.is_empty() { None } else { Some(t) }
            }
            _ => None,
        })
        .collect();

    assert!(
        all_text.iter().any(|t| t.contains("SDT Content")),
        "Expected 'SDT Content' in output, got: {all_text:?}"
    );
}

#[test]
fn test_docx_drawing_text_box_paragraph_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p>
                                                <w:r><w:t>Inside box</w:t></w:r>
                                            </w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Before".to_string(),
            "Inside box".to_string(),
            "After".to_string(),
        ]
    );
}

#[test]
fn test_docx_drawing_text_box_multiple_paragraphs_are_emitted_in_order() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 2"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>First line</w:t></w:r></w:p>
                                            <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Lead-in".to_string(),
            "First line".to_string(),
            "Second line".to_string(),
            "Tail".to_string(),
        ]
    );
}

#[test]
fn test_docx_drawing_text_box_table_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Before table box</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box Table"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:tbl>
                                                <w:tblPr/>
                                                <w:tblGrid>
                                                    <w:gridCol w:w="2000"/>
                                                    <w:gridCol w:w="2000"/>
                                                </w:tblGrid>
                                                <w:tr>
                                                    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                                                    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                                                </w:tr>
                                            </w:tbl>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>After table box</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let has_table = flow
        .content
        .iter()
        .any(|block| matches!(block, Block::Table(_)));
    assert!(has_table, "Expected a table extracted from text box");

    let table = first_table(&doc);
    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 2);

    let cell_text: Vec<String> = table.rows[0]
        .cells
        .iter()
        .map(|cell| {
            cell.content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(
                        p.runs
                            .iter()
                            .map(|run| run.text.as_str())
                            .collect::<String>(),
                    ),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect();
    assert_eq!(cell_text, vec!["A".to_string(), "B".to_string()]);
}

#[test]
fn test_docx_vml_text_box_paragraph_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox1" style="width:100pt;height:40pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Before".to_string(),
            "VML box".to_string(),
            "After".to_string(),
        ]
    );
}

#[test]
fn test_docx_vml_text_box_multiple_paragraphs_are_emitted_in_order() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox2" style="width:120pt;height:60pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>First VML line</w:t></w:r></w:p>
                                <w:p><w:r><w:t>Second VML line</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Lead-in".to_string(),
            "First VML line".to_string(),
            "Second VML line".to_string(),
            "Tail".to_string(),
        ]
    );
}

#[test]
fn test_docx_vml_floating_text_box_square_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox3"
                             style="position:absolute;margin-left:72pt;margin-top:36pt;width:144pt;height:72pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML floating box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="square"/>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating VML text box");

    let ftb = floating[0];
    assert_eq!(ftb.wrap_mode, WrapMode::Square);
    assert!((ftb.offset_x - 72.0).abs() < 0.5);
    assert!((ftb.offset_y - 36.0).abs() < 0.5);
    assert!((ftb.width - 144.0).abs() < 0.5);
    assert!((ftb.height - 72.0).abs() < 0.5);

    let texts: Vec<String> = ftb
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    assert_eq!(texts, vec!["VML floating box".to_string()]);
}

#[test]
fn test_docx_vml_floating_text_box_none_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox4"
                             style="position:absolute;margin-left:12pt;margin-top:18pt;width:90pt;height:40pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>No wrap box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="none"/>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating VML text box");
    assert_eq!(floating[0].wrap_mode, WrapMode::None);
}

/// Helper: find all FloatingTextBox blocks in a FlowPage.
fn find_floating_text_boxes(doc: &Document) -> Vec<&FloatingTextBox> {
    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|b| match b {
            Block::FloatingTextBox(ftb) => Some(ftb),
            _ => None,
        })
        .collect()
}

#[test]
fn test_docx_floating_text_box_square_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="0" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                        <wp:extent cx="1828800" cy="914400"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapSquare wrapText="bothSides"/>
                        <wp:docPr id="1" name="Anchored Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Inside anchored box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating text box");

    let ftb = floating[0];
    assert_eq!(ftb.wrap_mode, WrapMode::Square);
    assert!(
        (ftb.offset_x - 72.0).abs() < 0.5,
        "Expected offset_x ~72pt, got {}",
        ftb.offset_x
    );
    assert!(
        (ftb.offset_y - 36.0).abs() < 0.5,
        "Expected offset_y ~36pt, got {}",
        ftb.offset_y
    );
    assert!(
        (ftb.width - 144.0).abs() < 0.5,
        "Expected width ~144pt, got {}",
        ftb.width
    );
    assert!(
        (ftb.height - 72.0).abs() < 0.5,
        "Expected height ~72pt, got {}",
        ftb.height
    );

    let texts: Vec<String> = ftb
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    assert_eq!(texts, vec!["Inside anchored box".to_string()]);
}

#[test]
fn test_docx_floating_text_box_top_and_bottom_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="1" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>
                        <wp:extent cx="1270000" cy="635000"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapTopAndBottom/>
                        <wp:docPr id="2" name="Top Bottom Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Top and bottom box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating text box");
    assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
}

// ── Floating image (anchor) tests ──

/// Helper: find all FloatingImage blocks in a FlowPage.
fn find_floating_images(doc: &Document) -> Vec<&FloatingImage> {
    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|b| match b {
            Block::FloatingImage(fi) => Some(fi),
            _ => None,
        })
        .collect()
}

#[test]
fn test_docx_floating_image_square_wrap() {
    // Build a floating image with wrapSquare (allow_overlap = false, floating)
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(2_540_000, 1_270_000) // 200pt × 100pt
        .floating()
        .offset_x(914_400) // 72pt (1 inch)
        .offset_y(457_200); // 36pt (0.5 inch)
    // allow_overlap defaults to false for floating() → docx-rs writes wrapSquare

    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating image");

    let fi = floating[0];
    assert_eq!(fi.wrap_mode, WrapMode::Square);
    assert!(!fi.image.data.is_empty(), "Image data should not be empty");

    // Check dimensions
    let width = fi.image.width.expect("Expected width");
    let height = fi.image.height.expect("Expected height");
    assert!(
        (width - 200.0).abs() < 0.5,
        "Expected width ~200pt, got {width}"
    );
    assert!(
        (height - 100.0).abs() < 0.5,
        "Expected height ~100pt, got {height}"
    );
}

#[test]
fn test_docx_floating_image_top_and_bottom_wrap() {
    // Build a DOCX manually with wrapTopAndBottom
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000) // 100pt × 100pt
        .floating()
        .overlapping(); // allow_overlap=true → docx-rs writes wrapNone

    // Build DOCX bytes
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let mut data = cursor.into_inner();

    // Patch the DOCX XML to replace wrapNone with wrapTopAndBottom
    data = patch_docx_wrap_type(&data, "wp:wrapNone", "wp:wrapTopAndBottom");

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating image");
    assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
}

#[test]
fn test_docx_floating_image_behind_wrap() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000)
        .floating()
        .overlapping(); // generates wrapNone

    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let mut data = cursor.into_inner();

    // Patch to wrapNone → behindDoc attribute + wrapNone
    // Behind text is indicated by behindDoc="1" attribute on wp:anchor,
    // combined with wrapNone. Our scan should detect this.
    data = patch_docx_behind_doc(&data);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating image");
    // behindDoc with wrapNone → Behind wrap mode
    assert_eq!(floating[0].wrap_mode, WrapMode::Behind);
}

#[test]
fn test_docx_floating_image_position_offset() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000) // 100pt × 100pt
        .floating()
        .offset_x(914_400) // 72pt (1 inch in EMU)
        .offset_y(457_200); // 36pt (0.5 inch in EMU)

    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating image");

    let fi = floating[0];
    assert!(
        (fi.offset_x - 72.0).abs() < 0.5,
        "Expected offset_x ~72pt, got {}",
        fi.offset_x
    );
    assert!(
        (fi.offset_y - 36.0).abs() < 0.5,
        "Expected offset_y ~36pt, got {}",
        fi.offset_y
    );
}

#[test]
fn test_docx_inline_image_not_floating() {
    // Inline images should NOT become FloatingImage
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(
        floating.len(),
        0,
        "Inline images should not be floating images"
    );

    let images = find_images(&doc);
    assert_eq!(images.len(), 1, "Should still find the inline image");
}

/// Helper: Patch a DOCX ZIP by replacing a wrap element in document.xml.
fn patch_docx_wrap_type(data: &[u8], old_wrap: &str, new_wrap: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::FileOptions::default();
        new_zip.start_file(&name, options).unwrap();

        let mut contents = Vec::new();
        file.read_to_end(&mut contents).unwrap();

        if name == "word/document.xml" {
            let xml = String::from_utf8(contents).unwrap();
            // Replace self-closing: <wp:wrapNone /> → <wp:wrapTopAndBottom />
            let xml = xml
                .replace(&format!("<{old_wrap} />"), &format!("<{new_wrap} />"))
                .replace(&format!("<{old_wrap}/>"), &format!("<{new_wrap}/>"));
            std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
        } else {
            std::io::Write::write_all(&mut new_zip, &contents).unwrap();
        }
    }

    new_zip.finish().unwrap().into_inner()
}

/// Helper: Patch a DOCX ZIP to set behindDoc="1" on wp:anchor in document.xml.
fn patch_docx_behind_doc(data: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::FileOptions::default();
        new_zip.start_file(&name, options).unwrap();

        let mut contents = Vec::new();
        file.read_to_end(&mut contents).unwrap();

        if name == "word/document.xml" {
            let xml = String::from_utf8(contents).unwrap();
            // Replace existing behindDoc="0" with behindDoc="1"
            let xml = xml.replace("behindDoc=\"0\"", "behindDoc=\"1\"");
            std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
        } else {
            std::io::Write::write_all(&mut new_zip, &contents).unwrap();
        }
    }

    new_zip.finish().unwrap().into_inner()
}

// ── OMML math equation tests ──

/// Build a DOCX ZIP with a custom document.xml containing OMML math.
fn build_docx_with_math(document_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )
    .unwrap();

    // word/document.xml
    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_docx_with_display_math_fraction() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>Before math</w:t></w:r>
        </w:p>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:f>
                        <m:num><m:r><m:t>a</m:t></m:r></m:num>
                        <m:den><m:r><m:t>b</m:t></m:r></m:den>
                    </m:f>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:p>
            <w:r><w:t>After math</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    // Should find a MathEquation block
    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::MathEquation(m) => Some(m),
            _ => None,
        })
        .collect();

    assert!(
        !math_blocks.is_empty(),
        "Expected at least one MathEquation block, found none"
    );
    assert_eq!(math_blocks[0].content, "frac(a, b)");
    assert!(math_blocks[0].display);
}

#[test]
fn test_parse_docx_with_inline_math_superscript() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>The value of </w:t></w:r>
            <m:oMath>
                <m:sSup>
                    <m:e><m:r><m:t>x</m:t></m:r></m:e>
                    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
            </m:oMath>
            <w:r><w:t> is positive</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::MathEquation(m) => Some(m),
            _ => None,
        })
        .collect();

    assert!(
        !math_blocks.is_empty(),
        "Expected at least one MathEquation block"
    );
    assert_eq!(math_blocks[0].content, "x^2");
    assert!(!math_blocks[0].display);
}

#[test]
fn test_parse_docx_with_complex_math() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:r><m:t>E</m:t></m:r>
                    <m:r><m:t>=</m:t></m:r>
                    <m:r><m:t>m</m:t></m:r>
                    <m:sSup>
                        <m:e><m:r><m:t>c</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                    </m:sSup>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::MathEquation(m) => Some(m),
            _ => None,
        })
        .collect();

    assert!(!math_blocks.is_empty());
    // Space before sSup separates run "m" from base "c" to prevent identifier
    // concatenation (both are semantically equivalent in Typst math: m × c²)
    assert_eq!(math_blocks[0].content, "E=m c^2");
    assert!(math_blocks[0].display);
}

/// Build a DOCX ZIP with a chart embedded in it.
fn build_docx_with_chart(document_xml: &str, chart_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#,
        )
        .unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>"#,
        )
        .unwrap();

    // word/document.xml
    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    // word/charts/chart1.xml
    zip.start_file("word/charts/chart1.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, chart_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_docx_with_bar_chart() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt>
            <c:pt idx="1"><c:v>Q2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>100</c:v></c:pt>
            <c:pt idx="1"><c:v>200</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let data = build_docx_with_chart(document_xml, chart_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("parse should succeed");

    let content = match &doc.pages[0] {
        crate::ir::Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };
    let chart_blocks: Vec<&crate::ir::Chart> = content
        .iter()
        .filter_map(|b| match b {
            Block::Chart(c) => Some(c),
            _ => None,
        })
        .collect();

    assert_eq!(chart_blocks.len(), 1);
    assert_eq!(chart_blocks[0].chart_type, crate::ir::ChartType::Bar);
    assert_eq!(chart_blocks[0].title.as_deref(), Some("Sales"));
    assert_eq!(chart_blocks[0].categories, vec!["Q1", "Q2"]);
    assert_eq!(chart_blocks[0].series.len(), 1);
    assert_eq!(chart_blocks[0].series[0].name.as_deref(), Some("Revenue"));
    assert_eq!(chart_blocks[0].series[0].values, vec![100.0, 200.0]);
}

#[test]
fn test_parse_docx_with_pie_chart() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:cat><c:strLit>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
            <c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strLit></c:cat>
          <c:val><c:numLit>
            <c:pt idx="0"><c:v>30</c:v></c:pt>
            <c:pt idx="1"><c:v>50</c:v></c:pt>
            <c:pt idx="2"><c:v>20</c:v></c:pt>
          </c:numLit></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let data = build_docx_with_chart(document_xml, chart_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("parse should succeed");

    let content = match &doc.pages[0] {
        crate::ir::Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };
    let chart_blocks: Vec<&crate::ir::Chart> = content
        .iter()
        .filter_map(|b| match b {
            Block::Chart(c) => Some(c),
            _ => None,
        })
        .collect();

    assert_eq!(chart_blocks.len(), 1);
    assert_eq!(chart_blocks[0].chart_type, crate::ir::ChartType::Pie);
    assert!(chart_blocks[0].title.is_none());
    assert_eq!(chart_blocks[0].categories, vec!["A", "B", "C"]);
    assert_eq!(chart_blocks[0].series[0].values, vec![30.0, 50.0, 20.0]);
}

// ── Metadata extraction tests ──────────────────────────────────────

/// Build a minimal DOCX with docProps/core.xml containing metadata.
fn build_docx_with_metadata(core_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
    )
    .unwrap();

    zip.start_file("docProps/core.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, core_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_docx_extracts_metadata() {
    let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My DOCX Title</dc:title>
  <dc:creator>DOCX Author</dc:creator>
  <dc:subject>DOCX Subject</dc:subject>
  <dc:description>DOCX description text</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-03-15T08:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-04-20T12:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

    let data = build_docx_with_metadata(core_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.metadata.title.as_deref(), Some("My DOCX Title"));
    assert_eq!(doc.metadata.author.as_deref(), Some("DOCX Author"));
    assert_eq!(doc.metadata.subject.as_deref(), Some("DOCX Subject"));
    assert_eq!(
        doc.metadata.description.as_deref(),
        Some("DOCX description text")
    );
    assert_eq!(
        doc.metadata.created.as_deref(),
        Some("2024-03-15T08:00:00Z")
    );
    assert_eq!(
        doc.metadata.modified.as_deref(),
        Some("2024-04-20T12:30:00Z")
    );
}

#[test]
fn test_parse_docx_without_metadata_no_crash() {
    // A minimal DOCX without docProps/core.xml → defaults
    let data = build_docx_with_math(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>No metadata</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    // Should not crash; fields are None
    assert!(doc.metadata.title.is_none());
    assert!(doc.metadata.author.is_none());
}

// --- Heading level IR tests (US-096) ---

#[test]
fn test_heading1_sets_heading_level_in_ir() {
    let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Title"))
                .style("Heading1"),
        ],
        vec![h1_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(
        para.style.heading_level,
        Some(1),
        "Heading 1 (outline_lvl 0) should set heading_level = 1"
    );
}

#[test]
fn test_heading2_sets_heading_level_in_ir() {
    let h2_style = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
        .name("Heading 2")
        .outline_lvl(1);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Subtitle"))
                .style("Heading2"),
        ],
        vec![h2_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(
        para.style.heading_level,
        Some(2),
        "Heading 2 (outline_lvl 1) should set heading_level = 2"
    );
}

#[test]
fn test_normal_paragraph_no_heading_level() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Normal text")),
    ]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);
    assert_eq!(
        para.style.heading_level, None,
        "Normal paragraph should not have heading_level"
    );
}

// --- US-103: Multi-column section layout tests ---

/// Helper: build a DOCX from raw document.xml (reuses build_docx_with_math pattern)
fn build_docx_with_columns(document_xml: &str) -> Vec<u8> {
    build_docx_with_math(document_xml)
}

#[test]
fn test_parse_docx_two_column_equal() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Column content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="2" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let cols = flow.columns.as_ref().expect("Should have column layout");
    assert_eq!(cols.num_columns, 2);
    // 720 twips = 36pt
    assert!(
        (cols.spacing - 36.0).abs() < 0.1,
        "spacing: {}",
        cols.spacing
    );
    assert!(
        cols.column_widths.is_none(),
        "Equal columns should not have per-column widths"
    );
}

#[test]
fn test_parse_docx_section_specific_column_layouts() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Section one intro</w:t></w:r></w:p>
        <w:p>
            <w:pPr>
                <w:sectPr>
                    <w:cols w:num="2" w:space="720"/>
                </w:sectPr>
            </w:pPr>
            <w:r><w:t>Section one end</w:t></w:r>
        </w:p>
        <w:p><w:r><w:t>Section two content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="1" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2, "Expected one FlowPage per section");

    let first = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    let second = match &doc.pages[1] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        first.columns.as_ref().map(|layout| layout.num_columns),
        Some(2),
        "First section should keep the two-column layout"
    );
    assert!(
        second.columns.is_none(),
        "Final single-column section should not expose a column layout"
    );
}

#[test]
fn test_parse_docx_three_column_equal() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="3" w:space="360"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let cols = flow.columns.as_ref().expect("Should have column layout");
    assert_eq!(cols.num_columns, 3);
    // 360 twips = 18pt
    assert!((cols.spacing - 18.0).abs() < 0.1);
}

#[test]
fn test_parse_docx_unequal_columns() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="2" w:space="720" w:equalWidth="0">
                <w:col w:w="6000" w:space="720"/>
                <w:col w:w="3000"/>
            </w:cols>
        </w:sectPr>
    </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let cols = flow.columns.as_ref().expect("Should have column layout");
    assert_eq!(cols.num_columns, 2);
    let widths = cols
        .column_widths
        .as_ref()
        .expect("Should have per-column widths");
    assert_eq!(widths.len(), 2);
    // 6000 twips = 300pt, 3000 twips = 150pt
    assert!((widths[0] - 300.0).abs() < 0.1, "width[0]: {}", widths[0]);
    assert!((widths[1] - 150.0).abs() < 0.1, "width[1]: {}", widths[1]);
}

#[test]
fn test_parse_docx_no_columns() {
    // Default document without w:cols should have no column layout
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Normal")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        flow.columns.is_none(),
        "Normal doc should not have column layout"
    );
}

#[test]
fn test_parse_docx_column_break() {
    // Test that w:br with type="column" produces Block::ColumnBreak
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Before"))
            .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Column))
            .add_run(docx_rs::Run::new().add_text("After")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };

    // Should have a ColumnBreak block
    let has_col_break = flow.content.iter().any(|b| matches!(b, Block::ColumnBreak));
    assert!(
        has_col_break,
        "Should have a ColumnBreak block. Blocks: {:?}",
        flow.content
            .iter()
            .map(std::mem::discriminant)
            .collect::<Vec<_>>()
    );
}

#[test]
fn test_parse_docx_single_column_no_layout() {
    // w:cols with num="1" should not produce column layout
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="1" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        flow.columns.is_none(),
        "Single column should not produce column layout"
    );
}

#[test]
fn test_extract_tab_stops_preserves_explicit_clear_override() {
    let tabs = vec![
        docx_rs::Tab::new()
            .val(docx_rs::TabValueType::Clear)
            .pos(1440),
    ];

    let tab_stops = extract_tab_stops(&tabs);

    assert_eq!(
        tab_stops,
        Some(vec![]),
        "A paragraph-level clear tab must remain an explicit empty override"
    );
}

#[test]
fn test_merge_paragraph_style_preserves_inherited_tabs_not_overridden() {
    let explicit_prop = docx_rs::ParagraphProperty::new().add_tab(
        docx_rs::Tab::new()
            .val(docx_rs::TabValueType::Left)
            .pos(2160),
    );
    let explicit = extract_paragraph_style(&explicit_prop);
    let explicit_tab_overrides = extract_tab_stop_overrides(&explicit_prop.tabs);
    let style = ResolvedStyle {
        text: TextStyle::default(),
        paragraph: ParagraphStyle {
            tab_stops: Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 144.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            ..ParagraphStyle::default()
        },
        paragraph_tab_overrides: None,
        heading_level: None,
    };

    let merged = merge_paragraph_style(&explicit, explicit_tab_overrides.as_deref(), Some(&style));

    assert_eq!(
        merged.tab_stops,
        Some(vec![
            TabStop {
                position: 72.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            },
            TabStop {
                position: 108.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            },
            TabStop {
                position: 144.0,
                alignment: TabAlignment::Right,
                leader: TabLeader::Dot,
            },
        ]),
        "Paragraph-level tabs should extend inherited style tabs instead of replacing them"
    );
}

#[test]
fn test_merge_paragraph_style_clears_only_targeted_inherited_tab_stop() {
    let explicit_prop = docx_rs::ParagraphProperty::new()
        .add_tab(
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Clear)
                .pos(2880),
        )
        .add_tab(
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Left)
                .pos(2160),
        );
    let explicit = extract_paragraph_style(&explicit_prop);
    let explicit_tab_overrides = extract_tab_stop_overrides(&explicit_prop.tabs);
    let style = ResolvedStyle {
        text: TextStyle::default(),
        paragraph: ParagraphStyle {
            tab_stops: Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 144.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            ..ParagraphStyle::default()
        },
        paragraph_tab_overrides: None,
        heading_level: None,
    };

    let merged = merge_paragraph_style(&explicit, explicit_tab_overrides.as_deref(), Some(&style));

    assert_eq!(
        merged.tab_stops,
        Some(vec![
            TabStop {
                position: 72.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            },
            TabStop {
                position: 108.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            },
        ]),
        "A clear tab should remove only the matching inherited stop, not the whole inherited list"
    );
}

#[test]
fn test_merge_paragraph_style_allows_clearing_inherited_tab_stops() {
    let inherited = TabStop {
        position: 72.0,
        alignment: TabAlignment::Left,
        leader: TabLeader::None,
    };
    let explicit = ParagraphStyle {
        tab_stops: Some(vec![]),
        ..ParagraphStyle::default()
    };
    let style = ResolvedStyle {
        text: TextStyle::default(),
        paragraph: ParagraphStyle {
            tab_stops: Some(vec![inherited]),
            ..ParagraphStyle::default()
        },
        paragraph_tab_overrides: None,
        heading_level: None,
    };

    let merged = merge_paragraph_style(&explicit, None, Some(&style));

    assert_eq!(
        merged.tab_stops,
        Some(vec![]),
        "Explicit paragraph tab clearing must override inherited style tab stops"
    );
}

// ── BiDi / RTL tests ──────────────────────────────────────────────

/// Helper: create a bidi paragraph with the given text.
fn make_bidi_paragraph(text: &str) -> docx_rs::Paragraph {
    let mut para = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(text));
    para.property = docx_rs::ParagraphProperty::new().bidi(true);
    para
}

#[test]
fn test_parse_docx_bidi_paragraph() {
    // Build a DOCX with a bidi paragraph containing Arabic text
    let para = make_bidi_paragraph("مرحبا بالعالم");
    let data = build_docx_bytes(vec![para]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let para_block = flow.content.iter().find_map(|b| match b {
        Block::Paragraph(p) => Some(p),
        _ => None,
    });
    let p = para_block.expect("Should have a paragraph");
    assert_eq!(
        p.style.direction,
        Some(TextDirection::Rtl),
        "bidi paragraph should have RTL direction"
    );
}

#[test]
fn test_parse_docx_no_bidi_paragraph() {
    // Normal LTR paragraph should have direction: None
    let para = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello World"));
    let data = build_docx_bytes(vec![para]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let para_block = flow.content.iter().find_map(|b| match b {
        Block::Paragraph(p) => Some(p),
        _ => None,
    });
    let p = para_block.expect("Should have a paragraph");
    assert!(
        p.style.direction.is_none(),
        "Non-bidi paragraph should have no direction"
    );
}

#[test]
fn test_parse_docx_mixed_bidi_paragraphs() {
    // Mixed: first paragraph is RTL Arabic, second is LTR English
    let para_rtl = make_bidi_paragraph("مرحبا 123");
    let para_ltr = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello World"));
    let data = build_docx_bytes(vec![para_rtl, para_ltr]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    let paras: Vec<&Paragraph> = flow
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert!(paras.len() >= 2, "Should have at least 2 paragraphs");
    assert_eq!(
        paras[0].style.direction,
        Some(TextDirection::Rtl),
        "First paragraph (Arabic) should be RTL"
    );
    assert!(
        paras[1].style.direction.is_none(),
        "Second paragraph (English) should have no direction"
    );
}

#[test]
fn test_resolve_highlight_color_named_colors() {
    assert_eq!(
        resolve_highlight_color("yellow"),
        Some(Color::new(255, 255, 0))
    );
    assert_eq!(
        resolve_highlight_color("green"),
        Some(Color::new(0, 255, 0))
    );
    assert_eq!(
        resolve_highlight_color("cyan"),
        Some(Color::new(0, 255, 255))
    );
    assert_eq!(resolve_highlight_color("red"), Some(Color::new(255, 0, 0)));
    assert_eq!(
        resolve_highlight_color("darkBlue"),
        Some(Color::new(0, 0, 128))
    );
    assert_eq!(resolve_highlight_color("black"), Some(Color::new(0, 0, 0)));
    assert_eq!(
        resolve_highlight_color("white"),
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(resolve_highlight_color("none"), None);
    assert_eq!(resolve_highlight_color("unknown"), None);
}

#[test]
fn test_highlight_parsing_from_docx() {
    let para = docx_rs::Paragraph::new().add_run(
        docx_rs::Run::new()
            .add_text("Highlighted")
            .highlight("yellow"),
    );
    let data: Vec<u8> = build_docx_bytes(vec![para]);
    let (doc, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let pages: Vec<&FlowPage> = doc
        .pages
        .iter()
        .filter_map(|p| match p {
            Page::Flow(fp) => Some(fp),
            _ => None,
        })
        .collect();
    let runs: Vec<&Run> = pages
        .iter()
        .flat_map(|p| &p.content)
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(&p.runs),
            _ => None,
        })
        .flatten()
        .collect();
    let highlighted: Vec<&&Run> = runs
        .iter()
        .filter(|r| r.style.highlight.is_some())
        .collect();
    assert!(
        !highlighted.is_empty(),
        "Should have at least one run with highlight color"
    );
    assert_eq!(
        highlighted[0].style.highlight,
        Some(Color::new(255, 255, 0)),
        "Yellow highlight should map to (255, 255, 0)"
    );
}
