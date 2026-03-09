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
