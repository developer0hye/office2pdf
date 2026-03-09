use super::*;

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
