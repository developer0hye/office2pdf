use super::*;

#[test]
fn test_text_box_extraction() {
    let shape = make_text_box(0, 0, 1_000_000, 500_000, "Hello World");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 element");

    let blocks = text_box_blocks(&page.elements[0]);
    assert!(!blocks.is_empty(), "Expected at least one block");

    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 1);
    assert_eq!(para.runs[0].text, "Hello World");
}

#[test]
fn test_text_box_position_and_size() {
    let x = 1_000_000i64;
    let y = 500_000i64;
    let cx = 5_000_000i64;
    let cy = 2_000_000i64;
    let shape = make_text_box(x, y, cx, cy, "Positioned");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];

    let expected_x = x as f64 / 12700.0;
    let expected_y = y as f64 / 12700.0;
    let expected_w = cx as f64 / 12700.0;
    let expected_h = cy as f64 / 12700.0;

    assert!(
        (elem.x - expected_x).abs() < 0.1,
        "Expected x ~{expected_x}, got {}",
        elem.x
    );
    assert!(
        (elem.y - expected_y).abs() < 0.1,
        "Expected y ~{expected_y}, got {}",
        elem.y
    );
    assert!(
        (elem.width - expected_w).abs() < 0.1,
        "Expected width ~{expected_w}, got {}",
        elem.width
    );
    assert!(
        (elem.height - expected_h).abs() < 0.1,
        "Expected height ~{expected_h}, got {}",
        elem.height
    );
}

#[test]
fn test_text_box_bold_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Bold text");
    assert_eq!(para.runs[0].style.bold, Some(true));
}

#[test]
fn test_text_box_italic_formatting() {
    let runs_xml = r#"<a:r><a:rPr i="1"/><a:t>Italic text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Italic text");
    assert_eq!(para.runs[0].style.italic, Some(true));
}

#[test]
fn test_text_box_font_size() {
    let runs_xml = r#"<a:r><a:rPr sz="2400"/><a:t>Large text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].style.font_size, Some(24.0));
}

#[test]
fn test_text_box_combined_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Styled text");
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.italic, Some(true));
    assert_eq!(run.style.underline, Some(true));
    assert_eq!(run.style.strikethrough, Some(true));
    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    assert_eq!(run.style.font_family, Some("Arial".to_string()));
}

#[test]
fn test_multiple_text_boxes() {
    let shape1 = make_text_box(100_000, 100_000, 2_000_000, 500_000, "Box 1");
    let shape2 = make_text_box(100_000, 700_000, 2_000_000, 500_000, "Box 2");
    let slide = make_slide_xml(&[shape1, shape2]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 text boxes");

    let get_text = |elem: &FixedElement| -> String {
        let blocks = text_box_blocks(elem);
        blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect()
    };
    assert_eq!(get_text(&page.elements[0]), "Box 1");
    assert_eq!(get_text(&page.elements[1]), "Box 2");
}

#[test]
fn test_multiple_slides() {
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 3")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 3, "Expected 3 pages");
    for page in &doc.pages {
        assert!(matches!(page, Page::Fixed(_)));
    }
}

#[test]
fn test_text_box_multiple_paragraphs() {
    let paras_xml = r#"<a:p><a:r><a:rPr/><a:t>Paragraph 1</a:t></a:r></a:p><a:p><a:r><a:rPr/><a:t>Paragraph 2</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 3_000_000, 2_000_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paras: Vec<&Paragraph> = blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert!(paras.len() >= 2, "Expected at least 2 paragraphs");
    assert_eq!(paras[0].runs[0].text, "Paragraph 1");
    assert_eq!(paras[1].runs[0].text, "Paragraph 2");
}

#[test]
fn test_text_box_multiple_runs() {
    let runs_xml =
        r#"<a:r><a:rPr b="1"/><a:t>Bold </a:t></a:r><a:r><a:rPr i="1"/><a:t>Italic</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Bold ");
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert_eq!(para.runs[1].text, "Italic");
    assert_eq!(para.runs[1].style.italic, Some(true));
}

#[test]
fn test_paragraph_alignment_center() {
    let paras_xml = r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr/><a:t>Centered</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 2_000_000, 500_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.style.alignment, Some(Alignment::Center));
}
