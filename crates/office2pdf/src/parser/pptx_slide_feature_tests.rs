use super::*;

// ── Slide background tests ───────────────────────────────────────────

#[test]
fn test_slide_solid_color_background() {
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let slide = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_slide_no_background() {
    let slide = make_empty_slide_xml();
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert!(page.background_color.is_none());
}

#[test]
fn test_slide_background_with_scheme_color() {
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let slide = make_slide_xml_with_bg(bg_xml, &[]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(0x44, 0x72, 0xC4)));
}

#[test]
fn test_slide_background_with_text_content() {
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let text_box = make_text_box(100000, 100000, 5000000, 500000, "Hello");
    let slide = make_slide_xml_with_bg(bg_xml, &[text_box]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(0, 0, 255)));
    assert_eq!(page.elements.len(), 1);
}

#[test]
fn test_slide_inherits_master_background() {
    let slide_xml = make_empty_slide_xml();
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(0, 255, 0)));
}

fn make_layout_xml(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sldLayout>");
    xml
}

fn make_master_xml(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sldMaster>");
    xml
}

// ── US-025: Slide master and layout inheritance tests ────────────────

#[test]
fn test_master_shape_appears_on_slide() {
    let slide_xml = make_empty_slide_xml();
    let layout_xml = make_layout_xml(&[]);
    let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Master Logo");
    let master_xml = make_master_xml(&[master_shape]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Master Logo");
}

#[test]
fn test_layout_shape_appears_on_slide() {
    let slide_xml = make_empty_slide_xml();
    let layout_shape = make_text_box(100_000, 100_000, 3_000_000, 500_000, "Layout Title");
    let layout_xml = make_layout_xml(&[layout_shape]);
    let master_xml = make_master_xml(&[]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Layout Title");
}

#[test]
fn test_inheritance_element_ordering() {
    let slide_shape = make_text_box(0, 0, 1_000_000, 500_000, "Slide Content");
    let slide_xml = make_slide_xml(&[slide_shape]);
    let layout_shape = make_text_box(0, 0, 1_000_000, 500_000, "Layout Content");
    let layout_xml = make_layout_xml(&[layout_shape]);
    let master_shape = make_text_box(0, 0, 1_000_000, 500_000, "Master Content");
    let master_xml = make_master_xml(&[master_shape]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 3);

    let master_blocks = text_box_blocks(&page.elements[0]);
    let master_para = match &master_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(master_para.runs[0].text, "Master Content");

    let layout_blocks = text_box_blocks(&page.elements[1]);
    let layout_para = match &layout_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(layout_para.runs[0].text, "Layout Content");

    let slide_blocks = text_box_blocks(&page.elements[2]);
    let slide_para = match &slide_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(slide_para.runs[0].text, "Slide Content");
}

#[test]
fn test_master_elements_appear_on_all_slides() {
    let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Company Logo");
    let master_xml = make_master_xml(&[master_shape]);
    let layout_xml = make_layout_xml(&[]);

    let slide1_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 1");
    let slide1_xml = make_slide_xml(&[slide1_shape]);
    let slide2_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 2");
    let slide2_xml = make_slide_xml(&[slide2_shape]);

    let data = build_test_pptx_with_layout_master_multi_slide(
        SLIDE_CX,
        SLIDE_CY,
        &[slide1_xml, slide2_xml],
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2);

    for (i, page) in doc.pages.iter().enumerate() {
        let fixed_page = match page {
            Page::Fixed(p) => p,
            _ => panic!("Expected FixedPage"),
        };
        assert_eq!(
            fixed_page.elements.len(),
            2,
            "Slide {} should have 2 elements (master + slide)",
            i + 1
        );

        let master_blocks = text_box_blocks(&fixed_page.elements[0]);
        let master_para = match &master_blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(master_para.runs[0].text, "Company Logo");
    }
}

#[test]
fn test_slide_without_layout_master_has_only_slide_elements() {
    let shape = make_text_box(0, 0, 1_000_000, 500_000, "Just Slide");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Just Slide");
}

#[test]
fn test_slide_inherits_layout_background_over_master() {
    let slide_xml = make_empty_slide_xml();
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(255, 0, 255)));
}

// ----- US-029: Slide selection tests -----

#[test]
fn test_slide_filter_single_slide() {
    use crate::config::SlideRange;

    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(2, 2)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 1, "Should only include slide 2");
    let page = first_fixed_page(&doc);
    let text = match &page.elements[0].kind {
        FixedElementKind::TextBox(text_box) => match &text_box.content[0] {
            Block::Paragraph(p) => p.runs[0].text.clone(),
            _ => panic!("Expected Paragraph"),
        },
        _ => panic!("Expected TextBox"),
    };
    assert_eq!(text, "Slide 2");
}

#[test]
fn test_slide_filter_range() {
    use crate::config::SlideRange;

    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
    let slide4 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 4")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3, slide4]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(2, 3)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 2, "Should include slides 2 and 3");
}

#[test]
fn test_slide_filter_none_includes_all() {
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: None,
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 2, "None should include all slides");
}

#[test]
fn test_slide_filter_range_beyond_total() {
    use crate::config::SlideRange;

    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(5, 10)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(
        doc.pages.len(),
        0,
        "Range beyond total slides should produce empty document"
    );
}
