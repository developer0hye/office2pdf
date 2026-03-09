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
fn test_text_box_auto_numbered_paragraphs_group_into_list() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Ordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("1.")
    );
    assert_eq!(list.items[0].content[0].runs[0].text, "First");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second");
}

#[test]
fn test_text_box_bulleted_paragraphs_group_into_list() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>Second bullet</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Unordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
}

#[test]
fn test_text_box_bulleted_paragraph_preserves_char_marker_and_uses_run_style() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buFontTx/><a:buChar char="-"/></a:pPr>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Pretendard"/></a:rPr><a:t>First bullet</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.marker_text.as_deref(), Some("-"));
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Pretendard")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(14.0)
    );
    assert_eq!(
        style.marker_style.as_ref().and_then(|style| style.color),
        Some(Color::new(0x11, 0x22, 0x33))
    );
}

#[test]
fn test_text_box_bulleted_paragraph_preserves_explicit_marker_font() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Symbol bullet</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.marker_text.as_deref(), Some("è"));
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Wingdings")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(14.0)
    );
}

#[test]
fn test_text_box_paragraph_line_spacing_pct_extracted() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Proportional(factor)) => {
            assert!((factor - 1.5).abs() < f64::EPSILON);
        }
        other => panic!("Expected proportional line spacing, got {other:?}"),
    }
}

#[test]
fn test_text_box_body_pr_defaults_and_center_anchor_extracted() {
    let shape = make_text_box_with_body_pr(
        0,
        0,
        1_000_000,
        500_000,
        r#"<a:bodyPr anchor="ctr"><a:spAutoFit/></a:bodyPr>"#,
        "Centered",
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let text_box = match &page.elements[0].kind {
        FixedElementKind::TextBox(text_box) => text_box,
        other => panic!("Expected TextBox, got {other:?}"),
    };
    assert!((text_box.padding.left - 7.2).abs() < 0.001);
    assert!((text_box.padding.right - 7.2).abs() < 0.001);
    assert!((text_box.padding.top - 3.6).abs() < 0.001);
    assert!((text_box.padding.bottom - 3.6).abs() < 0.001);
    assert_eq!(
        text_box.vertical_align,
        crate::ir::TextBoxVerticalAlign::Center
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_start_override_sets_list_start() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod" startAt="3"/></a:pPr><a:r><a:t>Gamma</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod"/></a:pPr><a:r><a:t>Delta</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Ordered);
    assert_eq!(list.items[0].start_at, Some(3));
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("A.")
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_extracts_hanging_indent() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };

    let paragraph = &list.items[0].content[0];
    assert_eq!(paragraph.style.indent_left, Some(36.0));
    assert_eq!(paragraph.style.indent_first_line, Some(-36.0));
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("1)")
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_resolves_marker_style_from_text() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
        r#"<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buAutoNum type="arabicParenR"/>"#,
        r#"</a:pPr>"#,
        r#"<a:r><a:rPr lang="ko-KR" sz="2000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Pretendard Medium"/></a:rPr><a:t>First</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.numbering_pattern.as_deref(), Some("1)"));
    assert_eq!(style.marker_text, None);
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Pretendard Medium")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(20.0)
    );
    assert_eq!(
        style.marker_style.as_ref().and_then(|style| style.color),
        Some(Color::black())
    );
}

#[test]
fn test_text_box_paragraph_preserves_soft_line_breaks() {
    let paragraphs_xml = concat!(
        r#"<a:p>"#,
        r#"<a:r><a:t>Line 1</a:t></a:r>"#,
        r#"<a:br/>"#,
        r#"<a:r><a:t>Line 2</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "Line 1\u{000B}Line 2");
}

#[test]
fn test_text_box_plain_paragraph_between_bullets_breaks_list_sequence() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>1) First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>2) Second bullet</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 3, "Expected list / paragraph / list split");
    match &blocks[0] {
        Block::List(list) => {
            assert_eq!(list.items.len(), 1);
            assert_eq!(
                list.level_styles
                    .get(&1)
                    .and_then(|style| style.marker_text.as_deref()),
                Some("-")
            );
        }
        other => panic!("Expected first block to be a list, got {other:?}"),
    }
    match &blocks[1] {
        Block::Paragraph(paragraph) => {
            let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
            assert_eq!(text, "-> Continuation paragraph");
        }
        other => panic!("Expected middle block to be a paragraph, got {other:?}"),
    }
    match &blocks[2] {
        Block::List(list) => {
            assert_eq!(list.items.len(), 1);
            assert_eq!(
                list.level_styles
                    .get(&1)
                    .and_then(|style| style.marker_text.as_deref()),
                Some("-")
            );
        }
        other => panic!("Expected last block to be a list, got {other:?}"),
    }
}

#[test]
fn test_text_box_plain_paragraph_preserves_leading_arrow_text() {
    let paragraphs_xml = r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "-> Continuation paragraph");
}

#[test]
fn test_text_box_plain_paragraph_preserves_escaped_gt_entity() {
    let paragraphs_xml = r#"<a:p><a:r><a:t>-&gt; Continuation paragraph</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "-> Continuation paragraph");
}

#[test]
fn test_text_box_trailing_empty_bullets_do_not_override_nested_marker_style() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Arrow bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.items.len(), 1);
    assert_eq!(list.items[0].level, 1);
    assert_eq!(
        list.level_styles
            .get(&1)
            .and_then(|style| style.marker_text.as_deref()),
        Some("è")
    );
    assert!(
        !list.level_styles.contains_key(&0),
        "Trailing empty dash bullets should not create a level-0 marker style"
    );
}

#[test]
fn test_text_box_lst_style_default_run_props_are_applied_to_runs() {
    let shape = String::from(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="032543"/></a:solidFill><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p></p:txBody></p:sp>"#,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let run = &paragraph.runs[0];
    assert_eq!(
        run.style.font_family.as_deref(),
        Some("Pretendard SemiBold")
    );
    assert_eq!(run.style.font_size, Some(14.0));
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.color, Some(Color::new(0x03, 0x25, 0x43)));
}

#[test]
fn test_non_placeholder_shape_inherits_master_other_style_run_defaults() {
    let slide_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Caption"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#,
        r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>신</a:t></a:r><a:r><a:rPr lang="ko-KR" sz="1800"/><a:t>형</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp>"#,
    );
    let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="224466"/></a:solidFill><a:latin typeface="Pretendard"/><a:ea typeface="Pretendard"/><a:cs typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "신형");
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.font_size == Some(18.0))
    );
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.font_family.as_deref() == Some("Pretendard"))
    );
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.color == Some(Color::new(0x22, 0x44, 0x66)))
    );
}

#[test]
fn test_text_box_lst_style_overrides_master_other_style_run_defaults() {
    let slide_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle>"#,
        r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp>"#,
    );
    let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:latin typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    assert_eq!(paragraph.runs[0].style.font_size, Some(24.0));
    assert_eq!(
        paragraph.runs[0].style.font_family.as_deref(),
        Some("Pretendard SemiBold")
    );
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
