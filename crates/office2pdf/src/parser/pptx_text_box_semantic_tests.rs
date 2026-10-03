use super::*;
use crate::ir::{TabAlignment, TabLeader, TabStop};

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
fn test_text_box_inherited_bullet_ignores_run_and_paragraph_mark_underline() {
    let shape = String::from(concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="457200" indent="-228600">"#,
        r#"<a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1400"><a:latin typeface="Arial"/></a:defRPr>"#,
        r#"</a:lvl1pPr></a:lstStyle><a:p><a:pPr lvl="0"/>"#,
        r#"<a:r><a:rPr lang="en-US" u="sng"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>www.commoncore.org</a:t></a:r>"#,
        r#"<a:endParaRPr lang="en-US" u="sng"/></a:p></p:txBody></p:sp>"#,
    ));
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
    assert_eq!(paragraph.runs[0].style.underline, Some(true));
    assert!(
        list.level_styles[&0]
            .marker_style
            .as_ref()
            .is_none_or(|style| style.underline != Some(true)),
        "A run or endParaRPr underline must not decorate an inherited marker"
    );
}

#[test]
fn test_text_box_inherited_bullet_keeps_list_level_underline() {
    let shape = String::from(concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="457200" indent="-228600">"#,
        r#"<a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1400" u="sng"><a:latin typeface="Arial"/></a:defRPr>"#,
        r#"</a:lvl1pPr></a:lstStyle><a:p><a:pPr lvl="0"/>"#,
        r#"<a:r><a:rPr lang="en-US" u="none"/><a:t>Plain text</a:t></a:r>"#,
        r#"<a:endParaRPr lang="en-US" u="none"/></a:p></p:txBody></p:sp>"#,
    ));
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
    assert_eq!(paragraph.runs[0].style.underline, Some(false));
    assert_eq!(
        list.level_styles[&0]
            .marker_style
            .as_ref()
            .and_then(|style| style.underline),
        Some(true),
        "The list level's own underline must remain on its marker"
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
fn test_text_box_paragraph_tab_stops_are_extracted() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr defTabSz="914400"><a:tabLst>"#,
        r#"<a:tab pos="914400" algn="l"/>"#,
        r#"<a:tab pos="1828800" algn="ctr"/>"#,
        r#"<a:tab pos="2743200" algn="r"/>"#,
        r#"<a:tab pos="3657600" algn="dec"/>"#,
        r#"</a:tabLst></a:pPr><a:r><a:t>Label&#9;Value</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 6_000_000, 500_000, paragraphs_xml);
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

    assert_eq!(
        paragraph.style.tab_stops,
        Some(vec![
            TabStop {
                position: 72.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            },
            TabStop {
                position: 144.0,
                alignment: TabAlignment::Center,
                leader: TabLeader::None,
            },
            TabStop {
                position: 216.0,
                alignment: TabAlignment::Right,
                leader: TabLeader::None,
            },
            TabStop {
                position: 288.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            },
        ])
    );
    assert_eq!(paragraph.style.default_tab_stop_pt, Some(72.0));
}

#[test]
fn test_text_box_paragraph_inherits_tab_stops_from_list_style() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr defTabSz="914400"><a:tabLst><a:tab pos="914400" algn="r"/></a:tabLst></a:lvl1pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:t>Label&#9;Value</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
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

    assert_eq!(
        paragraph.style.tab_stops,
        Some(vec![TabStop {
            position: 72.0,
            alignment: TabAlignment::Right,
            leader: TabLeader::None,
        }])
    );
    assert_eq!(paragraph.style.default_tab_stop_pt, Some(72.0));
}

#[test]
fn test_custom_geo_page_46_tabs_are_relative_to_the_inner_text_origin() {
    let data = include_bytes!("../../../../tests/fixtures/pptx/customGeo.pptx");
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[45] {
        Page::Fixed(page) => page,
        other => panic!("Expected fixed page, got {other:?}"),
    };
    // Located by content rather than by index: the page's element order shifts
    // whenever a shape gains a background element, and this test is about the
    // body's tab stops, not about where it sits in the vector.
    let blocks = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::TextBox(_)))
        .map(text_box_blocks)
        .find(|blocks| blocks.len() == 3)
        .expect("the three-paragraph body text box");

    assert_eq!(blocks.len(), 3);
    for block in blocks {
        let paragraph = match block {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected paragraph, got {other:?}"),
        };
        let tab_stops = paragraph
            .style
            .tab_stops
            .as_deref()
            .expect("Expected explicit tab stop");
        assert_eq!(tab_stops.len(), 1);
        assert!((tab_stops[0].position - 252.0).abs() < 0.001);
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
    // `auto_fit` drives text scaling, and `<a:spAutoFit/>` does not ask for
    // any: it resizes the shape to the text and leaves the run's declared
    // size alone (ECMA-376 §21.1.2.1.2). This assertion used to require the
    // opposite, which is what shrank an 8pt label to 4.9pt in issue #898.
    assert!(!text_box.auto_fit);
    assert!(!text_box.no_wrap);
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
fn test_text_box_auto_numbered_changed_start_begins_a_new_sequence() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr><a:r><a:t>Third</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod" startAt="8"/></a:pPr><a:r><a:t>Eighth</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(
        blocks.len(),
        2,
        "a changed start value must restart the list"
    );
    let starts: Vec<Option<u32>> = blocks
        .iter()
        .map(|block| match block {
            Block::List(list) => list.items[0].start_at,
            other => panic!("Expected List block, got {other:?}"),
        })
        .collect();
    assert_eq!(starts, vec![Some(3), Some(8)]);
}

#[test]
fn test_text_box_auto_numbered_repeated_start_continues_the_sequence() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr><a:r><a:t>Third</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr><a:r><a:t>Fourth</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(
        blocks.len(),
        1,
        "identical consecutive start metadata must not restart the list"
    );
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].start_at, Some(3));
    assert_eq!(list.items[1].start_at, None);
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
fn test_text_box_blank_soft_break_keeps_break_and_paragraph_mark_sizes() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:r><a:rPr sz="900"/><a:t>Caption</a:t></a:r></a:p>"#,
        r#"<a:p><a:endParaRPr sz="1000"/></a:p>"#,
        r#"<a:p><a:br><a:rPr sz="1100"/></a:br><a:endParaRPr sz="1000"/></a:p>"#,
        r#"<a:p><a:r><a:rPr sz="1100"/><a:t>Row 1</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let leading_blank_paragraph = match &blocks[1] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected the leading blank paragraph, got {other:?}"),
    };
    assert!(leading_blank_paragraph.runs.is_empty());
    assert_eq!(
        leading_blank_paragraph
            .style
            .paragraph_mark_text_style
            .as_deref()
            .and_then(|style| style.font_size),
        Some(10.0),
        "a paragraph with no runs uses its declared paragraph-mark size"
    );

    let blank_paragraph = match &blocks[2] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected the blank paragraph, got {other:?}"),
    };
    assert_eq!(blank_paragraph.runs.len(), 1);
    assert_eq!(blank_paragraph.runs[0].text, "\u{000B}");
    assert_eq!(
        blank_paragraph.runs[0].style.font_size,
        Some(11.0),
        "the empty line uses the size declared on its <a:br>"
    );
    assert_eq!(
        blank_paragraph
            .style
            .paragraph_mark_text_style
            .as_deref()
            .and_then(|style| style.font_size),
        Some(10.0),
        "the final empty line uses the size declared on <a:endParaRPr>"
    );
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
fn test_text_box_empty_bullet_paragraphs_keep_blank_lines_without_overriding_marker_style() {
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
    assert_eq!(
        list.items.len(),
        3,
        "Blank paragraphs must keep their lines"
    );
    assert_eq!(list.items[0].level, 1);
    assert!(pptx_paragraph_is_blank_list_item(&list.items[1].content[0]));
    assert!(pptx_paragraph_is_blank_list_item(&list.items[2].content[0]));
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
fn test_text_box_empty_numbered_paragraph_does_not_split_or_restart_sequence() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr><a:endParaRPr sz="1400"/></a:p>"#,
        r#"<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 1, "A blank item must not split numbering");
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.items.len(), 3);
    assert_eq!(list.items[0].start_at, Some(5));
    assert!(pptx_paragraph_is_blank_list_item(&list.items[1].content[0]));
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.marker_style.as_ref())
            .and_then(|style| style.font_size),
        Some(18.0),
        "The blank paragraph mark must not replace the visible number style"
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
fn test_text_box_font_ref_default_color_is_not_overridden_by_run_line_fill() {
    let slide_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="OrgBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="900000"/></a:xfrm>"#,
        r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></p:spPr>"#,
        r#"<p:style>"#,
        r#"<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>"#,
        r#"<a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef>"#,
        r#"<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>"#,
        r#"<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>"#,
        r#"</p:style>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#,
        r#"<a:p><a:pPr algn="ctr"/>"#,
        r#"<a:r><a:rPr lang="ko-KR">"#,
        r#"<a:ln><a:solidFill><a:sysClr val="window" lastClr="FFFFFF"><a:alpha val="0"/></a:sysClr></a:solidFill></a:ln>"#,
        r#"<a:latin typeface="Pretendard"/><a:ea typeface="Pretendard"/></a:rPr><a:t>이동욱 이사</a:t></a:r>"#,
        r#"</a:p></p:txBody></p:sp>"#,
    );
    let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Pretendard", "Pretendard");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    assert_eq!(paragraph.runs.len(), 1);
    assert_eq!(paragraph.runs[0].text, "이동욱 이사");
    assert_eq!(
        paragraph.runs[0].style.color,
        Some(Color::new(0x00, 0x00, 0x00)),
        "Run line fill should not overwrite the fontRef default text color"
    );
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
fn test_text_box_auto_numbered_list_continues_when_run_has_explicit_default_attrs() {
    // Reproduces a bug where three consecutive `arabicParenR` paragraphs
    // were split into two lists (1,2 then 1) because the third paragraph's
    // text runs had explicit default attributes (b="0", i="0", solidFill
    // black, etc.) that made the resolved marker style differ structurally
    // from the first two paragraphs even though the visual appearance was
    // identical.
    let paragraphs_xml = concat!(
        // Paragraph 1: simple run properties
        r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
        r#"<a:buAutoNum type="arabicParenR"/>"#,
        r#"</a:pPr>"#,
        r#"<a:r><a:rPr lang="ko-KR" sz="2000"><a:latin typeface="Pretendard Medium"/></a:rPr><a:t>First</a:t></a:r>"#,
        r#"</a:p>"#,
        // Paragraph 2: simple run properties
        r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
        r#"<a:buAutoNum type="arabicParenR"/>"#,
        r#"</a:pPr>"#,
        r#"<a:r><a:rPr lang="ko-KR" sz="2000"><a:latin typeface="Pretendard Medium"/></a:rPr><a:t>Second</a:t></a:r>"#,
        r#"</a:p>"#,
        // Paragraph 3: explicit default attributes (b="0", i="0", solidFill black, etc.)
        r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
        r#"<a:buAutoNum type="arabicParenR"/>"#,
        r#"</a:pPr>"#,
        r#"<a:r><a:rPr lang="en-US" sz="2000" b="0" i="0" u="none" strike="noStrike">"#,
        r#"<a:solidFill><a:prstClr val="black"/></a:solidFill>"#,
        r#"<a:latin typeface="Pretendard Medium"/>"#,
        r#"</a:rPr><a:t>Third</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 5_000_000, 3_000_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(
        blocks.len(),
        1,
        "All three paragraphs should be grouped into a single list block, got {blocks:?}"
    );

    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Ordered);
    assert_eq!(list.items.len(), 3, "List should have 3 items");
    assert_eq!(list.items[0].content[0].runs[0].text, "First");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second");
    assert_eq!(list.items[2].content[0].runs[0].text, "Third");
}
