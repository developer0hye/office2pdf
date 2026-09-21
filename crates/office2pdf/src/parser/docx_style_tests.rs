use super::*;

#[test]
fn test_heading1_style_applies_defaults() {
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

    assert_eq!(run.style.font_size, Some(20.0));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_custom_heading1_and_linked_character_typography_do_not_synthesize_bold() {
    // Reduced from `Place your event title here.docx`: the paragraph style
    // states its display size, and the linked character style states the face
    // and size. Word treats the omitted w:b as regular rather than layering
    // a synthesized Heading 1 bold default on top (issue #1457).
    let styles_xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
        <w:rPr><w:sz w:val="132"/><w:szCs w:val="132"/></w:rPr>
      </w:style>
      <w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
        <w:name w:val="Heading 1 Char"/>
        <w:link w:val="Heading1"/>
        <w:rPr>
          <w:rFonts w:ascii="The Hand Black" w:hAnsi="The Hand Black"/>
          <w:sz w:val="132"/><w:szCs w:val="132"/>
        </w:rPr>
      </w:style>
    </w:styles>"#;
    let body_xml = r#"<w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="Heading1Char"/></w:rPr><w:t>ALL AGES WELCOME</w:t></w:r>
    </w:p>"#;

    let data = super::page_feature_tests::build_docx_with_raw_styles(styles_xml, body_xml);
    let (doc, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("custom Heading1 package parses");
    let run = first_run(&doc);

    assert_eq!(run.style.font_family.as_deref(), Some("The Hand Black"));
    assert_eq!(run.style.font_size, Some(66.0));
    assert_eq!(run.style.bold, None);
}

#[test]
fn test_custom_heading2_typography_without_w_b_remains_regular() {
    // The same fixture defines Heading2's face and size directly but omits
    // w:b. That is a complete document-defined display treatment, not a bare
    // unformatted heading definition that needs synthesized defaults (issue #1457).
    let styles_xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="Avenir Next W1G Medium" w:hAnsi="Avenir Next W1G Medium"/>
          <w:sz w:val="28"/><w:szCs w:val="28"/>
        </w:rPr>
      </w:style>
    </w:styles>"#;
    let body_xml = r#"<w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>GATES OPEN AT 9AM</w:t></w:r>
    </w:p>"#;

    let data = super::page_feature_tests::build_docx_with_raw_styles(styles_xml, body_xml);
    let (doc, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("custom Heading2 package parses");
    let run = first_run(&doc);

    assert_eq!(
        run.style.font_family.as_deref(),
        Some("Avenir Next W1G Medium")
    );
    assert_eq!(run.style.font_size, Some(14.0));
    assert_eq!(run.style.bold, None);
}

#[test]
fn test_derived_heading_typography_without_own_outline_level_remains_regular() {
    // A custom style may inherit Heading1's outline level while supplying its
    // own text treatment. The child does not repeat w:outlineLvl, but it is
    // still a document-defined heading rather than a bare, unformatted one
    // (issue #1457).
    let styles_xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:customStyle="1" w:styleId="EventHeading">
        <w:name w:val="Event Heading"/>
        <w:basedOn w:val="Heading1"/>
        <w:rPr>
          <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
          <w:sz w:val="36"/><w:szCs w:val="36"/>
        </w:rPr>
      </w:style>
    </w:styles>"#;
    let body_xml = r#"<w:p>
      <w:pPr><w:pStyle w:val="EventHeading"/></w:pPr>
      <w:r><w:t>Derived event heading</w:t></w:r>
    </w:p>"#;

    let data = super::page_feature_tests::build_docx_with_raw_styles(styles_xml, body_xml);
    let (doc, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("derived heading package parses");
    let run = first_run(&doc);

    assert_eq!(run.style.font_family.as_deref(), Some("Georgia"));
    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.bold, None);
}

#[test]
fn test_heading3_through_6_defaults() {
    let expected: Vec<(usize, &str, f64)> = vec![
        (2, "Heading3", 16.0),
        (3, "Heading4", 14.0),
        (4, "Heading5", 12.0),
        (5, "Heading6", 11.0),
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
    let custom = docx_rs::Style::new("CustomStyle", docx_rs::StyleType::Paragraph)
        .name("Custom Style")
        .size(36)
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
    let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
        .name("Heading 1")
        .outline_lvl(0);

    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Small heading").size(20))
                .style("Heading1"),
        ],
        vec![h1_style],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    assert_eq!(run.style.font_size, Some(10.0));
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
fn test_based_on_chain_inherits_parent_paragraph_and_run_properties() {
    let child = docx_rs::Style::new("Child", docx_rs::StyleType::Paragraph)
        .name("Child")
        .based_on("Middle")
        .bold();
    let parent = docx_rs::Style::new("Middle", docx_rs::StyleType::Paragraph)
        .name("Middle")
        .based_on("Normal")
        .size(36);
    let normal = docx_rs::Style::new("Normal", docx_rs::StyleType::Paragraph)
        .name("Normal")
        .align(docx_rs::AlignmentType::Center)
        .fonts(docx_rs::RunFonts::new().ascii("Georgia"));
    let paragraph = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Inherited"))
        .style("Child");

    // Deliberately put the child before its parents: style resolution must
    // follow IDs rather than relying on declaration order.
    let data = build_docx_bytes_with_styles(vec![paragraph], vec![child, parent, normal]);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraph = first_paragraph(&doc);
    let run = &paragraph.runs[0];

    assert_eq!(paragraph.style.alignment, Some(Alignment::Center));
    assert_eq!(run.style.font_family.as_deref(), Some("Georgia"));
    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_character_style_based_on_chain_inherits_run_properties() {
    let child = docx_rs::Style::new("ChildCharacter", docx_rs::StyleType::Character)
        .name("Child character")
        .based_on("ParentCharacter")
        .bold();
    let parent = docx_rs::Style::new("ParentCharacter", docx_rs::StyleType::Character)
        .name("Parent character")
        .fonts(docx_rs::RunFonts::new().ascii("Georgia"))
        .italic();
    let paragraph = docx_rs::Paragraph::new().add_run(
        docx_rs::Run::new()
            .add_text("Inherited character style")
            .style("ChildCharacter"),
    );

    // Character styles use the same child-to-root cascade as paragraph
    // styles, while remaining independent from paragraph defaults.
    let data = build_docx_bytes_with_styles(vec![paragraph], vec![child, parent]);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    assert_eq!(run.style.font_family.as_deref(), Some("Georgia"));
    assert_eq!(run.style.italic, Some(true));
    assert_eq!(run.style.bold, Some(true));
}

#[test]
fn test_child_and_direct_formatting_override_based_on_parent() {
    let parent = docx_rs::Style::new("Parent", docx_rs::StyleType::Paragraph)
        .name("Parent")
        .align(docx_rs::AlignmentType::Center)
        .fonts(docx_rs::RunFonts::new().ascii("Georgia"));
    let child = docx_rs::Style::new("Child", docx_rs::StyleType::Paragraph)
        .name("Child")
        .based_on("Parent")
        .align(docx_rs::AlignmentType::Right)
        .fonts(docx_rs::RunFonts::new().ascii("Arial"));
    let paragraph = docx_rs::Paragraph::new()
        .add_run(
            docx_rs::Run::new()
                .add_text("Overrides")
                .fonts(docx_rs::RunFonts::new().ascii("Courier New")),
        )
        .style("Child")
        .align(docx_rs::AlignmentType::Left);

    let data = build_docx_bytes_with_styles(vec![paragraph], vec![parent, child]);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraph = first_paragraph(&doc);

    assert_eq!(paragraph.style.alignment, Some(Alignment::Left));
    assert_eq!(
        paragraph.runs[0].style.font_family.as_deref(),
        Some("Courier New")
    );
}

#[test]
fn test_cyclic_based_on_chain_is_ignored_for_every_cycle_member() {
    let first = docx_rs::Style::new("First", docx_rs::StyleType::Paragraph)
        .name("First")
        .based_on("Second");
    let second = docx_rs::Style::new("Second", docx_rs::StyleType::Paragraph)
        .name("Second")
        .based_on("First")
        .align(docx_rs::AlignmentType::Center);
    let paragraphs = vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("First paragraph"))
            .style("First"),
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Second paragraph"))
            .style("Second"),
    ];

    let data = build_docx_bytes_with_styles(paragraphs, vec![first, second]);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let paragraphs: Vec<&Paragraph> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();

    assert_eq!(paragraphs[0].style.alignment, None);
    assert_eq!(paragraphs[1].style.alignment, Some(Alignment::Center));
}

#[test]
fn test_missing_or_cross_type_based_on_parent_is_ignored() {
    let missing_parent = docx_rs::Style::new("MissingChild", docx_rs::StyleType::Paragraph)
        .name("Missing child")
        .based_on("DoesNotExist")
        .bold();
    let character_parent = docx_rs::Style::new("CharacterParent", docx_rs::StyleType::Character)
        .name("Character parent")
        .align(docx_rs::AlignmentType::Center);
    let cross_type_child = docx_rs::Style::new("CrossTypeChild", docx_rs::StyleType::Paragraph)
        .name("Cross-type child")
        .based_on("CharacterParent")
        .italic();
    let paragraphs = vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Missing"))
            .style("MissingChild"),
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Cross type"))
            .style("CrossTypeChild"),
    ];

    let data = build_docx_bytes_with_styles(
        paragraphs,
        vec![missing_parent, character_parent, cross_type_child],
    );
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let paragraphs: Vec<&Paragraph> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();

    assert_eq!(paragraphs[0].style.alignment, None);
    assert_eq!(paragraphs[0].runs[0].style.bold, Some(true));
    assert_eq!(paragraphs[1].style.alignment, None);
    assert_eq!(paragraphs[1].runs[0].style.italic, Some(true));
}

#[test]
fn test_normal_style_no_heading_defaults() {
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

    assert!(run.style.font_size.is_none());
    assert!(run.style.bold.is_none());
}

#[test]
fn test_heading_with_mixed_paragraphs() {
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

    if let Block::Paragraph(p) = &blocks[0] {
        assert_eq!(p.runs[0].style.font_size, Some(24.0));
        assert_eq!(p.runs[0].style.bold, Some(true));
    } else {
        panic!("Expected Paragraph");
    }

    if let Block::Paragraph(p) = &blocks[1] {
        assert!(p.runs[0].style.font_size.is_none());
        assert!(p.runs[0].style.bold.is_none());
    } else {
        panic!("Expected Paragraph");
    }

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

#[test]
fn test_direct_jc_center_applied_to_paragraph() {
    // Direct <w:jc w:val="center"/> in the paragraph's own pPr (no style).
    let data = build_docx_bytes_with_styles(
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Centered directly"))
                .align(docx_rs::AlignmentType::Center),
        ],
        vec![],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    assert_eq!(para.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_default_paragraph_style_applies_without_pstyle() {
    // w:default="1" marks the style that paragraphs without an explicit
    // pStyle inherit (issue #288): its spacing must survive the cascade.
    let mut normal = docx_rs::Style::new("Normal", docx_rs::StyleType::Paragraph)
        .name("Normal")
        .size(24)
        .line_spacing(
            docx_rs::LineSpacing::new()
                .after(160)
                .line(360)
                .line_rule(docx_rs::LineSpacingType::Auto),
        );
    normal.default = true;

    let data = build_docx_bytes_with_styles(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("본문 문단"))],
        vec![normal],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let Block::Paragraph(paragraph) = flow
        .content
        .iter()
        .find(|b| matches!(b, Block::Paragraph(_)))
        .expect("paragraph")
    else {
        unreachable!()
    };
    assert_eq!(
        paragraph.style.space_after,
        Some(8.0),
        "default style spacing (160 twips = 8pt) must apply to pStyle-less paragraphs"
    );
    assert_eq!(paragraph.runs[0].style.font_size, Some(12.0));
}

#[test]
fn test_scan_default_paragraph_style_id_from_raw_styles_xml() {
    let xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="character" w:default="1" w:styleId="DefaultCharacter"/>
      <w:style w:type="paragraph" w:default="1" w:styleId="BodyDefault"/>
    </w:styles>"#;

    assert_eq!(
        styles::scan_default_paragraph_style_id(xml).as_deref(),
        Some("BodyDefault")
    );
}

#[test]
fn test_scan_defines_default_paragraph_style() {
    // The two shapes real packages declare their default paragraph style in:
    // Word marks it `w:default="1"`; docx-rs-built files write a style whose
    // id is `Normal` without the flag. Either replaces Word's built-in Normal,
    // which is what reactivates the East Asian auto space (issue #732).
    let by_default_flag = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Standard"/>
    </w:styles>"#;
    let by_normal_id = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    </w:styles>"#;
    // The corpus shape: named styles only, no default, no Normal. A character
    // style named Normal must not count — only a paragraph style replaces the
    // built-in paragraph default.
    let corpus_shaped = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="character" w:default="1" w:styleId="Normal"/>
      <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>
    </w:styles>"#;

    assert!(styles::scan_defines_default_paragraph_style(
        by_default_flag
    ));
    assert!(styles::scan_defines_default_paragraph_style(by_normal_id));
    assert!(!styles::scan_defines_default_paragraph_style(corpus_shaped));
}

#[test]
fn test_doc_default_theme_font_resolves_via_theme() {
    // docDefaults referencing asciiTheme="minorHAnsi" must resolve to the
    // theme's minor latin typeface instead of falling back to the renderer
    // default (issue #287). docx-rs's builder can't author theme slots, so
    // exercise the resolver directly.
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements><a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
    <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
  </a:fontScheme></a:themeElements>
</a:theme>"#;
    let theme = parse_theme_fonts(theme_xml);
    assert_eq!(theme.minor_latin.as_deref(), Some("Calibri"));
    assert_eq!(theme.major_latin.as_deref(), Some("Calibri Light"));

    let run_property = serde_json::json!({ "fonts": { "asciiTheme": "minorHAnsi" } });
    assert_eq!(
        resolve_theme_font_family(&run_property, &theme).as_deref(),
        Some("Calibri")
    );
    let heading_property = serde_json::json!({ "fonts": { "asciiTheme": "majorHAnsi" } });
    assert_eq!(
        resolve_theme_font_family(&heading_property, &theme).as_deref(),
        Some("Calibri Light")
    );
    let no_theme = serde_json::json!({ "fonts": { "ascii": "Arial" } });
    assert_eq!(resolve_theme_font_family(&no_theme, &theme), None);
}

#[test]
fn test_doc_default_falls_back_to_words_face_only_when_none_resolves() {
    // Triangulation for issue #1196 on a package of a different shape from the
    // fixture that measured it: raw `word/styles.xml`, no header, no footer,
    // no theme part. The fallback is what a package resolving nothing takes,
    // so a `w:docDefaults` that does name a face must keep it.
    const STYLES_TEMPLATE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>{FONTS}<w:sz w:val="20"/></w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"#;
    let body_xml = r#"<w:p><w:r><w:t>Quarterly revenue held flat.</w:t></w:r></w:p>"#;

    for (fonts_xml, expected) in [
        ("", "Times New Roman"),
        (
            r#"<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>"#,
            "Verdana",
        ),
    ] {
        let styles_xml = STYLES_TEMPLATE.replace("{FONTS}", fonts_xml);
        let data = super::page_feature_tests::build_docx_with_raw_styles(&styles_xml, body_xml);

        let (doc, _warnings) = DocxParser
            .parse(&data, &ConvertOptions::default())
            .expect("the package parses");

        assert_eq!(
            first_run(&doc).style.font_family.as_deref(),
            Some(expected),
            "a `w:rPrDefault` of {fonts_xml:?} must resolve {expected}"
        );
    }
}

#[test]
fn test_paragraph_shading_extracted_as_background() {
    // Word paints w:pPr/w:shd behind the whole paragraph (code blocks in
    // the CLI-manual fixture); the fill must reach the IR (issue #351).
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:shd w:val="clear" w:fill="F4F4F4"/></w:pPr>
      <w:r><w:t>$ cargo install office2pdf-cli</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>"#;
    let data = build_docx_with_columns(document_xml);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    assert_eq!(para.style.background, Some(Color::new(0xF4, 0xF4, 0xF4)));
}

#[test]
fn test_paragraph_bottom_border_extracted() {
    // w:pBdr bottom rules (resume header underline, letterhead frames) must
    // reach the IR with Word's eighth-point width unit (issue #368).
    let mut ruled = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("JAMIE PARKER"));
    ruled.property = ruled.property.set_borders(
        docx_rs::ParagraphBorders::with_empty().set(
            docx_rs::ParagraphBorder::new(docx_rs::ParagraphBorderPosition::Bottom)
                .val(docx_rs::BorderType::Single)
                .size(6)
                .color("1E2761"),
        ),
    );
    let data = build_docx_bytes(vec![ruled]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let para = first_paragraph(&doc);

    let border = para.style.border.as_ref().expect("border must be parsed");
    let bottom = border.bottom.as_ref().expect("bottom side present");
    assert_eq!(bottom.width, 0.75, "w:sz is eighths of a point");
    assert_eq!(bottom.color, Color::new(0x1E, 0x27, 0x61));
    assert_eq!(bottom.style, BorderLineStyle::Solid);
    assert!(border.top.is_none());
}

/// `w:docDefaults` carrying the justification, line spacing, and space-after a
/// generated document states once for its whole body.
const DOC_DEFAULT_STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="100" w:line="278" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:pPr><w:spacing w:after="150"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
</w:styles>"#;

fn doc_default_paragraph(document_xml: &str) -> Paragraph {
    let data = build_docx_with_styles_xml(document_xml, DOC_DEFAULT_STYLES_XML);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let Block::Paragraph(paragraph) = flow
        .content
        .iter()
        .find(|block| matches!(block, Block::Paragraph(_)))
        .expect("paragraph")
    else {
        unreachable!()
    };
    paragraph.clone()
}

fn proportional_line_spacing(style: &ParagraphStyle) -> Option<f64> {
    match style.line_spacing {
        Some(LineSpacing::Proportional(factor)) => Some(factor),
        _ => None,
    }
}

#[test]
fn test_doc_default_paragraph_properties_reach_a_paragraph_without_a_style() {
    // w:docDefaults/w:pPrDefault sits below every named style and below the
    // w:default="1" style. Reading only w:rPrDefault left a document that
    // states its body layout there ragged, single-spaced, and gapless, which
    // repaginated the technical brief from 39 pages to 31 (issue #574).
    let paragraph = doc_default_paragraph(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>본문 문단</w:t></w:r></w:p></w:body>
</w:document>"#,
    );

    assert_eq!(
        paragraph.style.alignment,
        Some(Alignment::Justify),
        "w:pPrDefault w:jc=both must reach a paragraph with no pStyle"
    );
    assert_eq!(
        paragraph.style.space_after,
        Some(5.0),
        "w:pPrDefault w:after=100 twips must reach a paragraph with no pStyle"
    );
    assert_eq!(
        proportional_line_spacing(&paragraph.style),
        Some(278.0 / 240.0),
        "w:pPrDefault w:line=278 auto must reach a paragraph with no pStyle"
    );
}

#[test]
fn test_named_style_overrides_doc_defaults_but_inherits_what_it_does_not_state() {
    // A named style states only what it changes. Heading1 here sets its own
    // alignment and space-after, so those win, but it states no line spacing
    // and must still inherit the document default's (issue #574).
    let paragraph = doc_default_paragraph(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1. 개요</w:t></w:r></w:p>
  </w:body>
</w:document>"#,
    );

    assert_eq!(
        paragraph.style.alignment,
        Some(Alignment::Left),
        "the style's own w:jc must win over the document default"
    );
    assert_eq!(
        paragraph.style.space_after,
        Some(7.5),
        "the style's own 150 twips must win over the document default's 100"
    );
    assert_eq!(
        proportional_line_spacing(&paragraph.style),
        Some(278.0 / 240.0),
        "the line spacing the style does not state must come from w:pPrDefault"
    );
}

#[test]
fn test_direct_paragraph_formatting_overrides_doc_defaults() {
    // Direct w:pPr sits above both the style and the document default.
    let paragraph = doc_default_paragraph(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:spacing w:after="240" w:line="480" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:t>가운데 문단</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#,
    );

    assert_eq!(paragraph.style.alignment, Some(Alignment::Center));
    assert_eq!(paragraph.style.space_after, Some(12.0));
    assert_eq!(proportional_line_spacing(&paragraph.style), Some(2.0));
}

#[test]
fn test_tracked_insertion_renders_and_tracked_deletion_does_not() {
    // Word's final view — "No Markup", and what accepting every revision
    // produces — keeps w:ins content and drops w:del content. Both were
    // falling through the paragraph child match, so an accepted insertion
    // vanished from the output along with the deletion (issue #583).
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">v0.6.4 기준 </w:t></w:r>
      <w:del w:id="901" w:author="문서 검토자" w:date="2026-07-26T10:00:00Z">
        <w:r><w:delText xml:space="preserve">기능 목록 재확인</w:delText></w:r>
      </w:del>
      <w:ins w:id="902" w:author="문서 검토자" w:date="2026-07-26T10:05:00Z">
        <w:r><w:t xml:space="preserve">기능·품질 지표 실측값 반영</w:t></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let Block::Paragraph(paragraph) = flow
        .content
        .iter()
        .find(|block| matches!(block, Block::Paragraph(_)))
        .expect("paragraph")
    else {
        unreachable!()
    };

    let text: String = paragraph
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect::<String>();

    assert!(
        text.contains("기능·품질 지표 실측값 반영"),
        "the tracked insertion is ordinary final-document text: {text:?}"
    );
    assert!(
        !text.contains("기능 목록 재확인"),
        "the tracked deletion is not in the final document: {text:?}"
    );
    assert!(
        text.starts_with("v0.6.4 기준"),
        "the untracked run keeps its place: {text:?}"
    );
}

#[test]
fn test_tracked_move_keeps_its_destination_and_drops_its_origin() {
    // A tracked move is both kinds of revision at once: Word's final view
    // shows the text where it was moved to (w:moveTo) and not where it came
    // from (w:moveFrom). docx-rs reads both as their own paragraph children,
    // so falling through the child match dropped the destination as well.
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:moveToRangeStart w:id="905" w:author="Reviewer" w:date="2026-09-13T10:00:00Z" w:name="move1"/>
      <w:moveTo w:id="906" w:author="Reviewer" w:date="2026-09-13T10:00:00Z">
        <w:r><w:t xml:space="preserve">Quarterly summary</w:t></w:r>
      </w:moveTo>
      <w:moveToRangeEnd w:id="905"/>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">Revenue grew. </w:t></w:r>
      <w:moveFromRangeStart w:id="907" w:author="Reviewer" w:date="2026-09-13T10:00:00Z" w:name="move1"/>
      <w:moveFrom w:id="908" w:author="Reviewer" w:date="2026-09-13T10:00:00Z">
        <w:r><w:t xml:space="preserve">Quarterly summary</w:t></w:r>
      </w:moveFrom>
      <w:moveFromRangeEnd w:id="907"/>
    </w:p>
  </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let paragraph_texts: Vec<String> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect();

    assert_eq!(
        paragraph_texts.len(),
        2,
        "both paragraphs remain: {paragraph_texts:?}"
    );
    assert_eq!(
        paragraph_texts[0], "Quarterly summary",
        "the move destination is final-document text"
    );
    assert_eq!(
        paragraph_texts[1].trim_end(),
        "Revenue grew.",
        "the move origin is not in the final document"
    );
}

#[test]
fn test_insertion_that_was_later_deleted_is_dropped() {
    // A w:del nested inside a w:ins is text that was inserted and then
    // deleted again, so the final document does not contain it (issue #583).
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">남는 문장</w:t></w:r>
      <w:ins w:id="903" w:author="검토자" w:date="2026-07-26T10:05:00Z">
        <w:del w:id="904" w:author="검토자" w:date="2026-07-26T10:06:00Z">
          <w:r><w:delText xml:space="preserve">되돌린 문장</w:delText></w:r>
        </w:del>
      </w:ins>
    </w:p>
  </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let Block::Paragraph(paragraph) = flow
        .content
        .iter()
        .find(|block| matches!(block, Block::Paragraph(_)))
        .expect("paragraph")
    else {
        unreachable!()
    };

    let text: String = paragraph
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect::<String>();

    assert_eq!(text, "남는 문장");
}

#[test]
fn test_tracked_changes_resolve_the_same_way_in_a_footer() {
    // A header or footer resolves to the final view like the body does; the
    // two loops share one flattening so they cannot drift (issue #583).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    // The brief's footers carry a PAGE and a NUMPAGES field, which travel the
    // same match as the tracked-change variants.
    let footer_fields: usize = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => flow.footer.as_ref(),
            _ => None,
        })
        .flat_map(|footer| footer.paragraphs.iter())
        .flat_map(|paragraph| paragraph.elements.iter())
        .filter(|element| {
            matches!(
                element,
                crate::ir::HFInline::PageNumber(_) | crate::ir::HFInline::TotalPages(_)
            )
        })
        .count();

    assert!(
        footer_fields >= 12,
        "every section's footer keeps its PAGE and NUMPAGES fields, got {footer_fields}"
    );
}

#[test]
fn test_header_and_footer_runs_inherit_the_document_default_run_style() {
    // Header and footer parts are read before the stylesheet is, so their runs
    // were left with only what they state themselves and fell through to the
    // renderer's own family and size. Word resolves them through the same run
    // cascade as the body (issue #578).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let Page::Flow(flow) = doc
        .pages
        .iter()
        .find(|page| matches!(page, Page::Flow(flow) if flow.header.is_some()))
        .expect("a section with a header")
    else {
        unreachable!()
    };
    let header = flow.header.as_ref().expect("header");

    // The brief's `w:rPrDefault` names Calibri at 10pt. The header's first run
    // states a colour, a weight, and 8pt, but no family.
    let styled_run = header
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .find_map(|element| match element {
            crate::ir::HFInline::Run(run) if !run.text.trim().is_empty() => Some(&run.style),
            _ => None,
        })
        .expect("a header text run");

    assert_eq!(
        styled_run.font_family.as_deref(),
        Some("Calibri"),
        "the family the run does not state comes from w:rPrDefault"
    );
    assert_eq!(
        styled_run.font_size,
        Some(8.0),
        "the size the run does state still wins"
    );

    let footer = flow.footer.as_ref().expect("footer");
    let page_number_style = footer
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .find_map(|element| match element {
            crate::ir::HFInline::PageNumber(style) => Some(style),
            _ => None,
        })
        .expect("a PAGE field");

    assert_eq!(
        page_number_style.font_family.as_deref(),
        Some("Calibri"),
        "a PAGE field resolves through the same cascade as the literals beside it"
    );
}

#[test]
fn test_footnote_text_resolves_its_paragraph_style_and_run_properties() {
    // A note is read before the stylesheet is, so its runs used to arrive as
    // one unstyled string and rendered at the engine's own footnote size and
    // face. Word resolves them through the same cascade as the body: the
    // note's w:pStyle supplies what the runs leave unstated (issue #580).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let note_runs: Vec<&Run> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow),
            _ => None,
        })
        .flat_map(|flow| flow.content.iter())
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| paragraph.runs.iter())
        .filter_map(|run| run.footnote.as_ref())
        .flatten()
        .collect();

    assert!(!note_runs.is_empty(), "the brief carries footnotes");

    // `FootnoteBody` sets `w:sz="16"` and `w:color="404040"`, and inherits the
    // document default family.
    for note_run in &note_runs {
        assert_eq!(
            note_run.style.font_size,
            Some(8.0),
            "the note's style sets 8pt: {:?}",
            note_run.text
        );
        assert_eq!(
            note_run
                .style
                .color
                .map(|color| (color.r, color.g, color.b)),
            Some((0x40, 0x40, 0x40)),
            "the note's style sets #404040: {:?}",
            note_run.text
        );
        assert_eq!(
            note_run.style.font_family.as_deref(),
            Some("Calibri"),
            "the family the note does not state comes from the document default"
        );
    }
}

#[test]
fn test_footnote_run_properties_override_the_notes_style() {
    // A run inside a note states its own properties over the note's style.
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
    <w:footnoteReference w:id="2"/></w:r></w:p></w:body>
</w:document>"#;
    let footnotes_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="2">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteBody"/></w:pPr>
      <w:r><w:t xml:space="preserve">보통 </w:t></w:r>
      <w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">강조</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="FootnoteBody">
    <w:name w:val="Footnote Body"/>
    <w:rPr><w:sz w:val="16"/></w:rPr>
  </w:style>
</w:styles>"#;

    let data = build_docx_with_notes_xml(document_xml, styles_xml, footnotes_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected flow page");
    };
    let note: &Vec<Run> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| paragraph.runs.iter())
        .find_map(|run| run.footnote.as_ref())
        .expect("a footnote");

    assert_eq!(note.len(), 2, "both runs survive: {note:?}");
    assert_eq!(note[0].text, "보통 ");
    assert_eq!(note[0].style.font_size, Some(8.0), "from the note's style");
    assert_eq!(note[1].text, "강조");
    assert_eq!(
        note[1].style.font_size,
        Some(12.0),
        "the run's own size wins"
    );
    assert_eq!(note[1].style.bold, Some(true), "the run's own weight wins");
}

#[test]
fn test_section_page_numbering_restarts_and_picks_its_numerals() {
    // A front matter that restarts at `i` states both the restart and the
    // format in w:pgNumType. office2pdf read neither and warned that it was
    // falling back to one global decimal counter (issue #582).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let numbering: Vec<Option<crate::ir::PageNumbering>> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow.page_numbering),
            _ => None,
        })
        .collect();

    // The cover states nothing, the front matter restarts at i, the body
    // restarts at 1 in decimal.
    assert_eq!(
        numbering[1],
        Some(crate::ir::PageNumbering {
            start: Some(1),
            format: crate::ir::PageNumberFormat::LowerRoman
        }),
        "the front matter restarts at i"
    );
    assert_eq!(
        numbering[2],
        Some(crate::ir::PageNumbering {
            start: Some(1),
            format: crate::ir::PageNumberFormat::Decimal
        }),
        "the body restarts at 1 in decimal"
    );

    assert!(
        !warnings.iter().any(|warning| matches!(
            warning,
            crate::error::ConvertWarning::FallbackUsed { from, .. }
                if from == "section page number restart"
        )),
        "the restart is no longer a fallback: {warnings:?}"
    );
}

#[test]
fn test_seq_fields_number_captions_in_document_order() {
    // Word keeps a caption's number in a SEQ field, not in its text. Dropping
    // the field left every caption in the brief without one (issue #577).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let caption_text = |lead: &str| -> Vec<String> {
        doc.pages
            .iter()
            .filter_map(|page| match page {
                Page::Flow(flow) => Some(flow),
                _ => None,
            })
            .flat_map(|flow| flow.content.iter())
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph),
                Block::Caption(caption) => Some(&caption.paragraph),
                _ => None,
            })
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>()
            })
            // A caption leads with the label and the field's number; the
            // contents-page headings share the label but not a number.
            .filter(|text| {
                text.strip_prefix(lead)
                    .is_some_and(|rest| rest.starts_with(|c: char| c.is_ascii_digit()))
            })
            .collect()
    };

    let tables = caption_text("표 ");
    let figures = caption_text("그림 ");
    assert_eq!(tables.len(), 33, "every table caption: {tables:?}");
    assert_eq!(figures.len(), 5, "every figure caption: {figures:?}");

    // Each identifier counts from one, independently, in document order.
    for (index, caption) in tables.iter().enumerate() {
        assert!(
            caption.starts_with(&format!("표 {}", index + 1)),
            "table caption {} reads {caption:?}",
            index + 1
        );
    }
    for (index, caption) in figures.iter().enumerate() {
        assert!(
            caption.starts_with(&format!("그림 {}", index + 1)),
            "figure caption {} reads {caption:?}",
            index + 1
        );
    }
}

#[test]
fn test_east_asian_font_family_survives_beside_the_latin_one() {
    // Word shapes a run's Latin codepoints with w:ascii and its East Asian
    // ones with w:eastAsia. Collapsing the two into one family dropped
    // whichever came second, so Hangul was shaped by falling back from the
    // Latin family instead (issue #575).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let body_run = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow),
            _ => None,
        })
        .flat_map(|flow| flow.content.iter())
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| paragraph.runs.iter())
        .find(|run| run.text.contains("문서 변환은"))
        .expect("a body run");

    assert_eq!(
        body_run.style.font_family.as_deref(),
        Some("Calibri"),
        "the Latin family is unchanged"
    );
    assert_eq!(
        body_run.style.east_asian_font_family.as_deref(),
        Some("맑은 고딕"),
        "the East Asian family survives beside it"
    );
}

#[test]
fn test_dirty_toc_field_becomes_a_contents_block() {
    // A generated document ships its TOC field empty for Word to fill on
    // open, so the paragraph holding it has no text and the contents page
    // came out blank (issue #576).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let contents: Vec<&crate::ir::TableOfContents> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow),
            _ => None,
        })
        .flat_map(|flow| flow.content.iter())
        .filter_map(|block| match block {
            Block::TableOfContents(contents) => Some(contents),
            _ => None,
        })
        .collect();

    assert_eq!(
        contents,
        vec![
            &crate::ir::TableOfContents::Headings { depth: 3 },
            &crate::ir::TableOfContents::Captions {
                identifier: "Figure".to_string()
            },
            &crate::ir::TableOfContents::Captions {
                identifier: "Table".to_string()
            },
        ],
        "the brief's three dirty fields each become the list they name"
    );
}

#[test]
fn test_seq_captions_become_caption_blocks_a_list_can_collect() {
    // A `TOC \a` list collects the paragraphs a `SEQ` identifier numbers, so
    // those paragraphs are marked as they are parsed rather than matched by
    // their text afterwards (issue #576).
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/office2pdf_technical_brief_ko.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let captions: Vec<&crate::ir::Caption> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow),
            _ => None,
        })
        .flat_map(|flow| flow.content.iter())
        .filter_map(|block| match block {
            Block::Caption(caption) => Some(caption),
            _ => None,
        })
        .collect();

    let figures: Vec<&crate::ir::Caption> = captions
        .iter()
        .copied()
        .filter(|caption| caption.identifier == "Figure")
        .collect();
    let tables: Vec<&crate::ir::Caption> = captions
        .iter()
        .copied()
        .filter(|caption| caption.identifier == "Table")
        .collect();

    assert_eq!(figures.len(), 5, "every figure caption is collectable");
    assert_eq!(tables.len(), 33, "every table caption is collectable");

    // Word lists a caption without the label and number that precede it.
    assert_eq!(tables[0].entry_text, "문서 서지 정보");
    assert!(
        !tables[0].entry_text.starts_with("표"),
        "the list entry drops the label and the field's number: {:?}",
        tables[0].entry_text
    );
    assert!(
        tables[0].paragraph.runs.iter().any(|run| run.text == "1"),
        "the caption itself keeps its number"
    );
}

// ----- Pair kerning (issue #628) -----

#[test]
fn test_run_without_kern_element_resolves_to_never_kerned() {
    // Word writes no `w:kern` unless the user turns kerning on, and every
    // English business mock is written that way — so every glyph advances by
    // its nominal width.
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("THE MONTHLY RENDER")),
    ]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(
        first_run(&doc).style.pair_kerning,
        Some(PairKerning::Never),
        "an absent w:kern must resolve to 'never kerned', not to 'unstated'"
    );
    assert_eq!(
        doc.styles
            .default_text
            .as_ref()
            .and_then(|text| text.pair_kerning),
        Some(PairKerning::Never),
        "the document default carries the same decision for unwrapped runs"
    );
}

#[test]
fn test_kern_element_parses_as_a_half_point_size_threshold() {
    // `w:kern` is a threshold, not a switch: `w:val` is in half-points, so
    // `32` means "kern from 16pt up".
    let with_threshold = serde_json::json!({ "kern": 32 });
    assert_eq!(
        extract_pair_kerning(&with_threshold),
        Some(PairKerning::AtOrAbovePt(16.0))
    );

    let nested = serde_json::json!({ "kern": { "val": 28 } });
    assert_eq!(
        extract_pair_kerning(&nested),
        Some(PairKerning::AtOrAbovePt(14.0))
    );

    // Word records "off" as `w:val="0"`, which is not "kern from 0pt up".
    let explicit_off = serde_json::json!({ "kern": 0 });
    assert_eq!(
        extract_pair_kerning(&explicit_off),
        Some(PairKerning::Never)
    );

    // Absent properties state nothing at all: at run level that is "inherit",
    // and only the document default may read absence as a decision.
    let absent = serde_json::json!({ "sz": 44 });
    assert_eq!(extract_pair_kerning(&absent), None);
}

#[test]
fn test_doc_defaults_kern_is_read_from_the_raw_styles_part() {
    // docx-rs has no field for `w:kern`, so waiting on its parse reports a
    // document that asks Word to kern everything as unkerned. Word writes the
    // "kern everything" case as a 1pt threshold (issue #628 review).
    let rules = PairKerningRules::from_styles_xml(Some(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:rPrDefault><w:rPr>
               <w:kern w:val="2"/><w:sz w:val="22"/>
             </w:rPr></w:rPrDefault></w:docDefaults>
           </w:styles>"#,
    ));

    assert_eq!(rules.document_default(), PairKerning::AtOrAbovePt(1.0));
}

#[test]
fn test_absent_doc_defaults_kern_is_a_decision_not_an_inheritance() {
    let rules = PairKerningRules::from_styles_xml(Some(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
           </w:styles>"#,
    ));

    assert_eq!(rules.document_default(), PairKerning::Never);
    assert_eq!(
        PairKerningRules::from_styles_xml(None).document_default(),
        PairKerning::Never
    );
}

#[test]
fn test_named_style_kern_is_read_and_scoped_to_that_style() {
    // Word's Title style states its own threshold — 28 half-points, so a 28pt
    // title kerns while the 11pt body under the same document default does
    // not (issue #628 review).
    let rules = PairKerningRules::from_styles_xml(Some(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
             <w:style w:type="paragraph" w:styleId="Title">
               <w:name w:val="Title"/>
               <w:pPr><w:spacing w:after="80"/></w:pPr>
               <w:rPr><w:kern w:val="28"/><w:sz w:val="56"/></w:rPr>
             </w:style>
             <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
           </w:styles>"#,
    ));

    assert_eq!(
        rules.for_style("Title"),
        Some(PairKerning::AtOrAbovePt(14.0))
    );
    assert_eq!(
        rules.for_style("Normal"),
        None,
        "a style that states nothing inherits rather than deciding"
    );
    assert_eq!(rules.for_style("NoSuchStyle"), None);
}

#[test]
fn test_word_document_stating_doc_defaults_kern_keeps_kerning() {
    // 1-page.docx is genuine Word output: `<w:kern w:val="2"/>` in
    // `w:docDefaults` asks Word to kern from 1pt up, i.e. everything. Reading
    // it as unkerned re-wrapped its body away from Word's line breaks.
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/1-page.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let body_run = first_run(&doc);
    assert_eq!(
        body_run.style.pair_kerning,
        Some(PairKerning::AtOrAbovePt(1.0)),
        "the document default reaches the body"
    );
    assert!(
        body_run
            .style
            .pair_kerning
            .expect("stated")
            .applies_at(body_run.style.font_size),
        "11pt body text is above a 1pt threshold, so Word kerns it"
    );
}

#[test]
fn test_named_style_threshold_reaches_the_paragraphs_that_use_it() {
    // The brief's Title style states 28 half-points and 56 half-points of
    // size: the threshold has to survive the cascade for the 28pt title to
    // keep the kerning Word gives it.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:rPr><w:kern w:val="28"/><w:sz w:val="56"/></w:rPr>
  </w:style>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Quarterly Review</w:t></w:r></w:p>
  <w:p><w:r><w:t>Body copy that Word leaves unkerned.</w:t></w:r></w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraphs: Vec<&Paragraph> = collect_paragraphs(&doc);

    let title: &Run = &paragraphs[0].runs[0];
    assert_eq!(title.style.font_size, Some(28.0));
    assert_eq!(
        title.style.pair_kerning,
        Some(PairKerning::AtOrAbovePt(14.0)),
        "the Title style's own threshold, not the document default"
    );
    assert!(
        title
            .style
            .pair_kerning
            .expect("stated")
            .applies_at(title.style.font_size),
        "a 28pt title is above a 14pt threshold"
    );

    let body: &Run = &paragraphs[1].runs[0];
    assert_eq!(
        body.style.pair_kerning,
        Some(PairKerning::Never),
        "a paragraph outside the Title style keeps the document default"
    );
}

#[test]
fn test_run_stating_no_kern_keeps_the_style_threshold() {
    // An absent `w:kern` in a run's `w:rPr` means "inherit", not "off": the
    // 28pt Title below states its size and nothing else, and must still be
    // kerned by its style's 14pt threshold (issue #628 review, defect 3).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:rPr><w:kern w:val="28"/></w:rPr>
  </w:style>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>
    <w:r><w:rPr><w:sz w:val="56"/><w:b/></w:rPr><w:t>Quarterly Review</w:t></w:r>
  </w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let run = first_run(&doc);

    assert_eq!(run.style.font_size, Some(28.0));
    assert_eq!(
        run.style.pair_kerning,
        Some(PairKerning::AtOrAbovePt(14.0)),
        "a silent run must not overwrite its style's threshold"
    );
}

/// The `w:spacing w:after` a paragraph ends up with when the document states
/// nothing, keyed by what `word/styles.xml` declares.
///
/// Measured against native Word by the issue #1085 probes over
/// `korean_alignment_autospace.docx` — Malgun Gothic 10.5pt, no `w:spacing`
/// anywhere, intra-paragraph pitch 18.24pt:
///
/// - `issue-1085-space-after-declared` (patching every `w:pPr`): `w:after="0"`
///   pulls the page up 24.00pt over the three gaps, `w:after="240"` pushes it
///   down 12.00pt, and `w:after="160"` reproduces the untouched export exactly
///   — so the silent document already lays out as 8pt after.
/// - `issue-1085-space-after-pprdefault` and `-default-shape` (patching
///   `styles.xml`): any `w:docDefaults/w:pPrDefault` — carrying `w:after="0"`,
///   carrying only `w:before`, holding an empty `w:pPr`, or written as the bare
///   `<w:pPrDefault/>` element — collapses the gap to 0 (−24.00pt), while a
///   `Normal` style holding its own `w:pPr`, and deleting the `Normal`
///   definition outright, both leave the 8pt untouched.
const BUILT_IN_NORMAL_SPACE_AFTER_PT: f64 = 8.0;
const DECLARED_DEFAULTS_SPACE_AFTER_PT: f64 = 0.0;

#[test]
fn test_document_without_paragraph_property_defaults_opens_words_built_in_8pt() {
    // The fixture the issue measured: four paragraphs, none stating
    // `w:spacing`, over a `styles.xml` whose `w:docDefaults` holds only
    // `w:rPrDefault`. Word opens 8pt below each of them.
    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/korean_alignment_autospace.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraphs: Vec<&Paragraph> = collect_paragraphs(&doc);

    assert_eq!(paragraphs.len(), 4);
    for (index, paragraph) in paragraphs.iter().enumerate() {
        assert_eq!(
            paragraph.style.space_after,
            Some(BUILT_IN_NORMAL_SPACE_AFTER_PT),
            "paragraph {index} states no w:spacing and the package declares no \
             w:pPrDefault, so Word's built-in Normal supplies w:after=\"160\""
        );
    }
}

#[test]
fn test_declared_paragraph_property_defaults_replace_the_built_in_8pt() {
    // A `w:pPrDefault` is the document taking over its own paragraph defaults,
    // and ECMA-376 leaves an unstated `w:after` at zero — so even one holding
    // an empty `w:pPr` closes the gap Word's built-in Normal would have opened.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:pPrDefault><w:pPr/></w:pPrDefault>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>First clause of the agreement.</w:t></w:r></w:p>
  <w:p><w:r><w:t>Second clause of the agreement.</w:t></w:r></w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    for paragraph in collect_paragraphs(&doc) {
        assert_eq!(
            paragraph.style.space_after,
            Some(DECLARED_DEFAULTS_SPACE_AFTER_PT),
            "a declared w:pPrDefault silences the built-in default"
        );
    }
}

#[test]
fn test_bare_paragraph_property_default_element_replaces_the_built_in_8pt() {
    // `<w:pPrDefault/>` carries no `w:pPr` at all, so nothing distinguishes it
    // from absence in the parsed model — yet Word's export moves by the same
    // 24pt as `w:after="0"` does. The element's presence is the whole signal.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:pPrDefault/>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Notice of the annual general meeting.</w:t></w:r></w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(
        collect_paragraphs(&doc)[0].style.space_after,
        Some(DECLARED_DEFAULTS_SPACE_AFTER_PT)
    );
}

#[test]
fn test_style_paragraph_properties_do_not_silence_the_built_in_8pt() {
    // A `Normal` definition holding its own `w:pPr` is not the document taking
    // over the defaults: the probe's `Normal style w:spacing w:before only`
    // variant exported at the untouched baselines, i.e. still 8pt after. Only
    // `w:docDefaults/w:pPrDefault` moves it.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:before="0"/></w:pPr>
  </w:style>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Notice of the annual general meeting.</w:t></w:r></w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(
        collect_paragraphs(&doc)[0].style.space_after,
        Some(BUILT_IN_NORMAL_SPACE_AFTER_PT)
    );
}

#[test]
fn test_stated_spacing_outranks_the_built_in_default_either_way() {
    // Triangulation: the default only fills a gap nothing states. A paragraph
    // stating `w:after="240"` keeps its 12pt with no `w:pPrDefault` in sight,
    // and one stating `w:after="0"` keeps its zero.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:t>Widely spaced clause.</w:t></w:r></w:p>
  <w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Tightly spaced clause.</w:t></w:r></w:p>
  <w:p><w:r><w:t>Silent clause.</w:t></w:r></w:p>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraphs: Vec<&Paragraph> = collect_paragraphs(&doc);

    assert_eq!(paragraphs[0].style.space_after, Some(12.0));
    assert_eq!(paragraphs[1].style.space_after, Some(0.0));
    assert_eq!(
        paragraphs[2].style.space_after,
        Some(BUILT_IN_NORMAL_SPACE_AFTER_PT)
    );
}

#[test]
fn test_paragraph_property_defaults_are_scanned_from_the_raw_styles_part() {
    // docx-rs models a missing `w:pPrDefault` and one holding an empty `w:pPr`
    // identically, so the presence has to come off the raw part — the reason
    // `PairKerningRules` reads `w:kern` there (issue #628).
    let declares = |xml: &str| styles::scan_declares_paragraph_property_defaults(xml);

    assert!(declares(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
           </w:styles>"#
    ));
    assert!(
        declares(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:pPrDefault/><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>
           </w:styles>"#
        ),
        "the bare element declares the defaults just as a populated one does"
    );
    assert!(
        !declares(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
             <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
             <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
               <w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr>
             </w:style>
           </w:styles>"#
        ),
        "a style's own w:pPr is not a document-level default"
    );
    assert!(!declares(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
    ));
}

fn collect_paragraphs(doc: &Document) -> Vec<&Paragraph> {
    doc.pages
        .iter()
        .flat_map(|page| match page {
            Page::Flow(flow) => flow.content.iter(),
            _ => [].iter(),
        })
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect()
}

#[test]
fn a_centred_title_style_with_an_outline_level_prints_centred() {
    // The Title style zh-CN Word ships is centred, bold 16pt and carries
    // `w:outlineLvl 0`, so the parser resolves it as a level-1 heading. A
    // template's `{Title}` placeholder in it printed flush left, while the
    // same paragraph in a style without the outline level printed centred
    // (issue #1820). The two must land on the same x.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="a3">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>
      <w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="CentredLine">
    <w:name w:val="Centred Line"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>
      <w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
</w:styles>"#;
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr><w:r><w:t>{Title}</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="CentredLine"/></w:pPr><w:r><w:t>{Title}</w:t></w:r></w:p>
  <w:p><w:r><w:t>Body copy under the title.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>
</w:body></w:document>"#;

    let data = build_docx_with_styles_xml(document_xml, styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let paragraphs: Vec<&Paragraph> = collect_paragraphs(&doc);
    assert_eq!(paragraphs[0].style.heading_level, Some(1));
    assert_eq!(paragraphs[1].style.heading_level, None);

    let source: String = crate::render::typst_gen::generate_typst(&doc)
        .unwrap()
        .source;
    let lefts: Vec<f64> = crate::render::pdf::compiled_text_runs(&source, 0)
        .unwrap()
        .into_iter()
        .filter(|run| run.text == "{Title}")
        .map(|run| run.left_pt)
        .collect();
    let [title, centred_line] = lefts[..] else {
        panic!("expected both placeholders on page 1, got {lefts:?}");
    };
    // w:left="1800" is a 90pt margin; a centred 16pt placeholder starts far
    // inside it on a 415pt measure.
    assert!(
        centred_line > 90.0 + 100.0,
        "the style without an outline level is centred: {centred_line}pt"
    );
    assert!(
        (title - centred_line).abs() < 0.01,
        "the outline-level Title starts at {title}pt, not where the same \
         centred paragraph does ({centred_line}pt)"
    );
}
