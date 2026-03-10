use super::*;

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
    assert!((widths[0] - 300.0).abs() < 0.1, "width[0]: {}", widths[0]);
    assert!((widths[1] - 150.0).abs() < 0.1, "width[1]: {}", widths[1]);
}

#[test]
fn test_parse_docx_no_columns() {
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

fn make_bidi_paragraph(text: &str) -> docx_rs::Paragraph {
    let mut para = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(text));
    para.property = docx_rs::ParagraphProperty::new().bidi(true);
    para
}

#[test]
fn test_parse_docx_bidi_paragraph() {
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
