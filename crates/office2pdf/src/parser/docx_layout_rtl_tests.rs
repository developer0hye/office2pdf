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
    assert!(!cols.has_separator, "missing w:sep defaults to no rule");
}

#[test]
fn test_parse_docx_column_separator_from_real_fixture() {
    let fixture: &[u8] = include_bytes!(
        "../../../../tests/fixtures/docx/libreoffice/tdf46940_dontEquallyDistributeColumns.docx"
    );
    let parser = DocxParser;
    let (document, _warnings) = parser.parse(fixture, &ConvertOptions::default()).unwrap();

    let column_layouts: Vec<_> = document
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) | Page::FlowContinuous(flow) => flow.columns.as_ref(),
            _ => None,
        })
        .collect();
    assert_eq!(column_layouts.len(), 3, "fixture section layout changed");
    assert!(
        column_layouts.iter().all(|columns| columns.has_separator),
        "the three w:cols/@w:sep flags must survive section parsing"
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
            <w:cols w:num="2" w:space="720" w:equalWidth="0" w:sep="1">
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
    assert!(cols.has_separator, "w:sep=1 must enable column rules");
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
fn test_parse_docx_run_page_break() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Before"))
            .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page))
            .add_run(docx_rs::Run::new().add_text("After")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    assert!(
        flow.content
            .iter()
            .any(|block| matches!(block, Block::PageBreak)),
        "a run-level page break should remain a structural page break"
    );
}

#[test]
fn test_inline_image_before_run_page_break_stays_before_break() {
    let image_bytes: Vec<u8> = super::image_tests::make_test_bmp();
    let picture: docx_rs::Pic = docx_rs::Pic::new(&image_bytes).size(9_525, 9_525);
    let paragraph: docx_rs::Paragraph = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_image(picture))
        .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page))
        .add_run(docx_rs::Run::new().add_text("Next-page content"));
    let data: Vec<u8> = build_docx_bytes(vec![paragraph]);
    let parser: DocxParser = DocxParser;
    let (document, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("synthetic DOCX with inline image and page break should parse");
    let flow = match &document.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let image_index: usize = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::Image(_) | Block::InlineImages(_)))
        .expect("the image remains in the paragraph flow");
    let page_break_index: usize = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::PageBreak))
        .expect("the run-level page break remains structural");
    let next_page_index: usize = flow
        .content
        .iter()
        .position(|block| match block {
            Block::Paragraph(paragraph) => paragraph
                .runs
                .iter()
                .any(|run| run.text.contains("Next-page content")),
            _ => false,
        })
        .expect("the text after the break remains in the paragraph flow");

    assert!(
        image_index < page_break_index,
        "inline content before a page break must stay on the preceding page: {:?}",
        flow.content
            .iter()
            .map(std::mem::discriminant)
            .collect::<Vec<_>>()
    );
    assert!(
        page_break_index < next_page_index,
        "content after a page break must start after it"
    );
}

#[test]
fn test_shaded_paragraph_fragment_before_leading_page_break_is_preserved() {
    let mut paragraph: docx_rs::Paragraph = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page))
        .add_run(docx_rs::Run::new().add_text("Visible after"));
    paragraph.property = docx_rs::ParagraphProperty::new()
        .shading(docx_rs::Shading::new().fill("00E4D5"))
        .set_borders(docx_rs::ParagraphBorders::new());
    let data: Vec<u8> = build_docx_bytes(vec![paragraph]);
    let parser: DocxParser = DocxParser;
    let (document, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("synthetic shaded DOCX with a leading page break should parse");
    let flow = match &document.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let page_break_index: usize = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::PageBreak))
        .expect("the run-level page break remains structural");
    let has_shaded_empty_fragment_before_break: bool =
        flow.content[..page_break_index]
            .iter()
            .any(|block| match block {
                Block::Paragraph(paragraph) => {
                    paragraph.runs.is_empty()
                        && paragraph.style.background.is_some()
                        && paragraph.style.border.is_none()
                }
                _ => false,
            });

    assert!(
        has_shaded_empty_fragment_before_break,
        "paragraph shading before a leading page break must remain on the preceding page"
    );
}

#[test]
fn test_text_before_and_after_break_in_same_run_keeps_order() {
    let paragraph: docx_rs::Paragraph = docx_rs::Paragraph::new().add_run(
        docx_rs::Run::new()
            .add_text("Visible before")
            .add_break(docx_rs::BreakType::Page)
            .add_text(" visible after"),
    );
    let data: Vec<u8> = build_docx_bytes(vec![paragraph]);
    let parser: DocxParser = DocxParser;
    let (document, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("synthetic DOCX with text around a page break should parse");
    let flow = match &document.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let page_break_index: usize = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::PageBreak))
        .expect("the run-level page break remains structural");
    let text_before_break: String = flow.content[..page_break_index]
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
    let text_after_break: String = flow.content[page_break_index + 1..]
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

    assert_eq!(text_before_break, "Visible before");
    assert_eq!(text_after_break, " visible after");
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
        paragraph_space_after_is_explicit: false,
        paragraph_tab_overrides: None,
        heading_level: None,
        heading_has_document_run_formatting: false,
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
        paragraph_space_after_is_explicit: false,
        paragraph_tab_overrides: None,
        heading_level: None,
        heading_has_document_run_formatting: false,
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
        paragraph_space_after_is_explicit: false,
        paragraph_tab_overrides: None,
        heading_level: None,
        heading_has_document_run_formatting: false,
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

/// Word uses a paragraph holding only `<w:br w:type="page"/>` to force a page
/// break; it contributes no line box, so the next block starts at the top of
/// the new page.
#[test]
fn test_parse_docx_page_break_carrier_paragraph_adds_no_blank_line() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Page one body")),
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page)),
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Appendix A. Exit Codes")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let break_index = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::PageBreak))
        .expect("the carrier paragraph still forces a page break");
    match flow.content.get(break_index + 1) {
        Some(Block::Paragraph(paragraph)) => {
            let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
            assert_eq!(
                text, "Appendix A. Exit Codes",
                "the next-page content must follow the break directly"
            );
        }
        other => panic!("expected the next-page paragraph after the break, got {other:?}"),
    }
}

/// An intentionally empty paragraph is a real blank line in Word and must be
/// kept; only break carriers collapse.
#[test]
fn test_parse_docx_empty_paragraph_without_break_is_kept() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First")),
        docx_rs::Paragraph::new(),
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let paragraphs: Vec<&crate::ir::Paragraph> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    assert_eq!(
        paragraphs.len(),
        3,
        "the blank spacer paragraph must survive"
    );
    assert!(
        paragraphs[1].runs.iter().all(|run| run.text.is_empty()),
        "the middle paragraph stays empty"
    );
}

/// Text placed after an in-paragraph page break still forms a paragraph on the
/// new page.
#[test]
fn test_parse_docx_text_after_run_page_break_still_renders() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page))
            .add_run(docx_rs::Run::new().add_text("Second page text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let break_index = flow
        .content
        .iter()
        .position(|block| matches!(block, Block::PageBreak))
        .expect("page break present");
    match flow.content.get(break_index + 1) {
        Some(Block::Paragraph(paragraph)) => {
            let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
            assert_eq!(text, "Second page text");
        }
        other => panic!("expected the trailing text paragraph, got {other:?}"),
    }
}

// ── w:wordWrap (issue #730) ────────────────────────────────────────────

/// The property reaches the IR from the raw XML, with `0` and `1` distinct
/// from each other and from an absent element. It cannot come from `docx-rs`:
/// the published crate does not parse `w:wordWrap`, and reading the patched
/// fork's field made the package unpublishable (issue #1041).
#[test]
fn test_word_wrap_reaches_the_paragraph_style() {
    let xml = r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:p>
<w:p><w:pPr><w:wordWrap w:val="1"/></w:pPr></w:p>
<w:p><w:pPr><w:wordWrap/></w:pPr></w:p>
<w:p/>
</w:body></w:document>"#;
    let context = super::super::contexts::WordWrapContext::from_xml(Some(xml));
    assert_eq!(context.next_word_wrap(None), Some(false));
    assert_eq!(context.next_word_wrap(None), Some(true));
    assert_eq!(
        context.next_word_wrap(None),
        Some(true),
        "an absent w:val means on, the element's own default"
    );
    assert_eq!(context.next_word_wrap(None), None);
}

/// A table cell's paragraph consumes its own slot, so nesting cannot shift
/// the pairing between the raw scan and the structured walk.
#[test]
fn test_word_wrap_scan_covers_nested_table_paragraphs() {
    let xml = r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p/>
<w:tbl><w:tr><w:tc><w:p><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:p></w:tc></w:tr></w:tbl>
<w:p/>
</w:body></w:document>"#;
    let context = super::super::contexts::WordWrapContext::from_xml(Some(xml));
    assert_eq!(context.next_word_wrap(None), None);
    assert_eq!(context.next_word_wrap(None), Some(false));
    assert_eq!(context.next_word_wrap(None), None);
}

/// A paragraph style's `w:wordWrap` comes from the raw styles.xml the same
/// way (issue #730's style chain).
#[test]
fn test_style_word_wrap_is_scanned_by_style_id() {
    let xml = r#"<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="NoWrap"><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Plain"><w:pPr/></w:style>
<w:style w:type="character" w:styleId="CharNoWrap"><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:style>
</w:styles>"#;
    let map = super::super::contexts::scan_style_word_wrap(Some(xml));
    assert_eq!(map.get("NoWrap"), Some(&false));
    assert_eq!(map.get("Plain"), None);
    assert_eq!(
        map.get("CharNoWrap"),
        None,
        "a character style's stray wordWrap stays out of the paragraph chain"
    );
}

/// Measured on Word: a paragraph's own `w:wordWrap` beats the one its style
/// carries. A `ListParagraph` with `w:val="0"` breaks mid-eojeol even though
/// the style alone keeps eojeol whole.
#[test]
fn test_explicit_word_wrap_overrides_the_style_chain() {
    let explicit = ParagraphStyle {
        word_wrap: Some(false),
        ..ParagraphStyle::default()
    };
    let style = ResolvedStyle {
        text: TextStyle::default(),
        paragraph: ParagraphStyle {
            word_wrap: Some(true),
            ..ParagraphStyle::default()
        },
        paragraph_space_after_is_explicit: false,
        paragraph_tab_overrides: None,
        heading_level: None,
        heading_has_document_run_formatting: false,
    };

    let merged = merge_paragraph_style(&explicit, None, Some(&style));
    assert_eq!(
        merged.word_wrap,
        Some(false),
        "the paragraph's own value wins"
    );

    // And the style still supplies the value when the paragraph says nothing.
    let inherited = merge_paragraph_style(&ParagraphStyle::default(), None, Some(&style));
    assert_eq!(inherited.word_wrap, Some(true));
}

// ── Alignment does not gate the CJK/Latin auto-space (issue #1053) ─────

/// Measured on `02_contract_ko` page 1: every digit-to-Hangul boundary of
/// the centred date line advances 5.78pt in Word against 8.41pt in the list
/// paragraph above it, and 8.41 − 5.78 is the 0.25em space at 10.5pt (issue
/// #728). The #732 probe reframed that corpus case — the date line is a bare
/// paragraph in a package defining no default style, flush for *that* reason —
/// and the #1053 probe finished the job: in a package that does define one,
/// native Word gives a centred line and a justified line the same +2.588pt at
/// every boundary as a left-aligned one. So alignment gates nothing; the style
/// predicate alone decides, and this docx-rs document defines `Normal`.
#[test]
fn test_alignment_does_not_gate_the_east_asian_auto_space() {
    // Built and parsed for real, so this exercises the parser's own gate
    // rather than restating the condition.
    let body_text = |alignment: Option<docx_rs::AlignmentType>| -> String {
        let mut paragraph =
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("2026년"));
        if let Some(a) = alignment {
            paragraph = paragraph.align(a);
        }
        let mut cursor = Cursor::new(Vec::new());
        docx_rs::Docx::new()
            .add_paragraph(paragraph)
            .build()
            .pack(&mut cursor)
            .unwrap();
        let (doc, _warnings) = DocxParser
            .parse(&cursor.into_inner(), &ConvertOptions::default())
            .unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(page) => page,
            other => panic!("Expected FlowPage, got {other:?}"),
        };
        page.content
            .iter()
            .find_map(|block| match block {
                Block::Paragraph(p) => p.runs.first().map(|r| r.text.clone()),
                _ => None,
            })
            .expect("the paragraph survives parsing")
    };

    let widened = body_text(None);
    assert_ne!(
        widened, "2026년",
        "an unaligned body paragraph still gets the auto-space"
    );
    assert_eq!(
        body_text(Some(docx_rs::AlignmentType::Center)),
        widened,
        "a centred one gets exactly the same"
    );
    assert_eq!(
        body_text(Some(docx_rs::AlignmentType::Justified)),
        widened,
        "and so does a justified one"
    );
}
