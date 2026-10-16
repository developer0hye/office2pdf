use super::*;

fn paragraph(properties: &str, text: &str) -> String {
    format!(r#"<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn parse(body: &str, styles: &str) -> Document {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{styles}</w:styles>"#
    );
    let bytes = build_docx_with_styles_xml(&document, &styles);
    DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap()
        .0
}

fn style(id: &str, parent: &str, properties: &str) -> String {
    let parent = if parent.is_empty() {
        String::new()
    } else {
        format!(r#"<w:basedOn w:val="{parent}"/>"#)
    };
    format!(
        r#"<w:style w:type="paragraph" w:styleId="{id}">{parent}<w:pPr>{properties}</w:pPr></w:style>"#
    )
}

fn sequence(document: &Document) -> Vec<String> {
    all_blocks(document)
        .iter()
        .map(|block| match block {
            Block::PageBreak | Block::PageBreakBefore => "BREAK".to_owned(),
            Block::Paragraph(paragraph) => {
                paragraph.runs.iter().map(|run| run.text.as_str()).collect()
            }
            other => panic!("unexpected block: {other:?}"),
        })
        .collect()
}

#[test]
fn page_break_before_inherits_named_style_and_direct_false_clears_it() {
    let styles = [
        style("Base", "", "<w:pageBreakBefore/>"),
        style("Child", "Base", ""),
    ]
    .concat();
    let mut body = paragraph("", "Opening");
    body += &paragraph(r#"<w:pStyle w:val="Child"/>"#, "Inherited");
    for value in ["0", "false", "off"] {
        body += &paragraph(
            &format!(r#"<w:pStyle w:val="Child"/><w:pageBreakBefore w:val="{value}"/>"#),
            value,
        );
    }
    body += &paragraph("", "Ordinary");
    assert_eq!(
        sequence(&parse(&body, &styles)),
        [
            "Opening",
            "BREAK",
            "Inherited",
            "0",
            "false",
            "off",
            "Ordinary"
        ]
    );
}

#[test]
fn page_break_before_child_false_stops_parent_and_direct_true_reenables() {
    for value in ["0", "false", "off"] {
        let styles = [
            style("Base", "", "<w:pageBreakBefore/>"),
            style(
                "Child",
                "Base",
                &format!(r#"<w:pageBreakBefore w:val="{value}"/>"#),
            ),
            style("Grandchild", "Child", ""),
        ]
        .concat();
        let body = [
            paragraph("", "Opening"),
            paragraph(r#"<w:pStyle w:val="Grandchild"/>"#, "Cleared"),
            paragraph(
                r#"<w:pStyle w:val="Grandchild"/><w:pageBreakBefore/>"#,
                "Explicit",
            ),
        ]
        .concat();
        assert_eq!(
            sequence(&parse(&body, &styles)),
            ["Opening", "Cleared", "BREAK", "Explicit"]
        );
    }
}

#[test]
fn page_break_before_inherits_document_defaults_and_default_style() {
    let body = [
        paragraph("", "Implicit"),
        paragraph(r#"<w:pStyle w:val="Unknown"/>"#, "Unknown"),
        paragraph(r#"<w:pStyle w:val="Named"/>"#, "Named"),
    ]
    .concat();
    for defaults in ["<w:docDefaults><w:pPrDefault><w:pPr><w:pageBreakBefore/></w:pPr></w:pPrDefault></w:docDefaults>".to_owned(), r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>"#.to_owned()] {
        let expected = if defaults.contains("docDefaults") { vec!["BREAK", "Implicit", "BREAK", "Unknown", "BREAK", "Named"] } else { vec!["BREAK", "Implicit", "BREAK", "Unknown", "Named"] };
        assert_eq!(sequence(&parse(&body, &(defaults + &style("Named", "", "")))), expected);
    }
}

#[test]
fn page_break_before_ignores_revision_history_in_paragraph_and_style() {
    let styles = style(
        "Base",
        "",
        r#"<w:pageBreakBefore/><w:pPrChange w:id="1"><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr></w:pPrChange>"#,
    );
    let body = [paragraph("", "Opening"), paragraph(r#"<w:pStyle w:val="Base"/>"#, "CurrentStyle"), paragraph(r#"<w:pStyle w:val="Base"/><w:pageBreakBefore w:val="0"/><w:pPrChange w:id="2"><w:pPr><w:pageBreakBefore/></w:pPr></w:pPrChange>"#, "CurrentParagraph")].concat();
    assert_eq!(
        sequence(&parse(&body, &styles)),
        ["Opening", "BREAK", "CurrentStyle", "CurrentParagraph"]
    );
}

#[test]
fn page_break_before_cyclic_or_character_ancestors_do_not_leak() {
    let styles = [style("A", "B", ""), style("B", "A", "<w:pageBreakBefore/>"), style("Paragraph", "Character", ""), r#"<w:style w:type="character" w:styleId="Character"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>"#.to_owned()].concat();
    let body = [
        paragraph("", "Opening"),
        paragraph(r#"<w:pStyle w:val="A"/>"#, "Cycle"),
        paragraph(r#"<w:pStyle w:val="Paragraph"/>"#, "WrongType"),
    ]
    .concat();
    assert_eq!(
        sequence(&parse(&body, &styles)),
        ["Opening", "Cycle", "WrongType"]
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_places_inherited_paragraph_on_a_new_compiled_page() {
    let styles = [
        style("Base", "", "<w:pageBreakBefore/>"),
        style("Child", "Base", ""),
    ]
    .concat();
    let body = [
        paragraph("", "Opening"),
        paragraph(r#"<w:pStyle w:val="Child"/>"#, "Inherited"),
        paragraph(
            r#"<w:pStyle w:val="Child"/><w:pageBreakBefore w:val="false"/>"#,
            "Cleared",
        ),
        paragraph("", "Following"),
    ]
    .concat();
    let document = parse(&body, &styles);
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    assert_eq!(
        crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
        2
    );
    let text = |page| {
        crate::render::pdf::compiled_text_runs(&source, page)
            .unwrap()
            .into_iter()
            .map(|run| run.text)
            .collect::<String>()
    };
    assert_eq!(text(0), "Opening");
    assert_eq!(text(1), "InheritedClearedFollowing");
}

#[test]
fn page_break_before_empty_properties_and_style_revision_do_not_leak() {
    let styles = [
        style(
            "Base",
            "",
            r#"<w:pPrChange w:id="3"><w:pPr><w:pageBreakBefore/></w:pPr></w:pPrChange>"#,
        ),
        style("Child", "Base", ""),
    ]
    .concat();
    let body = [
        paragraph(
            "<w:pPrChange w:id=\"4\"><w:pPr><w:pageBreakBefore/></w:pPr></w:pPrChange>",
            "Opening",
        ),
        paragraph(r#"<w:pStyle w:val="Child"/>"#, "Current"),
        paragraph("", "Following"),
    ]
    .concat();
    assert_eq!(
        sequence(&parse(&body, &styles)),
        ["Opening", "Current", "Following"]
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_at_document_start_does_not_insert_a_blank_page() {
    let document = parse(
        &paragraph(r#"<w:pStyle w:val="Base"/>"#, "First"),
        &style("Base", "", "<w:pageBreakBefore/>"),
    );
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    assert_eq!(
        crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
        1
    );
    assert_eq!(
        crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap()
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>(),
        "First"
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_after_explicit_break_does_not_add_another_blank_page() {
    let body = [
        paragraph("", "Opening"),
        "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>".to_owned(),
        paragraph(r#"<w:pStyle w:val="Base"/>"#, "Next"),
    ]
    .concat();
    let document = parse(&body, &style("Base", "", "<w:pageBreakBefore/>"));
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    assert_eq!(
        crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
        2
    );
    assert!(
        crate::render::pdf::compiled_text_runs(&source, 1)
            .unwrap()
            .iter()
            .any(|run| run.text == "Next")
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_preserves_boundaries_in_equal_and_unequal_columns() {
    for widths in [None, Some(vec![240.0, 180.0])] {
        for (prefix, expected) in [
            (String::new(), 1),
            (
                paragraph("", "Opening") + "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>",
                2,
            ),
        ] {
            let body = prefix + &paragraph(r#"<w:pStyle w:val="Base"/>"#, "Next");
            let mut document = parse(&body, &style("Base", "", "<w:pageBreakBefore/>"));
            let Page::Flow(page) = &mut document.pages[0] else {
                unreachable!()
            };
            page.columns = Some(ColumnLayout {
                num_columns: 2,
                spacing: 36.0,
                column_widths: widths.clone(),
                has_separator: false,
            });
            let source = crate::render::typst_gen::generate_typst(&document)
                .unwrap()
                .source;
            assert_eq!(
                crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
                expected,
                "widths={widths:?}, {source}"
            );
            assert!(
                crate::render::pdf::compiled_text_runs(&source, (expected - 1) as usize)
                    .unwrap()
                    .iter()
                    .any(|run| run.text == "Next")
            );
        }
    }
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_restarts_a_continuous_section_without_a_blank_page() {
    let mut document = parse(&paragraph("", "Opening"), "");
    let mut next = parse(
        &paragraph(r#"<w:pStyle w:val="Base"/>"#, "Next"),
        &style("Base", "", "<w:pageBreakBefore/>"),
    );
    let Page::Flow(page) = next.pages.remove(0) else {
        unreachable!()
    };
    document.pages.push(Page::FlowContinuous(page));
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    assert_eq!(
        crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
        2
    );
    assert!(
        crate::render::pdf::compiled_text_runs(&source, 1)
            .unwrap()
            .iter()
            .any(|run| run.text == "Next")
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn page_break_before_does_not_weaken_explicit_hard_breaks() {
    let body = paragraph("", "Opening")
        + "<w:p><w:r><w:br w:type=\"page\"/><w:br w:type=\"page\"/></w:r></w:p>"
        + &paragraph(r#"<w:pStyle w:val="Base"/>"#, "Next");
    let document = parse(&body, &style("Base", "", "<w:pageBreakBefore/>"));
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    assert_eq!(
        crate::render::pdf::compile_page_count_with_fonts(&source, &[], &[], &[]).unwrap(),
        3
    );
    assert!(
        crate::render::pdf::compiled_text_runs(&source, 2)
            .unwrap()
            .iter()
            .any(|run| run.text == "Next")
    );
}
