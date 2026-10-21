//! Word's *Insert > Horizontal Line*: a VML `v:rect o:hr="t"` that docx-rs
//! drops, so a paragraph holding one converted to a blank line (issue #2251).

use super::*;
use crate::ir::{Alignment, HorizontalRule, HorizontalRulePaint, HorizontalRuleWidth};

fn document(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:o="urn:schemas-microsoft-com:office:office">
    <w:body>{body}<w:sectPr/></w:body>
</w:document>"#
    )
}

fn text_paragraph(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn rule_run(attributes: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:pict><v:rect {attributes} stroked="f"/></w:pict></w:r>"#
    )
}

fn paragraphs(blocks: &[Block]) -> Vec<&Paragraph> {
    blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect()
}

fn text_of(paragraph: &Paragraph) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

fn rule_of(paragraph: &Paragraph) -> Option<HorizontalRule> {
    paragraph.style.horizontal_rule.as_deref().copied()
}

fn parse_body(body: &str) -> Vec<Block> {
    let data = build_docx_with_columns(&document(body));
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    match &doc.pages[0] {
        Page::Flow(flow) => flow.content.clone(),
        _ => panic!("Expected FlowPage"),
    }
}

/// A paragraph whose only content is the rule is the rule's line: it keeps
/// its place between its neighbours and carries the rule, which docx-rs
/// alone reduced to a blank line.
#[test]
fn a_rule_only_paragraph_carries_its_rule() {
    let blocks = parse_body(&format!(
        "{}<w:p><w:pPr><w:ind w:left=\"720\"/></w:pPr>{}</w:p>{}<w:p>{}</w:p>",
        text_paragraph("Above"),
        rule_run(
            r##"style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0""##
        ),
        text_paragraph("Between"),
        rule_run(
            r#"style="width:150pt;height:9pt" o:hrpct="0" o:hralign="right" o:hrnoshade="t" o:hr="t" fillcolor="teal""#
        ),
    ));

    let paragraphs = paragraphs(&blocks);
    assert_eq!(
        paragraphs.iter().map(|p| text_of(p)).collect::<Vec<_>>(),
        ["Above", "", "Between", ""],
        "each rule keeps its own line among the text: {blocks:#?}"
    );
    assert_eq!(rule_of(paragraphs[0]), None);
    assert_eq!(
        rule_of(paragraphs[1]),
        Some(HorizontalRule {
            width: HorizontalRuleWidth::Fraction(1.0),
            height: 1.5,
            alignment: Alignment::Center,
            paint: HorizontalRulePaint::Shaded,
        })
    );
    assert_eq!(
        paragraphs[1].style.indent_left,
        Some(36.0),
        "the rule's paragraph keeps its own properties"
    );
    assert_eq!(rule_of(paragraphs[2]), None);
    assert_eq!(
        rule_of(paragraphs[3]),
        Some(HorizontalRule {
            width: HorizontalRuleWidth::Points(150.0),
            height: 9.0,
            alignment: Alignment::Right,
            paint: HorizontalRulePaint::Solid(Color::new(0x00, 0x80, 0x80)),
        })
    );
}

/// A table cell's paragraphs convert through the same path as the body's,
/// so a cell's rule-only paragraph carries its rule too. The renderer does
/// not draw it yet (`TODO(#2251 rules in table cells)`).
#[test]
fn a_rule_in_a_table_cell_rides_its_cell_paragraph() {
    let blocks = parse_body(&format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>
            <w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
            {}<w:p>{}</w:p>
        </w:tc></w:tr></w:tbl>"#,
        text_paragraph("Cell"),
        rule_run(r#"style="width:0;height:1.5pt" o:hrpct="500" o:hr="t""#),
    ));

    let table = blocks
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("the table converts");
    let cell_paragraphs = paragraphs(&table.rows[0].cells[0].content);
    assert_eq!(
        cell_paragraphs
            .iter()
            .map(|paragraph| (text_of(paragraph), rule_of(paragraph)))
            .collect::<Vec<_>>(),
        [
            ("Cell".to_string(), None),
            (
                String::new(),
                Some(HorizontalRule {
                    width: HorizontalRuleWidth::Fraction(0.5),
                    height: 1.5,
                    alignment: Alignment::Left,
                    paint: HorizontalRulePaint::Shaded,
                })
            ),
        ]
    );
}

/// Only `o:hr` makes a rectangle a rule: any other VML rectangle in a run is
/// still not drawn, and its paragraph stays a plain blank line.
#[test]
fn a_rectangle_that_is_not_a_rule_draws_nothing() {
    let blocks = parse_body(&format!(
        "<w:p>{}</w:p>",
        rule_run(r#"style="width:100pt;height:20pt" fillcolor="red""#),
    ));

    let paragraphs = paragraphs(&blocks);
    assert_eq!(paragraphs.len(), 1, "{blocks:#?}");
    assert_eq!(rule_of(paragraphs[0]), None);
}

/// TODO(rule beside text): Word gives the rule a line of its own inside a
/// paragraph that also holds text. Until that line is laid out, such a
/// paragraph keeps its text and draws no rule, so the rule can never land on
/// a line of text.
#[test]
fn a_rule_beside_text_leaves_the_paragraph_its_text() {
    let blocks = parse_body(&format!(
        r#"<w:p>{}<w:r><w:br/></w:r><w:r><w:t>Source</w:t></w:r></w:p>"#,
        rule_run(r#"style="width:0;height:1.5pt" o:hr="t""#),
    ));

    let paragraphs = paragraphs(&blocks);
    assert_eq!(paragraphs.len(), 1, "{blocks:#?}");
    assert_eq!(text_of(paragraphs[0]), "\nSource");
    assert_eq!(rule_of(paragraphs[0]), None);
}
