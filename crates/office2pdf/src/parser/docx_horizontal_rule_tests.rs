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

/// The body's blocks as text and rule per paragraph, with each page and
/// column break in its place.
fn outline(blocks: &[Block]) -> Vec<(String, Option<HorizontalRule>)> {
    blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some((text_of(paragraph), rule_of(paragraph))),
            Block::PageBreak => Some(("<page>".to_string(), None)),
            Block::ColumnBreak => Some(("<column>".to_string(), None)),
            _ => None,
        })
        .collect()
}

/// A page or column break splits its paragraph into the lines before and
/// after it. A rule alone on its side of the break is that stretch's line, as
/// a rule-only paragraph is, so it is drawn on the page it ends or the page it
/// resumes on. A paragraph opening with a page break before a rule dropped the
/// rule as a mere break carrier (a POI fixture, `bug57031.docx`).
#[test]
fn a_break_gives_a_rule_alone_on_its_side_the_line_there() {
    let full = HorizontalRule {
        width: HorizontalRuleWidth::Fraction(1.0),
        height: 1.5,
        alignment: Alignment::Center,
        paint: HorizontalRulePaint::Shaded,
    };
    let full_rule = rule_run(r#"style="width:0;height:1.5pt" o:hralign="center" o:hr="t""#);
    let half_rule = rule_run(r#"style="width:0;height:1.5pt" o:hrpct="500" o:hr="t""#);
    let page_break = r#"<w:r><w:br w:type="page"/></w:r>"#;
    let column_break = r#"<w:r><w:br w:type="column"/></w:r>"#;
    let spacing = r#"<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>"#;
    let blocks = parse_body(&format!(
        "<w:p>{page_break}{full_rule}</w:p>\
         <w:p>{half_rule}{column_break}</w:p>\
         <w:p>{spacing}<w:r><w:t>Before</w:t></w:r>{page_break}{full_rule}</w:p>\
         <w:p>{spacing}{full_rule}{page_break}<w:r><w:t>After</w:t></w:r></w:p>",
    ));

    let half = HorizontalRule {
        width: HorizontalRuleWidth::Fraction(0.5),
        alignment: Alignment::Left,
        ..full
    };
    assert_eq!(
        outline(&blocks),
        [
            ("<page>".to_string(), None),
            (String::new(), Some(full)),
            (String::new(), Some(half)),
            ("<column>".to_string(), None),
            ("Before".to_string(), None),
            ("<page>".to_string(), None),
            (String::new(), Some(full)),
            (String::new(), Some(full)),
            ("<page>".to_string(), None),
            ("After".to_string(), None),
        ],
        "{blocks:#?}"
    );

    // The rule's line is one of its paragraph's lines: only the paragraph's
    // first line keeps the space before, and only its last the space after.
    let spaced = paragraphs(&blocks);
    let spacing_of =
        |paragraph: &Paragraph| (paragraph.style.space_before, paragraph.style.space_after);
    assert_eq!(spacing_of(spaced[2]), (Some(12.0), None), "Before");
    assert_eq!(spacing_of(spaced[3]), (None, Some(6.0)), "its rule");
    assert_eq!(spacing_of(spaced[4]), (Some(12.0), None), "the next rule");
    assert_eq!(spacing_of(spaced[5]), (None, Some(6.0)), "After");
}

/// A rule shares its side of a break with text, or sits in the same `w:r` as
/// the break, where the scan cannot tell which side it is on: such a rule is
/// not drawn, so it never lands on a line of text or on the wrong page.
#[test]
fn a_rule_beside_text_or_its_break_run_draws_nothing() {
    let rule_attributes = r#"style="width:0;height:1.5pt" o:hr="t""#;
    let blocks = parse_body(&format!(
        r#"<w:p><w:r><w:br w:type="page"/></w:r>{}<w:r><w:t>Beside</w:t></w:r></w:p>
           <w:p><w:r><w:br w:type="page"/><w:pict><v:rect {rule_attributes}/></w:pict></w:r></w:p>
           <w:p><w:r><w:pict><v:rect {rule_attributes}/></w:pict><w:br w:type="page"/></w:r></w:p>"#,
        rule_run(rule_attributes),
    ));

    assert!(
        outline(&blocks).iter().all(|(_, rule)| rule.is_none()),
        "{blocks:#?}"
    );
}

const VML_NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:o="urn:schemas-microsoft-com:office:office""#;

/// A package whose section references one header and one footer part, both
/// written out verbatim.
fn build_docx_with_header_and_footer(header_body: &str, footer_body: &str) -> Vec<u8> {
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::SimpleFileOptions;

    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    let parts: [(&str, String); 6] = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"#
                .to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#
                .to_string(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>"#
                .to_string(),
        ),
        (
            "word/header1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr {VML_NAMESPACES}>{header_body}</w:hdr>"#
            ),
        ),
        (
            "word/footer1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr {VML_NAMESPACES}>{footer_body}</w:ftr>"#
            ),
        ),
        (
            "word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId8"/>
      <w:footerReference w:type="default" r:id="rId9"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"#
                .to_string(),
        ),
    ];
    for (name, content) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Each header or footer paragraph's text and horizontal rule, in story order.
fn story_rules(story: &crate::ir::HeaderFooter) -> Vec<(String, Option<HorizontalRule>)> {
    story
        .paragraphs
        .iter()
        .map(|paragraph| {
            let text: String = paragraph
                .elements
                .iter()
                .filter_map(|element| match element {
                    crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
                .collect();
            (text, paragraph.style.horizontal_rule.as_deref().copied())
        })
        .collect()
}

fn header_and_footer(header_body: &str, footer_body: &str) -> (HeaderFooter, HeaderFooter) {
    let data = build_docx_with_header_and_footer(header_body, footer_body);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &doc.pages[0] else {
        panic!("Expected FlowPage");
    };
    (
        page.header.clone().expect("the section has a header"),
        page.footer.clone().expect("the section has a footer"),
    )
}

/// Header and footer stories convert through their own path, which read no
/// rule at all, so the issue's header rule was a blank line (issue #2251). A
/// story paragraph whose only content is one rule now carries it, counted
/// among the story's own paragraphs: a table's cell paragraphs and a text
/// box's are not story paragraphs and shift nothing.
///
/// TODO(#2411): a paragraph after a table in the story's flow is drawn at the
/// top of the story, over the table, so its rule is not carried until that
/// paragraph is drawn below the table. A positioned table is out of the flow
/// and leaves the paragraph after it its rule.
#[test]
fn a_header_or_footer_rule_only_paragraph_carries_its_rule() {
    let shaded_rule = r##"style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0""##;
    let partial_rule = r#"style="width:291.7pt;height:1.5pt" o:hrpct="609" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="gray""#;
    let flow_table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>
            <w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
            {}<w:p>{}</w:p>
        </w:tc></w:tr></w:tbl>"#,
        text_paragraph("Cell"),
        rule_run(shaded_rule),
    );
    let positioned_table = flow_table.replace(
        "<w:tblGrid>",
        r#"<w:tblPr><w:tblpPr w:vertAnchor="page" w:tblpY="2000"/></w:tblPr><w:tblGrid>"#,
    );
    let text_box_paragraph = format!(
        r#"<w:p><w:r><w:pict><v:shape style="width:100pt;height:20pt"><v:textbox><w:txbxContent>
              <w:p>{}</w:p>
            </w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:t>Boxed</w:t></w:r></w:p>"#,
        rule_run(shaded_rule),
    );
    let header_body = format!(
        "{}<w:p/>{}<w:p>{}</w:p>{}<w:p>{}</w:p>{}<w:p>{}</w:p>",
        text_paragraph("Title"),
        text_box_paragraph,
        rule_run(partial_rule),
        positioned_table,
        rule_run(shaded_rule),
        flow_table,
        rule_run(shaded_rule),
    );
    let footer_body = format!(
        "<w:p>{}</w:p>{}",
        rule_run(shaded_rule),
        text_paragraph("Footer")
    );

    let (header, footer) = header_and_footer(&header_body, &footer_body);

    let full_width = HorizontalRule {
        width: HorizontalRuleWidth::Fraction(1.0),
        height: 1.5,
        alignment: Alignment::Center,
        paint: HorizontalRulePaint::Shaded,
    };
    assert_eq!(
        story_rules(&header),
        [
            ("Title".to_string(), None),
            (String::new(), None),
            ("Boxed".to_string(), None),
            (
                String::new(),
                Some(HorizontalRule {
                    width: HorizontalRuleWidth::Fraction(0.609),
                    ..full_width
                })
            ),
            (String::new(), Some(full_width)),
            (String::new(), None),
        ],
        "{header:#?}"
    );
    assert_eq!(
        story_rules(&footer),
        [
            (String::new(), Some(full_width)),
            ("Footer".to_string(), None)
        ]
    );
}
