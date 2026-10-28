//! `w:contextualSpacing` through the whole DOCX parse (issue #1684). The
//! expected gaps are the ones native Word exports of the same shapes showed.

use super::*;

const WORDPROCESSINGML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const PACKAGE_RELATIONSHIPS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

/// Word's own shape: an 8pt `w:after` from `w:pPrDefault`, and the flag on
/// `ListParagraph`.
fn styles_xml() -> String {
    format!(
        r#"<w:styles xmlns:w="{WORDPROCESSINGML}">
<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
</w:styles>"#
    )
}

const BULLET_NUMBERING: &str = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;

fn document_xml(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORDPROCESSINGML}"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:v="urn:schemas-microsoft-com:vml"
            mc:Ignorable="wps"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    )
}

/// A package holding `document.xml` plus whichever optional parts are given.
fn build_package(
    document_xml: &str,
    styles_xml: Option<&str>,
    numbering_xml: Option<&str>,
) -> Vec<u8> {
    let mut content_types: String = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
    );
    let mut document_relationships: String = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut parts: Vec<(&str, &str)> = vec![("word/document.xml", document_xml)];
    if let Some(styles_xml) = styles_xml {
        content_types.push_str(r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#);
        document_relationships.push_str(r#"<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#);
        parts.push(("word/styles.xml", styles_xml));
    }
    if let Some(numbering_xml) = numbering_xml {
        content_types.push_str(r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#);
        document_relationships.push_str(r#"<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>"#);
        parts.push(("word/numbering.xml", numbering_xml));
    }
    content_types.push_str("</Types>");
    document_relationships.push_str("</Relationships>");

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for (name, content) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", PACKAGE_RELATIONSHIPS),
        (
            "word/_rels/document.xml.rels",
            document_relationships.as_str(),
        ),
    ]
    .into_iter()
    .chain(parts)
    {
        zip.start_file(name, options).unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn parse(data: &[u8]) -> Document {
    let (doc, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    doc
}

/// Every paragraph in reading order — list items and table cells included —
/// as its text with its `space_before` and `space_after`.
fn paragraph_gaps(blocks: &[Block]) -> Vec<(String, Option<f64>, Option<f64>)> {
    let gap = |paragraph: &Paragraph| {
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        (
            text,
            paragraph.style.space_before,
            paragraph.style.space_after,
        )
    };
    let mut gaps: Vec<(String, Option<f64>, Option<f64>)> = Vec::new();
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => gaps.push(gap(paragraph)),
            Block::List(list) => gaps.extend(
                list.items
                    .iter()
                    .flat_map(|item| item.content.iter())
                    .map(gap),
            ),
            Block::Table(table) => {
                for cell in table.rows.iter().flat_map(|row| row.cells.iter()) {
                    gaps.extend(paragraph_gaps(&cell.content));
                }
            }
            _ => {}
        }
    }
    gaps
}

fn space_after_by_text(doc: &Document) -> Vec<(String, Option<f64>)> {
    paragraph_gaps(all_blocks(doc))
        .into_iter()
        .map(|(text, _, after)| (text, after))
        .collect()
}

fn paragraph(style_id: Option<&str>, properties: &str, text: &str) -> String {
    let style: String = style_id
        .map(|style_id| format!(r#"<w:pStyle w:val="{style_id}"/>"#))
        .unwrap_or_default();
    format!(r#"<w:p><w:pPr>{style}{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

#[test]
fn issue_1684_flagged_paragraphs_lose_the_after_word_drops() {
    // The issue's own package: no styles part, eight bare paragraphs with
    // `w:after="160"`, the first four flagged. Word drops the fourth gap as
    // well, because the plain paragraph below shares the default style.
    let flagged = (1..=4).map(|item| {
        paragraph(
            None,
            r#"<w:spacing w:after="160"/><w:contextualSpacing/>"#,
            &format!("contextualSpacing item {item}"),
        )
    });
    let plain = (1..=4).map(|item| {
        paragraph(
            None,
            r#"<w:spacing w:after="160"/>"#,
            &format!("plain paragraph {item}"),
        )
    });
    let body: String = flagged.chain(plain).collect();

    let doc = parse(&build_package(&document_xml(&body), None, None));

    let after: Vec<Option<f64>> = space_after_by_text(&doc)
        .into_iter()
        .map(|(_, after)| after)
        .collect();
    assert_eq!(
        after,
        vec![
            Some(0.0),
            Some(0.0),
            Some(0.0),
            Some(0.0),
            Some(8.0),
            Some(8.0),
            Some(8.0),
            Some(8.0)
        ]
    );
}

#[test]
fn bulleted_list_items_keep_only_the_gaps_to_other_styles() {
    let body: String = [
        paragraph(None, "", "Quarterly results"),
        paragraph(
            Some("ListParagraph"),
            r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#,
            "Revenue grew 12%",
        ),
        paragraph(
            Some("ListParagraph"),
            r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#,
            "Churn fell to 3%",
        ),
        paragraph(
            Some("ListParagraph"),
            r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#,
            "Two regions opened",
        ),
        paragraph(None, "", "Details follow."),
    ]
    .concat();

    let doc = parse(&build_package(
        &document_xml(&body),
        Some(&styles_xml()),
        Some(BULLET_NUMBERING),
    ));

    assert!(
        all_blocks(&doc)
            .iter()
            .any(|block| matches!(block, Block::List(list) if list.items.len() == 3)),
        "the three items should form one list"
    );
    assert_eq!(
        space_after_by_text(&doc),
        vec![
            ("Quarterly results".to_string(), Some(8.0)),
            ("Revenue grew 12%".to_string(), Some(0.0)),
            ("Churn fell to 3%".to_string(), Some(0.0)),
            ("Two regions opened".to_string(), Some(8.0)),
            ("Details follow.".to_string(), Some(8.0)),
        ]
    );
}

#[test]
fn kept_before_is_offset_by_the_after_dropped_above_it() {
    // Probe c16: 10pt dropped above a kept 15pt `w:before` leaves 5pt.
    let body: String = [
        paragraph(
            None,
            r#"<w:spacing w:after="200"/><w:contextualSpacing/>"#,
            "Flagged",
        ),
        paragraph(None, r#"<w:spacing w:before="300"/>"#, "Keeps its before"),
    ]
    .concat();

    let doc = parse(&build_package(
        &document_xml(&body),
        Some(&styles_xml()),
        None,
    ));

    let gaps = paragraph_gaps(all_blocks(&doc));
    assert_eq!(gaps[0].2, Some(0.0), "the flagged paragraph's after drops");
    assert_eq!(gaps[1].1, Some(5.0), "15pt before less the 10pt dropped");
}

#[test]
fn table_cell_paragraphs_drop_the_gaps_between_each_other() {
    // Probe c22: a cell is a flow of its own, like the body.
    let body: String = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>{}{}</w:tc></w:tr></w:tbl>"#,
        paragraph(Some("ListParagraph"), "", "Cell item one"),
        paragraph(Some("ListParagraph"), "", "Cell item two"),
    );

    let doc = parse(&build_package(
        &document_xml(&body),
        Some(&styles_xml()),
        None,
    ));

    assert_eq!(
        space_after_by_text(&doc),
        vec![
            ("Cell item one".to_string(), Some(0.0)),
            ("Cell item two".to_string(), Some(8.0)),
        ]
    );
}

#[test]
fn text_box_ahead_of_a_list_leaves_the_gaps_on_their_own_paragraphs() {
    // The fallback copy of a DrawingML text box never reaches paragraph
    // conversion. Counting it would shift every later paragraph's rule by
    // one, dropping the wrong gaps.
    let text_box_run: String = format!(
        r#"<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Text Box 1"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:txbx><w:txbxContent>{box_paragraph}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>
</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>{box_paragraph}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>"#,
        box_paragraph = r#"<w:p><w:r><w:t>Inside box</w:t></w:r></w:p>"#,
    );
    let body: String = [
        format!(r#"<w:p><w:r><w:t>Anchor</w:t></w:r>{text_box_run}</w:p>"#),
        paragraph(Some("ListParagraph"), "", "First item"),
        paragraph(Some("ListParagraph"), "", "Second item"),
        paragraph(None, "", "Closing"),
    ]
    .concat();

    let doc = parse(&build_package(
        &document_xml(&body),
        Some(&styles_xml()),
        None,
    ));

    let after: Vec<(String, Option<f64>)> = space_after_by_text(&doc)
        .into_iter()
        .filter(|(text, _)| text != "Inside box" && !text.is_empty())
        .collect();
    assert_eq!(
        after,
        vec![
            ("Anchor".to_string(), Some(8.0)),
            ("First item".to_string(), Some(0.0)),
            ("Second item".to_string(), Some(8.0)),
            ("Closing".to_string(), Some(8.0)),
        ]
    );
}
