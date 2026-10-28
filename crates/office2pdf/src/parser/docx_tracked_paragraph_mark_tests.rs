//! A paragraph mark a tracked deletion or move removed (issue #1710).
//!
//! Word's final view has no paragraph break where the mark is gone, so the
//! paragraph merges into the one after it: it contributes no line, no list
//! number, and its surviving runs open the next paragraph. The shapes here are
//! the ones Word and LibreOffice write — the whole paragraph moved away, half a
//! paragraph deleted across a break, and a `w:rPrChange` that records a past
//! revision rather than the mark's current state.

use super::*;

const WORDPROCESSINGML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const PACKAGE_RELATIONSHIPS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

/// A decimal single-level list, so an item that should not exist shows up as a
/// skipped number rather than only as a blank line.
const DECIMAL_NUMBERING: &str = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;

fn document_xml(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORDPROCESSINGML}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    )
}

fn build_package(document_xml: &str, numbering_xml: Option<&str>) -> Vec<u8> {
    let mut content_types: String = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
    );
    let mut document_relationships: String = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut parts: Vec<(&str, &str)> = vec![("word/document.xml", document_xml)];
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

fn parse_body(document_body: &str, numbering_xml: Option<&str>) -> Document {
    let data = build_package(&document_xml(document_body), numbering_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc
}

fn paragraph_text(paragraph: &Paragraph) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

/// Every line the flow produces, in reading order, as its text — list items
/// included, so a phantom item shows up as an empty entry rather than vanishing
/// from the comparison.
fn line_texts(doc: &Document) -> Vec<String> {
    let mut texts: Vec<String> = Vec::new();
    for page in &doc.pages {
        let Page::Flow(flow) = page else { continue };
        for block in &flow.content {
            match block {
                Block::Paragraph(paragraph) => texts.push(paragraph_text(paragraph)),
                Block::List(list) => texts.extend(
                    list.items
                        .iter()
                        .flat_map(|item| item.content.iter())
                        .map(paragraph_text),
                ),
                _ => {}
            }
        }
    }
    texts
}

/// `w:moveFrom` on both the runs and the mark: the whole origin paragraph is
/// gone from the final view. LibreOffice's `tdf123460` shape.
#[test]
fn a_paragraph_moved_away_whole_leaves_no_line() {
    let doc = parse_body(
        r#"<w:p><w:r><w:t>Before.</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:moveFrom w:id="2" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr>
  <w:moveFromRangeStart w:id="3" w:author="Reviewer" w:date="2026-09-14T10:00:00Z" w:name="move1"/>
  <w:moveFrom w:id="4" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:t>Moved away.</w:t></w:r></w:moveFrom>
</w:p>
<w:moveFromRangeEnd w:id="3"/>
<w:p><w:r><w:t>After.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Before.", "After."]);
}

/// The same rule for a plain tracked deletion of the mark, which is what Word
/// writes when a reviewer deletes a whole paragraph.
#[test]
fn a_paragraph_deleted_whole_leaves_no_line() {
    let doc = parse_body(
        r#"<w:p><w:r><w:t>Before.</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:del w:id="2" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr>
  <w:del w:id="3" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:delText>Deleted.</w:delText></w:r></w:del>
</w:p>
<w:p><w:r><w:t>After.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Before.", "After."]);
}

/// An ordinary empty paragraph is a deliberate blank line and still keeps its
/// own line — the control that stops the rule above from dropping every empty
/// paragraph.
#[test]
fn an_ordinary_empty_paragraph_keeps_its_line() {
    let doc = parse_body(
        r#"<w:p><w:r><w:t>Before.</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:t>After.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Before.", "", "After."]);
}

/// The reported shape: item 2 of a four-item numbered list is moved below item
/// 3. The origin keeps its `w:numPr`, so leaving it in place both drew a bare
/// number and pushed every later item one too high. LibreOffice's `tdf149711`.
#[test]
fn a_moved_away_list_item_takes_no_number_at_its_origin() {
    let numbered_item = |mark: &str, content: &str| {
        format!(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>{mark}</w:pPr>{content}</w:p>"#
        )
    };
    let plain_item = |text: &str| numbered_item("", &format!("<w:r><w:t>{text}</w:t></w:r>"));
    let body = format!(
        "{}{}{}{}{}{}",
        plain_item("Item 1"),
        numbered_item(
            r#"<w:rPr><w:moveFrom w:id="0" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr>"#,
            r#"<w:moveFromRangeStart w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z" w:name="move1"/><w:moveFrom w:id="2" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:t>Item 2</w:t></w:r></w:moveFrom>"#,
        ),
        r#"<w:moveFromRangeEnd w:id="1"/>"#,
        plain_item("Item 3"),
        numbered_item(
            r#"<w:rPr><w:moveTo w:id="3" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr>"#,
            r#"<w:moveTo w:id="4" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:t>Item 2</w:t></w:r></w:moveTo>"#,
        ),
        plain_item("Item 4"),
    );

    let doc = parse_body(&body, Some(DECIMAL_NUMBERING));

    assert_eq!(
        line_texts(&doc),
        vec!["Item 1", "Item 3", "Item 2", "Item 4"],
        "the origin contributes no item, so the list runs 1 to 4"
    );
    let numbered_items: Vec<Option<u32>> = doc
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) => Some(flow.content.iter()),
            _ => None,
        })
        .flatten()
        .filter_map(|block| match block {
            Block::List(list) => Some(list.items.iter()),
            _ => None,
        })
        .flatten()
        .map(|item| item.start_at)
        .collect();
    assert_eq!(
        numbered_items,
        vec![Some(1), None, None, None],
        "one unbroken numbering run, so only the first item states a number"
    );
}

/// A deletion that spans a paragraph break: Word marks the tail of the first
/// paragraph, its mark, and the head of the second. The final view reads the
/// two surviving halves as one paragraph.
#[test]
fn a_removed_mark_merges_its_surviving_runs_into_the_next_paragraph() {
    let doc = parse_body(
        r#"<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr>
  <w:r><w:t xml:space="preserve">The head survives. </w:t></w:r>
  <w:del w:id="2" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:delText>Cut tail.</w:delText></w:r></w:del>
</w:p>
<w:p>
  <w:del w:id="3" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:r><w:delText>Cut head.</w:delText></w:r></w:del>
  <w:r><w:t>The tail survives.</w:t></w:r>
</w:p>"#,
        None,
    );

    assert_eq!(
        line_texts(&doc),
        vec!["The head survives. The tail survives."]
    );
}

/// The merged paragraph takes the surviving mark's formatting, because in Word
/// the paragraph mark is what carries it. Here that is the second paragraph's
/// centring, not the first paragraph's right alignment.
#[test]
fn a_merged_paragraph_takes_the_surviving_marks_formatting() {
    let doc = parse_body(
        r#"<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Head </w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>tail.</w:t></w:r></w:p>"#,
        None,
    );

    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected a flow page");
    };
    let paragraphs: Vec<&Paragraph> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    assert_eq!(paragraphs.len(), 1, "the two halves are one paragraph");
    assert_eq!(paragraph_text(paragraphs[0]), "Head tail.");
    assert_eq!(paragraphs[0].style.alignment, Some(Alignment::Center));
}

/// `w:rPrChange` records the run properties a revision replaced, so a
/// `w:moveFrom` inside it describes a past state of the mark rather than its
/// current one. `tdf123460` carries exactly this on a destination paragraph.
#[test]
fn a_past_revision_recorded_in_rprchange_does_not_remove_the_paragraph() {
    let doc = parse_body(
        r#"<w:p><w:r><w:t>Before.</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:rPrChange w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"><w:rPr><w:moveFrom w:id="2" w:author="Reviewer" w:date="2026-09-13T10:00:00Z"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>Still here.</w:t></w:r></w:p>
<w:p><w:r><w:t>After.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Before.", "Still here.", "After."]);
}

/// A removed mark with no paragraph after it cannot merge anywhere. Word never
/// writes that — a container's last mark is undeletable — so the surviving text
/// is kept rather than silently dropped.
#[test]
fn a_removed_mark_on_the_last_paragraph_keeps_its_text() {
    let doc = parse_body(
        r#"<w:p><w:r><w:t>Before.</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Orphaned tail.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Before.", "Orphaned tail."]);
}

/// A table between the two paragraphs is not something a paragraph merges
/// across, so the surviving runs stay in the flow before it instead of jumping
/// into the first cell.
#[test]
fn a_table_after_a_removed_mark_does_not_absorb_its_runs() {
    let doc = parse_body(
        r#"<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-09-14T10:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Orphaned tail.</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>After.</w:t></w:r></w:p>"#,
        None,
    );

    assert_eq!(line_texts(&doc), vec!["Orphaned tail.", "After."]);
    let Page::Flow(flow) = &doc.pages[0] else {
        panic!("expected a flow page");
    };
    let cell_texts: Vec<String> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Table(table) => Some(table.rows.iter()),
            _ => None,
        })
        .flatten()
        .flat_map(|row| row.cells.iter())
        .flat_map(|cell| cell.content.iter())
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph_text(paragraph)),
            _ => None,
        })
        .collect();
    assert_eq!(cell_texts, vec!["Cell."]);
}
