//! A `PAGE` or `NUMPAGES` field in a header or footer table shows the page it
//! is drawn on, as Word computes it, rather than the result Word cached when
//! it last saved the story (issue #2338).
use super::*;
use std::io::Cursor;

const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// Arial throughout: a package that names no face loses its exact line boxes
/// on CI's Linux runner.
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#;

/// A field written as Word writes one: begin, the instruction, separate, the
/// cached result and end, each in a run of its own.
fn field(instruction: &str, cached: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>{cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn text(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

/// A one-row story table whose cells hold one paragraph of `cells` each, then
/// the paragraph Word ends every story with.
fn table_story(cells: &[&str]) -> String {
    let grid: String = cells.iter().map(|_| r#"<w:gridCol w:w="3000"/>"#).collect();
    let cells: String = cells
        .iter()
        .map(|runs| {
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>{runs}</w:p></w:tc>"#
            )
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl><w:p/>"#
    )
}

/// A section's `w:sectPr` drawing the package's header and footer, numbered
/// as `page_numbering` (a `w:pgNumType`, or nothing) states.
fn section_properties(page_numbering: &str) -> String {
    format!(
        r#"<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="708" w:gutter="0"/>{page_numbering}</w:sectPr>"#
    )
}

const PAGE_BREAK: &str = r#"<w:r><w:br w:type="page"/></w:r>"#;

/// Three pages of body text in one section, numbered as `page_numbering`
/// states.
fn three_page_body(page_numbering: &str) -> String {
    format!(
        r#"<w:p><w:r><w:t>First</w:t></w:r>{PAGE_BREAK}<w:r><w:t>Second</w:t></w:r>{PAGE_BREAK}<w:r><w:t>Third</w:t></w:r></w:p>{}"#,
        section_properties(page_numbering)
    )
}

/// `body` under `header` and `footer`.
fn build_package(header: &str, footer: &str, body: &str) -> Vec<u8> {
    let document: String = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NAMESPACES}><w:body>{body}</w:body></w:document>"#
    );
    let parts: Vec<(&str, String)> = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"#
                .to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#
                .to_string(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#
                .to_string(),
        ),
        ("word/document.xml", document),
        ("word/styles.xml", STYLES.to_string()),
        (
            "word/header1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr {NAMESPACES}>{header}</w:hdr>"#
            ),
        ),
        (
            "word/footer1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr {NAMESPACES}>{footer}</w:ftr>"#
            ),
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn parse(header: &str, footer: &str) -> Document {
    DocxParser
        .parse(
            &build_package(header, footer, &three_page_body("")),
            &ConvertOptions::default(),
        )
        .unwrap()
        .0
}

/// The text of each paragraph in each cell of `story`'s tables.
fn story_table_cell_texts(story: &HeaderFooter) -> Vec<String> {
    story
        .shapes
        .iter()
        .filter_map(|shape| match &shape.content {
            HeaderFooterShapeContent::Table(table) => Some(table),
            _ => None,
        })
        .flat_map(|table| &table.rows)
        .flat_map(|row| &row.cells)
        .flat_map(|cell| &cell.content)
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
        .collect()
}

fn header_cell_texts(document: &Document) -> Vec<String> {
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected a flow page");
    };
    story_table_cell_texts(page.header.as_ref().expect("the section has a header"))
}

#[test]
fn a_story_table_page_field_shows_the_page_rather_than_its_cached_result() {
    let page: String = PAGE_NUMBER_FIELD_CHAR.to_string();
    let total: String = TOTAL_PAGES_FIELD_CHAR.to_string();
    for (cell, expected) in [
        // docxide-pdf `zimbabwe_gold_mining_judgment`'s header cell.
        (field(r"PAGE   \* MERGEFORMAT", "2"), page.clone()),
        (field("PAGE", "17"), page.clone()),
        (field("NUMPAGES", "8"), total.clone()),
        (
            format!(
                "{}{}{}{}",
                text("Page "),
                field("PAGE", "3"),
                text(" of "),
                field(r"NUMPAGES \* Arabic", "12")
            ),
            format!("Page {page} of {total}"),
        ),
    ] {
        let document: Document = parse(&table_story(&[&cell]), "");
        assert_eq!(
            header_cell_texts(&document),
            vec![expected.clone()],
            "header cell {cell}"
        );
        let document: Document = parse("", &table_story(&[&cell]));
        let Page::Flow(page) = &document.pages[0] else {
            panic!("expected a flow page");
        };
        assert_eq!(
            story_table_cell_texts(page.footer.as_ref().expect("the section has a footer")),
            vec![expected],
            "footer cell {cell}"
        );
    }
}

/// Only `PAGE` and `NUMPAGES` are computed. Every other field, including one
/// nested in another's instruction or result as the clean-energy act's footer
/// `IF` holds a `DOCPROPERTY`, keeps the text Word cached, and a page field
/// after it in the same paragraph still resolves.
#[test]
fn a_story_table_keeps_the_cached_result_of_any_other_field() {
    let nested_if: String = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve"> = "" "none" "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>17/10/2023</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("DOCPROPERTY RegisteredDate", "17 October 2023"),
        field(
            r#"DOCPROPERTY RegisteredDate \@ "dd/MM/yyyy""#,
            "17/10/2023"
        ),
    );
    let page: char = PAGE_NUMBER_FIELD_CHAR;
    for (cell, expected) in [
        (
            field("DOCPROPERTY ShortT", "Short Title Act 2023"),
            "Short Title Act 2023".to_string(),
        ),
        (field("PAGEREF _Toc1 \\h", "4"), "4".to_string()),
        (
            format!(
                "{}{}",
                field("STYLEREF ShortT", "Rule 2015"),
                field("PAGE", "i")
            ),
            format!("Rule 2015{page}"),
        ),
    ] {
        assert_eq!(
            header_cell_texts(&parse(&table_story(&[&cell]), "")),
            vec![expected],
            "header cell {cell}"
        );
    }

    let cell: String = format!("{nested_if}{}{}", text(" page "), field("PAGE", "9"));
    let texts: Vec<String> = header_cell_texts(&parse(&table_story(&[&cell]), ""));
    assert_eq!(texts.len(), 1, "{texts:?}");
    assert!(
        texts[0].contains("17/10/2023") && texts[0].ends_with(&format!(" page {page}")),
        "the IF keeps its result and the PAGE after it resolves, got {texts:?}"
    );
    assert_eq!(texts[0].matches(page).count(), 1, "{texts:?}");
}

/// A field inside another field's instruction is part of that field's code,
/// not of the page: the `IF` of docxide-pdf's `dutch_government_budget_letter`
/// first-page footer holds `PAGE` and `NUMPAGES` in its text arguments, and
/// Word prints only the `IF`'s own result. Its cached result here is empty.
#[test]
fn a_page_field_inside_another_fields_instruction_shows_no_page() {
    let instruction = |text: &str| -> String {
        format!(r#"<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>"#)
    };
    let nested = |name: &str, cached: &str| -> String {
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>{}<w:r><w:fldChar w:fldCharType="separate"/></w:r>{}<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            instruction(&format!(" {name} ")),
            instruction(cached)
        )
    };
    let if_field: String = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>{}{}{}{}{}{}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        instruction(r#" if "#),
        nested("NUMPAGES", "1"),
        instruction(r#" = "1" "" "Pagina "#),
        nested("PAGE", "1"),
        instruction(" van "),
        nested("NUMPAGES", "2"),
    );
    let page: char = PAGE_NUMBER_FIELD_CHAR;
    assert_eq!(
        header_cell_texts(&parse(&table_story(&[&if_field]), "")),
        vec![String::new()]
    );
    // A page field after the `IF` is the page's own again.
    let cell: String = format!("{if_field}{}", field("PAGE", "4"));
    assert_eq!(
        header_cell_texts(&parse(&table_story(&[&cell]), "")),
        vec![page.to_string()]
    );
}

/// A dirty field Word left without a cached result still counts the page.
#[test]
fn a_story_table_page_field_without_a_cached_result_still_shows_the_page() {
    let dirty: &str = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
    assert_eq!(
        header_cell_texts(&parse(&table_story(&[dirty]), "")),
        vec![PAGE_NUMBER_FIELD_CHAR.to_string()]
    );
}

#[cfg(not(target_arch = "wasm32"))]
mod rendered {
    use super::*;

    /// The text of every run each compiled page places, page by page.
    fn page_texts(data: &[u8]) -> Vec<Vec<String>> {
        let (document, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        (0..3)
            .map(|page_index| {
                crate::render::pdf::compiled_text_runs(&output.source, page_index)
                    .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source))
                    .into_iter()
                    .map(|run| run.text)
                    .collect()
            })
            .collect()
    }

    /// Each page prints its own number and the document's page count, in the
    /// section's numeral format. On docxide-pdf's
    /// `zimbabwe_gold_mining_judgment` Word's export prints each page's own
    /// number in the header table, where the cached `2` printed on every page.
    #[test]
    fn a_story_table_page_field_counts_each_page() {
        let header: String = table_story(&[
            &text("Judgment No 1"),
            &field(r"PAGE   \* MERGEFORMAT", "2"),
        ]);
        let footer: String = table_story(&[&format!(
            "{}{}{}{}",
            text("Page "),
            field("PAGE", "2"),
            text(" of "),
            field("NUMPAGES", "8")
        )]);
        // Under a restart the footer's page count is not checked: no Word
        // reference here shows what `NUMPAGES` prints there, and the table
        // draws the same count a story paragraph's `NUMPAGES` does.
        for (page_numbering, numerals, counts_pages) in [
            ("", ["1", "2", "3"], true),
            (
                r#"<w:pgNumType w:fmt="lowerRoman"/>"#,
                ["i", "ii", "iii"],
                true,
            ),
            (r#"<w:pgNumType w:start="5"/>"#, ["5", "6", "7"], false),
        ] {
            let pages = page_texts(&build_package(
                &header,
                &footer,
                &three_page_body(page_numbering),
            ));
            for (texts, numeral) in pages.iter().zip(numerals) {
                assert!(
                    texts.iter().any(|text| text == numeral),
                    "`{page_numbering}`: the header cell must print {numeral}, got {texts:?}"
                );
                let footer_line: String = if counts_pages {
                    format!("Page {numeral} of 3")
                } else {
                    format!("Page {numeral} of ")
                };
                assert!(
                    texts.concat().contains(&footer_line),
                    "`{page_numbering}`: the footer must print {footer_line:?}, got {texts:?}"
                );
                assert!(
                    !texts.concat().contains('8'),
                    "`{page_numbering}`: a cached result must not print, got {texts:?}"
                );
            }
        }
    }

    /// A section that restarts its numbering restarts the table's count on its
    /// own first page, where the page background is drawn before the restart:
    /// roman front matter, then the body from 1, as the carbon farming rule's
    /// and the clean energy act's footers number theirs.
    #[test]
    fn a_story_table_page_field_restarts_with_its_section() {
        let header: String = table_story(&[&field("PAGE", "i")]);
        let body: String = format!(
            r#"<w:p><w:r><w:t>Contents</w:t></w:r>{PAGE_BREAK}<w:r><w:t>Preface</w:t></w:r></w:p><w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>Part</w:t></w:r></w:p>{}"#,
            section_properties(r#"<w:pgNumType w:fmt="lowerRoman" w:start="1"/>"#),
            section_properties(r#"<w:pgNumType w:start="1"/>"#),
        );
        let pages: Vec<Vec<String>> = page_texts(&build_package(&header, "", &body));
        for (texts, numeral) in pages.iter().zip(["i", "ii", "1"]) {
            assert!(
                texts.iter().any(|text| text == numeral),
                "the header cell must print {numeral}, got {pages:?}"
            );
        }
    }
}
