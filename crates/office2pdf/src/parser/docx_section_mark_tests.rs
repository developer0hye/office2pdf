//! An empty paragraph carrying a `w:sectPr` (a section mark): Word lays it out
//! as a line only when it is the first block of its section, or, in
//! compatibility mode 15, when it follows an inline table before a continuous
//! section. Elsewhere it takes no height, whatever its `w:pPr` holds
//! (issue #2232). Before a continuous section, such a mark still passes on a
//! gap: the previous paragraph's `w:after` plus `max(0, next before − mark
//! after)`.

use super::*;

/// An A4 section's page size and 72pt margins.
const PAGE_SETUP: &str = r#"<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>"#;

/// The `w:pPr` contents a section mark is tried with: bare, as in
/// docxide-pdf's `page-header-ordinary-marker`; with an after-spacing and a
/// 14pt mark, as in `paragraph-offset-frame-no-frame`; with a 1.5-line
/// spacing, as in `section-mark-after-table-*`; and centred.
const MARK_PROPERTIES: [&str; 4] = [
    "",
    r#"<w:spacing w:after="80"/><w:rPr><w:sz w:val="28"/></w:rPr>"#,
    r#"<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>"#,
    r#"<w:jc w:val="center"/>"#,
];

/// How the section after a mark starts.
const NEXT_SECTION_TYPES: [&str; 2] = ["continuous", "nextPage"];

fn text_paragraph(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

/// An empty paragraph with `properties` that ends a section.
fn section_mark(properties: &str) -> String {
    format!(r#"<w:p><w:pPr>{properties}<w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr></w:p>"#)
}

/// A one-cell table; `position` is its `w:tblpPr`, if it floats.
fn table(position: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr>{position}<w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        text_paragraph("Cell")
    )
}

const FLOATING_POSITION: &str =
    r#"<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpY="1"/>"#;

/// A package whose body closes with a section of type `last_section_type`,
/// declaring `compatibilityMode` `compatibility_mode` if one is given.
fn section_mark_docx(
    body: &str,
    last_section_type: &str,
    compatibility_mode: Option<u32>,
) -> Vec<u8> {
    let document: String = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:type w:val="{last_section_type}"/>{PAGE_SETUP}</w:sectPr></w:body></w:document>"#
    );
    let settings: Option<String> = compatibility_mode.map(|mode| {
        format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>"#
        )
    });
    let mut parts: Vec<(&str, &str)> = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", &document),
    ];
    if let Some(settings) = settings.as_deref() {
        parts.push(("word/settings.xml", settings));
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Each section's body blocks: a paragraph as its text, where an empty one,
/// which the renderer lays out as a blank line, shows as `""`, and a table as
/// `<table>` or `<floating table>`.
fn section_blocks(data: &[u8]) -> Vec<Vec<String>> {
    let (doc, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    doc.pages
        .iter()
        .filter_map(|page| match page {
            Page::Flow(flow) | Page::FlowContinuous(flow) => Some(&flow.content),
            _ => None,
        })
        .map(|content| {
            content
                .iter()
                .map(|block| match block {
                    Block::Paragraph(paragraph) => paragraph
                        .runs
                        .iter()
                        .map(|run| run.text.as_str())
                        .collect::<String>(),
                    Block::Table(_) => "<table>".to_string(),
                    Block::FloatingTable(_) => "<floating table>".to_string(),
                    other => format!("{other:?}"),
                })
                .collect()
        })
        .collect()
}

/// A mark after other blocks of its section takes no line, whatever its
/// `w:pPr` holds and however the next section starts. Word's exports of
/// docxide-pdf's `paragraph-offset-frame-no-frame` (a 14pt mark with 4pt
/// after, after two empty paragraphs) and `section-mark-after-table-float-*`
/// (a 1.5-line mark after a floating table) give it zero height.
#[test]
fn an_empty_section_mark_after_other_blocks_takes_no_line() {
    let preceding_blocks: [(&str, String, &str); 3] = [
        ("a text paragraph", text_paragraph("Body"), "Body"),
        ("an empty paragraph", "<w:p/>".to_string(), ""),
        (
            "a floating table",
            table(FLOATING_POSITION),
            "<floating table>",
        ),
    ];
    for (preceding_name, preceding_xml, preceding_block) in &preceding_blocks {
        for properties in MARK_PROPERTIES {
            for next_type in NEXT_SECTION_TYPES {
                for mode in [14, 15] {
                    let body: String = format!(
                        "{preceding_xml}{}{}",
                        section_mark(properties),
                        text_paragraph("Next")
                    );
                    assert_eq!(
                        section_blocks(&section_mark_docx(&body, next_type, Some(mode))),
                        [vec![preceding_block.to_string()], vec!["Next".to_string()]],
                        "a mark with {properties:?} after {preceding_name}, before a {next_type} \
                         section in compatibility mode {mode}, must take no line"
                    );
                }
            }
        }
    }
}

/// A mark that is the first block of its section is laid out as a line. In
/// Word's export of docxide-pdf's `page-header-ordinary-marker`, a bare mark
/// that is its section's only block leaves one line above the next section's
/// text. Zero-width bookmarks before it are no block.
#[test]
fn an_empty_section_mark_opening_its_section_keeps_its_line() {
    for properties in MARK_PROPERTIES {
        for next_type in NEXT_SECTION_TYPES {
            let leading: String = format!(
                r#"<w:bookmarkStart w:id="0" w:name="start"/><w:bookmarkEnd w:id="0"/>{}{}"#,
                section_mark(properties),
                text_paragraph("Next")
            );
            assert_eq!(
                section_blocks(&section_mark_docx(&leading, next_type, None)),
                [vec![String::new()], vec!["Next".to_string()]],
                "a mark with {properties:?} opening the document must keep its line"
            );

            let later: String = format!(
                r#"{}<w:p><w:pPr><w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr><w:r><w:t>Closing</w:t></w:r></w:p>{}{}"#,
                text_paragraph("First"),
                section_mark(properties),
                text_paragraph("Next")
            );
            assert_eq!(
                section_blocks(&section_mark_docx(&later, next_type, None)),
                [
                    vec!["First".to_string(), "Closing".to_string()],
                    vec![String::new()],
                    vec!["Next".to_string()]
                ],
                "a mark with {properties:?} opening a later section must keep its line"
            );
        }
    }
}

/// In compatibility mode 15, Word keeps a mark's line after an inline table
/// when the next section continues on the same page: its export of
/// docxide-pdf's `section-mark-after-table-inline-compat15` sets the next
/// text a 1.5-line mark's line below the table. Mode 14, a floating table and
/// a section starting on a new page each drop the line again.
#[test]
fn an_empty_section_mark_keeps_its_line_after_an_inline_table_only_in_compatibility_mode_15() {
    for properties in MARK_PROPERTIES {
        let after_inline: String = format!(
            "{}{}{}",
            table(""),
            section_mark(properties),
            text_paragraph("Next")
        );
        assert_eq!(
            section_blocks(&section_mark_docx(&after_inline, "continuous", Some(15))),
            [
                vec!["<table>".to_string(), String::new()],
                vec!["Next".to_string()]
            ],
            "a mark with {properties:?} after an inline table keeps its line in mode 15"
        );
        for (mode, next_type) in [(14, "continuous"), (15, "nextPage")] {
            assert_eq!(
                section_blocks(&section_mark_docx(&after_inline, next_type, Some(mode))),
                [vec!["<table>".to_string()], vec!["Next".to_string()]],
                "a mark with {properties:?} after an inline table before a {next_type} \
                 section in mode {mode} must take no line"
            );
        }
        // The next section's own `w:type` decides, not the mark's.
        let continuous_mark: String = format!(
            r#"{}<w:p><w:pPr>{properties}<w:sectPr><w:type w:val="continuous"/>{PAGE_SETUP}</w:sectPr></w:pPr></w:p>{}"#,
            table(""),
            text_paragraph("Next")
        );
        assert_eq!(
            section_blocks(&section_mark_docx(&continuous_mark, "nextPage", Some(15))),
            [vec!["<table>".to_string()], vec!["Next".to_string()]],
            "a continuous mark with {properties:?} before a new-page section must take no line"
        );
    }
}

/// A mark that is not empty is an ordinary line of text, wherever it sits.
#[test]
fn a_section_mark_with_text_keeps_its_line() {
    let body: String = format!(
        r#"{}<w:p><w:pPr><w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr><w:r><w:t>Mark</w:t></w:r></w:p>{}"#,
        text_paragraph("Body"),
        text_paragraph("Next")
    );
    assert_eq!(
        section_blocks(&section_mark_docx(&body, "continuous", Some(15))),
        [
            vec!["Body".to_string(), "Mark".to_string()],
            vec!["Next".to_string()]
        ]
    );
}

/// Each page's painted lines, as their text.
#[cfg(not(target_arch = "wasm32"))]
fn page_texts(data: &[u8]) -> Vec<Vec<String>> {
    let (document, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    let output = crate::render::typst_gen::generate_typst(&document).expect("document generates");
    let pages: u32 =
        crate::render::pdf::compile_page_count_with_fonts(&output.source, &output.images, &[], &[])
            .expect("document compiles");
    (0..pages as usize)
        .map(|page| {
            crate::render::pdf::compiled_text_runs(&output.source, page)
                .expect("page compiles")
                .into_iter()
                .map(|run| run.text.trim().to_string())
                .filter(|text| !text.is_empty())
                .collect()
        })
        .collect()
}

/// A page break that ends a section beginning on the same page as the one
/// before it still moves the text after it to the next page when the next
/// section also continues there (issue #2370). That holds whether the break
/// ends a paragraph before an empty mark that takes no line, or sits in the
/// mark itself, and in one or two columns.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_page_break_ending_a_continuous_section_moves_the_next_section_to_a_new_page() {
    let break_run: &str = r#"<w:r><w:br w:type="page"/></w:r>"#;
    for columns in ["", r#"<w:cols w:num="2" w:space="720"/>"#] {
        let continuous: String = format!(r#"<w:type w:val="continuous"/>{PAGE_SETUP}{columns}"#);
        let mut endings: Vec<(String, String)> = MARK_PROPERTIES
            .iter()
            .map(|properties| {
                (
                    format!("a break before a mark with {properties:?}"),
                    format!(
                        r#"<w:p><w:r><w:t>Before</w:t></w:r>{break_run}</w:p><w:p><w:pPr>{properties}<w:sectPr>{continuous}</w:sectPr></w:pPr></w:p>"#
                    ),
                )
            })
            .collect();
        endings.push((
            "a break in the mark".to_string(),
            format!(
                r#"{}<w:p><w:pPr><w:sectPr>{continuous}</w:sectPr></w:pPr>{break_run}</w:p>"#,
                text_paragraph("Before")
            ),
        ));
        for (ending, second_section) in endings {
            let body: String = format!(
                r#"<w:p><w:pPr><w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>{second_section}{}"#,
                text_paragraph("After")
            );
            assert_eq!(
                page_texts(&section_mark_docx(&body, "continuous", None)),
                [vec!["First", "Before"], vec!["After"]],
                "{ending} in columns {columns:?} must start the next section on a new page"
            );
        }
    }
}

/// A package of `body` closed by a continuous section with `columns`
/// (`w:cols` markup, or none), whose numbering part holds a decimal list,
/// `w:numId` 1.
fn continuous_section_docx(body: &str, columns: &str) -> Vec<u8> {
    let document: String = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:type w:val="continuous"/>{PAGE_SETUP}{columns}</w:sectPr></w:body></w:document>"#
    );
    let parts: [(&str, &str); 5] = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#,
        ),
        (
            "word/numbering.xml",
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#,
        ),
        ("word/document.xml", &document),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A `w:spacing` stating `before` and `after`, in points.
fn spacing(before: f64, after: f64) -> String {
    format!(
        r#"<w:spacing w:before="{}" w:after="{}"/>"#,
        before * 20.0,
        after * 20.0
    )
}

/// "Before" with `after` points after it, then an empty mark with
/// `mark_after` points after it, ending its section; then the continuous
/// section's "Next" with `next_before` points before it, as a list item when
/// `listed`.
fn spaced_section_mark_body(after: f64, mark_after: f64, next_before: f64, listed: bool) -> String {
    let numbering: &str = if listed {
        r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#
    } else {
        ""
    };
    format!(
        r#"<w:p><w:pPr>{}</w:pPr><w:r><w:t>Before</w:t></w:r></w:p>{}<w:p><w:pPr>{numbering}{}</w:pPr><w:r><w:t>Next</w:t></w:r></w:p>"#,
        spacing(0.0, after),
        section_mark(&spacing(0.0, mark_after)),
        spacing(next_before, 0.0)
    )
}

/// The distance from the baseline of page one's "Before" down to that of
/// its "Next".
#[cfg(not(target_arch = "wasm32"))]
fn before_to_next_pt(data: &[u8]) -> f64 {
    let (document, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    let output = crate::render::typst_gen::generate_typst(&document).expect("document generates");
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).expect("page compiles");
    let baseline = |text: &str| -> f64 {
        runs.iter()
            .find(|run| run.text.trim() == text)
            .unwrap_or_else(|| panic!("page one paints {text:?}"))
            .baseline_pt
    };
    baseline("Next") - baseline("Before")
}

/// Before a continuous section, an empty mark that takes no line passes on the
/// whole `w:after` of the paragraph before it, and only the part of the next
/// paragraph's `w:before` beyond the mark's own `w:after`: the gap is
/// `after + max(0, next before − mark after)`, not the larger of the two
/// neighbours' gaps. Word's export of docxide-pdf's
/// `scraped/covid_insomnia_mental_health` starts its two-column section's
/// numbered heading 22pt below the paragraph before the mark (12pt after, a
/// 0pt mark, 10pt before), and docxide-pdf's Word probe of 12pt, 8pt and 10pt
/// gives 14pt. That holds in one column or two, in a section that later
/// breaks the page, and for a plain paragraph or a list item.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_dropped_section_mark_absorbs_only_its_own_after_of_the_next_before() {
    // (after, mark after, next before) and Word's gap beyond the line pitch.
    let cases: [(f64, f64, f64, f64); 8] = [
        (12.0, 8.0, 10.0, 14.0),
        (12.0, 0.0, 10.0, 22.0),
        (4.0, 0.0, 10.0, 14.0),
        (6.0, 2.0, 4.0, 8.0),
        (0.0, 0.0, 10.0, 10.0),
        (10.0, 20.0, 15.0, 10.0),
        (10.0, 10.0, 0.0, 10.0),
        (0.0, 4.0, 0.0, 0.0),
    ];
    let page_break: &str = r#"<w:p><w:r><w:br w:type="page"/><w:t>Later</w:t></w:r></w:p>"#;
    let layouts: [(&str, &str, &str); 3] = [
        ("one column", "", ""),
        ("two columns", r#"<w:cols w:num="2" w:space="720"/>"#, ""),
        ("a section breaking the page later", "", page_break),
    ];
    for (layout, columns, rest) in layouts {
        for listed in [false, true] {
            let pitch: f64 = before_to_next_pt(&continuous_section_docx(
                &(spaced_section_mark_body(0.0, 0.0, 0.0, listed) + rest),
                columns,
            ));
            for (after, mark_after, next_before, gap) in cases {
                let measured: f64 = before_to_next_pt(&continuous_section_docx(
                    &(spaced_section_mark_body(after, mark_after, next_before, listed) + rest),
                    columns,
                )) - pitch;
                assert!(
                    (measured - gap).abs() < 0.01,
                    "{after}pt after, a {mark_after}pt mark and {next_before}pt before \
                     ({layout}, listed {listed}) must leave {gap}pt, got {measured}pt"
                );
            }
        }
    }
}

/// A mark after a page break opens the next section at the top of the new
/// page, where the next paragraph's `w:before` is still dropped: LibreOffice's
/// `tdf170119_bottomSpacing` test expects no gap there.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_dropped_section_mark_after_a_page_break_leaves_the_next_page_top_alone() {
    let next_baseline = |next_before: f64| -> f64 {
        let body: String = format!(
            r#"<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>{}<w:p><w:pPr>{}</w:pPr><w:r><w:t>Next</w:t></w:r></w:p>"#,
            section_mark(&spacing(0.0, 0.0)),
            spacing(next_before, 0.0)
        );
        let (document, _warnings) = DocxParser
            .parse(
                &continuous_section_docx(&body, ""),
                &ConvertOptions::default(),
            )
            .expect("document parses");
        let output =
            crate::render::typst_gen::generate_typst(&document).expect("document generates");
        crate::render::pdf::compiled_text_runs(&output.source, 1)
            .expect("page two compiles")
            .iter()
            .find(|run| run.text.trim() == "Next")
            .expect("page two paints Next")
            .baseline_pt
    };
    assert!((next_baseline(10.0) - next_baseline(0.0)).abs() < 0.01);
}
