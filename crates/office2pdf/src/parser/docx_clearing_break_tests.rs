//! A text-wrapping break that clears floats (`w:br w:clear`): Word resumes
//! the line after it below the floating table beside it, where a plain break
//! resumes on the next line beside the table (issue #2230).

use super::*;

const MARGIN_PT: f64 = 72.0;
/// A4's 11906 twips less two 1440-twip margins.
const TEXT_WIDTH_PT: f64 = 451.3;
/// The `leftFromText`/`rightFromText` every table here states: 180 twips.
const FROM_TEXT_PT: f64 = 9.0;
const TABLE_WIDTH_TWIPS: u32 = 4820;

fn run(content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>{content}</w:r>"#
    )
}

fn text(content: &str) -> String {
    run(&format!("<w:t>{content}</w:t>"))
}

fn line_break(attributes: &str) -> String {
    run(&format!("<w:br {attributes}/>"))
}

fn clearing_break(clear: &str) -> String {
    line_break(&format!(r#"w:type="textWrapping" w:clear="{clear}""#))
}

fn paragraph_of(paragraph_properties: &str, runs: &[String]) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{paragraph_properties}</w:pPr>{}</w:p>"#,
        runs.concat()
    )
}

fn paragraph(content: &str) -> String {
    paragraph_of("", &[text(content)])
}

/// A one-cell floating table whose single row is exactly `height_twips` tall,
/// anchored to the text it precedes.
fn floating_table(height_twips: u32, position: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" {position}/><w:tblW w:w="{TABLE_WIDTH_TWIPS}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="{TABLE_WIDTH_TWIPS}"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="{height_twips}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{TABLE_WIDTH_TWIPS}" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph("Floating cell")
    )
}

/// An A4 package with 72pt margins, no settings part and no styles part.
fn clear_docx(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let parts: [(&str, &str); 3] = [
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
    for (name, content) in parts {
        zip.start_file(name, options).unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn body_paragraphs(data: &[u8]) -> Vec<Paragraph> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected a flow page");
    };
    page.content
        .iter()
        .filter_map(|block| match block {
            crate::ir::Block::Paragraph(paragraph) => Some(paragraph.clone()),
            _ => None,
        })
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn typst_source(data: &[u8]) -> String {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source
}

/// The painted body runs; the table's own cell text is left out.
#[cfg(not(target_arch = "wasm32"))]
fn placed(data: &[u8]) -> Vec<(crate::render::pdf::PlacedTextRun, f64)> {
    let source = typst_source(data);
    let text = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source, 0, true).unwrap();
    text.into_iter()
        .filter(|run| !run.text.trim().is_empty() && run.text != "Floating cell")
        .map(|run| {
            let width = glyphs
                .iter()
                .find(|glyph| glyph.text == run.text)
                .unwrap_or_else(|| panic!("no glyph run {:?}", run.text))
                .width_pt();
            (run, width)
        })
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn find<'a>(
    runs: &'a [(crate::render::pdf::PlacedTextRun, f64)],
    text: &str,
) -> &'a (crate::render::pdf::PlacedTextRun, f64) {
    runs.iter()
        .find(|(run, _)| run.text == text)
        .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_close(actual: f64, expected: f64, context: &str) {
    assert!(
        (actual - expected).abs() < 0.05,
        "{context}: {actual} != {expected}"
    );
}

/// The first line's baseline at the top of the page and the pitch between
/// two lines, from a document with no table.
#[cfg(not(target_arch = "wasm32"))]
fn page_top_baseline_and_line_pitch() -> (f64, f64) {
    let runs = placed(&clear_docx(
        &[paragraph("First"), paragraph("Second")].concat(),
    ));
    let (first, _) = find(&runs, "First");
    let (second, _) = find(&runs, "Second");
    (first.baseline_pt, second.baseline_pt - first.baseline_pt)
}

/// The left edge of the room right of a table at the left margin.
const WRAP_EDGE_PT: f64 = MARGIN_PT + TABLE_WIDTH_TWIPS as f64 / 20.0 + FROM_TEXT_PT;

#[test]
fn a_clearing_break_reaches_the_paragraph_as_a_line_break_that_clears() {
    let data = clear_docx(&paragraph_of(
        "",
        &[
            text("One"),
            line_break(""),
            text("Two"),
            clearing_break("all"),
            text("Three"),
            line_break(r#"w:clear="none""#),
            clearing_break("left"),
            clearing_break("right"),
            text("Four"),
        ],
    ));
    let paragraphs = body_paragraphs(&data);
    let text: String = paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect();
    assert_eq!(text, "One\nTwo\nThree\n\n\nFour");
    assert_eq!(
        &*paragraphs[0].style.clearing_breaks,
        &[
            ClearingBreak {
                line_break: 1,
                clear: BreakClear::All,
            },
            ClearingBreak {
                line_break: 3,
                clear: BreakClear::Left,
            },
            ClearingBreak {
                line_break: 4,
                clear: BreakClear::Right,
            },
        ]
    );
}

/// `w:clear` on a page break does not make it a line break, and a paragraph
/// split at that page break counts its line breaks from its own start.
#[test]
fn a_page_break_neither_clears_nor_shifts_the_count() {
    let data = clear_docx(&paragraph_of(
        "",
        &[
            text("One"),
            line_break(""),
            line_break(r#"w:type="page" w:clear="all""#),
            text("Two"),
            line_break(""),
            clearing_break("all"),
            text("Three"),
        ],
    ));
    let paragraphs = body_paragraphs(&data);
    assert!(paragraphs[0].style.clearing_breaks.is_empty());
    let last = paragraphs.last().unwrap();
    assert_eq!(
        &*last.style.clearing_breaks,
        &[ClearingBreak {
            line_break: 1,
            clear: BreakClear::All,
        }]
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_after_a_clearing_break_resumes_below_a_floating_table() {
    let (top_baseline, pitch) = page_top_baseline_and_line_pitch();
    for height_twips in [800_u32, 1600] {
        let data = clear_docx(
            &[
                floating_table(height_twips, ""),
                paragraph_of(
                    r#"<w:jc w:val="center"/>"#,
                    &[text("Before"), clearing_break("all"), text("After")],
                ),
                paragraph("Next"),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let context = format!("{height_twips}-twip float");
        let (before, before_width) = find(&runs, "Before");
        assert_close(
            before.left_pt + before_width / 2.0,
            (WRAP_EDGE_PT + MARGIN_PT + TEXT_WIDTH_PT) / 2.0,
            &format!("{context}: the line before the break stays beside the table"),
        );
        // The line after the break starts at the table's bottom edge and is
        // centred on the whole column again.
        let (after, after_width) = find(&runs, "After");
        assert_close(
            after.baseline_pt,
            top_baseline + f64::from(height_twips) / 20.0,
            &context,
        );
        assert_close(
            after.left_pt + after_width / 2.0,
            MARGIN_PT + TEXT_WIDTH_PT / 2.0,
            &context,
        );
        let (next, _) = find(&runs, "Next");
        assert_close(next.left_pt, MARGIN_PT, &context);
        assert_close(next.baseline_pt, after.baseline_pt + pitch, &context);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn each_further_clearing_break_adds_one_ordinary_line() {
    let (top_baseline, pitch) = page_top_baseline_and_line_pitch();
    let height_twips: u32 = 1200;
    for break_count in [1_usize, 2, 3] {
        let mut runs: Vec<String> = vec![text("Before")];
        runs.extend(std::iter::repeat_n(clearing_break("all"), break_count));
        runs.push(text("After"));
        let data =
            clear_docx(&[floating_table(height_twips, ""), paragraph_of("", &runs)].concat());
        let placed_runs = placed(&data);
        let (after, _) = find(&placed_runs, "After");
        let context = format!("{break_count} clearing breaks");
        assert_close(after.left_pt, MARGIN_PT, &context);
        assert_close(
            after.baseline_pt,
            top_baseline + f64::from(height_twips) / 20.0 + (break_count - 1) as f64 * pitch,
            &context,
        );
    }
}

/// With no table beside its line, a clearing break is a plain line break: the
/// document compiles to the same markup, and a line already below the table
/// stays where a plain break puts it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_clearing_break_with_no_float_beside_its_line_is_a_plain_break() {
    let lines = |last_break: String| {
        paragraph_of(
            "",
            &[
                text("Line 1"),
                line_break(""),
                text("Line 2"),
                line_break(""),
                text("Line 3"),
                last_break,
                text("Line 4"),
            ],
        )
    };
    assert_eq!(
        typst_source(&clear_docx(&lines(clearing_break("all")))),
        typst_source(&clear_docx(&lines(line_break("")))),
        "no table: the clear changes nothing"
    );

    // A 20pt table: the fourth line is well below it.
    let table: String = floating_table(400, "");
    let cleared = placed(&clear_docx(
        &[table.clone(), lines(clearing_break("all"))].concat(),
    ));
    let plain = placed(&clear_docx(&[table, lines(line_break(""))].concat()));
    let (cleared_line, _) = find(&cleared, "Line 4");
    let (plain_line, _) = find(&plain, "Line 4");
    assert_close(cleared_line.baseline_pt, plain_line.baseline_pt, "Line 4");
    assert_close(cleared_line.left_pt, MARGIN_PT, "Line 4 below the table");
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_break_that_clears_nothing_keeps_the_next_line_beside_the_table() {
    let (top_baseline, pitch) = page_top_baseline_and_line_pitch();
    for break_attributes in ["", r#"w:type="textWrapping" w:clear="none""#] {
        let data = clear_docx(
            &[
                floating_table(1600, ""),
                paragraph_of(
                    "",
                    &[text("Before"), line_break(break_attributes), text("After")],
                ),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let (after, _) = find(&runs, "After");
        assert_close(after.left_pt, WRAP_EDGE_PT, break_attributes);
        assert_close(after.baseline_pt, top_baseline + pitch, break_attributes);
    }
}

/// `w:clear="left"` resumes below a float on the line's left, which a table at
/// the left margin is, and `"right"` below one on its right.
///
/// FIXME(unprobed one-sided clear): the Word references at hand only clear
/// `all`; this follows ECMA-376's ST_BrClear wording.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_one_sided_clearing_break_clears_only_a_table_on_that_side() {
    let (top_baseline, pitch) = page_top_baseline_and_line_pitch();
    let height_twips: u32 = 1600;
    let below: f64 = top_baseline + f64::from(height_twips) / 20.0;
    let beside: f64 = top_baseline + pitch;
    for (position, clear, expected_baseline) in [
        ("", "left", below),
        ("", "right", beside),
        (r#"w:tblpXSpec="right""#, "right", below),
        (r#"w:tblpXSpec="right""#, "left", beside),
    ] {
        let data = clear_docx(
            &[
                floating_table(height_twips, position),
                paragraph_of("", &[text("Before"), clearing_break(clear), text("After")]),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let (after, _) = find(&runs, "After");
        assert_close(
            after.baseline_pt,
            expected_baseline,
            &format!("clear={clear} beside a table at {position:?}"),
        );
    }
}

/// docxide-pdf's `leading-table-clear-leading-clear`, the issue's case. Its
/// greeting follows a clearing break: Word sets it below the table, centred on
/// the whole column, and the body paragraph two lines below at the margin.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn issue_case_greeting_after_a_clearing_break_lies_below_its_floating_table() {
    let data: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_table_clear_break.docx");
    let source = typst_source(data);
    let runs: Vec<crate::render::pdf::PlacedTextRun> =
        crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source, 0, true).unwrap();
    let find_run = |text: &str| {
        runs.iter()
            .find(|run| run.text == text)
            .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
    };
    let topic = find_run("Public floating topic");
    let second = find_run("Second topic line");
    let pitch: f64 = second.baseline_pt - topic.baseline_pt;
    // The table's four 20pt cell lines end two lines below its second topic
    // line, where the greeting resumes.
    let greeting = find_run("Public greeting");
    assert_close(
        greeting.baseline_pt,
        second.baseline_pt + 2.0 * pitch,
        "greeting",
    );
    let greeting_width: f64 = glyphs
        .iter()
        .find(|glyph| glyph.text == "Public greeting")
        .unwrap()
        .width_pt();
    assert_close(
        greeting.left_pt + greeting_width / 2.0,
        MARGIN_PT + TEXT_WIDTH_PT / 2.0,
        "greeting centred on the column",
    );
    let body = find_run("Public body paragraph");
    assert_close(body.left_pt, MARGIN_PT, "body paragraph");
    assert_close(
        body.baseline_pt,
        greeting.baseline_pt + 2.0 * pitch,
        "body paragraph",
    );
}
