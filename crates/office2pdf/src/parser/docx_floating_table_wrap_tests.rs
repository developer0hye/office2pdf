//! Body text beside a `w:tblpPr` floating table: Word lays the lines that
//! fall within the table's vertical extent in the room beside it, and moves
//! them below the table when neither side leaves room (issue #2229).

use super::*;

const MARGIN_PT: f64 = 72.0;
/// A4's 11906 twips less two 1440-twip margins.
const TEXT_WIDTH_PT: f64 = 451.3;
/// The `leftFromText`/`rightFromText` every table here states: 180 twips.
const FROM_TEXT_PT: f64 = 9.0;

fn run(content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>{content}</w:r>"#
    )
}

fn paragraph(paragraph_properties: &str, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{paragraph_properties}</w:pPr>{}</w:p>"#,
        run(&format!("<w:t>{text}</w:t>"))
    )
}

/// A one-cell floating table `width_twips` wide whose single row is exactly
/// `height_twips` tall, anchored to the text it precedes.
fn floating_table(width_twips: u32, height_twips: u32, position: &str) -> String {
    anchored_floating_table(
        width_twips,
        height_twips,
        &format!(r#"w:vertAnchor="text" {position}"#),
    )
}

/// [`floating_table`] with its whole `w:tblpPr` position stated.
fn anchored_floating_table(width_twips: u32, height_twips: u32, position: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" {position}/><w:tblW w:w="{width_twips}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="{width_twips}"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="{height_twips}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{width_twips}" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph("", "Floating cell")
    )
}

/// An A4 section's page size and 72pt margins.
const PAGE_SETUP: &str = r#"<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>"#;

/// An A4 package with 72pt margins, no settings part and no styles part.
fn float_docx(body: &str) -> Vec<u8> {
    section_float_docx(body, PAGE_SETUP)
}

/// [`float_docx`] with `section_properties` in the body's closing `w:sectPr`.
fn section_float_docx(body: &str, section_properties: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr>{section_properties}</w:sectPr></w:body></w:document>"#
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

/// The painted body runs with each run's advance width; the table's own
/// cell text is left out.
#[cfg(not(target_arch = "wasm32"))]
fn placed(data: &[u8]) -> Vec<(crate::render::pdf::PlacedTextRun, f64)> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
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

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_after_a_floating_table_is_laid_out_in_the_room_beside_it() {
    for width_twips in [3000_u32, 4820] {
        let wrap_edge: f64 = MARGIN_PT + f64::from(width_twips) / 20.0 + FROM_TEXT_PT;
        let data = float_docx(
            &[
                floating_table(width_twips, 1600, ""),
                paragraph("", "Beside left"),
                paragraph(r#"<w:jc w:val="center"/>"#, "Beside centred"),
                paragraph(r#"<w:jc w:val="right"/>"#, "Beside right"),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let context = format!("{width_twips}-twip float");
        let (left, _) = find(&runs, "Beside left");
        assert_close(left.left_pt, wrap_edge, &context);
        // Alignment resolves against the room beside the table, not the column.
        let (centred, centred_width) = find(&runs, "Beside centred");
        assert_close(
            centred.left_pt + centred_width / 2.0,
            (wrap_edge + MARGIN_PT + TEXT_WIDTH_PT) / 2.0,
            &context,
        );
        let (right, right_width) = find(&runs, "Beside right");
        assert_close(
            right.left_pt + right_width,
            MARGIN_PT + TEXT_WIDTH_PT,
            &context,
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_beside_a_floating_table_on_the_right_uses_the_room_on_its_left() {
    for width_twips in [3000_u32, 4820] {
        let wrap_edge: f64 =
            MARGIN_PT + TEXT_WIDTH_PT - f64::from(width_twips) / 20.0 - FROM_TEXT_PT;
        let data = float_docx(
            &[
                floating_table(width_twips, 1600, r#"w:tblpXSpec="right""#),
                paragraph("", "Beside left"),
                paragraph(r#"<w:jc w:val="right"/>"#, "Beside right"),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let context = format!("{width_twips}-twip float on the right");
        let (left, _) = find(&runs, "Beside left");
        assert_close(left.left_pt, MARGIN_PT, &context);
        let (right, right_width) = find(&runs, "Beside right");
        assert_close(right.left_pt + right_width, wrap_edge, &context);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn lines_below_a_floating_table_return_to_the_full_column() {
    for height_twips in [800_u32, 1200] {
        let band_bottom: f64 = MARGIN_PT + f64::from(height_twips) / 20.0;
        let wrap_edge: f64 = MARGIN_PT + 4820.0 / 20.0 + FROM_TEXT_PT;
        let lines: Vec<String> = (1..=10)
            .map(|line| paragraph("", &format!("Line {line}")))
            .collect();
        let data = float_docx(&[floating_table(4820, height_twips, ""), lines.concat()].concat());
        let runs = placed(&data);
        let (mut beside, mut below) = (0, 0);
        for (run, _) in &runs {
            // A baseline above the table's bottom puts the line top inside
            // it; a line whose top is a full em below the bottom is clear.
            if run.baseline_pt < band_bottom {
                assert_close(run.left_pt, wrap_edge, &run.text);
                beside += 1;
            } else if run.baseline_pt - run.font_size_pt > band_bottom {
                assert_close(run.left_pt, MARGIN_PT, &run.text);
                below += 1;
            }
        }
        assert!(
            beside > 0 && below > 0,
            "{height_twips}-twip float: {beside} lines beside, {below} below in {runs:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_reaching_a_margin_anchored_table_down_the_page_wraps_beside_it() {
    let wrap_edge: f64 = MARGIN_PT + 4820.0 / 20.0 + FROM_TEXT_PT;
    for (offset_twips, height_twips) in [(1600_u32, 1200_u32), (2400, 1600)] {
        let band_top: f64 = MARGIN_PT + f64::from(offset_twips) / 20.0;
        let band_bottom: f64 = band_top + f64::from(height_twips) / 20.0;
        let lines: Vec<String> = (1..=16)
            .map(|line| paragraph("", &format!("Line {line}")))
            .collect();
        let data = float_docx(
            &[
                anchored_floating_table(
                    4820,
                    height_twips,
                    &format!(r#"w:vertAnchor="margin" w:tblpY="{offset_twips}""#),
                ),
                lines.concat(),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let (mut above, mut beside, mut below) = (0, 0, 0);
        for (run, _) in &runs {
            // A line whose bottom clears the table's top, one inside the
            // table's extent, and one whose top clears its bottom.
            if run.baseline_pt + 0.3 * run.font_size_pt < band_top {
                assert_close(run.left_pt, MARGIN_PT, &run.text);
                above += 1;
            } else if run.baseline_pt > band_top + run.font_size_pt && run.baseline_pt < band_bottom
            {
                assert_close(run.left_pt, wrap_edge, &run.text);
                beside += 1;
            } else if run.baseline_pt - run.font_size_pt > band_bottom {
                assert_close(run.left_pt, MARGIN_PT, &run.text);
                below += 1;
            }
        }
        assert!(
            above > 0 && beside > 0 && below > 0,
            "{offset_twips}-twip offset: {above} above, {beside} beside, {below} below in {runs:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn wrapping_beside_a_floating_table_keeps_every_paragraph_spacing() {
    let spaced: Vec<String> = (1..=6)
        .map(|line| {
            format!(
                r#"<w:p><w:pPr><w:spacing w:before="{}" w:after="{}"/></w:pPr>{}</w:p>"#,
                // The first paragraph states no before-gap: a page-top
                // paragraph's before-gap is not what this test varies.
                (line - 1) * 40,
                line * 20,
                run(&format!("<w:t>Spaced {line}</w:t>"))
            )
        })
        .collect();
    let control = placed(&float_docx(&spaced.concat()));
    let data = float_docx(&[floating_table(4820, 1200, ""), spaced.concat()].concat());
    let runs = placed(&data);
    assert_eq!(runs.len(), control.len());
    for ((run, _), (expected, _)) in runs.iter().zip(&control) {
        assert_eq!(run.text, expected.text);
        assert_close(run.baseline_pt, expected.baseline_pt, &run.text);
    }
    assert!(
        runs.iter().any(|(run, _)| run.left_pt > MARGIN_PT + 1.0),
        "the first paragraphs must sit beside the table: {runs:?}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_moves_below_a_floating_table_that_leaves_no_room_beside_it() {
    let control = placed(&float_docx(&paragraph("", "Following")));
    let (control_run, _) = find(&control, "Following");
    // A table as wide as the column, and one leaving a strip too narrow for
    // a line beside it.
    for (width_twips, height_twips) in [(9026_u32, 800_u32), (9026, 1200), (8400, 1200)] {
        let data = float_docx(
            &[
                floating_table(width_twips, height_twips, ""),
                paragraph("", "Following"),
                paragraph("", "Next"),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let context = format!("{width_twips}x{height_twips}-twip float");
        let (following, _) = find(&runs, "Following");
        assert_close(following.left_pt, MARGIN_PT, &context);
        assert_close(
            following.baseline_pt,
            control_run.baseline_pt + f64::from(height_twips) / 20.0,
            &context,
        );
        let (next, _) = find(&runs, "Next");
        assert_close(next.left_pt, MARGIN_PT, &context);
        assert!(next.baseline_pt > following.baseline_pt, "{context}");
    }
}

/// A floating table that ends its section keeps the text of a continuous
/// section after it, laid out in the same text area, off its extent, as a
/// text frame does. Word's exports of docxide-pdf's
/// `section-mark-after-table-float-compat14` and `-compat15` start that
/// section's text right below a table as wide as the column.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_floating_table_keeps_text_off_its_extent_in_the_continuous_section_after_it() {
    let control = placed(&float_docx(&paragraph("", "Following")));
    let (control_run, _) = find(&control, "Following");
    let section_mark: String =
        format!(r#"<w:p><w:pPr><w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr></w:p>"#);
    let continuous: String = format!(r#"<w:type w:val="continuous"/>{PAGE_SETUP}"#);
    for (width_twips, height_twips) in [(3000_u32, 1200_u32), (9026, 800), (9026, 1200)] {
        let data = section_float_docx(
            &[
                floating_table(width_twips, height_twips, ""),
                section_mark.clone(),
                paragraph("", "Following"),
            ]
            .concat(),
            &continuous,
        );
        let runs = placed(&data);
        let context = format!("{width_twips}x{height_twips}-twip float");
        let (following, _) = find(&runs, "Following");
        if width_twips == 3000 {
            assert_close(
                following.left_pt,
                MARGIN_PT + f64::from(width_twips) / 20.0 + FROM_TEXT_PT,
                &context,
            );
            assert_close(following.baseline_pt, control_run.baseline_pt, &context);
        } else {
            assert_close(following.left_pt, MARGIN_PT, &context);
            assert_close(
                following.baseline_pt,
                control_run.baseline_pt + f64::from(height_twips) / 20.0,
                &context,
            );
        }
    }
}

/// docxide-pdf's `leading-table-clear-leading-no-clear`, the issue's case:
/// Word centres the greeting in the room right of the table and starts the
/// body paragraph at that room's left edge.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn issue_case_greeting_and_body_lie_beside_their_floating_table() {
    let data: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_table_wrap.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected a flow page");
    };
    let table = page
        .content
        .iter()
        .find_map(|block| match block {
            crate::ir::Block::FloatingTable(table) => Some(table),
            _ => None,
        })
        .expect("the table keeps its floating frame");
    let wrap_edge: f64 = MARGIN_PT
        + table.frame.x.unwrap_or_default()
        + table.table.column_widths.iter().sum::<f64>()
        + table.frame.right_from_text;
    assert_close(table.frame.right_from_text, FROM_TEXT_PT, "rightFromText");

    let runs = placed(data);
    let (body, _) = find(&runs, "Public body paragraph");
    assert_close(body.left_pt, wrap_edge, "body paragraph");
    let (greeting, width) = find(&runs, "Public greeting");
    assert_close(
        greeting.left_pt + width / 2.0,
        (wrap_edge + MARGIN_PT + TEXT_WIDTH_PT) / 2.0,
        "centred greeting",
    );
}
