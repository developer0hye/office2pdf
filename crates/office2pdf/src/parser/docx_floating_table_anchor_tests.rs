//! Where a `w:tblpPr` floating table with `w:vertAnchor="text"` is drawn:
//! `w:tblpY` below the top edge of its anchor paragraph, the paragraph after
//! it. That edge is where the flow stands before the paragraph's own
//! `w:before`, so earlier content above the table moves it down the page
//! (issue #2236).

use super::floating_table_wrap_tests::{
    FROM_TEXT_PT, MARGIN_PT, assert_close, find, float_docx, paragraph, placed, run,
};
use super::*;

/// A one-column floating table `width_twips` wide, positioned by
/// `position`, with one row exactly 640 twips tall for each of `rows`, its
/// cell reading that text.
fn floating_rows(width_twips: u32, position: &str, rows: &[String]) -> String {
    let rows: String = rows
        .iter()
        .map(|text| {
            format!(
                r#"<w:tr><w:trPr><w:trHeight w:val="640" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{width_twips}" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr>"#,
                paragraph("", text)
            )
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" {position}/><w:tblW w:w="{width_twips}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="{width_twips}"/></w:tblGrid>{rows}</w:tbl>"#
    )
}

/// A one-cell floating table 5600 twips wide whose single row is exactly
/// 640 twips tall, its cell reading `text`, positioned by `position`. Text
/// fits in the room beside it.
fn positioned_table(position: &str, text: &str) -> String {
    floating_rows(5600, position, &[text.to_string()])
}

/// A floating table as wide as the column, which leaves no room beside it,
/// hung `offset_twips` below its anchor paragraph.
fn full_width_table(offset_twips: u32, rows: &[String]) -> String {
    floating_rows(
        9026,
        &format!(r#"w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{offset_twips}""#),
        rows,
    )
}

/// [`positioned_table`] hung `offset_twips` below its anchor paragraph.
fn text_anchored_table(offset_twips: u32, text: &str) -> String {
    positioned_table(
        &format!(r#"w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{offset_twips}""#),
        text,
    )
}

/// A one-cell flow table as wide as the column whose single row is exactly
/// `height_twips` tall.
fn flow_table(height_twips: u32) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="{height_twips}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph("", "Preceding cell")
    )
}

/// A single-spaced paragraph with `before_twips` and `after_twips` of
/// paragraph spacing.
fn spaced_paragraph(before_twips: u32, after_twips: u32, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="{before_twips}" w:after="{after_twips}" w:line="240" w:lineRule="auto"/></w:pPr>{}</w:p>"#,
        run(&format!("<w:t>{text}</w:t>"))
    )
}

/// Every painted run on the first page, the floating tables' cells included.
#[cfg(not(target_arch = "wasm32"))]
fn page_runs(data: &[u8]) -> Vec<crate::render::pdf::PlacedTextRun> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap()
}

#[cfg(not(target_arch = "wasm32"))]
fn baseline(runs: &[crate::render::pdf::PlacedTextRun], text: &str) -> f64 {
    runs.iter()
        .find(|run| run.text == text)
        .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
        .baseline_pt
}

#[cfg(not(target_arch = "wasm32"))]
fn page_runs_at(data: &[u8], page_index: usize) -> Vec<crate::render::pdf::PlacedTextRun> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    crate::render::pdf::compiled_text_runs(&source.source, page_index).unwrap()
}

/// The baseline of a floating table's cell text when the table hangs at the
/// top of the text area: the first block on the page, with no offset.
#[cfg(not(target_arch = "wasm32"))]
fn page_top_float_baseline() -> f64 {
    let data = float_docx(&[text_anchored_table(0, "Floating"), paragraph("", "Anchor")].concat());
    baseline(&page_runs(&data), "Floating")
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_text_anchored_floating_table_hangs_from_the_end_of_the_table_before_it() {
    let page_top: f64 = page_top_float_baseline();
    for height_twips in [1000_u32, 2000] {
        for offset_twips in [0_u32, 390, 800] {
            let data = float_docx(
                &[
                    flow_table(height_twips),
                    text_anchored_table(offset_twips, "Floating"),
                    spaced_paragraph(400, 400, "Anchor"),
                ]
                .concat(),
            );
            assert_close(
                baseline(&page_runs(&data), "Floating"),
                page_top + f64::from(height_twips + offset_twips) / 20.0,
                &format!("{height_twips}-twip table, {offset_twips}-twip offset"),
            );
        }
    }
}

/// The margin anchor is the control: the table before it does not move it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_margin_anchored_floating_table_stays_put_below_a_table() {
    let page_top: f64 = page_top_float_baseline();
    for height_twips in [1000_u32, 2000] {
        for offset_twips in [390_u32, 800] {
            let data = float_docx(
                &[
                    flow_table(height_twips),
                    positioned_table(
                        &format!(r#"w:vertAnchor="margin" w:tblpY="{offset_twips}""#),
                        "Floating",
                    ),
                    paragraph("", "Anchor"),
                ]
                .concat(),
            );
            assert_close(
                baseline(&page_runs(&data), "Floating"),
                page_top + f64::from(offset_twips) / 20.0,
                &format!("{height_twips}-twip table, {offset_twips}-twip offset"),
            );
        }
    }
}

/// The anchor paragraph's top edge is the end of the gap after the
/// paragraph before the table, wherever the anchor's own `w:before` puts its
/// first line, as Word's export of docxide-pdf's `case32` shows for a 10pt
/// `w:after` and the issue's case for a 20pt `w:before`.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_text_anchored_floating_table_hangs_below_the_gap_after_the_paragraph_before_it() {
    let page_top: f64 = page_top_float_baseline();
    let control = page_runs(&float_docx(
        &[
            spaced_paragraph(0, 0, "Above"),
            spaced_paragraph(0, 0, "Below"),
        ]
        .concat(),
    ));
    let line_pt: f64 = baseline(&control, "Below") - baseline(&control, "Above");
    for after_twips in [0_u32, 200, 400] {
        for before_twips in [0_u32, 400] {
            let data = float_docx(
                &[
                    spaced_paragraph(0, after_twips, "Above"),
                    text_anchored_table(390, "Floating"),
                    spaced_paragraph(before_twips, 0, "Anchor"),
                ]
                .concat(),
            );
            assert_close(
                baseline(&page_runs(&data), "Floating"),
                page_top + line_pt + f64::from(after_twips + 390) / 20.0,
                &format!("{after_twips}-twip after, {before_twips}-twip before"),
            );
        }
    }
}

/// Text-anchored tables in a row all hang from the one anchor paragraph
/// after them.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_anchored_floating_tables_in_a_row_share_their_anchor() {
    let page_top: f64 = page_top_float_baseline();
    for after_twips in [0_u32, 300] {
        let data = float_docx(
            &[
                flow_table(1000),
                spaced_paragraph(0, after_twips, "Between"),
                text_anchored_table(0, "First"),
                text_anchored_table(800, "Second"),
                paragraph("", "Anchor"),
            ]
            .concat(),
        );
        let runs = page_runs(&data);
        let anchor_top: f64 = baseline(&runs, "First") - page_top;
        let context = format!("{after_twips}-twip after");
        assert_close(
            baseline(&runs, "Second"),
            page_top + anchor_top + 40.0,
            &context,
        );
        assert!(
            anchor_top > 50.0 + f64::from(after_twips) / 20.0,
            "{context}: the first table hangs at {anchor_top}pt, above the end of the flow"
        );
    }
}

/// Body text after a text-anchored table that hangs below earlier content
/// keeps off the table where it is drawn.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_after_a_text_anchored_floating_table_down_the_page_lies_beside_it() {
    let wrap_edge: f64 = MARGIN_PT + 5600.0 / 20.0 + FROM_TEXT_PT;
    for (height_twips, offset_twips) in [(1000_u32, 390_u32), (2000, 0)] {
        let band_top: f64 = MARGIN_PT + f64::from(height_twips + offset_twips) / 20.0;
        let band_bottom: f64 = band_top + 640.0 / 20.0;
        let lines: Vec<String> = (1..=12)
            .map(|line| paragraph("", &format!("Line {line}")))
            .collect();
        let data = float_docx(
            &[
                flow_table(height_twips),
                text_anchored_table(offset_twips, "Floating cell"),
                lines.concat(),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let (mut beside, mut below) = (0, 0);
        for (run, _) in runs.iter().filter(|(run, _)| run.text.starts_with("Line")) {
            if run.baseline_pt > band_top + run.font_size_pt && run.baseline_pt < band_bottom {
                assert_close(run.left_pt, wrap_edge, &run.text);
                beside += 1;
            } else if run.baseline_pt - run.font_size_pt > band_bottom {
                assert_close(run.left_pt, MARGIN_PT, &run.text);
                below += 1;
            }
        }
        assert!(
            beside > 0 && below > 0,
            "{height_twips}-twip table, {offset_twips}-twip offset: {beside} lines beside, \
             {below} below in {runs:?}"
        );
    }
}

/// docxide-pdf's `empty-float-float-indent-0-0`, the issue's case: Word
/// draws the table's row 19.5pt below the 100pt table before it, at 191.52pt.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn issue_case_floating_topic_hangs_below_the_preceding_table() {
    let data: &[u8] = include_bytes!(
        "../../../../tests/fixtures/docx/docxide/text_anchored_float_after_table.docx"
    );
    let runs = placed(data);
    let (preceding, _) = find(&runs, "Preceding table");
    let (topic, _) = find(&runs, "Floating topic");
    // Both cells seat their 13pt line the same way, so the baselines are as
    // far apart as the row tops.
    assert_close(
        topic.baseline_pt - preceding.baseline_pt,
        100.0 + 19.5,
        "Floating topic",
    );
}

/// No Word reference aligns a text-anchored table with `w:tblpYSpec`, so
/// such a table stays aligned in the text area, as before, whatever comes
/// before it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_aligned_text_anchored_floating_table_stays_aligned_in_the_text_area() {
    let aligned = |body_before: &str| {
        let data = float_docx(
            &[
                body_before,
                &positioned_table(
                    r#"w:vertAnchor="text" w:horzAnchor="margin" w:tblpYSpec="bottom""#,
                    "Floating",
                ),
                &paragraph("", "Anchor"),
            ]
            .concat(),
        );
        baseline(&page_runs(&data), "Floating")
    };
    let at_page_top: f64 = aligned("");
    assert!(
        at_page_top > MARGIN_PT + 600.0,
        "a bottom-aligned table sits at the text area's foot, not at {at_page_top}pt"
    );
    for height_twips in [1000_u32, 2000] {
        assert_close(
            aligned(&flow_table(height_twips)),
            at_page_top,
            &format!("{height_twips}-twip table before it"),
        );
    }
}

/// A text-anchored table that leaves no room beside it hangs from its anchor
/// like any other, and the anchor paragraph starts its own `w:before` below
/// the table's bottom edge, as every line Word meets the table with does.
/// Each offset here is less than the anchor's first line, which therefore
/// meets the table.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_full_width_text_anchored_table_hangs_from_its_anchor_with_the_text_below_it() {
    let page_top: f64 = page_top_float_baseline();
    let control = page_runs(&float_docx(
        &[
            spaced_paragraph(0, 0, "Above"),
            spaced_paragraph(0, 0, "Below"),
        ]
        .concat(),
    ));
    let line_pt: f64 = baseline(&control, "Below") - baseline(&control, "Above");
    let rows: Vec<String> = vec!["Floating".to_string(), "Second row".to_string()];
    for (after_twips, offset_twips) in [(0_u32, 0_u32), (200, 100), (400, 200)] {
        for before_twips in [0_u32, 300] {
            let data = float_docx(
                &[
                    spaced_paragraph(0, after_twips, "Above"),
                    full_width_table(offset_twips, &rows),
                    spaced_paragraph(before_twips, 0, "Anchor"),
                ]
                .concat(),
            );
            let runs = page_runs(&data);
            let context = format!(
                "{after_twips}-twip after, {offset_twips}-twip offset, {before_twips}-twip before"
            );
            let table_top: f64 = line_pt + f64::from(after_twips + offset_twips) / 20.0;
            assert_close(baseline(&runs, "Floating"), page_top + table_top, &context);
            // The anchor's first line sits as far below the table's two 32pt
            // rows and its `w:before` as "Above" sits below the page top.
            assert_close(
                baseline(&runs, "Anchor"),
                baseline(&runs, "Above") + table_top + 64.0 + f64::from(before_twips) / 20.0,
                &context,
            );
        }
    }
}

/// Word continues a text-anchored table that leaves no room beside it on the
/// next page when it runs past the text area, and the text after it follows
/// its last row there: docxide-pdf's `scraped/bulgarian_road_safety_program`
/// spreads its 24-row table over five pages.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_full_width_text_anchored_table_past_the_page_end_continues_on_the_next_page() {
    // A4's 841.9pt page less two 72pt margins.
    let text_bottom: f64 = 841.9 - MARGIN_PT;
    for lines_before in [20_usize, 40] {
        let before: Vec<String> = (1..=lines_before)
            .map(|line| paragraph("", &format!("Before {line}")))
            .collect();
        let rows: Vec<String> = (1..=20).map(|row| format!("Row {row}")).collect();
        let data = float_docx(
            &[
                before.concat(),
                full_width_table(0, &rows),
                paragraph("", "After table"),
            ]
            .concat(),
        );
        let first = page_runs_at(&data, 0);
        let second = page_runs_at(&data, 1);
        let context = format!("{lines_before} lines before");
        let rows_on = |runs: &[crate::render::pdf::PlacedTextRun]| -> Vec<f64> {
            runs.iter()
                .filter(|run| run.text.starts_with("Row "))
                .map(|run| run.baseline_pt)
                .collect()
        };
        let (first_rows, second_rows) = (rows_on(&first), rows_on(&second));
        assert_eq!(
            first_rows.len() + second_rows.len(),
            rows.len(),
            "{context}: every row is on the two pages"
        );
        assert!(
            !first_rows.is_empty() && !second_rows.is_empty(),
            "{context}: the table starts on page 1 and continues on page 2"
        );
        assert!(
            first_rows.iter().all(|baseline| *baseline < text_bottom),
            "{context}: no row runs past the text area: {first_rows:?}"
        );
        let last_row: f64 = second_rows.iter().copied().fold(f64::MIN, f64::max);
        assert!(
            baseline(&second, "After table") > last_row,
            "{context}: the text after the table follows its last row"
        );
    }
}
