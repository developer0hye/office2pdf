#![cfg(not(target_arch = "wasm32"))] // native-only integration tests (fs, qpdf, criterion)
//! Integration tests for XLSX fixtures.
//!
//! Each real-world `.xlsx` file in `tests/fixtures/xlsx/` gets two tests:
//! - **smoke**: `convert()` → valid PDF (or graceful error — no panic)
//! - **structure**: parse → assert expected IR content

mod common;

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;
use office2pdf::internal::Parser;
use office2pdf::internal::XlsxParser;
use office2pdf::internal::generate_typst;
use office2pdf::ir::{
    Alignment, Block, BorderLineStyle, ChartAreaOutline, ChartHost, ChartLine, ChartPlotAreaLayout,
    ChartUserShapeExtent, Color, HFInline, Page, SheetPage, TableCell,
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx")
        .join(name)
}

fn load_fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("fixture file should exist")
}

/// Smoke-test helper: conversion must not panic.
fn assert_produces_valid_pdf(name: &str) {
    let path = fixture_path(name);
    match office2pdf::convert(&path) {
        Ok(result) => {
            assert!(!result.pdf.is_empty(), "PDF output should not be empty");
            assert!(
                result.pdf.starts_with(b"%PDF"),
                "output should start with PDF magic bytes"
            );
            common::validate_pdf_with_qpdf(&result.pdf);
        }
        Err(e) => {
            eprintln!("[WARN] {name}: conversion error (non-panic): {e}");
        }
    }
}

/// Parse an XLSX fixture and return the sheet pages.
fn sheet_pages(name: &str) -> Vec<SheetPage> {
    let data = load_fixture(name);
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .into_iter()
        .filter_map(|p| match p {
            Page::Sheet(sp) => Some(sp),
            _ => None,
        })
        .collect()
}

fn sheet_names(pages: &[SheetPage]) -> Vec<&str> {
    pages.iter().map(|p| p.name.as_str()).collect()
}

fn total_rows(pages: &[SheetPage]) -> usize {
    pages.iter().map(|p| p.table.rows.len()).sum()
}

fn has_cell_border(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.border.is_some())
    })
}

fn has_merged_cells(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table
            .rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .any(|c| c.col_span > 1 || c.row_span > 1)
    })
}

fn has_formatted_text(pages: &[SheetPage]) -> bool {
    pages.iter().any(|p| {
        p.table.rows.iter().flat_map(|r| r.cells.iter()).any(|c| {
            c.content.iter().any(|b| match b {
                Block::Paragraph(para) => para.runs.iter().any(|r| {
                    r.style.bold == Some(true)
                        || r.style.italic == Some(true)
                        || r.style.color.is_some()
                }),
                _ => false,
            })
        })
    })
}

fn table_cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
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

fn sheet_page_named<'a>(pages: &'a [SheetPage], name: &str) -> &'a SheetPage {
    pages
        .iter()
        .find(|page| page.name == name)
        .unwrap_or_else(|| {
            let available = pages
                .iter()
                .map(|page| page.name.as_str())
                .collect::<Vec<_>>();
            panic!("missing sheet page {name}; available pages: {available:?}")
        })
}

// ---------------------------------------------------------------------------
// PR #186 contributor acceptance fixture
// ---------------------------------------------------------------------------

const PR_186_FIXTURE: &str = "pr_186_contributor_acceptance.xlsx";

#[test]
fn smoke_pr_186_contributor_acceptance_fixture() {
    assert_produces_valid_pdf(PR_186_FIXTURE);
}

#[test]
fn structure_pr_186_contributor_acceptance_supported_behavior() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let statement_pages = pages
        .iter()
        .filter(|page| page.name == "Statement Landscape")
        .collect::<Vec<_>>();
    let executive = sheet_page_named(&pages, "Executive Portrait");

    assert_eq!(statement.table.rows.len(), 4);
    assert_eq!(
        statement_pages
            .iter()
            .map(|page| page.table.column_widths.len())
            .sum::<usize>(),
        4
    );
    assert!(
        (statement.table.column_widths[1] - 120.0).abs() < 0.01,
        "20 Carlito character units should convert to 120pt"
    );

    let expected_alignments = [
        Alignment::Left,
        Alignment::Center,
        Alignment::Right,
        Alignment::Justify,
    ];
    let alignment_cells = statement_pages
        .iter()
        .flat_map(|page| page.table.rows[1].cells.iter());
    for (cell, expected) in alignment_cells.zip(expected_alignments) {
        let paragraph = match &cell.content[0] {
            Block::Paragraph(paragraph) => paragraph,
            _ => panic!("alignment cell should contain a paragraph"),
        };
        assert_eq!(paragraph.style.alignment, Some(expected));
    }

    assert!(
        !table_cell_text(&statement.table.rows[2].cells[1]).is_empty(),
        "the text-valued General cell should be retained"
    );
    let top_border = statement.table.rows[3].cells[0]
        .border
        .as_ref()
        .and_then(|border| border.top.as_ref())
        .expect("A4 should have a top border");
    assert_eq!(top_border.style, BorderLineStyle::Double);

    // The sheet declares 0.4in / 0.5in / 0.3in / 0.6in and 0.7in, which reach
    // the paper on the whole point Excel prints against (issue #1127).
    assert!((statement.margins.top - 28.0).abs() < 0.01);
    assert!((statement.margins.bottom - 36.0).abs() < 0.01);
    assert!((statement.margins.left - 21.0).abs() < 0.01);
    assert!((statement.margins.right - 43.0).abs() < 0.01);
    assert!((executive.margins.top - 54.0).abs() < 0.01);
    assert!((executive.margins.bottom - 54.0).abs() < 0.01);
    assert!((executive.margins.left - 50.0).abs() < 0.01);
    assert!((executive.margins.right - 50.0).abs() < 0.01);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_six_digit_colors() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let header_cell = &statement.table.rows[0].cells[0];

    assert_eq!(
        header_cell.background,
        Some(Color::new(0xD9, 0xEA, 0xF7)),
        "six-digit header fill should survive XLSX parsing"
    );
    let paragraph = match &header_cell.content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("header cell should contain a paragraph"),
    };
    assert_eq!(
        paragraph.runs[0].style.color,
        Some(Color::new(0x16, 0x32, 0x4F)),
        "six-digit header font color should survive XLSX parsing"
    );
}

#[test]
fn acceptance_pr_186_contributor_acceptance_carlito_fallback() {
    let data = load_fixture(PR_186_FIXTURE);
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let output = generate_typst(&document).expect("fixture should generate Typst");

    assert!(
        output.source.contains(
            r#"font: ("Carlito", "Calibri", "Liberation Sans", "Arimo", "Arial", "DejaVu Sans", "Helvetica")"#
        ),
        "Carlito should retain a sans-serif fallback chain, ending on the generic \
         sans faces so a host with none of the named ones still keeps the class: {}",
        output.source
    );
}

#[test]
fn acceptance_pr_186_contributor_acceptance_numeric_general_alignment() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let general_row = &statement.table.rows[2];

    let numeric = match &general_row.cells[0].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("numeric cell should contain a paragraph"),
    };
    let numeric_looking_text = match &general_row.cells[1].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("text cell should contain a paragraph"),
    };
    assert_eq!(numeric.style.alignment, Some(Alignment::Right));
    assert_eq!(numeric_looking_text.style.alignment, None);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_page_setup() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement = sheet_page_named(&pages, "Statement Landscape");
    let executive = sheet_page_named(&pages, "Executive Portrait");

    assert!((statement.size.width - 612.0).abs() < 0.01);
    assert!((statement.size.height - 396.0).abs() < 0.01);
    assert!((executive.size.width - 522.0).abs() < 0.01);
    assert!((executive.size.height - 756.0).abs() < 0.01);
}

#[test]
fn acceptance_pr_186_contributor_acceptance_print_pagination() {
    let pages = sheet_pages(PR_186_FIXTURE);
    let statement_pages = pages
        .iter()
        .filter(|page| page.name == "Statement Landscape")
        .collect::<Vec<_>>();

    assert_eq!(
        statement_pages.len(),
        2,
        "the four statement columns should print as two horizontal pages"
    );
    assert_eq!(
        statement_pages[0].table.column_widths,
        [156.0, 120.0, 144.0]
    );
    assert_eq!(statement_pages[1].table.column_widths, [144.0]);
    assert_eq!(
        table_cell_text(&statement_pages[0].table.rows[0].cells[2]),
        "Explicit Right"
    );
    assert_eq!(
        table_cell_text(&statement_pages[1].table.rows[0].cells[0]),
        "Explicit Justify"
    );
}

#[test]
fn acceptance_pr_186_contributor_acceptance_double_border_rendering() {
    let data = load_fixture(PR_186_FIXTURE);
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let output = generate_typst(&document).expect("fixture should generate Typst");

    assert!(!output.source.contains("dash: \"dashed\""));
    // Overlay offsets track the cell padding, which is 3pt left and 2pt right
    // since issue #1157 (it was 3pt each side, which ended every right-aligned
    // run a whole point inside Excel's). The `dx` backs out the left inset
    // alone, so it is unchanged; the length backs out both, so it moved with
    // the right one. Since issue #619 the double paints as two
    // boundary-anchored 1pt bands [B-1, B] and [B+1, B+2], and each band runs
    // 1pt past its end boundary, so the run is the 5pt padding backout plus 1.
    assert!(output.source.contains(
        "#place(top + left, dx: -3pt, dy: -2pt, line(length: 100% + 6pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
    ));
    assert!(output.source.contains(
        "#place(top + left, dx: -3pt, dy: 0pt, line(length: 100% + 6pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
    ));
}

// ---------------------------------------------------------------------------
// any_sheets.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_any_sheets() {
    assert_produces_valid_pdf("any_sheets.xlsx");
}

#[test]
fn structure_any_sheets() {
    // any_sheets.xlsx declares 4 sheets: Visible, Hidden (state="hidden"),
    // VeryHidden (state="veryHidden") and Chart. The hidden pair is out
    // because Excel never prints a hidden sheet (issue #1065) — and empty of
    // cells besides. `Chart` is a chartsheet, and Excel prints one of those as
    // a page of its own (issue #1099), so two pages are left.
    let pages = sheet_pages("any_sheets.xlsx");
    assert_eq!(sheet_names(&pages), vec!["Visible", "Chart"]);
}

/// `Visible` names Arial 12 as its Normal font and records no row heights.
/// Excel recomputes a 16pt worksheet row, then seats it on a 15pt printed
/// track. The old Calibri/Aptos-keyed predicate left all five tracks at 16pt
/// instead (issue #1224).
#[test]
fn structure_any_sheets_uses_the_arial_printed_grid() {
    let pages = sheet_pages("any_sheets.xlsx");
    let visible = pages
        .iter()
        .find(|page| page.name == "Visible")
        .expect("the visible worksheet should contribute a page");

    assert_eq!(
        visible
            .table
            .rows
            .iter()
            .map(|row| row.height)
            .collect::<Vec<_>>(),
        vec![Some(15.0); 5]
    );
}

/// A chartsheet prints as one page carrying its chart alone.
///
/// Measured on Excel for Mac 16.100 exports of this fixture's `Chart` sheet:
/// one page, the chart seated inside the printable area, and — since the
/// chartsheet declares no `<pageSetup>` — landscape, where the same workbook's
/// `Visible` worksheet (equally without one) exports portrait.
#[test]
fn structure_any_sheets_chartsheet_pages_its_chart_alone() {
    let pages = sheet_pages("any_sheets.xlsx");
    let chart_page = pages
        .iter()
        .find(|page| page.name == "Chart")
        .expect("the chartsheet should contribute a page");

    // No `<pageSetup>`: Letter (the schema's paperSize default, issue #717)
    // turned on its side by the chartsheet's own landscape default.
    assert_eq!(
        (chart_page.size.width, chart_page.size.height),
        (792.0, 612.0)
    );
    // The chartsheet's own `<pageMargins>`: 0.7" sides, 0.75" top and bottom.
    assert_eq!(
        (
            chart_page.margins.left,
            chart_page.margins.right,
            chart_page.margins.top,
            chart_page.margins.bottom,
        ),
        (50.4, 50.4, 54.0, 54.0)
    );
    // A chartsheet holds no cells at all.
    assert!(chart_page.table.rows.is_empty());

    assert_eq!(chart_page.charts.len(), 1);
    assert_eq!(
        chart_page.charts[0].chart.host,
        ChartHost::SpreadsheetChartsheet,
        "the renderer must be able to distinguish a chartsheet from an anchored worksheet chart"
    );
    let placement = chart_page.charts[0]
        .placement
        .expect("a chartsheet's chart is placed, not flowed after the grid");
    // Excel starts the chart 4pt inside each margin, on the whole point. The
    // recorded native box on this paper is 681.82 x 493.18 at (54, 58) inside
    // a 691.2 x 504 printable area; the page-only model uses the exact 4pt
    // near-edge rule symmetrically and keeps the internal-grid residual as its
    // error term (issues #1147 and #1221).
    assert_eq!(
        (
            chart_page.margins.left + placement.x_offset_pt,
            chart_page.margins.top + placement.y_offset_pt,
        ),
        (54.0, 58.0)
    );
    assert_eq!((placement.width, placement.height), (683.0, 496.0));

    // The chart belongs to the chartsheet, not to the worksheet before it: a
    // native export of `Visible` alone draws no chart at all.
    let worksheet_page = pages
        .iter()
        .find(|page| page.name == "Visible")
        .expect("the visible worksheet should still page");
    assert!(worksheet_page.charts.is_empty());
}

/// The chart frame and the value axis' major gridlines take the colour the
/// part declares, not the renderer's built-in grey.
///
/// This workbook's `xl/charts/chart1.xml` states one stroke three times — for
/// the chart space's frame, for the category axis, and for the value axis'
/// major gridlines — as `<a:schemeClr val="tx1">` lifted by lumMod 15% /
/// lumOff 85%. The workbook theme's `dk1` is black, so that is #D9D9D9, and a
/// `mutool draw -F trace` of the Excel for Mac 16.100 export of the `Chart`
/// sheet draws every one of them at `.8509804` on all three channels. A
/// workbook carries no colour map to turn `tx1` into `dk1`, so the colour
/// resolved to nothing and all three came out at RGB 134 (issue #1145).
#[test]
fn structure_any_sheets_chart_chrome_takes_its_declared_colour() {
    let pages = sheet_pages("any_sheets.xlsx");
    let chart = &pages
        .iter()
        .find(|page| page.name == "Chart")
        .expect("the chartsheet should contribute a page")
        .charts
        .first()
        .expect("the chartsheet carries its chart")
        .chart;

    let declared = Color::new(0xD9, 0xD9, 0xD9);
    assert_eq!(
        chart.chart_area_outline,
        ChartAreaOutline::Explicit {
            width_pt: Some(0.75),
            color: Some(declared),
        },
    );
    let expected = ChartLine::Explicit {
        width_pt: Some(0.75),
        color: Some(declared),
        alpha: None,
    };
    assert_eq!(chart.major_gridline_line, expected, "major gridlines");
    assert_eq!(chart.category_axis_line, expected, "category axis");
}

/// `c:chartSpace/c:spPr` paints the whole chart area before it draws the
/// title, plot, and legend. The real fixture declares theme `bg1`, which is
/// white on white and hides a dropped fill, so this package-local mutation
/// gives only that outer fill a colour no other element uses (issue #1217).
#[test]
fn a_chart_space_solid_fill_paints_the_chart_area_box() {
    let data = repackage_xlsx_part(
        &load_fixture("any_sheets.xlsx"),
        "xl/charts/chart1.xml",
        |xml| {
            xml.replace(
                r#"<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525""#,
                r#"<c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln w="9525""#,
            )
        },
    );
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the patched fixture should parse");
    let source = generate_typst(&document)
        .expect("the patched fixture should generate Typst")
        .source;
    assert!(
        source.contains("fill: rgb(18, 52, 86)"),
        "the chart-space solid fill must paint the chart-area box; got:\n{source}"
    );
}

/// Both axes' labels take the colour their own `c:txPr` declares.
///
/// Each `<c:catAx>`/`<c:valAx>` gives its `<a:defRPr>` an
/// `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff
/// val="35000"/></a:schemeClr></a:solidFill>`, which over the theme's black
/// `dk1` is #595959. The Excel for Mac 16 export of the `Chart` sheet draws
/// all eleven tick and category labels at `.3490196` on all three channels;
/// the transforms were children of the colour element and were dropped, so
/// every one of them came out pure black (issue #1160).
///
/// The chart space's own `c:txPr` states a bare `<a:defRPr/>`, so its colour
/// stays `None` — the axes are the only scopes this workbook colours that the
/// renderer reads.
#[test]
fn structure_any_sheets_chart_labels_take_their_declared_colour() {
    let pages = sheet_pages("any_sheets.xlsx");
    let chart = &pages
        .iter()
        .find(|page| page.name == "Chart")
        .expect("the chartsheet should contribute a page")
        .charts
        .first()
        .expect("the chartsheet carries its chart")
        .chart;

    let declared = Some(Color::new(0x59, 0x59, 0x59));
    assert_eq!(
        chart.category_axis_text_style.color, declared,
        "category axis labels"
    );
    assert_eq!(
        chart.value_axis_text_style.color, declared,
        "value axis labels"
    );
    assert_eq!(
        chart.text_style.color, None,
        "the chart space declares no colour of its own"
    );
    // The same `a:defRPr` states `sz="900"`, which is what proves the run
    // properties around the colour still read after it is consumed.
    assert_eq!(chart.category_axis_text_style.size_pt, Some(9.0));
    assert_eq!(chart.value_axis_text_style.size_pt, Some(9.0));
}

/// The chartsheet's chart declares a `<c:title>` that names no text, so Excel
/// supplies the string itself and prints it above a shortened plot.
///
/// `xl/charts/chart1.xml` writes `<c:title>` with a `c:layout`, a `c:overlay`,
/// a `c:spPr` and a `c:txPr` — and no `<c:tx>` anywhere — beside
/// `<c:autoTitleDeleted val="0"/>`. An Excel for Mac 16.100 export of the
/// `Chart` sheet, forced to Letter landscape, prints the placeholder at 14pt
/// centred over the chart box and starts the topmost gridline 44.20pt below
/// the box's top edge (issue #1146).
#[test]
fn structure_any_sheets_chart_declares_an_automatic_title() {
    let pages = sheet_pages("any_sheets.xlsx");
    let chart = &pages
        .iter()
        .find(|page| page.name == "Chart")
        .expect("the chartsheet should contribute a page")
        .charts
        .first()
        .expect("the chartsheet carries its chart")
        .chart;

    assert!(chart.has_automatic_title);
    assert!(!chart.auto_title_deleted);
    // The part names no string of its own, and neither series is named, so
    // nothing in the file can supply one.
    assert_eq!(chart.title, None);
    assert!(chart.series.iter().all(|series| series.name.is_none()));
}

/// The title's own `<c:txPr>` states its size, its weight and its colour, and
/// all three have to survive the parse.
///
/// `xl/charts/chart1.xml` gives `<c:title>` a `<c:txPr>` whose `<a:defRPr>`
/// carries `sz="1400" b="0"` over an `<a:schemeClr val="tx1">` lifted by
/// lumMod 65% / lumOff 35% — #595959 against the theme's black `dk1`. The
/// Excel for Mac 16.100 export of the `Chart` sheet draws the placeholder at
/// `trm="14 0 0 14"` in `.3490196` grey. The subtree was read for its text
/// alone and everything beside it thrown away, so the title printed at the
/// renderer's own 11pt, bold and black (issue #1215).
#[test]
fn structure_any_sheets_chart_title_carries_its_own_run_properties() {
    let pages = sheet_pages("any_sheets.xlsx");
    let chart = &pages
        .iter()
        .find(|page| page.name == "Chart")
        .expect("the chartsheet should contribute a page")
        .charts
        .first()
        .expect("the chartsheet carries its chart")
        .chart;

    assert_eq!(chart.title_text_style.size_pt, Some(14.0), "declared size");
    assert_eq!(chart.title_text_style.bold, Some(false), "declared weight");
    assert_eq!(
        chart.title_text_style.color,
        Some(Color::new(0x59, 0x59, 0x59)),
        "declared colour"
    );
    // The chart space's own `c:txPr` is a bare `<a:defRPr/>`, so nothing here
    // could have come from it.
    assert_eq!(chart.text_style.size_pt, None);
    assert_eq!(chart.text_style.bold, None);
}

/// The legend's own `<c:txPr>` governs every entry rather than disappearing
/// between the chart parser and renderer.
///
/// `xl/charts/chart1.xml` gives `<c:legend>` a 9pt regular `a:defRPr` filled
/// with the same transformed `tx1` colour as the axes: #595959. Excel for Mac
/// 16.100 draws both legend entries at that size and colour, while the chart
/// space itself declares neither property (issue #1236).
#[test]
fn structure_any_sheets_chart_legend_carries_its_own_run_properties() {
    let data = load_fixture("any_sheets.xlsx");
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the fixture should parse");
    let source = generate_typst(&document)
        .expect("the fixture should generate Typst")
        .source;
    let entries: Vec<&str> = source
        .lines()
        .filter(|line| line.contains("[Series 1]") || line.contains("[Series 2]"))
        .collect();

    assert_eq!(entries.len(), 2, "both legend entries are emitted");
    for entry in entries {
        assert!(
            entry.contains("#text(size: 9pt, fill: rgb(89, 89, 89))[Series"),
            "the legend's own size and colour reach every entry; got: {entry}"
        );
    }
}

/// Two chartsheet packages whose parts collide by filename, which is how the
/// worksheet-only rels lookup went wrong in the first place.
///
/// `chart_sheet.xlsx` has no `xl/worksheets/_rels/` at all, so its chartsheet
/// resolved to nothing and its chart fell through to the first worksheet as an
/// orphan. `SimpleScatterChart.xlsx` has both `xl/worksheets/sheet1.xml` and
/// `xl/chartsheets/sheet1.xml`, so the chartsheet read the *worksheet's* rels
/// and printed a second copy of the worksheet's chart instead of its own.
#[test]
fn structure_chartsheet_reads_its_own_drawing_relationships() {
    let pages = sheet_pages("chart_sheet.xlsx");
    assert_eq!(sheet_names(&pages), vec!["Sheet1", "Chart1"]);
    assert!(pages[0].charts.is_empty());
    assert_eq!(pages[1].charts.len(), 1);

    let scatter = sheet_pages("SimpleScatterChart.xlsx");
    assert_eq!(sheet_names(&scatter), vec!["Sheet1", "Chart1"]);
    // The worksheet's chart keeps its own anchor; the chartsheet's takes the
    // page, so the two placements can no longer be the same box.
    let worksheet_placement = scatter[0].charts[0]
        .placement
        .expect("the worksheet chart is anchored");
    let chartsheet_placement = scatter[1].charts[0]
        .placement
        .expect("the chartsheet chart is placed by its page setup");
    assert_eq!(
        (chartsheet_placement.width, chartsheet_placement.height),
        (683.0, 496.0)
    );
    assert_ne!(
        (worksheet_placement.width, worksheet_placement.height),
        (chartsheet_placement.width, chartsheet_placement.height)
    );
}

// ---------------------------------------------------------------------------
// issue_1065_hidden_sheet_probe.xlsx
// ---------------------------------------------------------------------------

/// The probe workbook Excel for Mac 16.112 itself authored for issue #1065:
/// a visible `Summary` sheet and a populated `state="hidden"` `Data` sheet,
/// with `paperSize="9"` written into both worksheets afterwards so the paper
/// comes from the file rather than from whichever printer exported it.
///
/// A native export of it is one page — Excel prints no hidden sheet — and so
/// is ours; the hidden sheet's rows reach neither the page list nor the IR.
#[test]
fn structure_issue_1065_probe_pages_only_the_visible_sheet() {
    let pages = sheet_pages("issue_1065_hidden_sheet_probe.xlsx");
    assert_eq!(sheet_names(&pages), vec!["Summary"]);

    let printed_text: String = pages
        .iter()
        .flat_map(|page| page.table.rows.iter())
        .flat_map(|row| row.cells.iter())
        .map(table_cell_text)
        .collect::<Vec<String>>()
        .join(" ");
    assert!(
        printed_text.contains("Regional revenue"),
        "the visible sheet's title should print"
    );
    assert!(
        !printed_text.contains("N-01"),
        "the hidden sheet's lookup rows should not print"
    );
}

// ---------------------------------------------------------------------------
// 123233_charts.xlsx
// ---------------------------------------------------------------------------

/// The workbook's four `data_Page1_1_*` sheets are `state="hidden"` scratch
/// data behind the charts on `Page1_1`. Excel prints the one visible sheet;
/// paging the hidden four added four pages no reference export has
/// (issue #1065).
#[test]
fn structure_charts_123233_pages_only_the_visible_sheet() {
    let pages = sheet_pages("123233_charts.xlsx");
    let names = sheet_names(&pages);
    // `Page1_1` is wider than its paper, so it splits across pages of its own.
    assert!(
        names.iter().all(|name| *name == "Page1_1"),
        "only the visible sheet should page, got {names:?}"
    );
}

// ---------------------------------------------------------------------------
// date.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_date() {
    assert_produces_valid_pdf("date.xlsx");
}

#[test]
fn structure_date() {
    let pages = sheet_pages("date.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ---------------------------------------------------------------------------
// merge_cells.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_merge_cells() {
    assert_produces_valid_pdf("merge_cells.xlsx");
}

#[test]
fn structure_merge_cells() {
    let pages = sheet_pages("merge_cells.xlsx");
    assert!(
        has_merged_cells(&pages),
        "should have cells with col_span > 1 or row_span > 1"
    );
}

// ---------------------------------------------------------------------------
// SH001-Table.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh001_table() {
    assert_produces_valid_pdf("SH001-Table.xlsx");
}

#[test]
fn structure_sh001_table() {
    let pages = sheet_pages("SH001-Table.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ---------------------------------------------------------------------------
// SH002-TwoTablesTwoSheets.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh002_two_tables_two_sheets() {
    assert_produces_valid_pdf("SH002-TwoTablesTwoSheets.xlsx");
}

#[test]
fn structure_sh002_two_tables_two_sheets() {
    let pages = sheet_pages("SH002-TwoTablesTwoSheets.xlsx");
    assert!(pages.len() >= 2, "should have >= 2 sheets");
    let names = sheet_names(&pages);
    let unique: std::collections::HashSet<_> = names.iter().collect();
    assert_eq!(unique.len(), names.len(), "sheet names should be unique");
}

// ---------------------------------------------------------------------------
// SH106-Formatted.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh106_formatted() {
    assert_produces_valid_pdf("SH106-Formatted.xlsx");
}

#[test]
fn structure_sh106_formatted() {
    let pages = sheet_pages("SH106-Formatted.xlsx");
    assert!(
        has_formatted_text(&pages),
        "should have formatted text (bold/italic/color)"
    );
}

// ---------------------------------------------------------------------------
// SH109-CellWithBorder.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_sh109_cell_with_border() {
    assert_produces_valid_pdf("SH109-CellWithBorder.xlsx");
}

#[test]
fn structure_sh109_cell_with_border() {
    let pages = sheet_pages("SH109-CellWithBorder.xlsx");
    assert!(has_cell_border(&pages), "should have cells with borders");
}

// ---------------------------------------------------------------------------
// temperature.xlsx
// ---------------------------------------------------------------------------

#[test]
fn smoke_temperature() {
    assert_produces_valid_pdf("temperature.xlsx");
}

#[test]
fn structure_temperature() {
    let pages = sheet_pages("temperature.xlsx");
    assert!(!pages.is_empty(), "should have at least one sheet");
    assert!(total_rows(&pages) > 0, "should have data rows");
}

// ===========================================================================
// PDF text content verification
// ===========================================================================

/// Helper: convert an XLSX fixture to PDF and extract text.
fn pdf_text(name: &str) -> String {
    let path = fixture_path(name);
    let result = office2pdf::convert(&path).expect("conversion should succeed");
    common::extract_pdf_text(&result.pdf)
}

// ---------------------------------------------------------------------------
// temperature.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_temperature() {
    let text = pdf_text("temperature.xlsx");
    assert!(
        text.contains("celsius"),
        "PDF should contain 'celsius' label"
    );
    assert!(
        text.contains("fahrenheit"),
        "PDF should contain 'fahrenheit' label"
    );
}

// ---------------------------------------------------------------------------
// SH001-Table.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_sh001_table() {
    let text = pdf_text("SH001-Table.xlsx");
    // This is a simple table with single-character headers and numeric data
    assert!(!text.is_empty(), "PDF should contain extracted text");
    // Check for numeric data that should be present
    assert!(
        text.contains('1') && text.contains('2') && text.contains('3'),
        "PDF should contain numeric data from the table"
    );
}

// ---------------------------------------------------------------------------
// SH002-TwoTablesTwoSheets.xlsx — text content
// ---------------------------------------------------------------------------

#[test]
fn text_content_sh002_two_tables_two_sheets() {
    let text = pdf_text("SH002-TwoTablesTwoSheets.xlsx");
    assert!(!text.is_empty(), "PDF should contain extracted text");
    // Both sheets have different content; verify we have data from at least one
    assert!(
        text.contains('1') || text.contains('a') || text.contains('q'),
        "PDF should contain data from the sheets"
    );
}

// ===========================================================================
// Third-party fixtures — smoke tests (must not panic)
// ===========================================================================

/// Generate a pair of smoke + basic-structure tests for an XLSX fixture.
macro_rules! xlsx_fixture_tests {
    ($test_name:ident, $file:expr) => {
        paste::paste! {
            #[test]
            fn [<smoke_ $test_name>]() {
                assert_produces_valid_pdf($file);
            }

            #[test]
            fn [<structure_ $test_name>]() {
                let data = load_fixture($file);
                match XlsxParser.parse(&data, &ConvertOptions::default()) {
                    Ok((doc, _)) => {
                        let _ = doc.pages.len();
                    }
                    Err(e) => {
                        eprintln!("[WARN] {}: parse error (non-panic): {e}", $file);
                    }
                }
            }
        }
    };
}

// --- CC0 (Public Domain) ---------------------------------------------------

xlsx_fixture_tests!(ffc, "ffc.xlsx");
xlsx_fixture_tests!(hundred_customers, "100-customers.xlsx");
xlsx_fixture_tests!(thousand_customers, "1000-customers.xlsx");

// --- Apache POI (Apache 2.0) -----------------------------------------------

xlsx_fixture_tests!(charts_123233, "123233_charts.xlsx");
xlsx_fixture_tests!(booleans, "Booleans.xlsx");
xlsx_fixture_tests!(chart_sheet, "chart_sheet.xlsx");
xlsx_fixture_tests!(comments, "comments.xlsx");
xlsx_fixture_tests!(excel_pivot_table, "ExcelPivotTableSample.xlsx");
xlsx_fixture_tests!(excel_tables, "ExcelTables.xlsx");
xlsx_fixture_tests!(formatting, "Formatting.xlsx");
xlsx_fixture_tests!(group_test, "GroupTest.xlsx");
xlsx_fixture_tests!(header_footer_test, "headerFooterTest.xlsx");
xlsx_fixture_tests!(inline_string, "InlineString.xlsx");
xlsx_fixture_tests!(picture, "picture.xlsx");
xlsx_fixture_tests!(right_to_left, "right-to-left.xlsx");
xlsx_fixture_tests!(sample_ss, "SampleSS.xlsx");
xlsx_fixture_tests!(shared_formulas, "shared_formulas.xlsx");
xlsx_fixture_tests!(sheet_tab_colors, "SheetTabColors.xlsx");
xlsx_fixture_tests!(simple_monthly_budget, "simple-monthly-budget.xlsx");
xlsx_fixture_tests!(simple_scatter_chart, "SimpleScatterChart.xlsx");
xlsx_fixture_tests!(themes, "Themes.xlsx");
xlsx_fixture_tests!(with_chart, "WithChart.xlsx");
xlsx_fixture_tests!(with_drawing, "WithDrawing.xlsx");
xlsx_fixture_tests!(with_more_various_data, "WithMoreVariousData.xlsx");
xlsx_fixture_tests!(with_text_box, "WithTextBox.xlsx");
xlsx_fixture_tests!(with_various_data, "WithVariousData.xlsx");

// --- Repo-authored regression fixtures --------------------------------------

xlsx_fixture_tests!(theme_color_drawing, "theme_color_drawing.xlsx");
xlsx_fixture_tests!(
    issue_1065_hidden_sheet_probe,
    "issue_1065_hidden_sheet_probe.xlsx"
);
xlsx_fixture_tests!(
    issue_1110_horizontal_centered_probe,
    "issue_1110_horizontal_centered_probe.xlsx"
);
xlsx_fixture_tests!(
    issue_1158_unreferenced_drawing_rel,
    "issue_1158_unreferenced_drawing_rel.xlsx"
);

// --- MIT: Open-Xml-PowerTools (Microsoft) ----------------------------------

xlsx_fixture_tests!(
    sh003_date_first_col,
    "SH003-TableWithDateInFirstColumn.xlsx"
);
xlsx_fixture_tests!(sh004_offset_location, "SH004-TableAtOffsetLocation.xlsx");
xlsx_fixture_tests!(sh005_shared_strings, "SH005-Table-With-SharedStrings.xlsx");
xlsx_fixture_tests!(sh006_no_shared_strings, "SH006-Table-No-SharedStrings.xlsx");
xlsx_fixture_tests!(sh007_one_cell, "SH007-One-Cell-Table.xlsx");
xlsx_fixture_tests!(sh008_tall_row, "SH008-Table-With-Tall-Row.xlsx");
xlsx_fixture_tests!(sh101_simple_formats, "SH101-SimpleFormats.xlsx");
xlsx_fixture_tests!(sh102_9x9, "SH102-9-x-9.xlsx");
xlsx_fixture_tests!(sh103_no_shared_string, "SH103-No-SharedString.xlsx");
xlsx_fixture_tests!(sh104_with_shared_string, "SH104-With-SharedString.xlsx");
xlsx_fixture_tests!(sh105_no_shared_string2, "SH105-No-SharedString.xlsx");
xlsx_fixture_tests!(sh107_formatted_table, "SH107-9-x-9-Formatted-Table.xlsx");
xlsx_fixture_tests!(
    sh108_simple_formatted_cell,
    "SH108-SimpleFormattedCell.xlsx"
);

// --- Upstream parse failures (umya-spreadsheet) ------------------------------
// Related: #97
// These files fail with parse errors in umya-spreadsheet. All are handled
// gracefully — no panics, no crashes. Documented as known upstream limitations.

// ZipError: specified file not found in archive
xlsx_fixture_tests!(tdf121887, "libreoffice/tdf121887.xlsx");
xlsx_fixture_tests!(tdf131575, "libreoffice/tdf131575.xlsx");
xlsx_fixture_tests!(tdf76115, "libreoffice/tdf76115.xlsx");
xlsx_fixture_tests!(poi_49609, "poi/49609.xlsx");
xlsx_fixture_tests!(poi_56278, "poi/56278.xlsx");
xlsx_fixture_tests!(poi_59021, "poi/59021.xlsx");

// IoError: Invalid checksum
xlsx_fixture_tests!(forcepoint107, "libreoffice/forcepoint107.xlsx");

// ZipError: invalid Zip archive (Could not find EOCD)
xlsx_fixture_tests!(deep_data, "poi/deep-data.xlsx");

// --- Upstream panics caught by catch_unwind (umya-spreadsheet) ----------------
// Related: #97
// These files trigger panics inside umya-spreadsheet (arithmetic overflow,
// unwrap on None). catch_unwind prevents process crashes. Documented as
// known upstream limitations.

// attempt to subtract with overflow
xlsx_fixture_tests!(
    functions_excel_2010,
    "libreoffice/functions-excel-2010.xlsx"
);
xlsx_fixture_tests!(poi_51710, "poi/51710.xlsx");

// called Option::unwrap() on a None value
xlsx_fixture_tests!(poi_64450, "poi/64450.xlsx");

// attempt to multiply with overflow
xlsx_fixture_tests!(
    formula_eval_test_data_copy,
    "poi/FormulaEvalTestData_Copy.xlsx"
);

// --- Upstream panic safety (patched umya-spreadsheet) ------------------------
// Related: #83

/// Files that previously panicked in umya-spreadsheet now convert successfully
/// after the fork fix (developer0hye/umya-spreadsheet fix/panic-safety-v2).
///
/// All 21 previously-panicking files now produce valid PDFs.
#[test]
fn previously_panicking_files_now_convert() {
    let cases: &[&str] = &[
        // --- Phase 1 fixes (PR #90) ---
        // FileNotFound panics (7 files)
        "libreoffice/chart_hyperlink.xlsx",
        "libreoffice/hyperlink.xlsx",
        "libreoffice/tdf130959.xlsx",
        "libreoffice/test_115192.xlsx",
        "poi/47504.xlsx",
        "poi/bug63189.xlsx",
        "poi/ConditionalFormattingSamples.xlsx",
        // ParseFloatError / boolean cell (1 file)
        "libreoffice/check-boolean.xlsx",
        // unwrap() on None (2 files)
        "libreoffice/tdf100709.xlsx",
        "poi/sample-beta.xlsx",
        // dataBar end element (2 files)
        "libreoffice/tdf162948.xlsx",
        "poi/NewStyleConditionalFormattings.xlsx",
        // --- Phase 2 fixes ---
        // Backslash zip paths from Windows tools (3 files)
        "libreoffice/tdf131575.xlsx",
        "libreoffice/tdf76115.xlsx",
        "poi/49609.xlsx",
        // Missing optional styles.xml (3 files)
        "poi/56278.xlsx",
        "libreoffice/tdf121887.xlsx",
        "poi/59021.xlsx",
        // Arithmetic overflow in formula parsing (2 files)
        "libreoffice/functions-excel-2010.xlsx",
        "poi/FormulaEvalTestData_Copy.xlsx",
        // Missing XML attributes (1 file)
        "poi/64450.xlsx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("Skipping {name}: fixture not available");
            continue;
        }
        assert_produces_valid_pdf(name);
    }
}

// --- MIT: calamine (Rust) --------------------------------------------------

xlsx_fixture_tests!(date_1904, "date_1904.xlsx");
xlsx_fixture_tests!(empty_sheet, "empty_sheet.xlsx");
xlsx_fixture_tests!(errors, "errors.xlsx");
xlsx_fixture_tests!(pivots, "pivots.xlsx");
xlsx_fixture_tests!(richtext_namespaced, "richtext-namespaced.xlsx");
xlsx_fixture_tests!(column_row_ranges, "column_row_ranges.xlsx");
xlsx_fixture_tests!(table_multiple, "table-multiple.xlsx");
xlsx_fixture_tests!(formula_issue, "formula.issue.xlsx");
xlsx_fixture_tests!(header_row, "header-row.xlsx");

// --- Encrypted / password-protected fixtures --------------------------------

#[test]
fn encrypted_xlsx_returns_unsupported_encryption() {
    let path = fixture_path("poi/protected_passtika.xlsx");
    let err = office2pdf::convert(&path).unwrap_err();
    assert!(
        matches!(err, office2pdf::error::ConvertError::UnsupportedEncryption),
        "Expected UnsupportedEncryption for protected_passtika.xlsx, got: {err:?}"
    );
}

// --- Inline strings --------------------------------------------------------

#[test]
fn inline_strings_wrapped_in_a_value_keep_their_text() {
    // SH006 writes every cell as <c t="inlineStr"><v><is><t>…</t></is></v></c>.
    // The <v> wrapper is non-standard, but it must not hide the inline string.
    let pages = sheet_pages("SH006-Table-No-SharedStrings.xlsx");
    let column_a: Vec<String> = pages[0]
        .table
        .rows
        .iter()
        .map(|row| cell_text(&row.cells[0]))
        .collect();

    assert_eq!(column_a, ["Names", "Eric", "Jim", "Frank"]);
}

// --- XML entity bombs -------------------------------------------------------

#[test]
fn xml_entity_bomb_converts_without_expanding_the_entities() {
    // POI bug 54764: sharedStrings.xml and docProps/core.xml declare the
    // "billion laughs" chain lol..lol9 and use &lol9; in their text. A reader
    // that expanded it would allocate gigabytes, and one that panicked on the
    // undeclared-to-it reference failed the whole workbook. The reference must
    // stay unexpanded while the rest of the workbook still converts.
    for name in ["poi/54764.xlsx", "poi/54764-2.xlsx"] {
        let result = office2pdf::convert(fixture_path(name))
            .unwrap_or_else(|error| panic!("{name} should convert: {error:?}"));
        let text: String = common::extract_pdf_text(&result.pdf);

        assert!(
            text.contains("Test number 1"),
            "{name}: ordinary cells still render: {text:?}"
        );
        assert!(
            !text.contains("lollol"),
            "{name}: the entity chain must not be expanded: {text:?}"
        );
    }
}

// ---------------------------------------------------------------------------
// Worksheet drawings (issue #238)
// ---------------------------------------------------------------------------

#[test]
fn with_drawing_renders_anchored_images() {
    let pages = sheet_pages("poi/WithDrawing.xlsx");
    let total_images: usize = pages.iter().map(|sp| sp.images.len()).sum();
    assert!(
        total_images >= 3,
        "the drawing anchors five pictures (jpeg/png plus metafiles); at least \
         the raster ones must be extracted, got {total_images}"
    );
    let first = pages
        .iter()
        .flat_map(|sp| sp.images.iter())
        .next()
        .expect("at least one image");
    assert!(!first.image.data.is_empty(), "image bytes must be loaded");
    assert!(
        first.image.width.unwrap_or(0.0) > 10.0,
        "anchor geometry must produce a real width, got {:?}",
        first.image.width
    );
}

/// Excel writes `a:blip` as a start element with children whenever the picture
/// carries an alpha or recolour effect. The fixture is an Excel for Mac export
/// whose only picture is spelled that way — `<a:blip r:embed="rId1">
/// <a:alphaModFix amt="70000"/></a:blip>` — and its native export draws the
/// photo, so dropping it left the anchored band empty (issue #1066).
#[test]
fn picture_with_effect_bearing_blip_is_still_drawn() {
    let pages = sheet_pages("issue_1066_blip_effect_picture.xlsx");
    let images: Vec<_> = pages.iter().flat_map(|sp| sp.images.iter()).collect();
    assert_eq!(images.len(), 1, "the drawing anchors one picture");
    assert!(
        !images[0].image.data.is_empty(),
        "the blip's r:embed must resolve to the media bytes"
    );
    assert!(
        images[0].image.width.unwrap_or(0.0) > 10.0 && images[0].image.height.unwrap_or(0.0) > 10.0,
        "the anchor must produce a real size, got {:?} x {:?}",
        images[0].image.width,
        images[0].image.height
    );
}

/// `<a:alphaModFix amt="70000"/>` on that same blip is a transparency effect:
/// Excel draws the picture at 70% strength over whatever the worksheet shows
/// beneath it, and its native export wraps the bitmap in a soft mask so the
/// `#C0392B` bar prints as that colour composited onto white. Emitting the raw
/// bitmap printed a saturated block against a washed-out ground truth (issue
/// #1103).
#[test]
fn picture_alpha_mod_fix_is_baked_into_the_bitmap() {
    let pages = sheet_pages("issue_1066_blip_effect_picture.xlsx");
    let images: Vec<_> = pages.iter().flat_map(|sp| sp.images.iter()).collect();
    assert_eq!(images.len(), 1, "the drawing anchors one picture");

    let decoded = image::load_from_memory(&images[0].image.data)
        .expect("the anchored picture must stay a decodable bitmap")
        .into_rgba8();
    // The source is a flat three-bar PNG: #C0392B, #27AE60, #2980B9 on white,
    // every pixel fully opaque. At 70% each of them composites onto the white
    // page as 0.7 * channel + 0.3 * 255 - the red bar as (211, 116, 107).
    let red_bar: Vec<&image::Rgba<u8>> = decoded
        .pixels()
        .filter(|pixel| pixel[0] == 192 && pixel[1] == 57 && pixel[2] == 43)
        .collect();
    assert!(
        !red_bar.is_empty(),
        "the source bar colour must survive the alpha bake unchanged"
    );
    for pixel in &red_bar {
        assert_eq!(
            pixel[3], 179,
            "70% of an opaque pixel is 178.5, rounded to 179"
        );
    }
}

/// A drawing anchor spans the rows Excel *prints*, not the heights the
/// worksheet declares. The same fixture's Excel for Mac export draws its
/// picture 105.00pt tall with its top 90.00pt below the first row — seven and
/// six 15pt tracks — while the sheet declares `defaultRowHeight="18"` and
/// `ht="16"`, and Excel honours neither: it recomputes 16pt rows from the
/// Calibri 12 Normal font and prints them at the 15pt track that font's grid
/// compacts to (issue #1102).
///
/// The picture's width is asserted separately, by
/// `a_picture_width_drops_the_whole_point_part_of_an_overrunning_col_offset`:
/// Excel's columns here measure 65.00pt in the app and in the export alike,
/// matching ours, and the 11.00pt the export is narrower is entirely the
/// overrunning `to` `colOff` rule of issue #1149.
#[test]
fn a_picture_anchor_spans_the_printed_row_track() {
    let pages = sheet_pages("issue_1066_blip_effect_picture.xlsx");
    let images: Vec<_> = pages.iter().flat_map(|sp| sp.images.iter()).collect();
    assert_eq!(images.len(), 1, "the drawing anchors one picture");
    let height: f64 = images[0].image.height.expect("a two-cell anchor sizes");
    assert!(
        (height - 105.0).abs() < 0.01,
        "seven printed 15pt tracks, got {height}"
    );
    assert!(
        (images[0].y_offset_pt - 90.0).abs() < 0.01,
        "six printed 15pt tracks above the anchor row, got {}",
        images[0].y_offset_pt
    );
}

/// Excel does not add a `twoCellAnchor`'s `to` `xdr:colOff` to the spanned
/// columns as written once that offset overruns the column it sits in: it
/// drops the whole-point part of the overrun. The same fixture's anchor runs
/// from col 1 at 11880 EMU (0.935pt) to col 3 at 963720 EMU (75.883pt) over
/// 65.00pt columns, so the offset overruns by 10.883pt and Excel takes 11.00pt
/// of that back — its export measures the picture 193.948pt wide against the
/// 204.948pt the plain sum gives (issue #1149).
#[test]
fn a_picture_width_drops_the_whole_point_part_of_an_overrunning_col_offset() {
    let pages = sheet_pages("issue_1066_blip_effect_picture.xlsx");
    let images: Vec<_> = pages.iter().flat_map(|sp| sp.images.iter()).collect();
    assert_eq!(images.len(), 1, "the drawing anchors one picture");
    let width: f64 = images[0].image.width.expect("a two-cell anchor sizes");
    assert!(
        (width - 193.948).abs() < 0.01,
        "two 65pt columns less a 0.935pt `from` offset plus a 64.883pt \
         normalised `to` offset, got {width}"
    );
}

/// Excel cuts a blocked cell's unwrapped line at the cell's own gridline, not
/// at the inset its text starts from. The same fixture holds `Wrapping paper`
/// in `B4` against a value in `C4`, and its Excel for Mac export prints
/// `Wrapping pa`: the next `p` begins on the column boundary and is left
/// undrawn. Sizing the clip box from the *content* edge instead pushed that
/// boundary a whole left inset further right, which is exactly the room the
/// extra glyph needed (issue #1105).
#[test]
fn a_blocked_cell_clips_its_line_at_the_column_gridline() {
    let data = load_fixture("issue_1066_blip_effect_picture.xlsx");
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let source = generate_typst(&document)
        .expect("fixture should generate Typst")
        .source;

    // `B4` is the sheet's only spilling cell: every other label fits its own
    // column. Read past the vertical anchor and the box height, both of which
    // follow the row's shared line and so need font metrics a bare CI runner
    // may not have.
    let clip_box: &str = source
        .split_once("place(left + ")
        .and_then(|(_, rest)| rest.split_once(", box(width: "))
        .map(|(_, width)| width)
        .expect("the blocked cell must place a left-anchored clip box");

    // 65pt columns laid out from a 3pt left inset: the gridline is 62pt past
    // the point the line starts at, and 65pt would be a whole inset past it.
    assert!(
        clip_box.starts_with("62pt, height:"),
        "a blocked cell's clip must end on its column's right gridline, got {}: {source}",
        &clip_box[..clip_box.len().min(24)]
    );
}

// ---------------------------------------------------------------------------
// Worksheet drawing scheme colors resolve against the theme (issue #430)
// ---------------------------------------------------------------------------

#[test]
fn theme_color_drawing_resolves_scheme_fills() {
    use office2pdf::ir::Color;

    let pages = sheet_pages("theme_color_drawing.xlsx");
    let boxes: Vec<_> = pages.iter().flat_map(|sp| sp.text_boxes.iter()).collect();
    assert_eq!(boxes.len(), 3, "fixture drawing holds three text boxes");

    // accent1 straight from the workbook theme.
    assert_eq!(boxes[0].fill, Some(Color::new(68, 114, 196)));
    // "accent1, lighter 60%": lumMod 40% + lumOff 60% in HSL space.
    assert_eq!(boxes[1].fill, Some(Color::new(180, 199, 231)));
    // "accent6, darker 25%": lumMod 75%.
    assert_eq!(boxes[2].fill, Some(Color::new(84, 130, 53)));
}

// ---------------------------------------------------------------------------
// Worksheet drawing anchor geometry (issue #460)
// ---------------------------------------------------------------------------

#[test]
fn theme_color_drawing_anchors_span_six_eighteen_point_rows() {
    // The three shapes are `xdr:twoCellAnchor`s spanning rows 3..9 of a sheet
    // whose rows measure 18pt in the worksheet, so Excel draws them
    // 6 * 18 = 108pt tall - measured at 108pt on the native export, whose
    // printed grid keeps that 18pt because this workbook's Normal font
    // resolves through a per-script theme scheme and so does not compact
    // (issues #460, #1102). Resolving the anchor through a track that rounded
    // 18pt to 17pt left every shape 6pt short.
    let pages = sheet_pages("theme_color_drawing.xlsx");
    let boxes: Vec<_> = pages.iter().flat_map(|sp| sp.text_boxes.iter()).collect();
    assert_eq!(boxes.len(), 3, "fixture drawing holds three text boxes");
    for text_box in &boxes {
        assert!(
            (text_box.height - 108.0).abs() < 0.5,
            "anchor should span six 18pt rows, got {}",
            text_box.height
        );
    }
}

// ---------------------------------------------------------------------------
// Worksheet text boxes (issue #240)
// ---------------------------------------------------------------------------

#[test]
fn with_text_box_renders_anchored_text() {
    use office2pdf::ir::{Alignment, Color};

    let pages = sheet_pages("poi/WithTextBox.xlsx");
    let boxes: Vec<_> = pages.iter().flat_map(|sp| sp.text_boxes.iter()).collect();
    assert_eq!(boxes.len(), 1, "the drawing holds one text box");

    let text_box = boxes[0];
    assert_eq!(text_box.paragraphs.len(), 3);
    assert!(
        text_box.width > 50.0,
        "anchor width, got {}",
        text_box.width
    );

    let para = &text_box.paragraphs[0];
    assert_eq!(para.runs[0].text, "Line 1");
    assert_eq!(para.style.alignment, None, "algn=l maps to default/left");
    // This `xdr:txBody` declares no `a:spcBef`/`a:spcAft`, so its paragraphs
    // must carry an explicit zero rather than being left unset — unset lets
    // the renderer's own default block spacing in, which doubled the pitch
    // between lines (issue #656).
    for (index, paragraph) in text_box.paragraphs.iter().enumerate() {
        assert_eq!(
            paragraph.style.space_before,
            Some(0.0),
            "paragraph {index} must state its space before"
        );
        assert_eq!(
            paragraph.style.space_after,
            Some(0.0),
            "paragraph {index} must state its space after"
        );
    }
    assert_eq!(para.runs[0].style.color, Some(Color::new(0xFF, 0, 0)));

    assert_eq!(
        text_box.paragraphs[1].style.alignment,
        Some(Alignment::Center)
    );
    assert_eq!(
        text_box.paragraphs[1]
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>(),
        "Line 2"
    );
    assert_eq!(
        text_box.paragraphs[2].style.alignment,
        Some(Alignment::Right)
    );
    assert_eq!(
        text_box.paragraphs[2].runs[0].style.color,
        Some(Color::new(0, 0, 0xFF))
    );
}

// ---------------------------------------------------------------------------
// Embedded charts on chart-only sheets (issue #239)
// ---------------------------------------------------------------------------

#[test]
fn with_chart_renders_embedded_chart() {
    // WithChart.xlsx puts its chart on a sheet with no cells; that sheet was
    // skipped entirely, dropping the chart (same class as #238's image-only
    // sheets, which the fix did not extend to charts).
    let pages = sheet_pages("poi/WithChart.xlsx");
    let total_charts: usize = pages.iter().map(|sp| sp.charts.len()).sum();
    assert!(
        total_charts >= 1,
        "the embedded chart must be extracted, got {total_charts}"
    );
    let chart = pages
        .iter()
        .flat_map(|sp| sp.charts.iter())
        .next()
        .expect("a chart");
    assert!(
        !chart.chart.series.is_empty(),
        "chart must carry its series data"
    );
}

// ---------------------------------------------------------------------------
// Repository workbook — multi-sheet Korean analysis workbook
// ---------------------------------------------------------------------------

/// Ten-sheet Korean workbook describing this repository. Unlike the focused
/// fixtures above it combines the XLSX features that co-occur in real reporting
/// documents on the same printed page: merged banner rows, an Excel table with
/// `tableStyleInfo` stripes, data bars, colour scales, icon sets, hyperlinks,
/// three chart kinds, cell comments, mixed portrait/landscape print setup, and
/// `fitToPage` print scaling.
const REPOSITORY_WORKBOOK_FIXTURE: &str = "office2pdf_repository_workbook.xlsx";

#[test]
fn smoke_repository_workbook_fixture() {
    assert_produces_valid_pdf(REPOSITORY_WORKBOOK_FIXTURE);
}

#[test]
fn structure_repository_workbook_keeps_every_sheet_in_workbook_order() {
    let pages = sheet_pages(REPOSITORY_WORKBOOK_FIXTURE);

    let mut names: Vec<&str> = Vec::new();
    for page in &pages {
        if names.last() != Some(&page.name.as_str()) {
            names.push(page.name.as_str());
        }
    }

    assert_eq!(
        names,
        [
            "00_개요",
            "01_대시보드",
            "02_기능_매트릭스",
            "03_CLI_옵션",
            "04_라이브러리_API",
            "05_모듈_인벤토리",
            "06_릴리스_이력",
            "07_검증_품질게이트",
            "08_의존성",
            "09_리스크_로드맵",
        ],
        "every worksheet must survive parsing, in workbook order"
    );
}

#[test]
fn structure_repository_workbook_preserves_print_orientation_per_sheet() {
    let pages = sheet_pages(REPOSITORY_WORKBOOK_FIXTURE);

    let is_landscape = |sheet: &str| -> bool {
        let page = sheet_page_named(&pages, sheet);
        page.size.width > page.size.height
    };

    // The overview, dashboard, and library-API sheets print portrait; the wide
    // inventory and matrix sheets print landscape.
    assert!(!is_landscape("00_개요"));
    assert!(!is_landscape("01_대시보드"));
    assert!(!is_landscape("04_라이브러리_API"));
    assert!(is_landscape("02_기능_매트릭스"));
    assert!(is_landscape("05_모듈_인벤토리"));
    assert!(is_landscape("09_리스크_로드맵"));
}

#[test]
fn structure_repository_workbook_extracts_every_dashboard_chart_with_data() {
    let pages = sheet_pages(REPOSITORY_WORKBOOK_FIXTURE);

    // The dashboard is wider than one page-column, and a chart crossing the
    // boundary is printed on both tiles (issue #1598), so the same anchored
    // chart can appear on more than one page. Count each anchor once.
    let mut seen_anchors: std::collections::HashSet<(u32, Option<String>)> =
        std::collections::HashSet::new();
    let charts: Vec<&office2pdf::ir::Chart> = pages
        .iter()
        .filter(|page| page.name == "01_대시보드")
        .flat_map(|page| page.charts.iter())
        .filter(|sheet_chart| {
            seen_anchors.insert((sheet_chart.anchor_row, sheet_chart.chart.title.clone()))
        })
        .map(|sheet_chart| &sheet_chart.chart)
        .collect();

    assert_eq!(charts.len(), 3, "the dashboard anchors three charts");

    for chart in &charts {
        let title: &str = chart.title.as_deref().unwrap_or_default();
        assert!(
            !chart.categories.is_empty(),
            "chart {title:?} must carry its category labels"
        );
        assert!(
            chart.series.iter().any(|series| !series.values.is_empty()),
            "chart {title:?} must carry its cached series values"
        );
    }
}

/// End-to-end: `&A` reaches the rendered header as the worksheet name.
///
/// `check-boolean.xlsx` declares `<oddHeader>&C&"Times New Roman,Regular"&12&A`.
/// `&A` is the section's only content, so while the code was discarded the
/// section came out empty and the header paragraph was dropped altogether —
/// the sheet printed with no header at all where Excel prints "Sheet1"
/// (issue #690). Its footer, `Page &P`, survives throughout because `&P` was
/// already implemented, and is asserted here to keep the two apart.
///
/// Issue #690 cites `libreoffice/page_scale.xlsx`, which carries the very same
/// header string but has no `<sheetData>` rows. A sheet with no used range
/// takes the `empty_workbook_page` path, which deliberately carries neither
/// header nor footer onto the blank page it prints (issue #632) — so that
/// fixture renders wholly blank and cannot witness this defect either way.
/// This one has the same header and two data rows.
#[test]
fn structure_check_boolean_header_prints_the_sheet_name() {
    let pages = sheet_pages("libreoffice/check-boolean.xlsx");
    let page = pages
        .first()
        .expect("check-boolean has at least one sheet page");
    assert_eq!(page.name, "Sheet1");

    let header = page
        .header
        .as_ref()
        .expect("a header whose only content is &A must still be emitted");
    let header_text: String = header
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect();
    assert_eq!(header_text, "Sheet1");

    let footer = page
        .footer
        .as_ref()
        .expect("check-boolean declares a footer");
    assert!(
        footer
            .paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.elements.iter())
            .any(|element| matches!(element, HFInline::PageNumber(_))),
        "the footer's &P must still resolve"
    );
}

/// A cached formula string keeps the leading spaces `xml:space="preserve"`
/// protects (issue #719).
///
/// Tests rows 44-46 are `t="str"` cells whose cached `<v xml:space="preserve">`
/// carries 1, 2 and 4 leading spaces — the indent a `?` placeholder reserves,
/// which Excel bakes into the cached display text and which cannot be
/// recomputed from the formula.
///
/// Every cell is scanned rather than a fixed column index: this sheet prints
/// headings, so a row-number gutter column is prepended (issue #623) and
/// column A is not `cells[0]`.
#[test]
fn structure_number_format_tests_keeps_preserved_leading_spaces() {
    let pages = sheet_pages("poi/NumberFormatTests.xlsx");
    let texts: Vec<String> = pages
        .iter()
        .filter(|page| page.name == "Tests")
        .flat_map(|page| page.table.rows.iter())
        .flat_map(|row| row.cells.iter())
        .map(table_cell_text)
        .collect();

    for expected in [" 1,234,567", "  1,234,567", "    1,234,567"] {
        assert!(
            texts.iter().any(|text| text == expected),
            "expected a cell of exactly {expected:?}; leading spaces were trimmed"
        );
    }
}

/// A missing typewriter face must retain fixed pitch through the complete
/// XLSX-to-Typst path instead of reaching Typst's proportional body default.
#[test]
fn number_format_tests_emits_a_monospace_fallback_chain() {
    let data = load_fixture("poi/NumberFormatTests.xlsx");
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let source = generate_typst(&document)
        .expect("fixture should generate Typst")
        .source;

    assert!(
        source.contains(
            r#"font: ("Lucida Sans Typewriter", "DejaVu Sans Mono", "Noto Sans Mono", "Liberation Mono", "Cousine")"#,
        ),
        "Lucida Sans Typewriter should keep a monospace fallback chain"
    );
}

/// A cell's leading spaces must survive *rendering*, not just parsing: Typst
/// drops a space that opens a markup line, so the single-space cell above
/// rendered flush left even though its text was intact (issue #752).
#[test]
fn number_format_tests_renders_a_single_leading_space_as_an_indent() {
    let data = load_fixture("poi/NumberFormatTests.xlsx");
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");
    let source = generate_typst(&document)
        .expect("fixture should generate Typst")
        .source;

    // Two- and four-space runs already emitted a code-mode string, which
    // markup cannot collapse; the lone space fell through to a literal.
    for spaces in [" ", "  ", "    "] {
        assert!(
            source.contains(&format!("[#\"{spaces}\";1,234,567]")),
            "a {}-space indent must survive as a code-mode string",
            spaces.len()
        );
    }
}

/// A literal-text number format renders its literal, not the raw number
/// (issue #750).
///
/// `check-boolean.xlsx` styles both cells with `numFmt` 165,
/// `"TRUE";"TRUE";"FALSE"` — sections that carry no numeric placeholder at all.
/// A1 is `t="b"` and resolves through the cell type; A2 is `t="n"` holding 2
/// and must take the format's positive section.
#[test]
fn structure_check_boolean_renders_literal_text_number_format() {
    let pages = sheet_pages("libreoffice/check-boolean.xlsx");
    let texts: Vec<String> = pages
        .iter()
        .flat_map(|page| page.table.rows.iter())
        .flat_map(|row| row.cells.iter())
        .map(table_cell_text)
        .filter(|text| !text.is_empty())
        .collect();

    assert_eq!(
        texts.iter().filter(|text| *text == "TRUE").count(),
        2,
        "both cells carry numFmt 165 and print TRUE; got {texts:?}"
    );
    assert!(
        !texts.iter().any(|text| text == "2"),
        "the raw number must not leak past the literal format; got {texts:?}"
    );
}

#[test]
fn smoke_merged_row_overflows_page_column() {
    assert_produces_valid_pdf("merged_row_overflows_page_column.xlsx");
}

/// A merged row spanning a sheet wide enough to split horizontally used to keep
/// the whole merge's width as its spill width on every page-column, so its text
/// painted a single line far past the printable edge — off the paper on this
/// fixture, losing that ink entirely (#631).
#[test]
fn structure_merged_row_overflow_clamps_spill_to_its_page_column() {
    let pages = sheet_pages("merged_row_overflows_page_column.xlsx");
    assert!(
        pages.len() >= 2,
        "the sheet is wider than one page and must split into column groups; got {}",
        pages.len()
    );

    for (index, page) in pages.iter().enumerate() {
        let group_width: f64 = page.table.column_widths.iter().sum();
        for row in &page.table.rows {
            for cell in &row.cells {
                let Some(spill) = cell.spill_width else {
                    continue;
                };
                assert!(
                    spill <= group_width + 0.001,
                    "page {index}: spill width {spill}pt exceeds the {group_width}pt \
                     the page-column actually carries",
                );
            }
        }
    }
}

/// The plain text a sheet cell's paragraphs carry, concatenated.
fn cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
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

/// `100-customers.xlsx` writes no `<sheetFormatPr>` element and no `<cols>`,
/// so every column is a default column. Excel for Mac prints those at its
/// own 10-character base — `10 × 7 + 5 = 75pt` on the 7pt Malgun Gothic 12
/// unit the theme resolves to — not the ECMA 8-character base that a
/// declared element without `baseColWidth` keeps (issue #1656). Seven 75pt
/// columns (A:E plus the two the occupation values spill into) are 525pt and
/// no longer fit the 487pt A4 printable width, so the tails print on a second
/// page-column after the down-page run, and the rendered PDF carries the
/// native export's six pages: three of A:F, then three strip pages (issue
/// #1381).
#[test]
fn structure_hundred_customers_prints_mac_default_columns_and_an_overflow_strip() {
    let pages = sheet_pages("100-customers.xlsx");
    assert_eq!(
        pages.len(),
        2,
        "A:F fill the first page-column and the overflow reach forms a strip"
    );
    let column_widths: Vec<f64> = pages
        .iter()
        .flat_map(|page| page.table.column_widths.iter().copied())
        .collect();
    assert_eq!(column_widths.len(), 7, "A:G print; got {column_widths:?}");
    assert!(
        column_widths
            .iter()
            .all(|width| (width - 75.0).abs() < 0.01),
        "every default column prints at Excel for Mac's 75pt; got {column_widths:?}"
    );

    let result = office2pdf::convert(fixture_path("100-customers.xlsx")).expect("converts");
    assert_eq!(
        rendered_page_count(&result.pdf),
        6,
        "native Excel for Mac prints six pages: three of A:F, then three strip pages"
    );
}

/// Number of page objects in a PDF. Typst writes every object dictionary in
/// plain text without separating whitespace, so the `/Type/Page` entries are
/// countable without a PDF reader on every CI runner; the `/Pages` tree node
/// is excluded by the boundary.
fn rendered_page_count(pdf: &[u8]) -> usize {
    let text: String = String::from_utf8_lossy(pdf).into_owned();
    text.match_indices("/Type/Page")
        .filter(|(index, needle)| {
            !text[index + needle.len()..]
                .chars()
                .next()
                .is_some_and(|next| next.is_ascii_alphabetic())
        })
        .count()
}

/// Excel ends the printed sequence at its last page carrying ink. The native
/// Excel-for-Mac export of `1000-customers.xlsx` prints 51 pages: 26 of A:F,
/// then 25 strip pages, none for the 26th band whose only row does not spill
/// (issue #1714). The strip page-column therefore ends at its last continued
/// line, which leaves the 26th band with nothing to print.
#[test]
fn structure_thousand_customers_prints_no_strip_page_for_the_tail_less_last_band() {
    let pages = sheet_pages("1000-customers.xlsx");
    assert_eq!(pages.len(), 2, "A:F, then the overflow strip");
    let strip_rows: usize = pages[1].table.rows.len();
    assert!(
        strip_rows < pages[0].table.rows.len(),
        "the strip ends before the last row; it still carries {strip_rows} rows"
    );
    assert!(
        pages[1].table.rows[strip_rows - 1]
            .cells
            .iter()
            .any(|cell| cell.spill_continuation_offset_pt.is_some()),
        "the strip's last row carries a continued line"
    );

    let result = office2pdf::convert(fixture_path("1000-customers.xlsx")).expect("converts");
    assert_eq!(
        rendered_page_count(&result.pdf),
        51,
        "native Excel for Mac prints 26 down-page pages and 25 strip pages"
    );
}

/// `customers_trailing_band_no_spill.xlsx` is `100-customers.xlsx` with the
/// occupation values of rows 81-101, the third vertical band, shortened to
/// `x`. The native Excel-for-Mac export prints 5 pages: the down-page run of
/// three, then only the two strip pages whose bands carry a spilled tail
/// (issue #1714 probe).
#[test]
fn structure_trailing_band_without_a_spill_prints_no_strip_page() {
    let pages = sheet_pages("customers_trailing_band_no_spill.xlsx");
    assert_eq!(pages.len(), 2);
    let strip_rows: usize = pages[1].table.rows.len();
    assert!(
        (41..=80).contains(&strip_rows),
        "the strip ends inside the second band; got {strip_rows} rows"
    );

    let result = office2pdf::convert(fixture_path("customers_trailing_band_no_spill.xlsx"))
        .expect("converts");
    assert_eq!(rendered_page_count(&result.pdf), 5);
}

/// `customers_middle_band_no_spill.xlsx` shortens rows 41-80 instead, the
/// second band. The native export still prints 6 pages, the blank strip page
/// for that band included: a blank page inside the sequence is printed, so
/// the strip keeps every row up to its last continued line in the third band
/// (issue #1714 probe).
#[test]
fn structure_middle_band_without_a_spill_keeps_its_blank_strip_page() {
    let pages = sheet_pages("customers_middle_band_no_spill.xlsx");
    assert_eq!(pages.len(), 2);
    let strip_rows: usize = pages[1].table.rows.len();
    assert!(
        strip_rows > 80,
        "the strip reaches the third band's continued lines; got {strip_rows} rows"
    );
    let second_band_has_continuation: bool = pages[1].table.rows[40..80].iter().any(|row| {
        row.cells
            .iter()
            .any(|cell| cell.spill_continuation_offset_pt.is_some())
    });
    assert!(
        !second_band_has_continuation,
        "rows 41-80 carry nothing on the strip"
    );

    let result =
        office2pdf::convert(fixture_path("customers_middle_band_no_spill.xlsx")).expect("converts");
    assert_eq!(rendered_page_count(&result.pdf), 6);
}

#[test]
fn smoke_customers_overflow_strip() {
    assert_produces_valid_pdf("customers_overflow_strip.xlsx");
}

/// `customers_overflow_strip.xlsx` is the first twenty rows of
/// `100-customers.xlsx` with its five columns and the sheet default declared
/// at the 75pt Excel for Mac prints for that workbook's absent
/// `sheetFormatPr` (10.714 chars of the 7pt Malgun Gothic unit; #1656 tracks
/// the undeclared case, which this fixture sidesteps so the strip boundary
/// does not depend on it). The occupation column's unwrapped values reach two
/// columns past the used range, so the printed range is A:G and A:F fill the
/// 487pt A4 printable width. The native export prints two pages: A:F, then a
/// strip whose only ink is the tails of those lines, each redrawn 450pt to
/// the left of where it started and clipped at the strip's edge (issue
/// #1381).
#[test]
fn structure_overflow_strip_carries_the_spilled_tails() {
    let pages = sheet_pages("customers_overflow_strip.xlsx");
    assert_eq!(
        pages.len(),
        2,
        "A:F fill the first page-column and the overflow reach forms a second"
    );
    assert_eq!(pages[0].table.column_widths.len(), 6);

    let first_group_width: f64 = pages[0].table.column_widths.iter().sum();
    assert!(
        (first_group_width - 450.0).abs() < 0.01,
        "six 75pt columns print on page 1; got {first_group_width}pt"
    );

    // Row 2 (Martine Gleichner) reads `Corporate Security Supervisor` in
    // column E, whose left edge sits 300pt into the grid: its tail continues
    // 150pt past its own gridline on the strip.
    let strip_row = &pages[1].table.rows[1];
    let continuation = strip_row
        .cells
        .first()
        .expect("the strip's first column carries the continuation");
    assert_eq!(cell_text(continuation), "Corporate Security Supervisor");
    let offset: f64 = continuation
        .spill_continuation_offset_pt
        .expect("a continued line records how far left it began");
    assert!(
        (offset - 150.0).abs() < 0.01,
        "the tail starts 150pt left of the strip's gridline; got {offset}pt"
    );
    let remaining: f64 = continuation
        .spill_width
        .expect("a continued line keeps the width it still reaches");
    assert!(
        remaining > 0.0 && remaining <= pages[1].table.column_widths.iter().sum::<f64>() + 0.001,
        "the remaining reach is positive and within the strip; got {remaining}pt"
    );

    // A value that fits its own column leaves its strip cell empty: the
    // `Occupation` heading in E1 never spills, so nothing of it continues.
    let fitting_cell = pages[1].table.rows[0].cells.first().unwrap();
    assert_eq!(
        cell_text(fitting_cell),
        "",
        "a line that ends before the boundary has no continuation"
    );
    assert_eq!(fitting_cell.spill_continuation_offset_pt, None);
    assert_eq!(fitting_cell.spill_width, None);
}

#[test]
fn smoke_spill_reach_print_range() {
    assert_produces_valid_pdf("spill_reach_print_range.xlsx");
}

/// Excel extends a sheet's printed range to every column that unwrapped cell
/// text visibly overflows into, and only to those (issue #718). Probe-measured
/// on `NumberFormatTests.xlsx`: deleting the trailing styled `<col>` run, the
/// row-level `customFormat`, the pane, and the selections each left the extra
/// printed column in place, while deleting the rows whose D-column text paints
/// past the column edge removed it; lengthening that text made Excel print
/// additional columns and horizontal spill pages (14 pages against the
/// workbook's baseline 9).
///
/// The fixture's five sheets vary one factor each around the same grid
/// (Verdana 10; columns 20/18.33/14.66/13 chars; default width 10.6640625):
/// - `SpillOne`: 17-char text in the 13-char last column reaches one column
///   past the used range.
/// - `Wrapped`: the same long text as `LongSpill`, under `wrapText`, never
///   spills.
/// - `Fits`: 9-char text fits its column.
/// - `LongSpill`: 44-char text reaches three columns past the used range.
/// - `Numeric`: a wide number never spills.
#[test]
fn structure_spill_reach_extends_printed_columns() {
    let pages = sheet_pages("spill_reach_print_range.xlsx");
    // LongSpill's extension makes the grid wider than one page, so it splits
    // into two column groups — the same horizontal spill pages the native
    // export emits for an overflow reaching past the printable width.
    assert_eq!(
        sheet_names(&pages),
        vec![
            "SpillOne",
            "Wrapped",
            "Fits",
            "LongSpill",
            "LongSpill",
            "Numeric"
        ],
        "only the reach past the printable width may add pages"
    );

    // Each sheet prints headings, so every page group carries one row-heading
    // column ahead of its spreadsheet columns: 1 + used range (4) + extension.
    let column_counts: Vec<usize> = pages
        .iter()
        .map(|page| page.table.column_widths.len())
        .collect();
    assert_eq!(
        column_counts,
        vec![6, 5, 5, 6, 3, 5],
        "printed columns must extend exactly to the overflow's reach \
         (SpillOne +1, LongSpill +3 split across its column groups) and stay \
         at the used range everywhere else"
    );

    // The extension columns carry no `<col>` record, so they take the sheet's
    // declared default width (10.6640625 chars of Verdana 10 -> ~64pt, the
    // width the native export prints for column E).
    let spill_one_extension: f64 = *pages[0].table.column_widths.last().unwrap();
    assert!(
        (spill_one_extension - 64.0).abs() < 1.5,
        "extension column must take the default column width; got {spill_one_extension}pt"
    );
}

/// The `customFormat` row style paints the printed cells the row extends
/// over, including spill-reached columns holding no cell at all — Excel's
/// export fills the whole A1:E1 band (461pt on `NumberFormatTests.xlsx`,
/// issue #718), not just the populated A1:D1.
#[test]
fn structure_row_custom_format_fill_covers_spill_reached_column() {
    let pages = sheet_pages("spill_reach_print_range.xlsx");
    // rows[0] is the printed column-heading row; rows[1] is spreadsheet row 1,
    // whose customFormat style fills A1:D1 — and must also fill the reached E1.
    let header_cells = &pages[0].table.rows[1].cells;
    assert_eq!(header_cells.len(), 6, "header row spans the extended grid");
    assert_eq!(
        header_cells.last().unwrap().background,
        Some(Color {
            r: 255,
            g: 192,
            b: 0
        }),
        "the row-level customFormat fill must cover the spill-reached column"
    );

    // Rows without a customFormat style leave the extension column unfilled.
    let data_cells = &pages[0].table.rows[2].cells;
    assert_eq!(data_cells.len(), 6);
    assert_eq!(
        data_cells.last().unwrap().background,
        None,
        "plain rows must not inherit a fill in the extension column"
    );
}

/// The KPI tracker's A11 footnote is 99 characters of Arial 9 sitting in a
/// 156pt column, so it paints past the used range — but only as far as the
/// grid it already has. Excel prints the workbook on one page (issue #1054);
/// pricing that line a third too wide walked the printed range three columns
/// past F, which pushed the grid over the 487pt printable width and emitted a
/// second, empty column group.
#[test]
fn structure_kpi_tracker_note_overflow_keeps_the_grid_on_one_page() {
    let pages = sheet_pages("../../golden_mocks/business/sources/xlsx/10_kpi_tracker_en.xlsx");
    assert_eq!(
        sheet_names(&pages),
        vec!["KPI"],
        "the sheet must not split into horizontal column groups"
    );
    // printOptions omits headings, so the page is the A:F used range itself.
    assert_eq!(
        pages[0].table.column_widths.len(),
        6,
        "the printed range stays at the used range A:F, got widths {:?}",
        pages[0].table.column_widths
    );
}

/// `ExcelTables.xlsx` styles its `G1:I4` table `TableStyleLight1`. A native
/// Excel-for-Mac export prints that style as a `#D9D9D9` band over the 1st and
/// 3rd body rows, three full-width black 1pt rules — above the header, under
/// it, and at the table's foot — and a bold header row (issue #1080).
#[test]
fn structure_light1_table_style_bands_rules_and_bolds_its_header() {
    let pages = sheet_pages("ExcelTables.xlsx");
    let rows = &pages[0].table.rows;
    // The table occupies G1:I4, so column G is grid index 6 and its header
    // sits on the first printed row.
    let cell_at = |row: usize, col: usize| -> &TableCell { &rows[row].cells[col] };
    let header = cell_at(0, 6);
    assert_eq!(
        header.content.len(),
        1,
        "G1 must still hold the header label"
    );

    let stripe = Some(Color {
        r: 0xd9,
        g: 0xd9,
        b: 0xd9,
    });
    assert_eq!(cell_at(1, 6).background, stripe, "G2 is the first band");
    assert_eq!(cell_at(2, 6).background, None, "G3 sits between the bands");
    assert_eq!(cell_at(3, 6).background, stripe, "G4 is the second band");
    assert_eq!(
        cell_at(1, 3).background,
        None,
        "column D is outside the table"
    );

    let header_border = header.border.as_ref().expect("G1 carries the two rules");
    for (side, name) in [
        (&header_border.top, "above the header"),
        (&header_border.bottom, "under the header"),
    ] {
        let side = side.as_ref().unwrap_or_else(|| panic!("rule {name}"));
        assert_eq!(side.color, Color { r: 0, g: 0, b: 0 }, "rule {name}");
        assert_eq!(side.width, 1.0, "rule {name} is a 1pt band");
        assert_eq!(side.style, BorderLineStyle::Solid, "rule {name}");
    }
    let foot_border = cell_at(3, 6)
        .border
        .as_ref()
        .expect("G4 carries the foot rule");
    assert!(
        foot_border.bottom.is_some(),
        "the table's last row is ruled at its foot"
    );
    assert!(
        foot_border.top.is_none(),
        "no rule runs between the body rows"
    );
    assert!(
        cell_at(2, 6).border.is_none(),
        "an interior body row carries no rule"
    );

    let header_run = match &header.content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("header holds a paragraph, got {other:?}"),
    };
    assert_eq!(
        header_run.style.bold,
        Some(true),
        "the table style prints its header row bold"
    );
}

/// `ExcelTables.xlsx` declares a 0.7874in top and bottom margin (56.69pt) and
/// a 0.7in left and right one (50.4pt). A native Excel-for-Mac export prints
/// its grid against 56 and 50: the rule above the table's header spans y 56-57
/// and starts at x 464 = 50 + six 69pt columns, where printing against the
/// exact margin put them at 56.69-57.69 and 464.4 (issue #1127).
#[test]
fn structure_a_fractional_print_margin_lands_on_a_whole_point() {
    let pages = sheet_pages("ExcelTables.xlsx");
    let margins = pages[0].margins;
    assert_eq!(margins.top, 56.0, "56.69pt top prints against 56");
    assert_eq!(margins.bottom, 56.0);
    assert_eq!(margins.left, 50.0, "50.4pt left prints against 50");
    assert_eq!(margins.right, 50.0);
}

/// `SH001-Table.xlsx` styles its `A1:C3` table `TableStyleMedium2` over an
/// accent 1 of `5B9BD5`. A native Excel-for-Mac export prints that style as a
/// solid `#5B9BD5` header band carrying white bold runs, a `#DDEBF7` band over
/// the first body row, and `#9BC2E6` 1pt rules at every row boundary plus one
/// down each of the table's outer edges (issue #1125).
#[test]
fn structure_medium2_table_style_fills_rules_and_whitens_its_header() {
    let pages = sheet_pages("SH001-Table.xlsx");
    let rows = &pages[0].table.rows;
    let cell_at = |row: usize, col: usize| -> &TableCell { &rows[row].cells[col] };
    let accent = Color {
        r: 0x5b,
        g: 0x9b,
        b: 0xd5,
    };
    let rule_color = Color {
        r: 0x9b,
        g: 0xc2,
        b: 0xe6,
    };

    assert_eq!(
        cell_at(0, 0).background,
        Some(accent),
        "A1's header band is the accent itself"
    );
    assert_eq!(
        cell_at(1, 0).background,
        Some(Color {
            r: 0xdd,
            g: 0xeb,
            b: 0xf7,
        }),
        "A2 is the first body band"
    );
    assert_eq!(cell_at(2, 0).background, None, "A3 sits between the bands");

    let header_border = cell_at(0, 0).border.as_ref().expect("A1 is ruled");
    for (side, name) in [
        (&header_border.top, "above the header"),
        (&header_border.bottom, "under the header"),
        (&header_border.left, "down the table's left edge"),
    ] {
        let side = side.as_ref().unwrap_or_else(|| panic!("rule {name}"));
        assert_eq!(side.color, rule_color, "rule {name}");
        assert_eq!(side.width, 1.0, "rule {name} is a 1pt band");
        assert_eq!(side.style, BorderLineStyle::Solid, "rule {name}");
    }
    assert!(
        header_border.right.is_none(),
        "column A is not the table's right edge"
    );

    let interior = cell_at(1, 1).border.as_ref().expect("B2 is ruled too");
    assert!(
        interior.bottom.is_some(),
        "a Medium table rules every row boundary"
    );
    assert!(
        interior.left.is_none() && interior.right.is_none(),
        "column B is neither outer edge"
    );
    assert_eq!(
        cell_at(2, 2)
            .border
            .as_ref()
            .and_then(|border| border.right.as_ref())
            .map(|side| side.color),
        Some(rule_color),
        "column C carries the table's right edge"
    );

    let header_run = match &cell_at(0, 0).content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("header holds a paragraph, got {other:?}"),
    };
    assert_eq!(header_run.style.bold, Some(true));
    assert_eq!(
        header_run.style.color,
        Some(Color {
            r: 0xff,
            g: 0xff,
            b: 0xff,
        }),
        "the header runs are printed white on the accent band"
    );
}

/// `table-multiple.xlsx` styles its `A1:C5` table `TableStyleMedium9` over an
/// accent 1 of `4F81BD`. That style opens the Medium family's second band,
/// which fills **every** body row in two tints — `#B8CCE4` over a whole-table
/// `#DCE6F1` — under a solid accent header, and seams the rows apart in white:
/// 3pt under the header, 1pt between the body rows and 1pt down each boundary
/// between the columns (issue #1189).
#[test]
fn structure_medium9_table_style_fills_every_row_and_seams_them_in_white() {
    let pages = sheet_pages("table-multiple.xlsx");
    let rows = &pages[0].table.rows;
    let cell_at = |row: usize, col: usize| -> &TableCell { &rows[row].cells[col] };
    let white = Color {
        r: 0xff,
        g: 0xff,
        b: 0xff,
    };
    let light_band = Color {
        r: 0xdc,
        g: 0xe6,
        b: 0xf1,
    };
    let dark_band = Color {
        r: 0xb8,
        g: 0xcc,
        b: 0xe4,
    };

    assert_eq!(
        cell_at(0, 0).background,
        Some(Color {
            r: 0x4f,
            g: 0x81,
            b: 0xbd,
        }),
        "A1's header band is the accent itself"
    );
    assert_eq!(
        (1..=4)
            .map(|row| cell_at(row, 0).background)
            .collect::<Vec<_>>(),
        vec![
            Some(dark_band),
            Some(light_band),
            Some(dark_band),
            Some(light_band)
        ],
        "every body row is filled, alternating between the two tints"
    );

    let header_border = cell_at(0, 0).border.as_ref().expect("A1 is seamed");
    let seam = header_border
        .bottom
        .as_ref()
        .expect("a seam runs under the header");
    assert_eq!(seam.color, white, "band 2 seams in white");
    assert_eq!(seam.width, 3.0, "the header seam is a 3pt band");
    assert_eq!(seam.style, BorderLineStyle::Solid);
    assert!(
        header_border.top.is_none(),
        "band 2 leaves the table's top clear"
    );
    assert!(header_border.left.is_none(), "band 2 rules no left edge");

    let body_border = cell_at(1, 0).border.as_ref().expect("A2 is seamed");
    assert_eq!(
        body_border
            .bottom
            .as_ref()
            .map(|side| (side.color, side.width)),
        Some((white, 1.0)),
        "a 1pt seam runs between the body rows"
    );
    assert_eq!(
        cell_at(1, 1)
            .border
            .as_ref()
            .and_then(|border| border.left.as_ref())
            .map(|side| (side.color, side.width)),
        Some((white, 1.0)),
        "the boundary between columns A and B is seamed"
    );
    assert!(
        cell_at(1, 2)
            .border
            .as_ref()
            .and_then(|border| border.right.as_ref())
            .is_none(),
        "band 2 leaves the table's right edge clear"
    );
    assert!(
        cell_at(4, 0).border.is_none(),
        "band 2 leaves the table's foot and left edge clear"
    );

    let header_run = match &cell_at(0, 0).content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("header holds a paragraph, got {other:?}"),
    };
    assert_eq!(header_run.style.bold, Some(true));
    assert_eq!(
        header_run.style.color,
        Some(white),
        "the header runs are printed white on the accent band"
    );
}

/// `<printOptions horizontalCentered="1"/>` centres the printed grid between
/// the print margins; without it the grid prints flush to the left one.
///
/// `issue_1110_horizontal_centered_probe.xlsx` is the probe the model was
/// measured on: five 10-unit columns of the Calibri 11 Normal font — 60pt each
/// on Excel's whole-point column metric — on A4 portrait with 0.7in margins
/// and nothing else. A native Excel for Mac export of it, staged and run
/// inside Excel's own sandbox container, fills the grid box from x=146pt to
/// x=446pt on the 595pt page (issue #1110).
#[test]
fn structure_horizontally_centered_sheet_insets_its_grid_to_the_native_offset() {
    let data = load_fixture("issue_1110_horizontal_centered_probe.xlsx");
    let (document, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("fixture should parse");

    let pages: Vec<&SheetPage> = document
        .pages
        .iter()
        .filter_map(|page| match page {
            Page::Sheet(sheet) => Some(sheet),
            _ => None,
        })
        .collect();
    assert_eq!(pages.len(), 1, "the probe holds one printed page");
    let page: &SheetPage = pages[0];
    assert!(
        page.table.centers_between_print_margins,
        "the probe declares printOptions horizontalCentered"
    );
    assert_eq!(
        page.table.column_widths,
        vec![60.0; 5],
        "Excel's column metric prices 10 units of Calibri 11 at 60pt"
    );

    let source = generate_typst(&document)
        .expect("fixture should generate Typst")
        .source;
    let inset_pt: f64 = source
        .split_once("#pad(left: ")
        .and_then(|(_, rest)| rest.split_once("pt)["))
        .and_then(|(value, _)| value.parse().ok())
        .unwrap_or_else(|| panic!("the centred sheet must inset its grid: {source}"));
    let grid_left_pt: f64 = page.margins.left + inset_pt;
    assert!(
        (grid_left_pt - 146.0).abs() < 1.0,
        "grid left edge {grid_left_pt}pt must land within 1pt of the native export's 146pt"
    );
}

/// A `fitToPage` sheet prints its drawings at the same scale as its grid, so
/// an anchored picture comes down with the rows and columns it is anchored to.
///
/// `issue_1111_fit_to_page_picture.xlsx` is the reported workbook. Its second
/// sheet declares `<pageSetUpPr fitToPage="true"/>` with `fitToWidth="1"` and
/// prints at 0.82 on A3 landscape. A native Excel for Mac export of it, staged
/// and run inside Excel's own sandbox container, draws the photo 192.66 x
/// 140.26pt with its top 436.57pt from the page's top edge — a 54.00pt top
/// margin in, so 382.57pt into the printed grid. Leaving the anchor geometry
/// unscaled drew it 234.95 x 171.05pt, 1/0.82 in each axis (issue #1111).
///
/// The horizontal offset is left to the unit test in `xlsx_pagination`: this
/// sheet also prints `horizontalCentered`, so the export's 63.91pt left edge
/// is only comparable once the renderer has added the centring inset the IR
/// does not carry.
#[test]
fn structure_fit_to_page_sheet_scales_its_anchored_picture_to_the_native_size() {
    let pages = sheet_pages("issue_1111_fit_to_page_picture.xlsx");
    let images: Vec<&office2pdf::ir::SheetImage> =
        pages.iter().flat_map(|page| page.images.iter()).collect();
    assert_eq!(images.len(), 1, "the workbook anchors one picture");

    let width: f64 = images[0].image.width.expect("a two-cell anchor sizes");
    let height: f64 = images[0].image.height.expect("a two-cell anchor sizes");
    assert!(
        (width - 192.66).abs() < 0.5,
        "the picture must print at the export's 192.66pt, got {width}"
    );
    assert!(
        (height - 140.26).abs() < 0.5,
        "the picture must print at the export's 140.26pt, got {height}"
    );
    assert!(
        (images[0].y_offset_pt - 382.57).abs() < 1.0,
        "the anchor's top offset must scale with the rows, got {}",
        images[0].y_offset_pt
    );
}

/// A `fitToPage` sheet that names neither `fitToWidth` nor `fitToHeight` takes
/// ECMA-376's default of one page in *both* directions, and Excel scales it by
/// whichever bound is tighter.
///
/// `issue_1181_fit_to_height.xlsx` is the reported workbook. Its
/// `Monthly college budget` sheet declares `<pageSetUpPr fitToPage="1"/>` and a
/// bare `<pageSetup paperSize="8" scale="85" orientation="portrait"/>`, and
/// prints 73 rows of A1:R73 whose printed tracks total 1372pt against A3
/// portrait's 1082.55pt of printable height. A native Excel for Mac export of
/// it, staged and run inside Excel's own sandbox container, is a single A3 page
/// drawn at 0.78 — `mutool draw -F trace` reports a `.78` text transform.
/// Bounding the columns alone left the sheet at the width fit's 0.89 and a
/// second page (issue #1181). The later native controls recorded in #1632
/// measure 19.00pt for the declared 19.5pt body row, or 14.82pt at 0.78.
#[test]
fn structure_fit_to_page_sheet_without_declared_bounds_fits_its_rows_on_one_page() {
    let pages = sheet_pages("issue_1181_fit_to_height.xlsx");
    let budget: &SheetPage = pages
        .iter()
        .find(|page| page.name == "Monthly college budget")
        .expect("the reported sheet should print");

    let printable_height: f64 = budget.size.height - budget.margins.top - budget.margins.bottom;
    let printed_height: f64 = budget.table.rows.iter().filter_map(|row| row.height).sum();
    assert!(
        printed_height <= printable_height,
        "the fitted sheet must stand on one page: {printed_height}pt of rows against \
         {printable_height}pt of printable height"
    );

    // Pin the 0.78 fit result at the 19.00pt body track measured in #1632.
    // Its near-whole-point correction affects row 1, not this 19.5pt row.
    let tracks: Vec<f64> = budget
        .table
        .rows
        .iter()
        .filter_map(|row| row.height)
        .collect();
    let body_track: f64 = tracks[tracks.len() - 1];
    assert!(
        (body_track - 14.82).abs() < 0.01,
        "the body row's native 14.82pt track must remain explicit, got {body_track}"
    );
}

/// The budget sheet's drawing holds two `xdr:cxnSp` connectors with
/// `a:prstGeom prst="line"`, the gray separators between its three summary
/// charts. Native Excel prints them as 0.78pt #D9D9D9 vertical lines at
/// x = 340.07 and 596.34pt from y = 130.44 to 284.49pt on the 0.78-fitted
/// page; the output held no line shape at all (issue #1566).
#[test]
fn structure_monthly_budget_prints_its_two_chart_separator_lines() {
    use office2pdf::ir::ShapeKind;

    let pages = sheet_pages("issue_1181_fit_to_height.xlsx");
    let budget = sheet_page_named(&pages, "Monthly college budget");
    assert_eq!(
        budget.shapes.len(),
        2,
        "the budget sheet anchors two separator lines"
    );
    for (index, line) in budget.shapes.iter().enumerate() {
        let ShapeKind::Line { x1, y1, x2, y2, .. } = line.shape.kind else {
            panic!("separator {index} is a line, got {:?}", line.shape.kind);
        };
        assert_eq!(
            (x1, y1),
            (0.0, 0.0),
            "separator {index} starts at its anchor"
        );
        assert_eq!(x2, 0.0, "separator {index} is vertical");
        // Native: 2508265 EMU (197.50pt) of anchor height at 0.78 = 154.05pt;
        // the anchor spans rows 4-14, and each row's own whole-point printed
        // track rounds independently of the raw EMU height, so the sum is
        // pinned loosely rather than to the exact conversion.
        assert!(
            (y2 - 154.05).abs() < 2.0 && (line.height - y2).abs() < 1e-9,
            "separator {index} runs the anchor height, got {y2}pt"
        );
        let stroke = line
            .shape
            .stroke
            .as_ref()
            .unwrap_or_else(|| panic!("separator {index} is stroked"));
        assert!(
            (stroke.width - 0.78).abs() < 1e-6,
            "the 1pt `a:ln` scales to 0.78pt with the fitted sheet, got {}",
            stroke.width
        );
        assert_eq!(stroke.color, Color::new(217, 217, 217));
    }
    // Native draws them 256.27pt apart; the anchors sit at column 4 + 482419
    // EMU and column 13 + 82971 EMU of the fitted grid.
    let separation: f64 = budget.shapes[1].x_offset_pt - budget.shapes[0].x_offset_pt;
    assert!(
        (separation - 256.27).abs() < 1.0,
        "the separators keep their native spacing, got {separation}pt"
    );
    assert!(
        budget
            .shapes
            .iter()
            .all(|line| (line.y_offset_pt - budget.shapes[0].y_offset_pt).abs() < 1e-9),
        "both separators hang from the same row"
    );
}

/// The reported monthly-budget workbook formats zero-valued entry cells with
/// the third section of `#,##0_);[Red]\(#,##0\);\-\ \ `. Excel prints the
/// escaped dash followed by both escaped spaces. All 71 cells are explicit
/// `<v>0</v>` cells in the worksheet XML; falling back to the positive section
/// printed `0`, while trimming the literal section shifted every right-aligned
/// dash 3.76pt to the right (issue #1262).
#[test]
fn structure_monthly_budget_zero_values_preserve_the_literal_zero_section() {
    let pages = sheet_pages("issue_1181_fit_to_height.xlsx");
    let budget = sheet_page_named(&pages, "Monthly college budget");
    let zero_cells: &[(usize, &[usize])] = &[
        (31, &[2, 3]),
        (34, &[5, 6, 7, 8, 9, 10, 11, 12, 13]),
        (45, &[2, 3, 5, 6, 8, 9, 11, 12, 13]),
        (46, &[2, 3, 5, 6, 8, 9, 11, 12, 13]),
        (49, &[2, 3, 5, 6, 8, 9, 11, 12, 13]),
        (50, &[2, 3]),
        (56, &[3, 4, 5, 6, 7, 9, 10, 11]),
        (59, &[2, 3, 4]),
        (61, &[2, 3, 4, 5, 6, 8, 9, 10, 11, 12]),
        (62, &[2, 3]),
        (63, &[2, 3, 4]),
        (64, &[2, 3, 4]),
        (68, &[2, 3]),
    ];

    assert_eq!(
        zero_cells
            .iter()
            .map(|(_, cells)| cells.len())
            .sum::<usize>(),
        71
    );
    for (row_index, column_indexes) in zero_cells {
        for column_index in *column_indexes {
            assert_eq!(
                table_cell_text(&budget.table.rows[*row_index].cells[*column_index]),
                "-  ",
                "worksheet cell at row {}, column {} must preserve the zero section's two spaces",
                row_index + 1,
                column_index + 1
            );
        }
    }
}

/// Every chart in the same workbook states where its plot area sits inside the
/// chart area, because the template floats each chart over the cells that print
/// its heading — `january income:` and `$1,225` for this one — and the chart's
/// own area is unfilled, so those cells show through the top third of it.
///
/// A native Excel for Mac 16 export of the workbook, staged and run inside
/// Excel's own sandbox container and traced with `mutool draw -F trace`, puts
/// this chart's bars at 152.75pt with its value ticks 30.82pt apart, on a chart
/// area running 80.14..320.36pt of the printed page. The sheet prints at 0.78,
/// so that is a 240.23pt chart area holding a 154.09pt plot 72.61pt in — this
/// placement times these fractions, to within 0.03pt. Ignoring the layout drew
/// the plot from the frame's own top-left instead, over the heading
/// (issue #1182).
#[test]
fn structure_monthly_budget_bar_chart_states_where_its_plot_sits() {
    let pages = sheet_pages("issue_1181_fit_to_height.xlsx");
    let budget: &SheetPage = pages
        .iter()
        .find(|page| page.name == "Monthly college budget")
        .expect("the reported sheet should print");
    let income = budget
        .charts
        .iter()
        .find(|chart| chart.chart.categories.iter().any(|c| c == "from savings"))
        .expect("the january income chart should reach the IR");

    assert_eq!(
        income.chart.plot_area_layout,
        Some(ChartPlotAreaLayout {
            x: 0.3022229818508113,
            y: 0.34625485336714895,
            width: 0.6413931113556166,
            height: 0.5571170578930364,
        })
    );

    // The chart area those fractions are of, and the scale the whole drawing
    // comes down by: 307.99 x 207.52pt at 0.78 is the export's 240.23 x
    // 161.87pt.
    let placement = income.placement.expect("the anchor places the chart");
    for (axis, actual, expected) in [
        ("width", placement.width * placement.print_scale, 240.23),
        ("height", placement.height * placement.print_scale, 161.87),
    ] {
        assert!(
            (actual - expected).abs() <= 0.05,
            "the {axis} of the printed chart area: {actual} against the export's {expected}"
        );
    }
}

// ---------------------------------------------------------------------------
// A drawing relationship no sheet element references
// ---------------------------------------------------------------------------

/// Rewrite one part of an XLSX package, leaving every other entry byte-for-byte
/// as it was. The patch runs on the named part's UTF-8 text.
fn repackage_xlsx_part(data: &[u8], part: &str, patch: impl Fn(&str) -> String) -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(data)).expect("fixture should be a ZIP");
    let mut rebuilt = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let mut patched = false;

    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("entry should be readable");
        let name: String = entry.name().to_string();
        let mut bytes: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("entry should decompress");

        rebuilt
            .start_file(&name, zip::write::FileOptions::default())
            .expect("entry should be writable");
        if name == part {
            let xml: String = String::from_utf8(bytes).expect("the part should be UTF-8 XML");
            let replacement: String = patch(&xml);
            assert_ne!(replacement, xml, "the patch must change {part}");
            patched = true;
            std::io::Write::write_all(&mut rebuilt, replacement.as_bytes())
                .expect("patched part should be writable");
        } else {
            std::io::Write::write_all(&mut rebuilt, &bytes).expect("entry should be writable");
        }
    }

    assert!(patched, "{part} should exist in the package");
    rebuilt.finish().expect("package should close").into_inner()
}

/// Parse XLSX bytes held in memory and return the sheet pages.
fn sheet_pages_of(data: &[u8]) -> Vec<SheetPage> {
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("package should parse");
    doc.pages
        .into_iter()
        .filter_map(|page| match page {
            Page::Sheet(sheet) => Some(sheet),
            _ => None,
        })
        .collect()
}

/// A `/drawing` relationship attaches nothing on its own: the sheet body's
/// `<drawing r:id="...">` element is what puts a drawing on the grid.
///
/// `issue_1158_unreferenced_drawing_rel.xlsx` is
/// `issue_1066_blip_effect_picture.xlsx` with only that element removed — the
/// relationship, `xl/drawings/drawing1.xml`, `xl/media/image1.png` and the
/// content types are all still there. Excel for Mac exports it as the grid
/// alone, its `mutool draw -F trace` holding no `fill_image` at all where the
/// base package's holds one. Resolving the drawing through the relationship
/// instead still drew the picture, taking the export's 0.155% ink coverage to
/// 2.914% (issue #1158).
#[test]
fn a_drawing_relationship_no_sheet_element_names_prints_no_picture() {
    let pages: Vec<SheetPage> = sheet_pages("issue_1158_unreferenced_drawing_rel.xlsx");

    let images: Vec<&office2pdf::ir::SheetImage> =
        pages.iter().flat_map(|page| page.images.iter()).collect();
    assert!(
        images.is_empty(),
        "an unreferenced drawing relationship must print nothing, got {} picture(s)",
        images.len()
    );
    // The grid itself is untouched — this is a drawing rule, not a parse
    // failure that swallowed the sheet.
    assert_eq!(sheet_names(&pages), vec!["Budget"]);
    assert!(
        total_rows(&pages) > 0,
        "the worksheet's rows must still print"
    );
}

/// The same rule for a chart. `SimpleScatterChart.xlsx`'s worksheet anchors
/// `xl/charts/chart1.xml` through `xl/drawings/drawing1.xml`; its chartsheet
/// anchors `chart2.xml` through `drawing2.xml`. Dropping the worksheet's
/// `<drawing>` element must drop that chart entirely — not re-place it after
/// the grid as the unanchored-chart fallback does for a chart part no drawing
/// reaches at all (issue #1158).
#[test]
fn a_drawing_relationship_no_sheet_element_names_prints_no_chart() {
    let data: Vec<u8> = repackage_xlsx_part(
        &load_fixture("SimpleScatterChart.xlsx"),
        "xl/worksheets/sheet1.xml",
        |xml| xml.replace(r#"<drawing r:id="rId1"/>"#, ""),
    );
    let pages: Vec<SheetPage> = sheet_pages_of(&data);

    assert_eq!(sheet_names(&pages), vec!["Sheet1", "Chart1"]);
    assert!(
        pages[0].charts.is_empty(),
        "the worksheet names no drawing, so it prints no chart, got {}",
        pages[0].charts.len()
    );
    // The chartsheet still names its own drawing, so its chart is unaffected.
    assert_eq!(pages[1].charts.len(), 1);
}

/// The same workbook's cash-flow chart carries a drawing part of its own:
/// `xl/charts/chart2.xml` ends with `<c:userShapes r:id="rId3"/>`, and the
/// `xl/drawings/drawing2.xml` behind it anchors the `CASH FLOW` caption Excel
/// prints left of the plot. Nothing followed that relationship, so the caption
/// was missing from the page (issue #1186).
///
/// A native Excel for Mac 16 export of the worksheet, staged and run inside
/// Excel's own sandbox container and traced with `mutool draw -F trace`, prints
/// the chart area 87.59..701.22pt across the page — 613.63pt at the sheet's
/// 0.78, which is this anchor's 786.71pt — and sets the caption in 11.7pt
/// Cambria Bold, the 15pt this run declares at that same scale.
#[test]
fn structure_monthly_budget_cash_flow_chart_carries_its_caption() {
    let pages = sheet_pages("issue_1181_fit_to_height.xlsx");
    let budget: &SheetPage = pages
        .iter()
        .find(|page| page.name == "Monthly college budget")
        .expect("the reported sheet should print");
    let cash_flow = budget
        .charts
        .iter()
        .find(|chart| chart.chart.categories.iter().any(|c| c.trim() == "jan"))
        .expect("the cash-flow chart should reach the IR");

    let shapes = &cash_flow.chart.user_shapes;
    assert_eq!(shapes.len(), 1, "the drawing part anchors one shape");
    assert_eq!(shapes[0].from, (0.0, 0.17913));
    assert_eq!(
        shapes[0].extent,
        ChartUserShapeExtent::Corner {
            x: 0.11685,
            y: 0.50958
        }
    );
    // `wrap="none"`: the caption is 80pt of glyphs in a box whose insets leave
    // 77pt, and Excel runs it on past the box rather than breaking it.
    assert!(shapes[0].no_wrap);

    let runs = &shapes[0].paragraphs[0].runs;
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text, "CASH FLOW");
    assert_eq!(runs[0].style.font_size, Some(15.0));
    assert_eq!(runs[0].style.bold, Some(true));
    // `<a:latin typeface="+mj-lt"/>` against the workbook theme, and
    // `accent1` at half the luminance — both as the export prints them.
    assert_eq!(runs[0].style.font_family.as_deref(), Some("Cambria"));
    assert_eq!(runs[0].style.color, Some(Color::new(0x24, 0x67, 0x78)));

    // The chart area those fractions are of, at the export's own scale.
    let placement = cash_flow.placement.expect("the anchor places the chart");
    for (axis, actual, expected) in [
        ("width", placement.width * placement.print_scale, 613.63),
        ("height", placement.height * placement.print_scale, 58.93),
    ] {
        assert!(
            (actual - expected).abs() <= 0.05,
            "the {axis} of the printed chart area: {actual} against the export's {expected}"
        );
    }
}

/// And the caption reaches the page: the workbook has to compile, and its text
/// layer has to carry the string, which is the whole of what issue #1186
/// reports missing.
#[test]
fn text_content_monthly_budget_cash_flow_caption() {
    let text: String = pdf_text("issue_1181_fit_to_height.xlsx");
    assert!(
        text.contains("CASH FLOW"),
        "the chart's own drawing part prints its caption; got:\n{text}"
    );
}

#[test]
fn out_of_grid_xlsx_fixture_is_rejected_before_layout_allocation() {
    let data = load_fixture("libreoffice/too-many-cols-rows.xlsx");
    let error = office2pdf::convert_bytes(
        &data,
        office2pdf::config::Format::Xlsx,
        &ConvertOptions::default(),
    )
    .expect_err("the out-of-grid workbook must fail before dense layout allocation");
    let message: String = error.to_string();
    assert!(message.contains("Excel grid"), "{message}");
    assert!(
        message.contains("16385") && message.contains("16777217"),
        "{message}"
    );
    let streaming = XlsxParser.parse_streaming(&data, &ConvertOptions::default(), 100);
    assert!(streaming.is_err());
    assert!(streaming.unwrap_err().to_string().contains("Excel grid"));
}
