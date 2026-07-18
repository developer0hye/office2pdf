use super::*;
use crate::ir::{Block, Margins, PageSize, Paragraph, ParagraphStyle, Run, TextStyle};

fn cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 1,
        row_span: 1,
        border: None,
        background: None,
        data_bar: None,
        icon_text: None,
        icon_color: None,
        spill_width: None,
        vertical_align: None,
        padding: None,
    }
}

fn cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(
                paragraph
                    .runs
                    .iter()
                    .map(|r| r.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

fn make_page(column_widths: Vec<f64>, rows: Vec<TableRow>) -> SheetPage {
    SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize {
            width: 500.0,
            height: 800.0,
        },
        margins: Margins {
            top: 50.0,
            bottom: 50.0,
            left: 50.0,
            right: 50.0,
        },
        table: Table {
            rows,
            column_widths,
            header_row_count: 0,
            alignment: None,
            default_cell_padding: None,
            use_content_driven_row_heights: false,
            default_vertical_align: None,
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
    }
}

#[test]
fn test_narrow_sheet_stays_single_page() {
    // Printable width 400pt; two 150pt columns fit.
    let page = make_page(
        vec![150.0, 150.0],
        vec![TableRow {
            cells: vec![cell("A"), cell("B")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None);
    assert_eq!(pages.len(), 1);
}

#[test]
fn test_wide_sheet_splits_into_column_groups() {
    // Printable width 400pt; five 150pt columns -> groups of 2/2/1.
    let page = make_page(
        vec![150.0; 5],
        vec![TableRow {
            cells: vec![cell("A"), cell("B"), cell("C"), cell("D"), cell("E")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None);
    assert_eq!(pages.len(), 3);
    assert_eq!(pages[0].table.column_widths.len(), 2);
    assert_eq!(pages[1].table.column_widths.len(), 2);
    assert_eq!(pages[2].table.column_widths.len(), 1);
    assert_eq!(cell_text(&pages[0].table.rows[0].cells[0]), "A");
    assert_eq!(cell_text(&pages[1].table.rows[0].cells[0]), "C");
    assert_eq!(cell_text(&pages[2].table.rows[0].cells[0]), "E");
}

#[test]
fn test_merge_straddling_boundary_truncates_and_blanks_continuation() {
    // Columns 0-1 on page 1, columns 2-3 on page 2. The merged cell spans
    // columns 1-2, so page 1 shows its content truncated to one column and
    // page 2 shows a blank continuation cell.
    let merged = TableCell {
        col_span: 2,
        ..cell("MERGED")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![TableRow {
            cells: vec![cell("A"), merged, cell("D")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None);
    assert_eq!(pages.len(), 2);

    let first_row = &pages[0].table.rows[0];
    assert_eq!(first_row.cells.len(), 2);
    assert_eq!(cell_text(&first_row.cells[1]), "MERGED");
    assert_eq!(first_row.cells[1].col_span, 1u32);

    let second_row = &pages[1].table.rows[0];
    assert_eq!(second_row.cells.len(), 2);
    assert_eq!(cell_text(&second_row.cells[0]), "");
    assert_eq!(cell_text(&second_row.cells[1]), "D");
}

#[test]
fn test_charts_stay_on_first_column_group() {
    let mut page = make_page(
        vec![300.0, 300.0],
        vec![TableRow {
            cells: vec![cell("A"), cell("B")],
            height: None,
        }],
    );
    page.charts = vec![(
        1,
        crate::ir::Chart {
            chart_type: crate::ir::ChartType::Bar,
            title: None,
            categories: vec![],
            series: vec![],
        },
    )];
    let pages = split_sheet_page_by_width(page, None);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].charts.len(), 1);
    assert!(pages[1].charts.is_empty());
}

#[test]
fn test_pathologically_wide_sheet_is_capped() {
    // 100 columns x 150pt with 400pt printable would be 50 pages; the cap
    // keeps the tail on the last page instead of exploding the compiler.
    let cells: Vec<TableCell> = (0..100).map(|i| cell(&format!("c{i}"))).collect();
    let page = make_page(
        vec![150.0; 100],
        vec![TableRow {
            cells,
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None);
    assert_eq!(pages.len(), 12);
    let total_columns: usize = pages.iter().map(|p| p.table.column_widths.len()).sum();
    assert_eq!(total_columns, 100);
}
