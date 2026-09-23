use super::*;
use crate::ir::ChartAreaOutline;
use crate::ir::{
    AxisTickMark, Block, BorderLineStyle, BorderSide, CellBorder, Color, HFInline, HeaderFooter,
    HeaderFooterParagraph, LineJoin, Margins, PageSize, Paragraph, ParagraphStyle, Run,
    TableBorderPaintModel, TextStyle,
};

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
        background_alpha: None,
        data_bar: None,
        sparkline: None,
        icon_text: None,
        icon_color: None,
        icon_shading: None,
        spill_width: None,
        spill_continuation_offset_pt: None,
        spill_line_extent: None,
        vertical_align: None,
        padding: None,
        row_has_thick_bottom: false,
        wraps_text: false,
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

/// A sheet bounded in the column direction only — the shape every audited
/// `fitToWidth` workbook declares, with `fitToHeight="0"` leaving the rows
/// free.
fn fit_to_width(pages_wide: u32) -> SheetFit {
    SheetFit {
        pages_wide: Some(pages_wide),
        ..SheetFit::default()
    }
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
            non_repeating_header_row_count: 0,
            alignment: None,
            default_cell_padding: None,
            use_content_driven_row_heights: false,
            default_vertical_align: None,
            seats_bottom_aligned_text_on_descender: false,
            bottom_aligned_descent_floor_pt: 0.0,
            border_paint_model: TableBorderPaintModel::CenteredStroke,
            prints_gridlines: false,
            prints_headings: false,
            centers_between_print_margins: false,
            print_scale: None,
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    }
}

fn bar_chart() -> crate::ir::Chart {
    crate::ir::Chart {
        chart_type: crate::ir::ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec![],
        series: vec![],
        grouping: crate::ir::ChartGrouping::Clustered,
        legend_position: crate::ir::LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: crate::ir::BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn test_narrow_sheet_stays_single_page() {
    // Printable width 400pt; two 150pt columns fit.
    let page = make_page(
        vec![150.0, 150.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), cell("B")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 1);
}

#[test]
fn test_wide_sheet_splits_into_column_groups() {
    // Printable width 400pt; five 150pt columns -> groups of 2/2/1.
    let page = make_page(
        vec![150.0; 5],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), cell("B"), cell("C"), cell("D"), cell("E")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
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
            minimum_height: None,
            cells: vec![cell("A"), merged, cell("D")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
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
fn test_merge_spill_width_is_clamped_to_the_column_group() {
    // A merged cell carries the width of the whole merge as its spill width, so
    // its text paints one unwrapped line that far. Slicing the merge at a page
    // break has to clamp that too, or the line keeps the full-sheet width and
    // paints past the printable edge — off the paper entirely on a wide sheet
    // (#631).
    let merged = TableCell {
        col_span: 4,
        spill_width: Some(600.0),
        ..cell("MERGED")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![merged],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);

    // Page 1 keeps two of the four merged columns, so the line may run 300pt.
    assert_eq!(pages[0].table.rows[0].cells[0].spill_width, Some(300.0));
    // The continuation carries no content, so it claims no spill either.
    assert_eq!(pages[1].table.rows[0].cells[0].spill_width, None);
}

#[test]
fn test_unmerged_spill_width_is_clamped_to_the_remaining_group_width() {
    // A general-aligned cell spills right into empty neighbours. When the group
    // ends before those neighbours do, the spill has to stop at the group edge.
    let spilling = TableCell {
        spill_width: Some(600.0),
        ..cell("LONG")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![spilling, cell("B"), cell("C"), cell("D")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);

    // The cell sits at the group's left edge, so 300pt of the group remain.
    assert_eq!(pages[0].table.rows[0].cells[0].spill_width, Some(300.0));
}

#[test]
fn test_spill_crossing_the_group_boundary_continues_on_the_next_page_column() {
    // Printable width 400pt: columns 0-1 print on page 1, columns 2-3 on
    // page 2. The cell in column 1 spills 400pt rightwards across its empty
    // neighbours, so 250pt of its line lie past the boundary. Excel prints
    // that tail on the next page-column, shifted left by the 150pt already
    // printed and clipped at the page-column's edge (issue #1381).
    let spilling = TableCell {
        spill_width: Some(400.0),
        ..cell("Corporate Security Supervisor")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), spilling, cell(""), cell("")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);

    let first_page_cell = &pages[0].table.rows[0].cells[1];
    assert_eq!(first_page_cell.spill_width, Some(150.0));
    assert_eq!(first_page_cell.spill_continuation_offset_pt, None);

    let continuation = &pages[1].table.rows[0].cells[0];
    assert_eq!(cell_text(continuation), "Corporate Security Supervisor");
    assert_eq!(continuation.spill_continuation_offset_pt, Some(150.0));
    assert_eq!(continuation.spill_width, Some(250.0));
    assert_eq!(
        cell_text(&pages[1].table.rows[0].cells[1]),
        "",
        "the continuation occupies only the page-column's first cell"
    );
}

#[test]
fn test_spill_ending_before_the_group_boundary_leaves_the_next_page_column_blank() {
    // The same grid, but the line stops 50pt short of the boundary: nothing
    // reaches page 2, so its first cell stays empty and carries no offset.
    // A second row paints in column 3 so the page-column exists at all
    // (issue #1714 drops one that paints nothing).
    let spilling = TableCell {
        spill_width: Some(100.0),
        ..cell("Short")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![
            TableRow {
                minimum_height: None,
                cells: vec![cell("A"), spilling, cell(""), cell("")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![cell(""), cell(""), cell(""), cell("D")],
                height: None,
            },
        ],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    let next_page_cell = &pages[1].table.rows[0].cells[0];
    assert_eq!(cell_text(next_page_cell), "");
    assert_eq!(next_page_cell.spill_continuation_offset_pt, None);
    assert_eq!(next_page_cell.spill_width, None);
}

#[test]
fn test_spill_spanning_three_page_columns_continues_on_each() {
    // Printable width 400pt over six 150pt columns: three page-columns of
    // two. An 800pt line from column 0 crosses both boundaries, so each later
    // page-column redraws it from its own offset — 300pt in, then 600pt in —
    // and the middle one clamps the remainder to the width it carries.
    let spilling = TableCell {
        spill_width: Some(800.0),
        ..cell("A very long unwrapped line")
    };
    let page = make_page(
        vec![150.0; 6],
        vec![TableRow {
            minimum_height: None,
            cells: vec![spilling, cell(""), cell(""), cell(""), cell(""), cell("")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 3);

    assert_eq!(pages[0].table.rows[0].cells[0].spill_width, Some(300.0));

    let second = &pages[1].table.rows[0].cells[0];
    assert_eq!(cell_text(second), "A very long unwrapped line");
    assert_eq!(second.spill_continuation_offset_pt, Some(300.0));
    assert_eq!(second.spill_width, Some(300.0));

    let third = &pages[2].table.rows[0].cells[0];
    assert_eq!(cell_text(third), "A very long unwrapped line");
    assert_eq!(third.spill_continuation_offset_pt, Some(600.0));
    assert_eq!(third.spill_width, Some(200.0));
}

#[test]
fn test_a_populated_cell_on_the_next_page_column_is_never_overwritten_by_a_continuation() {
    // A spill width that claims more than the empty run to its right cannot
    // paint over a cell that holds its own value: the next page-column keeps
    // that cell as it is.
    let spilling = TableCell {
        spill_width: Some(600.0),
        ..cell("LONG")
    };
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![spilling, cell(""), cell("C"), cell("")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    let kept = &pages[1].table.rows[0].cells[0];
    assert_eq!(cell_text(kept), "C");
    assert_eq!(kept.spill_continuation_offset_pt, None);
}

fn placed_chart(x_offset_pt: f64, width: f64) -> crate::ir::SheetChart {
    crate::ir::SheetChart {
        anchor_row: 1,
        placement: Some(crate::ir::SheetChartPlacement {
            x_offset_pt,
            y_offset_pt: 20.0,
            width,
            height: 100.0,
            print_scale: 1.0,
            clip_window: None,
        }),
        chart: bar_chart(),
    }
}

fn chart_placements(page: &SheetPage) -> Vec<crate::ir::SheetChartPlacement> {
    page.charts
        .iter()
        .map(|chart| chart.placement.expect("every test chart is anchored"))
        .collect()
}

/// Two 300pt columns on a 400pt printable width: the page-columns are
/// `[0, 300)` and `[300, 600)` in sheet points.
fn two_page_column_sheet() -> SheetPage {
    make_page(
        vec![300.0, 300.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), cell("B")],
            height: None,
        }],
    )
}

#[test]
fn test_a_chart_crossing_the_page_column_boundary_continues_on_the_next_page_column() {
    let mut page = two_page_column_sheet();
    // 200..400pt straddles the 300pt boundary.
    page.charts = vec![placed_chart(200.0, 200.0)];
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);

    // Excel clips the chart at the printed sheet edge of the first tile and
    // paints the rest of it on the next one, shifted left by the width the
    // first tile already printed (issue #1598).
    let first: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[0]);
    assert_eq!(first.len(), 1);
    assert_eq!(first[0].x_offset_pt, 200.0);
    assert_eq!(
        first[0].clip_window,
        Some(crate::ir::SheetClipWindow {
            left_pt: 0.0,
            width_pt: 300.0,
        })
    );

    let second: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[1]);
    assert_eq!(second.len(), 1);
    assert_eq!(second[0].x_offset_pt, -100.0);
    assert_eq!(
        second[0].clip_window,
        Some(crate::ir::SheetClipWindow {
            left_pt: 0.0,
            width_pt: 300.0,
        })
    );
    // The translation touches nothing but the horizontal offset.
    assert_eq!(second[0].y_offset_pt, 20.0);
    assert_eq!(second[0].width, 200.0);
    assert_eq!(second[0].height, 100.0);
    assert_eq!(second[0].print_scale, 1.0);
}

#[test]
fn test_a_chart_inside_one_page_column_prints_only_there() {
    let mut page = two_page_column_sheet();
    // 0..200pt lies wholly inside the first tile; 350..450pt wholly inside
    // the second.
    page.charts = vec![placed_chart(0.0, 200.0), placed_chart(350.0, 100.0)];
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);

    let first: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[0]);
    assert_eq!(first.len(), 1);
    assert_eq!(first[0].x_offset_pt, 0.0);
    assert_eq!(first[0].width, 200.0);
    assert_eq!(
        first[0].clip_window.map(|window| window.width_pt),
        Some(300.0)
    );

    let second: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[1]);
    assert_eq!(second.len(), 1);
    assert_eq!(second[0].x_offset_pt, 50.0);
    assert_eq!(second[0].width, 100.0);
    assert_eq!(
        second[0].clip_window.map(|window| window.width_pt),
        Some(300.0)
    );
}

#[test]
fn test_a_chart_ending_exactly_at_the_boundary_does_not_continue() {
    let mut page = two_page_column_sheet();
    page.charts = vec![placed_chart(100.0, 200.0)];
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].charts.len(), 1);
    assert!(
        pages[1].charts.is_empty(),
        "a chart whose right edge sits on the boundary paints nothing past it"
    );
}

#[test]
fn test_an_unsplit_sheet_leaves_its_chart_unclipped() {
    let mut page = make_page(
        vec![100.0, 100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), cell("B")],
            height: None,
        }],
    );
    page.charts = vec![placed_chart(50.0, 100.0)];
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 1);
    let placements: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[0]);
    assert_eq!(placements[0].x_offset_pt, 50.0);
    assert_eq!(
        placements[0].clip_window, None,
        "a sheet that fits one page-column draws its chart exactly as before"
    );
}

#[test]
fn test_a_continued_chart_starts_after_the_repeated_title_columns() {
    let mut page = two_page_column_sheet();
    page.charts = vec![placed_chart(200.0, 200.0)];
    // Column A repeats on the overflow page-column, in front of column B.
    let pages = split_sheet_page_by_width(page, Some((0, 1)), SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[1].table.column_widths, vec![300.0, 300.0]);

    let second: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[1]);
    assert_eq!(second.len(), 1);
    // The tile's own content begins after the 300pt title column, so the
    // continuation is measured from there and its window excludes the
    // repeated titles.
    assert_eq!(second[0].x_offset_pt, 200.0);
    assert_eq!(
        second[0].clip_window,
        Some(crate::ir::SheetClipWindow {
            left_pt: 300.0,
            width_pt: 300.0,
        })
    );
}

#[test]
fn test_a_last_page_column_carrying_only_a_chart_continuation_is_printed() {
    let mut page = make_page(
        vec![300.0, 300.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("A"), TableCell::default()],
            height: None,
        }],
    );
    page.charts = vec![placed_chart(200.0, 200.0)];
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(
        pages.len(),
        2,
        "the chart continuation is ink on the second page-column, so Excel prints it"
    );
    assert_eq!(pages[1].charts.len(), 1);
}

#[test]
fn test_pathologically_wide_sheet_is_capped() {
    // 100 columns x 150pt with 400pt printable would be 50 pages; the cap
    // keeps the tail on the last page instead of exploding the compiler.
    let cells: Vec<TableCell> = (0..100).map(|i| cell(&format!("c{i}"))).collect();
    let page = make_page(
        vec![150.0; 100],
        vec![TableRow {
            minimum_height: None,
            cells,
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 12);
    let total_columns: usize = pages.iter().map(|p| p.table.column_widths.len()).sum();
    assert_eq!(total_columns, 100);
}

/// Excel's `fitToWidth` squeezes the sheet onto that many pages instead of
/// letting it spill sideways: 800pt of columns on a 400pt printable width
/// scales to half size and stays on one page.
#[test]
fn test_fit_to_width_scales_columns_onto_one_page() {
    let page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].table.column_widths, vec![200.0, 200.0]);
    assert_eq!(pages[0].table.rows[0].height, Some(10.0));
}

/// A chart is anchored to the sheet's columns and rows, so the fit-to-page
/// scale that shrinks those has to carry it along; leaving it full size slid
/// the fitted grid out from under it (issue #982).
#[test]
fn test_fit_to_width_scales_an_anchored_chart_with_the_grid() {
    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    page.charts = vec![crate::ir::SheetChart {
        anchor_row: 1,
        placement: Some(crate::ir::SheetChartPlacement {
            x_offset_pt: 100.0,
            y_offset_pt: 40.0,
            width: 600.0,
            height: 200.0,
            print_scale: 1.0,
            clip_window: None,
        }),
        chart: bar_chart(),
    }];

    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    let placement = pages[0].charts[0]
        .placement
        .expect("the chart keeps its placement");
    assert_eq!(placement.x_offset_pt, 50.0);
    assert_eq!(placement.y_offset_pt, 20.0);
    // The chart lays itself out in its full-size frame and is drawn shrunk,
    // so the box it occupies on the page is the frame times the scale.
    assert_eq!(placement.width * placement.print_scale, 300.0);
    assert_eq!(placement.height * placement.print_scale, 100.0);
}

/// Excel scales a printed sheet whole, drawings included, so the chart's own
/// text shrinks with its frame. Recording the scale beside the frame — rather
/// than baking it into the frame — lets the renderer draw the whole chart
/// shrunk; shrinking the frame alone printed the reported workbook's tick
/// labels and legend at the size the chart XML declares, about 22% larger than
/// Excel prints them (issue #1069).
#[test]
fn test_fit_to_width_records_the_print_scale_on_an_anchored_chart() {
    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    page.charts = vec![crate::ir::SheetChart {
        anchor_row: 1,
        placement: Some(crate::ir::SheetChartPlacement {
            x_offset_pt: 100.0,
            y_offset_pt: 40.0,
            width: 600.0,
            height: 200.0,
            print_scale: 1.0,
            clip_window: None,
        }),
        chart: bar_chart(),
    }];

    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    let placement = pages[0].charts[0]
        .placement
        .expect("the chart keeps its placement");
    assert_eq!(placement.print_scale, 0.5);
    assert_eq!(placement.width, 600.0);
    assert_eq!(placement.height, 200.0);
}

/// A picture is anchored to the sheet's columns and rows exactly as a chart
/// is, so the fit-to-page scale that shrinks them has to shrink it too. The
/// reported workbook prints at 0.82 and its Excel for Mac export draws the
/// photo 192.66 x 140.26pt with its top 436.57pt down the page; leaving the
/// anchor geometry unscaled drew it 234.95 x 171.05pt — 1/0.82 in both axes —
/// with its top 84.83pt lower (issue #1111).
///
/// Unlike a chart, a picture has no text of its own to size, so the scale
/// belongs in the frame rather than beside it.
#[test]
fn test_fit_to_width_scales_an_anchored_picture_with_the_grid() {
    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    let mut picture = sheet_image(100.0, 600.0);
    picture.y_offset_pt = 40.0;
    picture.image.height = Some(200.0);
    page.images = vec![picture];

    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    let picture = &pages[0].images[0];
    assert_eq!(picture.x_offset_pt, 50.0);
    assert_eq!(picture.y_offset_pt, 20.0);
    assert_eq!(picture.image.width, Some(300.0));
    assert_eq!(picture.image.height, Some(100.0));
}

/// A worksheet line shape is anchored to the same columns and rows a picture
/// is, so the fit scale shrinks its offsets, its extent and its stroke alike:
/// the budget workbook's `a:ln w="12700"` separators trace as 0.78pt lines
/// on the 0.78-fitted native page (issue #1566).
#[test]
fn test_fit_to_width_scales_an_anchored_line_shape_with_the_grid() {
    use crate::ir::{ArrowHead, Shape, ShapeKind};

    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    page.shapes = vec![crate::ir::SheetShape {
        anchor_row: 3,
        x_offset_pt: 100.0,
        y_offset_pt: 40.0,
        width: 0.0,
        height: 200.0,
        shape: Shape {
            kind: ShapeKind::Line {
                x1: 0.0,
                y1: 0.0,
                x2: 0.0,
                y2: 200.0,
                head_end: ArrowHead::None,
                tail_end: ArrowHead::None,
            },
            fill: None,
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 1.0,
                color: Color::new(217, 217, 217),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
        },
    }];

    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    let line = &pages[0].shapes[0];
    assert_eq!(line.x_offset_pt, 50.0);
    assert_eq!(line.y_offset_pt, 20.0);
    assert_eq!(line.height, 100.0);
    let ShapeKind::Line { y2, .. } = line.shape.kind else {
        panic!("the line keeps its geometry: {:?}", line.shape.kind);
    };
    assert_eq!(y2, 100.0);
    let stroke = line
        .shape
        .stroke
        .as_ref()
        .expect("the line keeps its stroke");
    assert_eq!(stroke.width, 0.5);
    assert_eq!(stroke.color, Color::new(217, 217, 217));
}

/// Excel's auto-fit scale is a whole percent, truncated so the content is
/// guaranteed to fit. 400/530 = 75.47% must land on 75%, not 75.47% —
/// otherwise every derived type size is off by a fraction of a point.
#[test]
fn test_fit_to_width_truncates_scale_to_whole_percent() {
    let mut row = TableRow {
        minimum_height: None,
        cells: vec![cell("wide")],
        height: None,
    };
    if let Block::Paragraph(paragraph) = &mut row.cells[0].content[0] {
        paragraph.runs[0].style.font_size = Some(10.0);
    }
    let page = make_page(vec![530.0], vec![row]);
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].table.column_widths, vec![397.5]);
    let Block::Paragraph(paragraph) = &pages[0].table.rows[0].cells[0].content[0] else {
        panic!("expected a paragraph");
    };
    assert_eq!(paragraph.runs[0].style.font_size, Some(7.5));
}

/// `fitToWidth` never enlarges: a sheet that already fits keeps its metrics.
#[test]
fn test_fit_to_width_does_not_upscale_a_sheet_that_already_fits() {
    let page = make_page(
        vec![100.0, 100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("a"), cell("b")],
            height: Some(20.0),
        }],
    );
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].table.column_widths, vec![100.0, 100.0]);
    assert_eq!(pages[0].table.rows[0].height, Some(20.0));
}

/// A sheet with `fitToWidth` spanning two pages scales only enough to fill
/// both, then splits — it does not squeeze onto a single page.
#[test]
fn test_fit_to_width_two_pages_scales_then_splits() {
    let page = make_page(
        vec![400.0, 400.0, 400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("a"), cell("b"), cell("c"), cell("d")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, None, fit_to_width(2), true);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].table.column_widths, vec![200.0, 200.0]);
    assert_eq!(pages[1].table.column_widths, vec![200.0, 200.0]);
}

/// `fitToHeight` bounds the row direction the way `fitToWidth` bounds the
/// columns: 1400pt of sheet on a 700pt printable height scales to half size
/// and stays on one page. The page's own rows do not supply that total — it
/// is the whole sheet's, measured before pagination.
#[test]
fn test_fit_to_height_scales_rows_onto_one_page() {
    let page = make_page(
        vec![100.0, 100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    let fit = SheetFit {
        pages_tall: Some(1),
        sheet_height_pt: 1400.0,
        ..SheetFit::default()
    };
    let pages = split_sheet_page_by_width(page, None, fit, true);
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].table.rows[0].height, Some(10.0));
    assert_eq!(pages[0].table.column_widths, vec![50.0, 50.0]);
}

/// Excel scales a fitted sheet by the tighter of its two bounds, not by
/// whichever one it reads first. The reported college-budget workbook fits
/// A3's width at 0.89 and its height at 0.78, and Excel prints it at 0.78
/// (issue #1181).
#[test]
fn test_fit_to_page_takes_the_tighter_of_the_two_bounds() {
    let tighter_in_rows = SheetFit {
        pages_wide: Some(1),
        pages_tall: Some(1),
        // 400pt of columns on a 400pt printable width needs no shrink;
        // 1400pt of rows on a 700pt printable height needs half.
        sheet_height_pt: 1400.0,
    };
    let pages = split_sheet_page_by_width(
        make_page(
            vec![200.0, 200.0],
            vec![TableRow {
                minimum_height: None,
                cells: vec![cell("left"), cell("right")],
                height: Some(20.0),
            }],
        ),
        None,
        tighter_in_rows,
        true,
    );
    assert_eq!(pages[0].table.column_widths, vec![100.0, 100.0]);
    assert_eq!(pages[0].table.rows[0].height, Some(10.0));

    let tighter_in_columns = SheetFit {
        pages_wide: Some(1),
        pages_tall: Some(1),
        // 800pt of columns on 400pt needs half; 875pt of rows on 700pt
        // needs only 0.80.
        sheet_height_pt: 875.0,
    };
    let pages = split_sheet_page_by_width(
        make_page(
            vec![400.0, 400.0],
            vec![TableRow {
                minimum_height: None,
                cells: vec![cell("left"), cell("right")],
                height: Some(20.0),
            }],
        ),
        None,
        tighter_in_columns,
        true,
    );
    assert_eq!(pages[0].table.column_widths, vec![200.0, 200.0]);
    assert_eq!(pages[0].table.rows[0].height, Some(10.0));
}

/// `fitToHeight="0"` is Excel's "as many pages tall as it takes", so a sheet
/// far past the printable height keeps its metrics and spills. This is the
/// shape every audited `fitToWidth` workbook declares, and the fix for
/// issue #1181 must not start shrinking them.
#[test]
fn test_an_unbounded_fit_to_height_leaves_the_rows_alone() {
    let page = make_page(
        vec![100.0, 100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    let unbounded = SheetFit {
        pages_wide: Some(1),
        pages_tall: None,
        sheet_height_pt: 1400.0,
    };
    let pages = split_sheet_page_by_width(page, None, unbounded, true);
    assert_eq!(pages[0].table.rows[0].height, Some(20.0));
    assert_eq!(pages[0].table.column_widths, vec![100.0, 100.0]);
}

/// The row bound truncates to a whole percent for the same reason the column
/// bound does: 700/900 = 77.78% has to land on 77%, so the content is
/// guaranteed to fit rather than to overshoot by a fraction of a point.
#[test]
fn test_fit_to_height_truncates_scale_to_whole_percent() {
    let page = make_page(
        vec![100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("only")],
            height: Some(100.0),
        }],
    );
    let fit = SheetFit {
        pages_tall: Some(1),
        sheet_height_pt: 900.0,
        ..SheetFit::default()
    };
    let pages = split_sheet_page_by_width(page, None, fit, true);
    assert_eq!(pages[0].table.rows[0].height, Some(77.0));
}

/// `fitToHeight="2"` asks for two pages tall, so a sheet twice the printable
/// height already fits and is left alone.
#[test]
fn test_fit_to_height_counts_the_pages_it_is_given() {
    let page = make_page(
        vec![100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("only")],
            height: Some(40.0),
        }],
    );
    let fit = SheetFit {
        pages_tall: Some(2),
        sheet_height_pt: 1400.0,
        ..SheetFit::default()
    };
    let pages = split_sheet_page_by_width(page, None, fit, true);
    assert_eq!(pages[0].table.rows[0].height, Some(40.0));
}

/// A sheet shorter than its printable height is never stretched to fill it.
#[test]
fn test_fit_to_height_does_not_upscale_a_short_sheet() {
    let page = make_page(
        vec![100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("only")],
            height: Some(20.0),
        }],
    );
    let fit = SheetFit {
        pages_tall: Some(1),
        sheet_height_pt: 20.0,
        ..SheetFit::default()
    };
    let pages = split_sheet_page_by_width(page, None, fit, true);
    assert_eq!(pages[0].table.rows[0].height, Some(20.0));
}

/// Cell padding is part of the sheet's printed metrics. Leaving it at full
/// size while the rows shrink adds a fixed overhead to every row, which
/// accumulates into whole extra pages over a long sheet.
#[test]
fn test_fit_to_width_scales_cell_padding() {
    let mut left = cell("left");
    left.padding = Some(crate::ir::Insets {
        top: 2.0,
        right: 4.0,
        bottom: 2.0,
        left: 4.0,
    });
    let page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![left, cell("right")],
            height: Some(20.0),
        }],
    );
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);
    let padding = pages[0].table.rows[0].cells[0]
        .padding
        .expect("padding survives the fit");
    assert_eq!(padding.top, 1.0);
    assert_eq!(padding.bottom, 1.0);
    assert_eq!(padding.left, 2.0);
    assert_eq!(padding.right, 2.0);
}

#[test]
fn test_width_split_repeats_the_heading_gutter_and_keeps_the_flag() {
    // Printable width 400pt. Gutter (23pt) + three 150pt data columns
    // overflow; the heading-adjusted title range (0,1) must repeat the gutter
    // — including its row numbers — on the overflow page, and the letter
    // cells must follow their own columns there (issue #623).
    let mut page = make_page(
        vec![23.0, 150.0, 150.0, 150.0],
        vec![
            TableRow {
                minimum_height: None,
                cells: vec![cell(""), cell("A"), cell("B"), cell("C")],
                height: Some(13.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![cell("1"), cell("a1"), cell("b1"), cell("c1")],
                height: Some(13.0),
            },
        ],
    );
    page.table.prints_headings = true;

    let pages = split_sheet_page_by_width(page, Some((0, 1)), SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    for split_page in &pages {
        assert!(
            split_page.table.prints_headings,
            "every column group keeps the heading flag"
        );
        assert_eq!(split_page.table.column_widths[0], 23.0);
        assert_eq!(cell_text(&split_page.table.rows[1].cells[0]), "1");
    }
    assert_eq!(cell_text(&pages[0].table.rows[0].cells[1]), "A");
    assert_eq!(cell_text(&pages[1].table.rows[0].cells[1]), "C");
    assert_eq!(cell_text(&pages[1].table.rows[1].cells[1]), "c1");
}

#[test]
fn test_first_group_packs_against_the_full_printable_width() {
    // Printable width 400pt. The title columns (here the 23pt heading
    // gutter) are physically part of the first group, so only overflow
    // groups — which get them prepended — reserve their width. Packing the
    // first group at 400-23 pushed a fitting 190pt column onto page 2
    // (issue #623 adversarial review, finding 3).
    let page = make_page(
        vec![23.0, 180.0, 190.0, 200.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell(""), cell("A"), cell("B"), cell("C")],
            height: None,
        }],
    );
    let pages = split_sheet_page_by_width(page, Some((0, 1)), SheetFit::default(), true);
    // 23+180+190 = 393 <= 400 fits page 1; the 200pt column overflows to
    // page 2 behind the repeated 23pt title column.
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].table.column_widths, vec![23.0, 180.0, 190.0]);
    assert_eq!(pages[1].table.column_widths, vec![23.0, 200.0]);
    assert_eq!(cell_text(&pages[1].table.rows[0].cells[1]), "C");
}

/// One footer paragraph carrying `runs`, in the shape `parse_hf_format_string`
/// builds for `&L…`.
fn footer_with_runs(runs: Vec<Run>) -> HeaderFooter {
    HeaderFooter {
        shapes: Vec::new(),
        paragraphs: vec![HeaderFooterParagraph {
            style: ParagraphStyle::default(),
            elements: runs.into_iter().map(HFInline::Run).collect(),
            border: None,
            border_space: None,
            sheet_section_is_rich: false,
            frame: None,
        }],
        distance_from_edge: None,
        sheet_print_scale: None,
    }
}

fn footer_run_sizes(page: &SheetPage) -> Vec<Option<f64>> {
    page.footer
        .as_ref()
        .expect("the page keeps its footer")
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.style.font_size),
            _ => None,
        })
        .collect()
}

/// `headerFooter/@scaleWithDoc` defaults to 1 (ECMA-376 §18.3.1.46), so Excel
/// shrinks the footer with the sheet. Leaving it at full size printed the Gantt
/// template's 8pt `&8` run at 8pt beside 5.85pt body text (issue #940).
#[test]
fn a_scaled_sheet_scales_its_footer_with_it() {
    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    page.footer = Some(footer_with_runs(vec![
        Run {
            text: "_x000D_".to_string(),
            // No `&<n>` before it, so it takes the renderer's default size.
            style: TextStyle::default(),
            href: None,
            footnote: None,
        },
        Run {
            text: " Sensitivity: Internal".to_string(),
            style: TextStyle {
                font_size: Some(8.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        },
    ]));
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    assert_eq!(pages.len(), 1);
    // 400 + 400 onto a 400pt printable width: scale 0.5.
    assert_eq!(
        pages[0]
            .footer
            .as_ref()
            .and_then(|footer| footer.sheet_print_scale),
        Some(0.5),
        "the footer layout keeps the fit factor for its coordinate box"
    );
    assert_eq!(
        footer_run_sizes(&pages[0]),
        vec![
            Some(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT * 0.5),
            Some(4.0)
        ],
        "both the sized run and the one taking the default shrink with the sheet"
    );
}

/// `scaleWithDoc="0"` leaves both the footer type and the coordinate box in
/// page space even though the worksheet grid still fits to the page (#1510).
#[test]
fn a_scaled_sheet_that_opts_out_leaves_its_footer_coordinate_box_alone() {
    let mut page = make_page(
        vec![400.0, 400.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("left"), cell("right")],
            height: Some(20.0),
        }],
    );
    page.footer = Some(footer_with_runs(vec![Run {
        text: "Footer".to_string(),
        style: TextStyle {
            font_size: Some(8.0),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    }]));

    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), false);

    assert_eq!(pages.len(), 1);
    assert_eq!(
        pages[0]
            .footer
            .as_ref()
            .and_then(|footer| footer.sheet_print_scale),
        None
    );
    assert_eq!(footer_run_sizes(&pages[0]), vec![Some(8.0)]);
}

/// A sheet that already fits is never scaled up, so its footer is untouched
/// (issue #940).
#[test]
fn an_unscaled_sheet_leaves_its_footer_alone() {
    let mut page = make_page(
        vec![100.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell("narrow")],
            height: Some(20.0),
        }],
    );
    page.footer = Some(footer_with_runs(vec![Run {
        text: "Footer".to_string(),
        style: TextStyle {
            font_size: Some(8.0),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    }]));
    let pages = split_sheet_page_by_width(page, None, fit_to_width(1), true);

    assert_eq!(footer_run_sizes(&pages[0]), vec![Some(8.0)]);
}

// ── Drawing-width pagination on empty sheets (issue #713) ────────────

fn sheet_image(x: f64, width: f64) -> crate::ir::SheetImage {
    crate::ir::SheetImage {
        anchor_row: 1,
        x_offset_pt: x,
        y_offset_pt: 0.0,
        image: crate::ir::ImageData {
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
            rotation_deg: None,
            width: Some(width),
            height: Some(80.0),
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
            flip_h: false,
            flip_v: false,
        },
        clip_width_pt: None,
    }
}

/// Excel prints a drawing that crosses the printable edge clipped there and
/// continued on the next page-column; a drawing-only sheet must split the
/// same way instead of overflowing the right margin (issue #713).
#[test]
fn a_drawing_past_the_printable_edge_adds_a_page_column() {
    // Printable width is 400pt (500 − 2×50).
    let mut page = make_page(Vec::new(), Vec::new());
    page.images.push(sheet_image(10.0, 100.0));
    page.images.push(sheet_image(350.0, 100.0));

    let pages = split_drawing_only_page(page);
    assert_eq!(pages.len(), 2);

    assert_eq!(pages[0].images.len(), 2);
    assert_eq!(pages[0].images[0].x_offset_pt, 10.0);
    assert_eq!(pages[0].images[1].x_offset_pt, 350.0);
    assert_eq!(pages[0].images[1].clip_width_pt, Some(400.0));

    // The crossing image continues on the second page-column, shifted left
    // by one printable width and clipped to the same window.
    assert_eq!(pages[1].images.len(), 1);
    assert_eq!(pages[1].images[0].x_offset_pt, -50.0);
    assert_eq!(pages[1].images[0].clip_width_pt, Some(400.0));
}

#[test]
fn a_chart_crossing_a_drawing_only_page_column_continues_on_the_next() {
    let mut page = make_page(Vec::new(), Vec::new());
    page.images.push(sheet_image(350.0, 100.0));
    // 350..450pt crosses the 400pt printable edge.
    page.charts = vec![placed_chart(350.0, 100.0)];
    let pages = split_drawing_only_page(page);
    assert_eq!(pages.len(), 2);

    let first: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[0]);
    assert_eq!(first[0].x_offset_pt, 350.0);
    assert_eq!(
        first[0].clip_window,
        Some(crate::ir::SheetClipWindow {
            left_pt: 0.0,
            width_pt: 400.0,
        })
    );
    let second: Vec<crate::ir::SheetChartPlacement> = chart_placements(&pages[1]);
    assert_eq!(second[0].x_offset_pt, -50.0);
    assert_eq!(
        second[0].clip_window,
        Some(crate::ir::SheetClipWindow {
            left_pt: 0.0,
            width_pt: 400.0,
        })
    );
}

#[test]
fn drawings_inside_the_printable_width_keep_one_page() {
    let mut page = make_page(Vec::new(), Vec::new());
    page.images.push(sheet_image(10.0, 100.0));

    let pages = split_drawing_only_page(page);
    assert_eq!(pages.len(), 1);
    assert_eq!(
        pages[0].images[0].clip_width_pt, None,
        "an unsplit page draws its images exactly as before"
    );
}

/// A row whose only text is in column 0 and whose remaining cells are empty.
fn row_with_first_cell(text: &str, spill_width: Option<f64>, columns: usize) -> TableRow {
    let mut cells: Vec<TableCell> = vec![TableCell {
        spill_width,
        ..cell(text)
    }];
    cells.extend((1..columns).map(|_| cell("")));
    TableRow {
        minimum_height: None,
        cells,
        height: None,
    }
}

/// Excel ends the printed page sequence at its last page carrying ink. A
/// native Excel-for-Mac export of `100-customers.xlsx` with the last band's
/// occupation values shortened prints 5 pages, not 6: the down-page run of
/// three, then only the two strip pages that carry a spilled tail. The strip
/// page-column therefore stops at its last inked row (issue #1714).
#[test]
fn test_last_page_column_drops_its_trailing_rows_with_no_ink() {
    // Printable width 400pt: columns 0-1 print on page 1, columns 2-3 on the
    // strip. Rows 0 and 2 spill 400pt past their own column, rows 1, 3 and 4
    // end inside the first page-column.
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![
            row_with_first_cell("Corporate Security Supervisor", Some(400.0), 4),
            row_with_first_cell("Short", None, 4),
            row_with_first_cell("Principal Configuration Orchestrator", Some(400.0), 4),
            row_with_first_cell("Short", None, 4),
            row_with_first_cell("Short", None, 4),
        ],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    assert_eq!(
        pages[0].table.rows.len(),
        5,
        "the first page-column keeps every row of the used range"
    );
    assert_eq!(
        pages[1].table.rows.len(),
        3,
        "the strip ends at its last continued line; rows 3 and 4 carry nothing there"
    );
    assert_eq!(
        cell_text(&pages[1].table.rows[1].cells[0]),
        "",
        "a blank row between two continued lines is kept, so the bands stay aligned"
    );
    assert_eq!(
        cell_text(&pages[1].table.rows[2].cells[0]),
        "Principal Configuration Orchestrator"
    );
}

/// A trailing strip row whose cell paints a fill or a border is ink, so the
/// page-column keeps it.
#[test]
fn test_trailing_strip_rows_with_a_fill_or_border_are_ink() {
    for (label, painted) in [
        (
            "fill",
            TableCell {
                background: Some(Color {
                    r: 255,
                    g: 255,
                    b: 0,
                }),
                ..cell("")
            },
        ),
        (
            "border",
            TableCell {
                border: Some(CellBorder {
                    top: Some(BorderSide {
                        width: 0.5,
                        color: Color { r: 0, g: 0, b: 0 },
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                    }),
                    ..CellBorder::default()
                }),
                ..cell("")
            },
        ),
    ] {
        let mut last_row = row_with_first_cell("Short", None, 4);
        last_row.cells[3] = painted;
        let page = make_page(
            vec![150.0, 150.0, 150.0, 150.0],
            vec![
                row_with_first_cell("Corporate Security Supervisor", Some(400.0), 4),
                row_with_first_cell("Short", None, 4),
                last_row,
            ],
        );
        let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
        assert_eq!(pages.len(), 2);
        assert_eq!(
            pages[1].table.rows.len(),
            3,
            "a {label} on the last row keeps the whole strip"
        );
    }
}

/// Only the last page-column ends early. The same probe series shortened the
/// middle band's values instead and the native export still printed 6 pages,
/// the blank strip page for that band included: a blank page inside the
/// sequence is printed, so a page-column before the last one keeps every
/// row even when its tail carries nothing.
#[test]
fn test_a_page_column_before_the_last_keeps_its_trailing_rows() {
    // Printable width 400pt over six 150pt columns: three page-columns of
    // two. Row 0 spills 800pt so it continues on both later page-columns;
    // row 1 spills 400pt, so it reaches the second page-column only; rows 2
    // and 3 end on the first.
    let page = make_page(
        vec![150.0; 6],
        vec![
            row_with_first_cell("Reaches the third page-column", Some(800.0), 6),
            row_with_first_cell("Reaches the second page-column", Some(400.0), 6),
            row_with_first_cell("Short", None, 6),
            row_with_first_cell("Short", None, 6),
        ],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 3);
    assert_eq!(pages[0].table.rows.len(), 4);
    assert_eq!(
        pages[1].table.rows.len(),
        4,
        "the middle page-column keeps rows 2 and 3 although nothing continues there"
    );
    assert_eq!(
        pages[2].table.rows.len(),
        1,
        "the last page-column ends at row 0, its only continued line"
    );
}

/// Printed gridlines and headings paint every row of the range, and no probe
/// has measured whether Excel still ends the sequence early then; the
/// page-column keeps its rows.
#[test]
fn test_printed_gridlines_keep_the_trailing_strip_rows() {
    for (label, gridlines, headings) in [("gridlines", true, false), ("headings", false, true)] {
        let mut page = make_page(
            vec![150.0, 150.0, 150.0, 150.0],
            vec![
                row_with_first_cell("Corporate Security Supervisor", Some(400.0), 4),
                row_with_first_cell("Short", None, 4),
            ],
        );
        page.table.prints_gridlines = gridlines;
        page.table.prints_headings = headings;
        let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
        assert_eq!(pages.len(), 2);
        assert_eq!(
            pages[1].table.rows.len(),
            2,
            "printed {label} keep the trailing strip row"
        );
    }
}

/// A last page-column with no ink at all is not emitted.
#[test]
fn test_a_last_page_column_without_ink_is_not_emitted() {
    // Column 3 is empty everywhere and nothing spills past column 1, yet the
    // four 150pt columns still exceed the 400pt printable width.
    let page = make_page(
        vec![150.0, 150.0, 150.0, 150.0],
        vec![
            row_with_first_cell("Short", Some(100.0), 4),
            row_with_first_cell("Short", None, 4),
        ],
    );
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 1, "only the inked page-column prints");
    assert_eq!(pages[0].table.column_widths.len(), 2);
}

/// A line with the reach to cross the boundary but whose own width ends
/// before it carries nothing onto the next page-column. On
/// `1000-customers.xlsx` every occupation cell has a three-column reach
/// (E:G), yet row 1001's `Direct Creative Liaison` is 122pt long and ends
/// inside column F, so the native export prints no strip page for its band
/// while the longer lines of the rows above continue (issue #1714).
#[test]
fn test_a_line_ending_before_the_boundary_is_not_ink_despite_its_reach() {
    let short_line = TableCell {
        spill_width: Some(225.0),
        spill_line_extent: Some(SheetLineExtent::Estimated(124.6)),
        ..cell("Direct Creative Liaison")
    };
    let long_line = TableCell {
        spill_width: Some(225.0),
        spill_line_extent: Some(SheetLineExtent::Estimated(159.1)),
        ..cell("Corporate Data Orchestrator")
    };
    let page = make_page(
        vec![75.0, 75.0, 75.0],
        vec![
            TableRow {
                minimum_height: None,
                cells: vec![long_line, cell(""), cell("")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![short_line, cell(""), cell("")],
                height: None,
            },
        ],
    );
    let mut page = page;
    // Printable width 150pt: columns 0-1 print first, column 2 on the strip.
    page.size.width = 250.0;
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(pages.len(), 2);
    let continued = &pages[1].table.rows[0].cells[0];
    assert_eq!(cell_text(continued), "Corporate Data Orchestrator");
    assert_eq!(continued.spill_continuation_offset_pt, Some(150.0));
    assert_eq!(
        pages[1].table.rows.len(),
        1,
        "the short line's row paints nothing on the strip, so the strip ends above it"
    );

    // A line ending within the estimate's slack of the boundary still counts
    // as painting: the estimate runs under the face's advances.
    let near_line = TableCell {
        spill_width: Some(225.0),
        spill_line_extent: Some(SheetLineExtent::Estimated(146.0)),
        ..cell("Customer Group Developer")
    };
    let mut page = make_page(
        vec![75.0, 75.0, 75.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![near_line, cell(""), cell("")],
            height: None,
        }],
    );
    page.size.width = 250.0;
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
    assert_eq!(
        pages.len(),
        2,
        "a line 4pt short of the boundary keeps its strip"
    );
}

/// Excel decides a continuation on the whole-point advance grid it paces sheet
/// glyphs with, and a line that ends exactly on the page-column boundary stops
/// there. `Customer Group Developer` in
/// `tests/fixtures/xlsx/customers_overflow_strip.xlsx` is that case: its
/// advances sum to 147pt from a pen 3pt inside a gridline 150pt before the
/// boundary, and the native Excel for Mac export prints no tail for it while
/// printing one for `Legacy Solutions Developer`, 2pt longer (issue #1659).
#[test]
fn test_a_measured_line_is_continued_only_past_the_boundary() {
    // Printable width 150pt over three 75pt columns: columns 0-1 print first,
    // column 2 on the strip, so a line from column 0 crosses at 150pt.
    let measured = |text: &str, width_pt: f64| TableCell {
        spill_width: Some(225.0),
        spill_line_extent: Some(SheetLineExtent::Measured(width_pt)),
        ..cell(text)
    };
    let rows = vec![
        TableRow {
            minimum_height: None,
            cells: vec![
                measured("Legacy Solutions Developer", 152.0),
                cell(""),
                cell(""),
            ],
            height: None,
        },
        TableRow {
            minimum_height: None,
            cells: vec![
                measured("Customer Group Developer", 150.0),
                cell(""),
                cell(""),
            ],
            height: None,
        },
        TableRow {
            minimum_height: None,
            cells: vec![
                measured("Regional Intranet Associate", 148.0),
                cell(""),
                cell(""),
            ],
            height: None,
        },
        TableRow {
            minimum_height: None,
            cells: vec![
                measured("Corporate Data Orchestrator", 163.0),
                cell(""),
                cell(""),
            ],
            height: None,
        },
    ];
    let mut page = make_page(vec![75.0, 75.0, 75.0], rows);
    page.size.width = 250.0;
    let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);

    assert_eq!(pages.len(), 2);
    let tails: Vec<String> = pages[1]
        .table
        .rows
        .iter()
        .map(|row| cell_text(&row.cells[0]))
        .collect();
    assert_eq!(
        tails,
        vec![
            "Legacy Solutions Developer".to_string(),
            String::new(),
            String::new(),
            "Corporate Data Orchestrator".to_string(),
        ],
        "only the lines past the 150pt boundary are redrawn on the strip"
    );
}

/// The line the boundary rule reads is the whole-point extent, not the reach:
/// the same measured width continues or stops as the columns before the
/// boundary grow, with no slack in either direction (issue #1659).
#[test]
fn test_the_measured_boundary_moves_with_the_page_column() {
    for (columns_before, continues) in [(1usize, true), (2, false)] {
        let boundary_pt: f64 = 75.0 * columns_before as f64;
        let page = {
            let mut page = make_page(
                vec![75.0; 3],
                vec![TableRow {
                    minimum_height: None,
                    cells: vec![
                        TableCell {
                            spill_width: Some(225.0),
                            spill_line_extent: Some(SheetLineExtent::Measured(150.0)),
                            ..cell("Customer Group Developer")
                        },
                        cell(""),
                        cell(""),
                    ],
                    height: None,
                }],
            );
            page.size.width = 100.0 + boundary_pt;
            page
        };
        let pages = split_sheet_page_by_width(page, None, SheetFit::default(), true);
        if continues {
            assert_eq!(
                cell_text(&pages[1].table.rows[0].cells[0]),
                "Customer Group Developer",
                "a 150pt line crosses a {boundary_pt}pt page-column"
            );
        } else {
            assert_eq!(
                pages.len(),
                1,
                "a 150pt line ends on the {boundary_pt}pt boundary and leaves the strip blank"
            );
        }
    }
}
