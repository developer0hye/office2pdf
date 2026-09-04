use super::*;
use crate::ir::ChartAreaOutline;
use crate::ir::TableBorderPaintModel;
use crate::ir::{
    Block, BorderLineStyle, Margins, PageSize, Paragraph, ParagraphStyle, Run, SheetImage,
    SheetTextBox, Table, TableCell, TableRow, TextStyle,
};

fn text_cell(text: &str) -> TableCell {
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
        ..TableCell::default()
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
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

fn make_page(column_widths: Vec<f64>, rows: Vec<TableRow>) -> SheetPage {
    SheetPage {
        name: "Tests".to_string(),
        size: PageSize {
            width: 612.0,
            height: 792.0,
        },
        margins: Margins {
            top: 72.0,
            bottom: 72.0,
            left: 54.0,
            right: 54.0,
        },
        table: Table {
            rows,
            column_widths,
            border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
            prints_gridlines: true,
            seats_bottom_aligned_text_on_descender: true,
            default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
    }
}

#[test]
fn test_column_letters_follow_excels_bijective_base_26() {
    assert_eq!(column_letters(1), "A");
    assert_eq!(column_letters(5), "E");
    assert_eq!(column_letters(26), "Z");
    assert_eq!(column_letters(27), "AA");
    assert_eq!(column_letters(52), "AZ");
    assert_eq!(column_letters(53), "BA");
    assert_eq!(column_letters(702), "ZZ");
    assert_eq!(column_letters(703), "AAA");
}

#[test]
fn test_augment_prepends_gutter_column_and_letter_strip_at_measured_geometry() {
    let mut page = make_page(
        vec![120.0, 110.0],
        vec![
            TableRow {
                minimum_height: None,
                cells: vec![text_cell("x"), text_cell("y")],
                height: Some(13.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![text_cell("z"), text_cell("w")],
                height: Some(13.0),
            },
        ],
    );
    augment_page_with_print_headings(&mut page, &[1, 2], 1, None);

    let table = &page.table;
    assert!(table.prints_headings);
    // Gutter track: GT interior 22pt + the 1pt black separator band puts the
    // gutter|data boundary 23pt right of the margin.
    assert_eq!(table.column_widths, vec![23.0, 120.0, 110.0]);
    // Strip row: GT interior 12pt + 1pt separator band = 13pt track.
    assert_eq!(table.rows[0].height, Some(13.0));
    assert_eq!(table.rows.len(), 3);

    // Corner cell is empty; letters cover the printed columns.
    let strip = &table.rows[0];
    assert_eq!(strip.cells.len(), 3);
    assert!(strip.cells[0].content.is_empty());
    assert_eq!(cell_text(&strip.cells[1]), "A");
    assert_eq!(cell_text(&strip.cells[2]), "B");

    // Gutter cells carry the actual sheet row numbers.
    assert_eq!(cell_text(&table.rows[1].cells[0]), "1");
    assert_eq!(cell_text(&table.rows[2].cells[0]), "2");
    // Data cells stay in place, one column later.
    assert_eq!(cell_text(&table.rows[1].cells[1]), "x");
    assert_eq!(cell_text(&table.rows[2].cells[2]), "w");
}

#[test]
fn test_heading_rules_use_measured_colors() {
    let mut page = make_page(
        vec![120.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![text_cell("x")],
            height: Some(13.0),
        }],
    );
    augment_page_with_print_headings(&mut page, &[1], 1, None);

    let gray = PRINT_HEADING_RULE_GRAY;
    assert_eq!((gray.r, gray.g, gray.b), (86, 86, 86));

    // Corner: gray right edge (over the strip height) and gray bottom edge.
    let corner = page.table.rows[0].cells[0].border.as_ref().unwrap();
    let right = corner.right.as_ref().unwrap();
    assert_eq!(right.color, gray);
    assert_eq!(right.width, 1.0);
    assert_eq!(right.style, BorderLineStyle::Solid);
    assert_eq!(corner.bottom.as_ref().unwrap().color, gray);
    assert!(corner.top.is_none() && corner.left.is_none());

    // Letter cell: black separator under it, gray rule at its right boundary.
    let letter = page.table.rows[0].cells[1].border.as_ref().unwrap();
    assert_eq!(letter.bottom.as_ref().unwrap().color, Color::black());
    assert_eq!(letter.right.as_ref().unwrap().color, gray);

    // Gutter cell: black separator on the data side, gray rules at the row
    // boundaries above and below.
    let gutter = page.table.rows[1].cells[0].border.as_ref().unwrap();
    assert_eq!(gutter.right.as_ref().unwrap().color, Color::black());
    assert_eq!(gutter.right.as_ref().unwrap().width, 1.0);
    assert_eq!(gutter.top.as_ref().unwrap().color, gray);
    assert_eq!(gutter.bottom.as_ref().unwrap().color, gray);
}

#[test]
fn test_heading_cells_pin_the_descender_seat() {
    // GT bottom-seats heading text: the baseline sits 3pt above the band
    // bottom at every row height — the descent bottom rests ON the grid
    // boundary (issue #623 evidence, digits 96.0 vs boundary 98 on
    // nft-sheet-0002). The #711 seat conditions therefore require each
    // heading cell to state Bottom itself (not lean on a table default a
    // segment copy might drop) and a zero effective bottom inset: the
    // declared padding must cancel the half border width #500/#503 reserve.
    let mut page = make_page(
        vec![120.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![text_cell("x")],
            height: Some(12.0),
        }],
    );
    augment_page_with_print_headings(&mut page, &[1], 1, None);

    let letter = &page.table.rows[0].cells[1];
    let gutter = &page.table.rows[1].cells[0];
    for (label, cell) in [("letter", letter), ("gutter", gutter)] {
        assert_eq!(
            cell.vertical_align,
            Some(crate::ir::CellVerticalAlign::Bottom),
            "{label} cell must bottom-align explicitly"
        );
        let padding = cell.padding.expect("heading cells carry their padding");
        let bottom_border_half: f64 = cell
            .border
            .as_ref()
            .and_then(|border| border.bottom.as_ref())
            .map(|side| side.width / 2.0)
            .expect("heading cells declare a bottom rule");
        assert_eq!(
            padding.bottom + bottom_border_half,
            0.0,
            "{label} cell's effective bottom inset must be zero so the \
             descent bottom seats on the grid boundary"
        );
    }
}

#[test]
fn test_heading_text_uses_the_workbook_normal_font_and_center_alignment() {
    let mut page = make_page(
        vec![120.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![text_cell("x")],
            height: Some(13.0),
        }],
    );
    let verdana = NormalFont {
        family: "Verdana".to_string(),
        size_pt: 10.0,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };
    augment_page_with_print_headings(&mut page, &[337], 1, Some(&verdana));

    let gutter_cell = &page.table.rows[1].cells[0];
    let Block::Paragraph(paragraph) = &gutter_cell.content[0] else {
        panic!("gutter cell must hold a paragraph");
    };
    // GT: digits/letters are the Normal face at the Normal size, regular
    // weight, centred on the cell.
    assert_eq!(
        paragraph.style.alignment,
        Some(crate::ir::Alignment::Center)
    );
    let run = &paragraph.runs[0];
    assert_eq!(run.text, "337");
    assert_eq!(run.style.font_family.as_deref(), Some("Verdana"));
    assert_eq!(run.style.font_size, Some(10.0));
    assert_ne!(run.style.bold, Some(true));
}

#[test]
fn test_augment_uses_sheet_row_numbers_and_first_column_for_pagination_slices() {
    // A later chunk: sheet rows 49..51, printed columns starting at sheet
    // column 3 (print area C..D).
    let mut page = make_page(
        vec![120.0, 110.0],
        vec![
            TableRow {
                minimum_height: None,
                cells: vec![text_cell("a"), text_cell("b")],
                height: Some(13.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![text_cell("c"), text_cell("d")],
                height: Some(13.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![text_cell("e"), text_cell("f")],
                height: Some(13.0),
            },
        ],
    );
    augment_page_with_print_headings(&mut page, &[49, 50, 51], 3, None);

    assert_eq!(cell_text(&page.table.rows[1].cells[0]), "49");
    assert_eq!(cell_text(&page.table.rows[2].cells[0]), "50");
    assert_eq!(cell_text(&page.table.rows[3].cells[0]), "51");
    assert_eq!(cell_text(&page.table.rows[0].cells[1]), "C");
    assert_eq!(cell_text(&page.table.rows[0].cells[2]), "D");
}

#[test]
fn test_merged_cell_spans_stay_on_their_data_columns() {
    // A1:B1 merged: the cell spans 2 data columns. Prepending the gutter
    // shifts it one cell later but must not change its span.
    let merged = TableCell {
        col_span: 2,
        ..text_cell("merged")
    };
    let mut page = make_page(
        vec![120.0, 110.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![merged],
            height: Some(13.0),
        }],
    );
    augment_page_with_print_headings(&mut page, &[1], 1, None);

    let row = &page.table.rows[1];
    assert_eq!(cell_text(&row.cells[0]), "1");
    assert_eq!(row.cells[0].col_span, 1);
    assert_eq!(cell_text(&row.cells[1]), "merged");
    assert_eq!(row.cells[1].col_span, 2);
}

#[test]
fn test_drawings_shift_with_the_inset_grid() {
    // Drawing offsets are absolute worksheet coordinates measured from cell
    // A1's top-left corner; the headings move that corner right by the gutter
    // track and down by the strip track.
    let mut page = make_page(
        vec![120.0],
        vec![TableRow {
            minimum_height: None,
            cells: vec![text_cell("x")],
            height: Some(13.0),
        }],
    );
    page.images.push(SheetImage {
        anchor_row: 1,
        x_offset_pt: 10.0,
        y_offset_pt: 5.0,
        clip_width_pt: None,
        image: crate::ir::ImageData {
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
            width: Some(50.0),
            height: Some(50.0),
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
        },
    });
    page.text_boxes.push(SheetTextBox {
        anchor_row: 1,
        x_offset_pt: 0.0,
        y_offset_pt: 0.0,
        width: 40.0,
        height: 20.0,
        paragraphs: Vec::new(),
        fill: None,
        border: None,
        vertical_center: false,
    });
    page.charts.push(crate::ir::SheetChart {
        anchor_row: 2,
        placement: Some(crate::ir::SheetChartPlacement {
            x_offset_pt: 30.0,
            y_offset_pt: 60.0,
            width: 200.0,
            height: 100.0,
            print_scale: 1.0,
        }),
        chart: crate::ir::Chart {
            chart_type: crate::ir::ChartType::Bar,
            hole_size_percent: None,
            title: None,
            categories: vec![],
            series: vec![],
            grouping: crate::ir::ChartGrouping::Clustered,
            legend_position: crate::ir::LegendPosition::Right,
            has_legend: true,
            category_axis_title: None,
            value_axis_title: None,
            category_axis_major_tick_mark: crate::ir::AxisTickMark::Outside,
            value_axis_major_tick_mark: crate::ir::AxisTickMark::Outside,
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
            category_axis_text_style: crate::ir::ChartTextStyle::default(),
            value_axis_text_style: crate::ir::ChartTextStyle::default(),
            value_axis_number_format: None,
            auto_title_deleted: false,
            has_automatic_title: false,
            title_layout: None,
            plot_area_layout: None,
            user_shapes: Vec::new(),
        },
    });

    augment_page_with_print_headings(&mut page, &[1], 1, None);

    assert_eq!(page.images[0].x_offset_pt, 10.0 + 23.0);
    assert_eq!(page.images[0].y_offset_pt, 5.0 + 13.0);
    assert_eq!(page.text_boxes[0].x_offset_pt, 23.0);
    assert_eq!(page.text_boxes[0].y_offset_pt, 13.0);
    // An anchored chart is overlaid at the same absolute coordinates, so it
    // moves with the grid exactly as the other drawings do.
    let placement = page.charts[0]
        .placement
        .expect("the chart keeps its placement");
    assert_eq!(placement.x_offset_pt, 30.0 + 23.0);
    assert_eq!(placement.y_offset_pt, 60.0 + 13.0);
}

#[test]
fn test_title_columns_adjust_to_include_the_gutter() {
    // Without headings the range is untouched.
    assert_eq!(
        heading_adjusted_title_columns(Some((0, 2)), false),
        Some((0, 2))
    );
    assert_eq!(heading_adjusted_title_columns(None, false), None);
    // With headings the gutter repeats on every overflow page; left-anchored
    // print-title columns extend the same contiguous block.
    assert_eq!(heading_adjusted_title_columns(None, true), Some((0, 1)));
    assert_eq!(
        heading_adjusted_title_columns(Some((0, 2)), true),
        Some((0, 3))
    );
}

#[test]
fn test_non_left_anchored_title_columns_keep_repeating_shifted_past_the_gutter() {
    // Print-title columns at data indices 2..4 do not touch the sheet's left
    // edge. The gutter insertion shifts them one column right; they must
    // keep repeating — dropping them for the gutter would regress behavior
    // that predates #623 (adversarial review, finding 4).
    assert_eq!(
        heading_adjusted_title_columns(Some((2, 4)), true),
        Some((3, 5))
    );
}
