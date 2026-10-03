use super::*;
use crate::ir::{BorderSide, CellBorder, Insets, LineCap, LineJoin, Table, TableCell, TableRow};

/// Helper to create a table cell with plain text.
pub(super) fn make_text_cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    }
}

#[test]
fn test_table_simple_2x2() {
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("columns: (100pt, 200pt)"),
        "Expected column widths in: {result}"
    );
    assert!(result.contains("A1"), "Expected A1 in: {result}");
    assert!(result.contains("B1"), "Expected B1 in: {result}");
    assert!(result.contains("A2"), "Expected A2 in: {result}");
    assert!(result.contains("B2"), "Expected B2 in: {result}");
}

#[test]
fn test_table_with_default_cell_padding() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Padded")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 2.0,
            right: 3.0,
            bottom: 1.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("inset: (top: 2pt, right: 3pt, bottom: 1pt, left: 4pt)"),
        "Expected table inset in: {result}"
    );
}

#[test]
fn test_table_cell_with_padding_override() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Inset".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        padding: Some(Insets {
            top: 5.0,
            right: 2.0,
            bottom: 3.0,
            left: 6.0,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 1.0,
            right: 2.0,
            bottom: 3.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("table.cell(inset: (top: 5pt, right: 2pt, bottom: 3pt, left: 6pt))"),
        "Expected cell inset override in: {result}"
    );
}

#[test]
fn test_table_alignment_center_wraps_table() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Centered table")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        non_repeating_header_row_count: 0,
        alignment: Some(Alignment::Center),
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
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#align(center)["),
        "Expected center wrapper in: {result}"
    );
    assert!(
        result.contains("#table("),
        "Expected table inside wrapper in: {result}"
    );
}

#[test]
fn test_table_with_repeating_header_rows_uses_table_header() {
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Header 1"), make_text_cell("Header 2")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        header_row_count: 1,
        non_repeating_header_row_count: 0,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("table.header("),
        "Expected table.header wrapper in: {result}"
    );
    assert!(
        result.contains("Header 1") && result.contains("Body 1"),
        "Expected header and body cell content in: {result}"
    );
}

#[test]
fn test_repeating_header_expands_to_cover_rowspan() {
    let spanning_header = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Spanning header".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        row_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![spanning_header, make_text_cell("Year")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Month")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Day")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        header_row_count: 1,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let output = generate_typst(&doc).unwrap();
    let header = first_table_header_block(&output.source);

    assert!(header.contains("Spanning header"), "{}", output.source);
    assert!(header.contains("Month"), "{}", output.source);
    assert!(header.contains("Day"), "{}", output.source);
    assert!(!header.contains("Body 1"), "{}", output.source);
    crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
}

#[test]
fn test_repeating_header_expansion_follows_rowspans_from_newly_covered_rows() {
    let mut first_header = make_text_cell("First header");
    first_header.row_span = 2;
    let mut extending_header = make_text_cell("Extending header");
    extending_header.row_span = 3;
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![first_header, make_text_cell("Year")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![extending_header],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Covered 1")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Covered 2")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        header_row_count: 1,
        ..Table::default()
    };
    let output =
        generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])])).unwrap();
    let header = first_table_header_block(&output.source);

    assert!(header.contains("Extending header"), "{}", output.source);
    assert!(header.contains("Covered 1"), "{}", output.source);
    assert!(header.contains("Covered 2"), "{}", output.source);
    assert!(!header.contains("Body 1"), "{}", output.source);
    crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
}

fn first_table_header_block(source: &str) -> &str {
    let header_start = source.find("table.header(").expect("header block");
    let header_end = source[header_start..]
        .find("\n  ),\n")
        .map(|offset| header_start + offset)
        .expect("header block must close");
    &source[header_start..header_end]
}

#[test]
fn test_table_with_colspan() {
    let merged_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Merged".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        col_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![merged_cell],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan: 2 in: {result}"
    );
    assert!(result.contains("Merged"), "Expected Merged in: {result}");
}

#[test]
fn test_table_with_rowspan() {
    let tall_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Tall".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![tall_cell, make_text_cell("B1")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("rowspan: 2"),
        "Expected rowspan: 2 in: {result}"
    );
    assert!(result.contains("Tall"), "Expected Tall in: {result}");
}

#[test]
fn test_table_with_explicit_row_sizes_and_cell_vertical_align() {
    let centered_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Centered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        vertical_align: Some(CellVerticalAlign::Center),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![centered_cell, make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("rows: (36pt, auto)"),
        "Expected explicit Typst row sizes in: {result}"
    );
    assert!(
        result.contains("align: horizon"),
        "Expected centered vertical alignment in: {result}"
    );
}

#[test]
fn test_table_with_content_driven_row_heights_omits_explicit_rows() {
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: Some(48.0),
            },
        ],
        column_widths: vec![100.0, 100.0],
        use_content_driven_row_heights: true,
        default_vertical_align: None,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("rows: ("),
        "Content-driven row-height tables should not emit exact Typst row sizes: {result}"
    );
}

#[test]
fn test_table_with_colspan_and_rowspan() {
    let big_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Big".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        col_span: 2,
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![big_cell, make_text_cell("C1")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("C2")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![
                    make_text_cell("A3"),
                    make_text_cell("B3"),
                    make_text_cell("C3"),
                ],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan: 2 in: {result}"
    );
    assert!(
        result.contains("rowspan: 2"),
        "Expected rowspan: 2 in: {result}"
    );
    assert!(result.contains("Big"), "Expected Big in: {result}");
}

#[test]
fn test_table_with_background_color() {
    let colored_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Colored".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        background: Some(Color::new(200, 200, 200)),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![colored_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("fill: rgb(200, 200, 200)"),
        "Expected fill color in: {result}"
    );
    assert!(result.contains("Colored"), "Expected Colored in: {result}");
}

#[test]
fn a_translucent_table_cell_fill_reaches_typst() {
    let cell = TableCell {
        background: Some(Color::new(200, 200, 200)),
        background_alpha: Some(0.5),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(20.0),
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])]))
        .unwrap()
        .source;

    assert!(source.contains("fill: rgb(200, 200, 200, 128)"), "{source}");
}

#[test]
fn test_table_with_cell_borders() {
    let bordered_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bordered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            bottom: Some(BorderSide {
                width: 2.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![bordered_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("stroke:"), "Expected stroke in: {result}");
    assert!(
        result.contains("Bordered"),
        "Expected Bordered in: {result}"
    );
}

#[test]
fn test_table_with_partial_cell_borders_does_not_fill_missing_grid_lines() {
    let header_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Header".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        border: Some(CellBorder {
            top: None,
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![header_cell],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![make_text_cell("Body")],
                height: None,
            },
        ],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("stroke: none"),
        "Expected table default stroke to be disabled so unbordered cells stay unbordered: {result}"
    );
    assert!(
        result.contains("stroke: (bottom: 1pt + rgb(0, 0, 0))"),
        "Expected explicit bottom border to remain on the header cell: {result}"
    );
}

#[test]
fn test_table_with_styled_text_in_cell() {
    let styled_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold cell".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    font_size: Some(14.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![styled_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("weight: \"bold\""),
        "Expected bold in table cell: {result}"
    );
    assert!(
        result.contains("size: 14pt"),
        "Expected font size in table cell: {result}"
    );
}

#[test]
fn test_table_cell_paragraph_preserves_right_alignment() {
    let right_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Right),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "N".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("greek"), right_cell],
            height: None,
        }],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#block(width: 100%)") && result.contains("#set align(right)"),
        "Expected table cell paragraph to preserve right alignment: {result}"
    );
}

#[test]
fn test_table_cell_paragraph_preserves_spacing() {
    let spaced_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                space_before: Some(2.0),
                space_after: Some(3.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Header".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![spaced_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#v(2pt)") && result.contains("#v(3pt)"),
        "Expected table cell paragraph spacing to be preserved: {result}"
    );
}

#[test]
fn test_table_cell_word_line_box() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                line_box: Some(LineBox {
                    ascent_em: 1.3125,
                    descent_em: 0.4375,
                }),
                space_after: Some(8.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Word line box".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#set text(top-edge: 1.3125em, bottom-edge: -0.4375em)"),
        "Expected Word-compatible text edges in: {result}"
    );
    assert!(
        result.contains("#set par(leading: 0pt)"),
        "Expected Word-compatible line stacking in: {result}"
    );
    assert!(
        result.contains("#v(8pt)"),
        "Expected space-after in: {result}"
    );
}

#[test]
fn test_table_empty_cells() {
    let empty_cell = TableCell::default();
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![empty_cell, make_text_cell("Has text")],
            height: None,
        }],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("Has text"),
        "Expected Has text in: {result}"
    );
}

#[test]
fn test_table_no_column_widths() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("A"), make_text_cell("B")],
            height: None,
        }],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(result.contains("A"), "Expected A in: {result}");
    assert!(result.contains("B"), "Expected B in: {result}");
}

#[path = "typst_gen_table_border_tests.rs"]
mod border_tests;

#[path = "typst_gen_table_cell_content_tests.rs"]
mod cell_content_tests;

#[test]
fn test_table_special_chars_in_cells() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Price: $100 #items")],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("\\$") && result.contains("\\#"),
        "Expected escaped special chars in: {result}"
    );
}

#[test]
fn test_table_in_flow_page_with_paragraphs() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Cell")],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Before table"),
        Block::Table(table),
        make_paragraph("After table"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("Before table"),
        "Expected Before table in: {result}"
    );
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("After table"),
        "Expected After table in: {result}"
    );
}

#[test]
fn test_generate_space_before_after() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            space_before: Some(12.0),
            space_after: Some(6.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Spaced paragraph".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
            inline_box: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("12pt") || result.contains("above"),
        "Expected space_before in: {result}"
    );
}

/// Excel rests a bottom-aligned cell's last-line descender on the bottom
/// inset edge. The spill wrapper used to hardcode a `horizon` anchor and an
/// ambient-sized `1.3em` box, so an unwrapped title in a tall row floated at
/// the row's vertical centre instead (issue #618). The clip box and the
/// in-flow strut must both be sized from the paragraph's own line box at the
/// run's font size, not the ambient text size.
#[test]
fn bottom_aligned_spill_cell_anchors_its_line_box_at_the_bottom() {
    let Some((ascender, descender, _word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let font_size: f64 = 10.0;
    // Excel rests the descent on the row's own bottom boundary, one cell inset
    // below the box's bottom edge (issue #1063). Here that seat falls inside
    // the 5pt inset, and Typst clamps a box's `bottom-edge` at the baseline,
    // so the single line spans the ascent alone at the run's own size and the
    // placed clip box is translated down by the remainder (issue #1545).
    let default_padding_bottom_pt: f64 = 5.0;
    let seat_shortfall_pt: f64 = default_padding_bottom_pt - (descender * font_size).round();
    assert!(
        seat_shortfall_pt > 0.0,
        "the control needs a seat inside the inset, descent {descender}"
    );
    let line_box_height_pt: f64 = ascender * font_size;
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Warehouse Inventory".to_string(),
                style: TextStyle {
                    font_family: Some("Libertinus Serif".to_string()),
                    font_size: Some(font_size),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        spill_width: Some(200.0),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(23.0),
        }],
        column_widths: vec![60.0],
        default_vertical_align: Some(CellVerticalAlign::Bottom),
        seats_bottom_aligned_text_on_descender: true,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(&format!(
            "place(left + bottom, dx: 0pt, dy: {}pt, box(width: 195pt, height: {}pt, clip: true)",
            crate::render::typst_gen::tables::format_geometry(seat_shortfall_pt),
            format_f64(line_box_height_pt)
        )),
        "bottom cell's spill box must anchor at the bottom, sized to its own line and \
         dropped by the seat's remainder: {result}"
    );
    assert!(
        result.contains(&format!(
            "#box(width: 0pt, height: {}pt)",
            format_f64(line_box_height_pt)
        )),
        "the in-flow strut must hold the same line height in points: {result}"
    );
    assert!(
        !result.contains("horizon"),
        "a bottom-aligned spill cell must not be vertically centred: {result}"
    );
}

/// A centred spill cell keeps the `horizon` anchor #618 measured correct, but
/// its clip box comes from the cell's own line, not from `1.3em` of whatever
/// text size happens to surround the table. `em` there resolved against the
/// ambient size, so a 42pt title on an 11pt sheet was clipped to 14.30pt and
/// lost every descender in `Prosjektplanlegging` (issue #927).
#[test]
fn center_aligned_spill_cell_sizes_its_clip_box_from_its_own_font() {
    for font_size in [10.0_f64, 42.0_f64] {
        let Some((ascender, descender, _word_pitch_em)) =
            crate::render::pdf::font_line_metrics_em("Libertinus Serif")
        else {
            return; // no font book available (e.g. exotic CI sandbox)
        };
        let line_box_height_pt: f64 = (ascender + descender) * font_size;
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Centered spill".to_string(),
                    style: TextStyle {
                        font_family: Some("Libertinus Serif".to_string()),
                        font_size: Some(font_size),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            spill_width: Some(200.0),
            vertical_align: Some(CellVerticalAlign::Center),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells: vec![cell],
                height: Some(23.0),
            }],
            column_widths: vec![60.0],
            default_vertical_align: Some(CellVerticalAlign::Bottom),
            seats_bottom_aligned_text_on_descender: true,
            bottom_aligned_descent_floor_pt: 0.0,
            border_paint_model: TableBorderPaintModel::CenteredStroke,
            prints_gridlines: false,
            prints_headings: false,
            centers_between_print_margins: false,
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(&format!(
                "place(left + horizon, box(width: 195pt, height: {}pt, clip: true)",
                format_f64(line_box_height_pt)
            )),
            "a centred spill cell stays centred but is sized to its own line \
             at {font_size}pt: {result}"
        );
        assert!(
            result.contains(&format!(
                "#box(width: 0pt, height: {}pt)",
                format_f64(line_box_height_pt)
            )),
            "the in-flow strut must hold the same line height at {font_size}pt: {result}"
        );
    }
}

/// Excel cuts an unwrapped line at the cell's own gridline, and `#place`
/// starts at the *content* box the inset has already pushed in. A clip box
/// given the whole spill width therefore overhangs the gridline by the inset
/// on the side it is anchored to, and every glyph in that overhang is one
/// Excel does not print (issue #1105).
///
/// A centred box keeps the whole width instead: it sits on the content box's
/// centre, which is the cell's own centre while the two insets match. Only
/// the difference between them offsets it, and only that difference has to
/// come off — as it does for an icon-set cell, whose left inset carries the
/// icon's reserve.
#[test]
fn a_spill_clip_box_stops_at_the_cell_edge_its_anchor_faces() {
    let cases: [(Option<Alignment>, Insets, &str, f64); 4] = [
        // A symmetric synthetic pair, which is what exercises the rule: the
        // real sheet box is 3pt left against 2pt right (issue #1157), and the
        // asymmetric case is the icon-set one below.
        (None, xlsx_test_padding(3.0), "left", 197.0),
        (Some(Alignment::Left), xlsx_test_padding(3.0), "left", 197.0),
        (
            Some(Alignment::Right),
            xlsx_test_padding(3.0),
            "right",
            197.0,
        ),
        (
            Some(Alignment::Center),
            xlsx_test_padding(3.0),
            "center",
            200.0,
        ),
        // An icon-set cell reserves the icon's advance on the left, so its
        // centre is 9.6pt right of the cell's (issue #652).
    ];
    let icon_set_case: (Option<Alignment>, Insets, &str, f64) = (
        Some(Alignment::Center),
        xlsx_test_padding(12.6),
        "center",
        190.4,
    );

    for (alignment, padding, anchor, expected_width_pt) in cases.into_iter().chain([icon_set_case])
    {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment,
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "Wrapping paper".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            spill_width: Some(200.0),
            padding: Some(padding),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells: vec![cell],
                height: Some(23.0),
            }],
            column_widths: vec![200.0],
            prints_gridlines: false,
            prints_headings: false,
            centers_between_print_margins: false,
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(&format!(
                "place({anchor} + horizon, box(width: {}pt,",
                format_f64(expected_width_pt)
            )),
            "a {alignment:?} spill cell anchored {anchor} must clip at the cell edge, \
             not {}pt past it: {result}",
            format_f64(200.0 - expected_width_pt)
        );
    }
}

/// Excel's horizontal cell inset, as the spill tests above state it.
fn xlsx_test_padding(left: f64) -> Insets {
    Insets {
        top: 1.0,
        right: 3.0,
        bottom: 1.5,
        left,
    }
}

// ---------------------------------------------------------------------------
// Printed headings (issue #623)
//
// `<printOptions headings="1"/>` prints the row-number gutter and the
// column-letter strip on every page. The XLSX parser materializes them in the
// IR (gutter as the first column, letter strip as `rows[0]`) and sets
// `prints_headings`; codegen must re-emit that first row as a repeating
// `table.header` so the letters appear on every paginated page, above any
// print-title headers.
// ---------------------------------------------------------------------------

/// The augmented shape the XLSX parser hands codegen: strip row first, gutter
/// cells at each row's front.
fn headings_table(extra_rows: Vec<TableRow>) -> Table {
    let mut rows: Vec<TableRow> = vec![TableRow {
        minimum_height: None,
        cells: vec![
            TableCell::default(),
            make_text_cell("A"),
            make_text_cell("B"),
        ],
        height: Some(13.0),
    }];
    rows.extend(extra_rows);
    Table {
        rows,
        column_widths: vec![23.0, 120.0, 110.0],
        prints_headings: true,
        ..Table::default()
    }
}

fn gutter_row(number: &str, first: &str, second: &str) -> TableRow {
    TableRow {
        minimum_height: None,
        cells: vec![
            make_text_cell(number),
            make_text_cell(first),
            make_text_cell(second),
        ],
        height: Some(13.0),
    }
}

#[test]
fn test_prints_headings_emits_the_letter_strip_as_a_repeating_header() {
    let table = headings_table(vec![gutter_row("1", "x", "y"), gutter_row("2", "z", "w")]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    let header_start: usize = result
        .find("table.header(repeat: true,")
        .expect("the letter strip must repeat on every page");
    let header_end: usize = result[header_start..]
        .find("),")
        .map(|offset| header_start + offset)
        .expect("header block must close");
    let header_block: &str = &result[header_start..header_end];
    assert!(
        header_block.contains("[A]") && header_block.contains("[B]"),
        "the strip header must carry the column letters: {result}"
    );
    assert!(
        !header_block.contains("[1]"),
        "gutter numbers flow with the body rows, not the strip: {result}"
    );
    let body: &str = &result[header_end..];
    assert!(
        body.contains("[1]") && body.contains("[2]"),
        "gutter numbers must follow in the body: {result}"
    );
    assert!(
        result.contains("columns: (23pt, 120pt, 110pt)"),
        "the gutter track must lead the column list: {result}"
    );
    assert!(
        result.contains("rows: (13pt, 13pt, 13pt)"),
        "the strip row's track must lead the row list: {result}"
    );
}

#[test]
fn test_prints_headings_keeps_print_title_headers_below_the_strip() {
    let mut table = headings_table(vec![
        gutter_row("1", "Title", "Title2"),
        gutter_row("2", "x", "y"),
    ]);
    // One print-title row (counted AFTER the strip row).
    table.header_row_count = 1;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    let strip_pos: usize = result
        .find("table.header(repeat: true,")
        .expect("strip header must exist");
    let title_pos: usize = result
        .find("table.header(level: 2,")
        .expect("print-title header must sit at the next level");
    assert!(
        strip_pos < title_pos,
        "the strip repeats above the print-title rows: {result}"
    );
    assert!(
        result[title_pos..].contains("Title"),
        "the title row must live in the level-2 header: {result}"
    );
}

#[test]
fn test_prints_headings_shifts_lead_and_title_header_levels() {
    let mut table = headings_table(vec![
        gutter_row("1", "Lead", "Lead2"),
        gutter_row("2", "Title", "Title2"),
        gutter_row("3", "x", "y"),
    ]);
    // One non-repeating lead row above one print-title row.
    table.non_repeating_header_row_count = 1;
    table.header_row_count = 1;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("table.header(repeat: true,"),
        "strip header must exist: {result}"
    );
    assert!(
        result.contains("table.header(repeat: false, level: 2,"),
        "the lead block must move to level 2 under the strip: {result}"
    );
    assert!(
        result.contains("table.header(level: 3,"),
        "the print-title block must move to level 3: {result}"
    );
}

#[test]
fn test_prints_headings_off_emits_no_repeating_strip() {
    let mut table = headings_table(vec![gutter_row("1", "x", "y")]);
    table.prints_headings = false;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("table.header("),
        "without the flag no header block may appear: {result}"
    );
}

/// The gray of the 1pt rules between heading cells (issue #623 GT model).
const HEADING_RULE_GRAY: crate::ir::Color = crate::ir::Color {
    r: 86,
    g: 86,
    b: 86,
};

fn solid_side(width: f64, color: crate::ir::Color) -> BorderSide {
    BorderSide {
        width,
        color,
        style: crate::ir::BorderLineStyle::Solid,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    }
}

/// The bordered shape the XLSX parser's heading augmentation hands codegen:
/// gray rules between heading cells, black 1pt separators against the data
/// grid (strip bottom, gutter right), boundary-band border regime.
fn separator_bordered_headings_table(data_cell: TableCell, prints_gridlines: bool) -> Table {
    let gray = || solid_side(1.0, HEADING_RULE_GRAY);
    let black_separator = || solid_side(1.0, crate::ir::Color::black());
    Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![
                    TableCell {
                        border: Some(CellBorder {
                            bottom: Some(gray()),
                            right: Some(gray()),
                            ..CellBorder::default()
                        }),
                        ..TableCell::default()
                    },
                    TableCell {
                        border: Some(CellBorder {
                            bottom: Some(black_separator()),
                            right: Some(gray()),
                            ..CellBorder::default()
                        }),
                        ..make_text_cell("A")
                    },
                ],
                height: Some(13.0),
            },
            TableRow {
                minimum_height: None,
                cells: vec![
                    TableCell {
                        border: Some(CellBorder {
                            top: Some(gray()),
                            bottom: Some(gray()),
                            right: Some(black_separator()),
                            ..CellBorder::default()
                        }),
                        ..make_text_cell("1")
                    },
                    data_cell,
                ],
                height: Some(13.0),
            },
        ],
        column_widths: vec![23.0, 120.0],
        prints_headings: true,
        prints_gridlines,
        border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
        ..Table::default()
    }
}

/// The `table.header(repeat: true, ...)` slice of the generated source.
fn strip_header_block(result: &str) -> &str {
    let header_start: usize = result
        .find("table.header(repeat: true,")
        .expect("strip header must exist");
    let header_end: usize = result[header_start..]
        .find("\n  ),")
        .map(|offset| header_start + offset)
        .expect("header block must close");
    &result[header_start..header_end]
}

#[test]
fn test_prints_headings_strip_bottom_separator_repeats_over_a_tying_body_border() {
    // The first body row declares a 1pt solid black top border — the same
    // conflict rank as the strip's black separator. The tie must resolve to
    // the strip side (the #619 repeating-header rule, applied at the strip's
    // own bottom boundary): a band left on the body row's top would vanish
    // on pages 2+, where only the header repeats (issue #623 adversarial
    // review, finding 1).
    let data_cell = TableCell {
        border: Some(CellBorder {
            top: Some(solid_side(1.0, crate::ir::Color::black())),
            ..CellBorder::default()
        }),
        ..make_text_cell("x")
    };
    let table = separator_bordered_headings_table(data_cell, false);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        strip_header_block(&result).contains("#place(bottom + left"),
        "the strip|data separator band must live inside the repeating header: {result}"
    );
}

#[test]
fn test_prints_headings_strip_adopts_a_heavier_body_border_into_the_repeat() {
    // A strictly heavier body declaration (2pt medium) outranks the strip's
    // 1pt separator; the #619 mechanism adopts it into the header side so
    // the heavier band repeats on every page alongside the body's own copy.
    let data_cell = TableCell {
        border: Some(CellBorder {
            top: Some(solid_side(2.0, crate::ir::Color::black())),
            ..CellBorder::default()
        }),
        ..make_text_cell("x")
    };
    let table = separator_bordered_headings_table(data_cell, false);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        strip_header_block(&result).contains("stroke: 2pt"),
        "the header must adopt the body's heavier band so repeats carry it: {result}"
    );
}

#[test]
fn test_prints_headings_excludes_the_heading_exterior_from_gridline_seeding() {
    // gridLines=1 AND headings=1: GT rules the heading exterior — the strip
    // row's top and the gutter column's left — as the 1pt black print FRAME,
    // not as gridlines (issue #623 evidence, nft-sheet-0002 trace: black
    // bands [54,538]x[72,73] and [54,55]x[72,710]). Gridline seeding stays
    // excluded there; the frame band replaces it rather than stacking, so
    // every frame boundary carries exactly one band. Data-area seeding
    // (#622) is unchanged where the frame does not own the boundary.
    let table = separator_bordered_headings_table(make_text_cell("x"), true);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#place(top + left"),
        "the strip-top frame band must paint: {result}"
    );
    // 10 black bands, each boundary painted once: corner top+left, letter
    // top+right (frame over the last column's gray rule) + bottom separator,
    // gutter left + right separator + bottom (frame over the last row's gray
    // rule), data cell right+bottom (frame replacing the #622 closure seeds).
    assert_eq!(
        result.matches("rgb(0, 0, 0)").count(),
        10,
        "the frame and separators paint exactly once per boundary — no \
         doubled bands and no leftover gridline seeds: {result}"
    );
}

#[test]
fn test_prints_headings_paints_the_black_exterior_frame_without_gridlines() {
    // GT (issue #623): with headings on, a 1pt black frame encloses the
    // heading bands and the data grid — the corner box's top and left edges
    // ARE the frame. The frame is gated on prints_headings alone; gridlines
    // are off here, so every exterior band below is the frame's.
    let table = separator_bordered_headings_table(make_text_cell("x"), false);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        strip_header_block(&result).contains("#place(top + left"),
        "the strip-top frame edge must live inside the repeating header so \
         pages 2+ carry it: {result}"
    );
    // Same 10-band census as the gridlines-on test: the frame does not
    // depend on gridline seeding.
    assert_eq!(
        result.matches("rgb(0, 0, 0)").count(),
        10,
        "frame top/left/right/bottom plus the strip and gutter separators \
         paint exactly once per boundary: {result}"
    );
}

/// A slide table cell with no explicit line spacing uses PowerPoint's flat 1.2em line.
///
/// A `<a:tbl>` reaches the shared table codegen, which gave its cells Word's
/// hhea line box. Measured on `office2pdf_introduction_ko` slide 16: an 11pt
/// cell advanced 17.46pt (1.587em) against PowerPoint's documented 1.2em, so
/// multi-line cells grew and the table's bottom border moved down with them.
/// Explicit proportional spacing above 100% uses a separate PowerPoint table
/// line-box rule; this test covers the plain `a:lnSpc`-absent case.
/// A LibreOffice render of the same slide advances 13.58pt (1.235em) — a
/// corroborating reference, not a native export (issue #663).
#[test]
fn test_slide_table_cell_uses_the_powerpoint_line_box() {
    let sized_cell = |text: &str| TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some("Liberation Sans".to_string()),
                    font_size: Some(11.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![sized_cell("wrapped cell text")],
            height: None,
        }],
        column_widths: vec![80.0],
        ..Table::default()
    };

    let slide = generate_typst(&make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 80.0,
            height: 40.0,
            kind: FixedElementKind::Table(table.clone()),
        }],
    )]))
    .unwrap()
    .source;
    let flow = generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])]))
        .unwrap()
        .source;

    // With no explicit line spacing, PowerPoint's plain table-cell line is a
    // flat 1.2em box. Word's line is the face's own hhea pitch.
    let (top, bottom) =
        emitted_slide_line_box_em(&slide, 11.0).expect("slide cell emits a line box");
    assert!(
        (top + bottom - 1.2).abs() < 0.001,
        "a slide cell's line must span 1.2em, got {top} + {bottom}: {slide}"
    );
    // The flow-page half needs the declared face's real hhea metrics, which a
    // runner without Liberation Sans cannot resolve — `word_cell_line_box`
    // then emits no box at all. The slide assertion above holds regardless,
    // because PowerPoint's 1.2em falls back to the default face.
    if crate::render::pdf::font_line_metrics_em("Liberation Sans").is_some()
        && let Some(flow_box) = emitted_line_box_em(&flow)
    {
        assert!(
            (flow_box.0 + flow_box.1 - 1.2).abs() > 0.001,
            "a flow-page cell must keep Word's hhea line, not PowerPoint's: {flow_box:?}"
        );
    }
}

fn tracked_slide_table_cell_source(
    text: &str,
    alignment: Alignment,
    spacing_pt: Option<f64>,
) -> String {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(alignment),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(18.0),
                    letter_spacing: spacing_pt,
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(36.0),
        }],
        column_widths: vec![120.0],
        ..Table::default()
    };

    generate_typst(&make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 72.0,
            width: 120.0,
            height: 36.0,
            kind: FixedElementKind::Table(table),
        }],
    )]))
    .unwrap()
    .source
}

/// PowerPoint places tracked table-cell lines by a width that includes one
/// `a:rPr/@spc` after the final glyph, just like a slide text box.
///
/// A native PowerPoint 16.112 one-factor probe varied only the tracking on the
/// seven-glyph headers in `table_test2.pptx`; its no-patch re-zip control was
/// layout-identical. The centred header's origin moved from 368.3125pt to
/// 364.8125, 357.8125 and 347.3125pt at 1, 3 and 6pt tracking: half of the
/// full 7, 21 and 42pt tracked widths. The right-aligned header moved from
/// 249.325pt to 242.325, 228.325 and 207.325pt: those widths in full. Typst
/// tracks only between glyphs, so both alignments need the final run's own
/// tracking appended to the line (issue #1256).
#[test]
fn tracked_slide_table_cell_reserves_its_trailing_letter_space() {
    for alignment in [Alignment::Center, Alignment::Right] {
        for spacing_pt in [1.0, 3.0, 6.0] {
            let source = tracked_slide_table_cell_source("Header2", alignment, Some(spacing_pt));
            let expected = format!("]#h({spacing_pt}pt)");

            assert!(
                source.contains(&expected),
                "{alignment:?} table-cell text must reserve its final {spacing_pt}pt \
                 letter-space: {source}"
            );
        }
    }
}

/// A left-aligned line starts at the content edge regardless of its measured
/// width, and an untracked line has no reserve to make. Neither case may emit
/// a spacer that could push a borderline table-cell line into a wrap.
#[test]
fn left_or_untracked_slide_table_cell_reserves_no_trailing_space() {
    for (alignment, spacing_pt) in [
        (Alignment::Left, Some(3.0)),
        (Alignment::Center, None),
        (Alignment::Center, Some(0.0)),
    ] {
        let source = tracked_slide_table_cell_source("Header2", alignment, spacing_pt);
        assert!(
            !source.contains("]#h("),
            "no trailing reserve is due for {alignment:?} at {spacing_pt:?}: {source}"
        );
    }
}

/// Every visible physical line owns the final glyph's tracking. A hard break
/// must not transfer the first line's reserve to the paragraph's last line or
/// omit the last line's reserve after enabling the shared per-break path.
#[test]
fn hard_broken_tracked_slide_table_cell_reserves_each_line_space() {
    let source = tracked_slide_table_cell_source("FIRST\nSECOND", Alignment::Center, Some(3.0));

    assert_eq!(
        source.matches("]#h(3pt)").count(),
        2,
        "both hard-broken table-cell lines must reserve their own final \
         letter-space: {source}"
    );
}

/// A hard break inside a centred slide-table cell clears the taller line above
/// it, not just the following line's own box.
///
/// The original slide 16 cell in `office2pdf_introduction_ko.pptx` is the
/// awkward boundary case: 13pt `DOCX` followed by 9.5pt `Word`. PowerPoint's
/// native export advances exactly 12.00pt, well past the 11.40pt the 9.5pt
/// line's own 1.2em box would give — a probe deck alternating the same two
/// sizes nine lines deep reads 12.00pt four times over, in a text box and in a
/// table cell alike. #683 fitted a 10pt line-box floor to that one number;
/// `1.2 x max(9.5, 10)` reaches 12.00 by coincidence, and the same floor paced
/// every hard-broken line under 10pt at a flat 12.00pt (issue #1172). What
/// actually carries the gap is the 13pt line's descent.
///
/// Each line now rounds its own seat (#1252), which predicts 11.60pt for this
/// boundary and remains within the native export's half-point dither band.
///
/// This compiles the real table-cell path so its vertical centring and line
/// box cannot make a source-only assertion pass by accident (issue #683).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn slide_table_hard_break_clears_the_taller_line_above_it() {
    let family = "Arial";
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                ..ParagraphStyle::default()
            },
            runs: vec![
                Run {
                    text: "DOCX\n".to_string(),
                    style: TextStyle {
                        font_family: Some(family.to_string()),
                        font_size: Some(13.0),
                        bold: Some(true),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                },
                Run {
                    text: "Word".to_string(),
                    style: TextStyle {
                        font_family: Some(family.to_string()),
                        font_size: Some(9.5),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                },
            ],
        })],
        vertical_align: Some(CellVerticalAlign::Center),
        padding: Some(Insets {
            top: 3.6,
            right: 7.2,
            bottom: 3.6,
            left: 7.2,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(66.24),
        }],
        column_widths: vec![120.0],
        ..Table::default()
    };
    let output = generate_typst(&make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 72.0,
            width: 120.0,
            height: 66.24,
            kind: FixedElementKind::Table(table),
        }],
    )]))
    .unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
    let docx_baseline = runs
        .iter()
        .find(|run| run.text.contains("DOCX"))
        .expect("DOCX run")
        .baseline_pt;
    let word_baseline = runs
        .iter()
        .find(|run| run.text.contains("Word"))
        .expect("Word run")
        .baseline_pt;

    let advance_pt: f64 = word_baseline - docx_baseline;
    let (plain_ascent_em, _) = crate::render::pdf::powerpoint_line_box_em(family)
        .expect("the Arial-compatible line metrics must resolve");
    let line_box_pt = |size_pt: f64| -> (f64, f64) {
        let (top_em, bottom_em) = powerpoint_percentage_line_box_em(plain_ascent_em, size_pt, 1.0);
        (top_em * size_pt, bottom_em * size_pt)
    };
    let (_, docx_bottom_pt) = line_box_pt(13.0);
    let (word_top_pt, _) = line_box_pt(9.5);
    let expected_pt = docx_bottom_pt + word_top_pt;
    assert!(
        (advance_pt - expected_pt).abs() < 0.01,
        "the break must combine the 13pt line's independently rounded descent \
         with the 9.5pt line's independently rounded seat: expected \
         {expected_pt}pt, got {advance_pt}pt from {docx_baseline}, \
         {word_baseline}\n{}",
        output.source
    );
    assert!(
        (advance_pt - 12.0).abs() < 0.5,
        "the break must land within half a point of the native export's 12pt: \
         {docx_baseline}, {word_baseline}\n{}",
        output.source
    );
}

/// An empty paragraph's blank line in a slide cell comes from the same model.
///
/// The strut that holds an empty `<a:p>`'s height was sized by
/// `word_cell_line_box` unconditionally, so a slide's empty cell would keep
/// Word's hhea height while the text beside it took PowerPoint's 1.2em one
/// (issues #625, #663).
#[test]
fn test_slide_table_empty_cell_blank_line_uses_the_powerpoint_line_box() {
    let sized_run = |text: &str| Run {
        text: text.to_string(),
        style: TextStyle {
            font_family: Some("Liberation Sans".to_string()),
            font_size: Some(20.0),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
        inline_box: None,
    };
    // A cell stacking a text paragraph and an empty one: the empty paragraph
    // takes its metrics from the neighbour.
    let cell = TableCell {
        content: vec![
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![sized_run("text")],
            }),
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: Vec::new(),
            }),
        ],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![120.0],
        ..Table::default()
    };

    let source = generate_typst(&make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 120.0,
            height: 60.0,
            kind: FixedElementKind::Table(table),
        }],
    )]))
    .unwrap()
    .source;

    // PowerPoint's line at 20pt is a flat 1.2em = 24pt, whatever the face.
    assert!(
        source.contains("height: 24pt"),
        "the blank line must be PowerPoint's 1.2em (24pt at 20pt), got: {source}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// A spill cell's text stays on one line.
///
/// The clip box states a width, and a Typst box wraps its content at the width
/// it states. So the line broke into several, the clip hid all but one, and
/// because the wrapper is anchored vertically the one left visible was the
/// *tail*: the merged title in `merged_row_overflows_page_column.xlsx`
/// rendered starting mid-sentence with its opening words gone (issue #811).
///
/// Asserted on baselines rather than on the emitted markup, because the defect
/// was invisible in the source — the wrapper read exactly as intended and it
/// was Typst's layout of it that differed.
#[test]
fn spill_cell_text_is_not_wrapped_by_its_clip_box() {
    let long_text: &str = "This merged full-width title is deliberately far wider \
                           than the first horizontal page-column so that it must \
                           either be clipped at the page break or continued on the \
                           following page-column of the printout.";
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: long_text.to_string(),
                style: TextStyle {
                    font_family: Some("Libertinus Serif".to_string()),
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        // Far narrower than the text, which is the whole point of a spill.
        spill_width: Some(120.0),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![60.0],
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);

    let output = generate_typst(&doc).expect("document should generate");
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
    let mut baselines: Vec<f64> = runs
        .iter()
        .filter(|run| run.text.contains("merged") || run.text.contains("printout"))
        .map(|run| run.baseline_pt)
        .collect();
    baselines.sort_by(f64::total_cmp);
    baselines.dedup_by(|left, right| (*left - *right).abs() < 0.01);
    assert_eq!(
        baselines.len(),
        1,
        "a spill cell's text must occupy exactly one line, got {} baselines: {baselines:?}",
        baselines.len()
    );
}

/// `w:tblPr/w:jc` positions the table box. Typst's `align` is inherited, so
/// the wrapper that centres the box also centred every cell paragraph that
/// declared no alignment of its own (issue #843).
#[test]
fn test_centred_table_does_not_centre_its_cell_paragraphs() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Inherits left")],
            height: None,
        }],
        column_widths: vec![100.0],
        alignment: Some(Alignment::Center),
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#align(center)["),
        "Expected the table box to stay centred in: {result}"
    );
    assert!(
        result.contains("#set align(start)"),
        "Expected the cell to reset the inherited alignment in: {result}"
    );
}

/// The reset must not override a paragraph that declares its own alignment,
/// or a right-aligned total inside a centred table would move.
#[test]
fn test_centred_table_keeps_a_cell_paragraph_own_alignment() {
    let mut cell = make_text_cell("Right");
    if let Some(Block::Paragraph(paragraph)) = cell.content.first_mut() {
        paragraph.style.alignment = Some(Alignment::Right);
    }
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        alignment: Some(Alignment::Center),
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#set align(right)"),
        "Expected the declared right alignment to survive in: {result}"
    );
}

/// A table with no `w:jc` emits no wrapper, so there is nothing to reset and
/// the cells must stay exactly as they were.
#[test]
fn test_unaligned_table_emits_no_alignment_reset() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Plain")],
            height: None,
        }],
        column_widths: vec![100.0],
        alignment: None,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("#set align(start)"),
        "Expected no alignment reset in: {result}"
    );
}

/// A `w:trHeight` floor is `max(floor, content)`, which neither a stated Typst
/// row length nor `auto` expresses: the first pins the row and the second drops
/// the floor. The strut grid carries it in one cell of the row (issue #965).
#[test]
fn a_row_minimum_height_emits_a_strut_rather_than_a_fixed_row() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("Left"), make_text_cell("Right")],
            height: None,
            minimum_height: Some(110.75),
        }],
        column_widths: vec![100.0, 100.0],
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
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // 110.75pt less the 5pt default padding on each side.
    assert!(
        result.contains(
            "#grid(columns: (0pt, 1fr), rows: (auto,), box(width: 0pt, height: 100.75pt), ["
        ),
        "{result}"
    );
    // One strut per row, not one per cell: the row is as tall as its tallest
    // cell, so repeating it would only restate the same constraint.
    assert_eq!(
        result.matches("box(width: 0pt, height: 100.75pt)").count(),
        1
    );
    // The floor must not become a stated row length, which would pin the row
    // and stop it growing for taller content.
    assert!(!result.contains("rows: (110.75pt"), "{result}");
}

/// A table in the document body pins both of its own vertical gaps.
///
/// Word gives a table no spacing of its own, so the paragraph after it starts
/// one line below the last row unless it states `w:spacing w:before`. Typst
/// resolves the gap between two blocks by weakness, and an unstated gap falls
/// back to `par.spacing` — 1.2em — at a *weaker* level than a stated one, so a
/// bare `#table` beside a paragraph that also states nothing opened 1.2em of
/// engine whitespace (issue #1688).
#[test]
fn test_body_table_pins_its_own_block_spacing() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Warehouse lease")],
            height: None,
        }],
        column_widths: vec![468.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Table(table),
        Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Questions go to the shared tracker.".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        }),
    ])]);
    let result = generate_typst(&doc).unwrap().source;

    let wrapper: usize = result
        .find("#block(width: 100%, above: 0pt, below: 0pt)[")
        .unwrap_or_else(|| panic!("table is not wrapped in a spacing-pinned block: {result}"));
    let table_start: usize = result.find("#table(").expect("table");
    assert!(
        wrapper < table_start,
        "the pinned block must open before the table: {result}"
    );
}

/// Triangulation for the rule above: the wrapper belongs to the body flow, so
/// a table nested inside a cell does not gain a second one. A cell's own
/// content already pins what Word's cell margins say, and Typst trims weak
/// spacing at a container's edges anyway.
#[test]
fn test_nested_table_does_not_gain_a_second_spacing_block() {
    let inner = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Inner")],
            height: None,
        }],
        column_widths: vec![80.0],
        ..Table::default()
    };
    let outer = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![TableCell {
                content: vec![Block::Table(inner)],
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(outer)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("#table(").count(),
        2,
        "expected the nested table to be emitted: {result}"
    );
    assert_eq!(
        result
            .matches("#block(width: 100%, above: 0pt, below: 0pt)[")
            .count(),
        1,
        "only the body-flow table takes the wrapper: {result}"
    );
}

/// A centred table keeps its `w:tblPr/w:jc` placement inside the wrapper: the
/// block spans the text column, so `#align(center)` still has the full measure
/// to centre the table box in (issue #843).
#[test]
fn test_centered_body_table_keeps_its_alignment_inside_the_wrapper() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell("Centred")],
            height: None,
        }],
        column_widths: vec![120.0],
        alignment: Some(Alignment::Center),
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    let wrapper: usize = result
        .find("#block(width: 100%, above: 0pt, below: 0pt)[")
        .unwrap_or_else(|| panic!("table is not wrapped in a spacing-pinned block: {result}"));
    let align: usize = result.find("#align(center)[").expect("alignment");
    assert!(
        wrapper < align,
        "the wrapper must span the text column the table centres in: {result}"
    );
}
