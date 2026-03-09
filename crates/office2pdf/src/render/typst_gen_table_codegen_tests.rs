use super::*;
use crate::ir::{BorderSide, CellBorder, Insets, Table, TableCell, TableRow};

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
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: None,
            },
            TableRow {
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
            cells: vec![make_text_cell("Padded")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 2.0,
            right: 3.0,
            bottom: 1.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
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
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 1.0,
            right: 2.0,
            bottom: 3.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
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
            cells: vec![make_text_cell("Centered table")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: Some(Alignment::Center),
        default_cell_padding: None,
        use_content_driven_row_heights: false,
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
                cells: vec![make_text_cell("Header 1"), make_text_cell("Header 2")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        header_row_count: 1,
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
fn test_table_with_colspan() {
    let merged_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Merged".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![merged_cell],
                height: None,
            },
            TableRow {
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
            }],
        })],
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![tall_cell, make_text_cell("B1")],
                height: None,
            },
            TableRow {
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
            }],
        })],
        vertical_align: Some(CellVerticalAlign::Center),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![centered_cell, make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
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
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: Some(48.0),
            },
        ],
        column_widths: vec![100.0, 100.0],
        use_content_driven_row_heights: true,
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
            }],
        })],
        col_span: 2,
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![big_cell, make_text_cell("C1")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("C2")],
                height: None,
            },
            TableRow {
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
            }],
        })],
        background: Some(Color::new(200, 200, 200)),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
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
fn test_table_with_cell_borders() {
    let bordered_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bordered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: Some(BorderSide {
                width: 2.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Solid,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
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
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
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
fn test_table_empty_cells() {
    let empty_cell = TableCell::default();
    let table = Table {
        rows: vec![TableRow {
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

#[test]
fn test_table_all_borders() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "All borders".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            left: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            right: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("top:"), "Expected top border in: {result}");
    assert!(
        result.contains("bottom:"),
        "Expected bottom border in: {result}"
    );
    assert!(
        result.contains("left:"),
        "Expected left border in: {result}"
    );
    assert!(
        result.contains("right:"),
        "Expected right border in: {result}"
    );
}

#[test]
fn test_table_dashed_border_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Dashed".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Dotted,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("dash: \"dashed\""),
        "Expected dashed dash pattern in: {result}"
    );
    assert!(
        result.contains("dash: \"dotted\""),
        "Expected dotted dash pattern in: {result}"
    );
}

#[test]
fn test_shape_dashed_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Rectangle,
            Some(Color::new(0, 128, 255)),
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("dash: \"dashed\""),
        "Expected dashed stroke in: {}",
        output.source
    );
}

#[test]
fn test_shape_dash_dot_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Ellipse,
            None,
            Some(BorderSide {
                width: 1.0,
                color: Color::new(0, 0, 255),
                style: BorderLineStyle::DashDot,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("dash: \"dash-dotted\""),
        "Expected dash-dotted stroke in: {}",
        output.source
    );
}

#[test]
fn test_border_line_style_to_typst_mapping() {
    assert_eq!(border_line_style_to_typst(BorderLineStyle::Solid), "solid");
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dashed),
        "dashed"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dotted),
        "dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDot),
        "dash-dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDotDot),
        "dash-dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Double),
        "dashed"
    );
    assert_eq!(border_line_style_to_typst(BorderLineStyle::None), "solid");
}

#[test]
fn test_solid_border_no_dash_param() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Solid".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: None,
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("dash:"),
        "Solid border should not have dash parameter in: {result}"
    );
    assert!(
        result.contains("1pt + rgb(0, 0, 0)"),
        "Expected simple solid format in: {result}"
    );
}

#[test]
fn test_table_cell_with_multiple_paragraphs() {
    let multi_para_cell = TableCell {
        content: vec![
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "First para".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Second para".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
        ],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![multi_para_cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("First para"),
        "Expected First para in: {result}"
    );
    assert!(
        result.contains("Second para"),
        "Expected Second para in: {result}"
    );
}

#[test]
fn test_table_cell_simple_list_uses_compact_fixed_text_layout() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#stack(dir: ttb"),
        "Expected compact stack-based list layout in: {result}"
    );
    assert!(
        !result.contains("#list("),
        "Compact table-cell lists should not use Typst list layout in: {result}"
    );
    assert!(result.contains("First item"));
    assert!(result.contains("Second item"));
}

#[test]
fn test_table_cell_simple_list_treats_default_and_explicit_left_as_same_style() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Left),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#stack(dir: ttb"),
        "Expected compact stack-based list layout when only left-alignment explicitness differs: {result}"
    );
    assert!(
        !result.contains("#list("),
        "Equivalent left-alignment styles should not force Typst list layout in: {result}"
    );
}

#[test]
fn test_table_cell_compact_list_uses_leading_without_extra_item_spacing() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        line_spacing: Some(LineSpacing::Proportional(1.5)),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle {
                            font_size: Some(24.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        line_spacing: Some(LineSpacing::Proportional(1.5)),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle {
                            font_size: Some(24.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#set par(leading: 12pt)"),
        "Expected paragraph leading derived from PPT line spacing in: {result}"
    );
    assert!(
        !result.contains("#stack(dir: ttb, spacing: 12pt"),
        "Compact table-cell lists should not add extra inter-item spacing in: {result}"
    );
}

#[test]
fn test_table_special_chars_in_cells() {
    let table = Table {
        rows: vec![TableRow {
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
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("12pt") || result.contains("above"),
        "Expected space_before in: {result}"
    );
}
