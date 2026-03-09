use super::*;

// ── DataBar / IconSet codegen tests ──────────────────────────────

#[test]
fn test_data_bar_codegen() {
    use crate::ir::DataBarInfo;

    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "50".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        data_bar: Some(DataBarInfo {
            color: Color::new(0x63, 0x8E, 0xC6),
            fill_pct: 50.0,
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
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("fill: rgb(99, 142, 198)"),
        "DataBar should contain bar color fill. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("width: 50%"),
        "DataBar should contain 50% width. Got: {}",
        output.source,
    );
}

#[test]
fn test_icon_text_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "90".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        icon_text: Some("↑".to_string()),
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
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("↑"),
        "Icon text should appear in output. Got: {}",
        output.source,
    );
}

#[test]
fn test_table_colspan_clamped_to_available_columns() {
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![wide_cell],
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
        "Expected colspan clamped to 2, got: {result}"
    );
    assert!(
        !result.contains("colspan: 3"),
        "colspan: 3 should have been clamped, got: {result}"
    );
}

#[test]
fn test_table_colspan_clamped_mid_row() {
    let normal_cell = make_text_cell("A1");
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![normal_cell, wide_cell],
            height: None,
        }],
        column_widths: vec![100.0, 100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan clamped to 2, got: {result}"
    );
}

#[test]
fn test_table_colspan_no_column_widths_inferred() {
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 5,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![wide_cell],
                height: None,
            },
            TableRow {
                cells: vec![
                    make_text_cell("A"),
                    make_text_cell("B"),
                    make_text_cell("C"),
                ],
                height: None,
            },
        ],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 3"),
        "Expected colspan clamped to 3 (inferred columns), got: {result}"
    );
    assert!(
        !result.contains("colspan: 5"),
        "colspan: 5 should have been clamped, got: {result}"
    );
}

// ── Extended geometry codegen tests (US-085) ──────────────────────────

#[test]
fn test_triangle_polygon_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            150.0,
            ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
            },
            Some(Color::new(255, 0, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon in: {}",
        output.source
    );
    assert!(
        output.source.contains("100pt"),
        "Expected 100pt vertex x in: {}",
        output.source
    );
    assert!(
        output.source.contains("fill: rgb(255, 0, 0)"),
        "Expected fill in: {}",
        output.source
    );
}

#[test]
fn test_rounded_rectangle_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            100.0,
            ShapeKind::RoundedRectangle {
                radius_fraction: 0.1,
            },
            Some(Color::new(0, 0, 255)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#rect("),
        "Expected #rect in: {}",
        output.source
    );
    assert!(
        output.source.contains("radius:"),
        "Expected radius parameter in: {}",
        output.source
    );
    assert!(
        output.source.contains("radius: 10pt"),
        "Expected radius: 10pt in: {}",
        output.source
    );
}

#[test]
fn test_arrow_polygon_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            300.0,
            150.0,
            ShapeKind::Polygon {
                vertices: vec![
                    (0.0, 0.25),
                    (0.6, 0.25),
                    (0.6, 0.0),
                    (1.0, 0.5),
                    (0.6, 1.0),
                    (0.6, 0.75),
                    (0.0, 0.75),
                ],
            },
            Some(Color::new(255, 136, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon for arrow in: {}",
        output.source
    );
    assert!(
        output.source.contains("300pt"),
        "Expected 300pt (arrow tip) in: {}",
        output.source
    );
}

#[test]
fn test_polygon_with_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            100.0,
            100.0,
            ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
            },
            None,
            Some(BorderSide {
                width: 2.0,
                color: Color::new(0, 0, 0),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon in: {}",
        output.source
    );
    assert!(
        output.source.contains("stroke: 2pt + rgb(0, 0, 0)"),
        "Expected stroke in: {}",
        output.source
    );
}
