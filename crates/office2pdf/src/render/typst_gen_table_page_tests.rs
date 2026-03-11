use super::*;

#[test]
fn test_table_page_basic() {
    let table = make_simple_table(vec![vec!["A1", "B1"], vec!["A2", "B2"]]);
    let doc = make_doc(vec![make_sheet_page(
        "Sheet1",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#set page("));
    assert!(output.source.contains("#table("));
    assert!(output.source.contains("A1"));
    assert!(output.source.contains("B1"));
    assert!(output.source.contains("A2"));
    assert!(output.source.contains("B2"));
}

#[test]
fn test_table_page_custom_page_size_and_margins() {
    let table = make_simple_table(vec![vec!["Data"]]);
    let doc = make_doc(vec![make_sheet_page(
        "Custom",
        800.0,
        600.0,
        Margins {
            top: 20.0,
            bottom: 20.0,
            left: 30.0,
            right: 30.0,
        },
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("width: 800pt"));
    assert!(output.source.contains("height: 600pt"));
    assert!(output.source.contains("top: 20pt"));
    assert!(output.source.contains("left: 30pt"));
}

#[test]
fn test_table_page_cell_data_types() {
    let table = make_simple_table(vec![
        vec!["Name", "Age", "Date"],
        vec!["Alice", "30", "2024-01-15"],
    ]);
    let doc = make_doc(vec![make_sheet_page(
        "Data",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Name"));
    assert!(output.source.contains("Age"));
    assert!(output.source.contains("Date"));
    assert!(output.source.contains("Alice"));
    assert!(output.source.contains("30"));
    assert!(output.source.contains("2024-01-15"));
}

#[test]
fn test_table_page_merged_cells() {
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![TableCell {
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
                }],
                height: None,
            },
            TableRow {
                cells: vec![
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Left".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Right".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                ],
                height: None,
            },
        ],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_sheet_page(
        "MergeSheet",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("colspan: 2"));
    assert!(output.source.contains("Merged"));
    assert!(output.source.contains("Left"));
    assert!(output.source.contains("Right"));
}

#[test]
fn test_table_page_with_column_widths() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![
                TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Col1".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    ..TableCell::default()
                },
                TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Col2".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    ..TableCell::default()
                },
            ],
            height: None,
        }],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_sheet_page(
        "Widths",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("columns: (100pt, 200pt)"));
}

#[test]
fn test_table_page_empty_table() {
    let table = Table {
        rows: vec![],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_sheet_page(
        "Empty",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#set page("));
}

#[test]
fn test_table_page_multiple_sheets() {
    let table1 = make_simple_table(vec![vec!["Sheet1Data"]]);
    let table2 = make_simple_table(vec![vec!["Sheet2Data"]]);
    let doc = make_doc(vec![
        make_sheet_page("Sheet1", 595.28, 841.89, Margins::default(), table1),
        make_sheet_page("Sheet2", 595.28, 841.89, Margins::default(), table2),
    ]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Sheet1Data"));
    assert!(output.source.contains("Sheet2Data"));
}

#[test]
fn test_table_page_rowspan_merge() {
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![
                    TableCell {
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
                    },
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Top".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                ],
                height: None,
            },
            TableRow {
                cells: vec![TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Bottom".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    ..TableCell::default()
                }],
                height: None,
            },
        ],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_sheet_page(
        "RowMerge",
        595.28,
        841.89,
        Margins::default(),
        table,
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rowspan: 2"));
    assert!(output.source.contains("Tall"));
    assert!(output.source.contains("Top"));
    assert!(output.source.contains("Bottom"));
}
