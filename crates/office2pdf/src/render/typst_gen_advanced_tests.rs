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

// ── Metadata codegen tests ─────────────────────────────────────────

#[test]
fn test_generate_typst_with_metadata_title_and_author() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Test Title".to_string()),
            author: Some("Test Author".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Test Title\", author: \"Test Author\")"),
        "Expected document metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_title_only() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Only Title".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Only Title\")"),
        "Expected title-only metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_without_metadata() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        runs: vec![Run {
            text: "Hello".to_string(),
            style: TextStyle::default(),
            footnote: None,
            href: None,
        }],
        style: ParagraphStyle::default(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#set document("),
        "Should not emit #set document when no metadata, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_created_date() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Dated Doc".to_string()),
            created: Some("2024-06-15T10:30:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("date: datetime(year: 2024, month: 6, day: 15"),
        "Expected document date from metadata created field, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_date_only() {
    let doc = Document {
        metadata: Metadata {
            created: Some("2023-12-25T08:00:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("date: datetime(year: 2023, month: 12, day: 25"),
        "Expected document date even without title/author, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_invalid_created_date() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Bad Date Doc".to_string()),
            created: Some("not-a-date".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("date: datetime("),
        "Invalid date should not produce document date, got: {result}"
    );
}

#[test]
fn test_parse_iso8601_date_full() {
    let result = parse_iso8601_date("2024-06-15T10:30:45Z");
    assert_eq!(result, Some((2024, 6, 15, 10, 30, 45)));
}

#[test]
fn test_parse_iso8601_date_date_only() {
    let result = parse_iso8601_date("2023-12-25");
    assert_eq!(result, Some((2023, 12, 25, 0, 0, 0)));
}

#[test]
fn test_parse_iso8601_date_invalid() {
    assert_eq!(parse_iso8601_date("not-a-date"), None);
    assert_eq!(parse_iso8601_date(""), None);
    assert_eq!(parse_iso8601_date("2024"), None);
    assert_eq!(parse_iso8601_date("2024-13-01T00:00:00Z"), None);
    assert_eq!(parse_iso8601_date("2024-00-01T00:00:00Z"), None);
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

#[test]
fn test_font_substitution_calibri_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Calibri text".to_string(),
            style: TextStyle {
                font_family: Some("Calibri".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Calibri", "Carlito", "Liberation Sans")"#),
        "Expected font fallback list for Calibri in: {result}"
    );
}

#[test]
fn test_font_substitution_arial_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Arial text".to_string(),
            style: TextStyle {
                font_family: Some("Arial".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Arial", "Liberation Sans", "Arimo")"#),
        "Expected font fallback list for Arial in: {result}"
    );
}

#[test]
fn test_font_substitution_unknown_font_no_fallback() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Custom text".to_string(),
            style: TextStyle {
                font_family: Some("Helvetica".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: "Helvetica""#),
        "Unknown font should use simple quoted string in: {result}"
    );
    assert!(
        !result.contains("font: (\"Helvetica\""),
        "Unknown font should not use array syntax in: {result}"
    );
}

#[test]
fn test_font_substitution_times_new_roman() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "TNR text".to_string(),
            style: TextStyle {
                font_family: Some("Times New Roman".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Times New Roman", "Liberation Serif", "Tinos")"#),
        "Expected font fallback list for Times New Roman in: {result}"
    );
}

#[test]
fn test_font_family_infers_medium_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard Medium".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "medium""#),
        "Expected medium weight inferred from family name in: {result}"
    );
}

#[test]
fn test_font_family_infers_extrabold_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Heading".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "extrabold""#),
        "Expected extrabold weight inferred from family name in: {result}"
    );
}

#[test]
fn test_generate_typst_prefers_office_font_order_when_context_present() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Apple SD Gothic Neo", "Malgun Gothic"],
        &["Malgun Gothic"],
        &[],
    );

    let output = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap();

    let apple_index = output
        .source
        .find("\"Apple SD Gothic Neo\"")
        .expect("Apple SD Gothic Neo should appear in Typst output");
    let malgun_index = output
        .source
        .find("\"Malgun Gothic\"")
        .expect("Malgun Gothic should appear in Typst output");
    assert!(
        malgun_index < apple_index,
        "Office-resolved font ordering should win in Typst output: {}",
        output.source
    );
}

// --- Heading level codegen tests (US-096) ---

#[test]
fn test_generate_heading_level_1() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Main Title".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)[Main Title]"),
        "H1 paragraph should emit #heading(level: 1): {result}"
    );
}

#[test]
fn test_generate_heading_level_2() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(2),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Sub Section".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 2)[Sub Section]"),
        "H2 paragraph should emit #heading(level: 2): {result}"
    );
}

#[test]
fn test_generate_heading_levels_3_to_6() {
    for level in 3..=6u8 {
        let text = format!("Heading {level}");
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(level),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.clone(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        let expected = format!("#heading(level: {level})[{text}]");
        assert!(
            result.contains(&expected),
            "H{level} should emit {expected}: {result}"
        );
    }
}

#[test]
fn test_generate_heading_with_styled_run() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Styled Heading".to_string(),
            style: TextStyle {
                bold: Some(true),
                font_size: Some(24.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)"),
        "Heading with styling should still emit #heading: {result}"
    );
}

#[test]
fn test_generate_regular_paragraph_no_heading() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Normal text".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#heading"),
        "Regular paragraph should not emit #heading: {result}"
    );
}
