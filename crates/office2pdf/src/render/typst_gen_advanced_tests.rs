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
                inline_box: None,
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
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        default_cell_padding: Some(Insets {
            top: 1.0,
            right: 3.0,
            bottom: 1.5,
            left: 3.0,
        }),
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear(rgb(99, 142, 198)"),
        "DataBar should be a gradient in the bar color. Got: {}",
        output.source,
    );
    assert!(
        output
            .source
            .contains("rgb(99, 142, 198).lighten(83%), space: rgb)"),
        "DataBar gradient should interpolate in sRGB. Got: {}",
        output.source,
    );
    // Excel's data-bar track has its own 2pt left / 1pt right inset, not the
    // cell text's 3pt per side. In a 100pt column the track is therefore
    // 97pt wide, and Excel quantises a 50% bar to 49 whole PDF points.
    assert!(
        output.source.contains("dx: -1pt"),
        "DataBar should start at Excel's track inset. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("width: 49pt"),
        "DataBar should resolve and quantise against Excel's track. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("#place("),
        "DataBar must be placed behind the value, not stacked above it. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("rgb(240, 240, 240)"),
        "Excel draws no gray track behind data bars. Got: {}",
        output.source,
    );
    // Auto-sized rows give the placed box no cell-bounded frame of
    // reference: a relative height resolves against the page and paints a
    // page-tall smear over neighboring rows (issue #362).
    assert!(
        !output.source.contains("height: 100%"),
        "DataBar height must not be page-relative in auto-height rows. Got: {}",
        output.source,
    );
    // The ramp's far end. Not a chosen number and not Excel's own endpoint
    // either: Excel's export fits a fade to 0.84, and 83% is the value that
    // reproduces it once our renderer's slightly lighter output is accounted
    // for. `write_table_cell` in typst_gen_tables.rs records the derivation.
    // Fading only 70% of the way left short bars reading near-solid (#654).
    assert!(
        output.source.contains(".lighten(83%)"),
        "DataBar gradient must fade to the endpoint that reproduces Excel's. Got: {}",
        output.source,
    );
}

#[test]
fn test_data_bar_fixed_row_height_codegen() {
    use crate::ir::DataBarInfo;

    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "75".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        data_bar: Some(DataBarInfo {
            color: Color::new(0x63, 0x8E, 0xC6),
            fill_pct: 75.0,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(24.0),
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // A 24pt row less Excel's 2pt clearance per side gives a concrete 20pt
    // bar instead of a height that resolves against the page.
    assert!(
        output.source.contains("height: 20pt"),
        "DataBar in a fixed-height row should be inset from the row edges. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("height: 100%"),
        "DataBar height must never be page-relative. Got: {}",
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
                inline_box: None,
            }],
        })],
        icon_text: Some("↑".to_string()),
        icon_color: Some(Color::new(214, 85, 50)),
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
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("↑"),
        "Icon text should appear in output. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("rgb(214, 85, 50)"),
        "Icon color should tint the icon glyph. Got: {}",
        output.source,
    );
    // Excel anchors the icon at the cell's left edge, on its own seat below
    // the row's top boundary. An in-flow icon wraps narrow cells onto a
    // second line, doubling the row height (issue #367) — the icon must be
    // placed out of layout. The `dx` and `dy` that carry it back out of the
    // cell's inset are this cell's own (issues #1087, #1202); which offsets
    // those are belongs to the tests below.
    assert!(
        output.source.contains("#place(top + left, dx: ")
            && output.source.contains("pt, dy: ")
            && output.source.contains("text("),
        "Icon must be placed out of layout at the cell's left edge. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains(")[↑] 90"),
        "Icon must not be an in-flow prefix of the value. Got: {}",
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
                inline_box: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![wide_cell],
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
                inline_box: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
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
                inline_box: None,
            }],
        })],
        col_span: 5,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![wide_cell],
                height: None,
            },
            TableRow {
                minimum_height: None,
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
                inline_box: None,
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
                inline_box: None,
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
            inline_box: None,
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
                inline_box: None,
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
                inline_box: None,
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
                inline_box: None,
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
                join: LineJoin::Round,
                cap: LineCap::Flat,
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
        output
            .source
            .contains("stroke: (paint: rgb(0, 0, 0), thickness: 2pt, join: \"round\")"),
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
            inline_box: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(
            r#"font: ("Calibri", "Carlito", "Liberation Sans", "Arimo", "DejaVu Sans", "Helvetica")"#
        ),
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
            inline_box: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(
            r#"font: ("Arial", "Liberation Sans", "Arimo", "Calibri", "DejaVu Sans", "Helvetica")"#
        ),
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
            inline_box: None,
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
            inline_box: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result
            .contains(r#"font: ("Times New Roman", "Liberation Serif", "Tinos", "DejaVu Serif")"#),
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
            inline_box: None,
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
            inline_box: None,
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
            inline_box: None,
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
            inline_box: None,
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
            inline_box: None,
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
                inline_box: None,
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
            inline_box: None,
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
            inline_box: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#heading"),
        "Regular paragraph should not emit #heading: {result}"
    );
}

#[test]
fn test_spill_continuation_redraws_the_line_shifted_left_and_clipped_at_the_gridline() {
    // The first cell of an overflow page-column carries the tail of a line
    // that began 150pt to its left. The clip box opens on the cell's own
    // gridline — the page-column edge Excel clips at — and runs the 40pt the
    // tail still reaches, while the line inside is moved left by the offset
    // less the inset the content box already sits behind. The placement is
    // gated on the laid-out line actually reaching that gridline, since the
    // spill width that put the continuation here is an estimate (issue
    // #1381).
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Corporate Security Supervisor".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        spill_width: Some(40.0),
        spill_continuation_offset_pt: Some(150.0),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![60.0],
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains(
            "if measure(o2p-spill).width > 145pt {place(left + horizon, dx: -5pt, dy: 0pt, \
             box(width: 40pt, height: 1.3em, clip: true)\
             [#move(dx: -145pt)[#box(width: measure(o2p-spill).width)[#o2p-spill]]])}}"
        ),
        "a continuation must open its clip on the gridline and shift the line back by \
         the width already printed. Got: {}",
        output.source
    );
}

#[test]
fn test_spill_width_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "spilling text".to_string(),
                style: TextStyle::default(),
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
            height: None,
        }],
        column_widths: vec![60.0],
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // The spill width less the left inset the box is anchored behind: the clip
    // ends on the cell's own gridline, not a whole inset past it (issue #1105).
    assert!(
        output.source.contains("width: 195pt"),
        "spilled cell must lay text out across the spill width. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("clip: true"),
        "spilled text must clip instead of wrapping. Got: {}",
        output.source,
    );
}

#[test]
fn test_table_default_vertical_align_codegen() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "value".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                        inline_box: None,
                    }],
                })],
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![100.0],
        default_vertical_align: Some(CellVerticalAlign::Bottom),
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("align: bottom"),
        "table-wide bottom alignment must be emitted. Got: {}",
        output.source,
    );
}

#[test]
fn test_vert_text_box_remaps_insets() {
    let text_box = TextBoxData {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "세로".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        padding: Insets {
            left: 7.2,
            right: 7.2,
            top: 3.6,
            bottom: 3.6,
        },
        vertical_align: TextBoxVerticalAlign::Top,
        fill: None,
        opacity: None,
        stroke: None,
        shape_kind: None,
        no_wrap: false,
        auto_fit: false,
        text_rotation_deg: Some(270.0),
        shape_rotation_deg: None,
    };
    let elem = FixedElement {
        x: 0.0,
        y: 0.0,
        width: 100.0,
        height: 50.0,
        kind: FixedElementKind::TextBox(text_box),
    };
    let page = Page::Fixed(FixedPage {
        size: PageSize::default(),
        elements: vec![elem],
        background_color: None,
        background_gradient: None,
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("inset: (top: 7.2pt, right: 3.6pt, bottom: 7.2pt, left: 3.6pt)"),
        "270° rotation must remap the bodyPr insets. Got: {}",
        output.source,
    );
}

/// Rows above a print-title range lead the table but must not repeat, so they
/// go into their own `repeat: false` header ahead of the repeating one.
#[test]
fn test_non_repeating_header_rows_emit_a_separate_header_block() {
    let row = |text: &str| TableRow {
        minimum_height: None,
        cells: vec![TableCell {
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
        }],
        height: Some(14.0),
    };
    let table = Table {
        rows: vec![row("Title"), row("Spacer"), row("SKU"), row("SKU-1000")],
        column_widths: vec![100.0],
        header_row_count: 1,
        non_repeating_header_row_count: 2,
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();

    let lead = output
        .source
        .find("table.header(repeat: false,")
        .expect("non-repeating header emitted");
    let repeating = output
        .source
        .find("table.header(level: 2,")
        .expect("repeating header emitted at a higher level");
    assert!(lead < repeating, "the non-repeating block comes first");
    let title = output.source.find("Title").expect("title row present");
    // Typst escapes the hyphen, so match on the numeric part.
    let sku = output.source.find("1000").expect("data row present");
    assert!(lead < title && title < repeating && repeating < sku);
}

/// Without non-repeating rows the header keeps its plain form.
#[test]
fn test_header_without_non_repeating_rows_stays_a_single_block() {
    let row = |text: &str| TableRow {
        minimum_height: None,
        cells: vec![TableCell {
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
        }],
        height: Some(14.0),
    };
    let table = Table {
        rows: vec![row("SKU"), row("SKU-1000")],
        column_widths: vec![100.0],
        header_row_count: 1,
        non_repeating_header_row_count: 0,
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("table.header(\n"));
    assert!(!output.source.contains("repeat: false"));
    assert!(!output.source.contains("level: 2"));
}

/// Excel insets data bars from the row's top and bottom edges instead of
/// filling the cell: native Excel PDFs print a 10 pt bar in a 14 pt row.
#[test]
fn test_data_bar_is_inset_from_the_row_edges() {
    use crate::ir::DataBarInfo;

    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "120".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        data_bar: Some(DataBarInfo {
            color: Color::new(0x63, 0x8E, 0xC6),
            fill_pct: 24.0,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: Some(14.0),
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("height: 10pt"),
        "a 14pt row must print a 10pt bar. Got: {}",
        output.source,
    );
}

/// The inset never collapses the bar in very short rows.
#[test]
fn test_data_bar_in_a_short_row_keeps_a_visible_height() {
    use crate::ir::DataBarInfo;

    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "3".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
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
            minimum_height: None,
            cells: vec![cell],
            height: Some(3.0),
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("height: 0pt"),
        "the bar must stay visible. Got: {}",
        output.source,
    );
    assert!(!output.source.contains("height: -"));
}

fn icon_cell(icon: &str, color: Color) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "107%".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        icon_text: Some(icon.to_string()),
        icon_color: Some(color),
        ..TableCell::default()
    }
}

fn icon_sheet(cell: TableCell) -> Page {
    Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells: vec![cell],
                height: Some(14.0),
            }],
            column_widths: vec![100.0],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })
}

/// Excel's arrow icon sets are drawn shapes, not characters, so an arrow icon
/// must render as a filled polygon in its band color.
#[test]
fn test_arrow_icon_set_renders_as_a_filled_polygon() {
    let doc = make_doc(vec![icon_sheet(icon_cell(
        crate::ir::ICON_ARROW_UP,
        Color::new(0x68, 0xA4, 0x90),
    ))]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("polygon(fill: rgb(104, 164, 144)"),
        "the arrow must be a polygon in the band color. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("darken(30%)"),
        "Excel outlines the arrow a shade darker"
    );
    assert!(
        !output.source.contains(crate::ir::ICON_ARROW_UP),
        "the triangle character must not also be drawn"
    );
}

/// Every `polygon(...)` the generated source draws, in emission order, as its
/// own points. They are emitted as `(Xpt, Ypt)`; pull them out without a regex
/// dependency, stopping at the first chunk past a polygon's coordinate list.
fn polygon_point_sets(source: &str) -> Vec<Vec<(f64, f64)>> {
    source
        .split("polygon(")
        .skip(1)
        .map(|call| {
            let mut points: Vec<(f64, f64)> = Vec::new();
            for chunk in call.split('(') {
                let point: Option<(f64, f64)> = chunk.split_once("pt, ").and_then(|(x, rest)| {
                    let (y, _) = rest.split_once("pt)")?;
                    Some((x.trim().parse().ok()?, y.parse().ok()?))
                });
                match point {
                    Some(point) => points.push(point),
                    None if points.is_empty() => continue,
                    None => break,
                }
            }
            points
        })
        .filter(|points| !points.is_empty())
        .collect()
}

/// Read an arrow's silhouette back out of the generated source: the filled
/// polygon, which is emitted before the outline ring stroked inside it.
fn arrow_polygon_points(source: &str) -> Vec<(f64, f64)> {
    let sets: Vec<Vec<(f64, f64)>> = polygon_point_sets(source);
    assert!(!sets.is_empty(), "no polygon points in: {source}");
    sets.into_iter().next().expect("checked non-empty")
}

/// Read the path the arrow's outline is stroked on (issue #1201).
fn arrow_outline_points(source: &str) -> Vec<(f64, f64)> {
    let sets: Vec<Vec<(f64, f64)>> = polygon_point_sets(source);
    assert_eq!(
        sets.len(),
        2,
        "an arrow is one filled silhouette under one outline ring: {source}"
    );
    sets.into_iter().nth(1).expect("checked length")
}

/// The layout box the icon's shapes are superimposed in, read back out of the
/// generated source. Every document's preamble carries a `box(width: 1fr, …)`
/// of its own, so take the first box measured in points.
fn icon_box_size(source: &str) -> (f64, f64) {
    source
        .match_indices("box(width: ")
        .find_map(|(index, marker)| {
            let (width, rest) = source[index + marker.len()..].split_once("pt, height: ")?;
            let (height, _) = rest.split_once("pt)")?;
            Some((width.parse().ok()?, height.parse().ok()?))
        })
        .unwrap_or_else(|| panic!("no icon box in: {source}"))
}

/// The silhouette an arrow polygon spans: its breadth across the shaft, its
/// length along the tip, the shaft's width and the head's length. Works for
/// either vertical orientation — the tip is the lone point at one end of the
/// length axis, the shaft's base the pair at the other, and the barbs the row
/// carrying both breadth extremes.
fn arrow_silhouette(points: &[(f64, f64)]) -> (f64, f64, f64, f64) {
    let min_x: f64 = points.iter().map(|p| p.0).fold(f64::MAX, f64::min);
    let max_x: f64 = points.iter().map(|p| p.0).fold(f64::MIN, f64::max);
    let min_y: f64 = points.iter().map(|p| p.1).fold(f64::MAX, f64::min);
    let max_y: f64 = points.iter().map(|p| p.1).fold(f64::MIN, f64::max);
    let row = |y: f64| -> Vec<(f64, f64)> {
        points
            .iter()
            .copied()
            .filter(|p| (p.1 - y).abs() < 1e-9)
            .collect()
    };
    let (tip_y, base) = if row(min_y).len() == 1 {
        (min_y, row(max_y))
    } else {
        (max_y, row(min_y))
    };
    assert_eq!(base.len(), 2, "the shaft's base is two points: {base:?}");
    let barb_y: f64 = points
        .iter()
        .find(|p| (p.0 - min_x).abs() < 1e-9)
        .expect("a barb reaches the breadth's left edge")
        .1;
    (
        max_x - min_x,
        max_y - min_y,
        (base[0].0 - base[1].0).abs(),
        (barb_y - tip_y).abs(),
    )
}

/// The arrow's shaft and head take Excel's fractions of the icon box.
///
/// The native export's sprites carry a soft mask, which is the silhouette
/// itself. The up arrow's is a 12 x 12px bitmap with 11 x 12px of ink: the
/// head occupies rows 0-5, half the length, and the shaft columns 3-7, 5 of
/// the 11 ink columns. We used to give the shaft 0.56 of the breadth under a
/// head 0.45 of the length — 23% too wide and 10% too short (issue #1135).
#[test]
fn test_arrow_icon_silhouette_matches_excels_sprite_mask() {
    for glyph in [crate::ir::ICON_ARROW_UP, crate::ir::ICON_ARROW_DOWN] {
        let output = generate_typst(&make_doc(vec![icon_sheet(icon_cell(
            glyph,
            Color::new(0x68, 0xA4, 0x90),
        ))]))
        .unwrap();
        let (breadth, length, shaft_width, head_length) =
            arrow_silhouette(&arrow_polygon_points(&output.source));
        assert!(
            (shaft_width / breadth - 5.0 / 11.0).abs() < 1e-6,
            "{glyph}: the shaft spans 5 of the mask's 11 ink columns, \
             got {shaft_width}pt of {breadth}pt",
        );
        assert!(
            (head_length / length - 6.0 / 12.0).abs() < 1e-6,
            "{glyph}: the head takes half the arrow's length, \
             got {head_length}pt of {length}pt",
        );
    }
}

/// Excel's own paint for the green `3Arrows` band, read off the sprite in the
/// native export of `10_kpi_tracker_en`.
fn green_arrow_shading() -> crate::ir::IconShading {
    crate::ir::IconShading {
        fill_start: Color::new(0x9F, 0xD8, 0xAE),
        fill_end: Color::new(0x28, 0xA5, 0x4A),
        outline: Color::new(0x25, 0x5E, 0x1B),
    }
}

/// A band Excel shades carries its measured ramp and its own outline colour,
/// not a flat fill under a `darken(30%)` derivative of it (issue #1134).
#[test]
fn test_shaded_arrow_icon_takes_excels_ramp_and_outline() {
    let mut cell = icon_cell(crate::ir::ICON_ARROW_UP, Color::new(0x59, 0xB0, 0x6D));
    cell.icon_shading = Some(green_arrow_shading());
    let output = generate_typst(&make_doc(vec![icon_sheet(cell)])).unwrap();
    assert!(
        output.source.contains(
            "polygon(fill: gradient.linear(angle: 45deg, space: rgb, \
             rgb(159, 216, 174), rgb(40, 165, 74)), stroke: none"
        ),
        "the icon must ramp along the box diagonal. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("pt + rgb(37, 94, 27)"),
        "and carry Excel's own outline hue. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("darken(30%)"),
        "a measured outline must not also be derived from the fill. Got: {}",
        output.source,
    );
}

/// The perpendicular distance each ring edge sits inside the silhouette edge
/// it follows, both endpoints reported separately so a corner that mitered the
/// wrong way shows up.
fn outline_inset_per_edge(silhouette: &[(f64, f64)], ring: &[(f64, f64)]) -> Vec<f64> {
    let count: usize = silhouette.len();
    let twice_area: f64 = (0..count)
        .map(|index| {
            let (x0, y0) = silhouette[index];
            let (x1, y1) = silhouette[(index + 1) % count];
            x0 * y1 - x1 * y0
        })
        .sum();
    let winding: f64 = if twice_area > 0.0 { 1.0 } else { -1.0 };
    (0..count)
        .flat_map(|index| {
            let (x0, y0) = silhouette[index];
            let (x1, y1) = silhouette[(index + 1) % count];
            let (dx, dy) = (x1 - x0, y1 - y0);
            let length: f64 = dx.hypot(dy);
            let normal: (f64, f64) = (-dy / length * winding, dx / length * winding);
            [ring[index], ring[(index + 1) % count]]
                .map(|(x, y)| (x - x0) * normal.0 + (y - y0) * normal.1)
        })
        .collect()
}

/// Excel's sprite draws the arrow's outline one whole sprite pixel wide and
/// wholly inside the silhouette: the up arrow's bottom row is the outline hue
/// across, its shaft's side columns likewise, and the interior ramp starts one
/// pixel in. The sprite is 12px in an 11pt box, so that pixel is 0.92pt.
///
/// A Typst stroke is centred on its path, so a 0.4pt stroke on the silhouette
/// itself painted 0.2pt of outline inward and spilled 0.2pt over the page,
/// where Excel paints none (issue #1201). Stroking the full width on a path
/// inset by half of it puts the stroke's outer edge back on the silhouette.
#[test]
fn test_arrow_icon_outline_is_stroked_inside_the_silhouette() {
    // One sprite pixel: the 12px bitmap prints in Excel's 11pt box.
    let width: f64 = 11.0 / 12.0;
    for glyph in [
        crate::ir::ICON_ARROW_UP,
        crate::ir::ICON_ARROW_DOWN,
        crate::ir::ICON_ARROW_RIGHT,
    ] {
        let mut cell = icon_cell(glyph, Color::new(0x59, 0xB0, 0x6D));
        cell.icon_shading = Some(green_arrow_shading());
        let source: String = generate_typst(&make_doc(vec![icon_sheet(cell)]))
            .unwrap()
            .source;

        let stroke: String = format!("stroke: {width}pt + rgb(37, 94, 27)");
        assert!(
            source.contains(&stroke),
            "{glyph}: the outline is one sprite pixel wide. Want {stroke} in: {source}",
        );
        let silhouette: Vec<(f64, f64)> = arrow_polygon_points(&source);
        let ring: Vec<(f64, f64)> = arrow_outline_points(&source);
        assert_eq!(
            ring.len(),
            silhouette.len(),
            "{glyph}: the ring follows every silhouette edge: {source}",
        );
        for (edge, inset) in outline_inset_per_edge(&silhouette, &ring)
            .into_iter()
            .enumerate()
        {
            assert!(
                (inset - width / 2.0).abs() < 1e-9,
                "{glyph}: edge {} sits {inset}pt inside the silhouette, \
                 so the stroke's outer edge misses it by {}pt",
                edge / 2,
                inset - width / 2.0,
            );
        }
    }
}

/// The arrow's layout box is Excel's 11 x 11pt sprite box, whatever the
/// silhouette in it spans.
///
/// The box is what the row seats (issue #1202), and the silhouette is flush
/// with its top-left corner because the mask's padding column falls on the
/// right and the transposed one's padding row at the bottom. An outline
/// stroked outside the silhouette would therefore leave the sprite's own
/// extent, which is what #1201 kept it inside of.
#[test]
fn test_arrow_icon_box_is_excels_sprite_box() {
    for glyph in [
        crate::ir::ICON_ARROW_UP,
        crate::ir::ICON_ARROW_DOWN,
        crate::ir::ICON_ARROW_RIGHT,
    ] {
        let source: String = generate_typst(&make_doc(vec![icon_sheet(icon_cell(
            glyph,
            Color::new(0x68, 0xA4, 0x90),
        ))]))
        .unwrap()
        .source;
        assert_eq!(
            icon_box_size(&source),
            (11.0, 11.0),
            "{glyph}: the box is the sprite's own. Got: {source}",
        );
        let silhouette: Vec<(f64, f64)> = arrow_polygon_points(&source);
        for axis in [0, 1] {
            let start: f64 = silhouette
                .iter()
                .map(|point| if axis == 0 { point.0 } else { point.1 })
                .fold(f64::MAX, f64::min);
            assert_eq!(
                start, 0.0,
                "{glyph}: the ink is flush with the box on axis {axis}. Got: {source}",
            );
        }
    }
}

/// The ramp runs down the box diagonal for every orientation: Excel shades the
/// sprite, not the arrow, so flipping or transposing the silhouette leaves the
/// light corner at the top left.
#[test]
fn test_every_shaded_arrow_orientation_ramps_along_the_same_diagonal() {
    for glyph in [
        crate::ir::ICON_ARROW_UP,
        crate::ir::ICON_ARROW_DOWN,
        crate::ir::ICON_ARROW_RIGHT,
    ] {
        let mut cell = icon_cell(glyph, Color::new(0x59, 0xB0, 0x6D));
        cell.icon_shading = Some(green_arrow_shading());
        let output = generate_typst(&make_doc(vec![icon_sheet(cell)])).unwrap();
        assert!(
            output.source.contains(
                "gradient.linear(angle: 45deg, space: rgb, rgb(159, 216, 174), rgb(40, 165, 74))"
            ),
            "{glyph}: the ramp must keep its direction and stops. Got: {}",
            output.source,
        );
    }
}

/// A band with no measured shading keeps the flat stand-in, so the sets with
/// no native export to read are unaffected.
#[test]
fn test_unshaded_arrow_icon_keeps_the_flat_fill() {
    let output = generate_typst(&make_doc(vec![icon_sheet(icon_cell(
        crate::ir::ICON_ARROW_UP,
        Color::new(0x68, 0xA4, 0x90),
    ))]))
    .unwrap();
    assert!(
        output
            .source
            .contains("polygon(fill: rgb(104, 164, 144), stroke: none"),
        "an unmeasured band still fills flat. Got: {}",
        output.source,
    );
    assert!(
        output
            .source
            .contains("pt + rgb(104, 164, 144).darken(30%)"),
        "under an outline derived from that fill. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("gradient"),
        "and must not invent a ramp. Got: {}",
        output.source,
    );
}

/// A down arrow points the other way, so its tip sits at the bottom.
#[test]
fn test_down_arrow_icon_is_flipped() {
    let up = generate_typst(&make_doc(vec![icon_sheet(icon_cell(
        crate::ir::ICON_ARROW_UP,
        Color::new(0x68, 0xA4, 0x90),
    ))]))
    .unwrap();
    let down = generate_typst(&make_doc(vec![icon_sheet(icon_cell(
        crate::ir::ICON_ARROW_DOWN,
        Color::new(0xD6, 0x55, 0x32),
    ))]))
    .unwrap();
    // Read the polygon back rather than pinning literal coordinates: the arrow
    // is sized from a GT measurement and those numbers move (issue #651). What
    // must hold is that flipping puts the tip at the other end of the same
    // shape.
    let tip = |source: &str, want_min_y: bool| -> (f64, f64) {
        let points: Vec<(f64, f64)> = arrow_polygon_points(source);
        let extreme = if want_min_y {
            points.iter().map(|p| p.1).fold(f64::MAX, f64::min)
        } else {
            points.iter().map(|p| p.1).fold(f64::MIN, f64::max)
        };
        let row: Vec<(f64, f64)> = points
            .into_iter()
            .filter(|p| (p.1 - extreme).abs() < 1e-6)
            .collect();
        assert_eq!(row.len(), 1, "the tip is the only point at that end");
        row[0]
    };

    let up_tip = tip(&up.source, true);
    let down_tip = tip(&down.source, false);
    assert_eq!(up_tip.1, 0.0, "the up arrow's tip is at the top");
    assert!(
        down_tip.1 > up_tip.1,
        "the down arrow's tip is at the other end: {down_tip:?} against {up_tip:?}"
    );
    assert!(
        (down_tip.0 - up_tip.0).abs() < 1e-6,
        "and on the same centre line: {down_tip:?} against {up_tip:?}"
    );
}

/// Icon sets that are not arrows keep their character rendering.
#[test]
fn test_non_arrow_icon_set_still_renders_as_text() {
    // Symbols, flags and stars have no drawn shape and stay characters; the
    // circles left this group when they became discs (#536).
    let doc = make_doc(vec![icon_sheet(icon_cell(
        "✓",
        Color::new(0xD6, 0x55, 0x32),
    ))]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("text(fill: rgb(214, 85, 50)"));
    assert!(!output.source.contains("polygon("));
    assert!(!output.source.contains("circle("));
}

/// Excel draws the traffic-light sets as filled discs, not a `●` character
/// at roughly half the diameter (issue #536).
#[test]
fn test_circle_icon_set_renders_as_a_filled_disc() {
    let doc = make_doc(vec![icon_sheet(icon_cell(
        crate::ir::ICON_CIRCLE,
        Color::new(0x62, 0xC1, 0x7A),
    ))]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("circle(radius: 4.48pt, fill: rgb(98, 193, 122)"),
        "the disc spans Excel's 8.96pt icon box in the band colour. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains(crate::ir::ICON_CIRCLE),
        "the character must not also be drawn"
    );
}

/// The inset an XLSX icon-set cell carries: the sheet's own 3pt text inset
/// plus the value reserve the icon's advance costs (issue #652).
fn icon_cell_inset() -> Insets {
    Insets {
        top: 1.0,
        right: 3.0,
        bottom: 1.5,
        left: 12.6,
    }
}

/// Excel anchors an icon-set icon at the cell's own left inset, not at the
/// text box the value reserve pushed inward.
///
/// Measured on the native export of `10_kpi_tracker_en` (issue #1087): column
/// E spans x 384-456pt and every `3Arrows` sprite is placed
/// `transform="11 0 0 11 386 …"`, so the icon starts 2pt inside the cell's
/// left boundary. Extracting the sprites puts their ink flush with that box's
/// left edge — the up and down arrows trim to 11x12 of 12x12 at +0+0 — so the
/// drawn polygon's own left edge belongs at 386.0. Our reserve leaves the
/// content box 12.6pt in, hence the 10.6pt pull-back.
#[test]
fn test_icon_set_icon_anchors_at_the_cells_left_inset() {
    let mut cell = icon_cell(crate::ir::ICON_ARROW_UP, Color::new(0x59, 0xB0, 0x6D));
    cell.padding = Some(icon_cell_inset());
    let output = generate_typst(&make_doc(vec![icon_sheet(cell)])).unwrap();
    assert!(
        output
            .source
            .contains("#place(top + left, dx: -10.6pt, dy: 0pt, box("),
        "the icon must be pulled back to 2pt inside the cell boundary. Got: {}",
        output.source,
    );
}

/// The pull-back is the cell's own inset less Excel's 2pt icon inset, not a
/// fixed offset: a cell laid out with a different inset needs a different
/// correction to land the icon in the same place.
#[test]
fn test_icon_anchor_follows_the_cells_own_inset() {
    let bare_inset = Insets {
        left: 3.0,
        ..icon_cell_inset()
    };
    let mut cell = icon_cell(crate::ir::ICON_ARROW_UP, Color::new(0x59, 0xB0, 0x6D));
    cell.padding = Some(bare_inset);
    let output = generate_typst(&make_doc(vec![icon_sheet(cell)])).unwrap();
    assert!(
        output
            .source
            .contains("#place(top + left, dx: -1pt, dy: 0pt, box("),
        "a 3pt inset leaves only 1pt to pull back. Got: {}",
        output.source,
    );
}

/// The anchor is the icon's, not the polygon's: the sets that stay characters
/// and the drawn discs sit at the same left inset.
#[test]
fn test_character_and_disc_icons_share_the_left_inset_anchor() {
    for glyph in ["✓", crate::ir::ICON_CIRCLE] {
        let mut cell = icon_cell(glyph, Color::new(0xD6, 0x55, 0x32));
        cell.padding = Some(icon_cell_inset());
        let output = generate_typst(&make_doc(vec![icon_sheet(cell)])).unwrap();
        assert!(
            output
                .source
                .contains("#place(top + left, dx: -10.6pt, dy: 0pt,"),
            "{glyph} must anchor at the cell's left inset too. Got: {}",
            output.source,
        );
    }
}

/// The alignment, `dx` and `dy` of the `#place` an icon is drawn by.
///
/// Only the icon's placement carries offsets in these fixtures, so the first
/// `#place(<align>, dx: …` in the source is it.
fn icon_placement(source: &str) -> (String, f64, f64) {
    source
        .match_indices("#place(")
        .find_map(|(index, marker)| {
            let call: &str = &source[index + marker.len()..];
            let (align, rest) = call.split_once(", dx: ")?;
            let (dx, rest) = rest.split_once("pt, ")?;
            let (dy, _) = rest.strip_prefix("dy: ")?.split_once("pt, ")?;
            Some((align.to_string(), dx.parse().ok()?, dy.parse().ok()?))
        })
        .unwrap_or_else(|| panic!("no placed icon in: {source}"))
}

/// The icon's own top edge, measured down from its row's top boundary: the
/// cell's inset, the placement's `dy`, and where in the drawn shape the ink
/// starts.
fn icon_top_below_row_boundary(source: &str, inset: Insets, ink_top_in_shape: f64) -> f64 {
    let (align, _, dy) = icon_placement(source);
    assert_eq!(
        align, "top + left",
        "the icon is seated from the row's top edge, not centred on it: {source}",
    );
    inset.top + dy + ink_top_in_shape
}

/// Excel seats an icon-set icon's sprite box a whole point below its row's
/// top boundary, whatever the row's height.
///
/// Measured on the native export of `10_kpi_tracker_en` (issue #1202): the
/// `KPI` sheet's six `3Arrows` sprites are placed
/// `transform="11 0 0 11 386 <y>"` with y = 133, 147, 161, 175, 189 and 203,
/// while the thin row rules under them fill 132-133, 146-147, … — so every
/// box starts 1.00pt below the boundary its rule is anchored on. The 14pt
/// tracks do not centre an 11pt box there; that would put it at 1.5pt.
#[test]
fn test_icon_sprite_box_seats_one_point_below_the_rows_top_boundary() {
    for glyph in [
        crate::ir::ICON_ARROW_UP,
        crate::ir::ICON_ARROW_DOWN,
        crate::ir::ICON_ARROW_RIGHT,
    ] {
        let mut cell = icon_cell(glyph, Color::new(0x59, 0xB0, 0x6D));
        cell.padding = Some(icon_cell_inset());
        let source: String = generate_typst(&make_doc(vec![icon_sheet(cell)]))
            .unwrap()
            .source;
        // Every arrow's ink starts at the top of the box it is drawn in: the
        // sprite's blank row falls at the bottom, never above the arrow.
        let ink_top: f64 = arrow_polygon_points(&source)
            .iter()
            .map(|point| point.1)
            .fold(f64::MAX, f64::min);
        let top: f64 = icon_top_below_row_boundary(&source, icon_cell_inset(), ink_top);
        assert!(
            (top - 1.0).abs() < 1e-9,
            "{glyph}: the icon starts {top}pt below the row's top edge, not 1pt. Got: {source}",
        );
    }
}

/// The seat is measured from the row boundary, so a cell laid out with a
/// different vertical inset needs a different `dy` to land the icon in the
/// same place — as the left anchor already does horizontally (issue #1087).
#[test]
fn test_icon_seat_follows_the_cells_own_vertical_inset() {
    for inset_top in [0.0, 1.0, 2.5] {
        let inset = Insets {
            top: inset_top,
            ..icon_cell_inset()
        };
        let mut cell = icon_cell(crate::ir::ICON_ARROW_UP, Color::new(0x59, 0xB0, 0x6D));
        cell.padding = Some(inset);
        let source: String = generate_typst(&make_doc(vec![icon_sheet(cell)]))
            .unwrap()
            .source;
        let ink_top: f64 = arrow_polygon_points(&source)
            .iter()
            .map(|point| point.1)
            .fold(f64::MAX, f64::min);
        let top: f64 = icon_top_below_row_boundary(&source, inset, ink_top);
        assert!(
            (top - 1.0).abs() < 1e-9,
            "a {inset_top}pt inset must still leave the icon 1pt below the boundary, \
             got {top}pt. Source: {source}",
        );
    }
}

/// A sprite shorter than its box keeps its ink against the box's top instead
/// of being re-centred in it.
///
/// The up and down arrows ink all 12 rows of the 12 x 12px bitmap; the right
/// one inks rows 0-10 and leaves row 11 blank, so its ink is 10.08pt in an
/// 11pt box and hangs from the top. Centring that shorter shape on the row
/// dropped it a further (11.00 - 10.08) / 2 = 0.46pt, for 0.71pt in total
/// (issue #1202).
#[test]
fn test_short_arrow_icon_hangs_from_its_sprite_boxs_top() {
    let mut cell = icon_cell(crate::ir::ICON_ARROW_RIGHT, Color::new(0x59, 0xB0, 0x6D));
    cell.padding = Some(icon_cell_inset());
    let source: String = generate_typst(&make_doc(vec![icon_sheet(cell)]))
        .unwrap()
        .source;
    let points: Vec<(f64, f64)> = arrow_polygon_points(&source);
    let ink_height: f64 = points.iter().map(|point| point.1).fold(f64::MIN, f64::max);
    assert!(
        (ink_height - 10.08).abs() < 1e-9,
        "the right arrow's ink is 0.92pt short of the box: {source}",
    );
    let (_, box_height) = icon_box_size(&source);
    assert!(
        (box_height - 11.0).abs() < 1e-9,
        "and is drawn in the sprite's own 11pt box, got {box_height}pt: {source}",
    );
    let top: f64 = icon_top_below_row_boundary(&source, icon_cell_inset(), 0.0);
    assert!(
        (top - 1.0).abs() < 1e-9,
        "so the box's top, not the ink's centre, is what the row seats: {top}pt",
    );
}

/// A disc has no native export to read, so it keeps the middle of the sprite
/// box Excel seats — the flush-top ink is a measured property of the arrow
/// masks, not of the box (issue #1202).
#[test]
fn test_disc_icon_centres_in_the_sprite_box() {
    let mut cell = icon_cell(crate::ir::ICON_CIRCLE, Color::new(0x62, 0xC1, 0x7A));
    cell.padding = Some(icon_cell_inset());
    let source: String = generate_typst(&make_doc(vec![icon_sheet(cell)]))
        .unwrap()
        .source;
    assert!(
        source.contains("box(height: 11pt)[#place(left + horizon, circle(radius: 4.48pt"),
        "the disc sits in the middle of the 11pt sprite box. Got: {source}",
    );
    let top: f64 = icon_top_below_row_boundary(&source, icon_cell_inset(), 0.0);
    assert!(
        (top - 1.0).abs() < 1e-9,
        "whose own top is still 1pt below the row boundary: {top}pt",
    );
}

/// A worksheet text box anchored after `anchor_row` at `x_offset_pt`.
fn make_sheet_text_box(anchor_row: u32, x_offset_pt: f64, height: f64) -> crate::ir::SheetTextBox {
    crate::ir::SheetTextBox {
        anchor_row,
        x_offset_pt,
        // Rows are 20pt in these fixtures, so the anchor row's top edge is
        // the summed height of the rows above it.
        y_offset_pt: f64::from(anchor_row - 1) * 20.0,
        width: 100.0,
        height,
        paragraphs: vec![Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: format!("shape at {x_offset_pt}"),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }],
        }],
        fill: None,
        border: None,
        vertical_center: false,
    }
}

fn sheet_page_with_text_boxes(text_boxes: Vec<crate::ir::SheetTextBox>) -> Page {
    Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells: vec![TableCell::default()],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes,
        shapes: Vec::new(),
    })
}

#[test]
fn test_sheet_drawings_overlay_the_grid_at_absolute_offsets() {
    // Excel floats drawings over the cells at absolute worksheet
    // coordinates. Reserving flow height per drawing stacked same-row shapes
    // diagonally and spilled the sheet onto a blank page (issue #459), and
    // even one reserved box per row could not match Excel's vertical
    // placement because our printed row heights differ from its print grid
    // (issue #474). All drawings are placed from the sheet's content origin
    // instead, in a page foreground that reserves no flow height at all.
    let doc = make_doc(vec![sheet_page_with_text_boxes(vec![
        make_sheet_text_box(3, 0.0, 60.0),
        make_sheet_text_box(3, 200.0, 60.0),
        make_sheet_text_box(3, 400.0, 60.0),
    ])]);
    let source = generate_typst(&doc).unwrap().source;

    // The sheet's content origin is its top-left margin corner, and the
    // foreground's offsets are measured from the page corner instead, so the
    // margins are what the two coordinate systems differ by.
    let margin: f64 = crate::defaults::DEFAULT_MARGIN_PT;
    assert_eq!(
        source
            .matches("#block(width: 100%, height: 0pt, spacing: 0pt)")
            .count(),
        1,
        "the drawing layer is pinned by one zero-height marker: {source}"
    );
    assert!(
        !source.contains("#box(width: 100%, height: 60pt)"),
        "no drawing may reserve flow height: {source}"
    );
    // A `box` is inline, so the paragraph holding it still lays out a line box
    // and the grid below drops by a whole line — 13.2pt of Typst's default
    // 11pt text, regardless of the sheet's own font (issue #1101). Only a
    // block-level container carries the marker without a strut.
    assert!(
        !source.contains("#box(width: 100%, height: 0pt)"),
        "the marker must not be inline content: {source}"
    );
    // Row 3's top edge is two 20pt rows down.
    assert_eq!(
        source.matches(&format!("dy: {}pt", margin + 40.0)).count(),
        3,
        "same-row drawings share one vertical origin: {source}"
    );
    for x_offset in [0.0, 200.0, 400.0] {
        let dx: String = format!("dx: {}pt", margin + x_offset);
        assert!(
            source.contains(&dx),
            "each drawing keeps its own horizontal offset ({dx}): {source}"
        );
    }
}

/// A custom geometry's subpaths are one path under one fill rule, so an inner
/// boundary carves a hole rather than painting solid. Filling each subpath
/// independently painted the frame in the deck on #870 as a blob.
#[test]
fn test_multi_subpath_shape_fills_even_odd() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            100.0,
            100.0,
            ShapeKind::Path {
                subpaths: vec![
                    crate::ir::Subpath::closed_outline(vec![
                        (0.0, 0.0),
                        (1.0, 0.0),
                        (1.0, 1.0),
                        (0.0, 1.0),
                    ]),
                    crate::ir::Subpath::closed_outline(vec![
                        (0.25, 0.25),
                        (0.75, 0.25),
                        (0.75, 0.75),
                        (0.25, 0.75),
                    ]),
                ],
            },
            Some(Color::new(0, 128, 0)),
            None,
        )],
    )]);
    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("#curve("),
        "a multi-subpath shape draws as a curve: {source}"
    );
    assert!(
        source.contains("fill-rule: \"even-odd\""),
        "the inner boundary must carve a hole: {source}"
    );
    assert_eq!(
        source.matches("curve.move(").count(),
        2,
        "one move per subpath: {source}"
    );
    assert!(
        source.contains("fill: rgb(0, 128, 0)"),
        "the fill still applies: {source}"
    );
}

/// An open subpath is drawn without `curve.close()`, so a stroked polyline
/// stops at its last point. The elbow connectors of the deck on issue #1205
/// are unclosed `moveTo lnTo lnTo lnTo` paths, and closing them drew a
/// diagonal back across the slide.
#[test]
fn test_open_subpath_is_not_closed() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            100.0,
            100.0,
            ShapeKind::Path {
                subpaths: vec![crate::ir::Subpath::open_outline(vec![
                    (0.0, 0.0),
                    (0.0, 1.0),
                    (1.0, 1.0),
                ])],
            },
            None,
            None,
        )],
    )]);
    let source = generate_typst(&doc).unwrap().source;

    assert_eq!(source.matches("curve.move(").count(), 1, "got {source}");
    assert!(
        !source.contains("curve.close()"),
        "an unclosed path must not be closed: {source}"
    );
}

/// A single-subpath geometry still draws as one closed outline, so the change
/// does not turn every custom shape into a hole-carving path.
#[test]
fn test_single_subpath_shape_still_closes_its_outline() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            100.0,
            100.0,
            ShapeKind::Path {
                subpaths: vec![crate::ir::Subpath::closed_outline(vec![
                    (0.0, 0.0),
                    (1.0, 0.0),
                    (1.0, 1.0),
                ])],
            },
            Some(Color::new(0, 0, 255)),
            None,
        )],
    )]);
    let source = generate_typst(&doc).unwrap().source;

    assert_eq!(source.matches("curve.move(").count(), 1);
    assert!(source.contains("curve.close()"), "got {source}");
}
