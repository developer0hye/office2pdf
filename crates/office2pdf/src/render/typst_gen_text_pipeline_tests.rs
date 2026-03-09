use super::*;

// ── Unicode NFC normalization tests ──────────────────────────────

#[test]
fn test_escape_typst_normalizes_korean_nfd_to_nfc() {
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let nfc_korean = "한글";
    let result = escape_typst(nfd_korean);
    assert_eq!(
        result, nfc_korean,
        "NFD Korean jamo should be normalized to composed hangul"
    );
}

#[test]
fn test_escape_typst_normalizes_combining_diacritics() {
    let nfd_cafe = "cafe\u{0301}";
    let nfc_cafe = "caf\u{00E9}";
    let result = escape_typst(nfd_cafe);
    assert_eq!(
        result, nfc_cafe,
        "Combining diacritics should be normalized to NFC"
    );
}

#[test]
fn test_escape_typst_nfc_with_special_chars() {
    let nfd_input = "cafe\u{0301} \\$5";
    let result = escape_typst(nfd_input);
    assert!(
        result.contains("caf\u{00E9}"),
        "Should contain NFC-normalized é: {result}"
    );
    assert!(
        result.contains("\\$"),
        "Should still escape $ sign: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_korean_in_paragraph() {
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_korean)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("한글"),
        "Generated Typst should contain NFC-composed Korean: {result}"
    );
    assert!(
        !result.contains('\u{1112}'),
        "Generated Typst should not contain decomposed jamo: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_diacritics_in_paragraph() {
    let nfd_resume = "re\u{0301}sume\u{0301}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_resume)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("r\u{00E9}sum\u{00E9}"),
        "Generated Typst should contain NFC-composed résumé: {result}"
    );
}

#[test]
fn test_escape_typst_already_nfc_unchanged() {
    let nfc_text = "Hello 한글 café";
    let result = escape_typst(nfc_text);
    assert_eq!(result, nfc_text, "Already-NFC text should be unchanged");
}

// --- US-103: Multi-column section layout codegen tests ---

#[test]
fn test_generate_flow_page_with_equal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Column text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(2, gutter: 36pt)"),
        "Should contain columns() call. Got: {result}"
    );
    assert!(
        result.contains("Column text"),
        "Should contain the text content. Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_three_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Three col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 3,
            spacing: 18.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(3, gutter: 18pt)"),
        "Should contain columns(3, ...). Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_unequal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Unequal col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: Some(vec![300.0, 150.0]),
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#grid(columns: (300pt, 150pt)"),
        "Unequal columns should use grid(). Got: {result}"
    );
}

#[test]
fn test_generate_column_break() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![
            make_paragraph("Before break"),
            Block::ColumnBreak,
            make_paragraph("After break"),
        ],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#colbreak()"),
        "Should contain colbreak(). Got: {result}"
    );
}

#[test]
fn test_generate_no_columns_no_wrapper() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Normal text")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#columns("),
        "Should not contain columns(). Got: {result}"
    );
    assert!(
        !result.contains("#grid(columns:"),
        "Should not contain grid(columns:). Got: {result}"
    );
}

// ── BiDi / RTL codegen tests ──────────────────────────────────────

#[test]
fn test_generate_rtl_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            direction: Some(TextDirection::Rtl),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "مرحبا بالعالم".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(dir: rtl)"),
        "RTL paragraph should emit #set text(dir: rtl). Got: {result}"
    );
}

#[test]
fn test_generate_ltr_paragraph_no_direction() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("dir: rtl"),
        "LTR paragraph should not emit dir: rtl. Got: {result}"
    );
}

#[test]
fn test_generate_mixed_rtl_ltr_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                direction: Some(TextDirection::Rtl),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "مرحبا 123".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }),
        make_paragraph("English text"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(dir: rtl)"),
        "Should contain RTL direction for Arabic paragraph. Got: {result}"
    );
    assert!(result.contains("مرحبا 123"), "Arabic text should appear");
    assert!(
        result.contains("English text"),
        "English text should appear"
    );
}

// --- US-204: Codegen/render robustness tests ---

#[test]
fn test_codegen_robustness_zero_pages() {
    let doc = make_doc(vec![]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.images.is_empty());
}

#[test]
fn test_codegen_robustness_flow_page_empty_content() {
    let doc = make_doc(vec![make_flow_page(vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_fixed_page_empty_elements() {
    let doc = make_doc(vec![Page::Fixed(FixedPage {
        size: PageSize::default(),
        elements: vec![],
        background_color: None,
        background_gradient: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_table_page_empty_rows() {
    let doc = make_doc(vec![Page::Table(TablePage {
        name: String::new(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![],
            column_widths: vec![],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_paragraph_all_alignment_variants() {
    for alignment in [
        Some(Alignment::Left),
        Some(Alignment::Center),
        Some(Alignment::Right),
        Some(Alignment::Justify),
        None,
    ] {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment,
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: format!("Alignment: {alignment:?}"),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not fail for alignment {alignment:?}"
        );
    }
}

#[test]
fn test_generate_shape_shadow_all_kinds() {
    let shadow = Shadow {
        blur_radius: 4.0,
        color: Color { r: 0, g: 0, b: 0 },
        opacity: 0.5,
        direction: 45.0,
        distance: 3.0,
    };

    let shape_kinds = vec![
        ShapeKind::Rectangle,
        ShapeKind::Ellipse,
        ShapeKind::Line { x2: 100.0, y2: 0.0 },
        ShapeKind::RoundedRectangle {
            radius_fraction: 0.1,
        },
        ShapeKind::Polygon {
            vertices: vec![(0.0, 0.0), (1.0, 0.0), (0.5, 1.0)],
        },
    ];

    for kind in shape_kinds {
        let doc = make_doc(vec![Page::Fixed(FixedPage {
            size: PageSize {
                width: 960.0,
                height: 540.0,
            },
            elements: vec![FixedElement {
                x: 100.0,
                y: 100.0,
                width: 200.0,
                height: 100.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: kind.clone(),
                    fill: Some(Color { r: 255, g: 0, b: 0 }),
                    gradient_fill: None,
                    stroke: None,
                    opacity: None,
                    shadow: Some(shadow.clone()),
                    rotation_deg: None,
                }),
            }],
            background_color: None,
            background_gradient: None,
        })]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not panic for shape kind {kind:?} with shadow"
        );
    }
}

#[test]
fn test_column_break_with_empty_content() {
    let segments = split_at_column_breaks(&[]);
    assert_eq!(segments.len(), 1);
    assert!(segments[0].is_empty());
}

#[test]
fn test_column_break_only_breaks() {
    let blocks = vec![Block::ColumnBreak, Block::ColumnBreak];
    let segments = split_at_column_breaks(&blocks);
    assert_eq!(segments.len(), 3);
    assert!(segments.iter().all(|segment| segment.is_empty()));
}

// --- US-315: text escaping for Typst-significant characters ---

#[test]
fn test_escape_typst_backslash() {
    assert_eq!(escape_typst("path\\to\\file"), "path\\\\to\\\\file");
}

#[test]
fn test_escape_typst_hash() {
    assert_eq!(escape_typst("#hashtag"), "\\#hashtag");
}

#[test]
fn test_escape_typst_dollar() {
    assert_eq!(escape_typst("$100"), "\\$100");
}

#[test]
fn test_escape_typst_brackets() {
    assert_eq!(escape_typst("[content]"), "\\[content\\]");
}

#[test]
fn test_escape_typst_braces() {
    assert_eq!(escape_typst("{code}"), "\\{code\\}");
}

#[test]
fn test_escape_typst_all_special_chars() {
    let input = r"#*_`<>@\~/$[]{}";
    let result = escape_typst(input);
    assert_eq!(result, "\\#\\*\\_\\`\\<\\>\\@\\\\\\~\\/\\$\\[\\]\\{\\}");
}

#[test]
fn test_escape_typst_in_paragraph_output() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 path\\to",
    )])]);
    let output = generate_typst(&doc).unwrap().source;
    assert!(
        output.contains("\\$100"),
        "Dollar sign should be escaped in output: {output}"
    );
    assert!(
        output.contains("path\\\\to"),
        "Backslash should be escaped in output: {output}"
    );
}

// --- US-316: single-stop gradient fallback ---

#[test]
fn test_gradient_single_stop_fallback_to_solid() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![GradientStop {
                position: 0.5,
                color: Color::new(255, 128, 0),
            }],
            angle: 0.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("gradient.linear"),
        "Single-stop gradient should fall back to solid fill: {}",
        output.source,
    );
    assert!(
        output.source.contains("rgb(255, 128, 0)"),
        "Single-stop gradient should use the stop color as solid fill: {}",
        output.source,
    );
}

#[test]
fn test_gradient_two_stops_still_works() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear"),
        "Two-stop gradient should still produce gradient.linear: {}",
        output.source,
    );
}

// --- US-382/383: unstyled run after styled run must not create `](` pattern ---

#[test]
fn test_unstyled_run_with_parens_after_styled_run() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "bold text".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "(parenthetical note)".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("](\\(") || !result.contains("]("),
        "Unstyled text with parens after styled run must be wrapped safely. Got: {result}"
    );
    assert!(
        result.contains("#[") || result.contains("\\("),
        "Unstyled text should be wrapped in #[...] to prevent syntax issues. Got: {result}"
    );
}
