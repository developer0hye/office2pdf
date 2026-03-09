use super::*;

#[test]
fn test_fixed_page_sets_page_size() {
    let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("width: 960pt"),
        "Expected slide width in: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 540pt"),
        "Expected slide height in: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_zero_margins() {
    let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("margin: 0pt"),
        "Expected zero margins for slide in: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_text_box() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_text_box(100.0, 200.0, 300.0, 50.0, "Slide Title")],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Slide Title"));
    assert!(output.source.contains("100pt"));
    assert!(output.source.contains("200pt"));
}

#[test]
fn test_fixed_page_text_box_uses_padding_and_center_vertical_align() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_text_box(
            100.0,
            200.0,
            300.0,
            50.0,
            Insets {
                top: 3.6,
                right: 7.2,
                bottom: 3.6,
                left: 7.2,
            },
            crate::ir::TextBoxVerticalAlign::Center,
            vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Centered".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("inset: (top: 3.6pt, right: 7.2pt, bottom: 3.6pt, left: 7.2pt)")
    );
    assert!(
        output
            .source
            .contains("#let text_box_content_0 = block(width: 285.6pt)[")
    );
    assert!(output.source.contains(
        "#context {\n    let text_box_slack_0 = calc.max(42.8pt - measure(text_box_content_0).height, 0pt)"
    ));
    assert!(output.source.contains("#v(text_box_slack_0 / 2)"));
    assert!(output.source.contains("let text_box_aligned_0 = ["));
}

#[test]
fn test_fixed_page_text_box_multiple_paragraphs_preserve_breaks() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 300.0,
            height: 100.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "First item".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }),
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Second item".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }),
                ],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("First item"));
    assert!(output.source.contains("Second item"));
    assert!(output.source.contains("First item\n\n  Second item"));
}

#[test]
fn test_fixed_page_text_box_ordered_list_preserves_textbox_styling() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 300.0,
            height: 100.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Ordered,
                    items: vec![
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle {
                                    line_spacing: Some(LineSpacing::Proportional(1.5)),
                                    ..ParagraphStyle::default()
                                },
                                runs: vec![Run {
                                    text: " First item".to_string(),
                                    style: TextStyle {
                                        font_size: Some(24.0),
                                        ..TextStyle::default()
                                    },
                                    href: None,
                                    footnote: None,
                                }],
                            }],
                            level: 0,
                            start_at: Some(1),
                        },
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle {
                                    line_spacing: Some(LineSpacing::Proportional(1.5)),
                                    ..ParagraphStyle::default()
                                },
                                runs: vec![Run {
                                    text: " Second item".to_string(),
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
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Ordered,
                            numbering_pattern: Some("1.".to_string()),
                            full_numbering: false,
                            marker_text: None,
                            marker_style: None,
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("#enum("));
    assert!(
        output
            .source
            .contains("#text(size: 24pt)[1.]#text(size: 24pt)[ First item]")
    );
    assert!(
        output
            .source
            .contains("#text(size: 24pt)[2.]#text(size: 24pt)[ Second item]")
    );
    assert!(!output.source.contains("\\\n2. Second item"));
    assert!(output.source.contains("#v(12pt)"));
    assert!(output.source.contains("#set par(leading: 12pt)"));
}

#[test]
fn test_fixed_page_text_box_compact_list_items_use_full_width_blocks() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Ordered,
                    items: vec![
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Long first item that should wrap inside the fixed text box width".to_string(),
                                    style: TextStyle {
                                        font_size: Some(20.0),
                                        ..TextStyle::default()
                                    },
                                    href: None,
                                    footnote: None,
                                }],
                            }],
                            level: 0,
                            start_at: Some(1),
                        },
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Long second item that should also wrap inside the fixed text box width".to_string(),
                                    style: TextStyle {
                                        font_size: Some(20.0),
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
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Ordered,
                            numbering_pattern: Some("1)".to_string()),
                            full_numbering: false,
                            marker_text: None,
                            marker_style: None,
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert_eq!(output.source.matches("#block(width: 320pt)[").count(), 2);
}

#[test]
fn test_fixed_page_text_box_compact_list_preserves_hanging_indent() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Ordered,
                    items: vec![ListItem {
                        content: vec![Paragraph {
                            style: ParagraphStyle {
                                indent_left: Some(36.0),
                                indent_first_line: Some(-36.0),
                                ..ParagraphStyle::default()
                            },
                            runs: vec![Run {
                                text: "Long first item that should wrap under the body text instead of the number".to_string(),
                                style: TextStyle {
                                    font_size: Some(20.0),
                                    ..TextStyle::default()
                                },
                                href: None,
                                footnote: None,
                            }],
                        }],
                        level: 0,
                        start_at: Some(1),
                    }],
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Ordered,
                            numbering_pattern: Some("1)".to_string()),
                            full_numbering: false,
                            marker_text: None,
                            marker_style: None,
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("hanging-indent: 36pt"));
    assert!(
        output
            .source
            .contains("tab_advance_1 = if tab_prefix_width_1 < 36pt")
    );
    assert!(!output.source.contains("first-line-indent"));
}

#[test]
fn test_fixed_page_text_box_compact_list_preserves_marker_origin_offset() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Ordered,
                    items: vec![ListItem {
                        content: vec![Paragraph {
                            style: ParagraphStyle {
                                indent_left: Some(54.0),
                                indent_first_line: Some(-36.0),
                                ..ParagraphStyle::default()
                            },
                            runs: vec![Run {
                                text: "Marker origin should stay inset while wrapped lines align to the text column"
                                    .to_string(),
                                style: TextStyle {
                                    font_size: Some(20.0),
                                    ..TextStyle::default()
                                },
                                href: None,
                                footnote: None,
                            }],
                        }],
                        level: 0,
                        start_at: Some(1),
                    }],
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Ordered,
                            numbering_pattern: Some("1)".to_string()),
                            full_numbering: false,
                            marker_text: None,
                            marker_style: None,
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(
        output
            .source
            .contains("inset: (top: 0pt, right: 0pt, bottom: 0pt, left: 18pt)")
    );
    assert!(output.source.contains("hanging-indent: 36pt"));
}

#[test]
fn test_fixed_page_text_box_compact_bulleted_list_uses_custom_marker_style() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Unordered,
                    items: vec![ListItem {
                        content: vec![Paragraph {
                            style: ParagraphStyle {
                                indent_left: Some(22.5),
                                indent_first_line: Some(-22.5),
                                ..ParagraphStyle::default()
                            },
                            runs: vec![Run {
                                text: "Symbol bullet".to_string(),
                                style: TextStyle {
                                    font_family: Some("Pretendard".to_string()),
                                    font_size: Some(14.0),
                                    ..TextStyle::default()
                                },
                                href: None,
                                footnote: None,
                            }],
                        }],
                        level: 0,
                        start_at: None,
                    }],
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Unordered,
                            numbering_pattern: None,
                            full_numbering: false,
                            marker_text: Some("è".to_string()),
                            marker_style: Some(TextStyle {
                                font_family: Some("Wingdings".to_string()),
                                font_size: Some(14.0),
                                ..TextStyle::default()
                            }),
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(!output.source.contains("Wingdings"));
    assert!(output.source.contains("➔"));
    assert!(output.source.contains("tab_advance_1"));
    assert!(output.source.contains("Symbol bullet"));
}

#[test]
fn test_escape_typst_escapes_leading_dash_list_prefix() {
    assert_eq!(escape_typst("- bullet"), "\\- bullet");
}

#[test]
fn test_fixed_page_text_box_dash_bullets_use_generic_list_path() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Unordered,
                    items: vec![
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle {
                                    indent_left: Some(22.5),
                                    indent_first_line: Some(-22.5),
                                    ..ParagraphStyle::default()
                                },
                                runs: vec![Run {
                                    text: "First dash bullet".to_string(),
                                    style: TextStyle {
                                        font_family: Some("Pretendard".to_string()),
                                        font_size: Some(14.0),
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
                                    indent_left: Some(22.5),
                                    indent_first_line: Some(-22.5),
                                    ..ParagraphStyle::default()
                                },
                                runs: vec![Run {
                                    text: "Second dash bullet".to_string(),
                                    style: TextStyle {
                                        font_family: Some("Pretendard".to_string()),
                                        font_size: Some(14.0),
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
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Unordered,
                            numbering_pattern: None,
                            full_numbering: false,
                            marker_text: Some("-".to_string()),
                            marker_style: Some(TextStyle {
                                font_family: Some("Pretendard".to_string()),
                                font_size: Some(14.0),
                                ..TextStyle::default()
                            }),
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("#list(marker: ["));
    assert!(!output.source.contains("tab_advance_1"));
}

#[test]
fn test_fixed_page_text_box_compact_list_preserves_soft_line_breaks() {
    use crate::ir::List;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 200.0,
            width: 320.0,
            height: 140.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::List(List {
                    kind: ListKind::Ordered,
                    items: vec![ListItem {
                        content: vec![Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Line 1\u{000B}Line 2".to_string(),
                                style: TextStyle {
                                    font_size: Some(20.0),
                                    ..TextStyle::default()
                                },
                                href: None,
                                footnote: None,
                            }],
                        }],
                        level: 0,
                        start_at: Some(1),
                    }],
                    level_styles: BTreeMap::from([(
                        0,
                        ListLevelStyle {
                            kind: ListKind::Ordered,
                            numbering_pattern: Some("1)".to_string()),
                            full_numbering: false,
                            marker_text: None,
                            marker_style: None,
                        },
                    )]),
                })],
                padding: Insets::default(),
                vertical_align: crate::ir::TextBoxVerticalAlign::Top,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("#linebreak()"));
    assert!(output.source.contains("#set text(size: 20pt"));
    assert!(output.source.contains("leading: 13pt"));
}

#[test]
fn test_fixed_page_text_box_with_width_height() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_text_box(50.0, 60.0, 400.0, 100.0, "Sized box")],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("400pt"));
    assert!(output.source.contains("100pt"));
}

#[test]
fn test_fixed_page_rectangle_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            150.0,
            ShapeKind::Rectangle,
            Some(Color::new(255, 0, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rect"));
    assert!(output.source.contains("200pt"));
    assert!(output.source.contains("rgb(255, 0, 0)"));
}

#[test]
fn test_fixed_page_ellipse_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            50.0,
            50.0,
            120.0,
            80.0,
            ShapeKind::Ellipse,
            Some(Color::new(0, 128, 255)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("ellipse"));
}

#[test]
fn test_fixed_page_line_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            300.0,
            0.0,
            ShapeKind::Line { x2: 300.0, y2: 0.0 },
            None,
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("line"));
}

#[test]
fn test_fixed_page_shape_with_stroke() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Rectangle,
            None,
            Some(BorderSide {
                width: 1.5,
                color: Color::new(0, 0, 255),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("stroke"));
    assert!(output.source.contains("1.5pt"));
}

#[test]
fn test_shape_rotation_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(255, 0, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: Some(90.0),
                opacity: None,
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rotate"));
    assert!(output.source.contains("90deg"));
}

#[test]
fn test_shape_opacity_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(0, 255, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: Some(0.5),
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rgb(0, 255, 0, 128)"));
}

#[test]
fn test_shape_rotation_and_opacity_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 50.0,
            y: 50.0,
            width: 100.0,
            height: 100.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Ellipse,
                fill: Some(Color::new(0, 0, 255)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: Some(45.0),
                opacity: Some(0.75),
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rotate"));
    assert!(output.source.contains("45deg"));
    assert!(output.source.contains("rgb(0, 0, 255, 191)"));
}

#[test]
fn test_fixed_page_image_element() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_image(
            100.0,
            150.0,
            400.0,
            300.0,
            ImageFormat::Png,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#image("));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_fixed_page_mixed_elements() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(50.0, 30.0, 800.0, 60.0, "Title"),
            make_shape_element(
                50.0,
                100.0,
                400.0,
                300.0,
                ShapeKind::Rectangle,
                Some(Color::new(200, 200, 200)),
                None,
            ),
            make_fixed_image(500.0, 100.0, 350.0, 300.0, ImageFormat::Jpeg),
            make_text_box(50.0, 420.0, 800.0, 40.0, "Footer text"),
        ],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Title"));
    assert!(output.source.contains("rect"));
    assert!(output.source.contains("#image("));
    assert!(output.source.contains("Footer text"));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_fixed_page_multiple_text_boxes() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(100.0, 50.0, 300.0, 40.0, "First"),
            make_text_box(100.0, 120.0, 300.0, 40.0, "Second"),
            make_text_box(100.0, 190.0, 300.0, 40.0, "Third"),
        ],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("First"));
    assert!(output.source.contains("Second"));
    assert!(output.source.contains("Third"));
}

#[test]
fn test_fixed_page_uses_place_for_positioning() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_text_box(100.0, 200.0, 300.0, 50.0, "Positioned")],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("place("));
}
