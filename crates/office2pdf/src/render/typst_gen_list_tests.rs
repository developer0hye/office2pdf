use super::*;
use std::collections::BTreeMap;

#[test]
fn test_generate_bulleted_list() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Apple".to_string(),
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
                        text: "Banana".to_string(),
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
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#list("));
    assert!(output.source.contains("Apple"));
    assert!(output.source.contains("Banana"));
}

#[test]
fn test_generate_numbered_list() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Step 1".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: Some(3),
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Step 2".to_string(),
                        style: TextStyle::default(),
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
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#enum("));
    assert!(output.source.contains("start: 3"));
    assert!(output.source.contains("numbering: \"1.\""));
    assert!(output.source.contains("Step 1"));
    assert!(output.source.contains("Step 2"));
}

#[test]
fn test_generate_numbered_list_preserves_hanging_indent_columns() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    indent_left: Some(36.0),
                    indent_first_line: Some(-18.0),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "Indented item".to_string(),
                    style: TextStyle::default(),
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
                numbering_pattern: Some("1.".to_string()),
                full_numbering: false,
                marker_text: None,
                marker_style: None,
            },
        )]),
    };

    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(output.source.contains("indent: 18pt"));
    assert!(output.source.contains("body-indent: 0pt"));
    assert!(output.source.contains("#box(width: 18pt"));
}

#[test]
fn test_generate_bulleted_list_preserves_nonstandard_hanging_indent_columns() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    indent_left: Some(45.0),
                    indent_first_line: Some(-15.0),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "Indented item".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }],
            level: 0,
            start_at: None,
        }],
        level_styles: BTreeMap::new(),
    };

    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(output.source.contains("indent: 30pt"));
    assert!(output.source.contains("body-indent: 0pt"));
    assert!(output.source.contains("marker: [#box(width: 15pt"));
}

#[test]
fn test_generate_list_preserves_paragraph_spacing_between_items() {
    use crate::ir::List;

    let make_item = |text: &str| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle {
                space_after: Some(8.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    let list = List {
        kind: ListKind::Unordered,
        items: vec![make_item("First"), make_item("Second")],
        level_styles: BTreeMap::new(),
    };

    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(output.source.contains("spacing: 19pt"), "{}", output.source);
    assert!(
        output.source.contains("#block(below: 19pt)"),
        "{}",
        output.source
    );
}

#[test]
fn test_generate_list_uses_word_line_box_and_boundary_spacing() {
    use crate::ir::List;

    let make_item = |text: &str| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle {
                line_box: Some(LineBox {
                    ascent_em: 1.3125,
                    descent_em: 0.4375,
                }),
                space_after: Some(8.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_size: Some(11.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    let list = List {
        kind: ListKind::Unordered,
        items: vec![make_item("First"), make_item("Second")],
        level_styles: BTreeMap::new(),
    };

    let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])]))
        .unwrap()
        .source;

    assert!(
        source.contains("#set text(top-edge: 1.3125em, bottom-edge: -0.4375em)"),
        "{source}"
    );
    assert!(source.contains("#set par(leading: 0pt)"), "{source}");
    assert!(source.contains("spacing: 8pt"), "{source}");
    assert!(
        source.contains("#block(above: 0pt, below: 8pt)"),
        "{source}"
    );
}

#[test]
fn test_generate_list_combines_exact_line_height_with_paragraph_spacing() {
    use crate::ir::List;

    let make_item = |text: &str| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle {
                line_spacing: Some(LineSpacing::Exact(18.0)),
                space_after: Some(6.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_size: Some(13.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    let list = List {
        kind: ListKind::Ordered,
        items: vec![make_item("First"), make_item("Second")],
        level_styles: BTreeMap::new(),
    };

    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(output.source.contains("spacing: 24pt"), "{}", output.source);
    assert!(
        output.source.contains("#block(below: 24pt)"),
        "{}",
        output.source
    );
}

#[test]
fn test_generate_numbered_list_marker_inherits_common_text_font() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Arial item".to_string(),
                    style: TextStyle {
                        font_family: Some("Arial".to_string()),
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
                numbering_pattern: Some("1.".to_string()),
                full_numbering: false,
                marker_text: None,
                marker_style: None,
            },
        )]),
    };
    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(
        output
            .source
            .contains("numbering: (..nums) => [#text(font: (\"Arial\"")
    );
}

#[test]
fn test_generate_symbol_bullet_uses_unicode_and_inherits_common_text_font() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Arial item".to_string(),
                    style: TextStyle {
                        font_family: Some("Arial".to_string()),
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
                marker_text: Some("\u{F0B7}".to_string()),
                marker_style: Some(TextStyle {
                    font_family: Some("Symbol".to_string()),
                    ..TextStyle::default()
                }),
            },
        )]),
    };
    let output = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])])).unwrap();

    assert!(output.source.contains("marker: [#text(font: (\"Arial\""));
    assert!(output.source.contains("[•]"));
    assert!(!output.source.contains("Symbol"));
    assert!(!output.source.contains('\u{F0B7}'));
}

#[test]
fn test_generate_numbered_list_emits_mid_list_restart() {
    use crate::ir::List;

    let make_item = |text: &str, start_at: Option<u32>| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at,
    };
    let list = List {
        kind: ListKind::Ordered,
        items: vec![
            make_item("First", Some(1)),
            make_item("Second", None),
            make_item("Restarted", Some(10)),
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
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);

    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("start: 1"));
    assert!(output.source.contains("enum.item(10)[Restarted]"));
}

#[test]
fn test_generate_nested_list() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Parent".to_string(),
                        style: TextStyle::default(),
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
                        text: "Child".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 1,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Sibling".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::from([
            (
                0,
                ListLevelStyle {
                    kind: ListKind::Ordered,
                    numbering_pattern: Some("1.".to_string()),
                    full_numbering: false,
                    marker_text: None,
                    marker_style: None,
                },
            ),
            (
                1,
                ListLevelStyle {
                    kind: ListKind::Unordered,
                    numbering_pattern: None,
                    full_numbering: false,
                    marker_text: None,
                    marker_style: None,
                },
            ),
        ]),
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Parent"));
    assert!(output.source.contains("Child"));
    assert!(output.source.contains("Sibling"));
    assert!(output.source.contains("#enum("));
    assert!(output.source.contains("#list("));
}

#[test]
fn test_nested_list_single_content_block() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Parent".to_string(),
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
                        text: "Child".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 1,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("][#list"));
    assert!(output.source.contains(" #list("));
}

#[test]
fn test_generate_nested_ordered_list_uses_full_numbering() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Parent".to_string(),
                        style: TextStyle::default(),
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
                        text: "Child".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 1,
                start_at: Some(1),
            },
        ],
        level_styles: BTreeMap::from([
            (
                0,
                ListLevelStyle {
                    kind: ListKind::Ordered,
                    numbering_pattern: Some("1.".to_string()),
                    full_numbering: false,
                    marker_text: None,
                    marker_style: None,
                },
            ),
            (
                1,
                ListLevelStyle {
                    kind: ListKind::Ordered,
                    numbering_pattern: Some("1.a.".to_string()),
                    full_numbering: true,
                    marker_text: None,
                    marker_style: None,
                },
            ),
        ]),
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("full: true"));
    assert!(output.source.contains("numbering: \"1.a.\""));
}

#[test]
fn test_generate_bulleted_list_with_custom_marker_text_and_style() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Dash marker".to_string(),
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
                marker_text: Some("-".to_string()),
                marker_style: Some(TextStyle {
                    font_family: Some("Pretendard".to_string()),
                    font_size: Some(14.0),
                    color: Some(Color::new(0x11, 0x22, 0x33)),
                    ..TextStyle::default()
                }),
            },
        )]),
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("#list("));
    assert!(output.source.contains("marker: [#text("));
    assert!(output.source.contains("Pretendard"));
    assert!(output.source.contains("fill: rgb(17, 34, 51)"));
    assert!(output.source.contains("Dash marker"));
}

#[test]
fn test_generate_ordered_list_with_custom_marker_style_uses_numbering_function() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Ordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Ordered marker".to_string(),
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
                marker_style: Some(TextStyle {
                    font_family: Some("Pretendard Medium".to_string()),
                    font_size: Some(20.0),
                    color: Some(Color::black()),
                    ..TextStyle::default()
                }),
            },
        )]),
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("#enum("));
    assert!(output.source.contains("numbering: (..nums) => ["));
    assert!(output.source.contains("#numbering(\"1)\", ..nums)"));
    assert!(output.source.contains("Pretendard Medium"));
}

#[test]
fn test_generate_bulleted_list_with_symbol_font_marker_uses_unicode_fallback() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Arrow bullet".to_string(),
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
                    color: Some(Color::black()),
                    ..TextStyle::default()
                }),
            },
        )]),
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("➔"));
    assert!(!output.source.contains("Wingdings"));
    assert!(output.source.contains("fill: rgb(0, 0, 0)"));
}

#[test]
fn test_generate_list_uses_first_item_level_marker_when_list_starts_nested() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Nested arrow bullet".to_string(),
                    style: TextStyle {
                        font_family: Some("Pretendard".to_string()),
                        font_size: Some(14.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                }],
            }],
            level: 1,
            start_at: None,
        }],
        level_styles: BTreeMap::from([(
            1,
            ListLevelStyle {
                kind: ListKind::Unordered,
                numbering_pattern: None,
                full_numbering: false,
                marker_text: Some("è".to_string()),
                marker_style: Some(TextStyle {
                    font_family: Some("Wingdings".to_string()),
                    font_size: Some(14.0),
                    color: Some(Color::black()),
                    ..TextStyle::default()
                }),
            },
        )]),
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains("➔"));
    assert!(!output.source.contains("marker: [#text(font: \"Wingdings\""));
    assert!(
        !output
            .source
            .contains("marker: [#text(font: \"맑은 고딕\", size: 14pt, fill: rgb(0, 0, 0))[-]]")
    );
}

#[test]
fn test_generate_list_metric_spacing_adds_gap_to_single_space_leading() {
    // Word adds `w:spacing w:after` directly to the single-space line
    // advance between list items: next item top = previous line advance +
    // after. Under metric text edges that whitespace is the single-space
    // leading plus the paragraph gap; adding a whole line height instead
    // stretched every list block by ~8pt per item (issue #384).
    use crate::ir::List;

    let Some((ascender, descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let font_size: f64 = 10.0;
    let single_space_leading_pt: f64 =
        ((word_pitch_em - (ascender + descender)) * font_size).max(0.0);

    let make_item = |text: &str| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle {
                space_after: Some(4.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some("Libertinus Serif".to_string()),
                    font_size: Some(font_size),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    let list = List {
        kind: ListKind::Unordered,
        items: vec![make_item("First"), make_item("Second")],
        level_styles: BTreeMap::new(),
    };

    let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])]))
        .unwrap()
        .source;

    assert!(
        source.contains("top-edge: \"ascender\""),
        "metric edges expected in: {source}"
    );
    let expected = format_f64(single_space_leading_pt + 4.0);
    assert!(
        source.contains(&format!("spacing: {expected}pt")),
        "expected inter-item spacing {expected}pt (leading + gap) in: {source}"
    );
    assert!(
        source.contains(&format!("below: {expected}pt")),
        "expected list below spacing {expected}pt (leading + gap) in: {source}"
    );
}
