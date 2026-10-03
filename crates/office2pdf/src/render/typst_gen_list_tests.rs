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
                        inline_box: None,
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
                        inline_box: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#list("));
    assert!(output.source.contains("Apple"));
    assert!(output.source.contains("Banana"));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_list_marker_sits_on_the_baseline_of_a_taller_first_line() {
    // Word seats a bullet on its item's first baseline even when a larger run
    // makes that line taller than the marker. typst 0.14 top-aligned native
    // list markers, leaving an 11pt bullet 15.94pt above a 28pt first line in
    // the native-Word probe of issue #1686; typst 0.15 keeps it on the line.
    let run = |text: &str, size: f64| Run {
        text: text.to_string(),
        style: TextStyle {
            font_size: Some(size),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
        inline_box: None,
    };
    let item = |runs: Vec<Run>| ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle::default(),
            runs,
        }],
        level: 0,
        start_at: None,
    };
    let list = crate::ir::List {
        kind: ListKind::Unordered,
        items: vec![
            item(vec![run("Quarterly", 28.0), run(" targets met", 11.0)]),
            item(vec![run("Second item", 11.0)]),
        ],
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let body = |needle: &str| {
        runs.iter()
            .find(|run| run.text.contains(needle))
            .unwrap_or_else(|| panic!("no {needle} run in {runs:?}"))
    };
    let (first_line, second_line) = (body("Quarterly"), body("Second"));
    // A marker sits in the hanging indent, left of the item's text.
    let mut markers: Vec<f64> = runs
        .iter()
        .filter(|run| !run.text.trim().is_empty() && run.left_pt < first_line.left_pt - 1.0)
        .map(|run| run.baseline_pt)
        .collect();
    markers.sort_by(f64::total_cmp);
    assert_eq!(markers.len(), 2, "one marker per item: {runs:?}");
    assert!(
        (markers[0] - first_line.baseline_pt).abs() < 0.01,
        "the bullet must sit on the taller first line's baseline: marker {} vs line {}",
        markers[0],
        first_line.baseline_pt
    );
    assert!(
        (markers[1] - second_line.baseline_pt).abs() < 0.01,
        "the control item's bullet stays on its own baseline: marker {} vs line {}",
        markers[1],
        second_line.baseline_pt
    );
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
                        inline_box: None,
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
                        inline_box: None,
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
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
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
                    inline_box: None,
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
                    inline_box: None,
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
                inline_box: None,
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
        output.source.contains("#block(width: 100%, below: 19pt)"),
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
                inline_box: None,
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
        source.contains("#block(width: 100%, above: 0pt, below: 8pt)"),
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
                inline_box: None,
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
        output.source.contains("#block(width: 100%, below: 24pt)"),
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
                    inline_box: None,
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
                    inline_box: None,
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
                inline_box: None,
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
                        inline_box: None,
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
                        inline_box: None,
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
                        inline_box: None,
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
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
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
                        inline_box: None,
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
                        inline_box: None,
                    }],
                }],
                level: 1,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::List(list)],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
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
                        inline_box: None,
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
                        inline_box: None,
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
                    inline_box: None,
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
                    inline_box: None,
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
                    inline_box: None,
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
                    inline_box: None,
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
fn test_generate_list_metric_spacing_is_the_raw_paragraph_gap() {
    // Word adds `w:spacing w:after` directly to the single-space line
    // advance between list items: next item top = previous line advance +
    // after. The wrapper's line box already spans that advance, so the
    // Typst list `spacing` that replaces the automatic leading is the raw
    // paragraph gap — adding a whole line height instead stretched every
    // list block by ~8pt per item (issues #384, #452).
    use crate::ir::List;

    let Some((ascender, descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let font_size: f64 = 10.0;
    let metric_em: f64 = ascender + descender;

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
                inline_box: None,
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

    let advance_pt: f64 = (word_pitch_em * font_size).max(metric_em * font_size);
    assert_line_advance(&source, "Libertinus Serif", font_size, advance_pt, 0.0);
    assert!(
        source.contains("spacing: 4pt"),
        "expected inter-item spacing to be the raw 4pt gap in: {source}"
    );
    assert!(
        source.contains("below: 4pt"),
        "expected list below spacing to be the raw 4pt gap in: {source}"
    );
}

#[test]
fn test_list_spacing_stays_the_paragraph_gap_when_items_state_proportional_line_spacing() {
    // Word's default template writes `w:spacing w:line="259" w:lineRule="auto"`
    // into `w:pPrDefault`, so nearly every authored list states a line
    // multiple. That multiple is folded into the wrapper's line box, exactly
    // as a bare single-spaced list's advance is, so the gap between two items
    // is still the raw `w:after` — office2pdf added a whole scaled line on top
    // of it, opening every item gap and the gap below the list by roughly one
    // line per item (issue #1685).
    //
    // Two multiples pin the rule: a gap that were hardcoded to one of them
    // would have to grow with the other.
    use crate::ir::List;

    let Some((ascender, descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let font_size: f64 = 10.0;
    let gap_pt: f64 = 8.0;
    let single_advance_pt: f64 =
        (word_pitch_em * font_size).max((ascender + descender) * font_size);

    for factor in [259.0 / 240.0, 1.5] {
        let make_item = |text: &str| ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    space_after: Some(gap_pt),
                    line_spacing: Some(LineSpacing::Proportional(factor)),
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
                    inline_box: None,
                }],
            }],
            level: 0,
            start_at: None,
        };
        let list = List {
            kind: ListKind::Unordered,
            items: vec![make_item("First"), make_item("Second"), make_item("Third")],
            level_styles: BTreeMap::new(),
        };

        let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])]))
            .unwrap()
            .source;

        // The multiple belongs to the line box, not to the item gap.
        assert_line_advance(
            &source,
            "Libertinus Serif",
            font_size,
            single_advance_pt * factor,
            0.0,
        );
        assert!(
            source.contains(&format!("spacing: {}pt", format_f64(gap_pt))),
            "item gap at line multiple {factor} should stay the raw {gap_pt}pt in: {source}"
        );
        assert!(
            source.contains(&format!("below: {}pt", format_f64(gap_pt))),
            "gap below the list at line multiple {factor} should stay the raw {gap_pt}pt in: {source}"
        );
    }
}

#[test]
fn test_list_wrapper_block_carries_the_edge_spacing() {
    // The line-height wrapper used to leave `above`/`below` unset, so
    // Typst's own 1.2em default block spacing governed the gap between a
    // list and its neighbours while the computed gap sat on an inner block
    // where it could not reach the boundary. A numbered item followed by a
    // shaded code paragraph opened ~10pt too far (issue #463).
    use crate::ir::List;

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
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
                inline_box: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    let list = List {
        kind: ListKind::Ordered,
        items: vec![
            make_item("Convert a single file."),
            make_item("Batch convert."),
        ],
        level_styles: BTreeMap::new(),
    };
    let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::List(list)])]))
        .unwrap()
        .source;

    assert!(
        !source.contains("#block(width: 100%)["),
        "the list wrapper must pin its own spacing, not inherit Typst's default: {source}"
    );
    let wrapper_start = source
        .find("#block(width: 100%")
        .expect("the list is wrapped in a full-width block");
    let wrapper_params =
        &source[wrapper_start..source[wrapper_start..].find(")[").unwrap() + wrapper_start];
    assert!(
        wrapper_params.contains("below: 4pt"),
        "the outermost list block carries the item gap: {wrapper_params}"
    );
}

// ----- PowerPoint's per-item list spacing (issue #524) -----

/// A slide text box holding one bullet list whose items declare the given
/// `a:spcAft` gaps, in points.
fn slide_bullet_list_source(gaps: &[f64]) -> String {
    let items: Vec<ListItem> = gaps
        .iter()
        .enumerate()
        .map(|(index, gap)| ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    space_after: Some(*gap),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: format!("Bullet {index}"),
                    style: TextStyle {
                        font_family: Some("Libertinus Serif".to_string()),
                        font_size: Some(17.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            }],
            level: 0,
            start_at: None,
        })
        .collect();
    let list = List {
        kind: ListKind::Unordered,
        items,
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_text_box(
            50.0,
            50.0,
            600.0,
            400.0,
            Insets::default(),
            crate::ir::TextBoxVerticalAlign::Top,
            vec![Block::List(list)],
        )],
    )]);
    generate_typst(&doc).unwrap().source
}

/// A slide outline with the same root/child shape and alternating paragraph
/// spacing as `04_training_deck_ko` page 2 (issue #659).
#[cfg(not(target_arch = "wasm32"))]
fn nested_slide_bullet_list(gaps: &[f64], levels: &[u32]) -> List {
    assert_eq!(gaps.len(), levels.len());
    let items = gaps
        .iter()
        .zip(levels)
        .enumerate()
        .map(|(index, (gap, level))| ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    space_after: Some(*gap),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: format!("OutlineItem{index}"),
                    style: TextStyle {
                        font_family: Some("Liberation Sans".to_string()),
                        font_size: Some(17.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            }],
            level: *level,
            start_at: None,
        })
        .collect();
    List {
        kind: ListKind::Unordered,
        items,
        level_styles: BTreeMap::new(),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn nested_slide_item_pitches(gaps: &[f64], levels: &[u32]) -> Vec<f64> {
    let list = nested_slide_bullet_list(gaps, levels);
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_text_box(
            50.0,
            50.0,
            600.0,
            400.0,
            Insets::default(),
            crate::ir::TextBoxVerticalAlign::Top,
            vec![Block::List(list)],
        )],
    )]);
    let source = generate_typst(&doc).unwrap().source;
    let baselines: Vec<f64> = crate::render::pdf::compiled_text_runs(&source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{source}"))
        .into_iter()
        .filter(|run| run.text.contains("OutlineItem"))
        .map(|run| run.baseline_pt)
        .collect();
    assert_eq!(
        baselines.len(),
        gaps.len(),
        "all outline items compile: {source}"
    );

    baselines.windows(2).map(|pair| pair[1] - pair[0]).collect()
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn nested_slide_items_keep_their_paragraph_gaps_across_level_changes() {
    let pitches = nested_slide_item_pitches(
        &[6.0, 10.0, 6.0, 6.0, 10.0, 6.0, 0.0],
        &[0, 1, 0, 1, 1, 0, 1],
    );

    let expected = [26.4, 30.4, 26.4, 26.4, 30.4, 26.4];
    for (index, (actual, expected)) in pitches.iter().zip(expected).enumerate() {
        assert!(
            (actual - expected).abs() < 0.05,
            "boundary {index} must be one 20.4pt PowerPoint line plus the previous paragraph's gap: expected {expected}pt, got {actual}pt; all pitches {pitches:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_nested_slide_never_hoists_spacing_from_only_its_adjacent_root_items() {
    // `list(spacing:)` applies after every root item, including one whose last
    // paragraph is a nested child. Seeing the first root-to-root 6pt boundary
    // must not hoist 6pt across the later child-to-root 10pt boundary.
    let pitches = nested_slide_item_pitches(
        &[6.0, 6.0, 10.0, 0.0, 0.0, 0.0, 0.0],
        &[0, 0, 1, 0, 1, 0, 1],
    );
    let expected = [26.4, 26.4, 30.4, 20.4, 20.4, 20.4];
    for (index, (actual, expected)) in pitches.iter().zip(expected).enumerate() {
        assert!(
            (actual - expected).abs() < 0.05,
            "boundary {index} must use the preceding document paragraph's gap: expected {expected}pt, got {actual}pt; all pitches {pitches:?}"
        );
    }
}

#[test]
fn slide_list_items_keep_their_own_gaps_when_they_differ() {
    // `list(spacing:)` carries one value for the whole level, so a list whose
    // items declare different `a:spcAft` could not use it and got *no* spacing
    // at all. 04_training_deck_ko's outline alternates 6pt and 10pt and lost
    // every one, drifting up to 18.8pt by the last bullet (issue #524).
    let source = slide_bullet_list_source(&[6.0, 10.0, 6.0, 0.0]);

    assert!(
        source.contains("#v(6pt)") && source.contains("#v(10pt)"),
        "each item keeps the gap it declares: {source}"
    );
}

#[test]
fn a_slide_list_with_one_shared_gap_separates_its_items_by_it() {
    // Triangulation for the per-item fallback: a list whose items CAN share
    // one value takes the shared path, and must still be separated by the gap
    // there. This test used to accept the wrapper's `below: 8pt` as proof the
    // gap was "carried once" — but that block encloses the whole list, so it
    // is the list's outer bottom edge and puts nothing between two bullets
    // (issue #928).
    let source = slide_bullet_list_source(&[8.0, 8.0, 8.0]);

    assert!(
        source.contains("below: 8pt"),
        "the list's outer bottom edge stays the declared gap: {source}"
    );
    assert_eq!(
        source.matches("#v(8pt)").count(),
        2,
        "and each of the two boundaries between three items gets it: {source}"
    );
}

#[test]
fn an_item_declaring_no_gap_emits_none() {
    // Zero is zero, not a house value: a trailing `#v(0pt)` would be noise.
    let source = slide_bullet_list_source(&[6.0, 0.0, 6.0]);

    assert!(
        !source.contains("#v(0pt)"),
        "an item with no gap emits nothing: {source}"
    );
}

/// A hanging-indent bullet separates its glyph from the body with the tab
/// alone (issue #685).
///
/// `fixed_text_list_marker` appends a space after the glyph, and the
/// hanging-indent branch then appends the tab that carries the gap to the
/// indent. Both together put the body one space past the indent — 101.60pt
/// against a reference's 99.01pt on the audited deck, enough to move a wrap
/// point.
#[test]
fn a_hanging_indent_bullet_separates_with_the_tab_alone() {
    use crate::ir::List;

    let list = List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    indent_left: Some(27.0),
                    indent_first_line: Some(-27.0),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "Bulleted".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            }],
            level: 0,
            start_at: None,
        }],
        level_styles: BTreeMap::new(),
    };
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_text_box(
            50.0,
            50.0,
            600.0,
            400.0,
            Insets::default(),
            crate::ir::TextBoxVerticalAlign::Top,
            vec![Block::List(list)],
        )],
    )]);

    let source = generate_typst(&doc).unwrap().source;

    // The tab compiles into segments, so the marker is the first of them.
    assert!(
        source.contains("let tab_segment_0 = [•]"),
        "the glyph is the whole first tab segment: {source}"
    );
    assert!(
        !source.contains("let tab_segment_0 = [• ]"),
        "no space may trail the glyph before the tab: {source}"
    );
}

/// PowerPoint separates two bullets by the line advance PLUS `a:spcAft` of the
/// first and `a:spcBef` of the second; the two are added, not collapsed to the
/// larger (issue #928).
///
/// The `AGENDA` slide of `002.CONTOSO.pptx` (attached to #841) declares
/// `spcBef` 4pt and `spcAft` 6pt on every bullet. The reference puts 8.7pt more
/// between two bullets than between the wrapped lines of one; we put 2.4pt
/// *less*, because both gaps landed on the wrapper around the whole list and
/// none between its items.
///
/// Measured as a difference against the same list with no spacing, so the
/// assertion holds whatever line box the resolved face gives.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn slide_bullets_add_their_paragraph_spacing_between_items() {
    use crate::ir::{Insets, List};

    let inter_item_advance = |space_before: Option<f64>, space_after: Option<f64>| -> Option<f64> {
        let make_item = |text: &str| ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle {
                    space_before,
                    space_after,
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle {
                        font_family: Some("Libertinus Serif".to_string()),
                        font_size: Some(24.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            }],
            level: 0,
            start_at: None,
        };
        let list = List {
            kind: ListKind::Unordered,
            items: vec![
                make_item("Introduksjoner"),
                make_item("Viktige oppdateringer"),
                make_item("Avslutning"),
            ],
            level_styles: BTreeMap::new(),
        };
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_fixed_text_box(
                43.2,
                196.6,
                301.7,
                293.0,
                Insets::default(),
                crate::ir::TextBoxVerticalAlign::Top,
                vec![Block::List(list)],
            )],
        )]);
        let source = generate_typst(&doc).unwrap().source;
        let baselines: Vec<f64> = crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap_or_else(|error| panic!("compile failed: {error}\n{source}"))
            .into_iter()
            .filter(|run| run.text.contains("Viktige") || run.text.contains("Avslutning"))
            .map(|run| run.baseline_pt)
            .collect();
        (baselines.len() == 2).then(|| baselines[1] - baselines[0])
    };

    let Some(unspaced) = inter_item_advance(None, None) else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let spaced = inter_item_advance(Some(4.0), Some(6.0)).expect("the same list, spaced");

    assert!(
        (spaced - unspaced - 10.0).abs() < 0.05,
        "4pt before + 6pt after must add 10pt to the item advance, \
         got {spaced}pt against the unspaced {unspaced}pt"
    );
}

/// Two items of a bulleted slide text box are one line advance apart, the same
/// quantity that separates two wrapped lines of one item. Each item was its own
/// `#block`, so the distance between two of them was Typst's default block
/// spacing — 1.2em of the *ambient* size, which has nothing to do with the
/// list's own leading — while two wrapped lines took `par(leading:)`. On the
/// `AGENDA` list of #841's deck (24pt) that is 13.199pt against 15.600pt, and
/// it accumulates at 2.40pt per item (issue #934).
#[test]
fn two_slide_list_items_sit_one_line_advance_apart() {
    let items: Vec<ListItem> = ["20XX høyde- og lavpunkter", "Neste steg"]
        .iter()
        .map(|text| ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: (*text).to_string(),
                    style: TextStyle {
                        font_family: Some("Liberation Sans".to_string()),
                        font_size: Some(24.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            }],
            level: 0,
            start_at: None,
        })
        .collect();
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_text_box(
            50.0,
            50.0,
            600.0,
            400.0,
            Insets::default(),
            crate::ir::TextBoxVerticalAlign::Top,
            vec![Block::List(List {
                kind: ListKind::Unordered,
                items,
                level_styles: BTreeMap::new(),
            })],
        )],
    )]);
    let source = generate_typst(&doc).unwrap().source;

    // PowerPoint gives a 24pt line 1.2x the size. The item's block stands
    // exactly that tall and contributes no spacing of its own, so two items
    // that declare no `a:spcBef`/`a:spcAft` are one line apart with nothing
    // emitted between them.
    assert!(
        source.contains("above: 0pt, below: 0pt"),
        "the item blocks contribute no spacing of their own: {source}"
    );
    assert!(
        source.contains("#set par(leading: 0pt)"),
        "the fixed line box carries the advance: {source}"
    );
    assert!(
        vertical_skips_pt(&source)
            .iter()
            .all(|gap| *gap == 0.0 || *gap > 100.0),
        "no gap is emitted between the items themselves: {source}"
    );
}

/// An explicit 100% DrawingML line spacing is still one PowerPoint line per
/// list item. Keeping zero paragraph gaps explicit must not add Typst's
/// ambient block spacing on top of that line box (issue #1913).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn slide_list_items_with_explicit_single_spacing_sit_one_line_apart() {
    use crate::internal::{Parser, PptxParser};
    use crate::ir::{Insets, LineSpacing, List};

    let fixture: std::path::PathBuf = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/poi/with_japanese.pptx");
    let data: Vec<u8> = std::fs::read(fixture).expect("with_japanese fixture");
    let (document, _warnings) = PptxParser
        .parse(&data, &crate::config::ConvertOptions::default())
        .expect("with_japanese must parse");
    let Page::Fixed(page) = &document.pages[0] else {
        panic!("slide 1 must be fixed")
    };
    for item_text in ["Bullet 1", "Number bullet 1"] {
        let mut list: List = page
            .elements
            .iter()
            .filter_map(|element| match &element.kind {
                FixedElementKind::TextBox(text_box) => Some(&text_box.content),
                _ => None,
            })
            .flatten()
            .find_map(|block| match block {
                Block::List(list)
                    if list.items.iter().any(|item| {
                        item.content.iter().any(|paragraph| {
                            paragraph
                                .runs
                                .iter()
                                .any(|run| run.text.contains(item_text))
                        })
                    }) =>
                {
                    Some(list.clone())
                }
                _ => None,
            })
            .unwrap_or_else(|| panic!("the fixture's first slide has {item_text}"));
        assert_eq!(list.items.len(), 3, "the target list has three items");
        for paragraph in list
            .items
            .iter_mut()
            .flat_map(|item| item.content.iter_mut())
        {
            assert!(
                matches!(paragraph.style.line_spacing, Some(LineSpacing::Proportional(factor)) if (factor - 1.0).abs() < f64::EPSILON),
                "the fixture declares 100% DrawingML line spacing"
            );
            assert_eq!(paragraph.style.space_before, Some(0.0));
            assert_eq!(paragraph.style.space_after, Some(0.0));
            paragraph.style.paragraph_mark_font_family = Some("Libertinus Serif".into());
            for run in &mut paragraph.runs {
                run.style.font_family = Some("Libertinus Serif".into());
            }
        }
        let font_size_pt: f64 = list.items[0].content[0].runs[0]
            .style
            .font_size
            .expect("the fixture declares its list font size");
        let expected_line_advance_pt: f64 =
            crate::render::pdf::POWERPOINT_LINE_HEIGHT_FACTOR * font_size_pt;

        // The fixture's text box has unrelated malformed geometry. Isolate the
        // parsed list in a visible slide box while keeping its original paragraphs.
        let isolated_document = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_fixed_text_box(
                72.0,
                72.0,
                540.0,
                400.0,
                Insets::default(),
                crate::ir::TextBoxVerticalAlign::Top,
                vec![Block::List(list)],
            )],
        )]);
        let source: String = generate_typst(&isolated_document).unwrap().source;
        let compiled_runs = crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap_or_else(|error| panic!("compile failed: {error}\n{source}"));
        let mut item_baselines: Vec<f64> = compiled_runs
            .iter()
            .filter(|run| run.text.to_lowercase().contains("bullet"))
            .map(|run| run.baseline_pt)
            .collect();
        item_baselines.sort_by(f64::total_cmp);

        assert_eq!(
            item_baselines.len(),
            3,
            "each fixture item must paint once, got {compiled_runs:?}; source:\n{source}"
        );
        for gap in item_baselines.windows(2).map(|pair| pair[1] - pair[0]) {
            assert!(
                (gap - expected_line_advance_pt).abs() < 1.0,
                "100% spacing advances one PowerPoint line ({expected_line_advance_pt}pt), got {gap}pt; source:\n{source}"
            );
        }
    }
}

/// `customGeo.pptx` slide 2 gives every real body item a 12pt `a:spcAft`.
/// PowerPoint applies that gap after the paragraph, while the wrapped
/// `Curriculum` and `system` lines keep the paragraph's ordinary line advance
/// (issue #1335).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn slide_list_item_space_after_stays_out_of_wrapped_line_advance() {
    use crate::internal::{Parser, PptxParser};

    let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/customGeo.pptx");
    let data = std::fs::read(fixture).expect("customGeo fixture");
    let (mut document, _warnings) = PptxParser
        .parse(&data, &crate::config::ConvertOptions::default())
        .expect("customGeo must parse");
    let Page::Fixed(mut spaced_page) = document.pages[1].clone() else {
        panic!("slide 2 must be fixed")
    };
    spaced_page.elements.retain(|element| {
        let FixedElementKind::TextBox(text_box) = &element.kind else {
            return false;
        };
        text_box.content.iter().any(|block| {
            let Block::List(list) = block else {
                return false;
            };
            list.items.iter().any(|item| {
                item.content.iter().any(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .any(|run| run.text.contains("Curriculum"))
                })
            })
        })
    });
    assert_eq!(spaced_page.elements.len(), 1, "slide 2 has one body list");
    // Libertinus is narrower than the fixture's Calibri. Keep item 1 on one
    // line while items 2 and 3 still exercise their real natural wraps.
    spaced_page.elements[0].width = 640.0;
    let FixedElementKind::TextBox(spaced_text_box) = &mut spaced_page.elements[0].kind else {
        unreachable!("the retained element is the body text box")
    };
    for block in &mut spaced_text_box.content {
        let Block::List(list) = block else {
            continue;
        };
        for paragraph in list
            .items
            .iter_mut()
            .flat_map(|item| item.content.iter_mut())
        {
            // Calibri is not installed on every CI runner. The embedded face
            // keeps the fixture's real text, size and spacing while making
            // both the adjusted-measure wrap and its line box portable.
            paragraph.style.paragraph_mark_font_family = Some("Libertinus Serif".into());
            for run in &mut paragraph.runs {
                run.style.font_family = Some("Libertinus Serif".into());
            }
        }
    }

    let mut unspaced_page = spaced_page.clone();
    let FixedElementKind::TextBox(text_box) = &mut unspaced_page.elements[0].kind else {
        unreachable!("the retained element is the body text box")
    };
    for block in &mut text_box.content {
        let Block::List(list) = block else {
            continue;
        };
        for paragraph in list
            .items
            .iter_mut()
            .flat_map(|item| item.content.iter_mut())
        {
            paragraph.style.space_after = None;
        }
    }

    document.pages = vec![Page::Fixed(spaced_page), Page::Fixed(unspaced_page)];
    let output = generate_typst(&document).unwrap();
    let baselines = |page_index: usize| -> (f64, f64, f64, f64, f64, f64) {
        let runs = crate::render::pdf::compiled_text_runs(&output.source, page_index)
            .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
        let baseline = |needle: &str| -> f64 {
            runs.iter()
                .find(|run| run.text.contains(needle))
                .unwrap_or_else(|| panic!("missing {needle:?} on page {page_index}: {runs:?}"))
                .baseline_pt
        };
        (
            baseline("Overview"),
            baseline("Standards"),
            baseline("Description"),
            baseline("Curriculum"),
            baseline("Update"),
            baseline("system"),
        )
    };
    let spaced = baselines(0);
    let unspaced = baselines(1);
    assert!(
        (spaced.1 - spaced.0).abs() < 0.05,
        "item 1 must remain the one-line boundary control: {spaced:?}"
    );
    let first_item_boundary_advance_pt: f64 = spaced.2 - spaced.0;
    let item_2_wrapped_advance_pt: f64 = spaced.3 - spaced.2;
    let item_2_control_advance_pt: f64 = unspaced.3 - unspaced.2;
    let item_3_wrapped_advance_pt: f64 = spaced.5 - spaced.4;
    let item_3_control_advance_pt: f64 = unspaced.5 - unspaced.4;
    assert!(
        item_2_wrapped_advance_pt > 5.0 && item_3_wrapped_advance_pt > 5.0,
        "items 2 and 3 must retain their natural wraps: {spaced:?}"
    );

    assert!(
        first_item_boundary_advance_pt - item_2_wrapped_advance_pt > 5.0,
        "the 12pt spcAft boundary must stay visibly larger than item 2's internal line advance: boundary={first_item_boundary_advance_pt}pt, wrapped={item_2_wrapped_advance_pt}pt\n{}",
        output.source
    );
    assert!(
        first_item_boundary_advance_pt - item_3_wrapped_advance_pt > 5.0,
        "the 12pt spcAft boundary must stay visibly larger than item 3's internal line advance: boundary={first_item_boundary_advance_pt}pt, wrapped={item_3_wrapped_advance_pt}pt\n{}",
        output.source
    );

    assert!(
        (item_2_wrapped_advance_pt - item_2_control_advance_pt).abs() < 0.05,
        "spcAft must not change item 2's internal line advance: spaced={item_2_wrapped_advance_pt}pt, control={item_2_control_advance_pt}pt\n{}",
        output.source
    );
    assert!(
        (item_3_wrapped_advance_pt - item_3_control_advance_pt).abs() < 0.05,
        "spcAft must not change item 3's internal line advance: spaced={item_3_wrapped_advance_pt}pt, control={item_3_control_advance_pt}pt\n{}",
        output.source
    );
}

/// `customGeo.pptx` slide 23 gives every body bullet an absolute 33pt
/// `<a:lnSpc><a:spcPts>`. That value is the baseline advance of every visual
/// line, including lines PowerPoint creates by wrapping one list paragraph
/// (issue #1336).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn slide_list_exact_point_spacing_controls_every_wrapped_line() {
    use crate::internal::{Parser, PptxParser};

    let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/customGeo.pptx");
    let data = std::fs::read(fixture).expect("customGeo fixture");
    let (mut document, _warnings) = PptxParser
        .parse(&data, &crate::config::ConvertOptions::default())
        .expect("customGeo must parse");
    let Page::Fixed(mut exact_33_page) = document.pages[22].clone() else {
        panic!("slide 23 must be fixed")
    };
    exact_33_page.elements.retain(|element| {
        let FixedElementKind::TextBox(text_box) = &element.kind else {
            return false;
        };
        text_box.content.iter().any(|block| {
            let Block::List(list) = block else {
                return false;
            };
            list.items.iter().any(|item| {
                item.content.iter().any(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .any(|run| run.text.contains("Familiarize yourself"))
                })
            })
        })
    });
    assert_eq!(
        exact_33_page.elements.len(),
        1,
        "slide 23 has one body list"
    );
    exact_33_page.elements[0].x = 50.0;
    exact_33_page.elements[0].y = 50.0;
    exact_33_page.elements[0].width = 500.0;
    exact_33_page.elements[0].height = 440.0;
    let FixedElementKind::TextBox(text_box) = &mut exact_33_page.elements[0].kind else {
        unreachable!("the retained element is the body text box")
    };
    text_box.auto_fit = false;
    text_box.vertical_align = crate::ir::TextBoxVerticalAlign::Top;
    text_box
        .content
        .retain(|block| matches!(block, Block::List(_)));
    for block in &mut text_box.content {
        let Block::List(list) = block else {
            unreachable!("only the body list was retained")
        };
        list.items.truncate(1);
        let paragraph = &mut list.items[0].content[0];
        assert!(
            matches!(paragraph.style.line_spacing, Some(LineSpacing::Exact(points)) if (points - 33.0).abs() < 1e-9),
            "the fixture must preserve its exact 33pt line spacing: {:?}",
            paragraph.style.line_spacing
        );
        // Calibri is not installed on every CI runner. The embedded face keeps
        // the fixture text and exact spacing while making wrapping portable.
        paragraph.style.paragraph_mark_font_family = Some("Libertinus Serif".into());
        for run in &mut paragraph.runs {
            run.style.font_family = Some("Libertinus Serif".into());
        }
    }

    let mut exact_41_page = exact_33_page.clone();
    let FixedElementKind::TextBox(text_box) = &mut exact_41_page.elements[0].kind else {
        unreachable!("the retained element is the body text box")
    };
    for block in &mut text_box.content {
        let Block::List(list) = block else {
            unreachable!("only the body list was retained")
        };
        list.items[0].content[0].style.line_spacing = Some(LineSpacing::Exact(41.0));
    }

    document.pages = vec![Page::Fixed(exact_33_page), Page::Fixed(exact_41_page)];
    let output = generate_typst(&document).unwrap();
    let line_advances = |page_index: usize| -> Vec<f64> {
        let runs = crate::render::pdf::compiled_text_runs(&output.source, page_index)
            .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
        let mut baselines: Vec<f64> = runs
            .iter()
            .filter(|run| !run.text.trim().is_empty())
            .map(|run| run.baseline_pt)
            .collect();
        baselines.sort_by(f64::total_cmp);
        baselines.dedup_by(|left, right| (*left - *right).abs() < 0.05);
        assert!(
            baselines.len() >= 2,
            "the retained real list item must wrap: {runs:?}\n{}",
            output.source
        );
        baselines.windows(2).map(|pair| pair[1] - pair[0]).collect()
    };

    for (page_index, expected_advance_pt) in [(0, 33.0), (1, 41.0)] {
        let advances = line_advances(page_index);
        assert!(
            advances
                .iter()
                .all(|advance| (*advance - expected_advance_pt).abs() < 0.05),
            "page {page_index} must advance every wrapped line by its exact {expected_advance_pt}pt rule, got {advances:?}\n{}",
            output.source
        );
    }
}

/// Every `#v(Xpt)` skip in some markup, in points — the emitted numbers carry
/// binary noise (`43.199999999999996`), so a gap is asserted numerically.
pub(in crate::render) fn vertical_skips_pt(source: &str) -> Vec<f64> {
    source
        .split("#v(")
        .skip(1)
        .filter_map(|rest| rest.split("pt)").next()?.parse::<f64>().ok())
        .collect()
}
