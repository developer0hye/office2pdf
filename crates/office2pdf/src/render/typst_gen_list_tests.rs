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
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#enum("));
    assert!(output.source.contains("start: 3"));
    assert!(output.source.contains("numbering: \"1.\""));
    assert!(output.source.contains("Step 1"));
    assert!(output.source.contains("Step 2"));
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
