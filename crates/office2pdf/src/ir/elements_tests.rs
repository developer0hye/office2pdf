use super::*;

#[test]
fn test_table_cell_default() {
    let cell = TableCell::default();
    assert_eq!(cell.col_span, 1);
    assert_eq!(cell.row_span, 1);
    assert!(cell.content.is_empty());
    assert!(cell.border.is_none());
    assert!(cell.background.is_none());
}

#[test]
fn test_list_item_default() {
    let item = ListItem {
        content: vec![Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Item 1".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }],
        level: 0,
        start_at: None,
    };
    assert_eq!(item.level, 0);
    assert_eq!(item.content.len(), 1);
}

#[test]
fn test_list_unordered() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Bullet 1".to_string(),
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
                        text: "Bullet 2".to_string(),
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
    assert_eq!(list.kind, ListKind::Unordered);
    assert_eq!(list.items.len(), 2);
}

#[test]
fn test_list_ordered() {
    let list = List {
        kind: ListKind::Ordered,
        items: vec![ListItem {
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
    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(list.items.len(), 1);
    assert_eq!(list.items[0].start_at, Some(3));
}

#[test]
fn test_list_nested() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Top".to_string(),
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
                        text: "Nested".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 1,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::from([(
            1,
            ListLevelStyle {
                kind: ListKind::Unordered,
                numbering_pattern: None,
                full_numbering: false,
                marker_text: None,
                marker_style: None,
            },
        )]),
    };
    assert_eq!(list.items[0].level, 0);
    assert_eq!(list.items[1].level, 1);
    assert_eq!(
        list.level_styles.get(&1),
        Some(&ListLevelStyle {
            kind: ListKind::Unordered,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_paragraph_with_runs() {
    let para = Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Hello ".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: "world".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
        ],
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Hello ");
    assert_eq!(para.runs[1].style.bold, Some(true));
}

#[test]
fn test_header_footer_with_text() {
    let hf = HeaderFooter {
        paragraphs: vec![HeaderFooterParagraph {
            style: ParagraphStyle::default(),
            elements: vec![HFInline::Run(Run {
                text: "My Header".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            })],
        }],
    };
    assert_eq!(hf.paragraphs.len(), 1);
    assert_eq!(hf.paragraphs[0].elements.len(), 1);
    match &hf.paragraphs[0].elements[0] {
        HFInline::Run(r) => assert_eq!(r.text, "My Header"),
        _ => panic!("Expected Run"),
    }
}

#[test]
fn test_header_footer_with_page_number() {
    let hf = HeaderFooter {
        paragraphs: vec![HeaderFooterParagraph {
            style: ParagraphStyle::default(),
            elements: vec![
                HFInline::Run(Run {
                    text: "Page ".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }),
                HFInline::PageNumber,
            ],
        }],
    };
    assert_eq!(hf.paragraphs[0].elements.len(), 2);
    assert!(matches!(hf.paragraphs[0].elements[1], HFInline::PageNumber));
}
