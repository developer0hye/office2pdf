use super::*;
use crate::ir::*;
use std::collections::BTreeMap;
use std::io::Cursor;

/// Helper: build a minimal DOCX as bytes using docx-rs builder.
fn build_docx_bytes(paragraphs: Vec<docx_rs::Paragraph>) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new();
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with custom page size and margins.
fn build_docx_bytes_with_page_setup(
    paragraphs: Vec<docx_rs::Paragraph>,
    width_twips: u32,
    height_twips: u32,
    margin_top: i32,
    margin_bottom: i32,
    margin_left: i32,
    margin_right: i32,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new()
        .page_size(width_twips, height_twips)
        .page_margin(
            docx_rs::PageMargin::new()
                .top(margin_top)
                .bottom(margin_bottom)
                .left(margin_left)
                .right(margin_right),
        );
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: extract the first run from the first paragraph of a parsed document.
fn first_run(doc: &Document) -> &Run {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    &para.runs[0]
}

// ----- Paragraph formatting tests (US-005) -----

/// Helper: extract the first paragraph from a parsed document.
fn first_paragraph(doc: &Document) -> &Paragraph {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph block"),
    }
}

/// Helper: get all blocks from the first page.
fn all_blocks(doc: &Document) -> &[Block] {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    &page.content
}

#[path = "docx_foundation_tests.rs"]
mod foundation_tests;

// ----- Table parsing tests (US-007) -----

/// Helper: build a DOCX with a table using docx-rs builder.
fn build_docx_with_table(table: docx_rs::Table) -> Vec<u8> {
    let docx = docx_rs::Docx::new().add_table(table);
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: extract the first table block from a parsed document.
fn first_table(doc: &Document) -> &crate::ir::Table {
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    for block in &page.content {
        if let Block::Table(t) = block {
            return t;
        }
    }
    panic!("No Table block found");
}

#[path = "docx_table_tests.rs"]
mod table_tests;

#[path = "docx_image_tests.rs"]
mod image_tests;

// ----- List parsing tests -----

/// Helper: build a DOCX with numbering definitions and list paragraphs.
fn build_docx_with_numbering(
    abstract_nums: Vec<docx_rs::AbstractNumbering>,
    numberings: Vec<docx_rs::Numbering>,
    paragraphs: Vec<docx_rs::Paragraph>,
) -> Vec<u8> {
    let mut nums = docx_rs::Numberings::new();
    for an in abstract_nums {
        nums = nums.add_abstract_numbering(an);
    }
    for n in numberings {
        nums = nums.add_numbering(n);
    }

    let mut docx = docx_rs::Docx::new().numberings(nums);
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn body_bookmark_and_comment_range_markers_preserve_list_continuity() {
    #[derive(Debug, Clone, Copy)]
    enum MarkerPosition {
        BookmarkStartBetweenItems,
        BookmarkEndBetweenItems,
        CommentStartBetweenItems,
        CommentEndBetweenItems,
    }

    let markers: [MarkerPosition; 4] = [
        MarkerPosition::BookmarkStartBetweenItems,
        MarkerPosition::BookmarkEndBetweenItems,
        MarkerPosition::CommentStartBetweenItems,
        MarkerPosition::CommentEndBetweenItems,
    ];

    for marker in markers {
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ));
        let numberings = docx_rs::Numberings::new()
            .add_abstract_numbering(abstract_num)
            .add_numbering(docx_rs::Numbering::new(1, 0));
        let first_item = || {
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First item"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0))
        };
        let second_item = || {
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Second item"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0))
        };
        let document = docx_rs::Docx::new().numberings(numberings);
        let document = match marker {
            MarkerPosition::BookmarkStartBetweenItems => document
                .add_paragraph(first_item())
                .add_bookmark_start(9, "marker")
                .add_paragraph(second_item())
                .add_bookmark_end(9),
            MarkerPosition::BookmarkEndBetweenItems => document
                .add_bookmark_start(9, "marker")
                .add_paragraph(first_item())
                .add_bookmark_end(9)
                .add_paragraph(second_item()),
            MarkerPosition::CommentStartBetweenItems => {
                let mut document = document.add_paragraph(first_item());
                document.document = document
                    .document
                    .add_comment_start(docx_rs::Comment::new(9));
                document = document.add_paragraph(second_item());
                document.document = document.document.add_comment_end(9);
                document
            }
            MarkerPosition::CommentEndBetweenItems => {
                let mut document = document;
                document.document = document
                    .document
                    .add_comment_start(docx_rs::Comment::new(9));
                document = document.add_paragraph(first_item());
                document.document = document.document.add_comment_end(9);
                document.add_paragraph(second_item())
            }
        };
        let mut cursor = Cursor::new(Vec::new());
        document.build().pack(&mut cursor).unwrap();

        let parser = DocxParser;
        let (parsed, _warnings) = parser
            .parse(&cursor.into_inner(), &ConvertOptions::default())
            .unwrap();
        let page = match &parsed.pages[0] {
            Page::Flow(page) => page,
            _ => panic!("Expected FlowPage"),
        };
        let lists: Vec<&List> = page
            .content
            .iter()
            .filter_map(|block| match block {
                Block::List(list) => Some(list),
                _ => None,
            })
            .collect();

        assert_eq!(
            lists.len(),
            1,
            "body range marker {marker:?} split one numbered list into separate blocks"
        );
        assert_eq!(lists[0].items.len(), 2);
        assert_eq!(lists[0].items[0].start_at, Some(1));
        assert_eq!(
            lists[0].items[1].start_at, None,
            "the second item should continue numbering through {marker:?}"
        );
    }
}

#[test]
fn bookmark_end_in_real_libreoffice_fixture_preserves_the_four_item_list() {
    let data: &[u8] = include_bytes!("../../../../tests/fixtures/docx/libreoffice/tdf149711.docx");
    let parser = DocxParser;
    let (parsed, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    let page = match &parsed.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected FlowPage"),
    };
    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .collect();

    assert_eq!(
        lists.len(),
        1,
        "the bookmark end between list paragraphs must not split the list"
    );
    assert_eq!(lists[0].items.len(), 4);
}

#[test]
fn test_parse_simple_bulleted_list() {
    // Create a bullet list: abstractNum with format "bullet", numId=1, ilvl=0
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("bullet"),
        docx_rs::LevelText::new("•"),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item A"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item B"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item C"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    // Should produce a single List block with 3 items
    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].kind, ListKind::Unordered);
    assert_eq!(lists[0].items.len(), 3);
    assert_eq!(lists[0].items[0].level, 0);
    assert_eq!(
        lists[0].level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Unordered,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: Some("•".to_string()),
            marker_style: None,
        })
    );

    // Verify item content
    let text0: String = lists[0].items[0]
        .content
        .iter()
        .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
        .collect();
    assert_eq!(text0, "Item A");
}

#[test]
fn test_numbering_level_hanging_indent_applies_to_list_paragraphs() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(
        docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        )
        .indent(
            Some(900),
            Some(docx_rs::SpecialIndentType::Hanging(300)),
            None,
            None,
        ),
    );
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented item"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Direct override"))
                .indent(
                    Some(1200),
                    Some(docx_rs::SpecialIndentType::Hanging(200)),
                    None,
                    None,
                )
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("numbered paragraph should become a list");
    let style = &list.items[0].content[0].style;

    assert_eq!(style.indent_left, Some(45.0));
    assert_eq!(style.indent_first_line, Some(-15.0));

    let override_style = &list.items[1].content[0].style;
    assert_eq!(override_style.indent_left, Some(60.0));
    assert_eq!(override_style.indent_first_line, Some(-10.0));
}

#[test]
fn test_list_paragraphs_use_word_compatible_spacing_when_unspecified() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Default spacing"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Explicit spacing"))
                .line_spacing(docx_rs::LineSpacing::new().after(120))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("numbered paragraphs should become a list");

    assert_eq!(list.items[0].content[0].style.space_after, Some(0.0));
    assert_eq!(list.items[1].content[0].style.space_after, Some(6.0));
}

#[test]
fn test_parse_simple_numbered_list() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Second"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].kind, ListKind::Ordered);
    assert_eq!(lists[0].items.len(), 2);
    assert_eq!(lists[0].items[0].start_at, Some(1));
    assert_eq!(
        lists[0].level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Ordered,
            numbering_pattern: Some("1.".to_string()),
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_nested_multi_level_list() {
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("◦"),
            docx_rs::LevelJc::new("left"),
        ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Top level"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Nested item"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Back to top"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|b| match b {
            Block::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "Expected 1 list block");
    assert_eq!(lists[0].items.len(), 3);
    assert_eq!(lists[0].items[0].level, 0);
    assert_eq!(lists[0].items[1].level, 1);
    assert_eq!(lists[0].items[2].level, 0);
    assert_eq!(
        lists[0].level_styles.get(&1),
        Some(&ListLevelStyle {
            kind: ListKind::Unordered,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: Some("◦".to_string()),
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_numbered_list_start_override() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering =
        docx_rs::Numbering::new(1, 0).add_override(docx_rs::LevelOverride::new(0).start(3));

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Third"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Fourth"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("Expected list block");

    assert_eq!(list.items[0].start_at, Some(3));
    assert_eq!(list.items[1].start_at, None);
    assert_eq!(
        list.level_styles.get(&0),
        Some(&ListLevelStyle {
            kind: ListKind::Ordered,
            numbering_pattern: Some("1.".to_string()),
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_parse_mixed_ordered_and_bulleted_levels() {
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Step"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bullet child"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let list = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .expect("Expected list block");

    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(
        list.level_styles,
        BTreeMap::from([
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
                    marker_text: Some("•".to_string()),
                    marker_style: None,
                },
            ),
        ])
    );
}

#[test]
fn test_parse_mixed_list_and_paragraphs() {
    // A list followed by a regular paragraph should produce two separate blocks
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item 1"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Item 2"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Regular paragraph")),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    // Should have at least a List block and a Paragraph block
    let list_count = page
        .content
        .iter()
        .filter(|b| matches!(b, Block::List(_)))
        .count();
    let para_count = page
        .content
        .iter()
        .filter(|b| matches!(b, Block::Paragraph(_)))
        .count();
    assert!(list_count >= 1, "Expected at least 1 list block");
    assert!(para_count >= 1, "Expected at least 1 paragraph block");
}

#[test]
fn test_merges_adjacent_lists_with_different_num_ids() {
    // pandoc/LibreOffice fragment a single logical list across several numIds
    // (issue #176). Adjacent list paragraphs must merge into one list so ordered
    // numbering continues (1., 2.) instead of restarting, and `ilvl` nesting is
    // preserved instead of flattening into a separate bullet list.
    // One abstract: ordered level 0, bulleted level 1 — the same shape the
    // passing `test_parse_mixed_ordered_and_bulleted_levels` relies on, so its
    // resolution is trusted. Two distinct numIds both reference it, mirroring
    // the issue's document where consecutive items carry different numId values.
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("\u{2022}"),
            docx_rs::LevelJc::new("left"),
        ));

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0), docx_rs::Numbering::new(2, 0)],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Second"))
                .numbering(docx_rs::NumberingId::new(2), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Sub"))
                .numbering(docx_rs::NumberingId::new(2), docx_rs::IndentLevel::new(1)),
        ],
    );

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let lists: Vec<&List> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::List(list) => Some(list),
            _ => None,
        })
        .collect();

    assert_eq!(
        lists.len(),
        1,
        "adjacent list paragraphs must merge into a single list"
    );
    let list = lists[0];
    assert_eq!(list.kind, ListKind::Ordered);
    assert_eq!(list.items.len(), 3);
    assert_eq!(list.items[0].level, 0);
    assert_eq!(list.items[0].start_at, Some(1));
    assert_eq!(list.items[1].level, 0);
    assert_eq!(
        list.items[1].start_at, None,
        "the second ordered item continues counting (-> 2.), it must not restart"
    );
    assert_eq!(
        list.items[2].level, 1,
        "the sub-item stays nested at level 1"
    );
    assert_eq!(
        list.level_styles.get(&0).map(|style| style.kind),
        Some(ListKind::Ordered)
    );
    assert_eq!(
        list.level_styles.get(&1).map(|style| style.kind),
        Some(ListKind::Unordered)
    );
}

#[test]
fn test_preserves_empty_paragraph_after_drawing_only_anchor() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="column"><wp:posOffset>366395</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>141605</wp:posOffset></wp:positionV>
<wp:extent cx="1590675" cy="733425"/>
<wp:wrapNone/>
<wp:docPr id="1" name="Shape 1"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="1590840" cy="733320"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="729fcf"/></a:solidFill>
<a:ln w="0"><a:solidFill><a:srgbClr val="3465a4"/></a:solidFill></a:ln>
</wps:spPr></wps:wsp>
</a:graphicData></a:graphic>
</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>
<w:p><w:r><w:t>After drawing</w:t></w:r></w:p>
<w:sectPr/>
</w:body></w:document>"#;

    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(
            &build_docx_with_math(document_xml),
            &ConvertOptions::default(),
        )
        .unwrap();
    let blocks = all_blocks(&doc);

    assert!(
        matches!(blocks[0], Block::FloatingShape(_)),
        "drawing-only paragraph should emit the floating shape first"
    );
    assert!(
        matches!(&blocks[1], Block::Paragraph(paragraph) if paragraph.runs.is_empty()),
        "drawing-only paragraph mark must remain as an empty paragraph spacer"
    );
    assert!(
        matches!(&blocks[2], Block::Paragraph(paragraph) if paragraph.runs[0].text == "After drawing"),
        "following content should stay after the preserved paragraph mark"
    );
}

#[path = "docx_page_feature_tests.rs"]
mod page_feature_tests;

// ----- Document styles tests (US-022) -----

/// Helper: build a DOCX with custom styles and paragraphs.
fn build_docx_bytes_with_styles(
    paragraphs: Vec<docx_rs::Paragraph>,
    styles: Vec<docx_rs::Style>,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new();
    for s in styles {
        docx = docx.add_style(s);
    }
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Helper: build a DOCX with an explicit stylesheet and paragraphs.
fn build_docx_bytes_with_stylesheet(
    paragraphs: Vec<docx_rs::Paragraph>,
    styles: docx_rs::Styles,
) -> Vec<u8> {
    let mut docx = docx_rs::Docx::new().styles(styles);
    for p in paragraphs {
        docx = docx.add_paragraph(p);
    }
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[path = "docx_style_tests.rs"]
mod style_tests;

// ----- Hyperlink tests (US-030) -----

#[test]
fn test_hyperlink_single_link_in_paragraph() {
    let link = docx_rs::Hyperlink::new("https://example.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Click here"));
    let para = docx_rs::Paragraph::new().add_hyperlink(link);
    let data = build_docx_bytes(vec![para]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };

    assert_eq!(para.runs.len(), 1);
    assert_eq!(para.runs[0].text, "Click here");
    assert_eq!(para.runs[0].href, Some("https://example.com".to_string()));
}

#[test]
fn test_hyperlink_mixed_text_and_link() {
    let link = docx_rs::Hyperlink::new("https://rust-lang.org", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Rust"));
    let para = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Visit "))
        .add_hyperlink(link)
        .add_run(docx_rs::Run::new().add_text(" for more."));
    let data = build_docx_bytes(vec![para]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };

    // Should have 3 runs: "Visit ", hyperlink "Rust", " for more."
    assert_eq!(para.runs.len(), 3);

    assert_eq!(para.runs[0].text, "Visit ");
    assert_eq!(para.runs[0].href, None);

    assert_eq!(para.runs[1].text, "Rust");
    assert_eq!(para.runs[1].href, Some("https://rust-lang.org".to_string()));

    assert_eq!(para.runs[2].text, " for more.");
    assert_eq!(para.runs[2].href, None);
}

#[test]
fn test_hyperlink_multiple_links_in_paragraph() {
    let link1 = docx_rs::Hyperlink::new("https://first.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("First"));
    let link2 = docx_rs::Hyperlink::new("https://second.com", docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Second"));
    let para = docx_rs::Paragraph::new()
        .add_hyperlink(link1)
        .add_run(docx_rs::Run::new().add_text(" and "))
        .add_hyperlink(link2);
    let data = build_docx_bytes(vec![para]);

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let para = match &page.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };

    assert_eq!(para.runs.len(), 3);

    assert_eq!(para.runs[0].text, "First");
    assert_eq!(para.runs[0].href, Some("https://first.com".to_string()));

    assert_eq!(para.runs[1].text, " and ");
    assert_eq!(para.runs[1].href, None);

    assert_eq!(para.runs[2].text, "Second");
    assert_eq!(para.runs[2].href, Some("https://second.com".to_string()));
}

#[path = "docx_notes_textbox_tests.rs"]
mod notes_textbox_tests;

// ── OMML math equation tests ──

/// Build a DOCX ZIP with a custom document.xml containing OMML math.
fn build_docx_with_math(document_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )
    .unwrap();

    // word/document.xml
    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

/// Helper: build a DOCX from raw document.xml using the minimal ZIP scaffold.
fn build_docx_with_columns(document_xml: &str) -> Vec<u8> {
    build_docx_with_math(document_xml)
}

/// Helper: build a DOCX from raw `document.xml`, `styles.xml`, and
/// `footnotes.xml`.
fn build_docx_with_notes_xml(document_xml: &str, styles_xml: &str, footnotes_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#,
        )
        .unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#,
        )
        .unwrap();

    for (name, contents) in [
        ("word/document.xml", document_xml),
        ("word/styles.xml", styles_xml),
        ("word/footnotes.xml", footnotes_xml),
    ] {
        zip.start_file(name, options).unwrap();
        std::io::Write::write_all(&mut zip, contents.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

/// Helper: build a DOCX from raw `document.xml` and `styles.xml`.
///
/// docx-rs's `Styles` builder keeps `doc_defaults` private, so a test that
/// needs to author `w:docDefaults` — the only place a document can state the
/// justification and line spacing every paragraph falls back to — has to write
/// the part itself.
fn build_docx_with_styles_xml(document_xml: &str, styles_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#,
        )
        .unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
        )
        .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.start_file("word/styles.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, styles_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[path = "docx_contextual_spacing_tests.rs"]
mod contextual_spacing_tests;
#[path = "docx_layout_rtl_tests.rs"]
mod layout_rtl_tests;
#[path = "docx_math_chart_metadata_tests.rs"]
mod math_chart_metadata_tests;
#[path = "docx_paragraph_cursor_tests.rs"]
mod paragraph_cursor_tests;
#[path = "docx_tracked_paragraph_mark_tests.rs"]
mod tracked_paragraph_mark_tests;

#[test]
fn issue_189_footer_preserves_inline_image_and_rtl_text() {
    let data = include_bytes!("../../../../tests/fixtures/docx/issue_189_footer_image_rtl.docx");
    let parser = DocxParser;
    let (document, warnings) = parser
        .parse(data, &ConvertOptions::default())
        .expect("issue #189 fixture should parse");

    assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected flow page");
    };
    let footer = page.footer.as_ref().expect("default footer");

    assert_eq!(footer.paragraphs.len(), 3);
    assert_eq!(footer.paragraphs[0].elements.len(), 1, "footer image");
    assert_eq!(footer.paragraphs[1].elements.len(), 1, "French footer text");
    assert_eq!(
        footer.paragraphs[2].style.direction,
        Some(TextDirection::Rtl),
        "Arabic footer paragraph direction"
    );
    let footer_text: String = footer
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.elements)
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect();
    assert!(footer_text.contains("Généré par m3llm.cafe"));
    assert!(footer_text.contains("صنع بواسطة m3llm.cafe"));

    let typst = crate::render::typst_gen::generate_typst(&document)
        .expect("issue #189 fixture should generate Typst");
    assert_eq!(typst.images.len(), 1, "footer image asset");
    assert!(typst.source.contains("#image(\"img-0.png\""));
    assert!(typst.source.contains("#text(dir: rtl)["));
    // w:bottom="1800" (90pt) with w:footer="708" (35.4pt) pins the footer's
    // bottom 54.6pt above the bottom margin line; the image paragraph grows the
    // footer upward from there.
    assert!(typst.source.contains("footer-descent: 0pt"));
    assert!(typst.source.contains("block(width: 100%, height: 54.6pt)"));
    assert!(typst.source.contains("place(bottom"));
    assert!(typst.source.contains("Généré par m3llm.cafe"));
    assert!(typst.source.contains("صنع بواسطة m3llm.cafe"));

    let result = crate::convert_bytes(data, crate::Format::Docx, &ConvertOptions::default())
        .expect("issue #189 fixture should convert to PDF");
    let pdf_text = pdf_extract::extract_text_from_mem(&result.pdf)
        .expect("issue #189 PDF text should extract");
    assert!(pdf_text.contains("Généré par m3llm.cafe"));
    // PDF text extraction exposes RTL glyphs in visual order with layout spacing.
    assert!(pdf_text.contains("عنص"), "extracted PDF text: {pdf_text:?}");
    assert!(
        pdf_text.contains("ةطساوب"),
        "extracted PDF text: {pdf_text:?}"
    );
}

#[test]
fn test_bullet_level_marker_glyph_preserved() {
    // Word's level-2 bullets use the numbering definition's lvlText glyph
    // (e.g. ○); dropping it made every level render the level-1 disc
    // (issue #356).
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("○"),
            docx_rs::LevelJc::new("left"),
        ));
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("level one"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("level two"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
        ],
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let list = doc
        .pages
        .iter()
        .find_map(|page| match page {
            Page::Flow(flow) => flow.content.iter().find_map(|block| match block {
                Block::List(list) => Some(list),
                _ => None,
            }),
            _ => None,
        })
        .expect("expected a list block");
    assert_eq!(
        list.level_styles
            .get(&1)
            .and_then(|style| style.marker_text.as_deref()),
        Some("○"),
        "level-2 bullet glyph must come from lvlText"
    );
}

#[test]
fn test_zero_indent_numbering_renders_inline_number_with_tab() {
    // Korean clause numbering (제N조) uses w:ind left=0 hanging=0: Word puts
    // the number inline, follows it with a tab, and wraps continuation
    // lines back to the margin. Rendering these as hanging-indent lists
    // indented every continuation line (issue #357).
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(
        docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("제%1조"),
            docx_rs::LevelJc::new("left"),
        )
        .indent(
            Some(0),
            Some(docx_rs::SpecialIndentType::Hanging(0)),
            None,
            None,
        ),
    );
    let numbering = docx_rs::Numbering::new(1, 0);

    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![numbering],
        vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("(목적) 본 계약은"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("(계약 기간) 본 계약의"))
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("expected flow page"),
    };
    let paragraphs: Vec<&Paragraph> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    assert_eq!(
        paragraphs.len(),
        2,
        "flush numbering must not become a list"
    );
    assert!(
        paragraphs[0].runs[0].text.starts_with("제1조\t"),
        "first clause number inline with tab: {:?}",
        paragraphs[0].runs[0].text
    );
    assert!(
        paragraphs[1].runs[0].text.starts_with("제2조\t"),
        "second clause number advances: {:?}",
        paragraphs[1].runs[0].text
    );
}

/// Rewrites `word/settings.xml` inside a DOCX, replacing the
/// `w:defaultTabStop` element with `replacement` (empty string removes it).
fn rewrite_settings_default_tab_stop(docx_bytes: &[u8], replacement: &str) -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(docx_bytes.to_vec())).expect("read zip");
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let name: String = file.name().to_string();
        let mut content: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut content).expect("read entry");
        if name == "word/settings.xml" {
            let xml = String::from_utf8(content).expect("settings utf8");
            let rewritten = match (xml.find("<w:defaultTabStop"), xml.find("/>")) {
                (Some(start), Some(_)) => {
                    let end = xml[start..].find("/>").expect("self-closing") + start + 2;
                    format!("{}{}{}", &xml[..start], replacement, &xml[end..])
                }
                _ => xml,
            };
            content = rewritten.into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("start entry");
        std::io::Write::write_all(&mut out, &content).expect("write entry");
    }
    out.finish().expect("finish zip").into_inner()
}

#[test]
fn test_explicit_default_tab_stop_is_parsed() {
    // Korean Word writes w:defaultTabStop val="800" (40pt); honoring it
    // places suffix tabs at Word's stops (issue #393).
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("본문")),
    ]);
    let data = rewrite_settings_default_tab_stop(&data, r#"<w:defaultTabStop w:val="800"/>"#);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.styles.default_tab_stop_pt, Some(40.0));
}

#[test]
fn test_absent_default_tab_stop_is_none() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("body")),
    ]);
    let data = rewrite_settings_default_tab_stop(&data, "");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.styles.default_tab_stop_pt, None);
}

/// Rewrites the `compatibilityMode` compatibility setting inside a DOCX's
/// `word/settings.xml`, replacing the whole `w:compat` element with
/// `replacement` (empty string removes it).
fn rewrite_settings_compat(docx_bytes: &[u8], replacement: &str) -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(docx_bytes.to_vec())).expect("read zip");
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let name: String = file.name().to_string();
        let mut content: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut content).expect("read entry");
        if name == "word/settings.xml" {
            let xml = String::from_utf8(content).expect("settings utf8");
            let start: usize = xml.find("<w:compat>").expect("docx-rs writes w:compat");
            let end: usize = xml.find("</w:compat>").expect("closing tag") + "</w:compat>".len();
            content = format!("{}{}{}", &xml[..start], replacement, &xml[end..]).into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("start entry");
        std::io::Write::write_all(&mut out, &content).expect("write entry");
    }
    out.finish().expect("finish zip").into_inner()
}

/// The `w:compat` element a package declaring `mode` carries. `value_first`
/// selects the attribute order native Word writes, which is not the order
/// docx-rs writes, so both have to read the same (issue #1130).
fn compat_element(mode: u32, value_first: bool) -> String {
    const URI: &str = "http://schemas.microsoft.com/office/word";
    let attributes: String = if value_first {
        format!(r#"w:val="{mode}" w:name="compatibilityMode" w:uri="{URI}""#)
    } else {
        format!(r#"w:name="compatibilityMode" w:uri="{URI}" w:val="{mode}""#)
    };
    format!("<w:compat><w:compatSetting {attributes}/></w:compat>")
}

/// Word 2013's layout engine, which native exports of every business-corpus
/// source declare.
#[test]
fn test_compatibility_mode_15_reads_as_word_2013() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("본문")),
    ]);
    let data = rewrite_settings_compat(&data, &compat_element(15, true));
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(
        doc.styles.word_compatibility_mode,
        Some(WordCompatibilityMode::Word2013OrLater)
    );
}

#[test]
fn test_compatibility_mode_before_15_reads_as_legacy() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("본문")),
    ]);
    let data = rewrite_settings_compat(&data, &compat_element(14, false));
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(
        doc.styles.word_compatibility_mode,
        Some(WordCompatibilityMode::Legacy)
    );
}

/// A package that states no mode is a pre-2013 document to Word, which is the
/// shape every hand-authored fixture in this repository has.
#[test]
fn test_absent_compatibility_mode_reads_as_legacy() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("본문")),
    ]);
    let data = rewrite_settings_compat(&data, "");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(
        doc.styles.word_compatibility_mode,
        Some(WordCompatibilityMode::Legacy)
    );
}
