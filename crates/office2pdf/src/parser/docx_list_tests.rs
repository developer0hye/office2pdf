use super::*;

#[test]
fn test_parse_simple_bulleted_list() {
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
            marker_text: None,
            marker_style: None,
        })
    );

    let text0: String = lists[0].items[0]
        .content
        .iter()
        .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
        .collect();
    assert_eq!(text0, "Item A");
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
            marker_text: None,
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
                    marker_text: None,
                    marker_style: None,
                },
            ),
        ])
    );
}

#[test]
fn test_parse_mixed_list_and_paragraphs() {
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
