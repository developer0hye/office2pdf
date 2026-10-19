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

#[path = "docx_namespace_tests.rs"]
mod namespace_tests;

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

#[path = "docx_image_recovery_tests.rs"]
mod image_recovery_tests;

#[path = "docx_smartart_tests.rs"]
mod smartart_tests;

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
        BookmarkStart,
        BookmarkEnd,
        CommentStart,
        CommentEnd,
    }

    let markers: [MarkerPosition; 4] = [
        MarkerPosition::BookmarkStart,
        MarkerPosition::BookmarkEnd,
        MarkerPosition::CommentStart,
        MarkerPosition::CommentEnd,
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
            MarkerPosition::BookmarkStart => document
                .add_paragraph(first_item())
                .add_bookmark_start(9, "marker")
                .add_paragraph(second_item())
                .add_bookmark_end(9),
            MarkerPosition::BookmarkEnd => document
                .add_bookmark_start(9, "marker")
                .add_paragraph(first_item())
                .add_bookmark_end(9)
                .add_paragraph(second_item()),
            MarkerPosition::CommentStart => {
                let mut document = document.add_paragraph(first_item());
                document.document = document
                    .document
                    .add_comment_start(docx_rs::Comment::new(9));
                document = document.add_paragraph(second_item());
                document.document = document.document.add_comment_end(9);
                document
            }
            MarkerPosition::CommentEnd => {
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
            numbering_pattern: Some(NumberingPattern::parse("1.")),
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        })
    );
}

#[test]
fn test_leading_layout_break_keeps_numbered_paragraph_as_a_list() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let paragraph = docx_rs::Paragraph::new()
        .add_run(
            docx_rs::Run::new()
                .add_break(docx_rs::BreakType::Column)
                .add_text("Numbered continuation"),
        )
        .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0));
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        vec![paragraph],
    );

    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("Expected a flow page");
    };
    assert!(matches!(page.content.first(), Some(Block::ColumnBreak)));
    let Some(Block::List(list)) = page.content.get(1) else {
        panic!(
            "a paragraph after a leading column break must stay numbered: {:#?}",
            page.content
        );
    };
    assert_eq!(list.items.len(), 1);
    assert_eq!(list.items[0].start_at, Some(1));
    assert_eq!(
        list.items[0].content[0].runs[0].text,
        "Numbered continuation"
    );
}

#[test]
fn test_empty_numbered_paragraph_remains_a_list_item() {
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let paragraph = docx_rs::Paragraph::new()
        .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0));
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        vec![paragraph],
    );

    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("Expected a flow page");
    };
    let Some(Block::List(list)) = page.content.first() else {
        panic!(
            "an empty numbered paragraph must remain a list item: {:#?}",
            page.content
        );
    };
    assert_eq!(list.items.len(), 1);
    assert_eq!(list.items[0].start_at, Some(1));
    assert!(list.items[0].content[0].runs.is_empty());
}

#[test]
fn test_numbered_paragraph_frame_survives_a_leading_column_break() {
    const FIXTURE: &[u8] = include_bytes!(
        "../../../../tests/fixtures/docx/libreoffice/tdf153964_numberingAfterBreak14.docx"
    );
    let (document, _warnings) = DocxParser
        .parse(FIXTURE, &ConvertOptions::default())
        .unwrap();
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;

    assert!(
        source.matches("fill: rgb(226, 239, 217)").count() >= 2,
        "the visible numbered paragraph must retain its shading in addition to its empty break fragment: {source}"
    );
    assert!(
        source.contains("if tab_prefix_width_1 < 17.85pt"),
        "the inline number must place its suffix tab at the paragraph's 17.85pt text indent: {source}"
    );
    assert!(
        source.contains(", outset: (x: 5.44pt)"),
        "the paragraph border space must extend the frame outside the text measure: {source}"
    );
    assert!(
        source.contains(", inset: (top: 2.5pt, bottom: 2.5pt)"),
        "horizontal border spacing must not reduce the available line width: {source}"
    );
    let second_page_start: usize = source
        .find("#pagebreak()")
        .expect("the fixture starts its next page with a page break");
    let second_page: &str = &source[second_page_start..];
    let columns_start: usize = second_page
        .find("#columns(2,")
        .expect("the page-break continuation uses two columns");
    let before_spacing: usize = second_page
        .find("#v(52pt, weak: false)")
        .expect("the continued paragraph keeps before-spacing after paragraph-after collapse");
    assert!(
        before_spacing > columns_start,
        "page-top paragraph spacing must stay in the first column rather than shifting every column: {second_page}"
    );
}

#[test]
fn test_separate_paragraph_at_later_column_page_top_drops_before_spacing() {
    const FIXTURE: &[u8] = include_bytes!(
        "../../../../tests/fixtures/docx/libreoffice/tdf153964_topMarginAfterBreak14.docx"
    );
    let (document, _warnings) = DocxParser
        .parse(FIXTURE, &ConvertOptions::default())
        .unwrap();
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    let text: usize = source
        .find("page break1")
        .expect("page-break paragraph emitted");
    let page_break: usize = source[..text]
        .rfind("#pagebreak()")
        .expect("the paragraph follows a page break");
    let page_segment: &str = &source[page_break..text];

    assert!(
        !page_segment.contains("#v(60pt, weak: false)") && !page_segment.contains("above: 60pt"),
        "before-spacing from a later column-break paragraph must not move the separate paragraph at the page top: {page_segment}"
    );
    assert!(
        source.contains(", outset: (x: 5.44pt)")
            && source.contains(", inset: (top: 2.5pt, bottom: 2.5pt)"),
        "horizontal paragraph-border spacing must expand the frame without narrowing text measure: {source}"
    );
}

#[test]
fn test_empty_shaded_paragraphs_keep_their_before_spacing_across_breaks() {
    const FIXTURE: &[u8] = include_bytes!(
        "../../../../tests/fixtures/docx/libreoffice/tdf153964_topMarginAfterBreak14.docx"
    );
    let (document, _warnings) = DocxParser
        .parse(FIXTURE, &ConvertOptions::default())
        .unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected one flow page");
    };
    let red_fragments: Vec<&Paragraph> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph)
                if paragraph.runs.is_empty()
                    && paragraph.style.background == Some(Color::new(0xFF, 0, 0)) =>
            {
                Some(paragraph)
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        red_fragments.len(),
        2,
        "the empty red paragraph has one painted fragment before each layout break"
    );
    let paragraph_mark_style: &TextStyle = red_fragments[0]
        .style
        .paragraph_mark_text_style
        .as_deref()
        .expect("DOCX empty paragraphs retain their resolved paragraph-mark style");
    assert_eq!(paragraph_mark_style.font_family.as_deref(), Some("Calibri"));
    assert_eq!(paragraph_mark_style.font_size, Some(11.0));
    assert!(
        red_fragments
            .iter()
            .all(|paragraph| paragraph.style.space_before.is_none()),
        "break-only fragments must not carry a 60pt before-gap into the next column or page"
    );
    assert!(
        red_fragments[0].style.border.is_none(),
        "an earlier empty fragment paints shading without a paragraph border"
    );
    assert!(
        red_fragments[1].style.border.is_some(),
        "the last empty fragment retains the paragraph border"
    );

    let purple_fragment: &Paragraph = page
        .content
        .windows(2)
        .find_map(|pair| match (&pair[0], &pair[1]) {
            (Block::Paragraph(paragraph), Block::PageBreak)
                if paragraph.runs.is_empty()
                    && paragraph.style.background == Some(Color::new(0x70, 0x30, 0xA0)) =>
            {
                Some(paragraph)
            }
            _ => None,
        })
        .expect("the purple empty paragraph ends with its page break");
    assert!(
        purple_fragment.style.border.is_none(),
        "the prefix before a page break paints shading without a paragraph border"
    );

    let yellow_fragments: Vec<&Paragraph> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph)
                if paragraph.runs.is_empty()
                    && paragraph.style.background == Some(Color::new(0xFF, 0xF2, 0xCC)) =>
            {
                Some(paragraph)
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        yellow_fragments.len(),
        2,
        "the empty yellow paragraph paints both the column-break prefix and its continuation"
    );
    assert!(yellow_fragments[0].style.border.is_none());
    assert!(yellow_fragments[1].style.border.is_some());

    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;
    let yellow_color: &str = "fill: rgb(255, 242, 204)";
    let yellow_prefix_color: usize = source
        .find(yellow_color)
        .expect("the yellow prefix precedes its column break");
    let yellow_prefix_start: usize = source[..yellow_prefix_color]
        .rfind("#block(")
        .expect("the yellow prefix has its own paragraph block");
    let yellow_column_break: usize = source[yellow_prefix_start..]
        .find("#colbreak()")
        .map(|relative_break| yellow_prefix_start + relative_break)
        .expect("the yellow paragraph mark continues in the next column");
    let yellow_continuation_color: usize = source[yellow_column_break..]
        .find(yellow_color)
        .map(|relative_color| yellow_column_break + relative_color)
        .expect("the second yellow fragment is after its column break");
    let yellow_continuation_start: usize = source[..yellow_continuation_color]
        .rfind("#block(")
        .expect("the yellow continuation is in its own paragraph block");
    let preceding_column_break: usize = source[..yellow_continuation_start]
        .rfind("#colbreak()")
        .expect("the yellow paragraph mark continues in the next column");
    assert!(
        source[preceding_column_break..yellow_continuation_start].contains("#v(52pt, weak: false)"),
        "the empty paragraph's before-spacing places its continuation after the column break: {source}"
    );
    let yellow_prefix_header_end: usize = source[yellow_prefix_start..]
        .find(")[\n")
        .map(|relative_end| yellow_prefix_start + relative_end)
        .expect("the yellow prefix block header closes");
    let yellow_continuation_header_end: usize = source[yellow_continuation_start..]
        .find(")[\n")
        .map(|relative_end| yellow_continuation_start + relative_end)
        .expect("the yellow continuation block header closes");
    assert!(
        !source[yellow_prefix_start..yellow_prefix_header_end].contains("stroke:")
            && source[yellow_continuation_start..yellow_continuation_header_end]
                .contains("stroke:"),
        "the prefix has no painted border while the final fragment retains the paragraph border: {source}"
    );
    assert!(
        source.contains("height: 8pt, above: 0pt, below: 0pt, fill: rgb(112, 48, 160)")
            && source.contains("height: 8pt, above: 0pt, below: 0pt, fill: rgb(226, 239, 217)"),
        "empty paragraph shading extends through after-spacing at a page break: {source}"
    );
    assert!(
        source.contains(
            "dy: -54.5pt, rect(width: 100% + 10.88pt, height: 52pt, fill: rgb(226, 239, 217)"
        ),
        "the standalone green paragraph shades the collapsed before-spacing outside its border inset: {source}"
    );
}

#[test]
fn test_paragraph_starting_with_page_break_keeps_before_spacing_on_later_page() {
    let before_break = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_text("Paragraph before the explicit page break"));
    let after_break = docx_rs::Paragraph::new()
        .add_run(
            docx_rs::Run::new()
                .add_break(docx_rs::BreakType::Page)
                .add_text("Paragraph content after the explicit page break"),
        )
        .line_spacing(docx_rs::LineSpacing::new().before(1200));
    let data = build_docx_bytes(vec![before_break, after_break]);

    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("Expected a flow page");
    };
    assert!(page.content.iter().any(|block| matches!(
        block,
        Block::Paragraph(paragraph) if paragraph.style.starts_after_layout_break
    )));
    let source = crate::render::typst_gen::generate_typst(&document)
        .unwrap()
        .source;

    assert!(
        source.contains("#pagebreak()\n\n#v(60pt, weak: false)"),
        "before-spacing on a paragraph whose own first run is a page break must survive at the new page top: {source}"
    );
}

#[test]
fn test_layout_break_continuation_does_not_repeat_paragraph_edges_or_first_line_indent() {
    let paragraph = docx_rs::Paragraph::new()
        .add_run(
            docx_rs::Run::new()
                .add_text("Before break")
                .add_break(docx_rs::BreakType::Page)
                .add_text("After break"),
        )
        .line_spacing(docx_rs::LineSpacing::new().before(1200).after(120))
        .indent(
            None,
            Some(docx_rs::SpecialIndentType::FirstLine(1134)),
            None,
            None,
        );
    let data = build_docx_bytes(vec![paragraph]);

    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("Expected a flow page");
    };
    let [
        Block::Paragraph(first),
        Block::PageBreak,
        Block::Paragraph(continuation),
    ] = page.content.as_slice()
    else {
        panic!(
            "expected two paragraph fragments around the page break: {:#?}",
            page.content
        );
    };

    assert_eq!(first.runs[0].text, "Before break");
    assert_eq!(first.style.space_before, Some(60.0));
    assert_eq!(first.style.space_after, None);
    assert_eq!(first.style.indent_first_line, Some(56.7));
    assert!(!first.style.starts_after_layout_break);
    assert_eq!(continuation.runs[0].text, "After break");
    assert_eq!(continuation.style.space_before, None);
    assert_eq!(continuation.style.space_after, Some(6.0));
    assert_eq!(continuation.style.indent_first_line, None);
    assert!(!continuation.style.starts_after_layout_break);
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
            numbering_pattern: Some(NumberingPattern::parse("1.")),
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
                    numbering_pattern: Some(NumberingPattern::parse("1.")),
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
    let options = zip::write::SimpleFileOptions::default();

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
    let options = zip::write::SimpleFileOptions::default();

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
    let options = zip::write::SimpleFileOptions::default();

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

/// A numbering level with a hanging indent, as Word writes list levels.
fn hanging_list_level(
    level: usize,
    number_format: &str,
    level_text: &str,
    left_twips: i32,
    hanging_twips: i32,
) -> docx_rs::Level {
    docx_rs::Level::new(
        level,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new(number_format),
        docx_rs::LevelText::new(level_text),
        docx_rs::LevelJc::new("left"),
    )
    .indent(
        Some(left_twips),
        Some(docx_rs::SpecialIndentType::Hanging(hanging_twips)),
        None,
        None,
    )
}

/// Compiles a DOCX and returns, for each body text, the label painted before
/// it on its baseline and the left edge of the run that paints the body.
#[cfg(not(target_arch = "wasm32"))]
fn painted_list_labels(data: &[u8], bodies: &[&str]) -> Vec<(String, f64)> {
    let (doc, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&doc).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    bodies
        .iter()
        .map(|body| {
            let body_run = runs
                .iter()
                .find(|run| run.text.contains(body))
                .unwrap_or_else(|| panic!("{body:?} is not painted: {runs:?}"));
            let mut label_runs: Vec<_> = runs
                .iter()
                .filter(|run| {
                    (run.baseline_pt - body_run.baseline_pt).abs() < 0.5
                        && run.left_pt < body_run.left_pt
                })
                .collect();
            label_runs.sort_by(|left, right| left.left_pt.total_cmp(&right.left_pt));
            let mut label: String = label_runs.iter().map(|run| run.text.as_str()).collect();
            label.push_str(&body_run.text[..body_run.text.find(body).unwrap()]);
            (label.trim().to_string(), body_run.left_pt)
        })
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_level_text_literal_text_prints_as_written() {
    // Typst reads every `1`, `a`, `A`, `i`, `I`, `*`, ... in a numbering
    // pattern as a counter, but Word prints lvlText's literal text as written:
    // `1.%1` is `1.1`, `1.2`, never `1`, `2` (issue #2221).
    let bodies = ["Alpha", "Beta", "Gamma"];
    let mut body_lefts: BTreeMap<&str, Vec<f64>> = BTreeMap::new();
    // A 9pt hanging indent is narrower than `1.1`, so the body's tab follows
    // the label's measured width. A label with a space gets room for itself
    // (issue #2322).
    for (level_text, number_format, hanging_twips, expected) in [
        ("1.%1", "decimal", 180, ["1.1", "1.2", "1.3"]),
        // `9` is no counting symbol, so this control already prints right.
        ("9.%1", "decimal", 180, ["9.1", "9.2", "9.3"]),
        (
            "Article %1",
            "upperRoman",
            1440,
            ["Article I", "Article II", "Article III"],
        ),
        ("*%1)", "lowerLetter", 180, ["*a)", "*b)", "*c)"]),
        // Without a placeholder, lvlText is static text on every item.
        ("1)", "decimal", 180, ["1)", "1)", "1)"]),
        ("*", "lowerLetter", 180, ["*", "*", "*"]),
        ("", "decimal", 180, ["", "", ""]),
        ("", "none", 180, ["", "", ""]),
    ] {
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(hanging_list_level(
            0,
            number_format,
            level_text,
            1440,
            hanging_twips,
        ));
        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![docx_rs::Numbering::new(1, 0)],
            bodies
                .iter()
                .map(|body| {
                    docx_rs::Paragraph::new()
                        .add_run(docx_rs::Run::new().add_text(*body))
                        .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0))
                })
                .collect(),
        );
        let painted = painted_list_labels(&data, &bodies);
        let labels: Vec<&str> = painted.iter().map(|(label, _)| label.as_str()).collect();
        assert_eq!(labels, expected, "lvlText {level_text:?}");
        body_lefts.insert(level_text, painted.iter().map(|(_, left)| *left).collect());
    }
    // `1.1` and `9.1` are equally wide, so their bodies tab to the same stop.
    for (literal, control) in body_lefts["1.%1"].iter().zip(&body_lefts["9.%1"]) {
        assert!(
            (literal - control).abs() < 0.01,
            "body after `1.N` at {literal}pt, after `9.N` at {control}pt"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_level_text_literal_text_survives_parent_numbers() {
    let abstract_num = docx_rs::AbstractNumbering::new(0)
        .add_level(hanging_list_level(0, "decimal", "1.%1", 720, 360))
        .add_level(hanging_list_level(1, "lowerLetter", "1.%1.%2", 1440, 360));
    let items = [
        (0, "Alpha"),
        (1, "Child one"),
        (1, "Child two"),
        (0, "Beta"),
    ];
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        items
            .iter()
            .map(|(level, body)| {
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text(*body))
                    .numbering(
                        docx_rs::NumberingId::new(1),
                        docx_rs::IndentLevel::new(*level),
                    )
            })
            .collect(),
    );
    let bodies: Vec<&str> = items.iter().map(|(_, body)| *body).collect();
    let labels: Vec<String> = painted_list_labels(&data, &bodies)
        .into_iter()
        .map(|(label, _)| label)
        .collect();
    assert_eq!(labels, ["1.1", "1.1.a", "1.1.b", "1.2"]);
}

/// Rewrites `word/settings.xml` inside a DOCX, replacing the
/// `w:defaultTabStop` element with `replacement` (empty string removes it).
fn rewrite_settings_default_tab_stop(docx_bytes: &[u8], replacement: &str) -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(docx_bytes.to_vec())).expect("read zip");
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let name: String = file.name().expect("entry name should decode").into_owned();
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
        out.start_file(name, zip::write::SimpleFileOptions::default())
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

const LANGUAGE_TEST_DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Aftale</w:t></w:r></w:p></w:body></w:document>"#;

/// A `word/styles.xml` with `doc_defaults_rpr` inside `w:rPrDefault` and
/// `normal_rpr` inside the default paragraph style.
fn language_test_styles_xml(doc_defaults_rpr: &str, normal_rpr: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>{doc_defaults_rpr}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>{normal_rpr}</w:rPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr><w:lang w:val="fr-FR"/></w:rPr></w:style></w:styles>"#
    )
}

/// Danish Word writes the editing language into the document defaults.
#[test]
fn test_document_language_comes_from_doc_defaults() {
    let styles_xml = language_test_styles_xml(
        r#"<w:sz w:val="22"/><w:lang w:val="da-DK" w:eastAsia="en-US" w:bidi="ar-SA"/>"#,
        "",
    );
    let data = build_docx_with_styles_xml(LANGUAGE_TEST_DOCUMENT_XML, &styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(
        doc.styles.default_language,
        Some(DocumentLanguage {
            language: "da".to_string(),
            region: Some("DK".to_string()),
        })
    );
}

/// Body text takes the default paragraph style over the document defaults,
/// and a style no body text defaults to does not decide the language.
#[test]
fn test_default_paragraph_style_language_overrides_doc_defaults() {
    let styles_xml =
        language_test_styles_xml(r#"<w:lang w:val="en-US"/>"#, r#"<w:lang w:val="nb-NO"/>"#);
    let data = build_docx_with_styles_xml(LANGUAGE_TEST_DOCUMENT_XML, &styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(
        doc.styles.default_language,
        Some(DocumentLanguage {
            language: "nb".to_string(),
            region: Some("NO".to_string()),
        })
    );
}

/// A document that states no language leaves the renderer's default alone.
#[test]
fn test_document_without_language_has_none() {
    let styles_xml = language_test_styles_xml(r#"<w:sz w:val="22"/>"#, "");
    let data = build_docx_with_styles_xml(LANGUAGE_TEST_DOCUMENT_XML, &styles_xml);
    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.styles.default_language, None);
}

#[test]
fn test_document_language_ignores_historical_run_properties() {
    for (current, historical) in [("de-DE", "en-US"), ("da-DK", "fr-FR")] {
        let current_properties = format!(r#"<w:lang w:val="{current}"/>"#);
        let historical_properties = format!(
            r#"<w:rPrChange w:id="1" w:author="Synthetic reviewer" w:date="2026-01-01T00:00:00Z"><w:rPr><w:lang w:val="{historical}"/></w:rPr></w:rPrChange>"#
        );
        for normal_properties in [
            format!("{current_properties}{historical_properties}"),
            historical_properties.clone(),
        ] {
            let xml = language_test_styles_xml(&current_properties, &normal_properties);
            assert_eq!(
                styles::scan_default_language(&xml, "Normal"),
                DocumentLanguage::from_office_tag(current),
                "historical properties must not replace the current style or document default"
            );
        }
    }
}

#[test]
fn test_document_language_ignores_paragraph_mark_language() {
    let xml = language_test_styles_xml(r#"<w:lang w:val="da-DK"/>"#, "").replace(
        "<w:name w:val=\"Normal\"/>",
        r#"<w:name w:val="Normal"/><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>"#,
    );
    assert_eq!(
        styles::scan_default_language(&xml, "Normal"),
        DocumentLanguage::from_office_tag("da-DK")
    );
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
        let name: String = file.name().expect("entry name should decode").into_owned();
        let mut content: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut content).expect("read entry");
        if name == "word/settings.xml" {
            let xml = String::from_utf8(content).expect("settings utf8");
            let start: usize = xml.find("<w:compat>").expect("docx-rs writes w:compat");
            let end: usize = xml.find("</w:compat>").expect("closing tag") + "</w:compat>".len();
            content = format!("{}{}{}", &xml[..start], replacement, &xml[end..]).into_bytes();
        }
        out.start_file(name, zip::write::SimpleFileOptions::default())
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

/// Word hyphenates only a document whose settings carry
/// `w:autoHyphenation`; docx-rs writes none.
#[test]
fn test_auto_hyphenation_setting_is_read() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("body")),
    ]);
    let data = rewrite_settings_default_tab_stop(
        &data,
        r#"<w:autoHyphenation/><w:defaultTabStop w:val="720"/>"#,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(doc.styles.hyphenates_automatically);
}

#[test]
fn test_absent_auto_hyphenation_reads_as_off() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("body")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(!doc.styles.hyphenates_automatically);
}

/// `w:val` is an ST_OnOff; Word writes the element with an explicit off value
/// after the option has been switched on and off again.
#[test]
fn test_auto_hyphenation_switched_off_reads_as_off() {
    for value in ["false", "0", "off"] {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("body")),
        ]);
        let data = rewrite_settings_default_tab_stop(
            &data,
            &format!(r#"<w:autoHyphenation w:val="{value}"/><w:defaultTabStop w:val="720"/>"#),
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        assert!(
            !doc.styles.hyphenates_automatically,
            "w:val=\"{value}\" must read as off"
        );
    }
}

#[test]
fn test_auto_hyphenation_uses_xml_values_and_settings_scope() {
    let base = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Quarterly report")),
    ]);
    for (setting, expected) in [
        (r#"<w:autoHyphenation w:val='false'/>"#, false),
        (r#"<w:autoHyphenation w:val = "off"/>"#, false),
        (r#"<w:autoHyphenation w:val="&#48;"/>"#, false),
        (r#"<!-- <w:autoHyphenation/> -->"#, false),
        (r#"<w:compat><w:autoHyphenation/></w:compat>"#, false),
        (
            r#"<other:autoHyphenation xmlns:other="urn:extension"/>"#,
            false,
        ),
        (
            r#"<x:autoHyphenation xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" x:val='on'/>"#,
            true,
        ),
        (
            r#"<x:autoHyphenation xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" x:val='0'/>"#,
            false,
        ),
        (
            r#"<w:autoHyphenation w:val='true'></w:autoHyphenation>"#,
            true,
        ),
    ] {
        let data = rewrite_settings_default_tab_stop(&base, setting);
        assert_eq!(extract_auto_hyphenation(&data), expected, "{setting}");
    }
}

/// The upstream synthetic fixture isolates section decorations from body text.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_docx_page_borders_from_word_reference_are_rendered() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/page_borders.docx");
    let (document, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let paints =
        crate::render::pdf::compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let blue = typst::visualize::Color::from_u8(31, 78, 121, 255);
    let edges: Vec<_> = paints
        .iter()
        .filter(|paint| {
            paint
                .stroke
                .as_ref()
                .is_some_and(|stroke| stroke.color.as_ref() == Some(&blue))
        })
        .collect();
    assert_eq!(
        edges.len(),
        4,
        "all four declared page-border edges must paint"
    );
    for edge in &edges {
        assert!((edge.stroke.as_ref().unwrap().thickness_pt - 3.0).abs() < 0.01);
    }
    let expected = [
        (24.0, 25.5, 588.0, 25.5),
        (24.0, 766.5, 588.0, 766.5),
        (25.5, 25.5, 25.5, 766.5),
        (586.5, 25.5, 586.5, 766.5),
    ];
    for target in expected {
        assert!(
            edges.iter().any(|edge| {
                let b = edge.bounds;
                (b.0 - target.0).abs() < 0.01
                    && (b.1 - target.1).abs() < 0.01
                    && (b.2 - target.2).abs() < 0.01
                    && (b.3 - target.3).abs() < 0.01
            }),
            "missing declared page-space edge {target:?}: {edges:?}"
        );
    }
}

#[test]
fn test_docx_page_borders_scan_current_sections_and_resolve_namespaces() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-word"><w:body>
    <w:p><w:pPr><w:sectPr><w:pgBorders w:display="firstPage" w:zOrder="back"><w:top w:val="dashed" w:sz="8" w:space="10" w:color="E02030"/><w:right w:val="nil"/></w:pgBorders>
    <w:sectPrChange><w:sectPr><w:pgBorders><w:left w:val="single" w:sz="80"/></w:pgBorders></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>
    <w:sectPr/><x:sectPr><x:pgBorders><x:top x:val="single"/></x:pgBorders></x:sectPr>
    </w:body></w:document>"#;
    let sections = contexts::scan_page_borders(xml);
    assert_eq!(sections.len(), 2);
    let border = sections[0].as_ref().unwrap();
    assert_eq!(border.display, PageBorderDisplay::FirstPage);
    assert!(border.behind_text);
    assert!(!border.offset_from_page);
    assert_eq!(border.spacing.top, 10.0);
    let top = border.edges.top.as_ref().unwrap();
    assert_eq!(top.width, 1.0);
    assert_eq!(top.color, Color::new(224, 32, 48));
    assert_eq!(top.style, BorderLineStyle::Dashed);
    assert!(border.edges.right.is_none());
    assert!(border.edges.left.is_none());
    assert!(sections[1].is_none());
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_docx_page_borders_display_and_section_reset_affect_actual_pages() {
    use crate::render::pdf::{PaintedKind, compiled_page_paint_sequences};
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/page_borders.docx");
    let (original, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(base) = &original.pages[0] else {
        panic!("expected a flow section")
    };
    let blue = typst::visualize::Color::from_u8(31, 78, 121, 255);
    for (display, expected) in [
        (PageBorderDisplay::AllPages, [4, 4]),
        (PageBorderDisplay::FirstPage, [4, 0]),
        (PageBorderDisplay::NotFirstPage, [0, 4]),
    ] {
        let mut page = base.clone();
        page.page_borders.as_mut().unwrap().display = display;
        page.content.push(Block::PageBreak);
        page.content.extend(base.content.clone());
        let mut doc = original.clone();
        doc.pages = vec![Page::Flow(page)];
        let output = crate::render::typst_gen::generate_typst(&doc).unwrap();
        let pages = compiled_page_paint_sequences(&output.source, &output.images).unwrap();
        assert_eq!(pages.len(), 2);
        let counts: Vec<_> = pages
            .iter()
            .map(|page| {
                page.iter()
                    .filter(|paint| {
                        paint.kind == PaintedKind::Shape
                            && paint
                                .stroke
                                .as_ref()
                                .is_some_and(|stroke| stroke.color.as_ref() == Some(&blue))
                    })
                    .count()
            })
            .collect();
        assert_eq!(counts, expected, "display mode {display:?}");
    }
    let mut plain = base.clone();
    plain.page_borders = None;
    let mut doc = original;
    doc.pages.push(Page::Flow(plain));
    let output = crate::render::typst_gen::generate_typst(&doc).unwrap();
    let pages = compiled_page_paint_sequences(&output.source, &output.images).unwrap();
    assert_eq!(pages.len(), 2);
    assert!(
        !pages[1].iter().any(|paint| paint
            .stroke
            .as_ref()
            .is_some_and(|stroke| stroke.color.as_ref() == Some(&blue))),
        "a section without borders must not inherit the preceding decoration"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_docx_page_borders_text_offsets_partial_edges_and_paint_order() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/page_borders.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("expected flow page")
    };
    page.margins = Margins {
        top: 40.0,
        left: 50.0,
        bottom: 55.0,
        right: 65.0,
    };
    let borders = page.page_borders.as_mut().unwrap();
    borders.offset_from_page = false;
    borders.behind_text = true;
    borders.spacing.top = 6.0;
    borders.spacing.left = 8.0;
    borders.edges.bottom = None;
    borders.edges.right = None;
    let red = Color::new(160, 40, 90);
    for (side, width) in [
        (&mut borders.edges.top, 1.0),
        (&mut borders.edges.left, 2.0),
    ] {
        let side = side.as_mut().unwrap();
        side.width = width;
        side.color = red;
    }
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let red = typst::visualize::Color::from_u8(160, 40, 90, 255);
    let edges: Vec<_> = paints
        .iter()
        .enumerate()
        .filter(|(_, paint)| {
            paint
                .stroke
                .as_ref()
                .is_some_and(|stroke| stroke.color.as_ref() == Some(&red))
        })
        .collect();
    assert_eq!(edges.len(), 2);
    let expected = [(40.0, 33.5, 547.0, 33.5), (41.0, 33.5, 41.0, 737.0)];
    for bounds in expected {
        assert!(
            edges.iter().any(|(_, edge)| {
                let b = edge.bounds;
                (b.0 - bounds.0).abs() < 0.01
                    && (b.1 - bounds.1).abs() < 0.01
                    && (b.2 - bounds.2).abs() < 0.01
                    && (b.3 - bounds.3).abs() < 0.01
            }),
            "missing text-relative border {bounds:?}: {edges:?}"
        );
    }
    let first_text = paints
        .iter()
        .position(|paint| paint.kind == PaintedKind::Text)
        .unwrap();
    assert!(
        edges.iter().all(|(index, _)| *index < first_text),
        "back borders paint before the body"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_docx_page_borders_back_layer_stays_behind_header_artwork() {
    use crate::render::pdf::compiled_paint_sequence;
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/page_borders.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("expected flow page")
    };
    page.page_borders.as_mut().unwrap().behind_text = true;
    page.header = Some(HeaderFooter {
        paragraphs: Vec::new(),
        distance_from_edge: None,
        sheet_print_scale: None,
        shapes: vec![HeaderFooterShape {
            content: HeaderFooterShapeContent::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(240, 120, 20)),
                gradient_fill: None,
                pattern_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: None,
                shadow: None,
                top_bevel: None,
            }),
            frame: HeaderFooterFrame {
                x: Some(10.0),
                y: Some(10.0),
                width: Some(80.0),
                height: Some(80.0),
                horizontal_anchor: FrameAnchor::Page,
                vertical_anchor: FrameAnchor::Page,
                horizontal_align: None,
                vertical_align: None,
                inset_left: 0.0,
                inset_top: 0.0,
                bottom_offset: None,
                wraps_text: true,
            },
            anchor_paragraph_index: None,
            width: 80.0,
            height: 80.0,
            behind_text: true,
        }],
    });
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let blue = typst::visualize::Color::from_u8(31, 78, 121, 255);
    let orange = typst::visualize::Color::from_u8(240, 120, 20, 255);
    let banner = paints
        .iter()
        .position(|paint| paint.rectangle_fill.as_ref() == Some(&orange))
        .unwrap();
    let borders: Vec<_> = paints
        .iter()
        .enumerate()
        .filter(|(_, paint)| {
            paint
                .stroke
                .as_ref()
                .is_some_and(|stroke| stroke.color.as_ref() == Some(&blue))
        })
        .collect();
    assert_eq!(borders.len(), 4);
    assert!(
        borders.iter().all(|(index, _)| *index < banner),
        "back borders must not cover the page's header artwork"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_section_line_numbers_are_painted_in_the_margin() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/line_numbering.docx");
    let pdf = crate::convert_bytes(
        data,
        crate::config::Format::Docx,
        &ConvertOptions::default(),
    )
    .unwrap();
    let text = pdf_extract::extract_text_from_mem(&pdf.pdf).unwrap();
    let separate_numbers: Vec<&str> = text
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && line.chars().all(|c| c.is_ascii_digit()))
        .collect();
    assert_eq!(
        separate_numbers,
        [
            "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"
        ],
        "line numbers must be separate searchable margin text: {text}"
    );
}

#[test]
fn docx_line_numbering_only_reads_current_word_sections() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:body>
      <w:p><w:pPr><w:sectPr><w:lnNumType w:start="7" w:countBy="3" w:distance="240" w:restart="newSection"/>
        <w:sectPrChange><w:sectPr><w:lnNumType w:start="99"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>
      <w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr><w:lnNumType/></w:sectPr></w:pPr></w:p></w:tc></w:tr></w:tbl>
      <w:sectPr><x:lnNumType x:start="99"/></w:sectPr><x:sectPr><w:lnNumType/></x:sectPr>
      <w:sectPr><w:lnNumType w:countBy="1" w:distance="NaN"/></w:sectPr>
    </w:body></w:document>"#;
    for xml in [
        xml.to_string(),
        xml.replace(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ),
    ] {
        let sections = contexts::scan_line_numbering(&xml);
        assert_eq!(
            sections,
            vec![
                Some(LineNumbering {
                    count_by: 3,
                    start: 7,
                    distance: Some(12.0),
                    restart: LineNumberRestart::NewSection
                }),
                None,
                Some(LineNumbering {
                    count_by: 1,
                    start: 0,
                    distance: None,
                    restart: LineNumberRestart::NewPage
                })
            ]
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn line_number_probe_document() -> Document {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/line_numbering.docx");
    DocxParser
        .parse(data, &ConvertOptions::default())
        .unwrap()
        .0
}

#[cfg(not(target_arch = "wasm32"))]
fn compiled_margin_numbers(
    document: &Document,
    page: usize,
) -> Vec<crate::render::pdf::PlacedTextRun> {
    let output = crate::render::typst_gen::generate_typst(document).unwrap();
    crate::render::pdf::compiled_text_runs(&output.source, page)
        .unwrap()
        .into_iter()
        .filter(|run| {
            run.left_pt < 80.0
                && !run.text.is_empty()
                && run.text.chars().all(|c| c.is_ascii_digit())
        })
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_line_numbers_obey_page_restarts_intervals_and_distance() {
    for (restart, start, count_by, expected) in [
        (LineNumberRestart::NewPage, 1, 1, vec!["2", "3", "4"]),
        (LineNumberRestart::NewSection, 1, 1, vec!["5", "6", "7"]),
        (LineNumberRestart::Continuous, 1, 1, vec!["5", "6", "7"]),
        (LineNumberRestart::NewPage, 0, 2, vec!["2"]),
    ] {
        let mut document = line_number_probe_document();
        let Page::Flow(page) = &mut document.pages[0] else {
            panic!("flow section")
        };
        page.content.truncate(3);
        let next = page.content.clone();
        page.content.push(Block::PageBreak);
        page.content.extend(next);
        page.line_numbering = Some(LineNumbering {
            restart,
            start,
            count_by,
            distance: Some(24.0),
        });
        let runs = compiled_margin_numbers(&document, 1);
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            expected,
            "{restart:?}"
        );
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        let body = crate::render::pdf::compiled_text_runs(&output.source, 1)
            .unwrap()
            .into_iter()
            .find(|run| run.left_pt >= 90.0 && run.text.contains("Numbered"))
            .unwrap();
        assert!(
            runs.iter().all(|run| run.left_pt < 66.0
                && run.family == body.family
                && run.font_size_pt == body.font_size_pt),
            "{runs:?}; body {body:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_line_numbers_change_at_continuous_section_boundaries() {
    let mut document = line_number_probe_document();
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.content.truncate(2);
    page.line_numbering.as_mut().unwrap().restart = LineNumberRestart::NewSection;
    let mut next = page.clone();
    next.line_numbering.as_mut().unwrap().start = 8;
    let mut plain = page.clone();
    plain.line_numbering.as_mut().unwrap().count_by = 0;
    document.pages.push(Page::FlowContinuous(next));
    document.pages.push(Page::FlowContinuous(plain));
    let numbers = compiled_margin_numbers(&document, 0);
    assert_eq!(
        numbers
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>(),
        ["2", "3", "9", "10"]
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_line_number_page_start_matches_native_word() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/line_numbering_page_restart.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let numbers = compiled_margin_numbers(&document, 0);
    assert_eq!(
        numbers
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>(),
        [
            "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"
        ]
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_line_number_default_start_displays_one() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>First numbered line of a filing.</w:t></w:r></w:p><w:sectPr><w:pgMar w:left="1800" w:right="1800" w:top="1440" w:bottom="1440"/><w:lnNumType w:countBy="1"/></w:sectPr></w:body></w:document>"#;
    let data = build_docx_with_math(xml);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let numbers = compiled_margin_numbers(&document, 0);
    assert_eq!(
        numbers
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>(),
        ["1"]
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_line_numbers_stay_disabled_without_a_positive_interval() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/line_numbering_disabled.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    assert!(compiled_margin_numbers(&document, 0).is_empty());
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    assert_eq!(page.content.len(), 12);
    page.line_numbering.as_mut().unwrap().count_by = 0;
    assert!(compiled_margin_numbers(&document, 0).is_empty());
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.line_numbering.as_mut().unwrap().count_by = 5;
    assert_eq!(
        compiled_margin_numbers(&document, 0)
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>(),
        ["5", "10"]
    );
}

#[cfg(not(target_arch = "wasm32"))]
fn compiled_section_titles(
    document: &Document,
    page: usize,
) -> (bool, Vec<crate::render::pdf::PlacedTextRun>) {
    let output = crate::render::typst_gen::generate_typst(document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, page).unwrap();
    let uses_native_fonts = runs.iter().all(|run| run.family == "Cambria");
    (
        uses_native_fonts,
        runs.into_iter()
            .filter(|run| run.text == "Vertically Centered Title")
            .collect(),
    )
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_centered_section_matches_native_title_page() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_vertical_center.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let (native_fonts, centered) = compiled_section_titles(&document, 0);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.vertical_alignment = Some(PageVerticalAlignment::Top);
    let (_, top) = compiled_section_titles(&document, 0);
    assert!(centered[0].baseline_pt - top[0].baseline_pt > 200.0);
    if native_fonts {
        assert!(
            (centered[0].baseline_pt - 377.76002).abs() < 0.5,
            "{:?}",
            centered[0]
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_continuous_sections_use_a_uniform_page_alignment() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_alignment_continuous.docx");
    let (mut top_document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    for page in &mut top_document.pages {
        let (Page::Flow(page) | Page::FlowContinuous(page)) = page else {
            panic!("flow section")
        };
        page.vertical_alignment = Some(PageVerticalAlignment::Top);
    }
    let (_, top_titles) = compiled_section_titles(&top_document, 0);
    let mut center_baselines = Vec::new();
    for (first, second, expected) in [
        (
            PageVerticalAlignment::Center,
            PageVerticalAlignment::Center,
            [336.72002, 418.80002],
        ),
        (
            PageVerticalAlignment::Bottom,
            PageVerticalAlignment::Bottom,
            [578.88002, 660.72002],
        ),
        (
            PageVerticalAlignment::Center,
            PageVerticalAlignment::Top,
            [94.80002, 176.88002],
        ),
        (
            PageVerticalAlignment::Top,
            PageVerticalAlignment::Center,
            [94.80002, 176.88002],
        ),
        (
            PageVerticalAlignment::Center,
            PageVerticalAlignment::Bottom,
            [94.80002, 176.88002],
        ),
    ] {
        let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
        assert_eq!(document.pages.len(), 2);
        assert!(matches!(document.pages[1], Page::FlowContinuous(_)));
        for (page, alignment) in document.pages.iter_mut().zip([first, second]) {
            let (Page::Flow(page) | Page::FlowContinuous(page)) = page else {
                panic!("flow section")
            };
            page.vertical_alignment = Some(alignment);
        }
        let (native_fonts, titles) = compiled_section_titles(&document, 0);
        assert_eq!(titles.len(), 2);
        for (i, (title, expected)) in titles.iter().zip(expected).enumerate() {
            if native_fonts {
                assert!(
                    (title.baseline_pt - expected).abs() < 0.5,
                    "{first:?}/{second:?}: {title:?}"
                );
            }
            if first != second {
                assert!((title.baseline_pt - top_titles[i].baseline_pt).abs() < 0.001);
            } else if first == PageVerticalAlignment::Center {
                assert!(title.baseline_pt - top_titles[i].baseline_pt > 200.0);
                center_baselines.push(title.baseline_pt);
            } else {
                assert!(
                    (title.baseline_pt
                        - top_titles[i].baseline_pt
                        - 2.0 * (center_baselines[i] - top_titles[i].baseline_pt))
                        .abs()
                        < 0.001
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_section_alignment_applies_on_each_page_and_resets() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_vertical_center.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let (_, centered) = compiled_section_titles(&document, 0);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.vertical_alignment = Some(PageVerticalAlignment::Top);
    let (native_fonts, top) = compiled_section_titles(&document, 0);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.vertical_alignment = Some(PageVerticalAlignment::Bottom);
    let (_, bottom) = compiled_section_titles(&document, 0);
    assert!(
        (bottom[0].baseline_pt
            - top[0].baseline_pt
            - 2.0 * (centered[0].baseline_pt - top[0].baseline_pt))
            .abs()
            < 0.001
    );
    if native_fonts {
        assert!((bottom[0].baseline_pt - 660.72002).abs() < 0.5);
    }
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.vertical_alignment = Some(PageVerticalAlignment::Center);
    let copy = page.content.clone();
    page.content.push(Block::PageBreak);
    page.content.extend(copy);
    let (_, second_page) = compiled_section_titles(&document, 1);
    assert!((second_page[0].baseline_pt - centered[0].baseline_pt).abs() < 0.001);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.content.truncate(2);
    let mut plain = page.clone();
    plain.vertical_alignment = None;
    document.pages.push(Page::Flow(plain));
    let (_, reset) = compiled_section_titles(&document, 1);
    assert!((reset[0].baseline_pt - top[0].baseline_pt).abs() < 0.001);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_section_alignment_counts_declared_paragraph_spaces() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_vertical_center.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let (_, original) = compiled_section_titles(&document, 0);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    let Block::Paragraph(body) = &mut page.content[1] else {
        panic!("body paragraph")
    };
    body.style.space_before = Some(20.0);
    body.style.space_after = Some(30.0);
    let (native_fonts, changed) = compiled_section_titles(&document, 0);
    assert!((original[0].baseline_pt - changed[0].baseline_pt - 15.0).abs() < 0.001);
    if native_fonts {
        assert!((changed[0].baseline_pt - 362.88002).abs() < 0.5);
    }
}

#[test]
fn docx_section_alignment_reads_only_current_namespace_properties() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:body>
        <w:p><w:pPr><w:sectPr><w:vAlign w:val="center"/><w:sectPrChange><w:sectPr><w:vAlign w:val="bottom"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>
        <w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr><w:vAlign w:val="bottom"/></w:sectPr></w:pPr></w:p></w:tc></w:tr></w:tbl>
        <w:sectPr><w:vAlign x:val="bottom"/></w:sectPr><x:sectPr><w:vAlign w:val="bottom"/></x:sectPr>
        <w:sectPr><w:vAlign w:val="bottom"/></w:sectPr><w:sectPr><w:vAlign w:val="both"/></w:sectPr>
    </w:body></w:document>"#;
    for xml in [
        xml.to_string(),
        xml.replace(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ),
    ] {
        assert_eq!(
            contexts::scan_section_alignment(&xml),
            [
                Some(PageVerticalAlignment::Center),
                None,
                Some(PageVerticalAlignment::Bottom),
                Some(PageVerticalAlignment::Justified)
            ]
        );
    }
}

#[test]
fn docx_section_vertical_justification_warns_without_losing_text() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Filing body retained.</w:t></w:r></w:p><w:sectPr><w:vAlign w:val="both"/></w:sectPr></w:body></w:document>"#;
    let data = build_docx_with_math(xml);
    let (document, warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(
        warnings
            .iter()
            .any(|warning| warning.to_string().contains("vertical justification"))
    );
    let Page::Flow(page) = &document.pages[0] else {
        panic!("flow section")
    };
    assert!(page.content.iter().any(|block| matches!(block,Block::Paragraph(paragraph) if paragraph.runs.iter().any(|run|run.text == "Filing body retained."))));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_section_alignment_keeps_header_and_footer_anchors() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_vertical_center.docx");
    let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let story = |label: &str| HeaderFooter {
        distance_from_edge: Some(36.0),
        sheet_print_scale: None,
        shapes: Vec::new(),
        paragraphs: vec![HeaderFooterParagraph {
            style: ParagraphStyle {
                space_after: Some(20.0),
                ..ParagraphStyle::default()
            },
            elements: vec![HFInline::Run(Run {
                text: label.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            })],
            border: None,
            border_space: None,
            frame: None,
            sheet_section_is_rich: false,
        }],
    };
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.header = Some(story("Filing header"));
    page.footer = Some(story("Filing footer"));
    page.vertical_alignment = Some(PageVerticalAlignment::Top);
    let positions = |doc: &Document| {
        let output = crate::render::typst_gen::generate_typst(doc).unwrap();
        crate::render::pdf::compiled_text_runs(&output.source, 0)
            .unwrap()
            .into_iter()
            .filter(|run| run.text.starts_with("Filing "))
            .map(|run| (run.text, run.baseline_pt))
            .collect::<Vec<_>>()
    };
    let top = positions(&document);
    assert_eq!(top.len(), 2);
    let Page::Flow(page) = &mut document.pages[0] else {
        panic!("flow section")
    };
    page.vertical_alignment = Some(PageVerticalAlignment::Center);
    let centered = positions(&document);
    assert_eq!(top, centered);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_mixed_continuous_alignment_only_changes_the_shared_page() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_alignment_multipage.docx");
    let (document, warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    assert!(
        warnings
            .iter()
            .any(|warning| warning.to_string().contains("automatic pagination"))
    );
    let original =
        include_bytes!("../../../../tests/fixtures/docx/docxide/section_vertical_center.docx");
    let (original, _) = DocxParser
        .parse(original, &ConvertOptions::default())
        .unwrap();
    let (_, centered) = compiled_section_titles(&original, 0);
    let (native_fonts, first) = compiled_section_titles(&document, 0);
    let (_, second) = compiled_section_titles(&document, 1);
    assert_eq!(first.len(), 1);
    assert_eq!(second.len(), 2);
    assert!((first[0].baseline_pt - centered[0].baseline_pt).abs() < 0.001);
    if native_fonts {
        assert!((second[0].baseline_pt - 94.80002).abs() < 0.5);
        assert!((second[1].baseline_pt - 176.88002).abs() < 0.5);
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn rewrite_numbering_part(data: &[u8], edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    rewrite_docx_part(data, "word/numbering.xml", edit)
}

fn rewrite_docx_part(data: &[u8], part: &str, edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut edit = Some(edit);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().expect("entry name should decode").into_owned();
        let mut content = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut content).unwrap();
        if name == part {
            content = edit.take().unwrap()(&String::from_utf8(content).unwrap()).into_bytes();
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut output, &content).unwrap();
    }
    output.finish().unwrap().into_inner()
}

#[cfg(not(target_arch = "wasm32"))]
fn converted_numbering_text(data: &[u8]) -> String {
    let result = crate::convert_bytes(
        data,
        crate::config::Format::Docx,
        &ConvertOptions::default(),
    )
    .unwrap();
    pdf_extract::extract_text_from_mem(&result.pdf).unwrap()
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_numbering_never_restart_keeps_child_labels_across_parents() {
    let original =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_restarts.docx");
    for start in [1, 7] {
        for never_restarts in [false, true] {
            for inline in [false, true] {
                let data = rewrite_numbering_part(original, |xml| {
                    let (prefix, custom) = xml
                        .split_once("<w:abstractNum w:abstractNumId=\"100\">")
                        .unwrap();
                    let (parent, child) = custom.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
                    let child = child.replacen(
                        "<w:start w:val=\"1\"/>",
                        &format!("<w:start w:val=\"{start}\"/>"),
                        1,
                    );
                    let mut xml = format!(
                        "{prefix}<w:abstractNum w:abstractNumId=\"100\">{parent}<w:lvl w:ilvl=\"1\">{child}"
                    );
                    if !never_restarts {
                        xml = xml.replace("<w:lvlRestart w:val=\"0\"/>", "");
                    }
                    if inline {
                        xml = xml
                            .replace(
                                "w:left=\"360\" w:hanging=\"360\"",
                                "w:left=\"0\" w:hanging=\"0\"",
                            )
                            .replace(
                                "w:left=\"720\" w:hanging=\"360\"",
                                "w:left=\"0\" w:hanging=\"0\"",
                            );
                    }
                    xml
                });
                let text = converted_numbering_text(&data);
                let labels: Vec<String> = text
                    .split('(')
                    .skip(1)
                    .filter_map(|part| part.split_once(')').map(|(value, _)| format!("({value})")))
                    .collect();
                let expected: Vec<String> = (0..4)
                    .map(|index| {
                        format!(
                            "({})",
                            start + if never_restarts { index } else { index % 2 }
                        )
                    })
                    .collect();
                assert_eq!(
                    labels, expected,
                    "start {start}, never {never_restarts}, inline {inline}: {text}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_numbering_restart_uses_the_requested_earlier_level() {
    let original =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_restart_policy.docx");
    for (restart, expected) in [
        ("<w:lvlRestart w:val=\"0\"/>", [1, 2, 3]),
        ("", [1, 1, 1]),
        ("<w:lvlRestart w:val=\"1\"/>", [1, 2, 1]),
    ] {
        let data = rewrite_numbering_part(original, |xml| {
            xml.replace("<w:lvlRestart w:val=\"0\"/>", restart)
        });
        let text = converted_numbering_text(&data);
        let normalized: String = text.split_whitespace().collect();
        for (item, number) in ['C', 'E', 'G'].into_iter().zip(expected) {
            assert!(
                normalized.contains(&format!("({number})Item{item}")),
                "restart {restart}: {text}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_numbering_restart_in_level_overrides_keeps_the_abstract_policy() {
    let original =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_restarts.docx");
    for abstract_never_restarts in [false, true] {
        let data = rewrite_numbering_part(original, |xml| {
            let level = xml
                .split_once("<w:lvl w:ilvl=\"1\">")
                .unwrap()
                .1
                .split_once("</w:lvl>")
                .unwrap()
                .0;
            let override_level = format!("<w:lvl w:ilvl=\"1\">{level}</w:lvl>");
            let mut xml = if abstract_never_restarts {
                xml.to_string()
            } else {
                xml.replace("<w:lvlRestart w:val=\"0\"/>", "")
            };
            let override_level = if abstract_never_restarts {
                override_level.replace("<w:lvlRestart w:val=\"0\"/>", "<w:lvlRestart w:val=\"1\"/>")
            } else {
                override_level
            };
            xml = xml.replace("<w:num w:numId=\"100\"><w:abstractNumId w:val=\"100\"/></w:num>", &format!("<w:num w:numId=\"100\"><w:abstractNumId w:val=\"100\"/><w:lvlOverride w:ilvl=\"1\">{override_level}</w:lvlOverride></w:num>"));
            xml
        });
        let text = converted_numbering_text(&data);
        let labels: Vec<String> = text
            .split('(')
            .skip(1)
            .filter_map(|part| part.split_once(')').map(|(value, _)| format!("({value})")))
            .collect();
        let expected: Vec<String> = (0..4)
            .map(|index| {
                format!(
                    "({})",
                    1 + if abstract_never_restarts {
                        index
                    } else {
                        index % 2
                    }
                )
            })
            .collect();
        assert_eq!(
            labels, expected,
            "abstract never {abstract_never_restarts}: {text}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn compiled_legal_list_baselines(document: &Document) -> (Vec<f64>, String) {
    let output = crate::render::typst_gen::generate_typst(document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let labels = [
        "First section",
        "Sub one",
        "Sub two",
        "Second section",
        "Sub one again",
    ];
    let baselines: Vec<f64> = labels
        .iter()
        .map(|label| {
            runs.iter()
                .find(|run| run.text.trim() == *label)
                .unwrap_or_else(|| panic!("missing {label}: {runs:?}\n{}", output.source))
                .baseline_pt
        })
        .collect();
    (baselines, output.source)
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_nested_list_spacing_survives_every_level_transition() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    let (original, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let positions = |gap: f64, missing_font_metrics: bool| {
        let mut document = original.clone();
        let Page::Flow(page) = &mut document.pages[0] else {
            panic!("flow section")
        };
        for block in &mut page.content {
            if let Block::List(list) = block {
                for item in &mut list.items {
                    for paragraph in &mut item.content {
                        paragraph.style.space_before = Some(0.0);
                        paragraph.style.space_after = Some(gap);
                        if missing_font_metrics {
                            for run in &mut paragraph.runs {
                                run.style.font_family = Some("Unavailable filing font".to_string());
                            }
                        }
                    }
                }
            }
        }
        compiled_legal_list_baselines(&document)
    };
    for missing_font_metrics in [false, true] {
        let (unspaced, _) = positions(0.0, missing_font_metrics);
        for gap in [7.5, 10.0, 20.0] {
            let (spaced, source) = positions(gap, missing_font_metrics);
            for index in 1..spaced.len() {
                let added = (spaced[index] - spaced[0]) - (unspaced[index] - unspaced[0]);
                let expected = index as f64 * gap;
                assert!(
                    (added - expected).abs() < 0.15,
                    "paragraph {index}, missing metrics {missing_font_metrics}: expected {expected}pt accumulated gap, got {added}pt; {source}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_nested_list_spacing_collapses_actual_adjacent_paragraph_gaps() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    let (original, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let before: [f64; 5] = [0.0, 2.0, 12.0, 1.0, 9.0];
    let after: [f64; 5] = [3.0, 8.0, 0.0, 5.0, 17.0];
    for levels in [[0, 1, 1, 0, 1], [0, 1, 2, 1, 0], [0, 0, 1, 0, 1]] {
        let positions = |spaced: bool| {
            let mut document = original.clone();
            let Page::Flow(page) = &mut document.pages[0] else {
                panic!("flow section")
            };
            let list = page
                .content
                .iter_mut()
                .find_map(|block| match block {
                    Block::List(list) => Some(list),
                    _ => None,
                })
                .unwrap();
            list.level_styles.insert(2, list.level_styles[&1].clone());
            for (index, item) in list.items.iter_mut().enumerate() {
                item.level = levels[index];
                for paragraph in &mut item.content {
                    paragraph.style.space_before = Some(if spaced { before[index] } else { 0.0 });
                    paragraph.style.space_after = Some(if spaced { after[index] } else { 0.0 });
                }
            }
            compiled_legal_list_baselines(&document)
        };
        let (unspaced, _) = positions(false);
        let (spaced, source) = positions(true);
        let mut expected: f64 = 0.0;
        for index in 1..spaced.len() {
            expected += after[index - 1].max(before[index]);
            let added = (spaced[index] - spaced[0]) - (unspaced[index] - unspaced[0]);
            assert!(
                (added - expected).abs() < 0.15,
                "levels {levels:?}, paragraph {index}: expected {expected}pt, got {added}pt; {source}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_legal_numbering_uses_decimal_references_only_when_enabled() {
    let original = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    for (flag, enabled) in [
        ("<w:isLgl/>", true),
        ("<w:isLgl w:val=\"true\"/>", true),
        ("<w:isLgl w:val=\"1\"/>", true),
        ("<w:isLgl w:val=\"on\"/>", true),
        ("<w:isLgl w:val=\"false\"/>", false),
        ("<w:isLgl w:val=\"0\"/>", false),
        ("<w:isLgl w:val=\"off\"/>", false),
        ("", false),
        ("<f:isLgl xmlns:f=\"urn:foreign\"/>", false),
        ("<w:isLgl xmlns:f=\"urn:foreign\" f:val=\"false\"/>", true),
    ] {
        let data = rewrite_numbering_part(original, |xml| xml.replace("<w:isLgl/>", flag));
        let text = converted_numbering_text(&data);
        let (expected, unexpected) = if enabled {
            (["1.1.", "1.2.", "2.1."], ["I.1.", "I.2.", "II.1."])
        } else {
            (["I.1.", "I.2.", "II.1."], ["1.1.", "1.2.", "2.1."])
        };
        for label in expected {
            assert!(text.contains(label), "{flag}: missing {label}: {text}");
        }
        for label in unexpected {
            assert!(!text.contains(label), "{flag}: unexpected {label}: {text}");
        }
        for body in [
            "First section",
            "Sub one",
            "Sub two",
            "Second section",
            "Sub one again",
        ] {
            assert!(text.contains(body), "{flag}: missing body {body}");
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_legal_numbering_forces_current_and_ancestor_formats_to_decimal() {
    let original = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    let data = rewrite_numbering_part(original, |xml| {
        xml.replace("upperRoman", "upperLetter")
            .replace("w:val=\"decimal\"", "w:val=\"lowerRoman\"")
    });
    let text = converted_numbering_text(&data);
    for label in ["A.", "B.", "1.1.", "1.2.", "2.1."] {
        assert!(text.contains(label), "missing {label}: {text}");
    }
    assert!(!text.contains("A.i."));
    assert!(!text.contains("B.i."));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_legal_numbering_override_can_disable_abstract_level_flag() {
    let original = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    let data = rewrite_numbering_part(original, |xml| {
        let start = xml.find("<w:num w:numId=\"100\">").unwrap();
        let end = start + xml[start..].find("</w:num>").unwrap();
        let level = r#"<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl w:val="false"/><w:lvlText w:val="%1.%2."/><w:pPr><w:ind w:left="1440" w:hanging="432"/></w:pPr></w:lvl></w:lvlOverride>"#;
        format!("{}{}{}", &xml[..end], level, &xml[end..])
    });
    let text = converted_numbering_text(&data);
    for label in ["I.1.", "I.2.", "II.1."] {
        assert!(text.contains(label), "missing {label}: {text}");
    }
    assert!(!text.contains("1.1."));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_skipped_list_levels_keep_absolute_marker_and_body_origins() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_skipped_levels.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let position = |label: &str| runs.iter().find(|run| run.text == label).unwrap();
    let first = position("Item C");
    let skipped = position("Item G");
    assert!(
        (first.left_pt - skipped.left_pt).abs() < 0.01,
        "same-level body origins: {first:?} / {skipped:?}"
    );
    let marker = runs
        .iter()
        .find(|run| {
            run.text.starts_with('(') && (run.baseline_pt - skipped.baseline_pt).abs() < 0.01
        })
        .unwrap();
    assert_eq!(
        marker.family, skipped.family,
        "marker must use the skipped level's text face"
    );
    assert!((marker.font_size_pt - skipped.font_size_pt).abs() < 0.01);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_skipped_list_subtrees_handle_returns_to_shallower_child_levels() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_restart_policy.docx");
    let (base, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let original_source = crate::render::typst_gen::generate_typst(&base).unwrap();
    let original_runs = crate::render::pdf::compiled_text_runs(&original_source.source, 0).unwrap();
    for levels in [
        [0, 2, 3, 1, 2, 0, 2],
        [0, 3, 2, 3, 1, 2, 0],
        [0, 2, 2, 1, 3, 0, 1],
    ] {
        let mut document = base.clone();
        let Page::Flow(page) = &mut document.pages[0] else {
            panic!("flow page")
        };
        let margin = page.margins.left;
        let list = page
            .content
            .iter_mut()
            .find_map(|element| match element {
                Block::List(list) => Some(list),
                _ => None,
            })
            .unwrap();
        let mut deepest = list.level_styles[&2].clone();
        deepest.numbering_pattern = Some(NumberingPattern::parse("[1]"));
        deepest.marker_style = Some(TextStyle {
            font_size: Some(11.0),
            ..TextStyle::default()
        });
        list.level_styles.insert(3, deepest);
        for (item, level) in list.items.iter_mut().zip(levels) {
            item.level = level;
            item.start_at = Some(1);
            item.content[0].style.indent_left = Some(18.0 * f64::from(level + 1));
            item.content[0].style.indent_first_line = Some(-18.0);
        }
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
        for (index, level) in levels.into_iter().enumerate() {
            let label = format!("Item {}", char::from(b'A' + index as u8));
            let body = runs.iter().find(|run| run.text == label).unwrap();
            let expected = margin + 18.0 * f64::from(level + 1);
            assert!(
                (body.left_pt - expected).abs() < 0.01,
                "levels {levels:?}, {label}: {} expected {expected}",
                body.left_pt
            );
            let original = original_runs.iter().find(|run| run.text == label).unwrap();
            assert!(
                (body.baseline_pt - original.baseline_pt).abs() < 0.5,
                "levels {levels:?}, {label}: baseline {} expected {}",
                body.baseline_pt,
                original.baseline_pt
            );
            if level == 3 {
                let marker = runs
                    .iter()
                    .find(|run| {
                        run.text == "[1]" && (run.baseline_pt - body.baseline_pt).abs() < 5.0
                    })
                    .unwrap();
                assert!((marker.font_size_pt - 11.0).abs() < 0.01);
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_skipped_ancestors_advance_later_visible_level_numbers() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    for start in [1, 7] {
        for (levels, labels) in [
            ([0, 2, 3, 1, 2, 0, 2], vec![("Item D", start + 1)]),
            (
                [0, 3, 2, 3, 1, 2, 0],
                vec![("Item C", 2), ("Item E", start + 1)],
            ),
            ([0, 2, 2, 1, 3, 0, 1], vec![("Item D", start + 1)]),
        ] {
            let data = rewrite_numbering_part(base, |xml| {
                let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
                format!(
                    "{prefix}<w:lvl w:ilvl=\"1\">{}",
                    child.replacen(
                        "<w:start w:val=\"1\"/>",
                        &format!("<w:start w:val=\"{start}\"/>"),
                        1
                    )
                )
            });
            let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
                let mut parts = xml.split("<w:ilvl w:val=\"");
                let mut output = parts.next().unwrap().to_string();
                for (part, level) in parts.zip(levels) {
                    let (_, rest) = part.split_once('"').unwrap();
                    output.push_str(&format!("<w:ilvl w:val=\"{level}\"{rest}"));
                }
                output
            });
            let text = converted_numbering_text(&data);
            let normalized: String = text.split_whitespace().collect();
            for (label, value) in labels {
                let label: String = label.split_whitespace().collect();
                assert!(
                    normalized.contains(&format!("({value}){label}")),
                    "levels {levels:?}, start {start}: {text}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_implicit_ancestor_numbers_respect_never_restart_across_roots() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    for never in [false, true] {
        let data = rewrite_docx_part(base, "word/document.xml", |xml| {
            let paragraph = xml
                .split("</w:p>")
                .find(|part| part.contains("<w:t>Item D</w:t>"))
                .unwrap();
            let (_, content) = paragraph.rsplit_once("<w:p>").unwrap();
            let extra = format!("<w:p>{content}</w:p>").replace("Item D", "Item H");
            xml.replacen("<w:sectPr", &format!("{extra}<w:sectPr"), 1)
        });
        let data = rewrite_numbering_part(&data, |xml| {
            if !never {
                return xml.to_string();
            }
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
            format!(
                "{prefix}<w:lvl w:ilvl=\"1\">{}",
                child.replacen("</w:lvl>", "<w:lvlRestart w:val=\"0\"/></w:lvl>", 1)
            )
        });
        let text = converted_numbering_text(&data);
        let normalized: String = text.split_whitespace().collect();
        assert!(
            normalized.contains(if never { "(1)ItemD" } else { "(2)ItemD" }),
            "never {never}: {text}"
        );
        assert!(normalized.contains("(2)ItemH"), "never {never}: {text}");
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_implicit_ancestors_follow_each_levels_restart_policy() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    let data = rewrite_docx_part(base, "word/document.xml", |xml| {
        let mut parts = xml.split("<w:ilvl w:val=\"");
        let mut output = parts.next().unwrap().to_string();
        for (part, level) in parts.zip([0, 3, 2, 3, 1, 2, 0]) {
            let (_, rest) = part.split_once('"').unwrap();
            output.push_str(&format!("<w:ilvl w:val=\"{level}\"{rest}"));
        }
        output
    });
    for (restart, first, later) in [(None, 2, 1), (Some(0), 1, 2), (Some(1), 2, 3)] {
        let data = rewrite_numbering_part(&data, |xml| {
            let Some(restart) = restart else {
                return xml.to_string();
            };
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"2\">").unwrap();
            format!(
                "{prefix}<w:lvl w:ilvl=\"2\">{}",
                child.replacen(
                    "</w:lvl>",
                    &format!("<w:lvlRestart w:val=\"{restart}\"/></w:lvl>"),
                    1
                )
            )
        });
        let text = converted_numbering_text(&data);
        let normalized: String = text.split_whitespace().collect();
        for (item, number) in [('C', first), ('E', 2), ('F', later)] {
            assert!(
                normalized.contains(&format!("({number})Item{item}")),
                "restart {restart:?}: {text}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_restart_threshold_suppresses_leaf_markers() {
    let data = include_bytes!(
        "../../../../tests/fixtures/docx/docxide/numbering_restart_suppression.docx"
    );
    let text = converted_numbering_text(data);
    let normalized: String = text.split_whitespace().collect();
    assert_eq!(
        normalized, "Numberingrestartpolicy1.ItemA(1)ItemBItemC(2)ItemDItemE2.ItemFItemG",
        "{text}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_restart_threshold_suppression_follows_the_current_level() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    for (level, restart, labels) in [
        (0, 1, ["1.", "(1)", "[1]", "(2)", "(1)", "2.", "(1)"]),
        (1, 1, ["1.", "(1)", "[1]", "", "(1)", "2.", "(1)"]),
        (1, 2, ["1.", "(1)", "[1]", "", "(1)", "2.", "(1)"]),
        (2, 2, ["1.", "", "[1]", "(2)", "", "2.", ""]),
        (2, 3, ["1.", "", "[1]", "(2)", "", "2.", ""]),
        (3, 2, ["1.", "(1)", "[1]", "(2)", "(1)", "2.", "(1)"]),
        (3, 3, ["1.", "(1)", "", "(2)", "(1)", "2.", "(1)"]),
        (3, 4, ["1.", "(1)", "", "(2)", "(1)", "2.", "(1)"]),
    ] {
        let data = rewrite_numbering_part(base, |xml| {
            let opener = format!("<w:lvl w:ilvl=\"{level}\">");
            let (prefix, child) = xml.split_once(&opener).unwrap();
            format!(
                "{prefix}{opener}{}",
                child.replacen(
                    "</w:lvl>",
                    &format!("<w:lvlRestart w:val=\"{restart}\"/></w:lvl>"),
                    1
                )
            )
        });
        let text = converted_numbering_text(&data);
        let normalized: String = text.split_whitespace().collect();
        let expected: String = std::iter::once("Numberingrestartpolicy".to_string())
            .chain(
                labels
                    .into_iter()
                    .zip('A'..='G')
                    .map(|(label, letter)| format!("{label}Item{letter}")),
            )
            .collect();
        assert_eq!(
            normalized, expected,
            "level {level}, restart {restart}: {text}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_suppressed_restart_preserves_paragraph_and_descendant_origins() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    let positions = |data: &[u8], use_fallback: bool, omit_line_multiple: bool| {
        let (mut document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
        if use_fallback {
            for page in &mut document.pages {
                let Page::Flow(page) = page else { continue };
                for block in &mut page.content {
                    let paragraphs: Vec<&mut Paragraph> = match block {
                        Block::Paragraph(paragraph) => vec![paragraph],
                        Block::List(list) => list
                            .items
                            .iter_mut()
                            .flat_map(|item| &mut item.content)
                            .collect(),
                        _ => continue,
                    };
                    for paragraph in paragraphs {
                        if omit_line_multiple {
                            paragraph.style.line_spacing = None;
                        }
                        paragraph.style.paragraph_mark_font_family =
                            Some("Office2Pdf Missing Numbering Face".into());
                        for run in &mut paragraph.runs {
                            run.style.font_family =
                                Some("Office2Pdf Missing Numbering Face".into());
                        }
                    }
                }
            }
        }
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap()
    };
    for (use_fallback, omit_line_multiple) in [(false, false), (true, false), (true, true)] {
        let control = positions(base, use_fallback, omit_line_multiple);
        for bullet in [false, true] {
            for indent in [0, 1080] {
                let data = rewrite_numbering_part(base, |xml| {
                    let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
                    let (level, rest) = child.split_once("</w:lvl>").unwrap();
                    let level = if bullet {
                        level
                            .replace("w:val=\"decimal\"", "w:val=\"bullet\"")
                            .replace("(%2)", "•")
                    } else {
                        level.to_string()
                    };
                    format!(
                        "{prefix}<w:lvl w:ilvl=\"1\">{level}<w:lvlRestart w:val=\"1\"/></w:lvl>{rest}"
                    )
                });
                let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
                    let paragraph = xml
                        .split("</w:p>")
                        .find(|part| part.contains("<w:t>Item D</w:t>"))
                        .unwrap();
                    let (_, content) = paragraph.rsplit_once("<w:p>").unwrap();
                    let original = format!("<w:p>{content}</w:p>");
                    let replacement = original.replace(
                        "</w:pPr>",
                        &format!("<w:ind w:left=\"{indent}\" w:hanging=\"0\"/></w:pPr>"),
                    );
                    xml.replace(&original, &replacement)
                });
                let runs = positions(&data, use_fallback, omit_line_multiple);
                for letter in 'A'..='G' {
                    let label = format!("Item {letter}");
                    let before = control.iter().find(|run| run.text == label).unwrap();
                    let after = runs.iter().find(|run| run.text == label).unwrap();
                    let expected_x = if letter == 'D' {
                        90.0 + f64::from(indent) / 20.0
                    } else {
                        before.left_pt
                    };
                    assert!(
                        (after.left_pt - expected_x).abs() < 0.01,
                        "{label}, bullet {bullet}, indent {indent}: {after:?}"
                    );
                    assert!(
                        (after.baseline_pt - before.baseline_pt).abs() < 0.15,
                        "physical baseline {label}, bullet {bullet}, indent {indent}: {before:?} / {after:?}"
                    );
                }
                let body = runs.iter().find(|run| run.text == "Item D").unwrap();
                assert!(
                    !runs.iter().any(|run| run.text != "Item D"
                        && (run.baseline_pt - body.baseline_pt).abs() < 0.01),
                    "suppressed marker: {runs:?}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_out_of_range_restart_continues_child_labels() {
    let data = include_bytes!(
        "../../../../tests/fixtures/docx/docxide/numbering_out_of_range_restarts.docx"
    );
    let text = converted_numbering_text(data);
    assert_eq!(
        text.split_whitespace().collect::<String>(),
        "Numberingrestartpolicy1.ItemA(1)ItemB[1]ItemC(2)ItemD(2)ItemE2.ItemF(3)ItemG",
        "{text}"
    );
    let boundary = rewrite_numbering_part(data, |xml| {
        xml.replace("<w:lvlRestart w:val=\"8\"/>", "<w:lvlRestart w:val=\"7\"/>")
    });
    let text = converted_numbering_text(&boundary);
    assert_eq!(
        text.split_whitespace().collect::<String>(),
        "Numberingrestartpolicy1.ItemAItemB[1]ItemC(2)ItemDItemE2.ItemFItemG",
        "{text}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_out_of_range_restart_counts_physical_items_across_levels() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_implicit_ancestors.docx");
    for level in [1, 2, 3] {
        for restart in [8, 9, 65535] {
            for start in [1, 7] {
                let data = rewrite_numbering_part(base, |xml| {
                    let opener = format!("<w:lvl w:ilvl=\"{level}\">");
                    let (prefix, child) = xml.split_once(&opener).unwrap();
                    let (current, rest) = child.split_once("</w:lvl>").unwrap();
                    format!(
                        "{prefix}{opener}{}<w:lvlRestart w:val=\"{restart}\"/></w:lvl>{rest}",
                        current.replace(
                            "<w:start w:val=\"1\"/>",
                            &format!("<w:start w:val=\"{start}\"/>")
                        )
                    )
                });
                let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
                    if level == 1 {
                        let paragraph = xml
                            .split("</w:p>")
                            .find(|part| part.contains("<w:t>Item D</w:t>"))
                            .unwrap();
                        let (_, content) = paragraph.rsplit_once("<w:p>").unwrap();
                        let next = format!("<w:p>{content}</w:p>").replace("Item D", "Item H");
                        xml.replace("<w:sectPr", &format!("{next}<w:sectPr"))
                    } else if level == 3 {
                        let mut levels = [0, 3, 2, 3, 1, 2, 0].into_iter();
                        xml.split_inclusive("<w:ilvl w:val=\"")
                            .enumerate()
                            .map(|(index, part)| {
                                if index == 0 {
                                    return part.to_string();
                                }
                                let (_, rest) = part.split_once('"').unwrap();
                                format!("{}\"{rest}", levels.next().unwrap())
                            })
                            .collect()
                    } else {
                        xml.to_string()
                    }
                });
                let expected = match level {
                    1 => format!(
                        "Numberingrestartpolicy1.ItemA(1)ItemB[1]ItemC({start})ItemD(1)ItemE2.ItemF(1)ItemG({})ItemH",
                        start + 1
                    ),
                    2 => format!(
                        "Numberingrestartpolicy1.ItemA({start})ItemB[1]ItemC(2)ItemD({})ItemE2.ItemF({})ItemG",
                        start + 1,
                        start + 2
                    ),
                    3 => format!(
                        "Numberingrestartpolicy1.ItemA[{start}]ItemB(2)ItemC[{}]ItemD(2)ItemE(1)ItemF2.ItemG",
                        start + 1
                    ),
                    _ => unreachable!(),
                };
                let text = converted_numbering_text(&data);
                assert_eq!(
                    text.split_whitespace().collect::<String>(),
                    expected,
                    "level {level}, restart {restart}, start {start}: {text}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_out_of_range_abstract_restart_survives_complete_overrides() {
    let base = include_bytes!(
        "../../../../tests/fixtures/docx/docxide/numbering_out_of_range_restarts.docx"
    );
    for override_restart in [1, 8, 65535] {
        let data = rewrite_numbering_part(base, |xml| {
            let level = xml
                .split_once("<w:lvl w:ilvl=\"2\">")
                .unwrap()
                .1
                .split_once("</w:lvl>")
                .unwrap()
                .0;
            let overridden = level.replace(
                "<w:lvlRestart w:val=\"8\"/>",
                &format!("<w:lvlRestart w:val=\"{override_restart}\"/>"),
            );
            xml.replace("<w:num w:numId=\"100\"><w:abstractNumId w:val=\"100\"/></w:num>",
                &format!("<w:num w:numId=\"100\"><w:abstractNumId w:val=\"100\"/><w:lvlOverride w:ilvl=\"2\"><w:lvl w:ilvl=\"2\">{overridden}</w:lvl></w:lvlOverride></w:num>"))
        });
        let text = converted_numbering_text(&data);
        assert_eq!(
            text.split_whitespace().collect::<String>(),
            "Numberingrestartpolicy1.ItemA(1)ItemB[1]ItemC(2)ItemD(2)ItemE2.ItemF(3)ItemG",
            "override restart {override_restart}: {text}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_wide_numbering_body_tabs(data: &[u8], interval: f64, explicit_stop: Option<f64>) {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
    let mut checked = 0;
    for marker in text.iter().filter(|run| run.text.starts_with('(')) {
        let body = text
            .iter()
            .find(|run| {
                run.text.starts_with("Continuous item")
                    && (run.baseline_pt - marker.baseline_pt).abs() < 0.01
            })
            .unwrap();
        let shape = glyphs.iter().find(|run| run.text == marker.text).unwrap();
        let end = marker.left_pt + shape.width_pt();
        // The native probes preserve the nominal body origin for a fitting
        // marker, otherwise use a custom stop or the document's tab grid.
        let expected = if end < 126.0 - 0.01 {
            126.0
        } else if let Some(stop) = explicit_stop.filter(|stop| *stop > end + 0.01) {
            stop
        } else {
            90.0 + (((end - 90.0) / interval).floor() + 1.0) * interval
        };
        assert!(
            (body.left_pt - expected).abs() < 0.01,
            "{} ends at {end}; body at {}, expected {expected}",
            marker.text,
            body.left_pt
        );
        assert!((marker.left_pt - 108.0).abs() < 0.01);
        checked += 1;
    }
    assert_eq!(checked, 4);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_wide_numbering_markers_follow_default_and_custom_tabs() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    for start in [7, 98, 998] {
        let data = rewrite_numbering_part(base, |xml| {
            xml.replace(
                "<w:start w:val=\"7\"/>",
                &format!("<w:start w:val=\"{start}\"/>"),
            )
        });
        assert_wide_numbering_body_tabs(&data, 36.0, None);
        let shorter = rewrite_docx_part(&data, "word/settings.xml", |xml| {
            xml.replace(
                "<w:defaultTabStop w:val=\"720\"/>",
                "<w:defaultTabStop w:val=\"480\"/>",
            )
        });
        assert_wide_numbering_body_tabs(&shorter, 24.0, None);
        let custom = rewrite_numbering_part(&data, |xml| {
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
            format!(
                "{prefix}<w:lvl w:ilvl=\"1\">{}",
                child.replacen(
                    "<w:pPr>",
                    "<w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"1260\"/></w:tabs>",
                    1
                )
            )
        });
        assert_wide_numbering_body_tabs(&custom, 36.0, Some(153.0));
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_wide_numbering_tab_keeps_continuation_lines_at_declared_indent() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    let data = rewrite_docx_part(base, "word/document.xml", |xml| {
        let (prefix, suffix) = xml.rsplit_once("Continuous item").unwrap();
        format!(
            "{prefix}Continuous item with a long explanation about the final wide numbering marker and the paragraph continuation origin. This sentence continues across several physical lines to test wrapping at the declared left margin.{suffix}"
        )
    });
    assert_wide_numbering_body_tabs(&data, 36.0, None);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let first = runs
        .iter()
        .find(|run| run.text.starts_with("Continuous item with"))
        .unwrap();
    let continuations: Vec<_> = runs
        .iter()
        .filter(|run| run.baseline_pt > first.baseline_pt + 1.0)
        .collect();
    assert!(continuations.len() >= 2, "{continuations:?}");
    for line in continuations {
        assert!((line.left_pt - 126.0).abs() < 0.01, "{line:?}");
    }
}

/// The level-1 geometry and tab grid a list-body tab probe runs on, in points
/// from the 90pt left margin, and the level's `w:suff`.
#[cfg(not(target_arch = "wasm32"))]
struct ListBodyTabGrid<'a> {
    body_origin: f64,
    hanging: f64,
    interval: f64,
    stops: &'a [f64],
    suffix: &'a str,
}

#[cfg(not(target_arch = "wasm32"))]
impl ListBodyTabGrid<'_> {
    /// Word's next stop strictly past `pen`: the first custom stop ahead of
    /// it, otherwise the next multiple of the default interval.
    fn next_stop(&self, pen: f64) -> f64 {
        self.stops
            .iter()
            .copied()
            .find(|stop| *stop > pen + 0.01)
            .unwrap_or_else(|| {
                ((pen + 0.01) / self.interval).floor() * self.interval + self.interval
            })
    }
}

/// Every `(n)` item's body begins with `leading_tabs` tabs, then `first`;
/// `second` follows one more tab when given. Word counts tab stops from the
/// margin, so the pen starts at the body origin, or where a wide marker's
/// suffix tab already left it, and each tab moves it to the next stop.
#[cfg(not(target_arch = "wasm32"))]
fn assert_list_body_tabs(
    data: &[u8],
    grid: &ListBodyTabGrid<'_>,
    leading_tabs: usize,
    first: &str,
    second: Option<&str>,
) {
    const MARGIN: f64 = 90.0;
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
    let width = |label: &str| {
        glyphs
            .iter()
            .find(|run| run.text == label)
            .unwrap_or_else(|| panic!("no glyph run {label:?} in {glyphs:?}"))
            .width_pt()
    };
    let mut checked = 0;
    for marker in text.iter().filter(|run| run.text.starts_with('(')) {
        let on_marker_line = |label: &str| {
            text.iter()
                .find(|run| {
                    run.text.starts_with(label)
                        && (run.baseline_pt - marker.baseline_pt).abs() < 0.01
                })
                .unwrap_or_else(|| panic!("no {label:?} beside {}: {text:?}", marker.text))
        };
        let marker_end = marker.left_pt + width(&marker.text) - MARGIN;
        let mut pen = match grid.suffix {
            "space" => marker_end + compiled_numbering_separator_width(marker.font_size_pt),
            "nothing" => marker_end,
            _ if marker_end < grid.body_origin - 0.01 => grid.body_origin,
            _ => grid.next_stop(marker_end),
        };
        for _ in 0..leading_tabs {
            pen = grid.next_stop(pen);
        }
        let body = on_marker_line(first);
        assert!(
            (body.left_pt - (MARGIN + pen)).abs() < 0.01,
            "{} ends at {marker_end}; {first:?} at {}, expected {}: {text:?}",
            marker.text,
            body.left_pt,
            MARGIN + pen
        );
        if let Some(second) = second {
            let expected = MARGIN + grid.next_stop(body.left_pt + width(first) - MARGIN);
            let after = on_marker_line(second);
            assert!(
                (after.left_pt - expected).abs() < 0.01,
                "{}: {second:?} at {}, expected {expected}",
                marker.text,
                after.left_pt
            );
        }
        let marker_origin = MARGIN + grid.body_origin - grid.hanging;
        assert!((marker.left_pt - marker_origin).abs() < 0.01, "{marker:?}");
        checked += 1;
    }
    assert_eq!(checked, 4);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_list_body_tabs_advance_from_the_body_origin_on_margin_stops() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    let level_one = |data: &[u8], edit: &dyn Fn(&str) -> String| {
        rewrite_numbering_part(data, |xml| {
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
            format!("{prefix}<w:lvl w:ilvl=\"1\">{}", edit(child))
        })
    };
    let with_body = |data: &[u8], body: &str| {
        rewrite_docx_part(data, "word/document.xml", |xml| {
            xml.replace("<w:t>Continuous item</w:t>", body)
        })
    };
    let on_grid = ListBodyTabGrid {
        body_origin: 36.0,
        hanging: 18.0,
        interval: 36.0,
        stops: &[],
        suffix: "tab",
    };
    for start in [7, 998] {
        let data = rewrite_numbering_part(base, |xml| {
            xml.replace(
                "<w:start w:val=\"7\"/>",
                &format!("<w:start w:val=\"{start}\"/>"),
            )
        });
        let one = with_body(&data, "<w:tab/><w:t>Continuous item</w:t>");
        assert_list_body_tabs(&one, &on_grid, 1, "Continuous item", None);
        let two = with_body(&data, "<w:tab/><w:tab/><w:t>Continuous item</w:t>");
        assert_list_body_tabs(&two, &on_grid, 2, "Continuous item", None);
        let shorter = rewrite_docx_part(&one, "word/settings.xml", |xml| {
            xml.replace(
                "<w:defaultTabStop w:val=\"720\"/>",
                "<w:defaultTabStop w:val=\"480\"/>",
            )
        });
        let shorter_grid = ListBodyTabGrid {
            interval: 24.0,
            ..on_grid
        };
        assert_list_body_tabs(&shorter, &shorter_grid, 1, "Continuous item", None);
        // A custom stop ahead of the pen wins over the default grid.
        let custom = level_one(&one, &|child| {
            child.replacen(
                "<w:pPr>",
                "<w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1260\"/></w:tabs>",
                1,
            )
        });
        let custom_grid = ListBodyTabGrid {
            stops: &[63.0],
            ..on_grid
        };
        assert_list_body_tabs(&custom, &custom_grid, 1, "Continuous item", None);
        // A body origin off the default grid, with a custom stop the pen has
        // already reached there: the tab skips it for the next default stop.
        let off_grid = level_one(&one, &|child| {
            child.replacen(
                "<w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/>",
                "<w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1152\"/></w:tabs><w:ind w:left=\"1152\" w:hanging=\"432\"/>",
                1,
            )
        });
        let off_grid_geometry = ListBodyTabGrid {
            body_origin: 57.6,
            hanging: 21.6,
            interval: 36.0,
            stops: &[57.6],
            suffix: "tab",
        };
        assert_list_body_tabs(&off_grid, &off_grid_geometry, 1, "Continuous item", None);
        // A space or no separator leaves the pen just past the marker, on
        // either side of the body origin, and the tab measures from there.
        for suffix in ["space", "nothing"] {
            let suffixed = level_one(&one, &|child| {
                child.replacen(
                    "<w:lvlText",
                    &format!("<w:suff w:val=\"{suffix}\"/><w:lvlText"),
                    1,
                )
            });
            let suffixed_grid = ListBodyTabGrid { suffix, ..on_grid };
            assert_list_body_tabs(&suffixed, &suffixed_grid, 1, "Continuous item", None);
        }
        // A tab inside the body measures the text before it from the same
        // origin.
        let inner = with_body(&data, "<w:t>Continuous</w:t><w:tab/><w:t>item</w:t>");
        assert_list_body_tabs(&inner, &on_grid, 0, "Continuous", Some("item"));
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_list_body_tab_keeps_continuation_lines_at_declared_indent() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    // Every item starts with one tab; the last one's body also wraps.
    let data = rewrite_docx_part(base, "word/document.xml", |xml| {
        let (prefix, suffix) = xml.rsplit_once("<w:t>Continuous item</w:t>").unwrap();
        format!(
            "{prefix}<w:t>Continuous item with a long explanation after a leading tab about the final wide numbering marker and the paragraph continuation origin. This sentence continues across several physical lines to test wrapping at the declared left margin.</w:t>{suffix}"
        )
        .replace("<w:t>Continuous item", "<w:tab/><w:t>Continuous item")
    });
    let grid = ListBodyTabGrid {
        body_origin: 36.0,
        hanging: 18.0,
        interval: 36.0,
        stops: &[],
        suffix: "tab",
    };
    assert_list_body_tabs(&data, &grid, 1, "Continuous item", None);
    // Check the requested face and the narrower CI fallback, which fits more
    // of the body on each line.
    for family in [None, Some("Libertinus Serif")] {
        let (mut document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        if let Some(family) = family {
            for page in &mut document.pages {
                if let crate::ir::Page::Flow(flow) = page {
                    for block in &mut flow.content {
                        if let Block::List(list) = block {
                            for item in &mut list.items {
                                for paragraph in &mut item.content {
                                    for run in &mut paragraph.runs {
                                        run.style.font_family = Some(family.to_string());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
        let first = runs
            .iter()
            .find(|run| run.text.starts_with("Continuous item with"))
            .unwrap();
        let continuations: Vec<_> = runs
            .iter()
            .filter(|run| run.baseline_pt > first.baseline_pt + 1.0)
            .collect();
        assert!(continuations.len() >= 2, "{family:?}: {continuations:?}");
        for line in continuations {
            assert!((line.left_pt - 126.0).abs() < 0.01, "{family:?}: {line:?}");
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn compiled_numbering_separator_width(size: f64) -> f64 {
    // Shape an isolated space between non-ligating glyphs: character offsets
    // in a body word need not match its glyph indices.
    let font = crate::render::font_subst::font_with_fallbacks_for_text("Arial", " ");
    let source = format!("#set text(font: {font}, size: {size}pt)\nx x");
    let runs = crate::render::pdf::compiled_glyph_runs(&source, 0, true).unwrap();
    let probe = runs.iter().find(|run| run.text == "x x").unwrap();
    assert_eq!(probe.advances_pt.len(), 3);
    probe.advances_pt[1]
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_non_tab_numbering_suffix_follows_actual_marker_advance() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    for marker_size in [11.0, 16.0] {
        for family in ["Cambria", "Libertinus Serif"] {
            for start in [7, 998] {
                for suffix in ["space", "nothing"] {
                    let data = rewrite_numbering_part(base, |xml| {
                        let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
                        format!(
                            "{prefix}<w:lvl w:ilvl=\"1\">{}",
                            child.replacen(
                                "<w:lvlText",
                                &format!("<w:suff w:val=\"{suffix}\"/><w:lvlText"),
                                1
                            )
                        )
                        .replace(
                            "<w:start w:val=\"7\"/>",
                            &format!("<w:start w:val=\"{start}\"/>"),
                        )
                    });
                    let (mut document, _) =
                        DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
                    // Exercise both the requested Word face and the CI fallback;
                    // a fitting fallback marker has no tab overflow to suppress.
                    for page in &mut document.pages {
                        if let crate::ir::Page::Flow(flow) = page {
                            for block in &mut flow.content {
                                if let Block::List(list) = block {
                                    for level in list.level_styles.values_mut() {
                                        let marker_style =
                                            level.marker_style.get_or_insert_with(Default::default);
                                        marker_style.font_family = Some(family.to_string());
                                        marker_style.font_size = Some(marker_size);
                                    }
                                    for item in &mut list.items {
                                        for paragraph in &mut item.content {
                                            for run in &mut paragraph.runs {
                                                run.style.font_family = Some(family.to_string());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
                    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
                    let label = format!("({})", start + 3);
                    let marker = runs.iter().find(|run| run.text == label).unwrap();
                    let body = runs
                        .iter()
                        .find(|run| {
                            run.text == "Continuous item"
                                && (run.baseline_pt - marker.baseline_pt).abs() < 0.01
                        })
                        .unwrap();
                    let glyphs =
                        crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
                    let shape = glyphs.iter().find(|run| run.text == marker.text).unwrap();
                    let end = marker.left_pt + shape.width_pt();
                    let space_advance = compiled_numbering_separator_width(marker.font_size_pt);
                    let expected = end
                        + if suffix == "space" {
                            space_advance
                        } else {
                            0.0
                        };
                    assert!(
                        (body.left_pt - expected).abs() < 0.01,
                        "suffix {suffix}, {family}, start {start} follows its marker at {end} and separator, expected {expected}: {body:?}"
                    );
                    if start == 998 {
                        assert!(
                            end > 126.0,
                            "the wide-marker case must exercise overflow: {marker:?}"
                        );
                        let next_tab = 90.0 + (((end - 90.0) / 36.0).floor() + 1.0) * 36.0;
                        assert!(
                            body.left_pt < next_tab - 0.01,
                            "suffix {suffix} must not gain a tab jump: body {body:?}, next tab {next_tab}"
                        );
                    }
                }
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_non_tab_numbering_suffix_keeps_wrapped_continuation_indent() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    for suffix in ["space", "nothing"] {
        let data = rewrite_numbering_part(base, |xml| {
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
            format!(
                "{prefix}<w:lvl w:ilvl=\"1\">{}",
                child.replacen(
                    "<w:lvlText",
                    &format!("<w:suff w:val=\"{suffix}\"/><w:lvlText"),
                    1
                )
            )
        });
        let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
            let (prefix, tail) = xml.rsplit_once("Continuous item").unwrap();
            format!(
                "{prefix}Continuous item with a long explanation about the final wide numbering marker and the paragraph continuation origin. This sentence continues across several physical lines to test wrapping at the declared left margin.{tail}"
            )
        });
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
        let marker = runs.iter().find(|run| run.text == "(10)").unwrap();
        let first = runs
            .iter()
            .find(|run| run.text.starts_with("Continuous item with"))
            .unwrap();
        let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
        let marker_shape = glyphs.iter().find(|run| run.text == "(10)").unwrap();
        let expected = marker.left_pt
            + marker_shape.width_pt()
            + if suffix == "space" {
                compiled_numbering_separator_width(marker.font_size_pt)
            } else {
                0.0
            };
        assert!(
            (first.left_pt - expected).abs() < 0.01,
            "{suffix}: {first:?}, expected {expected}"
        );
        let continuations: Vec<_> = runs
            .iter()
            .filter(|run| run.baseline_pt > first.baseline_pt + 1.0)
            .collect();
        assert!(continuations.len() >= 2, "{continuations:?}");
        for run in continuations {
            assert!((run.left_pt - 126.0).abs() < 0.01, "{suffix}: {run:?}");
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_bullet_suffix_follows_actual_marker_advance() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    for suffix in ["space", "nothing"] {
        let data = rewrite_numbering_part(base, |xml| {
            let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
            let (level, tail) = child.split_once("</w:lvl>").unwrap();
            let level = level
                .replace(
                    "<w:numFmt w:val=\"decimal\"/>",
                    "<w:numFmt w:val=\"bullet\"/>",
                )
                .replace(
                    "<w:lvlText w:val=\"(%2)\"/>",
                    &format!("<w:suff w:val=\"{suffix}\"/><w:lvlText w:val=\"•\"/>"),
                );
            format!("{prefix}<w:lvl w:ilvl=\"1\">{level}</w:lvl>{tail}")
        });
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
        let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
        let mut checked = 0;
        for marker in runs.iter().filter(|run| run.text == "•") {
            let body = runs
                .iter()
                .find(|run| {
                    run.text == "Continuous item"
                        && (run.baseline_pt - marker.baseline_pt).abs() < 0.01
                })
                .unwrap();
            let marker_shape = glyphs.iter().find(|run| run.text == marker.text).unwrap();
            let expected = marker.left_pt
                + marker_shape.width_pt()
                + if suffix == "space" {
                    compiled_numbering_separator_width(marker.font_size_pt)
                } else {
                    0.0
                };
            assert!(
                (body.left_pt - expected).abs() < 0.01,
                "bullet {suffix}: {body:?}, expected {expected}"
            );
            checked += 1;
        }
        assert_eq!(checked, 4);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_inline_numbering_honors_space_and_nothing_suffixes() {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    for size in [11, 16] {
        for start in [7, 998] {
            for suffix in ["space", "nothing"] {
                let data = rewrite_numbering_part(base, |xml| {
                    let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
                    let (level, tail) = child.split_once("</w:lvl>").unwrap();
                    let level = level
                        .replace(
                            "<w:ind w:left=\"720\" w:hanging=\"360\"/>",
                            "<w:ind w:left=\"0\" w:hanging=\"0\"/>",
                        )
                        .replace(
                            "<w:lvlText",
                            &format!("<w:suff w:val=\"{suffix}\"/><w:lvlText"),
                        );
                    format!("{prefix}<w:lvl w:ilvl=\"1\">{level}</w:lvl>{tail}").replace(
                        "<w:start w:val=\"7\"/>",
                        &format!("<w:start w:val=\"{start}\"/>"),
                    )
                });
                let data = rewrite_docx_part(&data, "word/styles.xml", |xml| {
                    let (prefix, properties) = xml.split_once("<w:rPrDefault>").unwrap();
                    let (properties, tail) = properties.split_once("</w:rPrDefault>").unwrap();
                    format!(
                        "{prefix}<w:rPrDefault>{}</w:rPrDefault>{tail}",
                        properties.replace(
                            "<w:sz w:val=\"22\"/>",
                            &format!("<w:sz w:val=\"{}\"/>", size * 2)
                        )
                    )
                });
                let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
                let source = crate::render::typst_gen::generate_typst(&document).unwrap();
                let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
                let glyphs =
                    crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
                let mut checked = 0;
                for marker in runs.iter().filter(|run| run.text.starts_with('(')) {
                    let body = runs
                        .iter()
                        .find(|run| {
                            run.text == "Continuous item"
                                && (run.baseline_pt - marker.baseline_pt).abs() < 0.01
                        })
                        .unwrap();
                    let shape = glyphs.iter().find(|run| run.text == marker.text).unwrap();
                    let expected = marker.left_pt
                        + shape.width_pt()
                        + if suffix == "space" {
                            compiled_numbering_separator_width(marker.font_size_pt)
                        } else {
                            0.0
                        };
                    assert!(
                        (body.left_pt - expected).abs() < 0.01,
                        "inline {suffix}, size {size}, start {start}: {marker:?}, {body:?}, expected {expected}"
                    );
                    assert!((marker.left_pt - 90.0).abs() < 0.01);
                    checked += 1;
                }
                assert_eq!(checked, 4);
            }
        }
    }
}
#[cfg(not(target_arch = "wasm32"))]
fn numbering_marker_properties_fixture(
    properties: &str,
    full_override: bool,
    inline: bool,
) -> Vec<u8> {
    let base =
        include_bytes!("../../../../tests/fixtures/docx/docxide/numbering_wide_markers.docx");
    rewrite_numbering_part(base, |xml| {
        let (prefix, child) = xml.split_once("<w:lvl w:ilvl=\"1\">").unwrap();
        let end = child.find("</w:lvl>").unwrap();
        let mut child_level = format!(
            "<w:lvl w:ilvl=\"1\">{}<w:rPr>{properties}</w:rPr></w:lvl>",
            &child[..end]
        );
        if inline {
            child_level = child_level.replace(
                "w:left=\"720\" w:hanging=\"360\"",
                "w:left=\"0\" w:hanging=\"0\"",
            );
        }
        if full_override {
            let start_num = xml.find("<w:num w:numId=\"100\">").unwrap();
            let end_num = start_num + xml[start_num..].find("</w:num>").unwrap();
            format!(
                "{}<w:lvlOverride w:ilvl=\"1\">{child_level}</w:lvlOverride>{}",
                &xml[..end_num],
                &xml[end_num..]
            )
        } else {
            format!("{prefix}{child_level}{}", &child[end + "</w:lvl>".len()..])
        }
    })
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_numbering_marker_run_properties_survive_abstract_and_override_levels() {
    for full_override in [false, true] {
        for (family, size, bold, italic, color) in [
            (
                "Times New Roman",
                16.0,
                true,
                false,
                Color::new(192, 32, 64),
            ),
            ("Arial", 22.0, false, true, Color::new(16, 80, 160)),
        ] {
            let properties = format!(
                r#"<w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="{}"/><w:b w:val="{bold}"/><w:i w:val="{italic}"/><w:color w:val="{:02X}{:02X}{:02X}"/>"#,
                size * 2.0,
                color.r,
                color.g,
                color.b
            );
            let data = numbering_marker_properties_fixture(&properties, full_override, false);
            let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
            let list = all_blocks(&document)
                .iter()
                .find_map(|block| match block {
                    Block::List(list) => Some(list),
                    _ => None,
                })
                .unwrap();
            let marker = list.level_styles[&1]
                .marker_style
                .as_ref()
                .expect("numbering-level run properties retained");
            assert_eq!(marker.font_family.as_deref(), Some(family));
            assert_eq!(marker.font_size, Some(size));
            assert_eq!(marker.bold, Some(bold));
            assert_eq!(marker.italic, Some(italic));
            assert_eq!(marker.color, Some(color));
            assert_eq!(list.level_styles[&0].marker_style, None);
            for item in list.items.iter().filter(|item| item.level == 1) {
                let body = &item.content[0].runs[0].style;
                assert_eq!(body.font_size, Some(11.0));
                assert_ne!(body.color, Some(color));
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_numbering_marker_sizes_apply_without_changing_body_size() {
    for inline in [false, true] {
        for size in [16.0, 22.0] {
            let properties = format!(
                r#"<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="{}"/>"#,
                size * 2.0
            );
            let data = numbering_marker_properties_fixture(&properties, false, inline);
            let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
            let source = crate::render::typst_gen::generate_typst(&document).unwrap();
            let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
            let mut count = 0;
            for label in ["(7)", "(8)", "(9)", "(10)"] {
                let marker = runs.iter().find(|run| run.text.trim() == label).unwrap();
                assert!(
                    (marker.font_size_pt - size).abs() < 0.01,
                    "inline {inline}, marker {marker:?} expected {size}pt"
                );
                count += 1;
            }
            assert_eq!(count, 4);
            for body in runs.iter().filter(|run| run.text == "Continuous item") {
                assert!((body.font_size_pt - 11.0).abs() < 0.01, "{body:?}");
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn inline_numbering_formats_fixture(
    parent_format: &str,
    child_format: &str,
    parent_start: u32,
    child_start: u32,
    parent_is_legal: bool,
    child_is_legal: bool,
) -> Vec<u8> {
    let base = include_bytes!("../../../../tests/fixtures/docx/docxide/legal_numbering.docx");
    rewrite_numbering_part(base, |xml| {
        let start = xml.find("<w:abstractNum w:abstractNumId=\"100\">").unwrap();
        let end = start + xml[start..].find("</w:abstractNum>").unwrap();
        let mut definition = xml[start..end].to_string();
        for (level, format, counter_start, is_legal) in [
            (0, parent_format, parent_start, parent_is_legal),
            (1, child_format, child_start, child_is_legal),
        ] {
            let level_start = definition
                .find(&format!("<w:lvl w:ilvl=\"{level}\">"))
                .unwrap();
            let level_end = level_start + definition[level_start..].find("</w:lvl>").unwrap();
            let mut content = definition[level_start..level_end].to_string();
            content = content.replace(
                "<w:start w:val=\"1\"/>",
                &format!("<w:start w:val=\"{counter_start}\"/>"),
            );
            content = content.replace(
                if level == 0 {
                    "<w:numFmt w:val=\"upperRoman\"/>"
                } else {
                    "<w:numFmt w:val=\"decimal\"/>"
                },
                &format!("<w:numFmt w:val=\"{format}\"/>"),
            );
            content = content.replace("<w:isLgl/>", "");
            if is_legal {
                content = content.replace("<w:lvlText", "<w:isLgl/><w:lvlText");
            }
            content = content
                .replace(
                    "w:left=\"720\" w:hanging=\"432\"",
                    "w:left=\"0\" w:hanging=\"0\"",
                )
                .replace(
                    "w:left=\"1440\" w:hanging=\"432\"",
                    "w:left=\"0\" w:hanging=\"0\"",
                );
            definition.replace_range(level_start..level_end, &content);
        }
        format!("{}{definition}{}", &xml[..start], &xml[end..])
    })
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_inline_numbering_honors_ancestor_formats_counters_and_legal_flags() {
    for (parent, child, parent_start, child_start, parent_legal, child_legal, labels) in [
        (
            "upperRoman",
            "decimal",
            1,
            1,
            false,
            true,
            ["I.", "1.1.", "1.2.", "II.", "2.1."],
        ),
        (
            "upperRoman",
            "decimal",
            1,
            1,
            false,
            false,
            ["I.", "I.1.", "I.2.", "II.", "II.1."],
        ),
        (
            "upperLetter",
            "lowerRoman",
            27,
            7,
            false,
            false,
            ["AA.", "AA.vii.", "AA.viii.", "BB.", "BB.vii."],
        ),
        (
            "upperLetter",
            "lowerRoman",
            27,
            7,
            false,
            true,
            ["AA.", "27.7.", "27.8.", "BB.", "28.7."],
        ),
        (
            "lowerLetter",
            "upperRoman",
            26,
            4,
            false,
            false,
            ["z.", "z.IV.", "z.V.", "aa.", "aa.IV."],
        ),
        (
            "decimalZero",
            "lowerLetter",
            9,
            26,
            false,
            false,
            ["09.", "09.z.", "09.aa.", "10.", "10.z."],
        ),
        (
            "upperRoman",
            "decimal",
            1,
            1,
            true,
            false,
            ["1.", "I.1.", "I.2.", "2.", "II.1."],
        ),
    ] {
        let data = inline_numbering_formats_fixture(
            parent,
            child,
            parent_start,
            child_start,
            parent_legal,
            child_legal,
        );
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let prefixes: Vec<&str> = all_blocks(&document)
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) if paragraph.style.has_inline_list_number => {
                    Some(paragraph.runs[0].text.trim_end_matches('\t'))
                }
                _ => None,
            })
            .collect();
        assert_eq!(
            prefixes, labels,
            "parent {parent}/{parent_start}/{parent_legal}, child {child}/{child_start}/{child_legal}"
        );
        let text = converted_numbering_text(&data);
        for label in labels {
            assert!(text.contains(label), "missing {label}: {text}");
        }
        for body in [
            "First section",
            "Sub one",
            "Sub two",
            "Second section",
            "Sub one again",
        ] {
            assert!(text.contains(body), "missing {body}: {text}");
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_inline_numbering_preserves_literal_letters_and_percent_signs() {
    let base = inline_numbering_formats_fixture("upperRoman", "decimal", 1, 1, false, false);
    let data = rewrite_numbering_part(&base, |xml| {
        xml.replace("w:val=\"%1.\"", "w:val=\"Part A %1: 100%\"")
            .replace("w:val=\"%1.%2.\"", "w:val=\"Clause %1-%2\"")
    });
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let prefixes: Vec<&str> = all_blocks(&document)
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) if paragraph.style.has_inline_list_number => {
                Some(paragraph.runs[0].text.trim_end_matches('\t'))
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        prefixes,
        [
            "Part A I: 100%",
            "Clause I-1",
            "Clause I-2",
            "Part A II: 100%",
            "Clause II-1"
        ]
    );
    let text = converted_numbering_text(&data);
    for prefix in prefixes {
        assert!(
            text.contains(prefix),
            "missing literal label {prefix}: {text}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_large_inline_markers_keep_native_paragraph_advance() {
    let mut marker_height_deltas = std::collections::BTreeMap::new();
    for (marker_size, line, expected_pitch) in [
        (16, 240, 27.6),
        (16, 276, 29.52),
        (16, 480, 40.56),
        (22, 276, 35.28),
    ] {
        for suffix in ["space", "nothing", "tab"] {
            let data = numbering_marker_properties_fixture(
                &format!("<w:sz w:val=\"{}\"/>", marker_size * 2),
                false,
                true,
            );
            let data = rewrite_numbering_part(&data, |xml| {
                xml.replace(
                    "<w:lvlRestart w:val=\"0\"/>",
                    &format!("<w:lvlRestart w:val=\"0\"/><w:suff w:val=\"{suffix}\"/>"),
                )
            });
            let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
                xml.replace(
                    "<w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"100\"/></w:numPr></w:pPr>",
                    &format!("<w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"100\"/></w:numPr><w:spacing w:line=\"{line}\" w:lineRule=\"auto\"/></w:pPr>"),
                )
            });
            let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
            let source = crate::render::typst_gen::generate_typst(&document).unwrap();
            let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
            let body: Vec<_> = runs
                .iter()
                .filter(|run| run.text == "Continuous item")
                .collect();
            assert_eq!(body.len(), 4);
            // Metric lookup can resolve a substitute. Native absolute pitches
            // apply only when the actual painted label/body face is Cambria.
            let uses_native_fonts = body.iter().all(|run| run.family == "Cambria")
                && runs
                    .iter()
                    .filter(|run| run.text.starts_with('('))
                    .all(|run| run.family == "Cambria")
                && (suffix != "space"
                    || runs
                        .iter()
                        .filter(|run| run.text == " ")
                        .all(|run| run.family == "Arial"));
            for first in [0, 2] {
                let pitch = body[first + 1].baseline_pt - body[first].baseline_pt;
                if uses_native_fonts {
                    assert!(
                        (pitch - expected_pitch).abs() < 0.3,
                        "marker {marker_size}pt, line {line}, suffix {suffix}: pitch {pitch}, native {expected_pitch}"
                    );
                }
            }
            for label in ["(7)", "(8)", "(9)", "(10)"] {
                let marker = runs.iter().find(|run| run.text.trim() == label).unwrap();
                assert!((marker.font_size_pt - f64::from(marker_size)).abs() < 0.01);
            }
            assert!(
                body.iter()
                    .all(|run| (run.font_size_pt - 11.0).abs() < 0.01)
            );
            // The native calibration also checks that changing paragraph
            // leading leaves the label's extra ascent unchanged.
            let control = rewrite_numbering_part(&data, |xml| {
                xml.replace(
                    &format!("<w:sz w:val=\"{}\"/>", marker_size * 2),
                    "<w:sz w:val=\"22\"/>",
                )
            });
            let (control_document, _) = DocxParser
                .parse(&control, &ConvertOptions::default())
                .unwrap();
            let control_source =
                crate::render::typst_gen::generate_typst(&control_document).unwrap();
            let control_runs =
                crate::render::pdf::compiled_text_runs(&control_source.source, 0).unwrap();
            let control_body: Vec<_> = control_runs
                .iter()
                .filter(|run| run.text == "Continuous item")
                .collect();
            assert_eq!(control_body.len(), 4);
            let extra_height = (body[1].baseline_pt - body[0].baseline_pt)
                - (control_body[1].baseline_pt - control_body[0].baseline_pt);
            if uses_native_fonts {
                let natural_extra_height = marker_height_deltas
                    .entry((marker_size, suffix))
                    .or_insert(extra_height);
                assert!(
                    (extra_height - *natural_extra_height).abs() < 0.3,
                    "paragraph leading stretched label height for {marker_size}pt/{suffix}: {extra_height} vs {natural_extra_height}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_inline_markers_without_font_metrics_keep_ordinary_paragraph_fallback() {
    let data = numbering_marker_properties_fixture("<w:sz w:val=\"32\"/>", false, true);
    let (mut document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let missing_family = "Unavailable quarterly report face";
    assert!(crate::render::pdf::font_line_metrics_em(missing_family).is_none());
    for page in &mut document.pages {
        if let Page::Flow(page) = page {
            for block in &mut page.content {
                if let Block::Paragraph(paragraph) = block {
                    for run in &mut paragraph.runs {
                        run.style.font_family = Some(missing_family.to_string());
                    }
                }
            }
        }
    }
    let inline = crate::render::typst_gen::generate_typst(&document).unwrap();
    for page in &mut document.pages {
        if let Page::Flow(page) = page {
            for block in &mut page.content {
                if let Block::Paragraph(paragraph) = block {
                    paragraph.style.has_inline_list_number = false;
                }
            }
        }
    }
    let ordinary = crate::render::typst_gen::generate_typst(&document).unwrap();
    let placement = |source: &str| {
        crate::render::pdf::compiled_text_runs(source, 0)
            .unwrap()
            .into_iter()
            .map(|run| {
                (
                    run.text,
                    run.family,
                    run.font_size_pt,
                    run.left_pt,
                    run.baseline_pt,
                )
            })
            .collect::<Vec<_>>()
    };
    let inline_runs = placement(&inline.source);
    assert_eq!(
        inline_runs
            .iter()
            .filter(|run| run.0 == "Continuous item")
            .count(),
        4
    );
    assert_eq!(inline_runs, placement(&ordinary.source));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_large_inline_markers_preserve_body_continuation_advance() {
    let data = numbering_marker_properties_fixture("<w:sz w:val=\"32\"/>", false, true);
    let data = rewrite_numbering_part(&data, |xml| {
        xml.replace(
            "<w:lvlRestart w:val=\"0\"/>",
            "<w:lvlRestart w:val=\"0\"/><w:suff w:val=\"space\"/>",
        )
    });
    let text = format!(
        "Continuous item repeated body words {}",
        "describing realistic content across the entire available text measure. ".repeat(4)
    );
    let data = rewrite_docx_part(&data, "word/document.xml", |xml| {
        xml.replace("Continuous item", &text)
    });
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let first_lines: Vec<_> = runs
        .iter()
        .filter(|run| run.text.starts_with("Continuous item repeated body words"))
        .collect();
    assert_eq!(first_lines.len(), 4);
    let uses_native_fonts = first_lines.iter().all(|run| run.family == "Cambria")
        && runs
            .iter()
            .filter(|run| run.text.starts_with('('))
            .all(|run| run.family == "Cambria")
        && runs
            .iter()
            .filter(|run| run.text == " ")
            .all(|run| run.family == "Arial");
    for first in [0, 2] {
        let pitch = first_lines[first + 1].baseline_pt - first_lines[first].baseline_pt;
        if uses_native_fonts {
            assert!(
                (pitch - 73.92).abs() < 0.5,
                "item pitch {pitch}, native 73.92pt"
            );
        }
    }
    let first = first_lines[0];
    let second = first_lines[1];
    let mut continuation_baselines: Vec<f64> = runs
        .iter()
        .filter(|run| {
            run.baseline_pt > first.baseline_pt + 0.01
                && run.baseline_pt < second.baseline_pt - 0.01
                && (run.font_size_pt - 11.0).abs() < 0.01
        })
        .map(|run| run.baseline_pt)
        .collect();
    continuation_baselines.sort_by(f64::total_cmp);
    continuation_baselines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    assert_eq!(continuation_baselines.len(), 3);
    let mut previous = first.baseline_pt;
    for (baseline, native_pitch) in continuation_baselines.iter().zip([14.64, 14.88, 14.88]) {
        if uses_native_fonts {
            assert!(
                (baseline - previous - native_pitch).abs() < 0.3,
                "continuation pitch {}, native {native_pitch}",
                baseline - previous
            );
        }
        previous = *baseline;
    }
    // Triangulate on the same host's font: enlarging only the label must not
    // change later body-line advances, even when Cambria is substituted.
    let control = rewrite_numbering_part(&data, |xml| {
        xml.replace("<w:sz w:val=\"32\"/>", "<w:sz w:val=\"22\"/>")
    });
    let (control_document, _) = DocxParser
        .parse(&control, &ConvertOptions::default())
        .unwrap();
    let control_source = crate::render::typst_gen::generate_typst(&control_document).unwrap();
    let control_runs = crate::render::pdf::compiled_text_runs(&control_source.source, 0).unwrap();
    let control_first_lines: Vec<_> = control_runs
        .iter()
        .filter(|run| run.text.starts_with("Continuous item repeated body words"))
        .collect();
    assert_eq!(control_first_lines.len(), 4);
    let mut control_continuations: Vec<f64> = control_runs
        .iter()
        .filter(|run| {
            run.baseline_pt > control_first_lines[0].baseline_pt + 0.01
                && run.baseline_pt < control_first_lines[1].baseline_pt - 0.01
                && (run.font_size_pt - 11.0).abs() < 0.01
        })
        .map(|run| run.baseline_pt)
        .collect();
    control_continuations.sort_by(f64::total_cmp);
    control_continuations.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    assert_eq!(continuation_baselines.len(), control_continuations.len());
    for (actual, control) in continuation_baselines
        .windows(2)
        .zip(control_continuations.windows(2))
    {
        assert!(
            ((actual[1] - actual[0]) - (control[1] - control[0])).abs() < 0.3,
            "later body advances changed with marker size: {actual:?} vs {control:?}"
        );
    }
}
