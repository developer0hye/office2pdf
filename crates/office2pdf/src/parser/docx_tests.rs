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

#[path = "docx_inline_style_tests.rs"]
mod inline_style_tests;

#[path = "docx_paragraph_format_tests.rs"]
mod paragraph_format_tests;

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

#[path = "docx_table_structure_tests.rs"]
mod table_structure_tests;

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

#[path = "docx_list_tests.rs"]
mod list_tests;

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

#[path = "docx_hyperlink_tests.rs"]
mod hyperlink_tests;

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

#[path = "docx_layout_rtl_tests.rs"]
mod layout_rtl_tests;
#[path = "docx_math_chart_metadata_tests.rs"]
mod math_chart_metadata_tests;
