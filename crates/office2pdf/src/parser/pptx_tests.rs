use super::*;
use crate::ir::ImageCrop;
use std::io::Write;
use zip::write::FileOptions;

// ── Test helpers ─────────────────────────────────────────────────────

/// Build a minimal PPTX file as bytes from slide XML strings.
fn build_test_pptx(slide_cx_emu: i64, slide_cy_emu: i64, slide_xmls: &[String]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    for i in 0..slide_xmls.len() {
        ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
    }
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

    // ppt/presentation.xml
    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
        slide_cx_emu, slide_cy_emu
    );
    for i in 0..slide_xmls.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels
    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for i in 0..slide_xmls.len() {
        pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // Slides
    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();
    }

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Create a slide XML with the given shape elements.
fn make_slide_xml(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sld>");
    xml
}

/// Create an empty slide XML (no shapes).
fn make_empty_slide_xml() -> String {
    make_slide_xml(&[])
}

/// Create a simple text box shape XML.
fn make_text_box(x: i64, y: i64, cx: i64, cy: i64, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn make_text_box_with_body_pr(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    body_pr_xml: &str,
    text: &str,
) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody>{body_pr_xml}<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// Create a text box with formatted text runs.
fn make_formatted_text_box(x: i64, y: i64, cx: i64, cy: i64, runs_xml: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>{runs_xml}</a:p></p:txBody></p:sp>"#
    )
}

/// Create a text box with multiple paragraphs.
fn make_multi_para_text_box(x: i64, y: i64, cx: i64, cy: i64, paragraphs_xml: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>{paragraphs_xml}</p:txBody></p:sp>"#
    )
}

/// Create a slide XML with a background and optional shape elements.
fn make_slide_xml_with_bg(bg_xml: &str, shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>"#,
    );
    xml.push_str(bg_xml);
    xml.push_str(r#"<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#);
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sld>");
    xml
}

/// Standard 4:3 slide size in EMU (10" x 7.5").
const SLIDE_CX: i64 = 9_144_000;
const SLIDE_CY: i64 = 6_858_000;

/// Helper: get the first FixedPage from a Document.
fn first_fixed_page(doc: &Document) -> &FixedPage {
    match &doc.pages[0] {
        Page::Fixed(p) => p,
        _ => panic!("Expected FixedPage"),
    }
}

fn text_box_data(elem: &FixedElement) -> &TextBoxData {
    match &elem.kind {
        FixedElementKind::TextBox(text_box) => text_box,
        _ => panic!("Expected TextBox"),
    }
}

/// Helper: get the TextBox blocks from a FixedElement.
fn text_box_blocks(elem: &FixedElement) -> &[Block] {
    &text_box_data(elem).content
}

// ── Tests ────────────────────────────────────────────────────────────

#[test]
fn test_parse_empty_presentation() {
    // PPTX with zero slides → document with no pages
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(doc.pages.is_empty(), "Expected no pages");
}

#[test]
fn test_parse_single_slide() {
    let slide = make_empty_slide_xml();
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.pages.len(), 1, "Expected 1 page");
    assert!(matches!(&doc.pages[0], Page::Fixed(_)));
}

#[test]
fn test_slide_dimensions() {
    // 16:9 widescreen: 12192000 × 6858000 EMU = 960pt × 540pt
    let cx = 12_192_000i64;
    let cy = 6_858_000i64;
    let slide = make_empty_slide_xml();
    let data = build_test_pptx(cx, cy, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let expected_w = cx as f64 / 12700.0;
    let expected_h = cy as f64 / 12700.0;
    assert!(
        (page.size.width - expected_w).abs() < 0.1,
        "Expected width ~{expected_w}pt, got {}",
        page.size.width
    );
    assert!(
        (page.size.height - expected_h).abs() < 0.1,
        "Expected height ~{expected_h}pt, got {}",
        page.size.height
    );
}

#[test]
fn test_text_box_extraction() {
    let shape = make_text_box(0, 0, 1_000_000, 500_000, "Hello World");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 element");

    let blocks = text_box_blocks(&page.elements[0]);
    assert!(!blocks.is_empty(), "Expected at least one block");

    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 1);
    assert_eq!(para.runs[0].text, "Hello World");
}

#[test]
fn test_text_box_auto_numbered_paragraphs_group_into_list() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Ordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("1.")
    );
    assert_eq!(list.items[0].content[0].runs[0].text, "First");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second");
}

#[test]
fn test_text_box_bulleted_paragraphs_group_into_list() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>Second bullet</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Unordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
}

#[test]
fn test_text_box_bulleted_paragraph_preserves_char_marker_and_uses_run_style() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buFontTx/><a:buChar char="-"/></a:pPr>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Pretendard"/></a:rPr><a:t>First bullet</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.marker_text.as_deref(), Some("-"));
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Pretendard")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(14.0)
    );
    assert_eq!(
        style.marker_style.as_ref().and_then(|style| style.color),
        Some(Color::new(0x11, 0x22, 0x33))
    );
}

#[test]
fn test_text_box_bulleted_paragraph_preserves_explicit_marker_font() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Symbol bullet</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.marker_text.as_deref(), Some("è"));
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Wingdings")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(14.0)
    );
}

#[test]
fn test_text_box_paragraph_line_spacing_pct_extracted() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Proportional(factor)) => {
            assert!((factor - 1.5).abs() < f64::EPSILON);
        }
        other => panic!("Expected proportional line spacing, got {other:?}"),
    }
}

#[test]
fn test_text_box_body_pr_defaults_and_center_anchor_extracted() {
    let shape = make_text_box_with_body_pr(
        0,
        0,
        1_000_000,
        500_000,
        r#"<a:bodyPr anchor="ctr"><a:spAutoFit/></a:bodyPr>"#,
        "Centered",
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let text_box = match &page.elements[0].kind {
        FixedElementKind::TextBox(text_box) => text_box,
        other => panic!("Expected TextBox, got {other:?}"),
    };
    assert!((text_box.padding.left - 7.2).abs() < 0.001);
    assert!((text_box.padding.right - 7.2).abs() < 0.001);
    assert!((text_box.padding.top - 3.6).abs() < 0.001);
    assert!((text_box.padding.bottom - 3.6).abs() < 0.001);
    assert_eq!(
        text_box.vertical_align,
        crate::ir::TextBoxVerticalAlign::Center
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_start_override_sets_list_start() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod" startAt="3"/></a:pPr><a:r><a:t>Gamma</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod"/></a:pPr><a:r><a:t>Delta</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Ordered);
    assert_eq!(list.items[0].start_at, Some(3));
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("A.")
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_extracts_hanging_indent() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };

    let paragraph = &list.items[0].content[0];
    assert_eq!(paragraph.style.indent_left, Some(36.0));
    assert_eq!(paragraph.style.indent_first_line, Some(-36.0));
    assert_eq!(
        list.level_styles
            .get(&0)
            .and_then(|style| style.numbering_pattern.as_deref()),
        Some("1)")
    );
}

#[test]
fn test_text_box_auto_numbered_paragraph_resolves_marker_style_from_text() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
        r#"<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buAutoNum type="arabicParenR"/>"#,
        r#"</a:pPr>"#,
        r#"<a:r><a:rPr lang="ko-KR" sz="2000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Pretendard Medium"/></a:rPr><a:t>First</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    let style = list.level_styles.get(&0).expect("Expected level 0 style");
    assert_eq!(style.numbering_pattern.as_deref(), Some("1)"));
    assert_eq!(style.marker_text, None);
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_family.as_deref()),
        Some("Pretendard Medium")
    );
    assert_eq!(
        style
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(20.0)
    );
    assert_eq!(
        style.marker_style.as_ref().and_then(|style| style.color),
        Some(Color::black())
    );
}

#[test]
fn test_text_box_paragraph_preserves_soft_line_breaks() {
    let paragraphs_xml = concat!(
        r#"<a:p>"#,
        r#"<a:r><a:t>Line 1</a:t></a:r>"#,
        r#"<a:br/>"#,
        r#"<a:r><a:t>Line 2</a:t></a:r>"#,
        r#"</a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "Line 1\u{000B}Line 2");
}

#[test]
fn test_text_box_plain_paragraph_between_bullets_breaks_list_sequence() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>1) First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>2) Second bullet</a:t></a:r></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    assert_eq!(blocks.len(), 3, "Expected list / paragraph / list split");
    match &blocks[0] {
        Block::List(list) => {
            assert_eq!(list.items.len(), 1);
            assert_eq!(
                list.level_styles
                    .get(&1)
                    .and_then(|style| style.marker_text.as_deref()),
                Some("-")
            );
        }
        other => panic!("Expected first block to be a list, got {other:?}"),
    }
    match &blocks[1] {
        Block::Paragraph(paragraph) => {
            let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
            assert_eq!(text, "-> Continuation paragraph");
        }
        other => panic!("Expected middle block to be a paragraph, got {other:?}"),
    }
    match &blocks[2] {
        Block::List(list) => {
            assert_eq!(list.items.len(), 1);
            assert_eq!(
                list.level_styles
                    .get(&1)
                    .and_then(|style| style.marker_text.as_deref()),
                Some("-")
            );
        }
        other => panic!("Expected last block to be a list, got {other:?}"),
    }
}

#[test]
fn test_text_box_plain_paragraph_preserves_leading_arrow_text() {
    let paragraphs_xml = r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "-> Continuation paragraph");
}

#[test]
fn test_text_box_plain_paragraph_preserves_escaped_gt_entity() {
    let paragraphs_xml = r#"<a:p><a:r><a:t>-&gt; Continuation paragraph</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "-> Continuation paragraph");
}

#[test]
fn test_text_box_trailing_empty_bullets_do_not_override_nested_marker_style() {
    let paragraphs_xml = concat!(
        r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Arrow bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
    );
    let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let list = match &blocks[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.items.len(), 1);
    assert_eq!(list.items[0].level, 1);
    assert_eq!(
        list.level_styles
            .get(&1)
            .and_then(|style| style.marker_text.as_deref()),
        Some("è")
    );
    assert!(
        !list.level_styles.contains_key(&0),
        "Trailing empty dash bullets should not create a level-0 marker style"
    );
}

#[test]
fn test_text_box_lst_style_default_run_props_are_applied_to_runs() {
    let shape = String::from(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="032543"/></a:solidFill><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p></p:txBody></p:sp>"#,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let run = &paragraph.runs[0];
    assert_eq!(
        run.style.font_family.as_deref(),
        Some("Pretendard SemiBold")
    );
    assert_eq!(run.style.font_size, Some(14.0));
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.color, Some(Color::new(0x03, 0x25, 0x43)));
}

#[test]
fn test_non_placeholder_shape_inherits_master_other_style_run_defaults() {
    let slide_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Caption"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#,
        r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>신</a:t></a:r><a:r><a:rPr lang="ko-KR" sz="1800"/><a:t>형</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp>"#,
    );
    let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="224466"/></a:solidFill><a:latin typeface="Pretendard"/><a:ea typeface="Pretendard"/><a:cs typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "신형");
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.font_size == Some(18.0))
    );
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.font_family.as_deref() == Some("Pretendard"))
    );
    assert!(
        paragraph
            .runs
            .iter()
            .all(|run| run.style.color == Some(Color::new(0x22, 0x44, 0x66)))
    );
}

#[test]
fn test_text_box_lst_style_overrides_master_other_style_run_defaults() {
    let slide_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle>"#,
        r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp>"#,
    );
    let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:latin typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph block, got {other:?}"),
    };
    assert_eq!(paragraph.runs[0].style.font_size, Some(24.0));
    assert_eq!(
        paragraph.runs[0].style.font_family.as_deref(),
        Some("Pretendard SemiBold")
    );
}

#[test]
fn test_text_box_position_and_size() {
    // Position: 1000000 EMU x, 500000 EMU y → ~78.74pt, ~39.37pt
    // Size: 5000000 EMU cx, 2000000 EMU cy → ~393.70pt, ~157.48pt
    let x = 1_000_000i64;
    let y = 500_000i64;
    let cx = 5_000_000i64;
    let cy = 2_000_000i64;
    let shape = make_text_box(x, y, cx, cy, "Positioned");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];

    let expected_x = x as f64 / 12700.0;
    let expected_y = y as f64 / 12700.0;
    let expected_w = cx as f64 / 12700.0;
    let expected_h = cy as f64 / 12700.0;

    assert!(
        (elem.x - expected_x).abs() < 0.1,
        "Expected x ~{expected_x}, got {}",
        elem.x
    );
    assert!(
        (elem.y - expected_y).abs() < 0.1,
        "Expected y ~{expected_y}, got {}",
        elem.y
    );
    assert!(
        (elem.width - expected_w).abs() < 0.1,
        "Expected width ~{expected_w}, got {}",
        elem.width
    );
    assert!(
        (elem.height - expected_h).abs() < 0.1,
        "Expected height ~{expected_h}, got {}",
        elem.height
    );
}

#[test]
fn test_text_box_bold_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Bold text");
    assert_eq!(para.runs[0].style.bold, Some(true));
}

#[test]
fn test_text_box_italic_formatting() {
    let runs_xml = r#"<a:r><a:rPr i="1"/><a:t>Italic text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Italic text");
    assert_eq!(para.runs[0].style.italic, Some(true));
}

#[test]
fn test_text_box_font_size() {
    // sz="2400" means 24pt (hundredths of a point)
    let runs_xml = r#"<a:r><a:rPr sz="2400"/><a:t>Large text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].style.font_size, Some(24.0));
}

#[test]
fn test_text_box_combined_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Styled text");
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.italic, Some(true));
    assert_eq!(run.style.underline, Some(true));
    assert_eq!(run.style.strikethrough, Some(true));
    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    assert_eq!(run.style.font_family, Some("Arial".to_string()));
}

#[test]
fn test_multiple_text_boxes() {
    let shape1 = make_text_box(100_000, 100_000, 2_000_000, 500_000, "Box 1");
    let shape2 = make_text_box(100_000, 700_000, 2_000_000, 500_000, "Box 2");
    let slide = make_slide_xml(&[shape1, shape2]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 text boxes");

    // Check content of each box
    let get_text = |elem: &FixedElement| -> String {
        let blocks = text_box_blocks(elem);
        blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect()
    };
    assert_eq!(get_text(&page.elements[0]), "Box 1");
    assert_eq!(get_text(&page.elements[1]), "Box 2");
}

#[test]
fn test_multiple_slides() {
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 3")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 3, "Expected 3 pages");
    for page in &doc.pages {
        assert!(matches!(page, Page::Fixed(_)));
    }
}

#[test]
fn test_text_box_multiple_paragraphs() {
    let paras_xml = r#"<a:p><a:r><a:rPr/><a:t>Paragraph 1</a:t></a:r></a:p><a:p><a:r><a:rPr/><a:t>Paragraph 2</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 3_000_000, 2_000_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paras: Vec<&Paragraph> = blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert!(paras.len() >= 2, "Expected at least 2 paragraphs");
    assert_eq!(paras[0].runs[0].text, "Paragraph 1");
    assert_eq!(paras[1].runs[0].text, "Paragraph 2");
}

#[test]
fn test_text_box_multiple_runs() {
    let runs_xml =
        r#"<a:r><a:rPr b="1"/><a:t>Bold </a:t></a:r><a:r><a:rPr i="1"/><a:t>Italic</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Bold ");
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert_eq!(para.runs[1].text, "Italic");
    assert_eq!(para.runs[1].style.italic, Some(true));
}

#[test]
fn test_paragraph_alignment_center() {
    let paras_xml = r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr/><a:t>Centered</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 2_000_000, 500_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_parse_invalid_data() {
    let parser = PptxParser;
    let result = parser.parse(b"not a valid pptx file", &ConvertOptions::default());
    assert!(result.is_err());
    match result.unwrap_err() {
        ConvertError::Parse(_) => {}
        other => panic!("Expected Parse error, got: {other:?}"),
    }
}

#[test]
fn test_slide_default_dimensions_4x3() {
    // Standard 4:3: 9144000 × 6858000 EMU = 720pt × 540pt
    let slide = make_empty_slide_xml();
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert!(
        (page.size.width - 720.0).abs() < 0.1,
        "Expected width ~720pt, got {}",
        page.size.width
    );
    assert!(
        (page.size.height - 540.0).abs() < 0.1,
        "Expected height ~540pt, got {}",
        page.size.height
    );
}

// ── Shape test helpers ───────────────────────────────────────────────

/// Create a shape XML element with preset geometry, optional fill and border.
#[allow(clippy::too_many_arguments)]
fn make_shape(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    fill_hex: Option<&str>,
    border_width_emu: Option<i64>,
    border_hex: Option<&str>,
) -> String {
    let fill_xml = fill_hex
        .map(|h| format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#))
        .unwrap_or_default();

    let ln_xml = match (border_width_emu, border_hex) {
        (Some(w), Some(h)) => {
            format!(r#"<a:ln w="{w}"><a:solidFill><a:srgbClr val="{h}"/></a:solidFill></a:ln>"#)
        }
        _ => String::new(),
    };

    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</p:spPr></p:sp>"#
    )
}

/// Helper: extract the Shape from a FixedElement or panic.
fn get_shape(elem: &FixedElement) -> &Shape {
    match &elem.kind {
        FixedElementKind::Shape(s) => s,
        other => panic!("Expected Shape, got {other:?}"),
    }
}

// ── Image test helpers ───────────────────────────────────────────────

/// Create a minimal valid BMP (1×1 pixel, red) for test images.
fn make_test_bmp() -> Vec<u8> {
    let mut bmp = Vec::new();
    // BMP header (14 bytes)
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&70u32.to_le_bytes()); // file size
    bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
    bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset
    // DIB header (40 bytes)
    bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
    bmp.extend_from_slice(&1i32.to_le_bytes()); // width
    bmp.extend_from_slice(&1i32.to_le_bytes()); // height
    bmp.extend_from_slice(&1u16.to_le_bytes()); // planes
    bmp.extend_from_slice(&24u16.to_le_bytes()); // bpp
    bmp.extend_from_slice(&0u32.to_le_bytes()); // compression
    bmp.extend_from_slice(&16u32.to_le_bytes()); // image size
    bmp.extend_from_slice(&2835u32.to_le_bytes()); // h resolution
    bmp.extend_from_slice(&2835u32.to_le_bytes()); // v resolution
    bmp.extend_from_slice(&0u32.to_le_bytes()); // colors
    bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors
    // Pixel data: 1 pixel (BGR) + 1 byte padding to align to 4 bytes
    bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);
    bmp
}

/// Create a minimal valid SVG image for test images.
fn make_test_svg() -> Vec<u8> {
    br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1" viewBox="0 0 1 1"><rect width="1" height="1" fill="#ff0000"/></svg>"##.to_vec()
}

/// Create a picture XML element referencing an image via relationship ID.
fn make_pic_xml(x: i64, y: i64, cx: i64, cy: i64, r_embed: &str) -> String {
    make_custom_pic_xml(
        x,
        y,
        cx,
        cy,
        &format!(r#"<a:blip r:embed="{r_embed}"/><a:stretch><a:fillRect/></a:stretch>"#),
    )
}

/// Create a picture XML element with custom `<p:blipFill>` contents.
fn make_custom_pic_xml(x: i64, y: i64, cx: i64, cy: i64, blip_fill_xml: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip_fill_xml}</p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr></p:pic>"#
    )
}

/// Slide image for the test PPTX builder.
struct TestSlideImage {
    rid: String,
    path: String,
    data: Vec<u8>,
    relationship_type: Option<String>,
}

/// Build a PPTX file with slides that have image relationships.
fn build_test_pptx_with_images(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slides: &[(String, Vec<TestSlideImage>)],
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    ct.push_str(r#"<Default Extension="png" ContentType="image/png"/>"#);
    ct.push_str(r#"<Default Extension="bmp" ContentType="image/bmp"/>"#);
    ct.push_str(r#"<Default Extension="jpeg" ContentType="image/jpeg"/>"#);
    ct.push_str(r#"<Default Extension="svg" ContentType="image/svg+xml"/>"#);
    for i in 0..slides.len() {
        ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
    }
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

    // ppt/presentation.xml
    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
        slide_cx_emu, slide_cy_emu
    );
    for i in 0..slides.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels
    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for i in 0..slides.len() {
        pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // Slides and their .rels files
    for (i, (slide_xml, slide_images)) in slides.iter().enumerate() {
        let slide_num = i + 1;

        // Write slide XML
        zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        // Write slide .rels if there are images
        if !slide_images.is_empty() {
            let mut rels = String::from(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            );
            for img in slide_images {
                rels.push_str(&format!(
                    r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                    img.rid,
                    img.relationship_type.as_deref().unwrap_or(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                    ),
                    img.path
                ));
            }
            rels.push_str("</Relationships>");
            zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
                .unwrap();
            zip.write_all(rels.as_bytes()).unwrap();

            // Write image media files
            for img in slide_images {
                // Resolve the relative path (e.g., "../media/image1.png" → "ppt/media/image1.png")
                let media_path = resolve_relative_path("ppt/slides", &img.path);
                zip.start_file(media_path, opts).unwrap();
                zip.write_all(&img.data).unwrap();
            }
        }
    }

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Helper: get the ImageData from a FixedElement or panic.
fn get_image(elem: &FixedElement) -> &ImageData {
    match &elem.kind {
        FixedElementKind::Image(img) => img,
        other => panic!("Expected Image, got {other:?}"),
    }
}

// ── Shape tests ──────────────────────────────────────────────────────

#[test]
fn test_shape_rectangle_with_fill() {
    let shape = make_shape(
        1_000_000,
        500_000,
        3_000_000,
        2_000_000,
        "rect",
        Some("FF0000"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 shape element");

    let elem = &page.elements[0];
    assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
    assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
    assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
    assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

    let shape = get_shape(elem);
    assert!(matches!(shape.kind, ShapeKind::Rectangle));
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
    assert!(shape.stroke.is_none());
}

#[test]
fn test_shape_ellipse() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "ellipse",
        Some("00FF00"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(matches!(s.kind, ShapeKind::Ellipse));
    assert_eq!(s.fill, Some(Color::new(0, 255, 0)));
}

#[test]
fn test_shape_line() {
    let shape = make_shape(
        500_000,
        1_000_000,
        4_000_000,
        0,
        "line",
        None,
        Some(25400),
        Some("0000FF"),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    match &s.kind {
        ShapeKind::Line { x2, y2 } => {
            assert!((*x2 - emu_to_pt(4_000_000)).abs() < 0.1);
            assert!((*y2 - 0.0).abs() < 0.1);
        }
        _ => panic!("Expected Line shape"),
    }
    assert!(s.fill.is_none());
    let stroke = s.stroke.as_ref().expect("Expected stroke on line");
    assert!((stroke.width - 2.0).abs() < 0.1); // 25400 EMU = 2pt
    assert_eq!(stroke.color, Color::new(0, 0, 255));
}

#[test]
fn test_shape_with_fill_and_border() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "rect",
        Some("FFFF00"),
        Some(12700),
        Some("000000"),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert_eq!(s.fill, Some(Color::new(255, 255, 0)));
    let stroke = s.stroke.as_ref().expect("Expected stroke");
    assert!((stroke.width - 1.0).abs() < 0.1); // 12700 EMU = 1pt
    assert_eq!(stroke.color, Color::black());
}

#[test]
fn test_shape_no_fill_no_border() {
    let shape = make_shape(0, 0, 1_000_000, 1_000_000, "rect", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(s.fill.is_none());
    assert!(s.stroke.is_none());
}

#[test]
fn test_multiple_shapes_on_slide() {
    let rect = make_shape(
        0,
        0,
        1_000_000,
        1_000_000,
        "rect",
        Some("FF0000"),
        None,
        None,
    );
    let ellipse = make_shape(
        2_000_000,
        0,
        1_000_000,
        1_000_000,
        "ellipse",
        Some("00FF00"),
        None,
        None,
    );
    let slide = make_slide_xml(&[rect, ellipse]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 shape elements");
    assert!(matches!(
        get_shape(&page.elements[0]).kind,
        ShapeKind::Rectangle
    ));
    assert!(matches!(
        get_shape(&page.elements[1]).kind,
        ShapeKind::Ellipse
    ));
}

#[test]
fn test_shapes_and_text_boxes_mixed() {
    let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Hello");
    let rect = make_shape(
        0,
        1_000_000,
        2_000_000,
        500_000,
        "rect",
        Some("FF0000"),
        None,
        None,
    );
    let slide = make_slide_xml(&[text_box, rect]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 elements");
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
}

// ── Image tests ──────────────────────────────────────────────────────

#[test]
fn test_image_basic_extraction() {
    let bmp_data = make_test_bmp();
    let pic = make_pic_xml(1_000_000, 500_000, 3_000_000, 2_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let elem = &page.elements[0];
    assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
    assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
    assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
    assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

    let img = get_image(elem);
    assert!(!img.data.is_empty(), "Image data should not be empty");
    assert_eq!(img.data, bmp_data);
}

#[test]
fn test_image_format_detection() {
    let bmp_data = make_test_bmp();

    // Test BMP format
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Bmp);
}

#[test]
fn test_svg_image_extraction() {
    let svg_data = make_test_svg();

    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.svg".to_string(),
        data: svg_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Svg);
    assert_eq!(img.data, svg_data);
}

#[test]
fn test_image_blip_start_tag_with_children_is_extracted() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let img = get_image(&page.elements[0]);
    assert_eq!(img.data, bmp_data);
}

#[test]
fn test_svg_blip_is_preferred_over_base_raster() {
    let bmp_data = make_test_bmp();
    let svg_data = make_test_svg();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId4"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.svg".to_string(),
            data: svg_data.clone(),
            relationship_type: None,
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Svg);
    assert_eq!(img.data, svg_data);
}

#[test]
fn test_src_rect_crop_is_extracted() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        2_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"/><a:srcRect l="25000" t="10000" r="5000" b="20000"/><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(
        img.crop,
        Some(ImageCrop {
            left: 0.25,
            top: 0.10,
            right: 0.05,
            bottom: 0.20,
        })
    );
}

#[test]
fn test_unsupported_img_layer_emits_partial_warning_but_keeps_base_image() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId4"/></a14:imgProps></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.wdp".to_string(),
            data: vec![0x00, 0x01, 0x02],
            relationship_type: Some(
                "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
            ),
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Base image should still render");
    assert_eq!(get_image(&page.elements[0]).data, bmp_data);
    assert!(
        warnings.iter().any(|warning| matches!(
            warning,
            ConvertWarning::PartialElement { format, element, detail }
                if format == "PPTX"
                    && element.contains("slide 1")
                    && detail.contains("image layer")
                    && detail.contains("image2.wdp")
        )),
        "Expected partial warning for unsupported image layer, got: {warnings:?}"
    );
}

#[test]
fn test_wdp_only_picture_emits_unsupported_warning() {
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.wdp".to_string(),
        data: vec![0x00, 0x01, 0x02],
        relationship_type: Some(
            "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
        ),
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        0,
        "Unsupported WDP image should be omitted"
    );
    assert!(
        warnings.iter().any(|warning| matches!(
            warning,
            ConvertWarning::UnsupportedElement { format, element }
                if format == "PPTX"
                    && element.contains("slide 1")
                    && element.contains("image1.wdp")
        )),
        "Expected unsupported warning for WDP-only picture, got: {warnings:?}"
    );
}

#[test]
fn test_image_dimensions_preserved() {
    let bmp_data = make_test_bmp();
    // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000 EMU
    let pic = make_pic_xml(0, 0, 2_540_000, 1_270_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    let w = img.width.expect("Expected width");
    let h = img.height.expect("Expected height");
    assert!((w - 200.0).abs() < 0.1, "Expected ~200pt, got {w}");
    assert!((h - 100.0).abs() < 0.1, "Expected ~100pt, got {h}");
}

#[test]
fn test_image_with_shapes_and_text() {
    let bmp_data = make_test_bmp();
    let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Title");
    let rect = make_shape(
        0,
        600_000,
        1_000_000,
        500_000,
        "rect",
        Some("AABBCC"),
        None,
        None,
    );
    let pic = make_pic_xml(2_000_000, 600_000, 1_500_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[text_box, rect, pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 3, "Expected 3 elements");
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
    assert!(matches!(&page.elements[2].kind, FixedElementKind::Image(_)));
}

#[test]
fn test_image_missing_rid_ignored() {
    // Picture references rId3 but no image data for that rId
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId99");
    let slide_xml = make_slide_xml(&[pic]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        0,
        "Missing image ref should be skipped"
    );
}

#[test]
fn test_multiple_images_on_slide() {
    let bmp_data = make_test_bmp();
    let pic1 = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let pic2 = make_pic_xml(2_000_000, 0, 1_500_000, 1_000_000, "rId4");
    let slide_xml = make_slide_xml(&[pic1, pic2]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 image elements");
    assert!(matches!(&page.elements[0].kind, FixedElementKind::Image(_)));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Image(_)));
}

// ── Theme test helpers ────────────────────────────────────────────

/// Create a theme XML with the given color scheme and font scheme.
fn make_theme_xml(colors: &[(&str, &str)], major_font: &str, minor_font: &str) -> String {
    let mut color_xml = String::new();
    for (name, hex) in colors {
        // dk1/lt1 use sysClr in real files; others use srgbClr
        if *name == "dk1" || *name == "lt1" {
            color_xml.push_str(&format!(
                r#"<a:{name}><a:sysClr val="windowText" lastClr="{hex}"/></a:{name}>"#
            ));
        } else {
            color_xml.push_str(&format!(r#"<a:{name}><a:srgbClr val="{hex}"/></a:{name}>"#));
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Test">{color_xml}</a:clrScheme><a:fontScheme name="Test"><a:majorFont><a:latin typeface="{major_font}"/></a:majorFont><a:minorFont><a:latin typeface="{minor_font}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
    )
}

/// Standard theme color set used in tests.
fn standard_theme_colors() -> Vec<(&'static str, &'static str)> {
    vec![
        ("dk1", "000000"),
        ("dk2", "1F4D78"),
        ("lt1", "FFFFFF"),
        ("lt2", "E7E6E6"),
        ("accent1", "4472C4"),
        ("accent2", "ED7D31"),
        ("accent3", "A5A5A5"),
        ("accent4", "FFC000"),
        ("accent5", "5B9BD5"),
        ("accent6", "70AD47"),
        ("hlink", "0563C1"),
        ("folHlink", "954F72"),
    ]
}

/// Build a test PPTX with a theme file included.
fn build_test_pptx_with_theme(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xmls: &[String],
    theme_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    for i in 0..slide_xmls.len() {
        ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
    }
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

    // ppt/presentation.xml
    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
        slide_cx_emu, slide_cy_emu
    );
    for i in 0..slide_xmls.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels (includes theme relationship)
    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    // Theme relationship (rId1 in pres rels)
    pres_rels.push_str(
            r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        );
    for i in 0..slide_xmls.len() {
        pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // ppt/theme/theme1.xml
    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(theme_xml.as_bytes()).unwrap();

    // Slides
    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();
    }

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Build a test PPTX with a single slide that has layout and master relationships.
///
/// Creates: slide1 → slideLayout1 → slideMaster1
fn build_test_pptx_with_layout_master(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xml: &str,
    layout_xml: &str,
    master_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let ct = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#;
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

    // ppt/presentation.xml
    let pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    );
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels
    let pres_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#;
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // ppt/slides/slide1.xml
    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide_xml.as_bytes()).unwrap();

    // ppt/slides/_rels/slide1.xml.rels → points to layout
    let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(slide_rels.as_bytes()).unwrap();

    // ppt/slideLayouts/slideLayout1.xml
    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(layout_xml.as_bytes()).unwrap();

    // ppt/slideLayouts/_rels/slideLayout1.xml.rels → points to master
    let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
        .unwrap();
    zip.write_all(layout_rels.as_bytes()).unwrap();

    // ppt/slideMasters/slideMaster1.xml
    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(master_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Build a test PPTX with a single slide, layout/master chain, and theme.
fn build_test_pptx_with_theme_layout_master(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xml: &str,
    layout_xml: &str,
    master_xml: &str,
    theme_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    let ct = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#;
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

    let pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    );
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    let pres_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(theme_xml.as_bytes()).unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide_xml.as_bytes()).unwrap();

    let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(slide_rels.as_bytes()).unwrap();

    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(layout_xml.as_bytes()).unwrap();

    let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
        .unwrap();
    zip.write_all(layout_rels.as_bytes()).unwrap();

    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(master_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Build a test PPTX with multiple slides that all share the same layout and master.
fn build_test_pptx_with_layout_master_multi_slide(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xmls: &[String],
    layout_xml: &str,
    master_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    for i in 0..slide_xmls.len() {
        ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
    }
    ct.push_str(r#"<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>"#);
    ct.push_str(r#"<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>"#);
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

    // ppt/presentation.xml
    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst>"#,
    );
    for i in 0..slide_xmls.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels
    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    pres_rels.push_str(
            r#"<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#,
        );
    for i in 0..slide_xmls.len() {
        pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // Slides and their .rels
    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        let slide_num = i + 1;
        zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        // Each slide's .rels points to the shared layout
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();
    }

    // ppt/slideLayouts/slideLayout1.xml
    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(layout_xml.as_bytes()).unwrap();

    // ppt/slideLayouts/_rels/slideLayout1.xml.rels → points to master
    let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
        .unwrap();
    zip.write_all(layout_rels.as_bytes()).unwrap();

    // ppt/slideMasters/slideMaster1.xml
    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(master_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

// ── Theme unit tests ──────────────────────────────────────────────

#[test]
fn test_parse_theme_xml_colors() {
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let theme = parse_theme_xml(&theme_xml);

    assert_eq!(theme.colors.len(), 12);
    assert_eq!(theme.colors["dk1"], Color::new(0, 0, 0));
    assert_eq!(theme.colors["lt1"], Color::new(255, 255, 255));
    assert_eq!(theme.colors["accent1"], Color::new(0x44, 0x72, 0xC4));
    assert_eq!(theme.colors["accent2"], Color::new(0xED, 0x7D, 0x31));
    assert_eq!(theme.colors["hlink"], Color::new(0x05, 0x63, 0xC1));
    assert_eq!(theme.colors["folHlink"], Color::new(0x95, 0x4F, 0x72));
}

#[test]
fn test_parse_theme_xml_fonts() {
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let theme = parse_theme_xml(&theme_xml);

    assert_eq!(theme.major_font, Some("Calibri Light".to_string()));
    assert_eq!(theme.minor_font, Some("Calibri".to_string()));
}

#[test]
fn test_parse_theme_xml_sys_clr() {
    // dk1 and lt1 use sysClr with lastClr attribute
    let theme_xml = make_theme_xml(&[("dk1", "111111"), ("lt1", "EEEEEE")], "Arial", "Arial");
    let theme = parse_theme_xml(&theme_xml);

    assert_eq!(theme.colors["dk1"], Color::new(0x11, 0x11, 0x11));
    assert_eq!(theme.colors["lt1"], Color::new(0xEE, 0xEE, 0xEE));
}

#[test]
fn test_parse_theme_xml_empty() {
    let theme = parse_theme_xml("");
    assert!(theme.colors.is_empty());
    assert!(theme.major_font.is_none());
    assert!(theme.minor_font.is_none());
}

#[test]
fn test_resolve_theme_font_major() {
    let theme = ThemeData {
        major_font: Some("Calibri Light".to_string()),
        minor_font: Some("Calibri".to_string()),
        ..ThemeData::default()
    };
    assert_eq!(resolve_theme_font("+mj-lt", &theme), "Calibri Light");
}

#[test]
fn test_resolve_theme_font_minor() {
    let theme = ThemeData {
        major_font: Some("Calibri Light".to_string()),
        minor_font: Some("Calibri".to_string()),
        ..ThemeData::default()
    };
    assert_eq!(resolve_theme_font("+mn-lt", &theme), "Calibri");
}

#[test]
fn test_resolve_theme_font_explicit() {
    let theme = ThemeData::default();
    assert_eq!(resolve_theme_font("Arial", &theme), "Arial");
}

// ── Theme integration tests (full PPTX parsing) ───────────────────

#[test]
fn test_scheme_color_in_shape_fill() {
    // Shape with <a:schemeClr val="accent1"/> should resolve to accent1 color
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0x44, 0x72, 0xC4)));
}

#[test]
fn test_scheme_color_in_line_stroke() {
    // Shape border using scheme color
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:ln></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let stroke = shape.stroke.as_ref().expect("Expected stroke");
    assert_eq!(stroke.color, Color::new(0, 0, 0)); // dk1 = black
}

#[test]
fn test_scheme_color_in_text_run() {
    // Text run using <a:schemeClr val="accent2"/>
    let runs_xml = r#"<a:r><a:rPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:rPr><a:t>Themed text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Themed text");
    assert_eq!(para.runs[0].style.color, Some(Color::new(0xED, 0x7D, 0x31)));
}

#[test]
fn test_theme_major_font_in_text() {
    // Text with <a:latin typeface="+mj-lt"/> should resolve to major font
    let runs_xml = r#"<a:r><a:rPr><a:latin typeface="+mj-lt"/></a:rPr><a:t>Heading</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Heading");
    assert_eq!(
        para.runs[0].style.font_family,
        Some("Calibri Light".to_string())
    );
}

#[test]
fn test_theme_minor_font_in_text() {
    // Text with <a:latin typeface="+mn-lt"/> should resolve to minor font
    let runs_xml = r#"<a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>Body text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Body text");
    assert_eq!(para.runs[0].style.font_family, Some("Calibri".to_string()));
}

#[test]
fn test_pptx_with_theme_colors_and_fonts_combined() {
    // Full test: shape with scheme color + text with scheme color and theme font
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent5"/></a:solidFill></p:spPr></p:sp>"#;
    let runs_xml = r#"<a:r><a:rPr b="1" sz="2400"><a:solidFill><a:schemeClr val="dk2"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Theme styled</a:t></a:r>"#;
    let text_box = make_formatted_text_box(3_000_000, 0, 4_000_000, 1_000_000, runs_xml);
    let slide = make_slide_xml(&[shape_xml.to_string(), text_box]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2);

    // Shape fill = accent5
    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0x5B, 0x9B, 0xD5)));

    // Text run: color = dk2, font = major font, bold, 24pt
    let blocks = text_box_blocks(&page.elements[1]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Theme styled");
    assert_eq!(run.style.color, Some(Color::new(0x1F, 0x4D, 0x78)));
    assert_eq!(run.style.font_family, Some("Calibri Light".to_string()));
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.font_size, Some(24.0));
}

#[test]
fn test_no_theme_scheme_color_ignored() {
    // When there's no theme, schemeClr references should produce None
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    // Use regular build_test_pptx (no theme)
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    // No theme → scheme color not resolved → fill is None
    assert!(shape.fill.is_none());
}

#[test]
fn test_scheme_color_as_start_element() {
    // schemeClr can have children like <a:tint val="50000"/>, test it still works
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"><a:tint val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    // Color is resolved from the scheme (tint is ignored for now but base color is read)
    assert_eq!(shape.fill, Some(Color::new(0xA5, 0xA5, 0xA5)));
}

#[test]
fn test_scheme_color_lum_mod_applies_to_shape_fill() {
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(
        &[("dk1", "000000"), ("lt1", "FFFFFF"), ("accent1", "808080")],
        "Calibri Light",
        "Calibri",
    );
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0x40, 0x40, 0x40)));
}

#[test]
fn test_layout_shape_uses_master_color_map_with_luminance_offset() {
    let slide_xml = make_empty_slide_xml();
    let layout_shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="tx1"><a:lumOff val="50000"/></a:schemeClr></a:solidFill><a:ln w="6350"><a:noFill/></a:ln></p:spPr></p:sp>"#;
    let layout_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{layout_shape}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#
    );
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
    let theme_xml = make_theme_xml(
        &[
            ("dk1", "000000"),
            ("dk2", "222222"),
            ("lt1", "FFFFFF"),
            ("lt2", "EEEEEE"),
            ("accent1", "4472C4"),
            ("accent2", "ED7D31"),
            ("accent3", "A5A5A5"),
            ("accent4", "FFC000"),
            ("accent5", "5B9BD5"),
            ("accent6", "70AD47"),
            ("hlink", "0563C1"),
            ("folHlink", "954F72"),
        ],
        "Calibri Light",
        "Calibri",
    );
    let data = build_test_pptx_with_theme_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        master_xml,
        &theme_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0x80, 0x80, 0x80)));
}

// ── Slide background tests ───────────────────────────────────────────

#[test]
fn test_slide_solid_color_background() {
    // Slide with a solid red background via <p:bg>
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let slide = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_slide_no_background() {
    // Slide with no <p:bg> → background_color is None
    let slide = make_empty_slide_xml();
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert!(page.background_color.is_none());
}

#[test]
fn test_slide_background_with_scheme_color() {
    // Slide background using a theme scheme color reference
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let slide = make_slide_xml_with_bg(bg_xml, &[]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(0x44, 0x72, 0xC4)));
}

#[test]
fn test_slide_background_with_text_content() {
    // Slide with both background and text shapes — both should be present
    let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let text_box = make_text_box(100000, 100000, 5000000, 500000, "Hello");
    let slide = make_slide_xml_with_bg(bg_xml, &[text_box]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.background_color, Some(Color::new(0, 0, 255)));
    assert_eq!(page.elements.len(), 1);
}

#[test]
fn test_slide_inherits_master_background() {
    // Slide has no background, but its master does → should inherit
    let slide_xml = make_empty_slide_xml();
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

    // Build PPTX with slide → layout → master chain
    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    // Should inherit master's green background
    assert_eq!(page.background_color, Some(Color::new(0, 255, 0)));
}

/// Create a slide layout XML with the given shape elements.
fn make_layout_xml(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sldLayout>");
    xml
}

/// Create a slide master XML with the given shape elements.
fn make_master_xml(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sldMaster>");
    xml
}

// ── US-025: Slide master and layout inheritance tests ────────────────

#[test]
fn test_master_shape_appears_on_slide() {
    // Master has a rectangle shape → it should appear on the slide
    let slide_xml = make_empty_slide_xml();
    let layout_xml = make_layout_xml(&[]);
    let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Master Logo");
    let master_xml = make_master_xml(&[master_shape]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    // Master element should be present
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Master Logo");
}

#[test]
fn test_layout_shape_appears_on_slide() {
    // Layout has a text box → it should appear on the slide
    let slide_xml = make_empty_slide_xml();
    let layout_shape = make_text_box(100_000, 100_000, 3_000_000, 500_000, "Layout Title");
    let layout_xml = make_layout_xml(&[layout_shape]);
    let master_xml = make_master_xml(&[]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Layout Title");
}

#[test]
fn test_inheritance_element_ordering() {
    // Master, layout, and slide all have elements → order: master, layout, slide
    let slide_shape = make_text_box(0, 0, 1_000_000, 500_000, "Slide Content");
    let slide_xml = make_slide_xml(&[slide_shape]);
    let layout_shape = make_text_box(0, 0, 1_000_000, 500_000, "Layout Content");
    let layout_xml = make_layout_xml(&[layout_shape]);
    let master_shape = make_text_box(0, 0, 1_000_000, 500_000, "Master Content");
    let master_xml = make_master_xml(&[master_shape]);

    let data = build_test_pptx_with_layout_master(
        SLIDE_CX,
        SLIDE_CY,
        &slide_xml,
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 3);

    // Master element is first (behind)
    let master_blocks = text_box_blocks(&page.elements[0]);
    let master_para = match &master_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(master_para.runs[0].text, "Master Content");

    // Layout element is second
    let layout_blocks = text_box_blocks(&page.elements[1]);
    let layout_para = match &layout_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(layout_para.runs[0].text, "Layout Content");

    // Slide element is last (on top)
    let slide_blocks = text_box_blocks(&page.elements[2]);
    let slide_para = match &slide_blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(slide_para.runs[0].text, "Slide Content");
}

#[test]
fn test_master_elements_appear_on_all_slides() {
    // Build a PPTX with 2 slides and a master shape → both slides should have it
    let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Company Logo");
    let master_xml = make_master_xml(&[master_shape]);
    let layout_xml = make_layout_xml(&[]);

    let slide1_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 1");
    let slide1_xml = make_slide_xml(&[slide1_shape]);
    let slide2_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 2");
    let slide2_xml = make_slide_xml(&[slide2_shape]);

    // Build PPTX with 2 slides, both pointing to same layout/master
    let data = build_test_pptx_with_layout_master_multi_slide(
        SLIDE_CX,
        SLIDE_CY,
        &[slide1_xml, slide2_xml],
        &layout_xml,
        &master_xml,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2);

    // Both slides should have master element + their own element
    for (i, page) in doc.pages.iter().enumerate() {
        let fixed_page = match page {
            Page::Fixed(p) => p,
            _ => panic!("Expected FixedPage"),
        };
        assert_eq!(
            fixed_page.elements.len(),
            2,
            "Slide {} should have 2 elements (master + slide)",
            i + 1
        );

        // First element is the master shape
        let master_blocks = text_box_blocks(&fixed_page.elements[0]);
        let master_para = match &master_blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(master_para.runs[0].text, "Company Logo");
    }
}

#[test]
fn test_slide_without_layout_master_has_only_slide_elements() {
    // Standard PPTX without layout/master .rels → only slide elements
    let shape = make_text_box(0, 0, 1_000_000, 500_000, "Just Slide");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Just Slide");
}

#[test]
fn test_slide_inherits_layout_background_over_master() {
    // Layout has a background, master has a different one → layout wins
    let slide_xml = make_empty_slide_xml();
    let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
    let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

    let data =
        build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    // Should inherit layout's magenta background (not master's green)
    assert_eq!(page.background_color, Some(Color::new(255, 0, 255)));
}

// ── Table test helpers ──────────────────────────────────────────────

/// Create a graphicFrame XML containing a table.
/// `x`, `y`, `cx`, `cy` are in EMU.
fn make_table_graphic_frame(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    col_widths_emu: &[i64],
    rows_xml: &str,
) -> String {
    let mut grid = String::new();
    for w in col_widths_emu {
        grid.push_str(&format!(r#"<a:gridCol w="{w}"/>"#));
    }
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{rows_xml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

/// Create a simple table row with text-only cells.
fn make_table_row(cells: &[&str]) -> String {
    let mut xml = String::from(r#"<a:tr h="370840">"#);
    for text in cells {
        xml.push_str(&format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#
            ));
    }
    xml.push_str("</a:tr>");
    xml
}

/// Helper: get the Table from a FixedElement.
fn table_element(elem: &FixedElement) -> &Table {
    match &elem.kind {
        FixedElementKind::Table(t) => t,
        _ => panic!("Expected Table, got {:?}", elem.kind),
    }
}

// ── Table tests ─────────────────────────────────────────────────────

#[test]
fn test_slide_with_basic_table() {
    // A slide with a 2×2 table
    let rows = format!(
        "{}{}",
        make_table_row(&["A1", "B1"]),
        make_table_row(&["A2", "B2"]),
    );
    let table_frame = make_table_graphic_frame(
        914400,              // x = 72pt
        914400,              // y = 72pt
        3657600,             // cx = 288pt
        1828800,             // cy = 144pt
        &[1828800, 1828800], // 2 columns, 144pt each
        &rows,
    );
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let elem = &page.elements[0];
    assert!((elem.x - 72.0).abs() < 0.1);
    assert!((elem.y - 72.0).abs() < 0.1);

    let table = table_element(elem);
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 144.0).abs() < 0.1);

    // Check cell text
    let cell_00 = &table.rows[0].cells[0];
    assert_eq!(cell_00.content.len(), 1);
    if let Block::Paragraph(p) = &cell_00.content[0] {
        assert_eq!(p.runs[0].text, "A1");
    } else {
        panic!("Expected paragraph in cell");
    }

    let cell_11 = &table.rows[1].cells[1];
    if let Block::Paragraph(p) = &cell_11.content[0] {
        assert_eq!(p.runs[0].text, "B2");
    } else {
        panic!("Expected paragraph in cell");
    }
}

#[test]
fn test_slide_table_scales_geometry_to_graphic_frame_extent() {
    let rows_xml = format!(
        "{}{}",
        make_table_row(&["A1", "B1"]),
        make_table_row(&["A2", "B2"]),
    );
    let table_frame = make_table_graphic_frame(
        914400,
        914400,
        3_657_600,
        1_483_360,
        &[914_400, 914_400],
        &rows_xml,
    );
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];
    let table = table_element(elem);

    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 144.0).abs() < 0.1);
    assert!((table.column_widths.iter().sum::<f64>() - elem.width).abs() < 0.1);

    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].height, Some(58.4));
    assert_eq!(table.rows[1].height, Some(58.4));
    assert!(
        (table
            .rows
            .iter()
            .map(|row| row.height.unwrap_or(0.0))
            .sum::<f64>()
            - elem.height)
            .abs()
            < 0.1
    );
}

#[test]
fn test_slide_table_reads_column_widths_from_gridcol_with_extensions() {
    let rows_xml = make_table_row(&["A1", "B1"]);
    let table_frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="370840"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="1"/></a:ext></a:extLst></a:gridCol><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="2"/></a:ext></a:extLst></a:gridCol></a:tblGrid>"#.to_string()
            + &rows_xml
            + r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 72.0).abs() < 0.1);
    assert!((table.column_widths[1] - 72.0).abs() < 0.1);
}

#[test]
fn test_slide_table_cell_anchor_maps_to_vertical_alignment() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Centered</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bottom</a:t></a:r></a:p></a:txBody><a:tcPr anchor="b"/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame =
        make_table_graphic_frame(0, 0, 1_828_800, 370_840, &[914_400, 914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.rows[0].cells[0].vertical_align,
        Some(crate::ir::CellVerticalAlign::Center)
    );
    assert_eq!(
        table.rows[0].cells[1].vertical_align,
        Some(crate::ir::CellVerticalAlign::Bottom)
    );
}

#[test]
fn test_slide_table_cell_margins_map_to_padding() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Padded</a:t></a:r></a:p></a:txBody><a:tcPr marL="76200" marR="76200" marT="38100" marB="38100"/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.rows[0].cells[0].padding,
        Some(crate::ir::Insets {
            top: 3.0,
            right: 6.0,
            bottom: 3.0,
            left: 6.0,
        })
    );
}

#[test]
fn test_slide_table_uses_powerpoint_default_cell_padding() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>DefaultPadding</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.default_cell_padding,
        Some(crate::ir::Insets {
            top: 3.6,
            right: 7.2,
            bottom: 3.6,
            left: 7.2,
        })
    );
    assert_eq!(table.rows[0].cells[0].padding, None);
    assert!(table.use_content_driven_row_heights);
}

#[test]
fn test_slide_table_coalesces_adjacent_runs_with_same_style() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t>YOLOv8n + </a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100" err="1"><a:latin typeface="Arial"/></a:rPr><a:t>topk filtering on gpu(</a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>K</a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t> = 100)</a:t></a:r>"#,
        r#"</a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    let paragraph = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph, got {other:?}"),
    };

    assert_eq!(paragraph.runs.len(), 3);
    assert_eq!(paragraph.runs[0].text, "YOLOv8n + topk filtering on gpu(");
    assert_eq!(paragraph.runs[1].text, "K");
    assert_eq!(paragraph.runs[2].text, "\u{00A0}= 100)");
    assert_eq!(paragraph.runs[1].style.italic, Some(true));
}

#[test]
fn test_slide_table_cell_bulleted_paragraphs_group_into_list() {
    let rows_xml = concat!(
        r#"<a:tr h="740000">"#,
        r#"<a:tc><a:txBody><a:bodyPr/>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Second bullet</a:t></a:r></a:p>"#,
        r#"</a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 740_000, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    assert_eq!(table.rows[0].cells[0].content.len(), 1);

    let list = match &table.rows[0].cells[0].content[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Unordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
}

#[test]
fn test_slide_table_with_merged_cells() {
    // Table with gridSpan (horizontal merge) and vMerge (vertical merge)
    let mut rows_xml = String::new();
    // Row 0: cell spanning 2 columns
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");
    // Row 1: two normal cells
    rows_xml.push_str(&make_table_row(&["C1", "C2"]));

    let table_frame =
        make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    // Row 0: merged cell should have col_span=2
    assert_eq!(table.rows[0].cells.len(), 2);
    assert_eq!(table.rows[0].cells[0].col_span, 2);
    // The hMerge cell should have col_span=0 (covered by merge)
    assert_eq!(table.rows[0].cells[1].col_span, 0);

    // Row 1: normal cells
    assert_eq!(table.rows[1].cells[0].col_span, 1);
    assert_eq!(table.rows[1].cells[1].col_span, 1);
}

#[test]
fn test_slide_table_with_vertical_merge() {
    // Table with rowSpan (vertical merge)
    let mut rows_xml = String::new();
    // Row 0: first cell starts a rowSpan of 2
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>VMerged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");
    // Row 1: first cell is continuation of vMerge
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc vMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame =
        make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    // Row 0: first cell rowSpan=2
    assert_eq!(table.rows[0].cells[0].row_span, 2);
    // Row 1: first cell vMerge continuation (row_span=0)
    assert_eq!(table.rows[1].cells[0].row_span, 0);
}

#[test]
fn test_slide_table_with_formatted_text() {
    // Table cell with bold, colored text
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" b="1" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Bold Red</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    let cell = &table.rows[0].cells[0];
    if let Block::Paragraph(p) = &cell.content[0] {
        assert_eq!(p.runs[0].text, "Bold Red");
        assert_eq!(p.runs[0].style.bold, Some(true));
        assert_eq!(p.runs[0].style.font_size, Some(18.0));
        assert_eq!(p.runs[0].style.color, Some(Color::new(255, 0, 0)));
    } else {
        panic!("Expected paragraph in cell");
    }
}

#[test]
fn test_slide_table_with_cell_background() {
    // Table cell with background fill
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Filled</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    let cell = &table.rows[0].cells[0];
    assert_eq!(cell.background, Some(Color::new(0, 255, 0)));
}

#[test]
fn test_slide_table_with_cell_borders() {
    // Table cell with border specification
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bordered</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB></a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    let cell = &table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");
    assert!(border.left.is_some());
    assert!(border.right.is_some());
    assert!(border.top.is_some());
    assert!(border.bottom.is_some());
    let left = border.left.as_ref().unwrap();
    assert!((left.width - 1.0).abs() < 0.1); // 12700 EMU = 1pt
    assert_eq!(left.color, Color::new(0, 0, 0));
}

#[test]
fn test_slide_table_cell_border_dash_styles() {
    // Table cell with dashed top and dotted bottom borders
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Dashed</a:t></a:r></a:p></a:txBody><a:tcPr>"#);
    rows_xml.push_str(r#"<a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:lnT>"#);
    rows_xml.push_str(r#"<a:lnB w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dot"/></a:lnB>"#);
    rows_xml.push_str(r#"</a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    let cell = &table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");

    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Dashed, "Top should be dashed");

    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert_eq!(
        bottom.style,
        BorderLineStyle::Dotted,
        "Bottom should be dotted"
    );
}

#[test]
fn test_shape_outline_dash_style() {
    // Shape with dashed outline
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape_elem = &page.elements[0];
    if let FixedElementKind::Shape(ref s) = shape_elem.kind {
        let stroke = s.stroke.as_ref().expect("Expected stroke");
        assert_eq!(
            stroke.style,
            BorderLineStyle::Dashed,
            "Shape stroke should be dashed"
        );
    } else {
        panic!("Expected Shape element");
    }
}

#[test]
fn test_slide_table_coexists_with_shapes() {
    // A slide with both a text box and a table
    let text_box = make_text_box(0, 0, 914400, 457200, "Header");
    let rows = make_table_row(&["Cell"]);
    let table_frame = make_table_graphic_frame(0, 914400, 914400, 370840, &[914400], &rows);
    let slide = make_slide_xml(&[text_box, table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2);

    // First element: TextBox
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    // Second element: Table
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Table(_)));
}

// ----- US-029: Slide selection tests -----

#[test]
fn test_slide_filter_single_slide() {
    use crate::config::SlideRange;
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(2, 2)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 1, "Should only include slide 2");
    // Verify slide 2 content
    let page = first_fixed_page(&doc);
    let text = match &page.elements[0].kind {
        FixedElementKind::TextBox(text_box) => match &text_box.content[0] {
            Block::Paragraph(p) => p.runs[0].text.clone(),
            _ => panic!("Expected Paragraph"),
        },
        _ => panic!("Expected TextBox"),
    };
    assert_eq!(text, "Slide 2");
}

#[test]
fn test_slide_filter_range() {
    use crate::config::SlideRange;
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
    let slide4 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 4")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3, slide4]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(2, 3)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 2, "Should include slides 2 and 3");
}

#[test]
fn test_slide_filter_none_includes_all() {
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: None,
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 2, "None should include all slides");
}

#[test]
fn test_slide_filter_range_beyond_total() {
    use crate::config::SlideRange;
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

    let parser = PptxParser;
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(5, 10)),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(
        doc.pages.len(),
        0,
        "Range beyond total slides should produce empty document"
    );
}

// ── Group shape helpers ─────────────────────────────────────────────

/// Create a group shape XML with a coordinate transform and child shapes.
#[allow(clippy::too_many_arguments)]
fn make_group_shape(
    off_x: i64,
    off_y: i64,
    ext_cx: i64,
    ext_cy: i64,
    ch_off_x: i64,
    ch_off_y: i64,
    ch_ext_cx: i64,
    ch_ext_cy: i64,
    children: &[String],
) -> String {
    let mut xml = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/><a:chOff x="{ch_off_x}" y="{ch_off_y}"/><a:chExt cx="{ch_ext_cx}" cy="{ch_ext_cy}"/></a:xfrm></p:grpSpPr>"#
    );
    for child in children {
        xml.push_str(child);
    }
    xml.push_str("</p:grpSp>");
    xml
}

/// Create a rectangle shape XML (no text body) with a fill color.
fn make_shape_rect(x: i64, y: i64, cx: i64, cy: i64, fill_hex: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill></p:spPr></p:sp>"#
    )
}

// ── Group shape tests ───────────────────────────────────────────────

#[test]
fn test_group_shape_two_text_boxes() {
    // Group at (1000000, 500000) with 1:1 mapping (ext == chExt)
    let child_a = make_text_box(0, 0, 2_000_000, 1_000_000, "Shape A");
    let child_b = make_text_box(2_000_000, 1_000_000, 2_000_000, 1_000_000, "Shape B");
    let group = make_group_shape(
        1_000_000,
        500_000, // off
        4_000_000,
        2_000_000, // ext
        0,
        0, // chOff
        4_000_000,
        2_000_000, // chExt
        &[child_a, child_b],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 elements from group");

    // Shape A: (1000000, 500000) in EMU → ~78.74pt, ~39.37pt
    let a = &page.elements[0];
    assert!(
        (a.x - emu_to_pt(1_000_000)).abs() < 0.1,
        "Shape A x: got {}, expected {}",
        a.x,
        emu_to_pt(1_000_000)
    );
    assert!(
        (a.y - emu_to_pt(500_000)).abs() < 0.1,
        "Shape A y: got {}, expected {}",
        a.y,
        emu_to_pt(500_000)
    );

    // Shape B: (1000000+2000000, 500000+1000000) = (3000000, 1500000) EMU
    let b = &page.elements[1];
    assert!(
        (b.x - emu_to_pt(3_000_000)).abs() < 0.1,
        "Shape B x: got {}, expected {}",
        b.x,
        emu_to_pt(3_000_000)
    );
    assert!(
        (b.y - emu_to_pt(1_500_000)).abs() < 0.1,
        "Shape B y: got {}, expected {}",
        b.y,
        emu_to_pt(1_500_000)
    );

    // Verify text content
    let blocks_a = text_box_blocks(a);
    let para_a = match &blocks_a[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para_a.runs[0].text, "Shape A");
}

#[test]
fn test_group_shape_with_scaling() {
    // Group: ext is half of chExt → children scaled down by 0.5
    let child = make_text_box(0, 0, 4_000_000, 2_000_000, "Scaled");
    let group = make_group_shape(
        0,
        0, // off
        2_000_000,
        1_000_000, // ext (half)
        0,
        0, // chOff
        4_000_000,
        2_000_000, // chExt (full)
        &[child],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let elem = &page.elements[0];
    // Width: 4000000 * 0.5 = 2000000 EMU → ~157.48pt
    let expected_w = emu_to_pt(2_000_000);
    assert!(
        (elem.width - expected_w).abs() < 0.1,
        "Scaled width: got {}, expected {}",
        elem.width,
        expected_w
    );
    let expected_h = emu_to_pt(1_000_000);
    assert!(
        (elem.height - expected_h).abs() < 0.1,
        "Scaled height: got {}, expected {}",
        elem.height,
        expected_h
    );
}

#[test]
fn test_nested_group_shapes() {
    // Inner group at (0, 0) with 1:1 mapping containing a text box
    let inner_child = make_text_box(0, 0, 1_000_000, 1_000_000, "Nested");
    let inner_group = make_group_shape(
        0,
        0, // off
        2_000_000,
        2_000_000, // ext
        0,
        0, // chOff
        2_000_000,
        2_000_000, // chExt
        &[inner_child],
    );
    // Outer group at (1000000, 1000000) with 1:1 mapping
    let outer_group = make_group_shape(
        1_000_000,
        1_000_000, // off
        4_000_000,
        4_000_000, // ext
        0,
        0, // chOff
        4_000_000,
        4_000_000, // chExt
        &[inner_group],
    );
    let slide = make_slide_xml(&[outer_group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        1,
        "Expected 1 element (nested text box)"
    );

    // The inner child at (0,0) → inner group maps to (0,0) → outer group maps to (1000000, 1000000)
    let elem = &page.elements[0];
    assert!(
        (elem.x - emu_to_pt(1_000_000)).abs() < 0.1,
        "Nested x: got {}, expected {}",
        elem.x,
        emu_to_pt(1_000_000)
    );
    assert!(
        (elem.y - emu_to_pt(1_000_000)).abs() < 0.1,
        "Nested y: got {}, expected {}",
        elem.y,
        emu_to_pt(1_000_000)
    );
    assert_eq!(elem.width, emu_to_pt(1_000_000));
    assert_eq!(elem.height, emu_to_pt(1_000_000));
}

#[test]
fn test_group_shape_mixed_element_types() {
    // Group with a text box and a rectangle shape
    let text = make_text_box(0, 0, 2_000_000, 1_000_000, "Text");
    let rect = make_shape_rect(2_000_000, 0, 2_000_000, 1_000_000, "FF0000");
    let group = make_group_shape(
        0,
        0, // off
        4_000_000,
        2_000_000, // ext
        0,
        0, // chOff
        4_000_000,
        2_000_000, // chExt
        &[text, rect],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected TextBox + Shape");

    // First element: TextBox
    assert!(
        matches!(&page.elements[0].kind, FixedElementKind::TextBox(_)),
        "First element should be TextBox"
    );
    // Second element: Shape
    assert!(
        matches!(&page.elements[1].kind, FixedElementKind::Shape(_)),
        "Second element should be Shape"
    );

    // Verify shape position: (2000000, 0) in child space → (2000000, 0) in slide space
    let shape_elem = &page.elements[1];
    assert!(
        (shape_elem.x - emu_to_pt(2_000_000)).abs() < 0.1,
        "Shape x: got {}, expected {}",
        shape_elem.x,
        emu_to_pt(2_000_000)
    );
}

#[test]
fn test_group_shape_with_nonzero_child_offset() {
    // Group where chOff != (0,0) — children positioned relative to offset
    let child = make_text_box(1_000_000, 1_000_000, 2_000_000, 1_000_000, "Offset");
    let group = make_group_shape(
        500_000,
        500_000, // off (group position on slide)
        4_000_000,
        2_000_000, // ext
        1_000_000,
        1_000_000, // chOff (child space origin)
        4_000_000,
        2_000_000, // chExt
        &[child],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    // child_x=1000000, chOff_x=1000000 → (1000000-1000000)*1.0 + 500000 = 500000
    let elem = &page.elements[0];
    assert!(
        (elem.x - emu_to_pt(500_000)).abs() < 0.1,
        "Offset x: got {}, expected {}",
        elem.x,
        emu_to_pt(500_000)
    );
    assert!(
        (elem.y - emu_to_pt(500_000)).abs() < 0.1,
        "Offset y: got {}, expected {}",
        elem.y,
        emu_to_pt(500_000)
    );
}

// ── Shape style (rotation, transparency) test helpers ────────────────

/// Create a shape XML with optional rotation and fill alpha.
/// `rot` is in 60000ths of a degree (e.g. 5400000 = 90°).
/// `alpha_thousandths` is in 1000ths of percent (e.g. 50000 = 50%).
#[allow(clippy::too_many_arguments)]
fn make_styled_shape(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    fill_hex: Option<&str>,
    rot: Option<i64>,
    alpha_thousandths: Option<i64>,
) -> String {
    let rot_attr = rot.map(|r| format!(r#" rot="{r}""#)).unwrap_or_default();

    let fill_xml = match (fill_hex, alpha_thousandths) {
        (Some(h), Some(a)) => format!(
            r#"<a:solidFill><a:srgbClr val="{h}"><a:alpha val="{a}"/></a:srgbClr></a:solidFill>"#
        ),
        (Some(h), None) => {
            format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#)
        }
        _ => String::new(),
    };

    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}</p:spPr></p:sp>"#
    )
}

// ── Shape style tests (US-034) ──────────────────────────────────────

#[test]
fn test_shape_rotation() {
    // 90° rotation = 5400000 (60000ths of a degree)
    let shape = make_styled_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "rect",
        Some("FF0000"),
        Some(5_400_000),
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(s.rotation_deg.is_some(), "Expected rotation_deg to be set");
    assert!(
        (s.rotation_deg.unwrap() - 90.0).abs() < 0.01,
        "Expected 90°, got {}",
        s.rotation_deg.unwrap()
    );
}

#[test]
fn test_shape_transparency() {
    // 50% opacity = alpha val 50000 (in 1000ths of percent)
    let shape = make_styled_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "rect",
        Some("00FF00"),
        None,
        Some(50_000),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(s.opacity.is_some(), "Expected opacity to be set");
    assert!(
        (s.opacity.unwrap() - 0.5).abs() < 0.01,
        "Expected 0.5 opacity, got {}",
        s.opacity.unwrap()
    );
}

#[test]
fn test_shape_rotation_and_transparency() {
    // 45° rotation (2700000) + 75% opacity (75000)
    let shape = make_styled_shape(
        1_000_000,
        500_000,
        3_000_000,
        2_000_000,
        "ellipse",
        Some("0000FF"),
        Some(2_700_000),
        Some(75_000),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(
        (s.rotation_deg.unwrap() - 45.0).abs() < 0.01,
        "Expected 45°, got {}",
        s.rotation_deg.unwrap()
    );
    assert!(
        (s.opacity.unwrap() - 0.75).abs() < 0.01,
        "Expected 0.75 opacity, got {}",
        s.opacity.unwrap()
    );
    assert!(matches!(s.kind, ShapeKind::Ellipse));
}

#[path = "pptx_smartart_tests.rs"]
mod smartart_tests;

#[path = "pptx_chart_tests.rs"]
mod chart_tests;

// ── Gradient background tests (US-050) ──────────────────────────────

#[test]
fn test_gradient_background_two_stops() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    // Should have a gradient background
    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient background");
    assert_eq!(gradient.stops.len(), 2);

    // First stop: red at 0%
    assert!((gradient.stops[0].position - 0.0).abs() < 0.001);
    assert_eq!(gradient.stops[0].color, Color::new(255, 0, 0));

    // Second stop: blue at 100%
    assert!((gradient.stops[1].position - 1.0).abs() < 0.001);
    assert_eq!(gradient.stops[1].color, Color::new(0, 0, 255));

    // Angle: 5400000 / 60000 = 90 degrees
    assert!((gradient.angle - 90.0).abs() < 0.001);

    // Fallback solid color should be first stop color
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_gradient_background_three_stops() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert_eq!(gradient.stops.len(), 3);
    assert!((gradient.stops[1].position - 0.5).abs() < 0.001);
    assert_eq!(gradient.stops[1].color, Color::new(0, 255, 0));
    assert!((gradient.angle - 0.0).abs() < 0.001);
}

#[test]
fn test_gradient_background_with_scheme_colors() {
    // Use theme colors in gradient stops
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);

    // Build with theme
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert_eq!(gradient.stops.len(), 2);
    // angle = 2700000 / 60000 = 45 degrees
    assert!((gradient.angle - 45.0).abs() < 0.001);
}

#[test]
fn test_solid_background_no_gradient() {
    // Solid fill background should NOT produce a gradient
    let bg_xml =
        r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    assert!(
        page.background_gradient.is_none(),
        "Solid fill should not produce gradient"
    );
    assert_eq!(page.background_color, Some(Color::new(255, 204, 0)));
}

#[test]
fn test_gradient_shape_fill() {
    // Shape with gradient fill
    let shape_xml =
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></p:spPr></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    assert_eq!(page.elements.len(), 1);
    let shape = get_shape(&page.elements[0]);

    // Should have gradient fill
    let gf = shape
        .gradient_fill
        .as_ref()
        .expect("Expected gradient fill on shape");
    assert_eq!(gf.stops.len(), 2);
    assert_eq!(gf.stops[0].color, Color::new(255, 0, 0));
    assert_eq!(gf.stops[1].color, Color::new(0, 255, 0));
    assert!((gf.angle - 90.0).abs() < 0.001);

    // Solid fill fallback should be first stop color
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_shape_solid_fill_no_gradient() {
    // Shape with only solid fill — gradient_fill should be None
    let shape_xml =
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    assert!(
        shape.gradient_fill.is_none(),
        "Solid fill shape should have no gradient"
    );
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_gradient_background_no_angle() {
    // Gradient with no <a:lin> element → angle defaults to 0
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert!(
        (gradient.angle - 0.0).abs() < 0.001,
        "Default angle should be 0"
    );
}

// ── Shadow / effects tests ─────────────────────────────────────────

#[test]
fn test_shape_outer_shadow_parsed() {
    // Shape with <a:effectLst><a:outerShdw> inside <p:spPr>
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    let shadow = shape.shadow.as_ref().expect("Expected shadow");

    // blurRad=50800 EMU → 50800/12700 = 4.0 pt
    assert!(
        (shadow.blur_radius - 4.0).abs() < 0.01,
        "Expected blur_radius ~4.0, got {}",
        shadow.blur_radius
    );
    // dist=38100 EMU → 38100/12700 = 3.0 pt
    assert!(
        (shadow.distance - 3.0).abs() < 0.01,
        "Expected distance ~3.0, got {}",
        shadow.distance
    );
    // dir=2700000 → 2700000/60000 = 45.0 degrees
    assert!(
        (shadow.direction - 45.0).abs() < 0.01,
        "Expected direction ~45.0, got {}",
        shadow.direction
    );
    // color = black
    assert_eq!(shadow.color, Color::new(0, 0, 0));
    // alpha val=50000 → 50000/100000 = 0.5
    assert!(
        (shadow.opacity - 0.5).abs() < 0.01,
        "Expected opacity ~0.5, got {}",
        shadow.opacity
    );
}

#[test]
fn test_shape_no_effects_no_shadow() {
    // Shape with no <a:effectLst> → shadow should be None
    let shape_xml = make_shape(
        100_000,
        200_000,
        500_000,
        300_000,
        "rect",
        Some("00FF00"),
        None,
        None,
    );
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    assert!(
        shape.shadow.is_none(),
        "Shape without effectLst should have no shadow"
    );
}

#[test]
fn test_shape_shadow_default_opacity() {
    // Shadow with no <a:alpha> element → opacity defaults to 1.0
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="25400" dist="12700" dir="5400000"><a:srgbClr val="333333"/></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    let shadow = shape.shadow.as_ref().expect("Expected shadow");

    // blurRad=25400 EMU → 2.0 pt
    assert!(
        (shadow.blur_radius - 2.0).abs() < 0.01,
        "Expected blur ~2.0, got {}",
        shadow.blur_radius
    );
    // dist=12700 EMU → 1.0 pt
    assert!(
        (shadow.distance - 1.0).abs() < 0.01,
        "Expected dist ~1.0, got {}",
        shadow.distance
    );
    // dir=5400000 → 90.0 degrees
    assert!(
        (shadow.direction - 90.0).abs() < 0.01,
        "Expected dir ~90.0, got {}",
        shadow.direction
    );
    // color = #333333
    assert_eq!(shadow.color, Color::new(0x33, 0x33, 0x33));
    // No alpha element → defaults to 1.0
    assert!(
        (shadow.opacity - 1.0).abs() < 0.01,
        "Expected opacity ~1.0 (default), got {}",
        shadow.opacity
    );
}

#[path = "pptx_metadata_tests.rs"]
mod metadata_tests;

#[path = "pptx_preset_shape_tests.rs"]
mod preset_shape_tests;
