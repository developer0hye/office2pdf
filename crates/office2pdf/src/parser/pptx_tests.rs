use super::*;
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
    let expected_w = emu_to_pt(cx);
    let expected_h = emu_to_pt(cy);
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

#[path = "pptx_text_box_tests.rs"]
mod text_box_tests;

#[path = "pptx_text_box_semantic_tests.rs"]
mod text_box_semantic_tests;

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

#[path = "pptx_theme_tests.rs"]
mod theme_tests;
use self::theme_tests::{
    build_test_pptx_with_layout_master, build_test_pptx_with_layout_master_multi_slide,
    build_test_pptx_with_theme, make_theme_xml, standard_theme_colors,
};

#[path = "pptx_table_tests.rs"]
mod table_tests;

#[path = "pptx_slide_feature_tests.rs"]
mod slide_feature_tests;

#[path = "pptx_group_shape_tests.rs"]
mod group_shape_tests;

#[path = "pptx_smartart_tests.rs"]
mod smartart_tests;

#[path = "pptx_chart_tests.rs"]
mod chart_tests;

#[path = "pptx_image_tests.rs"]
mod image_tests;

#[path = "pptx_shape_style_tests.rs"]
mod shape_style_tests;

#[path = "pptx_metadata_tests.rs"]
mod metadata_tests;

#[path = "pptx_preset_shape_tests.rs"]
mod preset_shape_tests;

#[path = "pptx_connector_tests.rs"]
mod connector_tests;
