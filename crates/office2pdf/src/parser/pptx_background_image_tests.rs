use super::*;
use std::io::Write;
use zip::write::FileOptions;

// ── Helpers ──────────────────────────────────────────────────────────

const BG_IMAGE_RID: &str = "rId9";

fn make_picture_fill_bg() -> String {
    format!(
        r#"<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="0"><a:blip r:embed="{BG_IMAGE_RID}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>"#
    )
}

fn make_slide_xml(bg_xml: &str, shapes: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>{bg_xml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
    )
}

fn make_layout_xml(bg_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>{bg_xml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#
    )
}

fn make_master_xml(bg_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>{bg_xml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#
    )
}

/// Build a PPTX whose slide/layout/master can each carry a `<p:bg>` and where
/// EVERY layer's rels expose the same BMP image relationship, so any layer's
/// picture-fill background is resolvable.
fn build_test_pptx_with_bg_layers(slide_xml: &str, layout_xml: &str, master_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    let ct = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bmp" ContentType="image/bmp"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#;
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    let pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{SLIDE_CX}" cy="{SLIDE_CY}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    );
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#,
    )
    .unwrap();

    let image_rel = format!(
        r#"<Relationship Id="{BG_IMAGE_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.bmp"/>"#
    );

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide_xml.as_bytes()).unwrap();
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>{image_rel}</Relationships>"#).as_bytes())
        .unwrap();

    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(layout_xml.as_bytes()).unwrap();
    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
        .unwrap();
    zip.write_all(format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>{image_rel}</Relationships>"#).as_bytes())
        .unwrap();

    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(master_xml.as_bytes()).unwrap();
    zip.start_file("ppt/slideMasters/_rels/slideMaster1.xml.rels", opts)
        .unwrap();
    zip.write_all(format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_rel}</Relationships>"#).as_bytes())
        .unwrap();

    zip.start_file("ppt/media/image1.bmp", opts).unwrap();
    zip.write_all(&image_tests::make_test_bmp()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn parse_first_page(data: &[u8]) -> FixedPage {
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    match &doc.pages[0] {
        Page::Fixed(page) => page.clone(),
        _ => panic!("Expected FixedPage"),
    }
}

fn assert_full_page_image(element: &FixedElement, page: &FixedPage) {
    assert!(
        matches!(element.kind, FixedElementKind::Image(_)),
        "expected background image element, got {:?}",
        element.kind
    );
    assert!((element.x - 0.0).abs() < 0.01 && (element.y - 0.0).abs() < 0.01);
    assert!(
        (element.width - page.size.width).abs() < 0.01
            && (element.height - page.size.height).abs() < 0.01,
        "background image should cover the page: {}x{} vs page {}x{}",
        element.width,
        element.height,
        page.size.width,
        page.size.height
    );
}

// ── Picture-fill backgrounds ─────────────────────────────────────────

#[test]
fn test_master_picture_fill_background_renders_behind_content() {
    let text_box = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="5000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Content</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let data = build_test_pptx_with_bg_layers(
        &make_slide_xml("", text_box),
        &make_layout_xml(""),
        &make_master_xml(&make_picture_fill_bg()),
    );

    let page = parse_first_page(&data);
    assert!(
        page.elements.len() >= 2,
        "expected background image + content, got {} elements",
        page.elements.len()
    );
    assert_full_page_image(&page.elements[0], &page);
}

#[test]
fn test_layout_picture_fill_background_is_inherited() {
    let data = build_test_pptx_with_bg_layers(
        &make_slide_xml("", ""),
        &make_layout_xml(&make_picture_fill_bg()),
        &make_master_xml(""),
    );

    let page = parse_first_page(&data);
    assert_eq!(page.elements.len(), 1);
    assert_full_page_image(&page.elements[0], &page);
}

#[test]
fn test_slide_picture_fill_background_renders() {
    let data = build_test_pptx_with_bg_layers(
        &make_slide_xml(&make_picture_fill_bg(), ""),
        &make_layout_xml(""),
        &make_master_xml(""),
    );

    let page = parse_first_page(&data);
    assert_eq!(page.elements.len(), 1);
    assert_full_page_image(&page.elements[0], &page);
}

#[test]
fn test_slide_solid_background_wins_over_master_picture_fill() {
    let solid_bg = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
    let data = build_test_pptx_with_bg_layers(
        &make_slide_xml(solid_bg, ""),
        &make_layout_xml(""),
        &make_master_xml(&make_picture_fill_bg()),
    );

    let page = parse_first_page(&data);
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
    assert!(
        page.elements.is_empty(),
        "no background image element expected when the slide overrides with a solid fill"
    );
}
