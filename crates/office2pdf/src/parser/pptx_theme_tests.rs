use super::*;

// ── Theme test helpers ────────────────────────────────────────────

pub(super) fn make_theme_xml(
    colors: &[(&str, &str)],
    major_font: &str,
    minor_font: &str,
) -> String {
    let mut color_xml = String::new();
    for (name, hex) in colors {
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

pub(super) fn standard_theme_colors() -> Vec<(&'static str, &'static str)> {
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

pub(super) fn build_test_pptx_with_theme(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xmls: &[String],
    theme_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

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

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

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

    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
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

    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(theme_xml.as_bytes()).unwrap();

    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

pub(super) fn build_test_pptx_with_layout_master(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xml: &str,
    layout_xml: &str,
    master_xml: &str,
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

    let pres_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#;
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

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

    zip.finish().unwrap().into_inner()
}

pub(super) fn build_test_pptx_with_theme_layout_master(
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

    zip.finish().unwrap().into_inner()
}

pub(super) fn build_test_pptx_with_layout_master_multi_slide(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xmls: &[String],
    layout_xml: &str,
    master_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

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

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

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

    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        let slide_num = i + 1;
        zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();
    }

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

    zip.finish().unwrap().into_inner()
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
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:ln></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let stroke = shape.stroke.as_ref().expect("Expected stroke");
    assert_eq!(stroke.color, Color::new(0, 0, 0));
}

#[test]
fn test_scheme_color_in_text_run() {
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
fn test_hyperlink_run_inherits_theme_color_underline_and_keeps_boundary() {
    let runs_xml = concat!(
        r#"<a:r><a:rPr lang="en-US"/><a:t>Plain text, then </a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>Click here</a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US"/><a:t> after</a:t></a:r>"#,
    );
    let shape = make_formatted_text_box(0, 0, 4_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&[("hlink", "778BA2")], "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 3);
    assert_eq!(para.runs[0].text, "Plain text, then ");
    assert_eq!(para.runs[1].text, "Click here");
    assert_eq!(para.runs[1].style.color, Some(Color::new(0x77, 0x8B, 0xA2)));
    assert_eq!(para.runs[1].style.underline, Some(true));
    assert_eq!(para.runs[2].text, " after");
}

#[test]
fn test_hyperlink_run_uses_theme_color_over_explicit_fill_and_keeps_underline_none() {
    let runs_xml = concat!(
        r#"<a:r><a:rPr lang="en-US" u="none">"#,
        r#"<a:solidFill><a:schemeClr val="accent4"/></a:solidFill>"#,
        r#"<a:hlinkClick r:id="rId2"/></a:rPr><a:t>Styled link</a:t></a:r>"#,
    );
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(
        &[("accent4", "8064A2"), ("hlink", "0000FF")],
        "Calibri Light",
        "Calibri",
    );
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].style.color, Some(Color::new(0x00, 0x00, 0xFF)));
    assert_eq!(para.runs[0].style.underline, Some(false));
}

#[test]
fn test_theme_major_font_in_text() {
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

    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0x5B, 0x9B, 0xD5)));

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
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert!(shape.fill.is_none());
}

#[test]
fn test_scheme_color_tint_blends_toward_white() {
    // accent3=#A5A5A5 with tint 50% → each channel: 255 - (255-165)*0.5 = 210 = 0xD2
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"><a:tint val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(shape.fill, Some(Color::new(0xD2, 0xD2, 0xD2)));
}

#[test]
fn test_scheme_color_shade_blends_toward_black() {
    // accent1=#4472C4 with shade 50% → #2F528F (issue #667: the scale is in
    // linear light, so it is not each byte * 0.5)
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    // accent1 = (0x44, 0x72, 0xC4) = (68, 114, 196)
    // shade 50% in linear light: (47, 82, 143) = (0x2F, 0x52, 0x8F)
    assert_eq!(shape.fill, Some(Color::new(0x2F, 0x52, 0x8F)));
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

#[test]
fn test_parse_theme_line_style_widths() {
    // Theme lnStyleLst widths back <a:lnRef idx="N"> outline resolution
    // (issue #318): idx=1/2/3 map to the 1st/2nd/3rd entry's EMU width.
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X">
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let theme = parse_theme_xml(theme_xml);
    let widths: Vec<i64> = theme
        .line_styles
        .iter()
        .map(|style| style.width_emu)
        .collect();
    assert_eq!(widths, vec![6350, 12700, 19050]);
}

/// Each `lnStyleLst` entry keeps its own corner join, and an entry naming none
/// stays unstated so the DrawingML round default decides (issue #1090).
///
/// PowerPoint's stock themes write `<a:miter lim="800000"/>` on every entry,
/// so a deck built on one must keep mitering while `customGeo.pptx`'s
/// join-less theme rounds.
#[test]
fn test_parse_theme_line_style_joins() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X">
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:miter lim="800000"/></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:bevel/></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let theme = parse_theme_xml(theme_xml);
    let joins: Vec<Option<LineJoin>> = theme.line_styles.iter().map(|style| style.join).collect();
    assert_eq!(
        joins,
        vec![Some(LineJoin::Miter), Some(LineJoin::Bevel), None]
    );
}

/// A `<a:ln>` nested inside a `lnStyleLst` entry — the arrowhead and
/// compound-line sub-elements can carry one — must not be mistaken for a
/// fourth entry, which would shift every `<a:lnRef idx>` past its style.
#[test]
fn test_theme_line_styles_ignore_nested_lines() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X">
      <a:lnStyleLst>
        <a:ln w="6350"><a:miter lim="800000"/><a:extLst><a:ext><a:ln w="99999"><a:bevel/></a:ln></a:ext></a:extLst></a:ln>
        <a:ln w="12700"><a:round/></a:ln>
      </a:lnStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let theme = parse_theme_xml(theme_xml);
    assert_eq!(
        theme
            .line_styles
            .iter()
            .map(|style| (style.width_emu, style.join))
            .collect::<Vec<_>>(),
        vec![
            (6350, Some(LineJoin::Miter)),
            (12700, Some(LineJoin::Round))
        ]
    );
}

// ── `<a:effectRef>` resolution (issue #740) ────────────────────────────

/// Theme carrying three distinct effect styles. The first two are the shadows
/// PowerPoint's stock themes ship. The third states its shadow color as `phClr`
/// and also carries the bounded orthographic `scene3d` plus circular top
/// `sp3d/bevelT` fixture, exercising both theme paths.
fn theme_xml_with_effect_styles() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="T"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="76200" dist="38100" dir="10800000" rotWithShape="0"><a:srgbClr val="FF0000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="12700" dist="12700" dir="2700000" rotWithShape="0"><a:schemeClr val="phClr"/></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
}

/// `idx="1"` picks the first effect style. Values are the ones stock themes
/// carry: a 40000 EMU blur, a 20000 EMU drop straight down, 38% black.
#[test]
fn effect_ref_resolves_the_first_theme_effect_style() {
    let theme = parse_theme_xml(theme_xml_with_effect_styles());
    let shadow = resolve_effect_ref(1, None, &theme, &default_color_map())
        .expect("idx 1 resolves to the first effect style");

    assert!((shadow.blur_radius - emu_to_pt(40000)).abs() < 1e-9);
    assert!((shadow.distance - emu_to_pt(20000)).abs() < 1e-9);
    assert!(
        (shadow.direction - 90.0).abs() < 1e-9,
        "5400000 is straight down"
    );
    assert_eq!(shadow.color, Color::new(0, 0, 0));
    assert!((shadow.opacity - 0.38).abs() < 1e-6);
}

/// Triangulation: a different index resolves to that entry's own values, so
/// the resolution cannot be a fixed shadow.
#[test]
fn effect_ref_resolves_each_index_to_its_own_entry() {
    let theme = parse_theme_xml(theme_xml_with_effect_styles());
    let shadow = resolve_effect_ref(2, None, &theme, &default_color_map())
        .expect("idx 2 resolves to the second effect style");

    assert!((shadow.blur_radius - emu_to_pt(76200)).abs() < 1e-9);
    assert!((shadow.distance - emu_to_pt(38100)).abs() < 1e-9);
    assert!((shadow.direction - 180.0).abs() < 1e-9);
    assert_eq!(shadow.color, Color::new(0xFF, 0, 0));
    assert!((shadow.opacity - 0.60).abs() < 1e-6);
}

/// The reference's own color binds `phClr` inside the entry, matching how
/// `<p:bgRef>` binds it for a fill entry.
#[test]
fn effect_ref_binds_its_color_to_the_placeholder() {
    let theme = parse_theme_xml(theme_xml_with_effect_styles());
    let reference_color = Color::new(0x44, 0x72, 0xC4);
    let shadow = resolve_effect_ref(3, Some(reference_color), &theme, &default_color_map())
        .expect("idx 3 resolves to the third effect style");

    assert_eq!(shadow.color, reference_color);
}

/// The theme effect style keeps its `scene3d`/`sp3d` siblings alongside the
/// already-supported shadow. This is the 5 x 2pt top bevel used by
/// `customGeo.pptx` page 44 (issue #1298).
#[test]
fn effect_ref_resolves_the_theme_top_bevel() {
    let theme = parse_theme_xml(theme_xml_with_effect_styles());
    let bevel = resolve_effect_ref_top_bevel(3, &theme)
        .expect("idx 3 resolves its orthographic three-point top bevel");

    assert!((bevel.width - 5.0).abs() < 1e-9);
    assert!((bevel.height - 2.0).abs() < 1e-9);
}

/// `idx="0"` means no effect, and an index past the list resolves to nothing
/// rather than to the last entry.
#[test]
fn effect_ref_index_zero_and_out_of_range_resolve_to_nothing() {
    let theme = parse_theme_xml(theme_xml_with_effect_styles());
    assert!(resolve_effect_ref(0, None, &theme, &default_color_map()).is_none());
    assert!(resolve_effect_ref(4, None, &theme, &default_color_map()).is_none());
}

/// A shape whose only effect is the `<p:style>` reference gets the theme's
/// shadow — the end-to-end path issue #740 reported as broken.
fn slide_with_styled_shape(sp_pr_effect: &str, style_body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Panel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{sp_pr_effect}</p:spPr><p:style>{style_body}</p:style><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
    )
}

fn first_shape_shadow(slide: &str) -> Option<Shadow> {
    let data = build_test_pptx_with_theme_layout_master(
        9144000,
        6858000,
        slide,
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#,
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#,
        theme_xml_with_effect_styles(),
    );
    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .iter()
        .flat_map(|page| match page {
            Page::Fixed(fixed) => fixed.elements.iter().collect::<Vec<_>>(),
            _ => Vec::new(),
        })
        .find_map(|element| match element.kind {
            FixedElementKind::Shape(ref shape) => shape.shadow.clone(),
            _ => None,
        })
}

fn first_shape_top_bevel(slide: &str) -> Option<TopBevel> {
    let data = build_test_pptx_with_theme_layout_master(
        9144000,
        6858000,
        slide,
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#,
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#,
        theme_xml_with_effect_styles(),
    );
    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    doc.pages
        .iter()
        .flat_map(|page| match page {
            Page::Fixed(fixed) => fixed.elements.iter().collect::<Vec<_>>(),
            _ => Vec::new(),
        })
        .find_map(|element| match element.kind {
            FixedElementKind::Shape(ref shape) => shape.top_bevel.clone(),
            _ => None,
        })
}

#[test]
fn styled_shape_inherits_the_theme_shadow_through_effect_ref() {
    let slide = slide_with_styled_shape(
        "",
        r#"<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>"#,
    );

    let shadow = first_shape_shadow(&slide).expect("the theme effect reaches the shape");
    assert!((shadow.blur_radius - emu_to_pt(40000)).abs() < 1e-9);
    assert!((shadow.distance - emu_to_pt(20000)).abs() < 1e-9);
    assert!((shadow.opacity - 0.38).abs() < 1e-6);
}

#[test]
fn styled_shape_inherits_the_theme_top_bevel_through_effect_ref() {
    let slide = slide_with_styled_shape(
        "",
        r#"<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>"#,
    );

    let bevel = first_shape_top_bevel(&slide).expect("the theme bevel reaches the shape");
    assert!((bevel.width - 5.0).abs() < 1e-9);
    assert!((bevel.height - 2.0).abs() < 1e-9);
    assert!((bevel.light_rig_rotation_deg - 20.0).abs() < 1e-9);
}

/// An `<a:effectLst/>` of the shape's own means "no effect" and must win over
/// the reference — otherwise every shape that switches its theme shadow off
/// would get it back.
#[test]
fn an_empty_effect_list_on_the_shape_suppresses_the_reference() {
    let slide = slide_with_styled_shape(
        "<a:effectLst/>",
        r#"<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>"#,
    );

    assert!(first_shape_shadow(&slide).is_none());
}

/// The shape's own shadow wins over the reference when both are present.
#[test]
fn a_direct_shadow_wins_over_the_effect_ref() {
    let slide = slide_with_styled_shape(
        r#"<a:effectLst><a:outerShdw blurRad="25400" dist="25400" dir="0"><a:srgbClr val="00FF00"/></a:outerShdw></a:effectLst>"#,
        r#"<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>"#,
    );

    let shadow = first_shape_shadow(&slide).expect("the direct shadow survives");
    assert_eq!(shadow.color, Color::new(0, 0xFF, 0));
}

/// The paragraph mark's family reaches the IR because it shares the final
/// physical line's box with the text fonts (#1176, #1177).
///
/// A mark that declares no typeface inherits the presentation's default text
/// style, whose `<a:latin typeface="+mn-lt"/>` names the theme's minor Latin
/// font — so it is that face, not the run's and not the renderer's default,
/// that shares this single-line fixture's box with the text.
#[test]
fn a_bare_paragraph_mark_takes_the_theme_minor_latin_font() {
    let runs_xml = r#"<a:r><a:rPr sz="4000"><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/></a:rPr><a:t>2025년 4분기</a:t></a:r><a:endParaRPr lang="en-US" sz="4000"/>"#;
    let shape = make_formatted_text_box(0, 0, 10_000_000, 1_000_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let Block::Paragraph(para) = &blocks[0] else {
        panic!("Expected Paragraph");
    };

    assert_eq!(
        para.runs[0].style.font_family.as_deref(),
        Some("Malgun Gothic"),
        "the run keeps its own face"
    );
    assert_eq!(
        para.style.paragraph_mark_font_family.as_deref(),
        Some("Calibri"),
        "a mark declaring no typeface falls to the theme's minor Latin font"
    );
}

/// A mark that names its own typeface keeps it — the theme fallback applies
/// only when nothing else does (issue #1176).
#[test]
fn a_paragraph_mark_keeps_the_typeface_it_declares() {
    let runs_xml = r#"<a:r><a:rPr sz="4000"><a:latin typeface="Malgun Gothic"/></a:rPr><a:t>2025년</a:t></a:r><a:endParaRPr lang="en-US" sz="4000"><a:latin typeface="Verdana"/></a:endParaRPr>"#;
    let shape = make_formatted_text_box(0, 0, 10_000_000, 1_000_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let Block::Paragraph(para) = &blocks[0] else {
        panic!("Expected Paragraph");
    };

    assert_eq!(
        para.style.paragraph_mark_font_family.as_deref(),
        Some("Verdana"),
        "the mark's own typeface must beat the theme fallback"
    );
}

/// A text box whose body-level `<a:lstStyle>` names `list_style_face` for
/// level 1, holding `paragraphs_xml` verbatim.
///
/// The list style is the only place that face appears, so anything set in it
/// reached the paragraph through inheritance rather than from a run.
fn text_box_with_list_style_face(list_style_face: &str, paragraphs_xml: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="2000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2000"><a:latin typeface="{list_style_face}"/></a:defRPr></a:lvl1pPr></a:lstStyle>{paragraphs_xml}</p:txBody></p:sp>"#
    )
}

/// The `paragraph_mark_font_family` of every paragraph of a slide built from
/// one shape.
fn paragraph_mark_families(shape_xml: &str, theme_xml: &str) -> Vec<Option<String>> {
    let slide = make_slide_xml(&[shape_xml.to_string()]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);
    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    text_box_blocks(&page.elements[0])
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(para) => Some(
                para.style
                    .paragraph_mark_font_family
                    .as_deref()
                    .map(str::to_string),
            ),
            _ => None,
        })
        .collect()
}

/// A paragraph that writes no `<a:endParaRPr>` at all puts no inherited face
/// on PowerPoint's shared line box (issue #1645).
///
/// Measured with two native one-factor probes over eight sizes from 11pt to
/// 53pt, exported through `scripts/probe_harness.py --backend office`
/// (`scripts/probes/issue-1645-paragraph-mark-list-style-face.json` and
/// `issue-1645-absent-paragraph-mark-theme-face.json`): with the run naming
/// its own typeface and no `<a:endParaRPr>` present, replacing the body list
/// style's `<a:latin>` with Meiryo or Calibri — or the theme's minor Latin
/// font with either — moves no baseline at all, while declaring
/// `<a:endParaRPr><a:latin typeface="Meiryo"/></a:endParaRPr>` moves all eight
/// by up to 5.04pt. The mark is typed in the face of the text it follows, so
/// it contributes nothing the runs have not already put on the line.
#[test]
fn an_undeclared_paragraph_mark_adds_no_face_to_the_line() {
    let paragraph = r#"<a:p><a:r><a:rPr sz="1600"><a:latin typeface="Verdana"/></a:rPr><a:t>Manager</a:t></a:r></a:p>"#;
    let shape = text_box_with_list_style_face("Gill Sans MT", paragraph);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Gill Sans MT", "Calibri");

    assert_eq!(
        paragraph_mark_families(&shape, &theme_xml),
        vec![Some("Verdana".to_string())],
        "an absent mark must take the run's own face, not the list style's \
         and not the theme's minor Latin font"
    );
}

/// Triangulation for [`an_undeclared_paragraph_mark_adds_no_face_to_the_line`]
/// with a different face triple, so no single family can be hard-coded.
#[test]
fn an_undeclared_paragraph_mark_adds_no_face_for_another_face_triple() {
    let paragraph = r#"<a:p><a:r><a:rPr sz="1600"><a:latin typeface="Malgun Gothic"/></a:rPr><a:t>Key employee</a:t></a:r></a:p>"#;
    let shape = text_box_with_list_style_face("Meiryo", paragraph);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Meiryo", "MS Gothic");

    assert_eq!(
        paragraph_mark_families(&shape, &theme_xml),
        vec![Some("Malgun Gothic".to_string())],
        "the rule is the run's own face, not one particular family"
    );
}

/// The other branch of the same rule: an `<a:endParaRPr>` that is *present*
/// but names no typeface does inherit the list style's face.
///
/// The `lst-meiryo-bare-mark` variant of the native probe moves all eight
/// baselines by exactly what an explicitly declared Meiryo mark moves them,
/// so presence — not a declared `<a:latin>` — is what puts the inherited face
/// on the line.
#[test]
fn a_present_but_bare_paragraph_mark_inherits_the_list_style_face() {
    let paragraph = r#"<a:p><a:r><a:rPr sz="1600"><a:latin typeface="Verdana"/></a:rPr><a:t>Manager</a:t></a:r><a:endParaRPr lang="en-US" sz="1600"/></a:p>"#;
    let shape = text_box_with_list_style_face("Gill Sans MT", paragraph);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Gill Sans MT", "Calibri");

    assert_eq!(
        paragraph_mark_families(&shape, &theme_xml),
        vec![Some("Gill Sans MT".to_string())],
        "a mark element that is present inherits the list style's face"
    );
}

/// The shape issue #1645 reports: a name paragraph in the major Latin font
/// followed by a role paragraph that names the minor one, neither writing an
/// `<a:endParaRPr>`.
///
/// The role line's box must be the minor font's alone. Sharing it with the
/// major font seats the role baseline a point high, because the two models
/// straddle the whole point PowerPoint rounds the story position to.
#[test]
fn neither_paragraph_of_a_name_and_role_pair_takes_the_others_face() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>August Bergquist</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr lang="en-US" sz="1600" i="1"><a:latin typeface="+mn-lt"/></a:rPr><a:t>Manager</a:t></a:r></a:p>"#,
    );
    let shape = text_box_with_list_style_face("+mj-lt", paragraphs);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Gill Sans MT", "Arial");

    assert_eq!(
        paragraph_mark_families(&shape, &theme_xml),
        vec![Some("Gill Sans MT".to_string()), Some("Arial".to_string())],
        "each paragraph's mark stays on the face its own run is set in"
    );
}

/// End to end for issue #1645: the role line of a name-and-role pair is
/// seated by its own face alone.
///
/// The deck is the shape slide 10 of the #1220 deck uses — a placeholder list
/// style naming `+mj-lt`, a name paragraph inheriting it, and a role
/// paragraph naming `+mn-lt`, neither writing an `<a:endParaRPr>`. Both faces
/// are Typst's own embedded ones so the test does not depend on what the host
/// has installed, and the sizes are chosen as the first ones where sharing the
/// box changes the whole point PowerPoint seats the baseline on — at a size
/// where the two models agree the test would pass without the fix.
#[test]
fn a_role_paragraph_is_seated_by_its_own_face_alone() {
    let major: &str = "Libertinus Serif";
    let minor: &str = "DejaVu Sans Mono";
    let (Some((major_em, _)), Some((minor_em, _)), Some((shared_em, _))) = (
        crate::render::pdf::powerpoint_line_box_em(major),
        crate::render::pdf::powerpoint_line_box_em(minor),
        crate::render::pdf::powerpoint_line_box_em_for_families(&[major, minor]),
    ) else {
        return;
    };
    let name_size_pt: f64 = discriminating_size_pt(major_em, shared_em);
    let role_size_pt: f64 = discriminating_size_pt(minor_em, shared_em);

    let paragraphs = format!(
        concat!(
            r#"<a:p><a:r><a:rPr lang="en-US" sz="{name}"/><a:t>August Bergquist</a:t></a:r></a:p>"#,
            r#"<a:p><a:r><a:rPr lang="en-US" sz="{role}"><a:latin typeface="+mn-lt"/></a:rPr>"#,
            r#"<a:t>Manager</a:t></a:r></a:p>"#,
        ),
        name = (name_size_pt * 100.0) as i64,
        role = (role_size_pt * 100.0) as i64,
    );
    let shape = text_box_with_list_style_face("+mj-lt", &paragraphs);
    let slide = make_slide_xml(&[shape]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), major, minor);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&doc)
        .unwrap()
        .source;
    let seats: Vec<f64> = emitted_line_box_top_edges_pt(&source);

    assert_eq!(
        seats,
        vec![
            (major_em * name_size_pt).round(),
            (minor_em * role_size_pt).round(),
        ],
        "each paragraph seats on its own face's share of the 1.2em box; \
         sharing them would seat the role line at \
         {shared:?}",
        shared = (shared_em * role_size_pt).round(),
    );
}

/// The smallest slide size at which two ascent shares seat the baseline on
/// different whole points — the only sizes at which a test can tell them
/// apart, because PowerPoint rounds the seat to a point.
fn discriminating_size_pt(alone_em: f64, shared_em: f64) -> f64 {
    (8..=72)
        .map(f64::from)
        .find(|size| (alone_em * size).round() != (shared_em * size).round())
        .expect("the two models must round apart at some slide size")
}

/// Every `top-edge:` the generated source states, in points and in order.
fn emitted_line_box_top_edges_pt(source: &str) -> Vec<f64> {
    source
        .split("top-edge: ")
        .skip(1)
        .filter_map(|rest| rest.split_once("pt")?.0.parse().ok())
        .collect()
}
