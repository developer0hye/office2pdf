use super::*;

fn make_smartart_data_xml(items: &[&str]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>"#,
    );
    xml.push_str(
        r#"<dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
    );
    for (index, item) in items.iter().enumerate() {
        xml.push_str(&format!(
            r#"<dgm:pt modelId="{}" type="node"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>{item}</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
            index + 1
        ));
    }
    xml.push_str("</dgm:ptLst>");
    xml.push_str("<dgm:cxnLst>");
    for (index, _) in items.iter().enumerate() {
        xml.push_str(&format!(
            r#"<dgm:cxn modelId="{}" type="parOf" srcId="0" destId="{}"/>"#,
            100 + index,
            index + 1,
        ));
    }
    xml.push_str("</dgm:cxnLst>");
    xml.push_str("</dgm:dataModel>");
    xml
}

fn make_smartart_graphic_frame(x: i64, y: i64, cx: i64, cy: i64, dm_rid: &str) -> String {
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="{dm_rid}" r:lo="rId99" r:qs="rId98" r:cs="rId97"/></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

fn build_test_pptx_with_smartart(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xml: &str,
    data_rid: &str,
    data_xml: &str,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    let presentation_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#
    );
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(presentation_xml.as_bytes()).unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide_xml.as_bytes()).unwrap();

    let slide_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{data_rid}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#
    );
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(slide_rels.as_bytes()).unwrap();

    zip.start_file("ppt/diagrams/data1.xml", opts).unwrap();
    zip.write_all(data_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn get_smartart(elem: &FixedElement) -> &SmartArt {
    match &elem.kind {
        FixedElementKind::SmartArt(smartart) => smartart,
        _ => panic!("Expected SmartArt, got {:?}", elem.kind),
    }
}

#[test]
fn test_slide_with_smartart_produces_items() {
    let smartart_frame =
        make_smartart_graphic_frame(914_400, 1_828_800, 5_486_400, 3_086_100, "rId5");
    let slide_xml = make_slide_xml(&[smartart_frame]);
    let data_xml = make_smartart_data_xml(&["Step 1", "Step 2", "Step 3"]);
    let data = build_test_pptx_with_smartart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &data_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let smartart_elements: Vec<_> = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::SmartArt(_)))
        .collect();
    assert_eq!(smartart_elements.len(), 1);

    let smartart = get_smartart(smartart_elements[0]);
    let texts: Vec<&str> = smartart
        .items
        .iter()
        .map(|item| item.text.as_str())
        .collect();
    assert_eq!(texts, vec!["Step 1", "Step 2", "Step 3"]);
    assert!(smartart.items.iter().all(|item| item.depth == 0));
    assert!((smartart_elements[0].x - 72.0).abs() < 0.1);
    assert!((smartart_elements[0].y - 144.0).abs() < 0.1);
}

#[test]
fn test_slide_with_smartart_and_text_box() {
    let text_box = make_text_box(100_000, 100_000, 500_000, 200_000, "Title");
    let smartart_frame =
        make_smartart_graphic_frame(500_000, 500_000, 3_000_000, 2_000_000, "rId5");
    let slide_xml = make_slide_xml(&[text_box, smartart_frame]);
    let data_xml = make_smartart_data_xml(&["Item A", "Item B"]);
    let data = build_test_pptx_with_smartart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &data_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let smartart_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::SmartArt(_)))
        .count();
    let text_box_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::TextBox(_)))
        .count();
    assert_eq!(smartart_count, 1);
    assert!(text_box_count >= 1);

    let smartart_element = page
        .elements
        .iter()
        .find(|element| matches!(element.kind, FixedElementKind::SmartArt(_)))
        .unwrap();
    let smartart = get_smartart(smartart_element);
    let texts: Vec<&str> = smartart
        .items
        .iter()
        .map(|item| item.text.as_str())
        .collect();
    assert_eq!(texts, vec!["Item A", "Item B"]);
}

#[test]
fn test_slide_without_smartart_no_smartart_elements() {
    let text_box = make_text_box(0, 0, 500_000, 200_000, "No SmartArt");
    let slide_xml = make_slide_xml(&[text_box]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let smartart_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::SmartArt(_)))
        .count();
    assert_eq!(smartart_count, 0);
}
