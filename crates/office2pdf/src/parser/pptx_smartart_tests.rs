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
    let opts = zip::write::SimpleFileOptions::default();

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

#[test]
fn smartart_drawing_cache_preserves_text_body_list_and_run_styles() {
    let drawing_xml = r#"
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:sp>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="12700" y="25400"/>
                  <a:ext cx="1270000" cy="2540000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>
                <a:ln w="12700">
                  <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
                  <a:prstDash val="solid"/>
                </a:ln>
              </dsp:spPr>
              <dsp:style><a:fontRef idx="minor"/></dsp:style>
              <dsp:txBody>
                <a:bodyPr lIns="127000" tIns="254000" rIns="381000" bIns="508000" anchor="t"/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr lvl="1" marL="254000" indent="-127000" algn="l">
                    <a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
                    <a:spcAft><a:spcPts val="250"/></a:spcAft>
                    <a:buChar char="••"/>
                  </a:pPr>
                  <a:r>
                    <a:rPr sz="1700" b="0">
                      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                    </a:rPr>
                    <a:t>First item</a:t>
                  </a:r>
                  <a:endParaRPr sz="1700"/>
                </a:p>
                <a:p>
                  <a:pPr lvl="1" marL="254000" indent="-127000" algn="l">
                    <a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
                    <a:buChar char="••"/>
                  </a:pPr>
                  <a:r>
                    <a:rPr sz="1700" b="0">
                      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                    </a:rPr>
                    <a:t>Second item</a:t>
                  </a:r>
                  <a:endParaRPr sz="1700"/>
                </a:p>
              </dsp:txBody>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
    "#;

    let theme = ThemeData {
        minor_font: Some("Calibri".to_string()),
        ..ThemeData::default()
    };
    let elements =
        slides::parse_smartart_drawing(drawing_xml, &theme, &default_color_map(), 10.0, 20.0);
    let text_box = elements
        .iter()
        .find_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some((element, text_box)),
            _ => None,
        })
        .expect("the cache shape produces a styled text box");

    assert!((text_box.0.x - 11.0).abs() < 1e-9);
    assert!((text_box.0.y - 22.0).abs() < 1e-9);
    assert_eq!(
        text_box.1.padding,
        Insets {
            top: 20.0,
            right: 30.0,
            bottom: 40.0,
            left: 10.0,
        }
    );
    assert_eq!(text_box.1.vertical_align, TextBoxVerticalAlign::Top);
    assert_eq!(text_box.1.fill, Some(Color::new(0xDD, 0xEE, 0xFF)));
    assert_eq!(
        text_box
            .1
            .stroke
            .as_ref()
            .map(|stroke| (stroke.width, stroke.color)),
        Some((1.0, Color::new(0x11, 0x22, 0x33)))
    );

    let [Block::List(list)] = text_box.1.content.as_slice() else {
        panic!("the two cache paragraphs must remain a single bulleted list");
    };
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].level, 1);
    assert_eq!(list.items[0].content[0].runs[0].text, "First item");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second item");
    assert_eq!(
        list.level_styles
            .get(&1)
            .and_then(|style| style.marker_text.as_deref()),
        Some("•")
    );

    let first_paragraph = &list.items[0].content[0];
    assert_eq!(first_paragraph.style.alignment, Some(Alignment::Left));
    assert_eq!(first_paragraph.style.indent_left, Some(20.0));
    assert_eq!(first_paragraph.style.indent_first_line, Some(-10.0));
    assert_eq!(first_paragraph.style.space_after, Some(2.5));
    assert!(matches!(
        first_paragraph.style.line_spacing,
        Some(LineSpacing::Proportional(value)) if (value - 0.9).abs() < 1e-9
    ));
    assert_eq!(
        first_paragraph.runs[0].style.font_family.as_deref(),
        Some("Calibri")
    );
    assert_eq!(first_paragraph.runs[0].style.font_size, Some(17.0));
    assert_eq!(first_paragraph.runs[0].style.bold, Some(false));
    assert_eq!(first_paragraph.runs[0].style.color, Some(Color::black()));
}
