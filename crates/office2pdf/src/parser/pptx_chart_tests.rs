use super::*;

fn make_chart_graphic_frame(x: i64, y: i64, cx: i64, cy: i64, chart_rid: &str) -> String {
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{chart_rid}"/></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

fn make_bar_chart_xml(title: &str, categories: &[&str], values: &[f64]) -> String {
    let mut category_xml = String::new();
    for (index, category) in categories.iter().enumerate() {
        category_xml.push_str(&format!(
            r#"<c:pt idx="{index}"><c:v>{category}</c:v></c:pt>"#
        ));
    }
    let mut value_xml = String::new();
    for (index, value) in values.iter().enumerate() {
        value_xml.push_str(&format!(r#"<c:pt idx="{index}"><c:v>{value}</c:v></c:pt>"#));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache>{category_xml}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache>{value_xml}</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
    )
}

fn build_test_pptx_with_chart(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xml: &str,
    chart_rid: &str,
    chart_xml: &str,
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
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{chart_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#
    );
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(slide_rels.as_bytes()).unwrap();

    zip.start_file("ppt/charts/chart1.xml", opts).unwrap();
    zip.write_all(chart_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn get_chart(elem: &FixedElement) -> &Chart {
    match &elem.kind {
        FixedElementKind::Chart(chart) => chart,
        _ => panic!("Expected Chart, got {:?}", elem.kind),
    }
}

#[test]
fn test_slide_with_chart_produces_chart_element() {
    let chart_frame = make_chart_graphic_frame(914_400, 1_828_800, 5_486_400, 3_086_100, "rId5");
    let slide_xml = make_slide_xml(&[chart_frame]);
    let chart_xml = make_bar_chart_xml("Sales Data", &["Q1", "Q2", "Q3"], &[100.0, 200.0, 150.0]);
    let data = build_test_pptx_with_chart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &chart_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let chart_elements: Vec<_> = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::Chart(_)))
        .collect();
    assert_eq!(chart_elements.len(), 1);

    let chart = get_chart(chart_elements[0]);
    assert_eq!(chart.title.as_deref(), Some("Sales Data"));
    assert_eq!(chart.categories, vec!["Q1", "Q2", "Q3"]);
    assert_eq!(chart.series.len(), 1);
    assert_eq!(chart.series[0].values, vec![100.0, 200.0, 150.0]);
    assert!((chart_elements[0].x - 72.0).abs() < 0.1);
    assert!((chart_elements[0].y - 144.0).abs() < 0.1);
}

#[test]
fn test_slide_with_chart_and_text_box() {
    let text_box = make_text_box(100_000, 100_000, 500_000, 200_000, "Title");
    let chart_frame = make_chart_graphic_frame(500_000, 500_000, 3_000_000, 2_000_000, "rId5");
    let slide_xml = make_slide_xml(&[text_box, chart_frame]);
    let chart_xml = make_bar_chart_xml("Revenue", &["Jan", "Feb"], &[50.0, 75.0]);
    let data = build_test_pptx_with_chart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &chart_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let chart_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::Chart(_)))
        .count();
    let text_box_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::TextBox(_)))
        .count();
    assert_eq!(chart_count, 1);
    assert!(text_box_count >= 1);
}

#[test]
fn test_slide_without_chart_no_chart_elements() {
    let text_box = make_text_box(0, 0, 500_000, 200_000, "No Chart");
    let slide_xml = make_slide_xml(&[text_box]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let chart_count = page
        .elements
        .iter()
        .filter(|element| matches!(element.kind, FixedElementKind::Chart(_)))
        .count();
    assert_eq!(chart_count, 0);
}

#[test]
fn test_scan_chart_refs_basic() {
    let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <p:cSld><p:spTree>
            <p:graphicFrame>
              <p:nvGraphicFramePr>
                <p:cNvPr id="5" name="Chart"/>
                <p:cNvGraphicFramePr/>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="914400" y="1828800"/>
                <a:ext cx="5486400" cy="3086100"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart r:id="rId5"/>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#;

    let refs = scan_chart_refs(slide_xml);
    assert_eq!(refs.len(), 1);
    assert_eq!(refs[0].x, 914400);
    assert_eq!(refs[0].y, 1828800);
    assert_eq!(refs[0].cx, 5486400);
    assert_eq!(refs[0].cy, 3086100);
    assert_eq!(refs[0].chart_rid, "rId5");
}

#[test]
fn test_scan_chart_refs_no_chart() {
    let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
              <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sld>"#;

    let refs = scan_chart_refs(slide_xml);
    assert!(refs.is_empty());
}
