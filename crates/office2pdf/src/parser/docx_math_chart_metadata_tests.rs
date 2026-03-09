use super::*;
use std::io::Cursor;

#[test]
fn test_parse_docx_with_display_math_fraction() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>Before math</w:t></w:r>
        </w:p>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:f>
                        <m:num><m:r><m:t>a</m:t></m:r></m:num>
                        <m:den><m:r><m:t>b</m:t></m:r></m:den>
                    </m:f>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:p>
            <w:r><w:t>After math</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::MathEquation(math) => Some(math),
            _ => None,
        })
        .collect();

    assert!(
        !math_blocks.is_empty(),
        "Expected at least one MathEquation block, found none"
    );
    assert_eq!(math_blocks[0].content, "frac(a, b)");
    assert!(math_blocks[0].display);
}

#[test]
fn test_parse_docx_with_inline_math_superscript() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>The value of </w:t></w:r>
            <m:oMath>
                <m:sSup>
                    <m:e><m:r><m:t>x</m:t></m:r></m:e>
                    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
            </m:oMath>
            <w:r><w:t> is positive</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::MathEquation(math) => Some(math),
            _ => None,
        })
        .collect();

    assert!(
        !math_blocks.is_empty(),
        "Expected at least one MathEquation block"
    );
    assert_eq!(math_blocks[0].content, "x^2");
    assert!(!math_blocks[0].display);
}

#[test]
fn test_parse_docx_with_complex_math() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:r><m:t>E</m:t></m:r>
                    <m:r><m:t>=</m:t></m:r>
                    <m:r><m:t>m</m:t></m:r>
                    <m:sSup>
                        <m:e><m:r><m:t>c</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                    </m:sSup>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_math(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(fp) => fp,
        _ => panic!("Expected FlowPage"),
    };

    let math_blocks: Vec<&MathEquation> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::MathEquation(math) => Some(math),
            _ => None,
        })
        .collect();

    assert!(!math_blocks.is_empty());
    assert_eq!(math_blocks[0].content, "E=m c^2");
    assert!(math_blocks[0].display);
}

fn build_docx_with_chart(document_xml: &str, chart_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.start_file("word/charts/chart1.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, chart_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_docx_with_bar_chart() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt>
            <c:pt idx="1"><c:v>Q2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>100</c:v></c:pt>
            <c:pt idx="1"><c:v>200</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let data = build_docx_with_chart(document_xml, chart_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("parse should succeed");

    let content = match &doc.pages[0] {
        Page::Flow(flow_page) => &flow_page.content,
        _ => panic!("Expected FlowPage"),
    };
    let chart_blocks: Vec<&Chart> = content
        .iter()
        .filter_map(|block| match block {
            Block::Chart(chart) => Some(chart),
            _ => None,
        })
        .collect();

    assert_eq!(chart_blocks.len(), 1);
    assert_eq!(chart_blocks[0].chart_type, ChartType::Bar);
    assert_eq!(chart_blocks[0].title.as_deref(), Some("Sales"));
    assert_eq!(chart_blocks[0].categories, vec!["Q1", "Q2"]);
    assert_eq!(chart_blocks[0].series.len(), 1);
    assert_eq!(chart_blocks[0].series[0].name.as_deref(), Some("Revenue"));
    assert_eq!(chart_blocks[0].series[0].values, vec![100.0, 200.0]);
}

#[test]
fn test_parse_docx_with_pie_chart() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:cat><c:strLit>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
            <c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strLit></c:cat>
          <c:val><c:numLit>
            <c:pt idx="0"><c:v>30</c:v></c:pt>
            <c:pt idx="1"><c:v>50</c:v></c:pt>
            <c:pt idx="2"><c:v>20</c:v></c:pt>
          </c:numLit></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let data = build_docx_with_chart(document_xml, chart_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(&data, &ConvertOptions::default())
        .expect("parse should succeed");

    let content = match &doc.pages[0] {
        Page::Flow(flow_page) => &flow_page.content,
        _ => panic!("Expected FlowPage"),
    };
    let chart_blocks: Vec<&Chart> = content
        .iter()
        .filter_map(|block| match block {
            Block::Chart(chart) => Some(chart),
            _ => None,
        })
        .collect();

    assert_eq!(chart_blocks.len(), 1);
    assert_eq!(chart_blocks[0].chart_type, ChartType::Pie);
    assert!(chart_blocks[0].title.is_none());
    assert_eq!(chart_blocks[0].categories, vec!["A", "B", "C"]);
    assert_eq!(chart_blocks[0].series[0].values, vec![30.0, 50.0, 20.0]);
}

fn build_docx_with_metadata(core_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

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
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
    )
    .unwrap();

    zip.start_file("docProps/core.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, core_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_docx_extracts_metadata() {
    let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My DOCX Title</dc:title>
  <dc:creator>DOCX Author</dc:creator>
  <dc:subject>DOCX Subject</dc:subject>
  <dc:description>DOCX description text</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-03-15T08:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-04-20T12:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

    let data = build_docx_with_metadata(core_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.metadata.title.as_deref(), Some("My DOCX Title"));
    assert_eq!(doc.metadata.author.as_deref(), Some("DOCX Author"));
    assert_eq!(doc.metadata.subject.as_deref(), Some("DOCX Subject"));
    assert_eq!(
        doc.metadata.description.as_deref(),
        Some("DOCX description text")
    );
    assert_eq!(
        doc.metadata.created.as_deref(),
        Some("2024-03-15T08:00:00Z")
    );
    assert_eq!(
        doc.metadata.modified.as_deref(),
        Some("2024-04-20T12:30:00Z")
    );
}

#[test]
fn test_parse_docx_without_metadata_no_crash() {
    let data = build_docx_with_math(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>No metadata</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert!(doc.metadata.title.is_none());
    assert!(doc.metadata.author.is_none());
}
