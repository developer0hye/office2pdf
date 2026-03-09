use super::*;

fn build_test_pptx_with_metadata(core_xml: &str) -> Vec<u8> {
    let slide = make_slide_xml(&[make_text_box(0, 0, 9144000, 6858000, "Hello")]);
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

    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    )
    .unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide.as_bytes()).unwrap();

    zip.start_file("docProps/core.xml", opts).unwrap();
    zip.write_all(core_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_parse_pptx_extracts_metadata() {
    let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My PPTX Title</dc:title>
  <dc:creator>PPTX Author</dc:creator>
  <dc:subject>PPTX Subject</dc:subject>
  <dc:description>PPTX description</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-05-01T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-15T18:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

    let data = build_test_pptx_with_metadata(core_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.metadata.title.as_deref(), Some("My PPTX Title"));
    assert_eq!(doc.metadata.author.as_deref(), Some("PPTX Author"));
    assert_eq!(doc.metadata.subject.as_deref(), Some("PPTX Subject"));
    assert_eq!(
        doc.metadata.description.as_deref(),
        Some("PPTX description")
    );
    assert_eq!(
        doc.metadata.created.as_deref(),
        Some("2024-05-01T09:00:00Z")
    );
    assert_eq!(
        doc.metadata.modified.as_deref(),
        Some("2024-06-15T18:30:00Z")
    );
}

#[test]
fn test_parse_pptx_without_metadata_no_crash() {
    let slide = make_slide_xml(&[make_text_box(0, 0, 9144000, 6858000, "Hello")]);
    let data = build_test_pptx(9144000, 6858000, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert!(doc.metadata.title.is_none());
    assert!(doc.metadata.author.is_none());
}
