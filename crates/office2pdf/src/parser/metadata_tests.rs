use super::*;

#[test]
fn test_parse_core_xml_full_metadata() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document Title</dc:title>
  <dc:creator>John Doe</dc:creator>
  <dc:subject>Testing Subject</dc:subject>
  <dc:description>A test document description</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-06-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-07-20T14:00:00Z</dcterms:modified>
</cp:coreProperties>"#;

    let meta = parse_core_xml(xml);
    assert_eq!(meta.title.as_deref(), Some("Test Document Title"));
    assert_eq!(meta.author.as_deref(), Some("John Doe"));
    assert_eq!(meta.subject.as_deref(), Some("Testing Subject"));
    assert_eq!(
        meta.description.as_deref(),
        Some("A test document description")
    );
    assert_eq!(meta.created.as_deref(), Some("2024-06-15T10:30:00Z"));
    assert_eq!(meta.modified.as_deref(), Some("2024-07-20T14:00:00Z"));
}

#[test]
fn test_parse_core_xml_partial_metadata() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Only Title</dc:title>
</cp:coreProperties>"#;

    let meta = parse_core_xml(xml);
    assert_eq!(meta.title.as_deref(), Some("Only Title"));
    assert!(meta.author.is_none());
    assert!(meta.subject.is_none());
    assert!(meta.description.is_none());
    assert!(meta.created.is_none());
    assert!(meta.modified.is_none());
}

#[test]
fn test_parse_core_xml_empty_elements() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title></dc:title>
  <dc:creator></dc:creator>
</cp:coreProperties>"#;

    let meta = parse_core_xml(xml);
    assert!(meta.title.is_none());
    assert!(meta.author.is_none());
}

#[test]
fn test_parse_core_xml_no_metadata() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
</cp:coreProperties>"#;

    let meta = parse_core_xml(xml);
    assert!(meta.title.is_none());
    assert!(meta.author.is_none());
    assert!(meta.subject.is_none());
    assert!(meta.description.is_none());
    assert!(meta.created.is_none());
    assert!(meta.modified.is_none());
}

#[test]
fn test_parse_core_xml_invalid_xml() {
    let xml = "not valid xml at all <<<<";
    let meta = parse_core_xml(xml);
    // Should return default, not crash
    assert!(meta.title.is_none());
}

#[test]
fn test_extract_metadata_from_zip_with_core_xml() {
    use std::io::{Cursor, Write};

    let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ZIP Test Title</dc:title>
  <dc:creator>ZIP Author</dc:creator>
</cp:coreProperties>"#;

    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip_writer = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::default();
    zip_writer.start_file("docProps/core.xml", options).unwrap();
    zip_writer.write_all(core_xml.as_bytes()).unwrap();
    let cursor = zip_writer.finish().unwrap();

    let mut archive = ZipArchive::new(cursor).unwrap();
    let meta = extract_metadata_from_zip(&mut archive);
    assert_eq!(meta.title.as_deref(), Some("ZIP Test Title"));
    assert_eq!(meta.author.as_deref(), Some("ZIP Author"));
}

#[test]
fn test_extract_metadata_from_zip_without_core_xml() {
    use std::io::{Cursor, Write};

    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip_writer = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::default();
    zip_writer
        .start_file("some_other_file.xml", options)
        .unwrap();
    zip_writer.write_all(b"<root/>").unwrap();
    let cursor = zip_writer.finish().unwrap();

    let mut archive = ZipArchive::new(cursor).unwrap();
    let meta = extract_metadata_from_zip(&mut archive);
    assert!(meta.title.is_none());
    assert!(meta.author.is_none());
}
