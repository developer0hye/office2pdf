use super::*;

// =============================================================================
// Step 1: GUID parsing, deobfuscation, font format detection
// =============================================================================

#[test]
fn parse_guid_to_bytes_valid_braces() {
    // Standard GUID: {7B19B49C-2336-4F82-AAD2-5D2BAE389560}
    // Mixed-endian: first 4 LE, next 2 LE, next 2 LE, remaining 8 BE
    let key = parse_guid_to_bytes("{7B19B49C-2336-4F82-AAD2-5D2BAE389560}").unwrap();
    // Group 1: 7B19B49C → LE → [0x9C, 0xB4, 0x19, 0x7B]
    // Group 2: 2336 → LE → [0x36, 0x23]
    // Group 3: 4F82 → LE → [0x82, 0x4F]
    // Group 4+5: AAD2 5D2BAE389560 → BE → [0xAA, 0xD2, 0x5D, 0x2B, 0xAE, 0x38, 0x95, 0x60]
    assert_eq!(
        key,
        [
            0x9C, 0xB4, 0x19, 0x7B, 0x36, 0x23, 0x82, 0x4F, 0xAA, 0xD2, 0x5D, 0x2B, 0xAE, 0x38,
            0x95, 0x60
        ]
    );
}

#[test]
fn parse_guid_to_bytes_no_braces() {
    let key = parse_guid_to_bytes("7B19B49C-2336-4F82-AAD2-5D2BAE389560").unwrap();
    assert_eq!(
        key,
        [
            0x9C, 0xB4, 0x19, 0x7B, 0x36, 0x23, 0x82, 0x4F, 0xAA, 0xD2, 0x5D, 0x2B, 0xAE, 0x38,
            0x95, 0x60
        ]
    );
}

#[test]
fn parse_guid_to_bytes_lowercase() {
    let key = parse_guid_to_bytes("{7b19b49c-2336-4f82-aad2-5d2bae389560}").unwrap();
    assert_eq!(
        key,
        [
            0x9C, 0xB4, 0x19, 0x7B, 0x36, 0x23, 0x82, 0x4F, 0xAA, 0xD2, 0x5D, 0x2B, 0xAE, 0x38,
            0x95, 0x60
        ]
    );
}

#[test]
fn parse_guid_to_bytes_invalid_returns_none() {
    assert!(parse_guid_to_bytes("not-a-guid").is_none());
    assert!(parse_guid_to_bytes("").is_none());
    assert!(parse_guid_to_bytes("{ZZZZZZZZ-0000-0000-0000-000000000000}").is_none());
    // Wrong number of groups
    assert!(parse_guid_to_bytes("{7B19B49C-2336-4F82-AAD2}").is_none());
}

#[test]
fn deobfuscate_font_data_roundtrip() {
    // Start with a known TTF signature
    let original: Vec<u8> = {
        let mut v = vec![0x00, 0x01, 0x00, 0x00]; // TTF magic
        v.extend(vec![0xAB; 28]); // rest of first 32 bytes
        v.extend(vec![0xCD; 32]); // data beyond 32 bytes (untouched)
        v
    };

    let key: [u8; 16] = [
        0x9C, 0xB4, 0x19, 0x7B, 0x36, 0x23, 0x82, 0x4F, 0xAA, 0xD2, 0x5D, 0x2B, 0xAE, 0x38, 0x95,
        0x60,
    ];

    // Obfuscate: XOR first 32 bytes
    let mut obfuscated = original.clone();
    for i in 0..32 {
        obfuscated[i] ^= key[i % 16];
    }
    // Obfuscated data should differ from original in first 32 bytes
    assert_ne!(&obfuscated[..32], &original[..32]);
    // Data beyond 32 bytes unchanged
    assert_eq!(&obfuscated[32..], &original[32..]);

    // Deobfuscate
    deobfuscate_font_data(&mut obfuscated, &key);
    assert_eq!(obfuscated, original);
}

#[test]
fn deobfuscate_font_data_short_data() {
    // Data shorter than 32 bytes — should XOR only available bytes
    let key: [u8; 16] = [0x01; 16];
    let mut data = vec![0x00; 10];
    deobfuscate_font_data(&mut data, &key);
    assert_eq!(data, vec![0x01; 10]);
}

#[test]
fn detect_font_format_ttf() {
    let data = [0x00, 0x01, 0x00, 0x00, 0xFF, 0xFF];
    assert_eq!(detect_font_format(&data), Some(FontFileFormat::Ttf));
}

#[test]
fn detect_font_format_otf() {
    let data = b"OTTO\x00\x01";
    assert_eq!(detect_font_format(data), Some(FontFileFormat::Otf));
}

#[test]
fn detect_font_format_ttc() {
    let data = b"ttcf\x00\x02\x00\x00";
    assert_eq!(detect_font_format(data), Some(FontFileFormat::Ttc));
}

#[test]
fn detect_font_format_unknown() {
    assert_eq!(detect_font_format(b"XXXX"), None);
    assert_eq!(detect_font_format(b""), None);
    assert_eq!(detect_font_format(b"OT"), None);
}

// =============================================================================
// Step 2: PPTX XML parsing
// =============================================================================

#[test]
fn parse_pptx_embedded_font_list_single_font() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:embeddedFontLst>
    <p:embeddedFont>
      <p:font typeface="Pretendard" pitchFamily="34" charset="2"/>
      <p:regular r:id="rId5"/>
    </p:embeddedFont>
  </p:embeddedFontLst>
</p:presentation>"#;

    let entries = parse_pptx_embedded_font_list(xml);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].typeface, "Pretendard");
    assert_eq!(entries[0].variants.len(), 1);
    assert_eq!(entries[0].variants[0].style, "regular");
    assert_eq!(entries[0].variants[0].r_id, "rId5");
}

#[test]
fn parse_pptx_embedded_font_list_multiple_variants() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:embeddedFontLst>
    <p:embeddedFont>
      <p:font typeface="TestFont"/>
      <p:regular r:id="rId10"/>
      <p:bold r:id="rId11"/>
      <p:italic r:id="rId12"/>
      <p:boldItalic r:id="rId13"/>
    </p:embeddedFont>
  </p:embeddedFontLst>
</p:presentation>"#;

    let entries = parse_pptx_embedded_font_list(xml);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].typeface, "TestFont");
    assert_eq!(entries[0].variants.len(), 4);

    let styles: Vec<&str> = entries[0]
        .variants
        .iter()
        .map(|v| v.style.as_str())
        .collect();
    assert!(styles.contains(&"regular"));
    assert!(styles.contains(&"bold"));
    assert!(styles.contains(&"italic"));
    assert!(styles.contains(&"boldItalic"));
}

#[test]
fn parse_pptx_embedded_font_list_empty() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
</p:presentation>"#;

    let entries = parse_pptx_embedded_font_list(xml);
    assert!(entries.is_empty());
}

#[test]
fn parse_pptx_embedded_font_list_multiple_fonts() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:embeddedFontLst>
    <p:embeddedFont>
      <p:font typeface="FontA"/>
      <p:regular r:id="rId1"/>
    </p:embeddedFont>
    <p:embeddedFont>
      <p:font typeface="FontB"/>
      <p:regular r:id="rId2"/>
      <p:bold r:id="rId3"/>
    </p:embeddedFont>
  </p:embeddedFontLst>
</p:presentation>"#;

    let entries = parse_pptx_embedded_font_list(xml);
    assert_eq!(entries.len(), 2);
    assert_eq!(entries[0].typeface, "FontA");
    assert_eq!(entries[1].typeface, "FontB");
    assert_eq!(entries[1].variants.len(), 2);
}

#[test]
fn extract_guid_from_font_path_valid() {
    let guid =
        extract_guid_from_font_path("ppt/fonts/{7B19B49C-2336-4F82-AAD2-5D2BAE389560}.fntdata");
    assert_eq!(
        guid.as_deref(),
        Some("{7B19B49C-2336-4F82-AAD2-5D2BAE389560}")
    );
}

#[test]
fn extract_guid_from_font_path_non_guid() {
    assert!(extract_guid_from_font_path("ppt/fonts/somefile.fntdata").is_none());
    assert!(extract_guid_from_font_path("ppt/slides/slide1.xml").is_none());
}

// =============================================================================
// Step 3: DOCX XML parsing
// =============================================================================

#[test]
fn parse_docx_embedded_font_entries_regular() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Pretendard">
    <w:embedRegular w:fontKey="{7B19B49C-2336-4F82-AAD2-5D2BAE389560}" r:id="rId1"/>
  </w:font>
</w:fonts>"#;

    let entries = parse_docx_embedded_font_entries(xml);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].font_name, "Pretendard");
    assert_eq!(entries[0].variants.len(), 1);
    assert_eq!(entries[0].variants[0].style, "regular");
    assert_eq!(entries[0].variants[0].r_id, "rId1");
    assert_eq!(
        entries[0].variants[0].font_key,
        "{7B19B49C-2336-4F82-AAD2-5D2BAE389560}"
    );
}

#[test]
fn parse_docx_embedded_font_entries_multiple_variants() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="TestFont">
    <w:embedRegular w:fontKey="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}" r:id="rId1"/>
    <w:embedBold w:fontKey="{11111111-2222-3333-4444-555555555555}" r:id="rId2"/>
    <w:embedItalic w:fontKey="{66666666-7777-8888-9999-AAAAAAAAAAAA}" r:id="rId3"/>
    <w:embedBoldItalic w:fontKey="{BBBBBBBB-CCCC-DDDD-EEEE-FFFFFFFFFFFF}" r:id="rId4"/>
  </w:font>
</w:fonts>"#;

    let entries = parse_docx_embedded_font_entries(xml);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].variants.len(), 4);
}

#[test]
fn parse_docx_embedded_font_entries_no_embedded() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Arial"/>
  <w:font w:name="Times New Roman"/>
</w:fonts>"#;

    let entries = parse_docx_embedded_font_entries(xml);
    assert!(entries.is_empty());
}

// =============================================================================
// Step 4: End-to-end extraction (synthetic ZIPs)
// =============================================================================

#[cfg(not(target_arch = "wasm32"))]
mod integration {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::ZipWriter;
    use zip::write::FileOptions;

    /// Build a minimal PPTX ZIP with an obfuscated embedded font.
    fn build_pptx_with_embedded_font(ttf_data: &[u8], guid: &str) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::default();

        // presentation.xml with embedded font list
        let pres_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:embeddedFontLst>
    <p:embeddedFont>
      <p:font typeface="TestFont"/>
      <p:regular r:id="rId5"/>
    </p:embeddedFont>
  </p:embeddedFontLst>
</p:presentation>"#;
        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(pres_xml.as_bytes()).unwrap();

        // presentation.xml.rels
        let font_target = format!("fonts/{guid}.fntdata");
        let rels_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="{font_target}"/>
</Relationships>"#
        );
        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(rels_xml.as_bytes()).unwrap();

        // Obfuscated font data
        let key = parse_guid_to_bytes(guid).unwrap();
        let mut font_bytes = ttf_data.to_vec();
        for i in 0..std::cmp::min(32, font_bytes.len()) {
            font_bytes[i] ^= key[i % 16];
        }
        let font_path = format!("ppt/fonts/{guid}.fntdata");
        zip.start_file(font_path, options).unwrap();
        zip.write_all(&font_bytes).unwrap();

        zip.finish().unwrap().into_inner()
    }

    /// Build a minimal DOCX ZIP with an obfuscated embedded font.
    fn build_docx_with_embedded_font(ttf_data: &[u8], guid: &str) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::default();

        // word/fontTable.xml
        let font_table_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="TestFont">
    <w:embedRegular w:fontKey="{guid}" r:id="rId1"/>
  </w:font>
</w:fonts>"#
        );
        zip.start_file("word/fontTable.xml", options).unwrap();
        zip.write_all(font_table_xml.as_bytes()).unwrap();

        // word/_rels/fontTable.xml.rels
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>
</Relationships>"#;
        zip.start_file("word/_rels/fontTable.xml.rels", options)
            .unwrap();
        zip.write_all(rels_xml.as_bytes()).unwrap();

        // Obfuscated font data
        let key = parse_guid_to_bytes(guid).unwrap();
        let mut font_bytes = ttf_data.to_vec();
        for i in 0..std::cmp::min(32, font_bytes.len()) {
            font_bytes[i] ^= key[i % 16];
        }
        zip.start_file("word/fonts/font1.odttf", options).unwrap();
        zip.write_all(&font_bytes).unwrap();

        zip.finish().unwrap().into_inner()
    }

    fn make_fake_ttf(size: usize) -> Vec<u8> {
        let mut data = vec![0u8; size];
        // TTF magic: 00 01 00 00
        data[0] = 0x00;
        data[1] = 0x01;
        data[2] = 0x00;
        data[3] = 0x00;
        // Fill rest with recognizable pattern
        for (i, byte) in data[4..size].iter_mut().enumerate() {
            *byte = ((i + 4) & 0xFF) as u8;
        }
        data
    }

    #[test]
    fn extract_pptx_embedded_fonts_roundtrip() {
        let guid = "{7B19B49C-2336-4F82-AAD2-5D2BAE389560}";
        let original_ttf = make_fake_ttf(128);
        let zip_data = build_pptx_with_embedded_font(&original_ttf, guid);

        let result = extract_embedded_fonts(&zip_data, crate::config::Format::Pptx);
        assert!(result.is_some(), "should extract fonts from PPTX");

        let dir = result.unwrap();
        assert!(!dir.is_empty());
        assert!(dir.path().exists());

        // Find the extracted font file
        let entries: Vec<_> = std::fs::read_dir(dir.path())
            .unwrap()
            .filter_map(|e| e.ok())
            .collect();
        assert_eq!(entries.len(), 1, "should have extracted one font file");

        let extracted = std::fs::read(entries[0].path()).unwrap();
        assert_eq!(
            &extracted[..4],
            &[0x00, 0x01, 0x00, 0x00],
            "deobfuscated font should start with TTF magic"
        );
        assert_eq!(
            extracted, original_ttf,
            "deobfuscated font should match original"
        );
    }

    #[test]
    fn extract_docx_embedded_fonts_roundtrip() {
        let guid = "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}";
        let original_ttf = make_fake_ttf(64);
        let zip_data = build_docx_with_embedded_font(&original_ttf, guid);

        let result = extract_embedded_fonts(&zip_data, crate::config::Format::Docx);
        assert!(result.is_some(), "should extract fonts from DOCX");

        let dir = result.unwrap();
        assert!(!dir.is_empty());

        let entries: Vec<_> = std::fs::read_dir(dir.path())
            .unwrap()
            .filter_map(|e| e.ok())
            .collect();
        assert_eq!(entries.len(), 1);

        let extracted = std::fs::read(entries[0].path()).unwrap();
        assert_eq!(extracted, original_ttf);
    }

    #[test]
    fn extract_embedded_fonts_no_fonts_returns_none() {
        // Build a minimal PPTX with no embedded fonts
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::default();

        let pres_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
</p:presentation>"#;
        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(pres_xml.as_bytes()).unwrap();

        let zip_data = zip.finish().unwrap().into_inner();

        let result = extract_embedded_fonts(&zip_data, crate::config::Format::Pptx);
        assert!(result.is_none());
    }

    #[test]
    fn extract_embedded_fonts_xlsx_returns_none() {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::default();
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(b"<workbook/>").unwrap();
        let zip_data = zip.finish().unwrap().into_inner();

        let result = extract_embedded_fonts(&zip_data, crate::config::Format::Xlsx);
        assert!(result.is_none());
    }

    #[test]
    fn embedded_font_dir_cleaned_up_on_drop() {
        let guid = "{7B19B49C-2336-4F82-AAD2-5D2BAE389560}";
        let original_ttf = make_fake_ttf(64);
        let zip_data = build_pptx_with_embedded_font(&original_ttf, guid);

        let path = {
            let dir = extract_embedded_fonts(&zip_data, crate::config::Format::Pptx).unwrap();
            let p = dir.path().to_path_buf();
            assert!(p.exists());
            p
            // dir drops here
        };
        assert!(!path.exists(), "temp dir should be cleaned up on drop");
    }
}
