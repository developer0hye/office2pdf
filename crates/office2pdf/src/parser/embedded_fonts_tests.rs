use super::*;

// =============================================================================
// Step 1: GUID parsing, deobfuscation, font format detection
// =============================================================================

#[test]
fn parse_guid_to_bytes_valid_braces() {
    // ECMA-376 Part 1 §17.8.1: the key is the GUID's hex digits read as bytes
    // in string order, then reversed. It is not the mixed-endian in-memory GUID
    // layout, reversed or not.
    let key = parse_guid_to_bytes("{7B19B49C-2336-4F82-AAD2-5D2BAE389560}").unwrap();
    // String order: 7B 19 B4 9C 23 36 4F 82 AA D2 5D 2B AE 38 95 60
    assert_eq!(
        key,
        [
            0x60, 0x95, 0x38, 0xAE, 0x2B, 0x5D, 0xD2, 0xAA, 0x82, 0x4F, 0x36, 0x23, 0x9C, 0xB4,
            0x19, 0x7B
        ]
    );
}

#[test]
fn parse_guid_to_bytes_reverses_every_group_in_string_order() {
    // The fontKey of the Press Start 2P face in
    // `issue_2200_embedded_press_start_2p.docx`. Each group's bytes are
    // reversed in place, and the groups come out last-first.
    let key = parse_guid_to_bytes("{497CBF7F-245D-42A2-A83F-AD87E4FBAC57}").unwrap();
    assert_eq!(
        key,
        [
            0x57, 0xAC, 0xFB, 0xE4, 0x87, 0xAD, 0x3F, 0xA8, 0xA2, 0x42, 0x5D, 0x24, 0x7F, 0xBF,
            0x7C, 0x49
        ]
    );
}

#[test]
fn parse_guid_to_bytes_no_braces() {
    let key = parse_guid_to_bytes("7B19B49C-2336-4F82-AAD2-5D2BAE389560").unwrap();
    assert_eq!(
        key,
        [
            0x60, 0x95, 0x38, 0xAE, 0x2B, 0x5D, 0xD2, 0xAA, 0x82, 0x4F, 0x36, 0x23, 0x9C, 0xB4,
            0x19, 0x7B
        ]
    );
}

#[test]
fn parse_guid_to_bytes_lowercase() {
    let key = parse_guid_to_bytes("{7b19b49c-2336-4f82-aad2-5d2bae389560}").unwrap();
    assert_eq!(
        key,
        [
            0x60, 0x95, 0x38, 0xAE, 0x2B, 0x5D, 0xD2, 0xAA, 0x82, 0x4F, 0x36, 0x23, 0x9C, 0xB4,
            0x19, 0x7B
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
    use zip::write::SimpleFileOptions;

    /// Build a minimal PPTX ZIP with an obfuscated embedded font.
    fn build_pptx_with_embedded_font(ttf_data: &[u8], guid: &str) -> Vec<u8> {
        build_pptx_with_named_embedded_font(ttf_data, guid, "TestFont")
    }

    fn build_pptx_with_named_embedded_font(ttf_data: &[u8], guid: &str, typeface: &str) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = SimpleFileOptions::default();

        // presentation.xml with embedded font list
        let pres_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:embeddedFontLst>
    <p:embeddedFont>
      <p:font typeface="{typeface}"/>
      <p:regular r:id="rId5"/>
    </p:embeddedFont>
  </p:embeddedFontLst>
</p:presentation>"#
        );
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
        build_docx_with_named_embedded_font(ttf_data, guid, "TestFont")
    }

    fn build_docx_with_named_embedded_font(
        ttf_data: &[u8],
        guid: &str,
        font_name: &str,
    ) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = SimpleFileOptions::default();

        // word/fontTable.xml
        let font_table_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="{font_name}">
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

        let font_data = extract_embedded_font_data(&zip_data, crate::config::Format::Pptx)
            .expect("should retain deobfuscated PPTX font bytes in memory");
        assert_eq!(font_data.faces.len(), 1);
        assert_eq!(font_data.faces[0].format, FontFileFormat::Ttf);
        assert_eq!(font_data.faces[0].data, original_ttf);

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

        let font_data = extract_embedded_font_data(&zip_data, crate::config::Format::Docx)
            .expect("should retain deobfuscated DOCX font bytes in memory");
        assert_eq!(font_data.faces.len(), 1);
        assert_eq!(font_data.faces[0].format, FontFileFormat::Ttf);
        assert_eq!(font_data.faces[0].data, original_ttf);

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

    /// Removes a scratch file or directory even when the test panics, so a
    /// regression cannot leave escaped fonts behind in the temp directory.
    struct RemoveOnDrop(std::path::PathBuf);

    impl Drop for RemoveOnDrop {
        fn drop(&mut self) {
            let _ = std::fs::remove_file(&self.0);
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }

    /// A scratch path no other test or run shares, so a leaked write is
    /// attributable to the case that made it.
    fn unique_scratch_path(label: &str) -> std::path::PathBuf {
        use std::time::{SystemTime, UNIX_EPOCH};

        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time should be valid")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "office2pdf-font-name-escape-{label}-{}-{nanos}",
            std::process::id()
        ))
    }

    /// Materialize a document whose embedded font declares `declared_name`,
    /// then require the face to land as a direct child of the font directory
    /// under a name that does not echo the document's text.
    fn assert_declared_name_cannot_choose_the_path(
        format: crate::config::Format,
        declared_name: &str,
    ) {
        let guid = "{7B19B49C-2336-4F82-AAD2-5D2BAE389560}";
        let original_ttf = make_fake_ttf(64);
        let zip_data = match format {
            crate::config::Format::Pptx => {
                build_pptx_with_named_embedded_font(&original_ttf, guid, declared_name)
            }
            crate::config::Format::Docx => {
                build_docx_with_named_embedded_font(&original_ttf, guid, declared_name)
            }
            crate::config::Format::Xlsx => unreachable!("XLSX carries no embedded fonts"),
        };

        let dir = extract_embedded_fonts(&zip_data, format)
            .unwrap_or_else(|| panic!("{format:?} face named {declared_name:?} was dropped"));
        let entries: Vec<_> = std::fs::read_dir(dir.path())
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .collect();
        assert_eq!(
            entries.len(),
            1,
            "{format:?} face named {declared_name:?} must be written inside {:?}",
            dir.path()
        );
        assert_eq!(entries[0].parent(), Some(dir.path()));
        assert_eq!(std::fs::read(&entries[0]).unwrap(), original_ttf);
        let file_name = entries[0].file_name().unwrap().to_string_lossy();
        for component in declared_name.split(['/', '\\']).filter(|c| c.len() > 2) {
            assert!(
                !file_name.contains(component),
                "file name {file_name:?} echoes document text {component:?}"
            );
        }
    }

    #[test]
    fn embedded_font_relative_escape_name_stays_in_font_dir() {
        for format in [crate::config::Format::Pptx, crate::config::Format::Docx] {
            let sentinel = unique_scratch_path("relative");
            let sentinel_name = sentinel.file_name().unwrap().to_string_lossy().into_owned();
            let leaked =
                RemoveOnDrop(sentinel.with_file_name(format!("{sentinel_name}-regular.ttf")));

            assert_declared_name_cannot_choose_the_path(format, &format!("../{sentinel_name}"));

            assert!(
                !leaked.0.exists(),
                "{format:?} font escaped to {:?}",
                leaked.0
            );
        }
    }

    #[test]
    fn embedded_font_absolute_name_stays_in_font_dir() {
        for format in [crate::config::Format::Pptx, crate::config::Format::Docx] {
            // The target directory exists, so only the fix can stop the write.
            let target = RemoveOnDrop(unique_scratch_path("absolute"));
            std::fs::create_dir(&target.0).unwrap();
            let declared = target.0.join("escaped");

            assert_declared_name_cannot_choose_the_path(format, &declared.to_string_lossy());

            let leaked: Vec<_> = std::fs::read_dir(&target.0).unwrap().collect();
            assert!(
                leaked.is_empty(),
                "{format:?} font escaped into {:?}",
                target.0
            );
        }
    }

    #[test]
    fn embedded_font_with_path_separator_in_its_name_is_still_materialized() {
        // Family names are free text; one that is not a valid file name must
        // neither escape nor be dropped.
        for format in [crate::config::Format::Pptx, crate::config::Format::Docx] {
            assert_declared_name_cannot_choose_the_path(format, "Brand/Display Sans");
        }
    }

    #[test]
    fn extract_embedded_fonts_no_fonts_returns_none() {
        // Build a minimal PPTX with no embedded fonts
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = ZipWriter::new(cursor);
        let options = SimpleFileOptions::default();

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
        let options = SimpleFileOptions::default();
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

    /// Require an sfnt table directory that only the right key recovers.
    ///
    /// The first 32 bytes cover the sfnt header and the first table record.
    /// A wrong key can still restore the magic and `numTables` (issue #2200),
    /// so check every record's tag, bounds and checksum instead.
    fn assert_valid_table_directory(data: &[u8]) {
        let be_u16 = |at: usize| u16::from_be_bytes([data[at], data[at + 1]]);
        let be_u32 =
            |at: usize| u32::from_be_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]]);
        assert_eq!(&data[0..4], &[0x00, 0x01, 0x00, 0x00], "sfnt magic");
        let num_tables = usize::from(be_u16(4));
        assert!(num_tables > 0, "a face carries tables");
        let mut previous_tag: Option<&[u8]> = None;
        for index in 0..num_tables {
            let record = 12 + 16 * index;
            let tag = &data[record..record + 4];
            assert!(
                tag.iter().all(|byte| (0x20..0x7F).contains(byte)),
                "table {index} tag {tag:02X?} must be printable ASCII"
            );
            assert!(
                previous_tag.is_none_or(|previous| previous < tag),
                "table tags must ascend, {previous_tag:?} then {tag:?}"
            );
            previous_tag = Some(tag);
            let checksum = be_u32(record + 4);
            let offset = be_u32(record + 8) as usize;
            let length = be_u32(record + 12) as usize;
            assert!(
                offset + length <= data.len(),
                "table {:?} must lie inside the face",
                String::from_utf8_lossy(tag)
            );
            // `head` stores a whole-file adjustment in its own checksum field.
            if tag != b"head" {
                let mut padded = data[offset..offset + length].to_vec();
                padded.resize(length.div_ceil(4) * 4, 0);
                let sum = padded
                    .as_chunks::<4>()
                    .0
                    .iter()
                    .map(|word| u32::from_be_bytes(*word))
                    .fold(0u32, u32::wrapping_add);
                assert_eq!(
                    sum,
                    checksum,
                    "table {:?} checksum",
                    String::from_utf8_lossy(tag)
                );
            }
        }
    }

    fn only_embedded_face(docx: &[u8]) -> Vec<u8> {
        let fonts = extract_embedded_font_data(docx, crate::config::Format::Docx)
            .expect("the fixture should carry an embedded font");
        let faces: Vec<&[u8]> = fonts.font_bytes().collect();
        assert_eq!(faces.len(), 1, "the fixture embeds one face");
        faces[0].to_vec()
    }

    fn family_name(data: &[u8]) -> Option<String> {
        let face = ttf_parser::Face::parse(data, 0).ok()?;
        face.names()
            .into_iter()
            .find(|name| name.name_id == ttf_parser::name_id::FAMILY && name.is_unicode())
            .and_then(|name| name.to_string())
    }

    /// Issue #2200: the Press Start 2P face that Word drew in the docxide-pdf
    /// case8 reference, kept byte-for-byte as that document obfuscated it.
    #[test]
    fn a_word_obfuscated_face_deobfuscates_to_its_table_directory() {
        let docx = include_bytes!(
            "../../../../tests/fixtures/docx/issue_2200_embedded_press_start_2p.docx"
        );
        let face = only_embedded_face(docx);

        assert_valid_table_directory(&face);
        assert_eq!(family_name(&face).as_deref(), Some("Press Start 2P"));
    }

    /// A second face under a different key, so the key rule is triangulated
    /// rather than fitted to one GUID.
    #[test]
    fn the_embedded_cjk_fixture_deobfuscates_to_its_table_directory() {
        let docx = include_bytes!("../../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
        let face = only_embedded_face(docx);

        assert_valid_table_directory(&face);
        assert_eq!(family_name(&face).as_deref(), Some("Noto Sans SC"));
    }
}
