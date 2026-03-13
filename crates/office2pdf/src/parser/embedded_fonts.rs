//! Extract and deobfuscate embedded fonts from PPTX/DOCX archives.
//!
//! OOXML files can embed fonts as obfuscated binary data. This module extracts
//! them, deobfuscates using the GUID-based XOR scheme, and writes them to a
//! temporary directory for use during PDF compilation.

#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::parser::xml_util::get_attr_str;

#[cfg(test)]
#[path = "embedded_fonts_tests.rs"]
mod tests;

// =============================================================================
// Types
// =============================================================================

/// Font file format detected by magic bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FontFileFormat {
    Ttf,
    Otf,
    Ttc,
}

impl FontFileFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Ttf => "ttf",
            Self::Otf => "otf",
            Self::Ttc => "ttc",
        }
    }
}

/// A PPTX embedded font entry parsed from `presentation.xml`.
#[derive(Debug)]
struct PptxEmbeddedFontEntry {
    typeface: String,
    variants: Vec<FontVariantRef>,
}

/// A reference to a single font variant (regular, bold, italic, boldItalic).
#[derive(Debug)]
struct FontVariantRef {
    style: String,
    r_id: String,
}

/// A DOCX embedded font entry parsed from `word/fontTable.xml`.
#[derive(Debug)]
struct DocxEmbeddedFontEntry {
    font_name: String,
    variants: Vec<DocxFontVariantRef>,
}

/// A DOCX font variant reference including the fontKey for deobfuscation.
#[derive(Debug)]
struct DocxFontVariantRef {
    style: String,
    r_id: String,
    font_key: String,
}

/// Temporary directory containing extracted font files.
/// Cleaned up automatically when dropped.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) struct EmbeddedFontDir {
    path: PathBuf,
    font_count: usize,
}

#[cfg(not(target_arch = "wasm32"))]
impl EmbeddedFontDir {
    pub(crate) fn path(&self) -> &Path {
        &self.path
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.font_count == 0
    }
}

#[cfg(not(target_arch = "wasm32"))]
impl Drop for EmbeddedFontDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.path);
    }
}

// =============================================================================
// Public API
// =============================================================================

/// Extract embedded fonts from an OOXML archive.
///
/// Returns `None` if:
/// - The format doesn't support embedded fonts (XLSX)
/// - No embedded fonts are declared in the document
/// - The ZIP cannot be opened
/// - Extraction fails silently (best-effort)
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn extract_embedded_fonts(
    data: &[u8],
    format: crate::config::Format,
) -> Option<EmbeddedFontDir> {
    use crate::config::Format;

    let result = match format {
        Format::Pptx => extract_pptx_fonts(data),
        Format::Docx => extract_docx_fonts(data),
        Format::Xlsx => None,
    };

    if let Some(ref dir) = result {
        tracing::info!(
            font_count = dir.font_count,
            path = ?dir.path,
            "extracted embedded fonts from archive"
        );
    }

    result
}

// =============================================================================
// GUID parsing
// =============================================================================

/// Parse a GUID string into 16 bytes using OOXML mixed-endian encoding.
///
/// Input formats: `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` or without braces.
/// Encoding: first 4 bytes LE, next 2 LE, next 2 LE, remaining 8 BE.
fn parse_guid_to_bytes(guid: &str) -> Option<[u8; 16]> {
    let s = guid.trim_start_matches('{').trim_end_matches('}');
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 5 {
        return None;
    }

    let group1 = u32::from_str_radix(parts[0], 16).ok()?;
    let group2 = u16::from_str_radix(parts[1], 16).ok()?;
    let group3 = u16::from_str_radix(parts[2], 16).ok()?;
    let group4 = u16::from_str_radix(parts[3], 16).ok()?;
    let group5 = u64::from_str_radix(parts[4], 16).ok()?;

    if parts[4].len() != 12 {
        return None;
    }

    let mut key = [0u8; 16];
    // Group 1: LE
    key[0..4].copy_from_slice(&group1.to_le_bytes());
    // Group 2: LE
    key[4..6].copy_from_slice(&group2.to_le_bytes());
    // Group 3: LE
    key[6..8].copy_from_slice(&group3.to_le_bytes());
    // Group 4: BE (2 bytes)
    key[8..10].copy_from_slice(&group4.to_be_bytes());
    // Group 5: BE (6 bytes, from a u64 → take last 6 bytes)
    let g5_bytes = group5.to_be_bytes();
    key[10..16].copy_from_slice(&g5_bytes[2..8]);

    Some(key)
}

// =============================================================================
// Deobfuscation
// =============================================================================

/// XOR the first 32 bytes of font data with the 16-byte key (repeated twice).
fn deobfuscate_font_data(data: &mut [u8], key: &[u8; 16]) {
    let len = std::cmp::min(32, data.len());
    for i in 0..len {
        data[i] ^= key[i % 16];
    }
}

/// Detect font format from magic bytes at the start of the data.
fn detect_font_format(data: &[u8]) -> Option<FontFileFormat> {
    if data.len() < 4 {
        return None;
    }
    if data[0..4] == [0x00, 0x01, 0x00, 0x00] {
        Some(FontFileFormat::Ttf)
    } else if &data[0..4] == b"OTTO" {
        Some(FontFileFormat::Otf)
    } else if &data[0..4] == b"ttcf" {
        Some(FontFileFormat::Ttc)
    } else {
        None
    }
}

// =============================================================================
// PPTX extraction
// =============================================================================

/// Parse `<p:embeddedFontLst>` from `presentation.xml`.
fn parse_pptx_embedded_font_list(xml: &str) -> Vec<PptxEmbeddedFontEntry> {
    let mut reader = Reader::from_str(xml);
    let mut entries: Vec<PptxEmbeddedFontEntry> = Vec::new();
    let mut in_embedded_font = false;
    let mut current_typeface: Option<String> = None;
    let mut current_variants: Vec<FontVariantRef> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"embeddedFont" {
                    in_embedded_font = true;
                    current_typeface = None;
                    current_variants = Vec::new();
                }
            }
            Ok(Event::Empty(ref e)) if in_embedded_font => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"font" => {
                        current_typeface = get_attr_str(e, b"typeface");
                    }
                    b"regular" | b"bold" | b"italic" | b"boldItalic" => {
                        if let Some(r_id) = get_attr_str(e, b"r:id") {
                            let style = String::from_utf8_lossy(local_name.as_ref()).to_string();
                            current_variants.push(FontVariantRef { style, r_id });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"embeddedFont" && in_embedded_font {
                    in_embedded_font = false;
                    if let Some(typeface) = current_typeface.take()
                        && !current_variants.is_empty()
                    {
                        entries.push(PptxEmbeddedFontEntry {
                            typeface,
                            variants: std::mem::take(&mut current_variants),
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    entries
}

/// Extract GUID from a PPTX font file path like `ppt/fonts/{GUID}.fntdata`.
fn extract_guid_from_font_path(path: &str) -> Option<String> {
    let filename = path.rsplit('/').next()?;
    let stem = filename.strip_suffix(".fntdata")?;
    // Verify it looks like a GUID: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    if stem.starts_with('{') && stem.ends_with('}') && stem.len() == 38 {
        Some(stem.to_string())
    } else {
        None
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_pptx_fonts(data: &[u8]) -> Option<EmbeddedFontDir> {
    use std::io::Read;

    let mut archive = crate::parser::open_zip(data).ok()?;

    // Read presentation.xml
    let pres_xml = {
        let mut file = archive.by_name("ppt/presentation.xml").ok()?;
        let mut content = String::new();
        file.read_to_string(&mut content).ok()?;
        content
    };

    let font_entries = parse_pptx_embedded_font_list(&pres_xml);
    if font_entries.is_empty() {
        return None;
    }

    // Read presentation.xml.rels to resolve rId → target
    let rels_xml = {
        let mut file = archive.by_name("ppt/_rels/presentation.xml.rels").ok()?;
        let mut content = String::new();
        file.read_to_string(&mut content).ok()?;
        content
    };
    let rels = crate::parser::xml_util::parse_rels_id_target(&rels_xml);

    // Create temp dir
    let temp_dir = create_temp_font_dir("office2pdf-pptx-fonts")?;
    let mut font_count: usize = 0;

    for entry in &font_entries {
        for variant in &entry.variants {
            let Some(target) = rels.get(&variant.r_id) else {
                continue;
            };

            // Resolve path relative to ppt/
            let font_zip_path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("ppt/{target}")
            };

            // Extract GUID from filename
            let Some(guid_str) = extract_guid_from_font_path(&font_zip_path) else {
                continue;
            };
            let Some(key) = parse_guid_to_bytes(&guid_str) else {
                continue;
            };

            // Read font data from ZIP
            let mut font_data = Vec::new();
            {
                let mut file = match archive.by_name(&font_zip_path) {
                    Ok(f) => f,
                    Err(_) => continue,
                };
                if file.read_to_end(&mut font_data).is_err() {
                    continue;
                }
            }

            // Deobfuscate
            deobfuscate_font_data(&mut font_data, &key);

            // Detect format and write
            let ext = detect_font_format(&font_data)
                .map(|f| f.extension())
                .unwrap_or("ttf");

            let filename = format!("{}-{}.{}", entry.typeface, variant.style, ext);
            let out_path = temp_dir.join(&filename);
            if std::fs::write(&out_path, &font_data).is_ok() {
                font_count += 1;
            }
        }
    }

    if font_count == 0 {
        let _ = std::fs::remove_dir_all(&temp_dir);
        return None;
    }

    Some(EmbeddedFontDir {
        path: temp_dir,
        font_count,
    })
}

// =============================================================================
// DOCX extraction
// =============================================================================

/// Parse `<w:font>` elements with embedded font references from `word/fontTable.xml`.
fn parse_docx_embedded_font_entries(xml: &str) -> Vec<DocxEmbeddedFontEntry> {
    let mut reader = Reader::from_str(xml);
    let mut entries: Vec<DocxEmbeddedFontEntry> = Vec::new();
    let mut in_font = false;
    let mut current_name: Option<String> = None;
    let mut current_variants: Vec<DocxFontVariantRef> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"font" {
                    in_font = true;
                    current_name = get_attr_str(e, b"w:name").or_else(|| get_attr_str(e, b"name"));
                    current_variants = Vec::new();
                }
            }
            Ok(Event::Empty(ref e)) if in_font => {
                let local_name = e.local_name();
                let style = match local_name.as_ref() {
                    b"embedRegular" => Some("regular"),
                    b"embedBold" => Some("bold"),
                    b"embedItalic" => Some("italic"),
                    b"embedBoldItalic" => Some("boldItalic"),
                    _ => None,
                };
                if let Some(style) = style {
                    let r_id = get_attr_str(e, b"r:id");
                    let font_key =
                        get_attr_str(e, b"w:fontKey").or_else(|| get_attr_str(e, b"fontKey"));
                    if let (Some(r_id), Some(font_key)) = (r_id, font_key) {
                        current_variants.push(DocxFontVariantRef {
                            style: style.to_string(),
                            r_id,
                            font_key,
                        });
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"font" && in_font {
                    in_font = false;
                    if let Some(name) = current_name.take()
                        && !current_variants.is_empty()
                    {
                        entries.push(DocxEmbeddedFontEntry {
                            font_name: name,
                            variants: std::mem::take(&mut current_variants),
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    entries
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_docx_fonts(data: &[u8]) -> Option<EmbeddedFontDir> {
    use std::io::Read;

    let mut archive = crate::parser::open_zip(data).ok()?;

    // Read word/fontTable.xml
    let font_table_xml = {
        let mut file = archive.by_name("word/fontTable.xml").ok()?;
        let mut content = String::new();
        file.read_to_string(&mut content).ok()?;
        content
    };

    let font_entries = parse_docx_embedded_font_entries(&font_table_xml);
    if font_entries.is_empty() {
        return None;
    }

    // Read word/_rels/fontTable.xml.rels
    let rels_xml = {
        let mut file = archive.by_name("word/_rels/fontTable.xml.rels").ok()?;
        let mut content = String::new();
        file.read_to_string(&mut content).ok()?;
        content
    };
    let rels = crate::parser::xml_util::parse_rels_id_target(&rels_xml);

    let temp_dir = create_temp_font_dir("office2pdf-docx-fonts")?;
    let mut font_count: usize = 0;

    for entry in &font_entries {
        for variant in &entry.variants {
            let Some(target) = rels.get(&variant.r_id) else {
                continue;
            };

            // Resolve path relative to word/
            let font_zip_path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("word/{target}")
            };

            // Parse GUID from fontKey attribute
            let Some(key) = parse_guid_to_bytes(&variant.font_key) else {
                continue;
            };

            // Read font data from ZIP
            let mut font_data = Vec::new();
            {
                let mut file = match archive.by_name(&font_zip_path) {
                    Ok(f) => f,
                    Err(_) => continue,
                };
                if file.read_to_end(&mut font_data).is_err() {
                    continue;
                }
            }

            // Deobfuscate
            deobfuscate_font_data(&mut font_data, &key);

            // Detect format and write
            let ext = detect_font_format(&font_data)
                .map(|f| f.extension())
                .unwrap_or("ttf");

            let filename = format!("{}-{}.{}", entry.font_name, variant.style, ext);
            let out_path = temp_dir.join(&filename);
            if std::fs::write(&out_path, &font_data).is_ok() {
                font_count += 1;
            }
        }
    }

    if font_count == 0 {
        let _ = std::fs::remove_dir_all(&temp_dir);
        return None;
    }

    Some(EmbeddedFontDir {
        path: temp_dir,
        font_count,
    })
}

// =============================================================================
// Helpers
// =============================================================================

#[cfg(not(target_arch = "wasm32"))]
fn create_temp_font_dir(prefix: &str) -> Option<PathBuf> {
    use std::time::{SystemTime, UNIX_EPOCH};

    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time should be valid")
        .as_nanos();
    let path = std::env::temp_dir().join(format!("{prefix}-{unique}"));
    std::fs::create_dir_all(&path).ok()?;
    Some(path)
}
