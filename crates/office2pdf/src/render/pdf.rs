use std::collections::HashMap;
use std::path::PathBuf;

use typst::diag::FileResult;
use typst::foundations::{Bytes, Datetime};
use typst::syntax::{FileId, Source, VirtualPath};
use typst::text::Font;
use typst::utils::LazyHash;
use typst::{Library, LibraryExt, World};
use typst_kit::fonts::{FontSearcher, Fonts};

use crate::config::PdfStandard;
use crate::error::ConvertError;

use super::typst_gen::ImageAsset;

/// Compile Typst markup to PDF bytes.
///
/// When `pdf_standard` is `Some`, the output PDF will conform to the
/// specified standard (e.g., PDF/A-2b for archival).
/// When `font_paths` is non-empty, those directories are searched for
/// additional fonts (highest priority).
pub fn compile_to_pdf(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    font_paths: &[PathBuf],
) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new(typst_source, images, font_paths);

    let warned = typst::compile::<typst::layout::PagedDocument>(&world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    let standards = match pdf_standard {
        Some(PdfStandard::PdfA2b) => typst_pdf::PdfStandards::new(&[typst_pdf::PdfStandard::A_2b])
            .map_err(|e| ConvertError::Render(format!("PDF standard configuration error: {e}")))?,
        None => typst_pdf::PdfStandards::default(),
    };

    // PDF/A requires a document creation timestamp
    let timestamp = if pdf_standard.is_some() {
        let now = Datetime::from_ymd_hms(2024, 1, 1, 0, 0, 0).unwrap();
        Some(typst_pdf::Timestamp::new_utc(now))
    } else {
        None
    };

    let options = typst_pdf::PdfOptions {
        standards,
        timestamp,
        ..Default::default()
    };
    typst_pdf::pdf(&document, &options).map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("PDF export failed: {}", messages.join("; ")))
    })
}

/// Minimal World implementation providing Typst compiler with source, fonts, and images.
struct MinimalWorld {
    library: LazyHash<Library>,
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<typst_kit::fonts::FontSlot>,
    source: Source,
    images: HashMap<String, Vec<u8>>,
}

impl MinimalWorld {
    fn new(source_text: &str, images: &[ImageAsset], font_paths: &[PathBuf]) -> Self {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(true);
        let font_data: Fonts = if font_paths.is_empty() {
            searcher.search()
        } else {
            searcher.search_with(font_paths.iter().map(|p| p.as_path()))
        };

        let main_id = FileId::new(None, VirtualPath::new("main.typ"));
        let source = Source::new(main_id, source_text.to_string());

        let image_map: HashMap<String, Vec<u8>> = images
            .iter()
            .map(|a| (a.path.clone(), a.data.clone()))
            .collect();

        Self {
            library: LazyHash::new(Library::default()),
            book: LazyHash::new(font_data.book),
            fonts: font_data.fonts,
            source,
            images: image_map,
        }
    }
}

impl World for MinimalWorld {
    fn library(&self) -> &LazyHash<Library> {
        &self.library
    }

    fn book(&self) -> &LazyHash<typst::text::FontBook> {
        &self.book
    }

    fn main(&self) -> FileId {
        self.source.id()
    }

    fn source(&self, id: FileId) -> FileResult<Source> {
        if id == self.source.id() {
            Ok(self.source.clone())
        } else {
            Err(typst::diag::FileError::NotFound(
                id.vpath().as_rootless_path().into(),
            ))
        }
    }

    fn file(&self, id: FileId) -> FileResult<Bytes> {
        if id == self.source.id() {
            Ok(Bytes::new(self.source.text().as_bytes().to_vec()))
        } else {
            // Check if it's an embedded image file
            let path = id.vpath().as_rootless_path().to_string_lossy();
            if let Some(data) = self.images.get(path.as_ref()) {
                Ok(Bytes::new(data.clone()))
            } else {
                Err(typst::diag::FileError::NotFound(
                    id.vpath().as_rootless_path().into(),
                ))
            }
        }
    }

    fn font(&self, index: usize) -> Option<Font> {
        self.fonts.get(index).and_then(|slot| slot.get())
    }

    fn today(&self, _offset: Option<i64>) -> Option<Datetime> {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compile_simple_text() {
        let result = compile_to_pdf("Hello, World!", &[], None, &[]).unwrap();
        assert!(!result.is_empty(), "PDF bytes should not be empty");
        assert!(
            result.starts_with(b"%PDF"),
            "PDF should start with %PDF magic bytes"
        );
    }

    #[test]
    fn test_compile_with_page_setup() {
        let source = r#"#set page(width: 612pt, height: 792pt)
Hello from a US Letter page."#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_styled_text() {
        let source = r#"#text(weight: "bold", size: 16pt)[Bold Title]

#text(style: "italic")[Italic body text]

#underline[Underlined text]"#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_colored_text() {
        let source = r#"#text(fill: rgb(255, 0, 0))[Red text]
#text(fill: rgb(0, 128, 255))[Blue text]"#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_alignment() {
        let source = r#"#align(center)[Centered text]

#align(right)[Right-aligned text]"#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_invalid_source_returns_error() {
        // Invalid Typst source should produce a compilation error
        let result = compile_to_pdf("#invalid-func-that-does-not-exist()", &[], None, &[]);
        assert!(result.is_err(), "Invalid source should produce an error");
    }

    #[test]
    fn test_compile_empty_source() {
        // Empty source should still produce valid PDF (empty page)
        let result = compile_to_pdf("", &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_multiple_paragraphs() {
        let source = "First paragraph.\n\nSecond paragraph.\n\nThird paragraph.";
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    /// Compute CRC32 over PNG chunk type + data.
    fn png_crc32(chunk_type: &[u8], data: &[u8]) -> u32 {
        let mut crc: u32 = 0xFFFF_FFFF;
        for &byte in chunk_type.iter().chain(data.iter()) {
            crc ^= byte as u32;
            for _ in 0..8 {
                if crc & 1 != 0 {
                    crc = (crc >> 1) ^ 0xEDB8_8320;
                } else {
                    crc >>= 1;
                }
            }
        }
        crc ^ 0xFFFF_FFFF
    }

    /// Build a minimal valid 1x1 red PNG with correct CRC checksums.
    fn make_test_png() -> Vec<u8> {
        let mut png = Vec::new();
        // PNG signature
        png.extend_from_slice(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        // IHDR: 1x1, 8-bit RGB
        let ihdr_data: [u8; 13] = [
            0x00, 0x00, 0x00, 0x01, // width=1
            0x00, 0x00, 0x00, 0x01, // height=1
            0x08, // bit depth=8
            0x02, // color type=RGB
            0x00, 0x00, 0x00, // compression, filter, interlace
        ];
        let ihdr_type = b"IHDR";
        png.extend_from_slice(&(ihdr_data.len() as u32).to_be_bytes());
        png.extend_from_slice(ihdr_type);
        png.extend_from_slice(&ihdr_data);
        png.extend_from_slice(&png_crc32(ihdr_type, &ihdr_data).to_be_bytes());

        // IDAT: zlib-compressed row [filter=0, R=255, G=0, B=0]
        let idat_data: [u8; 15] = [
            0x78, 0x01, // zlib header
            0x01, // BFINAL=1, BTYPE=00 (stored)
            0x04, 0x00, 0xFB, 0xFF, // LEN=4, NLEN
            0x00, 0xFF, 0x00, 0x00, // filter + RGB
            0x03, 0x01, 0x01, 0x00, // adler32
        ];
        let idat_type = b"IDAT";
        png.extend_from_slice(&(idat_data.len() as u32).to_be_bytes());
        png.extend_from_slice(idat_type);
        png.extend_from_slice(&idat_data);
        png.extend_from_slice(&png_crc32(idat_type, &idat_data).to_be_bytes());

        // IEND
        let iend_type = b"IEND";
        png.extend_from_slice(&0u32.to_be_bytes());
        png.extend_from_slice(iend_type);
        png.extend_from_slice(&png_crc32(iend_type, &[]).to_be_bytes());

        png
    }

    #[test]
    fn test_embedded_fonts_are_available() {
        // MinimalWorld should always have embedded fallback fonts available
        // (Libertinus Serif, New Computer Modern, DejaVu Sans Mono)
        let world = MinimalWorld::new("", &[], &[]);
        assert!(
            !world.fonts.is_empty(),
            "MinimalWorld should have at least the embedded fallback fonts"
        );
    }

    #[test]
    fn test_system_fonts_enabled() {
        // With system font discovery enabled, on typical systems we should have
        // more fonts than just the embedded set. On minimal systems, we at least
        // have the embedded fonts.
        let world = MinimalWorld::new("", &[], &[]);
        let embedded_only_count = {
            let mut s = FontSearcher::new();
            s.include_system_fonts(false);
            s.search().fonts.len()
        };
        // At minimum, we should have the embedded fonts
        assert!(
            world.fonts.len() >= embedded_only_count,
            "System font discovery should not reduce available fonts: total {} vs embedded-only {}",
            world.fonts.len(),
            embedded_only_count
        );
    }

    #[test]
    fn test_compile_with_system_font_name() {
        // A document specifying a common system font should compile successfully.
        // Typst falls back to embedded fonts if the named font isn't available,
        // so this test always succeeds — but with system fonts enabled, the
        // named font will be used if present on the system.
        let source = r#"#set text(font: "Arial")
Hello with a system font."#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_embedded_fonts_still_available_as_fallback() {
        // Embedded fonts (Libertinus Serif) must still be available even with
        // system font discovery enabled.
        let source = r#"#set text(font: "Libertinus Serif")
Text in Libertinus Serif."#;
        let result = compile_to_pdf(source, &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_pdfa2b_produces_valid_pdf() {
        let result = compile_to_pdf(
            "Hello PDF/A!",
            &[],
            Some(crate::config::PdfStandard::PdfA2b),
            &[],
        )
        .unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_pdfa2b_contains_xmp_metadata() {
        let result = compile_to_pdf(
            "PDF/A metadata test",
            &[],
            Some(crate::config::PdfStandard::PdfA2b),
            &[],
        )
        .unwrap();
        // PDF/A-2b requires XMP metadata with pdfaid namespace
        let pdf_str = String::from_utf8_lossy(&result);
        assert!(
            pdf_str.contains("pdfaid") || pdf_str.contains("PDF/A"),
            "PDF/A output should contain PDF/A identification metadata"
        );
    }

    #[test]
    fn test_compile_default_no_pdfa_metadata() {
        let result = compile_to_pdf("Regular PDF", &[], None, &[]).unwrap();
        let pdf_str = String::from_utf8_lossy(&result);
        // A regular PDF should not have pdfaid conformance metadata
        assert!(
            !pdf_str.contains("pdfaid:conformance"),
            "Regular PDF should not contain PDF/A conformance metadata"
        );
    }

    #[test]
    fn test_compile_with_font_paths_empty() {
        // Empty font paths should work the same as without
        let result = compile_to_pdf("Hello!", &[], None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_with_nonexistent_font_path() {
        // Non-existent font path should not crash — FontSearcher skips invalid dirs
        let paths = vec![PathBuf::from("/nonexistent/font/path")];
        let result = compile_to_pdf("Hello!", &[], None, &paths).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_with_embedded_image() {
        let png_data = make_test_png();
        let images = vec![ImageAsset {
            path: "img-0.png".to_string(),
            data: png_data,
        }];
        let source = r#"#image("img-0.png", width: 100pt)"#;
        let result = compile_to_pdf(source, &images, None, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }
}
