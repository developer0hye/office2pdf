use std::collections::HashMap;

use typst::diag::FileResult;
use typst::foundations::{Bytes, Datetime};
use typst::syntax::{FileId, Source, VirtualPath};
use typst::text::Font;
use typst::utils::LazyHash;
use typst::{Library, LibraryExt, World};
use typst_kit::fonts::{FontSearcher, Fonts};

use crate::error::ConvertError;

use super::typst_gen::ImageAsset;

/// Compile Typst markup to PDF bytes.
pub fn compile_to_pdf(typst_source: &str, images: &[ImageAsset]) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new(typst_source, images);

    let warned = typst::compile::<typst::layout::PagedDocument>(&world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    let options = typst_pdf::PdfOptions::default();
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
    fn new(source_text: &str, images: &[ImageAsset]) -> Self {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(false);
        let font_data: Fonts = searcher.search();

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
        let result = compile_to_pdf("Hello, World!", &[]).unwrap();
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
        let result = compile_to_pdf(source, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_styled_text() {
        let source = r#"#text(weight: "bold", size: 16pt)[Bold Title]

#text(style: "italic")[Italic body text]

#underline[Underlined text]"#;
        let result = compile_to_pdf(source, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_colored_text() {
        let source = r#"#text(fill: rgb(255, 0, 0))[Red text]
#text(fill: rgb(0, 128, 255))[Blue text]"#;
        let result = compile_to_pdf(source, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_alignment() {
        let source = r#"#align(center)[Centered text]

#align(right)[Right-aligned text]"#;
        let result = compile_to_pdf(source, &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_invalid_source_returns_error() {
        // Invalid Typst source should produce a compilation error
        let result = compile_to_pdf("#invalid-func-that-does-not-exist()", &[]);
        assert!(result.is_err(), "Invalid source should produce an error");
    }

    #[test]
    fn test_compile_empty_source() {
        // Empty source should still produce valid PDF (empty page)
        let result = compile_to_pdf("", &[]).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_multiple_paragraphs() {
        let source = "First paragraph.\n\nSecond paragraph.\n\nThird paragraph.";
        let result = compile_to_pdf(source, &[]).unwrap();
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
    fn test_compile_with_embedded_image() {
        let png_data = make_test_png();
        let images = vec![ImageAsset {
            path: "img-0.png".to_string(),
            data: png_data,
        }];
        let source = r#"#image("img-0.png", width: 100pt)"#;
        let result = compile_to_pdf(source, &images).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }
}
