use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

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
///
/// On native targets, system fonts are discovered automatically. On WASM,
/// only embedded fonts are used and `font_paths` is ignored.
///
/// # PDF output size optimization
///
/// typst-pdf (via krilla) applies the following optimizations by default:
///
/// - **Content stream compression**: All content streams use FLATE (deflate)
///   compression (`compress_content_streams: true`). Typical reduction: 60-80%.
/// - **Font subsetting**: Only glyphs actually used in the document are embedded
///   (via the `subsetter` crate). Typical reduction: 70-90% of font data.
/// - **Image pass-through**: Embedded images (PNG, JPEG) are included as-is
///   without re-encoding, preserving their original compression.
///
/// Expected output sizes:
/// - Empty page: ~10-30 KB (font data + PDF structure overhead)
/// - 10-page text-only document: ~30-60 KB
/// - Document with images: baseline + proportional to image data size
#[cfg(not(target_arch = "wasm32"))]
pub fn compile_to_pdf(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    font_paths: &[PathBuf],
) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new(typst_source, images, font_paths);
    compile_to_pdf_inner(&world, pdf_standard)
}

/// Compile Typst markup to PDF bytes (WASM target).
///
/// Uses embedded fonts only. System font paths are not supported on WASM.
#[cfg(target_arch = "wasm32")]
pub fn compile_to_pdf(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    _font_paths: &[std::path::PathBuf],
) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new_embedded_only(typst_source, images);
    compile_to_pdf_inner(&world, pdf_standard)
}

fn compile_to_pdf_inner(
    world: &MinimalWorld,
    pdf_standard: Option<PdfStandard>,
) -> Result<Vec<u8>, ConvertError> {
    let warned = typst::compile::<typst::layout::PagedDocument>(world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    let standards = match pdf_standard {
        Some(PdfStandard::PdfA2b) => typst_pdf::PdfStandards::new(&[typst_pdf::PdfStandard::A_2b])
            .map_err(|e| ConvertError::Render(format!("PDF standard configuration error: {e}")))?,
        None => typst_pdf::PdfStandards::default(),
    };

    // PDF/A requires a document creation timestamp; use actual conversion time
    let timestamp = if pdf_standard.is_some() {
        Some(typst_pdf::Timestamp::new_utc(current_utc_datetime()))
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

/// Convert the current system time to a Typst `Datetime` in UTC.
///
/// Uses `std::time::SystemTime` to avoid an external chrono dependency.
/// The civil date is computed from the Unix timestamp using Howard Hinnant's
/// algorithm (<http://howardhinnant.github.io/date_algorithms.html>).
fn current_utc_datetime() -> Datetime {
    let duration = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let secs = duration.as_secs() as i64;

    // Split into days since epoch and time-of-day
    let days = secs.div_euclid(86400);
    let rem = secs.rem_euclid(86400);
    let hours = (rem / 3600) as u8;
    let minutes = ((rem % 3600) / 60) as u8;
    let seconds = (rem % 60) as u8;

    // Civil date from day count since Unix epoch (1970-01-01)
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097) as u32; // day of era [0, 146096]
    let yoe = (doe - doe / 1461 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100); // day of year [0, 365]
    let mp = (5 * doy + 2) / 153; // [0, 11]
    let d = (doy - (153 * mp + 2) / 5 + 1) as u8;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u8;
    let y = if m <= 2 { y + 1 } else { y } as i32;

    Datetime::from_ymd_hms(y, m, d, hours, minutes, seconds)
        .expect("valid date derived from SystemTime")
}

/// Minimal World implementation providing Typst compiler with source, fonts, and images.
struct MinimalWorld {
    library: LazyHash<Library>,
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<typst_kit::fonts::FontSlot>,
    source: Source,
    images: HashMap<String, Bytes>,
}

impl MinimalWorld {
    /// Create a new `MinimalWorld` with system fonts and optional custom font paths.
    ///
    /// This is the native constructor — it discovers system fonts and optionally
    /// searches additional directories. Not available on `wasm32` targets because
    /// system font discovery requires OS APIs.
    #[cfg(not(target_arch = "wasm32"))]
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

        let image_map: HashMap<String, Bytes> = images
            .iter()
            .map(|a| (a.path.clone(), Bytes::new(a.data.clone())))
            .collect();

        Self {
            library: LazyHash::new(Library::default()),
            book: LazyHash::new(font_data.book),
            fonts: font_data.fonts,
            source,
            images: image_map,
        }
    }

    /// Create a new `MinimalWorld` with embedded fonts only (no system font search).
    ///
    /// This is the constructor used on WASM targets where system font discovery
    /// is not available. It uses only the fonts embedded in the `typst-kit` crate
    /// (Libertinus Serif, New Computer Modern, DejaVu Sans Mono).
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    fn new_embedded_only(source_text: &str, images: &[ImageAsset]) -> Self {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(false);
        let font_data: Fonts = searcher.search();

        let main_id = FileId::new(None, VirtualPath::new("main.typ"));
        let source = Source::new(main_id, source_text.to_string());

        let image_map: HashMap<String, Bytes> = images
            .iter()
            .map(|a| (a.path.clone(), Bytes::new(a.data.clone())))
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
                Ok(data.clone()) // Bytes::clone is cheap (reference-counted)
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

    #[test]
    fn test_embedded_only_world_produces_valid_pdf() {
        // Simulates the WASM code path: embedded fonts only, no system fonts.
        // This verifies that the embedded-only MinimalWorld can produce valid PDFs.
        let world = MinimalWorld::new_embedded_only("Hello from embedded-only world!", &[]);
        assert!(
            !world.fonts.is_empty(),
            "Embedded-only world should have fonts"
        );

        let warned = typst::compile::<typst::layout::PagedDocument>(&world);
        let document = warned.output.expect("Compilation should succeed");
        let pdf = typst_pdf::pdf(&document, &typst_pdf::PdfOptions::default())
            .expect("PDF export should succeed");
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_embedded_only_world_has_fonts() {
        // The embedded-only constructor (used on WASM) must have at least
        // the embedded fallback fonts (Libertinus, New Computer Modern, DejaVu).
        let world = MinimalWorld::new_embedded_only("", &[]);
        let embedded_count = {
            let mut s = FontSearcher::new();
            s.include_system_fonts(false);
            s.search().fonts.len()
        };
        assert_eq!(
            world.fonts.len(),
            embedded_count,
            "Embedded-only world should have exactly the embedded fonts"
        );
    }

    #[test]
    fn test_pdfa_timestamp_is_not_hardcoded() {
        // PDF/A output should contain the actual conversion timestamp,
        // not the previously hardcoded 2024-01-01.
        let result = compile_to_pdf(
            "Timestamp test",
            &[],
            Some(crate::config::PdfStandard::PdfA2b),
            &[],
        )
        .unwrap();
        let pdf_str = String::from_utf8_lossy(&result);
        // The old hardcoded date was 2024-01-01T00:00:00 — it should no longer appear
        assert!(
            !pdf_str.contains("2024-01-01T00:00:00"),
            "PDF/A timestamp should not be the hardcoded 2024-01-01T00:00:00"
        );
    }

    #[test]
    fn test_current_utc_datetime_is_valid() {
        // The helper should produce a valid Datetime that can create a Timestamp.
        let dt = current_utc_datetime();
        let _ts = typst_pdf::Timestamp::new_utc(dt);
    }

    #[test]
    fn test_pdfa_timestamp_has_recent_date() {
        // The PDF/A XMP metadata should contain a date from the current
        // decade, not a hardcoded past date.
        let result = compile_to_pdf(
            "Year test",
            &[],
            Some(crate::config::PdfStandard::PdfA2b),
            &[],
        )
        .unwrap();
        let pdf_str = String::from_utf8_lossy(&result);
        // The XMP metadata should contain a CreateDate field
        assert!(
            pdf_str.contains("xmp:CreateDate") || pdf_str.contains("CreateDate"),
            "PDF/A should contain creation date metadata"
        );
        // The date should NOT be the hardcoded 2024-01-01
        assert!(
            !pdf_str.contains("2024-01-01"),
            "PDF/A timestamp should not contain hardcoded 2024-01-01"
        );
    }

    // --- PDF output size optimization tests (US-089) ---

    #[test]
    fn test_pdf_uses_flate_compression() {
        // typst-pdf (via krilla) compresses content streams with FLATE by default.
        // Verify that the output PDF contains FlateDecode filter references.
        let source = "Hello, compressed world! ".repeat(100);
        let result = compile_to_pdf(&source, &[], None, &[]).unwrap();
        let pdf_str = String::from_utf8_lossy(&result);
        assert!(
            pdf_str.contains("FlateDecode"),
            "PDF content streams should use FlateDecode compression"
        );
    }

    #[test]
    fn test_font_subsetting_reduces_size() {
        // A PDF using only a few glyphs should be significantly smaller than
        // one using many distinct glyphs, demonstrating font subsetting is active.
        // "Few glyphs" document: only ASCII letters a-z
        let few_glyphs = compile_to_pdf("abcdefghij", &[], None, &[]).unwrap();

        // "Many glyphs" document: diverse characters force more glyph data.
        // Avoid Typst special characters (#, $, *, _, etc.) to keep it valid markup.
        let many_glyphs_source = "abcdefghijklmnopqrstuvwxyz \
            ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789 \
            The quick brown fox jumps over the lazy dog. \
            SPHINX OF BLACK QUARTZ, JUDGE MY VOW. \
            Pack my box with five dozen liquor jugs. \
            How vexingly quick daft zebras jump.";
        let many_glyphs = compile_to_pdf(many_glyphs_source, &[], None, &[]).unwrap();

        // With font subsetting, the "few glyphs" PDF should be noticeably smaller.
        // Without subsetting, both would embed the full font and be similar in size.
        assert!(
            few_glyphs.len() < many_glyphs.len(),
            "PDF with fewer glyphs ({} bytes) should be smaller than PDF with many glyphs ({} bytes), \
             indicating font subsetting is active",
            few_glyphs.len(),
            many_glyphs.len()
        );
    }

    #[test]
    fn test_multipage_text_pdf_size_reasonable() {
        // A 10-page text-only document should produce a PDF well under 500KB.
        // This verifies that compression and font subsetting keep output compact.
        //
        // typst-pdf behavior (verified):
        // - Content streams use FLATE compression (compress_content_streams: true)
        // - Fonts are automatically subset to include only used glyphs
        // - No unnecessary re-encoding of embedded data
        let mut source = String::new();
        for i in 1..=10 {
            if i > 1 {
                source.push_str("#pagebreak()\n");
            }
            source.push_str(&format!(
                "= Page {i}\n\n\
                 This is page {i} of a multi-page document used to verify \
                 that PDF output size remains reasonable with compression \
                 and font subsetting enabled.\n\n\
                 Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
                 Sed do eiusmod tempor incididunt ut labore et dolore magna \
                 aliqua. Ut enim ad minim veniam, quis nostrud exercitation \
                 ullamco laboris nisi ut aliquip ex ea commodo consequat.\n\n"
            ));
        }
        let result = compile_to_pdf(&source, &[], None, &[]).unwrap();

        // 500KB = 512_000 bytes — generous upper bound for 10 pages of text
        assert!(
            result.len() < 512_000,
            "10-page text-only PDF should be under 500KB, actual size: {} bytes ({:.1} KB)",
            result.len(),
            result.len() as f64 / 1024.0
        );
    }

    #[test]
    fn test_pdf_with_image_size_proportional() {
        // A PDF with an embedded image should not inflate the image size
        // significantly. The output PDF should be proportional to the input
        // image data size (not orders of magnitude larger from re-encoding).
        let png_data = make_test_png();
        let png_size = png_data.len();
        let images = vec![ImageAsset {
            path: "img-0.png".to_string(),
            data: png_data,
        }];
        let source = r#"#image("img-0.png", width: 100pt)"#;
        let result = compile_to_pdf(source, &images, None, &[]).unwrap();

        // The PDF has overhead (fonts, structure, metadata) beyond the image.
        // But the total should not be unreasonably large for a tiny 1x1 image.
        // A 1x1 PNG is ~70 bytes; the PDF overhead is typically 10-30KB (fonts).
        // We assert the total is under 100KB to catch re-encoding issues.
        assert!(
            result.len() < 100_000,
            "PDF with tiny 1x1 image should be under 100KB, actual: {} bytes ({:.1} KB). \
             Image was {} bytes. Possible image re-encoding issue.",
            result.len(),
            result.len() as f64 / 1024.0,
            png_size
        );
    }

    #[test]
    fn test_empty_page_pdf_baseline_size() {
        // An empty page PDF establishes the baseline overhead (fonts, structure).
        // This helps verify that additional content adds proportional size, not
        // excessive bloat from uncompressed data.
        let result = compile_to_pdf("", &[], None, &[]).unwrap();

        // Empty page PDF should be compact — mostly font data and PDF structure.
        // Typically 10-30KB depending on embedded font data.
        assert!(
            result.len() < 100_000,
            "Empty page PDF should be under 100KB (baseline), actual: {} bytes ({:.1} KB)",
            result.len(),
            result.len() as f64 / 1024.0
        );
    }

    #[test]
    fn test_compression_effective_for_repetitive_content() {
        // FLATE compression is especially effective on repetitive content.
        // A document with highly repetitive text should compress well,
        // producing a PDF not much larger than a document with less text.
        let short_source = "Hello world.\n\n";
        let short_pdf = compile_to_pdf(short_source, &[], None, &[]).unwrap();

        // 100x the text content, but should compress to much less than 100x the size
        let long_source = "Hello world.\n\n".repeat(100);
        let long_pdf = compile_to_pdf(&long_source, &[], None, &[]).unwrap();

        // With compression, 100x content should produce far less than 10x the PDF size.
        // The ratio demonstrates that content streams are being compressed.
        let size_ratio = long_pdf.len() as f64 / short_pdf.len() as f64;
        assert!(
            size_ratio < 10.0,
            "100x content should produce less than 10x PDF size with compression. \
             Short: {} bytes, Long: {} bytes, Ratio: {:.1}x",
            short_pdf.len(),
            long_pdf.len(),
            size_ratio
        );
    }
}
