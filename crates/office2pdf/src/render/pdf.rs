use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{SystemTime, UNIX_EPOCH};

use typst::diag::FileResult;
use typst::foundations::{Bytes, Datetime};
use typst::syntax::{FileId, Source, VirtualPath};
use typst::text::Font;
use typst::utils::LazyHash;
use typst::{Library, LibraryExt, World};
use typst_kit::fonts::FontSearcher;

use crate::config::PdfStandard;
use crate::error::ConvertError;

use super::typst_gen::ImageAsset;

/// Cached font data (book + font slots). Font discovery is expensive because
/// it scans the filesystem; the result doesn't change during the process
/// lifetime, so we cache it in a global `OnceLock`.
struct CachedFontData {
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<typst_kit::fonts::FontSlot>,
}

/// Cached system fonts (with system font search). Used when no custom
/// font paths are provided, which is the common case.
#[cfg(not(target_arch = "wasm32"))]
static SYSTEM_FONTS: OnceLock<CachedFontData> = OnceLock::new();

/// Cached font data for resolved extra font path sets.
#[cfg(not(target_arch = "wasm32"))]
static EXTRA_FONT_PATHS_CACHE: OnceLock<Mutex<HashMap<Vec<PathBuf>, Arc<CachedFontData>>>> =
    OnceLock::new();

/// Cached embedded-only fonts (no system font search). Used on WASM
/// or when system fonts are not needed.
static EMBEDDED_FONTS: OnceLock<CachedFontData> = OnceLock::new();

/// Get or initialize cached system fonts (with system font discovery).
#[cfg(not(target_arch = "wasm32"))]
fn get_system_fonts() -> &'static CachedFontData {
    SYSTEM_FONTS.get_or_init(|| {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(true);
        let font_data = searcher.search();
        CachedFontData {
            book: LazyHash::new(font_data.book),
            fonts: font_data.fonts,
        }
    })
}

/// Get or initialize cached fonts for a resolved extra font path set.
#[cfg(not(target_arch = "wasm32"))]
fn get_fonts_for_extra_paths(font_paths: &[PathBuf]) -> Arc<CachedFontData> {
    let cache = EXTRA_FONT_PATHS_CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    {
        let cache_guard = cache
            .lock()
            .expect("font cache mutex should not be poisoned");
        if let Some(cached) = cache_guard.get(font_paths) {
            return Arc::clone(cached);
        }
    }

    let mut searcher = FontSearcher::new();
    searcher.include_system_fonts(true);
    let font_data = searcher.search_with(font_paths.iter().map(|path| path.as_path()));
    let cached = Arc::new(CachedFontData {
        book: LazyHash::new(font_data.book),
        fonts: font_data.fonts,
    });

    let mut cache_guard = cache
        .lock()
        .expect("font cache mutex should not be poisoned");
    let entry = cache_guard
        .entry(font_paths.to_vec())
        .or_insert_with(|| Arc::clone(&cached));
    Arc::clone(entry)
}

/// Get or initialize cached embedded-only fonts.
fn get_embedded_fonts() -> &'static CachedFontData {
    EMBEDDED_FONTS.get_or_init(|| {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(false);
        let font_data = searcher.search();
        CachedFontData {
            book: LazyHash::new(font_data.book),
            fonts: font_data.fonts,
        }
    })
}

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
    tagged: bool,
    pdf_ua: bool,
) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new(typst_source, images, font_paths);
    compile_to_pdf_inner(&world, pdf_standard, tagged, pdf_ua)
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
    tagged: bool,
    pdf_ua: bool,
) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new_embedded_only(typst_source, images);
    compile_to_pdf_inner(&world, pdf_standard, tagged, pdf_ua)
}

fn compile_to_pdf_inner(
    world: &MinimalWorld,
    pdf_standard: Option<PdfStandard>,
    tagged: bool,
    pdf_ua: bool,
) -> Result<Vec<u8>, ConvertError> {
    let warned = typst::compile::<typst::layout::PagedDocument>(world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    // Build PDF standards list
    let mut pdf_standards = Vec::new();
    if let Some(PdfStandard::PdfA2b) = pdf_standard {
        pdf_standards.push(typst_pdf::PdfStandard::A_2b);
    }
    if pdf_ua {
        pdf_standards.push(typst_pdf::PdfStandard::Ua_1);
    }
    let standards = if pdf_standards.is_empty() {
        typst_pdf::PdfStandards::default()
    } else {
        typst_pdf::PdfStandards::new(&pdf_standards)
            .map_err(|e| ConvertError::Render(format!("PDF standard configuration error: {e}")))?
    };

    // PDF/A and PDF/UA require a document creation timestamp
    let needs_timestamp = pdf_standard.is_some() || pdf_ua;
    let timestamp = if needs_timestamp {
        Some(typst_pdf::Timestamp::new_utc(current_utc_datetime()))
    } else {
        None
    };

    // Enable tagging when explicitly requested or when PDF/UA requires it
    let enable_tagged = tagged || pdf_ua;

    let options = typst_pdf::PdfOptions {
        standards,
        timestamp,
        tagged: enable_tagged,
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

/// Font data source: either a static reference to cached fonts or owned
/// data for custom font path searches.
enum FontSource {
    /// Reference to globally cached font data (common case).
    Cached(&'static CachedFontData),
    /// Shared cached font data for resolved extra font paths.
    Shared(Arc<CachedFontData>),
}

impl FontSource {
    fn book(&self) -> &LazyHash<typst::text::FontBook> {
        match self {
            Self::Cached(d) => &d.book,
            Self::Shared(d) => &d.book,
        }
    }

    fn fonts(&self) -> &[typst_kit::fonts::FontSlot] {
        match self {
            Self::Cached(d) => &d.fonts,
            Self::Shared(d) => &d.fonts,
        }
    }
}

/// Minimal World implementation providing Typst compiler with source, fonts, and images.
struct MinimalWorld {
    library: LazyHash<Library>,
    font_source: FontSource,
    source: Source,
    images: HashMap<String, Bytes>,
}

impl MinimalWorld {
    /// Create a new `MinimalWorld` with system fonts and optional custom font paths.
    ///
    /// When `font_paths` is empty (the common case), system fonts are loaded from
    /// a process-wide cache, avoiding expensive filesystem scanning on repeated calls.
    /// Resolved extra font path sets are also cached by path list.
    #[cfg(not(target_arch = "wasm32"))]
    fn new(source_text: &str, images: &[ImageAsset], font_paths: &[PathBuf]) -> Self {
        let font_source = if font_paths.is_empty() {
            FontSource::Cached(get_system_fonts())
        } else {
            FontSource::Shared(get_fonts_for_extra_paths(font_paths))
        };

        let main_id = FileId::new(None, VirtualPath::new("main.typ"));
        let source = Source::new(main_id, source_text.to_string());

        let image_map: HashMap<String, Bytes> = images
            .iter()
            .map(|a| (a.path.clone(), Bytes::new(a.data.clone())))
            .collect();

        Self {
            library: LazyHash::new(Library::default()),
            font_source,
            source,
            images: image_map,
        }
    }

    /// Create a new `MinimalWorld` with embedded fonts only (no system font search).
    ///
    /// Uses a process-wide cache for embedded font data. This is the constructor
    /// used on WASM targets where system font discovery is not available.
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    fn new_embedded_only(source_text: &str, images: &[ImageAsset]) -> Self {
        let main_id = FileId::new(None, VirtualPath::new("main.typ"));
        let source = Source::new(main_id, source_text.to_string());

        let image_map: HashMap<String, Bytes> = images
            .iter()
            .map(|a| (a.path.clone(), Bytes::new(a.data.clone())))
            .collect();

        Self {
            library: LazyHash::new(Library::default()),
            font_source: FontSource::Cached(get_embedded_fonts()),
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
        self.font_source.book()
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
        self.font_source
            .fonts()
            .get(index)
            .and_then(|slot| slot.get())
    }

    fn today(&self, _offset: Option<i64>) -> Option<Datetime> {
        None
    }
}

#[cfg(test)]
#[path = "pdf_tests.rs"]
mod tests;
