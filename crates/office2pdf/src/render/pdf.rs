use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;
use std::sync::Mutex;
use std::sync::{Arc, OnceLock};
// `SystemTime::now()` panics on wasm32-unknown-unknown; web-time shims it there
// and re-exports std elsewhere. Mirrors the `Instant` handling in lib_pipeline.
#[cfg(not(target_arch = "wasm32"))]
use std::time::{SystemTime, UNIX_EPOCH};
#[cfg(target_arch = "wasm32")]
use web_time::{SystemTime, UNIX_EPOCH};

use typst::diag::FileResult;
use typst::foundations::{Bytes, Datetime, Duration};
use typst::syntax::{FileId, RootedPath, Source, VirtualPath, VirtualRoot};
use typst::text::Font;
use typst::utils::LazyHash;
use typst::{Library, LibraryExt, World};
use typst_layout::{Page, PagedDocument};

use crate::config::PdfStandard;
use crate::error::ConvertError;

use super::typst_gen::ImageAsset;

/// Cached font data (book + font slots). Font discovery is expensive because
/// it scans the filesystem; the result doesn't change during the process
/// lifetime, so we cache it in a global `OnceLock`.
struct CachedFontData {
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<FontSlot>,
}

/// A discovered face, read from disk on first use.
///
/// typst-kit 0.15 keeps its own slots private, and [`line_metric_face_at`]
/// needs a face's file to tell an Office bundle's copy from the system's.
struct FontSlot {
    path: Option<std::path::PathBuf>,
    index: u32,
    font: OnceLock<Option<Font>>,
}

impl FontSlot {
    #[cfg(not(target_arch = "wasm32"))]
    fn on_disk(found: typst_kit::fonts::FontPath) -> Self {
        Self {
            path: Some(found.path),
            index: found.index,
            font: OnceLock::new(),
        }
    }

    #[cfg(any(feature = "embedded-fonts", target_arch = "wasm32"))]
    fn loaded(font: Font) -> Self {
        Self {
            path: None,
            index: font.index(),
            font: OnceLock::from(Some(font)),
        }
    }

    /// The file holding the face, or `None` for one embedded in the binary.
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    fn path(&self) -> Option<&std::path::Path> {
        self.path.as_deref()
    }

    /// The face, read from its file the first time it is asked for.
    fn get(&self) -> Option<Font> {
        self.font
            .get_or_init(|| {
                let data: Vec<u8> = std::fs::read(self.path.as_ref()?).ok()?;
                Font::new(Bytes::new(data), self.index)
            })
            .clone()
    }
}

/// Faces in the priority order typst-kit 0.14's `FontSearcher` gave them:
/// `font_dirs` in order, then the system's fonts, then typst's embedded ones.
/// typst-kit 0.15 hands out the three sources as separate iterators.
#[cfg(not(target_arch = "wasm32"))]
fn discover_fonts(
    font_dirs: &[PathBuf],
    include_system_fonts: bool,
    include_embedded_fonts: bool,
) -> (typst::text::FontBook, Vec<FontSlot>) {
    let mut book = typst::text::FontBook::new();
    let mut fonts: Vec<FontSlot> = Vec::new();
    for dir in font_dirs {
        for (found, info) in typst_kit::fonts::scan(dir) {
            book.push(info);
            fonts.push(FontSlot::on_disk(found));
        }
    }
    if include_system_fonts {
        for (found, info) in typst_kit::fonts::system() {
            book.push(info);
            fonts.push(FontSlot::on_disk(found));
        }
    }
    if include_embedded_fonts {
        push_embedded_fonts(&mut book, &mut fonts);
    }
    (book, fonts)
}

/// The book [`discover_fonts`] builds, for callers that only index families.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn discover_font_book(
    font_dirs: &[PathBuf],
    include_system_fonts: bool,
    include_embedded_fonts: bool,
) -> typst::text::FontBook {
    discover_fonts(font_dirs, include_system_fonts, include_embedded_fonts).0
}

/// Typst's fallback faces, empty on native builds without `embedded-fonts`.
fn embedded_fonts() -> (typst::text::FontBook, Vec<FontSlot>) {
    let mut book = typst::text::FontBook::new();
    let mut fonts: Vec<FontSlot> = Vec::new();
    push_embedded_fonts(&mut book, &mut fonts);
    (book, fonts)
}

/// Appends the faces embedded in typst-assets, which rank below every
/// discovered face. Always available on WASM, opt-out on native builds.
#[cfg(any(feature = "embedded-fonts", target_arch = "wasm32"))]
fn push_embedded_fonts(book: &mut typst::text::FontBook, fonts: &mut Vec<FontSlot>) {
    for (font, info) in typst_kit::fonts::embedded() {
        book.push(info);
        fonts.push(FontSlot::loaded(font));
    }
}

#[cfg(all(not(feature = "embedded-fonts"), not(target_arch = "wasm32")))]
fn push_embedded_fonts(_book: &mut typst::text::FontBook, _fonts: &mut Vec<FontSlot>) {}

/// Document- or caller-provided in-memory faces followed by cached fallback
/// slots. The combined book preserves the same priority order that native
/// `discover_fonts` gives an explicit font directory without eagerly loading all
/// fallback font bytes.
struct InMemoryFontData {
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<Font>,
    fallback: FallbackFontData,
}

impl InMemoryFontData {
    fn new(fonts: &[Font], fallback: FallbackFontData) -> Self {
        let fallback_data = fallback.data();
        let infos = fonts.iter().map(|font| font.info().clone()).chain(
            (0..fallback_data.fonts.len())
                .filter_map(|index| fallback_data.book.info(index).cloned()),
        );

        Self {
            book: LazyHash::new(typst::text::FontBook::from_infos(infos)),
            fonts: fonts.to_vec(),
            fallback,
        }
    }
}

/// Lazily loaded font slots that follow caller-provided in-memory faces.
enum FallbackFontData {
    Cached(&'static CachedFontData),
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    Shared(Arc<CachedFontData>),
}

impl FallbackFontData {
    fn data(&self) -> &CachedFontData {
        match self {
            Self::Cached(data) => data,
            Self::Shared(data) => data,
        }
    }
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

/// Clear process-global Typst memoization after this many independent Office
/// documents. The interval avoids evicting on every call while bounding the
/// number of completed documents retained between cleanup opportunities.
const TYPST_CACHE_EVICTION_INTERVAL: usize = 64;

struct TypstCacheState {
    active_compilations: usize,
    completed_since_eviction: usize,
}

impl TypstCacheState {
    fn begin_compilation(&mut self) -> bool {
        let should_evict = self.active_compilations == 0
            && self.completed_since_eviction >= TYPST_CACHE_EVICTION_INTERVAL;
        if should_evict {
            self.completed_since_eviction = 0;
        }
        self.active_compilations += 1;
        should_evict
    }

    fn finish_compilation(&mut self) {
        self.active_compilations = self
            .active_compilations
            .checked_sub(1)
            .expect("a Typst compilation guard must balance its registration");
        self.completed_since_eviction = self.completed_since_eviction.saturating_add(1);
    }
}

static TYPST_CACHE_STATE: Mutex<TypstCacheState> = Mutex::new(TypstCacheState {
    active_compilations: 0,
    completed_since_eviction: 0,
});

/// Keeps cache aging outside every overlapping group of Typst compilations.
struct TypstCompilationGuard;

impl TypstCompilationGuard {
    fn begin() -> Self {
        let mut state = TYPST_CACHE_STATE
            .lock()
            .expect("Typst cache state mutex should not be poisoned");
        if state.begin_compilation() {
            comemo::evict(0);
        }

        Self
    }
}

impl Drop for TypstCompilationGuard {
    fn drop(&mut self) {
        let mut state = TYPST_CACHE_STATE
            .lock()
            .expect("Typst cache state mutex should not be poisoned");
        state.finish_compilation();
    }
}

/// Get or initialize cached system fonts (with system font discovery).
#[cfg(not(target_arch = "wasm32"))]
fn get_system_fonts() -> &'static CachedFontData {
    SYSTEM_FONTS.get_or_init(|| {
        let (book, fonts) = discover_fonts(&[], true, true);
        CachedFontData {
            book: LazyHash::new(book),
            fonts,
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

    let (book, fonts) = discover_fonts(font_paths, true, true);
    let cached = Arc::new(CachedFontData {
        book: LazyHash::new(book),
        fonts,
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
        let (book, fonts) = embedded_fonts();
        CachedFontData {
            book: LazyHash::new(book),
            fonts,
        }
    })
}

/// The instance of `font` whose tables the metric helpers read.
///
/// typst 0.15 moved a face's tables onto a [`typst::text::FontInstance`] fixed
/// to variation coordinates. Instantiating at the face's own variant reads a
/// static face exactly as typst 0.14's `Font::ttf` did; Typst's 11pt default
/// text size stands in for the size an `opsz` axis would follow.
pub(crate) fn measured_instance(font: &Font) -> typst::text::FontInstance {
    font.clone().instantiate(
        font.info().variant,
        typst::layout::Abs::pt(11.0),
        &typst::text::FontVariations::default(),
    )
}

/// Parse standalone font or font-collection bytes into Typst faces.
pub(crate) fn load_fonts_from_bytes<'a>(
    font_data: impl IntoIterator<Item = &'a [u8]>,
) -> Vec<Font> {
    font_data
        .into_iter()
        .flat_map(|data| Font::iter(Bytes::new(data.to_vec())))
        .collect()
}

/// Compile Typst markup to PDF bytes.
///
/// When `pdf_standard` is `Some`, the output PDF will conform to the
/// specified standard (e.g., PDF/A-2b for archival).
/// When `font_paths` is non-empty, those directories are searched for
/// additional fonts (highest priority).
///
/// On native targets, system fonts are discovered automatically. On WASM,
/// built-in, document-embedded, and caller-provided in-memory fonts are used;
/// `font_paths` is ignored.
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
    compile_to_pdf_inner(&world, pdf_standard, tagged, pdf_ua).map(|(pdf, _pages)| pdf)
}

/// Compile Typst markup to PDF bytes (WASM target).
///
/// Uses built-in fonts plus any fonts embedded by the document conversion
/// pipeline. System font paths are not supported on WASM.
#[cfg(target_arch = "wasm32")]
#[cfg_attr(feature = "wasm-cjk-font", allow(dead_code))]
pub fn compile_to_pdf(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    _font_paths: &[std::path::PathBuf],
    tagged: bool,
    pdf_ua: bool,
) -> Result<Vec<u8>, ConvertError> {
    compile_to_pdf_with_fonts(typst_source, images, pdf_standard, &[], &[], tagged, pdf_ua)
}

/// Compile Typst markup with document- or caller-provided in-memory fonts.
pub(crate) fn compile_to_pdf_with_fonts(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    font_paths: &[std::path::PathBuf],
    in_memory_fonts: &[Font],
    tagged: bool,
    pdf_ua: bool,
) -> Result<Vec<u8>, ConvertError> {
    compile_to_pdf_with_fonts_counted(
        typst_source,
        images,
        pdf_standard,
        font_paths,
        in_memory_fonts,
        tagged,
        pdf_ua,
    )
    .map(|(pdf, _pages)| pdf)
}

/// Compile as `compile_to_pdf_with_fonts` does, also reporting how many pages
/// the layout produced.
///
/// The count comes from the laid-out document rather than the source IR: a
/// source page, including a DOCX flow section, may lay out across multiple PDF
/// pages. Callers reporting `ConvertMetrics::page_count` need the PDF count.
pub(crate) fn compile_to_pdf_with_fonts_counted(
    typst_source: &str,
    images: &[ImageAsset],
    pdf_standard: Option<PdfStandard>,
    font_paths: &[std::path::PathBuf],
    in_memory_fonts: &[Font],
    tagged: bool,
    pdf_ua: bool,
) -> Result<(Vec<u8>, u32), ConvertError> {
    #[cfg(not(target_arch = "wasm32"))]
    let world =
        MinimalWorld::new_with_in_memory_fonts(typst_source, images, font_paths, in_memory_fonts);
    #[cfg(target_arch = "wasm32")]
    let world = {
        let _ = font_paths;
        MinimalWorld::new_embedded_with_fonts(typst_source, images, in_memory_fonts)
    };
    compile_to_pdf_inner(&world, pdf_standard, tagged, pdf_ua)
}

/// Lay out Typst markup and report its page count without exporting a PDF.
///
/// Streaming XLSX conversion uses this cheaper probe to move each compilation
/// boundary to the end of a laid-out page. That keeps the merged PDF's page
/// count independent of arbitrary row chunk sizes without retaining a whole
/// workbook layout.
pub(crate) fn compile_page_count_with_fonts(
    typst_source: &str,
    images: &[ImageAsset],
    font_paths: &[std::path::PathBuf],
    in_memory_fonts: &[Font],
) -> Result<u32, ConvertError> {
    #[cfg(not(target_arch = "wasm32"))]
    let world =
        MinimalWorld::new_with_in_memory_fonts(typst_source, images, font_paths, in_memory_fonts);
    #[cfg(target_arch = "wasm32")]
    let world = {
        let _ = font_paths;
        MinimalWorld::new_embedded_with_fonts(typst_source, images, in_memory_fonts)
    };
    let _compilation_guard = TypstCompilationGuard::begin();
    let warned = typst::compile::<PagedDocument>(&world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors
            .iter()
            .map(|error| error.message.to_string())
            .collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;
    Ok(document.pages().len() as u32)
}

/// The passes that move Typst's completed line paint into place.
///
/// Typst supplies the shaping, wrapping and line boxes; each pass reads its
/// own codegen marker and leaves every other frame alone, so their order only
/// has to be the same wherever a compiled document is inspected or exported.
fn apply_completed_frame_passes(document: &mut PagedDocument) {
    // typst 0.15 hands pages out read-only, so the edited pages make a new
    // document, whose introspector is rebuilt from them.
    let mut pages: Vec<Page> = document.pages().to_vec();
    super::powerpoint_line_paint::adjust_paragraph_marks(&mut pages);
    super::excel_fill_paint::adjust_cell_fills(&mut pages);
    super::excel_glyph_pacing::pace_sheet_glyphs_on_whole_points(&mut pages);
    super::word_justified_gap_phases::spread_justified_gaps_as_word_does(&mut pages);
    let info: typst::model::DocumentInfo = typst::model::Document::info(document).clone();
    *document = PagedDocument::new(pages.into_iter().collect(), info);
}

fn compile_to_pdf_inner(
    world: &MinimalWorld,
    pdf_standard: Option<PdfStandard>,
    tagged: bool,
    pdf_ua: bool,
) -> Result<(Vec<u8>, u32), ConvertError> {
    // Typst's memoized layout results are process-global, while each
    // `MinimalWorld` here is an independent Office document. Age that cache at
    // a bounded document interval so long-running processes do not retain
    // results from every earlier document. The guard keeps eviction outside
    // overlapping conversions, whose live layout entries may still be in use.
    let _compilation_guard = TypstCompilationGuard::begin();

    let warned = typst::compile::<PagedDocument>(world);
    let mut document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    apply_completed_frame_passes(&mut document);

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
        typst_pdf::PdfStandards::new(&pdf_standards).map_err(|e| {
            ConvertError::Render(format!("PDF standard configuration error: {}", e.message()))
        })?
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
    let page_count = document.pages().len() as u32;
    let pdf = typst_pdf::pdf(&document, &options).map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("PDF export failed: {}", messages.join("; ")))
    })?;
    Ok((pdf, page_count))
}

/// One shaped text run as the layout engine actually placed it.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone)]
pub(crate) struct PlacedTextRun {
    /// Distance from the page's top edge down to the run's baseline origin, in
    /// points. A turned run reports the page-space image of that origin, so it
    /// is the point the transform moved rather than a horizontal baseline.
    pub baseline_pt: f64,
    /// Distance from the page's left edge to the run's origin, in points.
    pub left_pt: f64,
    /// The family the run was actually shaped with, which is not always the one
    /// the source asked for.
    pub family: String,
    pub text: String,
}

/// Every text run the compiled document places on `page_index`, in layout order.
///
/// The emitted source cannot show where a line ends up: `top-edge`, `place` and
/// `measure` are all resolved by the layout engine, and issue #629's first
/// attempt passed a source-string assertion while moving every wrapped line of
/// the paragraph it touched. Tests for placement therefore compile the source
/// and read the frames.
///
/// Group transforms are accumulated in full, not just their translation: a
/// rotated text box turns the group its runs sit in, and reading only `tx`/`ty`
/// reports every one of those runs at the unturned place (issue #1078).
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_text_runs(
    typst_source: &str,
    page_index: usize,
) -> Result<Vec<PlacedTextRun>, ConvertError> {
    compiled_text_runs_with_line_seating(typst_source, page_index, true, &[])
}

/// Inspect placement with caller-provided faces, independent of host fonts.
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_text_runs_with_fonts(
    typst_source: &str,
    page_index: usize,
    fonts: &[Font],
) -> Result<Vec<PlacedTextRun>, ConvertError> {
    compiled_text_runs_with_line_seating(typst_source, page_index, true, fonts)
}

/// Inspect Typst's fractional line advances before the completed-frame pass.
/// These positions test layout height independently of the painted line grid.
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_text_runs_before_line_seating(
    typst_source: &str,
    page_index: usize,
) -> Result<Vec<PlacedTextRun>, ConvertError> {
    compiled_text_runs_with_line_seating(typst_source, page_index, false, &[])
}

#[cfg(all(test, not(target_arch = "wasm32")))]
fn compiled_text_runs_with_line_seating(
    typst_source: &str,
    page_index: usize,
    apply_line_seating: bool,
    fonts: &[Font],
) -> Result<Vec<PlacedTextRun>, ConvertError> {
    use typst::layout::{Frame, FrameItem, Transform};

    fn collect(frame: &Frame, transform: Transform, out: &mut Vec<PlacedTextRun>) {
        for (position, item) in frame.items() {
            let at: Transform = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    collect(&group.frame, at.pre_concat(group.transform), out);
                }
                // The run's own origin is the local `(0, 0)` the accumulated
                // transform maps to, which is exactly its translation column.
                FrameItem::Text(text) => out.push(PlacedTextRun {
                    baseline_pt: at.ty.to_pt(),
                    left_pt: at.tx.to_pt(),
                    family: text.font.info().family.clone(),
                    text: text.text.to_string(),
                }),
                _ => {}
            }
        }
    }

    // The same font set the conversion pipeline compiles with, so a probe sees
    // the faces `font_hhea_ascender_em` measured rather than a substitute.
    let font_paths: &[PathBuf] = super::font_context::default_font_search_paths();
    let world = MinimalWorld::new_with_in_memory_fonts(typst_source, &[], font_paths, fonts);
    let warned = typst::compile::<PagedDocument>(&world);
    let mut document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;
    if apply_line_seating {
        apply_completed_frame_passes(&mut document);
    }
    let page = document.pages().get(page_index).ok_or_else(|| {
        ConvertError::Render(format!(
            "page {page_index} is past the document's {} pages",
            document.pages().len()
        ))
    })?;
    let mut runs: Vec<PlacedTextRun> = Vec::new();
    collect(&page.frame, Transform::identity(), &mut runs);
    Ok(runs)
}

/// One shaped run's glyph pacing as the layout engine actually placed it.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone)]
pub(crate) struct PlacedGlyphRun {
    /// Distance from the page's left edge to the run's origin, in points.
    pub left_pt: f64,
    pub text: String,
    /// The advance each glyph takes, in points and in layout order. The
    /// run's glyph origins are this list's running sums from `left_pt`.
    pub advances_pt: Vec<f64>,
}

#[cfg(all(test, not(target_arch = "wasm32")))]
impl PlacedGlyphRun {
    /// Where each glyph after the first starts, relative to the run's origin.
    pub fn glyph_origins_pt(&self) -> Vec<f64> {
        let mut pen: f64 = 0.0;
        self.advances_pt
            .iter()
            .take(self.advances_pt.len().saturating_sub(1))
            .map(|advance| {
                pen += advance;
                pen
            })
            .collect()
    }

    pub fn width_pt(&self) -> f64 {
        self.advances_pt.iter().sum()
    }
}

/// Every shaped run's glyph advances on `page_index`, in layout order.
///
/// The completed-frame passes rewrite glyph advances rather than emitted
/// source, so a pacing test has to read the frames. `apply_frame_passes`
/// selects the placement before or after those passes, which is what lets a
/// test assert that a pass conserved a run's width.
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_glyph_runs(
    typst_source: &str,
    page_index: usize,
    apply_frame_passes: bool,
) -> Result<Vec<PlacedGlyphRun>, ConvertError> {
    use typst::layout::{Frame, FrameItem, Transform};

    fn collect(frame: &Frame, transform: Transform, out: &mut Vec<PlacedGlyphRun>) {
        for (position, item) in frame.items() {
            let at: Transform = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    collect(&group.frame, at.pre_concat(group.transform), out);
                }
                FrameItem::Text(text) => out.push(PlacedGlyphRun {
                    left_pt: at.tx.to_pt(),
                    text: text.text.to_string(),
                    advances_pt: text
                        .glyphs
                        .iter()
                        .map(|glyph| glyph.x_advance.at(text.size).to_pt())
                        .collect(),
                }),
                _ => {}
            }
        }
    }

    let font_paths: &[PathBuf] = super::font_context::default_font_search_paths();
    let world = MinimalWorld::new_with_in_memory_fonts(typst_source, &[], font_paths, &[]);
    let warned = typst::compile::<PagedDocument>(&world);
    let mut document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;
    if apply_frame_passes {
        apply_completed_frame_passes(&mut document);
    }
    let page = document.pages().get(page_index).ok_or_else(|| {
        ConvertError::Render(format!(
            "page {page_index} is past the document's {} pages",
            document.pages().len()
        ))
    })?;
    let mut runs: Vec<PlacedGlyphRun> = Vec::new();
    collect(&page.frame, Transform::identity(), &mut runs);
    Ok(runs)
}

/// One image box as the layout engine actually placed it.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone, Copy)]
pub(crate) struct PlacedImageBox {
    /// Page-space positions of the source image's own corners, in points,
    /// running from its top-left clockwise. A mirrored or turned picture
    /// reports them in the same source order, so the corner a transform was
    /// supposed to move is the one to read.
    pub corners: [(f64, f64); 4],
}

/// Every image the compiled document paints on `page_index`, in layout order.
///
/// `#rotate` and `#scale` resolve their `origin` against a frame the layout
/// engine sizes, so the emitted source cannot show where a turned picture
/// ends up. Follow the frame tree instead and accumulate the full transform,
/// not just its translation, or a turned picture reports the wrong box
/// (issue #1032).
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_image_boxes(
    typst_source: &str,
    images: &[ImageAsset],
    page_index: usize,
) -> Result<Vec<PlacedImageBox>, ConvertError> {
    use typst::layout::{Frame, FrameItem, Transform};

    fn place(transform: Transform, x: f64, y: f64) -> (f64, f64) {
        (
            transform.sx.get() * x + transform.kx.get() * y + transform.tx.to_pt(),
            transform.ky.get() * x + transform.sy.get() * y + transform.ty.to_pt(),
        )
    }

    fn collect(frame: &Frame, transform: Transform, out: &mut Vec<PlacedImageBox>) {
        for (position, item) in frame.items() {
            let at: Transform = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    collect(&group.frame, at.pre_concat(group.transform), out);
                }
                FrameItem::Image(_, size, _) => {
                    let (width, height): (f64, f64) = (size.x.to_pt(), size.y.to_pt());
                    out.push(PlacedImageBox {
                        corners: [
                            place(at, 0.0, 0.0),
                            place(at, width, 0.0),
                            place(at, width, height),
                            place(at, 0.0, height),
                        ],
                    });
                }
                _ => {}
            }
        }
    }

    let world = MinimalWorld::new(typst_source, images, &[]);
    let warned = typst::compile::<PagedDocument>(&world);
    let mut document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;
    apply_completed_frame_passes(&mut document);
    let page = document.pages().get(page_index).ok_or_else(|| {
        ConvertError::Render(format!(
            "page {page_index} is past the document's {} pages",
            document.pages().len()
        ))
    })?;
    let mut boxes: Vec<PlacedImageBox> = Vec::new();
    collect(&page.frame, Transform::identity(), &mut boxes);
    Ok(boxes)
}

/// What one primitive a compiled page paints is.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PaintedKind {
    /// A filled or stroked shape: a cell background, a rule, a border band.
    Shape,
    /// A raster or vector picture.
    Image,
    /// A run of glyphs.
    Text,
}

/// One primitive a compiled page paints, and the box it covers.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone)]
pub(crate) struct PaintedPrimitive {
    pub kind: PaintedKind,
    /// Page-space bounding box in points, as `(min x, min y, max x, max y)`.
    pub bounds: (f64, f64, f64, f64),
    pub stroke: Option<PaintedStroke>,
    /// Solid rectangle fill; other shape geometry and paint types remain absent.
    pub rectangle_fill: Option<typst::visualize::Color>,
}

/// Stroke geometry after compilation, before the enclosing frame transform.
#[cfg(all(test, not(target_arch = "wasm32")))]
#[derive(Debug, Clone)]
pub(crate) struct PaintedStroke {
    pub color: Option<typst::visualize::Color>,
    pub thickness_pt: f64,
    pub cap: typst::visualize::LineCap,
    pub join: typst::visualize::LineJoin,
    pub miter_limit: f64,
}

#[cfg(all(test, not(target_arch = "wasm32")))]
impl PaintedPrimitive {
    /// Whether this primitive's box contains all of `other`'s.
    pub fn covers(&self, other: &PaintedPrimitive) -> bool {
        self.bounds.0 <= other.bounds.0
            && self.bounds.1 <= other.bounds.1
            && self.bounds.2 >= other.bounds.2
            && self.bounds.3 >= other.bounds.3
    }
}

/// Every primitive the compiled document paints on `page_index`, in the order
/// it paints them.
///
/// A later entry covers an earlier one wherever the two overlap, which is the
/// only way to read z-order: the emitted markup shows document order, and
/// Typst's page foreground paints after a body that precedes it in the source
/// (issue #1168).
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_paint_sequence(
    typst_source: &str,
    images: &[ImageAsset],
    page_index: usize,
) -> Result<Vec<PaintedPrimitive>, ConvertError> {
    let pages = compiled_page_paint_sequences(typst_source, images)?;
    pages.get(page_index).cloned().ok_or_else(|| {
        ConvertError::Render(format!(
            "page {page_index} is past the document's {} pages",
            pages.len()
        ))
    })
}

/// Compile a probe collection once and preserve each page's paint sequence.
#[cfg(all(test, not(target_arch = "wasm32")))]
pub(crate) fn compiled_page_paint_sequences(
    typst_source: &str,
    images: &[ImageAsset],
) -> Result<Vec<Vec<PaintedPrimitive>>, ConvertError> {
    use typst::layout::{Frame, FrameItem, Transform};
    use typst::visualize::Geometry;

    fn place(transform: Transform, x: f64, y: f64) -> (f64, f64) {
        (
            transform.sx.get() * x + transform.kx.get() * y + transform.tx.to_pt(),
            transform.ky.get() * x + transform.sy.get() * y + transform.ty.to_pt(),
        )
    }

    /// The box `(0, 0)`-`(width, height)` spans once `transform` has moved and
    /// turned it. A turned box reports the box around its turned corners,
    /// which is all a coverage test needs.
    fn transformed_bounds(transform: Transform, width: f64, height: f64) -> (f64, f64, f64, f64) {
        let corners: [(f64, f64); 4] = [
            place(transform, 0.0, 0.0),
            place(transform, width, 0.0),
            place(transform, width, height),
            place(transform, 0.0, height),
        ];
        corners.iter().fold(
            (f64::MAX, f64::MAX, f64::MIN, f64::MIN),
            |bounds, corner| {
                (
                    bounds.0.min(corner.0),
                    bounds.1.min(corner.1),
                    bounds.2.max(corner.0),
                    bounds.3.max(corner.1),
                )
            },
        )
    }

    fn collect(frame: &Frame, transform: Transform, out: &mut Vec<PaintedPrimitive>) {
        for (position, item) in frame.items() {
            let at: Transform = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    collect(&group.frame, at.pre_concat(group.transform), out);
                }
                FrameItem::Image(_, size, _) => out.push(PaintedPrimitive {
                    kind: PaintedKind::Image,
                    rectangle_fill: None,
                    bounds: transformed_bounds(at, size.x.to_pt(), size.y.to_pt()),
                    stroke: None,
                }),
                FrameItem::Shape(shape, _) => {
                    let bounds: (f64, f64, f64, f64) = match &shape.geometry {
                        Geometry::Line(to) => {
                            let end: (f64, f64) = place(at, to.x.to_pt(), to.y.to_pt());
                            let start: (f64, f64) = place(at, 0.0, 0.0);
                            (
                                start.0.min(end.0),
                                start.1.min(end.1),
                                start.0.max(end.0),
                                start.1.max(end.1),
                            )
                        }
                        Geometry::Rect(size) => {
                            transformed_bounds(at, size.x.to_pt(), size.y.to_pt())
                        }
                        Geometry::Curve(curve) => {
                            let box_ = curve.bbox(None);
                            let min: (f64, f64) = place(at, box_.min.x.to_pt(), box_.min.y.to_pt());
                            let max: (f64, f64) = place(at, box_.max.x.to_pt(), box_.max.y.to_pt());
                            (
                                min.0.min(max.0),
                                min.1.min(max.1),
                                min.0.max(max.0),
                                min.1.max(max.1),
                            )
                        }
                    };
                    out.push(PaintedPrimitive {
                        kind: PaintedKind::Shape,
                        rectangle_fill: match (&shape.geometry, &shape.fill) {
                            (Geometry::Rect(_), Some(typst::visualize::Paint::Solid(color))) => {
                                Some(color.clone())
                            }
                            _ => None,
                        },
                        bounds,
                        stroke: shape.stroke.as_ref().map(|stroke| PaintedStroke {
                            color: match &stroke.paint {
                                typst::visualize::Paint::Solid(color) => Some(color.clone()),
                                _ => None,
                            },
                            thickness_pt: stroke.thickness.to_pt(),
                            cap: stroke.cap,
                            join: stroke.join,
                            miter_limit: stroke.miter_limit.get(),
                        }),
                    });
                }
                FrameItem::Text(text) => {
                    let width: f64 = text.width().to_pt();
                    let size: f64 = text.size.to_pt();
                    // A run's origin is its baseline, so the box it inks runs
                    // from roughly one em above it down to its descender.
                    let top: (f64, f64) = place(at, 0.0, -size);
                    let bottom: (f64, f64) = place(at, width, size / 4.0);
                    out.push(PaintedPrimitive {
                        kind: PaintedKind::Text,
                        rectangle_fill: None,
                        stroke: None,
                        bounds: (
                            top.0.min(bottom.0),
                            top.1.min(bottom.1),
                            top.0.max(bottom.0),
                            top.1.max(bottom.1),
                        ),
                    });
                }
                FrameItem::Link(..) | FrameItem::Tag(_) => {}
            }
        }
    }

    let world = MinimalWorld::new(typst_source, images, &[]);
    let warned = typst::compile::<PagedDocument>(&world);
    let mut document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;
    apply_completed_frame_passes(&mut document);
    Ok(document
        .pages()
        .iter()
        .map(|page| {
            let mut painted: Vec<PaintedPrimitive> = Vec::new();
            collect(&page.frame, Transform::identity(), &mut painted);
            painted
        })
        .collect())
}

#[cfg(all(test, not(target_arch = "wasm32")))]
#[test]
fn powerpoint_line_seating_keeps_links_and_underlines_with_text() {
    use typst::layout::{Frame, FrameItem, Transform};

    fn positions(frame: &Frame, transform: Transform, output: &mut [Vec<f64>; 3]) {
        for (point, item) in frame.items() {
            let at = transform.pre_concat(Transform::translate(point.x, point.y));
            match item {
                FrameItem::Group(group) => {
                    positions(&group.frame, at.pre_concat(group.transform), output);
                }
                FrameItem::Text(_) => output[0].push(at.ty.to_pt()),
                FrameItem::Link(..) => output[1].push(at.ty.to_pt()),
                FrameItem::Shape(..) => output[2].push(at.ty.to_pt()),
                _ => {}
            }
        }
    }

    let source = r##"
#set page(width: 300pt, height: 200pt, margin: 0pt)
#set text(font: "Libertinus Serif", size: 16pt, top-edge: 16pt, bottom-edge: 4pt)
#set par(leading: 0pt)
#metadata(("office2pdf-pptx-paragraph-mark", 1.0))
#block(width: 150pt)[#underline[#link("https://example.com")[Server-side document conversion preserves links and emphasis across wrapped lines.]]]
#metadata("office2pdf-pptx-paragraph-end")
"##;
    for round_lines in [false, true] {
        let source = if round_lines {
            // A 20.4pt line advance carries a raw 16.4pt seat through a
            // first baseline already painted at 16pt. Later lines need
            // different movements, including +0.6pt rather than -0.4pt.
            source
                .replace("bottom-edge: 4pt", "bottom-edge: -4.4pt")
                .replace(
                    "(\"office2pdf-pptx-paragraph-mark\", 1.0)",
                    "(\"office2pdf-pptx-paragraph-mark\", 0.0, (16.4, 16.0, 16.4))",
                )
        } else {
            source.to_owned()
        };
        let world = MinimalWorld::new(&source, &[], &[]);
        let mut document = typst::compile::<PagedDocument>(&world).output.unwrap();
        let mut before = [Vec::new(), Vec::new(), Vec::new()];
        positions(
            &document.pages()[0].frame,
            Transform::identity(),
            &mut before,
        );
        apply_completed_frame_passes(&mut document);
        let mut after = [Vec::new(), Vec::new(), Vec::new()];
        positions(
            &document.pages()[0].frame,
            Transform::identity(),
            &mut after,
        );
        assert!(before[0].len() > 1, "the paragraph must wrap");
        for kind in 0..2 {
            assert_eq!(before[kind].len(), before[0].len());
            assert_eq!(after[kind].len(), before[kind].len());
            for line in 0..before[0].len() {
                let expected = if round_lines {
                    [0.0, 0.6, 0.2, 0.8, 0.4][line % 5]
                } else if line + 1 == before[0].len() {
                    0.0
                } else {
                    1.0
                };
                assert!(
                    (after[kind][line] - before[kind][line] - expected).abs() < 0.001,
                    "kind {kind}, line {line}: {before:?} -> {after:?}",
                );
            }
        }
        assert!(before[2].len() >= before[0].len());
        assert_eq!(before[2].len(), after[2].len());
        for (old, new) in before[2].iter().zip(&after[2]) {
            let line = before[0]
                .iter()
                .enumerate()
                .min_by(|(_, a), (_, b)| (*a - old).abs().total_cmp(&(*b - old).abs()))
                .unwrap()
                .0;
            let expected = after[0][line] - before[0][line];
            assert!((new - old - expected).abs() < 0.001);
        }
    }
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
    /// Only constructed on native (extra font paths need filesystem access).
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    Shared(Arc<CachedFontData>),
    /// Document- or caller-provided faces held in memory, followed by cached
    /// fallback slots.
    InMemory(InMemoryFontData),
}

impl FontSource {
    fn book(&self) -> &LazyHash<typst::text::FontBook> {
        match self {
            Self::Cached(d) => &d.book,
            Self::Shared(d) => &d.book,
            Self::InMemory(d) => &d.book,
        }
    }

    #[cfg(test)]
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    fn len(&self) -> usize {
        match self {
            Self::Cached(d) => d.fonts.len(),
            Self::Shared(d) => d.fonts.len(),
            Self::InMemory(d) => d.fonts.len() + d.fallback.data().fonts.len(),
        }
    }

    fn font(&self, index: usize) -> Option<Font> {
        match self {
            Self::Cached(d) => d.fonts.get(index).and_then(|slot| slot.get()),
            Self::Shared(d) => d.fonts.get(index).and_then(|slot| slot.get()),
            Self::InMemory(d) => d.fonts.get(index).cloned().or_else(|| {
                d.fallback
                    .data()
                    .fonts
                    .get(index.checked_sub(d.fonts.len())?)
                    .and_then(|slot| slot.get())
            }),
        }
    }
}

/// Minimal World implementation providing Typst compiler with source, fonts, and images.
struct MinimalWorld {
    library: LazyHash<Library>,
    font_source: FontSource,
    source: Source,
    images: HashMap<String, Bytes>,
    /// Faces as the shaper is to receive them, by font-book index.
    ///
    /// The compiler asks for the same face once per text run and a face that
    /// carries both kern sources is rebuilt from its own bytes, so the answer
    /// is memoized for the life of the compilation (issue #1116).
    shaped_faces: Mutex<HashMap<usize, Font>>,
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

        let main_id = main_file_id();
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
            shaped_faces: Mutex::new(HashMap::new()),
        }
    }

    /// Create a world without system font discovery, using the configured
    /// Typst embedded set (empty on native builds without `embedded-fonts`).
    ///
    /// Uses a process-wide cache for embedded font data. This is the constructor
    /// used on WASM targets where system font discovery is not available.
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    fn new_embedded_only(source_text: &str, images: &[ImageAsset]) -> Self {
        Self::new_with_font_source(
            source_text,
            images,
            FontSource::Cached(get_embedded_fonts()),
        )
    }

    /// Create an embedded-only world with per-conversion in-memory faces at
    /// higher priority than the configured Typst embedded set.
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    fn new_embedded_with_fonts(
        source_text: &str,
        images: &[ImageAsset],
        document_fonts: &[Font],
    ) -> Self {
        if document_fonts.is_empty() {
            return Self::new_embedded_only(source_text, images);
        }

        Self::new_with_font_source(
            source_text,
            images,
            FontSource::InMemory(InMemoryFontData::new(
                document_fonts,
                FallbackFontData::Cached(get_embedded_fonts()),
            )),
        )
    }

    /// Create a native world with in-memory faces ahead of the same system or
    /// custom-path slots used by the ordinary native compiler.
    #[cfg(not(target_arch = "wasm32"))]
    fn new_with_in_memory_fonts(
        source_text: &str,
        images: &[ImageAsset],
        font_paths: &[PathBuf],
        fonts: &[Font],
    ) -> Self {
        if fonts.is_empty() {
            return Self::new(source_text, images, font_paths);
        }

        let fallback = if font_paths.is_empty() {
            FallbackFontData::Cached(get_system_fonts())
        } else {
            FallbackFontData::Shared(get_fonts_for_extra_paths(font_paths))
        };
        Self::new_with_font_source(
            source_text,
            images,
            FontSource::InMemory(InMemoryFontData::new(fonts, fallback)),
        )
    }

    fn new_with_font_source(
        source_text: &str,
        images: &[ImageAsset],
        font_source: FontSource,
    ) -> Self {
        let main_id = main_file_id();
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
            shaped_faces: Mutex::new(HashMap::new()),
        }
    }
}

/// The id each compilation registers its single source file under.
fn main_file_id() -> FileId {
    let path = VirtualPath::new("main.typ").expect("`main.typ` is a valid virtual path");
    FileId::new(RootedPath::new(VirtualRoot::Project, path))
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
            Err(typst::diag::FileError::NotFound(std::path::PathBuf::from(
                id.vpath().get_without_slash(),
            )))
        }
    }

    fn file(&self, id: FileId) -> FileResult<Bytes> {
        if id == self.source.id() {
            Ok(Bytes::new(self.source.text().as_bytes().to_vec()))
        } else {
            // Check if it's an embedded image file
            let path: &str = id.vpath().get_without_slash();
            if let Some(data) = self.images.get(path) {
                Ok(data.clone()) // Bytes::clone is cheap (reference-counted)
            } else {
                Err(typst::diag::FileError::NotFound(std::path::PathBuf::from(
                    id.vpath().get_without_slash(),
                )))
            }
        }
    }

    fn font(&self, index: usize) -> Option<Font> {
        if let Some(cached) = self
            .shaped_faces
            .lock()
            .expect("shaped face cache mutex should not be poisoned")
            .get(&index)
        {
            return Some(cached.clone());
        }

        let loaded: Font = self.font_source.font(index)?;
        // Office positions text through CoreText, which keeps reading a
        // TrueType face's legacy `kern` table where the shaper would take the
        // GPOS feature instead (issue #1116).
        let font: Font = super::font_kern::face_preferring_legacy_kern(&loaded).unwrap_or(loaded);
        self.shaped_faces
            .lock()
            .expect("shaped face cache mutex should not be poisoned")
            .insert(index, font.clone());
        Some(font)
    }

    fn today(&self, _offset: Option<Duration>) -> Option<Datetime> {
        None
    }
}

#[cfg(test)]
#[path = "pdf_tests.rs"]
mod tests;

/// The face the compiler will shape `family` with.
///
/// The declared name may be one the font book does not register — a localized
/// East Asian family, or a face the machine does not have — so the same alias
/// and substitute chain rendering resolves through is walked here (issue #575).
/// Resolving through the same font set the compiler uses also primes the
/// compile-time cache.
#[cfg(not(target_arch = "wasm32"))]
fn best_face(family: &str) -> Option<typst::text::Font> {
    let data = active_font_data();
    let variant: typst::text::FontVariant = chain_variant(family, None);
    first_face_in_chain(family, variant, |candidate| {
        select_face_index(&data, candidate, variant)
            .and_then(|index| data.fonts.get(index))
            .and_then(|slot| slot.get())
    })
}

/// The variant a chain walk resolves `family` at.
///
/// A family name can state a weight in its style suffix — `Calibri Light`,
/// `Segoe UI Semibold`, `Arial Black` — and the font book files that face
/// under the base family with the suffix trimmed, so the chain reaches it only
/// through the base family, and only the weight tells the member apart from
/// the family's regular one. Their vertical metrics need not agree: Arial
/// Black declares an `hhea` ascender of 2254/2048 against Arial's 1854/2048,
/// and twelve single-spaced 20pt `Arial Black` paragraphs advance 28.189pt in
/// a native Word export — Arial Black's own 1.4102em sum, not Arial's
/// 1.1499em. Resolving the chain at the regular variant would read a line box
/// 18% short (issue #1643).
///
/// `explicit_weight` is the weight the *run* states on top of the name, and
/// the heavier of the two wins — the composition `write_text_params` already
/// emits, so the face measured is the face painted. A bold cell in `Segoe UI
/// Semibold` paints the family's bold member, and so measures it.
#[cfg(not(target_arch = "wasm32"))]
fn chain_variant(
    family: &str,
    explicit_weight: Option<typst::text::FontWeight>,
) -> typst::text::FontVariant {
    let stated: Option<typst::text::FontWeight> =
        super::font_subst::weight_stated_by_family_name(family);
    let weight: typst::text::FontWeight = match (stated, explicit_weight) {
        (Some(stated), Some(explicit)) => stated.max(explicit),
        (Some(stated), None) => stated,
        (None, Some(explicit)) => explicit,
        (None, None) => typst::text::FontVariant::default().weight,
    };
    typst::text::FontVariant {
        weight,
        ..typst::text::FontVariant::default()
    }
}

/// The first face the alias and substitute chain of `family` resolves to at
/// `variant`'s weight, taking each candidate's conversion-local in-memory
/// face before `on_disk`.
///
/// This is the order the compiler selects in: `MinimalWorld` prepends the
/// conversion's in-memory faces to the same fallback book, so a family held in
/// memory outranks a copy on disk, but a candidate later in the chain never
/// outranks one that resolves earlier. Walking the whole chain against the
/// in-memory faces first broke that: a deck naming Avenir Next LT Pro loads
/// the bundled Noto Serif into memory as that family's reproducible fallback,
/// and every metric lookup answered with Noto Serif's line box even when the
/// exact face was supplied on `--font-path` and painted the run, seating the
/// block two points high (issue #1629).
///
/// `on_disk` must itself resolve `candidate` at `variant`'s weight — this
/// only fixes the in-memory step's weight, since the disk step's lookup
/// (a plain book `select`, or the shadowed-system-face walk in
/// [`line_metric_face`]) differs by caller.
#[cfg(not(target_arch = "wasm32"))]
fn first_face_in_chain(
    family: &str,
    variant: typst::text::FontVariant,
    on_disk: impl Fn(&str) -> Option<typst::text::Font>,
) -> Option<typst::text::Font> {
    super::font_subst::family_candidates(family)
        .iter()
        .find_map(|candidate| {
            super::font_subst::active_in_memory_font_named(candidate, variant)
                .or_else(|| on_disk(candidate))
        })
}

/// Index into `data` of the face registered under exactly `candidate` that
/// sits nearest `variant`'s weight.
#[cfg(not(target_arch = "wasm32"))]
fn select_face_index(
    data: &CachedFontData,
    candidate: &str,
    variant: typst::text::FontVariant,
) -> Option<usize> {
    data.book.select(&candidate.to_lowercase(), variant)
}

/// The font set the compiler shapes with: the caller's search paths when a
/// conversion has set them, else the Office bundle, each ahead of the system.
#[cfg(not(target_arch = "wasm32"))]
fn active_font_data() -> Arc<CachedFontData> {
    let search_paths = super::font_subst::active_font_search_paths()
        .unwrap_or_else(|| super::font_context::default_font_search_paths().to_vec());
    get_fonts_for_extra_paths(&search_paths)
}

/// The face Word measures a line box by, which is not always the face it
/// draws with.
///
/// On macOS each Office application registers its bundled `DFonts` beside the
/// system's copies, and when a system-shipped face shares the bundled face's
/// PostScript name, Word paces the line on the *system* face's `hhea` while
/// its export embeds the bundled outlines. Measured on Times New Roman: the
/// bundle's `times.ttf` (Version 7.00;O365) declares an `hhea` line gap of 0
/// and the system's `Times New Roman.ttf` (5.01) declares 87/2048, and a
/// native Word export embeds the version 7 outlines yet advances 12pt lines by
/// 13.80pt = 2355/2048em — the system copy's sum. The same file installed
/// under another family name, so that no system copy shares its name, paces at
/// its own 2268/2048em (issue #1284).
///
/// Only a *system-shipped* face shadows the bundle. A user-installed copy in
/// `~/Library/Fonts` sharing the bundled face's PostScript name left Word's
/// Rockwell on the bundled metrics even after a relaunch, and a system face
/// with a different PostScript name (Apple's `Rockwell-Regular` beside the
/// bundle's `Rockwell`, Apple's `Symbol` beside the bundle's `SymbolMT`)
/// shadows nothing. `/Library/Fonts` was not measured and is treated as not
/// shadowing.
///
/// Returns the shaped face itself whenever no such system copy exists, so
/// every host without an Office bundle is unaffected.
#[cfg(not(target_arch = "wasm32"))]
fn line_metric_face(family: &str) -> Option<typst::text::Font> {
    let data = active_font_data();
    let variant: typst::text::FontVariant = chain_variant(family, None);
    first_face_in_chain(family, variant, |candidate| {
        let shaped_index: usize = select_face_index(&data, candidate, variant)?;
        line_metric_face_at(
            &data,
            super::font_context::default_font_search_paths(),
            macos_system_font_dirs(),
            shaped_index,
        )
    })
}

/// The face Word measures the line box of the face at `shaped_index` by: the
/// system copy that shadows it, if one does, else the face itself.
/// `office_font_dirs` hold the application-bundled faces and
/// `system_font_dirs` the faces that may shadow them.
#[cfg(not(target_arch = "wasm32"))]
fn line_metric_face_at(
    data: &CachedFontData,
    office_font_dirs: &[PathBuf],
    system_font_dirs: &[PathBuf],
    shaped_index: usize,
) -> Option<typst::text::Font> {
    let shaped_slot = data.fonts.get(shaped_index)?;
    let shaped: typst::text::Font = shaped_slot.get()?;
    let is_under =
        |path: &std::path::Path, dirs: &[PathBuf]| dirs.iter().any(|dir| path.starts_with(dir));
    let Some(shaped_path) = shaped_slot.path() else {
        return Some(shaped);
    };
    if !is_under(shaped_path, office_font_dirs) {
        return Some(shaped);
    }
    let Some(shaped_name) = postscript_name(&shaped) else {
        return Some(shaped);
    };
    let family_key: String = data.book.info(shaped_index)?.family.to_lowercase();
    let shadowing: Option<typst::text::Font> = data
        .book
        .select_family(&family_key)
        .filter(|&index| index != shaped_index)
        .filter_map(|index| {
            let slot = data.fonts.get(index)?;
            let path = slot.path()?;
            (is_under(path, system_font_dirs) && !is_under(path, office_font_dirs))
                .then(|| slot.get())
                .flatten()
        })
        .find(|candidate| postscript_name(candidate).as_deref() == Some(shaped_name.as_str()));
    if shadowing.is_some() {
        tracing::debug!(
            family = %family_key,
            postscript_name = %shaped_name,
            bundled = %shaped_path.display(),
            "line metrics follow the system face that shadows the Office-bundled copy"
        );
    }
    shadowing.or(Some(shaped))
}

/// The face's `name` table PostScript name (ID 6), the key macOS registers a
/// font under.
#[cfg(not(target_arch = "wasm32"))]
fn postscript_name(font: &typst::text::Font) -> Option<String> {
    const POST_SCRIPT_NAME_ID: u16 = 6;
    measured_instance(font)
        .ttf()
        .names()
        .into_iter()
        .filter(|name| name.name_id == POST_SCRIPT_NAME_ID)
        .find_map(|name| name.to_string())
}

/// Where macOS ships its own faces. Only these shadow an Office bundle's copy
/// of the same face (see [`line_metric_face`]); the per-user font directory
/// measurably does not.
#[cfg(not(target_arch = "wasm32"))]
fn macos_system_font_dirs() -> &'static [PathBuf] {
    static DIRS: std::sync::LazyLock<Vec<PathBuf>> = std::sync::LazyLock::new(|| {
        if cfg!(target_os = "macos") {
            vec![PathBuf::from("/System/Library/Fonts")]
        } else {
            Vec::new()
        }
    });
    &DIRS
}

#[cfg(target_arch = "wasm32")]
fn best_face(family: &str) -> Option<typst::text::Font> {
    // Mirrors the native arm's weight rule: a weight-suffixed name reaches its
    // member only through the base family, at the weight the name states
    // (issue #1643).
    let weight: typst::text::FontWeight = super::font_subst::weight_stated_by_family_name(family)
        .unwrap_or_else(|| typst::text::FontVariant::default().weight);
    super::font_subst::active_in_memory_font(
        family,
        typst::text::FontVariant {
            weight,
            ..typst::text::FontVariant::default()
        },
    )
}

/// Look a per-family `f64` metric up through a process-wide cache.
///
/// Font resolution walks the substitute chain and opens the face, so every
/// caller that asks the same question twice would pay for it twice.
#[cfg(not(target_arch = "wasm32"))]
fn cached_family_metric(
    cache: &OnceLock<Mutex<HashMap<String, Option<f64>>>>,
    family: &str,
    resolve: fn(&str) -> Option<typst::text::Font>,
    compute: impl FnOnce(&typst::text::Font) -> Option<f64>,
) -> Option<f64> {
    if super::font_subst::active_font_search_paths().is_some() {
        return resolve(family).and_then(|font| compute(&font));
    }

    let cache = cache.get_or_init(|| Mutex::new(HashMap::new()));
    let key: String = family.to_lowercase();
    if let Some(cached) = cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .get(&key)
    {
        return *cached;
    }

    let value: Option<f64> = resolve(family).and_then(|font| compute(&font));
    cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .insert(key, value);
    value
}

/// The `hhea` ascender of the face Word measures `family` by, in em units
/// (see [`line_metric_face`]).
///
/// The bare ascender, with the `hhea` line gap left out — deliberately *not*
/// [`font_line_metrics_em`]'s first element, which folds the gap in. Callers
/// that want Word's line top add the gap back themselves: a header seat does
/// (`typst_gen_text::word_header_line_ascent_em`, issue #1640), and an East
/// Asian line box does not, because it leaves the gap out at both ends (issue
/// #1638). Splitting it here is what lets those two share one ascender read.
///
/// Read out of the `hhea` table directly rather than through
/// `ttf_parser::Face::ascender`, whose name promises `hhea` but which returns
/// OS/2 `sTypoAscender` whenever `fsSelection` sets `USE_TYPO_METRICS`: 84 of
/// the 1109 faces installed on the calibration machine set that bit, and on 8
/// of them — Cambria Math among them, at 0.9502em against 0.7778em — the two
/// tables disagree, so the alias silently answers a different question.
///
/// Which table Word measures by is *not* settled by the corpus. Both calibrated
/// faces, Arial (0.9053em) and Malgun Gothic (1.0884em), carry `usWinAscent`
/// equal to their `hhea` ascender, and no header in the corpus uses a face where
/// they differ — though 407 of the 1109 local faces do, by up to 0.2275em
/// (Candara: 0.7246 against 0.9521). `hhea` is chosen because the ground truth
/// is macOS Word, whose text stack reports a face's ascent from `hhea`, while
/// `usWinAscent` is the GDI quantity; and because [`font_line_metrics_em`]
/// builds Word's single-line pitch from `hhea`'s ascender, descender and line
/// gap, calibrated to 0.0005em on Arial (issue #508) — taking the two halves of
/// one line box from two different tables would be the odd choice.
///
/// TODO(the corpus cannot separate `hhea` from `usWinAscent`): a native macOS
/// Word export of a Candara header would settle it outright. The attempt in this
/// session could not — every `save as … format PDF` died with AppleEvent -1712
/// and wrote a 0-byte file, across seven attempts and a clean relaunch. Note also that `font_line_metrics_em`
/// still reads its ascender through the alias, so the two disagree on those 8
/// faces until it is given the same explicit read.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn font_hhea_ascender_em(family: &str) -> Option<f64> {
    static ASCENDER_CACHE: OnceLock<Mutex<HashMap<String, Option<f64>>>> = OnceLock::new();
    cached_family_metric(&ASCENDER_CACHE, family, line_metric_face, |font| {
        let instance = measured_instance(font);
        let ttf = instance.ttf();
        Some(f64::from(ttf.tables().hhea.ascender) / f64::from(ttf.units_per_em()).max(1.0))
    })
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn font_hhea_ascender_em(family: &str) -> Option<f64> {
    let font = best_face(family)?;
    let instance = measured_instance(&font);
    let ttf = instance.ttf();
    Some(f64::from(ttf.tables().hhea.ascender) / f64::from(ttf.units_per_em()).max(1.0))
}

/// The cap height of the best face for `family`, in em units.
///
/// This is the ascent the compiler itself gives a text line: `top-edge` defaults
/// to `"cap-height"`, and the value is taken from the very `Font` the compile
/// will shape with, so it tracks Typst's own fallbacks (`OS/2 sCapHeight`, else
/// the typographic ascender, else `hhea`) instead of restating them. Reading it
/// is what lets the header band be shifted by the *difference* between Word's
/// seat and the compiler's, leaving every line box — and therefore the story's
/// baseline-to-baseline advance — exactly as the compiler would lay it out
/// (issue #629). The advance is set separately, as a story-level
/// `par(leading:)` that tops this cap-height edge up to Word's pitch (issue
/// #735) — reading the cap height here is what keeps the two independent, since
/// the leading is computed as the remainder after it.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn font_cap_height_em(family: &str) -> Option<f64> {
    static CAP_HEIGHT_CACHE: OnceLock<Mutex<HashMap<String, Option<f64>>>> = OnceLock::new();
    cached_family_metric(&CAP_HEIGHT_CACHE, family, best_face, |font| {
        Some(measured_instance(font).metrics().cap_height.get())
    })
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn font_cap_height_em(family: &str) -> Option<f64> {
    best_face(family).map(|font| measured_instance(&font).metrics().cap_height.get())
}

/// The bare `hhea` line gap of the face Word measures `family` by, in em
/// units (see [`line_metric_face`]).
///
/// [`font_line_metrics_em`] folds the gap into its first element, because that
/// is where Word puts the baseline. Excel does not: it rounds the ascender,
/// the line gap and the descender into whole points *separately* before it
/// composes a printed sheet cell's line box, so that path needs the gap on its
/// own (issue #1161). Word's East Asian line box needs it for the opposite
/// reason — to take the gap back out of the folded pair (issue #1638).
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn font_line_gap_em(family: &str) -> Option<f64> {
    use std::collections::HashMap;
    use std::sync::Mutex;
    static LINE_GAP_CACHE: OnceLock<Mutex<HashMap<String, Option<f64>>>> = OnceLock::new();
    cached_family_metric(&LINE_GAP_CACHE, family, line_metric_face, |font| {
        let instance = measured_instance(font);
        let ttf = instance.ttf();
        let upem = f64::from(ttf.units_per_em()).max(1.0);
        Some(f64::from(ttf.line_gap()) / upem)
    })
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn font_line_gap_em(family: &str) -> Option<f64> {
    let font = best_face(family)?;
    let instance = measured_instance(&font);
    let ttf = instance.ttf();
    let upem = f64::from(ttf.units_per_em()).max(1.0);
    Some(f64::from(ttf.line_gap()) / upem)
}

/// Line metrics of the face Word measures `family` by, in em units:
/// `(above baseline, below baseline, Word single-line pitch)`.
///
/// That is [`line_metric_face`], not always [`best_face`]: where a
/// system-shipped copy shadows the Office bundle's, Word paces on the system
/// copy's `hhea` (issue #1284).
///
/// The first two split the third at the point Word puts the baseline —
/// `hhea` ascender plus line gap — so they always sum to the pitch. Typst's
/// own `metrics.ascender`/`descender` pair cannot express that split: it is
/// normalised, summing to exactly 1.0 for faces like Malgun Gothic, which put
/// the baseline 0.2em off where Word does (issue #508).
///
/// The pitch itself is the `hhea` ascender + descender + line gap sum that
/// Word uses for "single" line spacing (issue #354). An East Asian line is the
/// exception at both ends: it leaves the gap out of the pitch *and* out of the
/// seat, so that path takes this triple through
/// `typst_gen_text::word_line_metrics_em` (issue #1638).
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn font_line_metrics_em(family: &str) -> Option<(f64, f64, f64)> {
    use std::collections::HashMap;
    use std::sync::Mutex;
    type LineMetricsEm = Option<(f64, f64, f64)>;
    static METRICS_CACHE: OnceLock<Mutex<HashMap<String, LineMetricsEm>>> = OnceLock::new();

    let metrics_for = |font: &typst::text::Font| {
        let instance = measured_instance(font);
        let ttf = instance.ttf();
        let upem = f64::from(ttf.units_per_em()).max(1.0);
        let hhea_pitch_em = (f64::from(ttf.ascender()) - f64::from(ttf.descender())
            + f64::from(ttf.line_gap()))
            / upem;
        let top_em: f64 = (f64::from(ttf.ascender()) + f64::from(ttf.line_gap())) / upem;
        (top_em, hhea_pitch_em - top_em, hhea_pitch_em)
    };
    if super::font_subst::active_font_search_paths().is_some() {
        return line_metric_face(family).map(|font| metrics_for(&font));
    }

    let cache = METRICS_CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let key: String = family.to_lowercase();
    if let Some(cached) = cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .get(&key)
    {
        return *cached;
    }

    let metrics: Option<(f64, f64, f64)> = line_metric_face(family).map(|font| {
        // Word seats the baseline `hhea ascender + lineGap` below the top
        // of the line, not at the font's ascender/descender proportion of
        // the box — measured to 0.0005em on Arial (issue #508). Typst's
        // `metrics` pair is normalised (Malgun Gothic's sums to exactly
        // 1.0) and cannot express that, so the split comes from `hhea`.
        //
        // Safe to change only since #512: the document-grid arms used to
        // derive their line count from this pair, so altering it silently
        // repaginated. They now choose between the grid pitch and the
        // natural line without consulting it.
        metrics_for(&font)
    });
    cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .insert(key, metrics);
    metrics
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn font_line_metrics_em(family: &str) -> Option<(f64, f64, f64)> {
    let font = best_face(family)?;
    let instance = measured_instance(&font);
    let ttf = instance.ttf();
    let upem = f64::from(ttf.units_per_em()).max(1.0);
    let hhea_pitch_em =
        (f64::from(ttf.ascender()) - f64::from(ttf.descender()) + f64::from(ttf.line_gap())) / upem;
    let top_em = (f64::from(ttf.ascender()) + f64::from(ttf.line_gap())) / upem;
    Some((top_em, hhea_pitch_em - top_em, hhea_pitch_em))
}

/// Maximum horizontal advance over the digits U+0030..=U+0039 of the face
/// for `family` at the requested weight, in em units.
///
/// Excel derives every column print metric from this value of the face it
/// resolves for the workbook Normal font: 17 one-factor native Excel-for-Mac
/// probes show the column character-unit is `round_half_up(advance × size)`
/// integer points, matching each face's real `hmtx` maximum (issue #621).
/// This resolves families outside the parser's reference table — the same
/// alias and substitute chain rendering uses, so the metric tracks the face
/// the glyphs will actually come from.
///
/// `bold` must reflect the *cell's own* weight, not just its family: a
/// family's digits are not always the same width at every weight. Cambria's
/// bold digits measure 1213/2048em against 1134/2048em regular (~7% wider)
/// and Verdana's 1456/2048em against 1302/2048em (~12%), while Calibri,
/// Arial and Times New Roman keep identical digit widths across weight. A
/// 42pt bold Cambria title priced on its regular digit width undershoots its
/// whole-point left inset by a full step, 7pt instead of 8pt, against a
/// native Excel for Mac export (issue #1623).
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn max_digit_advance_em(family: &str, bold: bool) -> Option<f64> {
    static ADVANCE_CACHE: OnceLock<FaceAdvanceCache> = OnceLock::new();

    let variant: typst::text::FontVariant =
        chain_variant(family, bold.then_some(typst::text::FontWeight::BOLD));

    face_advance_em(family, variant, &ADVANCE_CACHE, |font| {
        let instance = measured_instance(font);
        let ttf = instance.ttf();
        let upem: f64 = f64::from(ttf.units_per_em()).max(1.0);
        ('0'..='9')
            .filter_map(|digit| {
                ttf.glyph_index(digit)
                    .and_then(|glyph| ttf.glyph_hor_advance(glyph))
            })
            .map(|glyph_advance| f64::from(glyph_advance) / upem)
            .fold(None, |widest: Option<f64>, advance_em: f64| {
                Some(widest.map_or(advance_em, |width| width.max(advance_em)))
            })
    })
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn max_digit_advance_em(_family: &str, _bold: bool) -> Option<f64> {
    None
}

/// Horizontal advance of the space U+0020 on the best face for `family`, in
/// em units.
///
/// Excel prices a cell alignment's `indent` at three spaces of the workbook
/// Normal font, each rounded to a whole point: eleven one-factor native
/// Excel-for-Mac exports fix the unit at 3 pt for Calibri 6 and 21 pt for
/// Courier New 11 (issue #1109). Resolved through the same alias and
/// substitute chain rendering uses, so the metric tracks the face the glyphs
/// will actually come from.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn space_advance_em(family: &str) -> Option<f64> {
    static ADVANCE_CACHE: OnceLock<FaceAdvanceCache> = OnceLock::new();

    face_advance_em(
        family,
        chain_variant(family, None),
        &ADVANCE_CACHE,
        |font| {
            let instance = measured_instance(font);
            let ttf = instance.ttf();
            let upem: f64 = f64::from(ttf.units_per_em()).max(1.0);
            ttf.glyph_index(' ')
                .and_then(|glyph| ttf.glyph_hor_advance(glyph))
                .map(|advance| f64::from(advance) / upem)
        },
    )
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn space_advance_em(_family: &str) -> Option<f64> {
    None
}

/// A cached `hmtx`-derived metric, keyed by `(family, is_bold)`.
#[cfg(not(target_arch = "wasm32"))]
type FaceAdvanceCache = std::sync::Mutex<std::collections::HashMap<(String, bool), Option<f64>>>;

/// One `hmtx`-derived metric of the face `family` resolves to at `variant`'s
/// weight, cached per `(family, is_bold)` in `cache`.
///
/// `is_bold` narrows the key rather than defining it: the family name is part
/// of the key, so two weight-suffixed members of one family — resolved at
/// different weights by [`chain_variant`] — never share an entry.
///
/// Shared by the digit and space metrics so both walk the same resolution
/// order as [`best_face`]: the conversion's in-memory fonts and search paths
/// when one is active, else the font set the compiler itself will use.
#[cfg(not(target_arch = "wasm32"))]
fn face_advance_em(
    family: &str,
    variant: typst::text::FontVariant,
    cache: &'static OnceLock<FaceAdvanceCache>,
    advance_for: impl Fn(&typst::text::Font) -> Option<f64>,
) -> Option<f64> {
    use std::collections::HashMap;
    use std::sync::Mutex;

    let is_bold: bool = variant.weight == typst::text::FontWeight::BOLD;

    if super::font_subst::active_font_search_paths().is_some() {
        let data = active_font_data();
        // Shares `first_face_in_chain` with `best_face` (issue #1629's
        // interleaved in-memory-then-disk order per candidate), so this can
        // never drift from that order the way two independent copies could.
        // Every caller composes its weight through `chain_variant`, so a
        // weight-suffixed family measures the member its name denotes here too
        // (issue #1643).
        return first_face_in_chain(family, variant, |candidate| {
            data.book
                .select(&candidate.to_lowercase(), variant)
                .and_then(|index| data.fonts.get(index))
                .and_then(|slot| slot.get())
        })
        .and_then(|font| advance_for(&font));
    }

    let cache = cache.get_or_init(|| Mutex::new(HashMap::new()));
    let key: (String, bool) = (family.to_lowercase(), is_bold);
    if let Some(cached) = cache
        .lock()
        .expect("face advance cache mutex should not be poisoned")
        .get(&key)
    {
        return *cached;
    }

    // Use the same font set the compiler will use (system + discovered
    // Office font dirs); this also primes the compile-time cache.
    let data = get_fonts_for_extra_paths(super::font_context::default_font_search_paths());
    let advance: Option<f64> = super::font_subst::family_candidates(family)
        .iter()
        .find_map(|candidate| data.book.select(&candidate.to_lowercase(), variant))
        .and_then(|index| data.fonts.get(index))
        .and_then(|slot| slot.get())
        .and_then(|font| advance_for(&font));
    cache
        .lock()
        .expect("face advance cache mutex should not be poisoned")
        .insert(key, advance);
    advance
}

/// Total horizontal advance of `text`, in em units, on the face `family`
/// resolves to at the requested weight.
///
/// Word's auto table layout never compresses a column below its min-content
/// width — the advance of its widest unbreakable token — so the DOCX parser
/// needs advances from the same faces the renderer will draw with
/// (issue #624). The face is resolved once per `(family, weight)` through the
/// same alias and substitute chain rendering uses and cached; bold runs must
/// measure against the bold face because its advances differ (Libertinus
/// Serif's "Total" is 2.392em bold against 2.138em regular).
///
/// Kerning and ligatures are deliberately ignored: the per-glyph `hmtx` sum
/// reproduced Word's invoice column widths within 0.10pt, and callers assert
/// with tolerances, so the ≲1-2% shaping error is acceptable. Returns `None`
/// when no face resolves or any character lacks a glyph, so the caller can
/// degrade to a measurement-free path.
///
/// Per-conversion in-memory faces are consulted before the process-wide cache.
/// Native-only `ConvertOptions::font_paths` and materialized document-embedded
/// directories are still not visible to parser-time measurements; a family
/// only those paths provide degrades to the measurement-free path.
/// TODO(issue #624): include native path-provided faces in these measurement
/// helpers as well.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn text_advance_em(family: &str, bold: bool, text: &str) -> Option<f64> {
    Some(glyph_advances_em(family, bold, text)?.iter().sum())
}

/// Each character of `text` measured on its own, in em units, on the face
/// `family` resolves to at the requested weight.
///
/// [`text_advance_em`] is this sum. A caller that quantizes advances one glyph
/// at a time — Excel rounds every one to a whole point before accumulating it
/// (issue #1088) — cannot work from the sum, because rounding the total is a
/// different number from the total of the rounded parts.
///
/// The same caveats hold: kerning and ligatures are ignored, and `None` comes
/// back when no face resolves or any character lacks a glyph.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn glyph_advances_em(family: &str, bold: bool, text: &str) -> Option<Vec<f64>> {
    use std::collections::HashMap;
    use std::sync::Mutex;
    type ResolvedFaceCache = HashMap<(String, bool), Option<typst::text::Font>>;
    static FACE_CACHE: OnceLock<Mutex<ResolvedFaceCache>> = OnceLock::new();

    let variant: typst::text::FontVariant =
        chain_variant(family, bold.then_some(typst::text::FontWeight::BOLD));
    let active_font = super::font_subst::active_in_memory_font(family, variant);

    let cache = FACE_CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let key: (String, bool) = (family.to_lowercase(), bold);
    let cached_font: Option<Option<typst::text::Font>> = cache
        .lock()
        .expect("resolved face cache mutex should not be poisoned")
        .get(&key)
        .cloned();
    let font: Option<typst::text::Font> = match active_font {
        Some(font) => Some(font),
        None => match cached_font {
            Some(font) => font,
            None => {
                // Use the same font set the compiler will use (system + discovered
                // Office font dirs); this also primes the compile-time cache.
                let data =
                    get_fonts_for_extra_paths(super::font_context::default_font_search_paths());
                let resolved: Option<typst::text::Font> =
                    super::font_subst::family_candidates(family)
                        .iter()
                        .find_map(|candidate| data.book.select(&candidate.to_lowercase(), variant))
                        .and_then(|index| data.fonts.get(index))
                        .and_then(|slot| slot.get());
                cache
                    .lock()
                    .expect("resolved face cache mutex should not be poisoned")
                    .insert(key, resolved.clone());
                resolved
            }
        },
    };

    font_glyph_advances_em(&font?, text)
}

/// Measure `text` through the exact family list Typst receives, including its
/// built-in fallback families and the font-book fallback it selects after
/// those names are exhausted.
///
/// A DOCX can declare a face absent from the host. Typst then calls
/// [`typst::text::FontBook::select_fallback`], while a family-only metric lookup
/// returns `None`. Overlong-token wrapping needs the former answer because the
/// fallback's glyph widths decide Word's character boundary (issue #1454).
pub(crate) fn glyph_advances_em_with_typst_fallback(
    families: &[String],
    bold: bool,
    text: &str,
) -> Option<Vec<f64>> {
    let variant = typst::text::FontVariant {
        weight: if bold {
            typst::text::FontWeight::BOLD
        } else {
            typst::text::FontWeight::REGULAR
        },
        ..typst::text::FontVariant::default()
    };

    #[cfg(not(target_arch = "wasm32"))]
    let native_data = {
        let search_paths = super::font_subst::active_font_search_paths()
            .unwrap_or_else(|| super::font_context::default_font_search_paths().to_vec());
        get_fonts_for_extra_paths(&search_paths)
    };
    #[cfg(not(target_arch = "wasm32"))]
    let fallback_data: &CachedFontData = native_data.as_ref();
    #[cfg(target_arch = "wasm32")]
    let fallback_data: &CachedFontData = get_embedded_fonts();

    let in_memory_fonts: Vec<typst::text::Font> = super::font_subst::active_in_memory_fonts();
    if in_memory_fonts.is_empty() {
        return select_typst_font_advances(
            &fallback_data.book,
            |index| fallback_data.fonts.get(index).and_then(|slot| slot.get()),
            families,
            variant,
            text,
        );
    }

    // `MinimalWorld` prepends the conversion-local faces to this same
    // fallback book. Rebuild just the lightweight metadata index so selection
    // and tie-breaking stay identical without eagerly loading every slot.
    let in_memory_count = in_memory_fonts.len();
    let book = typst::text::FontBook::from_infos(
        in_memory_fonts
            .iter()
            .map(|font| font.info().clone())
            .chain(
                (0..fallback_data.fonts.len())
                    .filter_map(|index| fallback_data.book.info(index).cloned()),
            ),
    );
    select_typst_font_advances(
        &book,
        |index| {
            if index < in_memory_count {
                in_memory_fonts.get(index).cloned()
            } else {
                fallback_data
                    .fonts
                    .get(index - in_memory_count)
                    .and_then(|slot| slot.get())
            }
        },
        families,
        variant,
        text,
    )
}

fn select_typst_font_advances(
    book: &typst::text::FontBook,
    mut load: impl FnMut(usize) -> Option<typst::text::Font>,
    families: &[String],
    variant: typst::text::FontVariant,
    text: &str,
) -> Option<Vec<f64>> {
    const TYPST_IMPLICIT_FAMILIES: [&str; 5] = [
        "Libertinus Serif",
        "Twitter Color Emoji",
        "Noto Color Emoji",
        "Apple Color Emoji",
        "Segoe UI Emoji",
    ];

    let mut first_selected: Option<usize> = None;
    for family in families
        .iter()
        .map(String::as_str)
        .chain(TYPST_IMPLICIT_FAMILIES)
    {
        let Some(index) = book.select(&family.to_lowercase(), variant) else {
            continue;
        };
        first_selected.get_or_insert(index);
        let font = load(index)?;
        if let Some(advances) = font_glyph_advances_em(&font, text) {
            return Some(advances);
        }
        // A run needing several faces has per-segment shaping and kerning that
        // this hmtx-only helper cannot reproduce safely. Basic Latin absent
        // from a selected face can still continue to Typst's fallback below.
        let instance = measured_instance(&font);
        if text
            .chars()
            .any(|character| instance.ttf().glyph_index(character).is_some())
        {
            return None;
        }
    }

    let like = first_selected.and_then(|index| book.info(index));
    let fallback_index = book.select_fallback(like, variant, text)?;
    let fallback = load(fallback_index)?;
    font_glyph_advances_em(&fallback, text)
}

fn font_glyph_advances_em(font: &typst::text::Font, text: &str) -> Option<Vec<f64>> {
    let instance = measured_instance(font);
    let ttf = instance.ttf();
    let upem: f64 = f64::from(ttf.units_per_em()).max(1.0);
    let mut advances_em: Vec<f64> = Vec::with_capacity(text.chars().count());
    for character in text.chars() {
        let glyph_advance: u16 = ttf
            .glyph_index(character)
            .and_then(|glyph| ttf.glyph_hor_advance(glyph))?;
        advances_em.push(f64::from(glyph_advance) / upem);
    }
    Some(advances_em)
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn text_advance_em(_family: &str, _bold: bool, _text: &str) -> Option<f64> {
    None
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn glyph_advances_em(_family: &str, _bold: bool, _text: &str) -> Option<Vec<f64>> {
    None
}

/// PowerPoint's line height factor: it gives every line 1.2 times the font
/// size, whatever the font's own metrics say.
///
/// Measured on native exports from both platforms. On macOS PowerPoint, one
/// wrapped Arial paragraph advances `(last baseline - first) / 14 = 20.3657pt`
/// at 17pt (1.1980em) and `/ 18 = 28.8400pt` at 24pt (1.2017em) — the division
/// cancels the export's 0.24pt position grid, which is what made small samples
/// look font-dependent. Windows PowerPoint measures the same 1.2000em flat for
/// Arial, Calibri, and Malgun Gothic (issues #485, #513).
pub(crate) const POWERPOINT_LINE_HEIGHT_FACTOR: f64 = 1.2;

/// The `(above baseline, below baseline)` split of a unit line for one face,
/// from the OS/2 `usWinAscent`/`usWinDescent` pair PowerPoint measures it by.
///
/// `None` for a face that declares no ascent, which no split could seat.
fn powerpoint_face_unit_line(ascent: f64, descent: f64) -> Option<(f64, f64)> {
    let natural: f64 = ascent + descent;
    if ascent <= 0.0 || natural <= 0.0 {
        return None;
    }
    Some((ascent / natural, descent / natural))
}

/// PowerPoint's `(above baseline, below baseline)` split of its 1.2em line for
/// a line set in `faces`, each given as a positive `(ascent, descent)` pair of
/// em fractions.
///
/// **Every font on the line shares one box, and each is normalised to a unit
/// line before they are compared.** The line takes the largest normalised
/// ascent and the largest normalised descent of its faces, then scales that
/// pair back down to the 1.2em the advance is fixed at. With a single face the
/// two steps cancel and the share is just `1.2 x ascent / (ascent + descent)`.
///
/// **The metrics are OS/2's `usWin*` pair, and the hhea line gap plays no
/// part.** Both halves of that were measured together, because a corpus frame
/// cannot separate them: a native PowerPoint 16.111 probe deck of 14
/// top-anchored, zero-inset boxes at 11-61pt, exported once per face, puts
/// Arial (hhea 1854/-434/**67**, usWin 1854/434 per 2048 upem) on 0.97238em —
/// its gap-free share — at all 14 sizes, and the gap-inclusive 0.94471em
/// outside the export's 0.12pt half-grid at 12 of them. Nine further faces land
/// on their own gap-free share the same way, and Yu Gothic — the one whose
/// usWin pair differs from hhea's (2017/619 against 1802/-455/1024) — sides
/// with usWin, which hhea cannot express at all (issue #1176).
///
/// **The paragraph mark counts on the final physical line.** The same
/// single-line probe with only `<a:endParaRPr>` varied moves every seat: an Arial run
/// whose mark is Calibri (usWin 1950/550) seats at 0.94377em, one whose mark is
/// Verdana at 0.97621em, Malgun Gothic 0.97420em, MS Gothic 0.98306em and
/// Meiryo 0.88107em — 70 cells, all inside the half-grid, none of them the run
/// face's own share. That is why the golden mocks read as Arial's *gap-inclusive*
/// share for 500 issues: their marks carry no typeface and fall to the theme's
/// minor Latin font, Calibri, whose deeper descent pushes the shared box down
/// to 0.94377em — 0.001em from the gap-inclusive 0.94471em #1118 fitted, and
/// on the same whole point at every size those decks use.
///
/// A face that **overflows** the box is shared like any other. Measured on a
/// native PowerPoint 16.112 export of the #841 Contoso deck, set in Posterama
/// Bold (usWin 2134/590 per 2048 upem, a 1.3301em line): slide 1 carries no
/// `<a:lnSpc>` and paces its three 50pt baselines exactly 60.00pt = 1.2em
/// apart, so the box is 1.2em there, and it seats the first 0.9411em below the
/// content top. The share predicts 0.9401em; halving the leading says 0.9770em,
/// 1.79pt low at that size, and every one of the deck's 18 titles was low by
/// 1.8-3.7pt (issue #1020).
///
/// A face that **fits** reads the same way once the whole-point seat of #1074
/// is accounted for, which is what hid it: a one-factor probe deck of
/// bottom-anchored boxes with every inset zeroed, exported natively at 14 sizes
/// from 8 to 100pt, separates the two on Georgia (usWin 1878/449, a 1.13623em
/// line). Its share is 0.968457em and the halved leading 0.948877em; the export
/// sits within its 0.12pt half-grid of the share at all 14 sizes and outside it
/// at 9 of them for the halved leading, by up to 2.04pt. The five that agree
/// are the ones where the two round to the same point, which is how a 17pt
/// Arial frame read as halved leading in #660 (issue #1118).
///
/// Word reads the hhea gap instead, and seats the baseline *below* it rather
/// than sharing a fixed box: see [`font_line_metrics_em`].
///
/// The below-baseline share is the descent gap a bottom-anchored box keeps
/// under its last baseline, which we used to drop entirely (issue #513).
///
/// `None` when no face declares an ascent, which no split could seat.
pub(crate) fn powerpoint_line_box_split_em<I>(faces: I) -> Option<(f64, f64)>
where
    I: IntoIterator<Item = (f64, f64)>,
{
    let mut above_share: f64 = 0.0;
    let mut below_share: f64 = 0.0;
    let mut seated: bool = false;
    for (ascent, descent) in faces {
        let Some((above, below)) = powerpoint_face_unit_line(ascent, descent) else {
            continue;
        };
        seated = true;
        above_share = above_share.max(above);
        below_share = below_share.max(below);
    }
    let natural: f64 = above_share + below_share;
    if !seated || natural <= 0.0 {
        return None;
    }
    let above: f64 = (POWERPOINT_LINE_HEIGHT_FACTOR * above_share / natural)
        .clamp(0.0, POWERPOINT_LINE_HEIGHT_FACTOR);
    Some((above, POWERPOINT_LINE_HEIGHT_FACTOR - above))
}

/// The `(ascent, descent)` pair PowerPoint measures a face's line box by, as
/// positive em fractions: OS/2's `usWin*` pair, falling back to hhea's for a
/// face carrying no `OS/2` table at all.
fn powerpoint_face_metrics_em(font: &typst::text::Font) -> (f64, f64) {
    let instance = measured_instance(font);
    let ttf = instance.ttf();
    let upem: f64 = f64::from(ttf.units_per_em()).max(1.0);
    if let Some(os2) = ttf.tables().os2 {
        let ascent: f64 = f64::from(os2.windows_ascender()).max(0.0);
        let descent: f64 = f64::from(os2.windows_descender());
        if ascent > 0.0 {
            return (ascent / upem, descent.abs() / upem);
        }
    }
    let hhea = ttf.tables().hhea;
    (
        f64::from(hhea.ascender).abs() / upem,
        f64::from(hhea.descender).abs() / upem,
    )
}

/// The `(ascent, descent)` pair the best face resolved for `family` measures
/// its line box by, as positive em fractions.
#[cfg(not(target_arch = "wasm32"))]
fn powerpoint_family_metrics_em(family: &str) -> Option<(f64, f64)> {
    use std::collections::HashMap;
    use std::sync::Mutex;
    type FaceMetricsEm = Option<(f64, f64)>;
    static CACHE: OnceLock<Mutex<HashMap<String, FaceMetricsEm>>> = OnceLock::new();

    if super::font_subst::active_font_search_paths().is_some() {
        return best_face(family)
            .or_else(|| best_face(crate::defaults::TYPST_DEFAULT_FONT_FAMILY))
            .map(|font| powerpoint_face_metrics_em(&font));
    }

    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let key: String = family.to_lowercase();
    if let Some(cached) = cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .get(&key)
    {
        return *cached;
    }

    // Resolve through the same substitute chain Typst sees in the emitted font
    // list. If none exists, Typst uses its embedded default, so the metrics
    // lookup must end there too.
    // Otherwise a missing Office face can collapse consecutive paragraphs to
    // the fallback font's glyph height (issue #705).
    let metrics: Option<(f64, f64)> = best_face(family)
        .or_else(|| best_face(crate::defaults::TYPST_DEFAULT_FONT_FAMILY))
        .map(|font| powerpoint_face_metrics_em(&font));
    cache
        .lock()
        .expect("metrics cache mutex should not be poisoned")
        .insert(key, metrics);
    metrics
}

#[cfg(target_arch = "wasm32")]
fn powerpoint_family_metrics_em(family: &str) -> Option<(f64, f64)> {
    best_face(family).map(|font| powerpoint_face_metrics_em(&font))
}

/// [`powerpoint_line_box_split_em`] for the best face resolved for `family`.
pub(crate) fn powerpoint_line_box_em(family: &str) -> Option<(f64, f64)> {
    powerpoint_line_box_em_for_families(std::slice::from_ref(&family))
}

/// [`powerpoint_line_box_split_em`] for every font on one line — the runs' and
/// the paragraph mark's — named in `families`.
///
/// A family that resolves to no face contributes nothing rather than voiding
/// the line: the box is still the one the faces that did resolve share.
pub(crate) fn powerpoint_line_box_em_for_families(families: &[&str]) -> Option<(f64, f64)> {
    powerpoint_line_box_split_em(
        families
            .iter()
            .filter_map(|family| powerpoint_family_metrics_em(family)),
    )
}
