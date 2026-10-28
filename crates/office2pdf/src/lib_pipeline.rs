use std::collections::HashSet;
#[cfg(not(target_arch = "wasm32"))]
use std::time::Instant;
#[cfg(target_arch = "wasm32")]
use web_time::Instant;

use crate::config::{ConvertOptions, Format};
use crate::error::{ConvertError, ConvertMetrics, ConvertResult, ConvertWarning};
use crate::parser::Parser;
use crate::{ir, parser, render};

fn format_label(format: Format) -> &'static str {
    match format {
        Format::Docx => "DOCX",
        Format::Pptx => "PPTX",
        Format::Xlsx => "XLSX",
    }
}

fn dedup_warnings(warnings: &mut Vec<ConvertWarning>) {
    let mut seen: HashSet<String> = HashSet::new();
    warnings.retain(|warning| seen.insert(warning.to_string()));
}

/// Build a `ConvertResult`, deduplicating warnings automatically so callers
/// don't need to remember to call `dedup_warnings` before every return site.
fn build_convert_result(
    pdf: Vec<u8>,
    mut warnings: Vec<ConvertWarning>,
    metrics: Option<ConvertMetrics>,
) -> ConvertResult {
    dedup_warnings(&mut warnings);
    ConvertResult {
        pdf,
        warnings,
        metrics,
    }
}

fn extract_panic_message(payload: &Box<dyn std::any::Any + Send>) -> String {
    if let Some(s) = payload.downcast_ref::<String>() {
        s.clone()
    } else if let Some(s) = payload.downcast_ref::<&str>() {
        (*s).to_string()
    } else {
        "unknown panic".to_string()
    }
}

const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

pub(super) fn is_ole2(data: &[u8]) -> bool {
    data.len() >= OLE2_MAGIC.len() && data[..OLE2_MAGIC.len()] == OLE2_MAGIC
}

#[cfg(not(target_arch = "wasm32"))]
pub(super) fn should_resolve_font_context(
    doc: &ir::Document,
    options: &ConvertOptions,
    has_embedded_fonts: bool,
) -> bool {
    has_embedded_fonts
        || !options.font_paths.is_empty()
        || !options.font_bytes.is_empty()
        || options
            .last_resort_font_family
            .as_deref()
            .is_some_and(|family| !family.trim().is_empty())
        || render::font_subst::document_requests_font_families(doc)
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_font_context_with_embedded(
    doc: &ir::Document,
    options: &ConvertOptions,
    embedded_font_dir: Option<&parser::embedded_fonts::EmbeddedFontDir>,
    in_memory_fonts: &[typst::text::Font],
) -> Option<render::font_context::FontSearchContext> {
    let has_embedded = embedded_font_dir.is_some_and(|d| !d.is_empty());
    if !should_resolve_font_context(doc, options, has_embedded) {
        return None;
    }
    let mut all_paths: Vec<std::path::PathBuf> = options.font_paths.clone();
    if let Some(dir) = embedded_font_dir
        && !dir.is_empty()
    {
        all_paths.push(dir.path().to_path_buf());
    }
    Some(
        render::font_context::resolve_font_search_context(&all_paths)
            .with_in_memory_fonts(in_memory_fonts)
            .with_last_resort_family(options.last_resort_font_family.as_deref()),
    )
}

fn load_registered_fonts(options: &ConvertOptions) -> Result<Vec<typst::text::Font>, ConvertError> {
    let mut fonts = Vec::new();
    for (index, bytes) in options.font_bytes.iter().enumerate() {
        let parsed = render::pdf::load_fonts_from_bytes([bytes.as_slice()]);
        if parsed.is_empty() {
            return Err(ConvertError::Render(format!(
                "registered font at index {index} contains no usable font faces"
            )));
        }
        fonts.extend(parsed);
    }
    Ok(fonts)
}

pub(super) fn load_additional_fonts(
    options: &ConvertOptions,
) -> Result<Vec<typst::text::Font>, ConvertError> {
    let fonts = load_registered_fonts(options)?;
    #[cfg(all(target_arch = "wasm32", feature = "wasm-cjk-font"))]
    {
        let mut fonts = fonts;
        fonts.extend_from_slice(crate::bundled_fonts::cjk_fonts());
        Ok(fonts)
    }
    #[cfg(not(all(target_arch = "wasm32", feature = "wasm-cjk-font")))]
    {
        Ok(fonts)
    }
}

/// Add document-scoped bundled faces without changing Typst's fallback book
/// for unrelated conversions. Caller-registered bytes stay ahead of bundled
/// data when they provide the same target family (issues #1458, #1463).
pub(super) fn extend_document_fonts(fonts: &mut Vec<typst::text::Font>, doc: &ir::Document) {
    if render::font_subst::document_requests_bundled_noto_serif(doc)
        && !fonts
            .iter()
            .any(|font| font.info().family == crate::bundled_fonts::NOTO_SERIF_FAMILY)
    {
        fonts.extend_from_slice(crate::bundled_fonts::noto_serif_fonts());
    }
    if render::font_subst::document_requests_bundled_noto_sans(doc)
        && !fonts
            .iter()
            .any(|font| font.info().family == crate::bundled_fonts::NOTO_SANS_FAMILY)
    {
        fonts.extend_from_slice(crate::bundled_fonts::noto_sans_fonts());
    }
    if render::font_subst::document_requests_bundled_selawik(doc)
        && !fonts
            .iter()
            .any(|font| font.info().family == crate::bundled_fonts::SELAWIK_FAMILY)
    {
        fonts.extend_from_slice(crate::bundled_fonts::selawik_fonts());
    }
}

fn effective_last_resort_family(options: &ConvertOptions) -> Option<&str> {
    let configured = options
        .last_resort_font_family
        .as_deref()
        .map(str::trim)
        .filter(|family| !family.is_empty());
    #[cfg(all(target_arch = "wasm32", feature = "wasm-cjk-font"))]
    return configured.or(Some(crate::bundled_fonts::CJK_LAST_RESORT_FAMILY));
    #[cfg(not(all(target_arch = "wasm32", feature = "wasm-cjk-font")))]
    configured
}

#[cfg(not(target_arch = "wasm32"))]
pub(super) fn convert(path: impl AsRef<std::path::Path>) -> Result<ConvertResult, ConvertError> {
    convert_with_options(path, &ConvertOptions::default())
}

#[cfg(not(target_arch = "wasm32"))]
pub(super) fn convert_with_options(
    path: impl AsRef<std::path::Path>,
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .ok_or_else(|| ConvertError::UnsupportedFormat("no file extension".to_string()))?;

    let format = Format::from_extension(ext)
        .ok_or_else(|| ConvertError::UnsupportedFormat(ext.to_string()))?;

    let data = std::fs::read(path)?;
    convert_bytes(&data, format, options)
}

pub(super) fn convert_bytes(
    data: &[u8],
    format: Format,
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    if is_ole2(data) {
        return Err(ConvertError::UnsupportedEncryption);
    }

    #[cfg(feature = "pdf-ops")]
    if options.streaming && format == Format::Xlsx {
        return convert_bytes_streaming_xlsx(data, options);
    }

    let total_start: Instant = Instant::now();
    let input_size_bytes = data.len() as u64;
    let mut additional_fonts = load_additional_fonts(options)?;

    // Extract embedded fonts before parsing (PPTX/DOCX only). Native keeps the
    // materialized directory alive through compilation; WASM keeps parsed
    // faces in memory for its filesystem-free Typst world.
    #[cfg(not(target_arch = "wasm32"))]
    let embedded_font_dir = parser::embedded_fonts::extract_embedded_fonts(data, format);
    #[cfg(target_arch = "wasm32")]
    let embedded_font_data = parser::embedded_fonts::extract_embedded_font_data(data, format);
    #[cfg(target_arch = "wasm32")]
    let embedded_fonts = embedded_font_data.as_ref().map_or_else(Vec::new, |fonts| {
        render::pdf::load_fonts_from_bytes(fonts.font_bytes())
    });
    let parser: Box<dyn Parser> = match format {
        Format::Docx => Box::new(parser::docx::DocxParser),
        Format::Pptx => Box::new(parser::pptx::PptxParser),
        Format::Xlsx => Box::new(parser::xlsx::XlsxParser),
    };

    let parse_start: Instant = Instant::now();
    let parse_result =
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| parser.parse(data, options)));
    let (doc, mut warnings) = match parse_result {
        Ok(result) => result?,
        Err(panic_info) => {
            return Err(ConvertError::Parse(format!(
                "upstream parser panicked: {}",
                extract_panic_message(&panic_info)
            )));
        }
    };
    let parse_duration = parse_start.elapsed();
    extend_document_fonts(&mut additional_fonts, &doc);

    #[cfg(target_arch = "wasm32")]
    let in_memory_fonts: Vec<typst::text::Font> = embedded_fonts
        .into_iter()
        .chain(additional_fonts.iter().cloned())
        .collect();

    #[cfg(not(target_arch = "wasm32"))]
    let font_context = resolve_font_context_with_embedded(
        &doc,
        options,
        embedded_font_dir.as_ref(),
        &additional_fonts,
    );
    #[cfg(target_arch = "wasm32")]
    let font_context = (!in_memory_fonts.is_empty()
        || effective_last_resort_family(options).is_some()
        || render::font_subst::document_requests_font_families(&doc))
    .then(|| {
        render::font_context::resolve_font_search_context_from_fonts(&in_memory_fonts)
            .with_last_resort_family(effective_last_resort_family(options))
    });

    #[cfg(not(target_arch = "wasm32"))]
    if let Some(font_context) = font_context.as_ref() {
        warnings.extend(
            render::font_subst::detect_missing_font_fallbacks_with_context(&doc, font_context)
                .into_iter()
                .map(|(from, to)| ConvertWarning::FallbackUsed {
                    format: format_label(format).to_string(),
                    from,
                    to,
                }),
        );
    }

    #[cfg(target_arch = "wasm32")]
    if let Some(font_context) = font_context.as_ref() {
        warnings.extend(
            render::font_subst::detect_missing_font_fallbacks_with_context(&doc, font_context)
                .into_iter()
                .map(|(from, to)| ConvertWarning::FallbackUsed {
                    format: format_label(format).to_string(),
                    from,
                    to,
                }),
        );
    } else {
        warnings.extend(
            render::font_subst::detect_missing_font_fallbacks(&doc, &options.font_paths)
                .into_iter()
                .map(|(from, to)| ConvertWarning::FallbackUsed {
                    format: format_label(format).to_string(),
                    from,
                    to,
                }),
        );
    }

    let codegen_start: Instant = Instant::now();
    #[cfg(not(target_arch = "wasm32"))]
    let output = render::typst_gen::generate_typst_with_options_and_font_context(
        &doc,
        options,
        font_context.as_ref(),
    )?;
    #[cfg(target_arch = "wasm32")]
    let output = match font_context.as_ref() {
        Some(font_context) => render::typst_gen::generate_typst_with_options_and_font_context(
            &doc,
            options,
            Some(font_context),
        )?,
        None => render::typst_gen::generate_typst_with_options(&doc, options)?,
    };
    let codegen_duration = codegen_start.elapsed();

    let compile_start: Instant = Instant::now();
    #[cfg(not(target_arch = "wasm32"))]
    let (pdf, page_count) = render::pdf::compile_to_pdf_with_fonts_counted(
        &output.source,
        &output.images,
        options.pdf_standard,
        font_context
            .as_ref()
            .map(|context| context.search_paths())
            .unwrap_or(&[]),
        &additional_fonts,
        options.tagged,
        options.pdf_ua,
    )?;
    #[cfg(target_arch = "wasm32")]
    let (pdf, page_count) = render::pdf::compile_to_pdf_with_fonts_counted(
        &output.source,
        &output.images,
        options.pdf_standard,
        &[],
        &in_memory_fonts,
        options.tagged,
        options.pdf_ua,
    )?;
    let compile_duration = compile_start.elapsed();

    let total_duration = total_start.elapsed();
    let output_size_bytes = pdf.len() as u64;

    Ok(build_convert_result(
        pdf,
        warnings,
        Some(ConvertMetrics {
            parse_duration,
            codegen_duration,
            compile_duration,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count,
        }),
    ))
}

#[cfg(feature = "pdf-ops")]
struct PlannedStreamingChunk {
    document: ir::Document,
    page_count: u32,
}

#[cfg(feature = "pdf-ops")]
struct StreamingProbeContext<'a> {
    options: &'a ConvertOptions,
    font_context: Option<&'a render::font_context::FontSearchContext>,
    additional_fonts: &'a [typst::text::Font],
    codegen_duration: std::time::Duration,
    compile_duration: std::time::Duration,
}

#[cfg(feature = "pdf-ops")]
impl StreamingProbeContext<'_> {
    fn page_count(&mut self, document: &ir::Document) -> Result<u32, ConvertError> {
        let codegen_start: Instant = Instant::now();
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            document,
            self.options,
            self.font_context,
        )?;
        self.codegen_duration += codegen_start.elapsed();

        let compile_start: Instant = Instant::now();
        let page_count = render::pdf::compile_page_count_with_fonts(
            &output.source,
            &output.images,
            self.font_context
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            self.additional_fonts,
        )?;
        self.compile_duration += compile_start.elapsed();
        Ok(page_count)
    }
}

#[cfg(feature = "pdf-ops")]
fn sheet_row_sections(table: &ir::Table) -> (usize, usize, usize, usize) {
    let heading_rows: usize = usize::from(table.prints_headings && !table.rows.is_empty());
    let remaining_rows: usize = table.rows.len().saturating_sub(heading_rows);
    let lead_rows: usize = table.non_repeating_header_row_count.min(remaining_rows);
    let declared_title_rows: usize = table
        .header_row_count
        .min(remaining_rows.saturating_sub(lead_rows));
    let title_start: usize = heading_rows + lead_rows;
    let title_rows: usize = render::typst_gen::header_row_count_covering_rowspans(
        &table.rows[title_start..],
        declared_title_rows,
    );
    let body_start: usize = heading_rows + lead_rows + title_rows;
    (heading_rows, lead_rows, title_rows, body_start)
}

#[cfg(feature = "pdf-ops")]
fn fixed_row_page_aligned_end(
    page: &ir::SheetPage,
    body_start: usize,
    chunk_start: usize,
    requested_end: usize,
    options: &ConvertOptions,
) -> Option<usize> {
    if page.charts.iter().any(|chart| chart.placement.is_none())
        || page
            .table
            .rows
            .iter()
            .any(|row| row.height.is_none() || row.cells.iter().any(|cell| cell.row_span > 1))
    {
        return None;
    }

    let (heading_rows, lead_rows, title_rows, _) = sheet_row_sections(&page.table);
    let repeat_height: f64 = page.table.rows[..heading_rows]
        .iter()
        .chain(
            page.table.rows[heading_rows + lead_rows..heading_rows + lead_rows + title_rows].iter(),
        )
        .filter_map(|row| row.height)
        .sum();
    let lead_height: f64 = page.table.rows[heading_rows..heading_rows + lead_rows]
        .iter()
        .filter_map(|row| row.height)
        .sum();
    let size = render::typst_gen::resolve_page_size(&page.size, options);
    let page_height: f64 = size.height - page.margins.top - page.margins.bottom;
    let regular_capacity: f64 = page_height - repeat_height;
    if !regular_capacity.is_finite() || regular_capacity <= 0.0 {
        return None;
    }

    let first_page_lead: f64 = if chunk_start == 0 { lead_height } else { 0.0 };
    let mut remaining: f64 = regular_capacity - first_page_lead;
    if remaining < 0.0 {
        return None;
    }

    let body_rows: &[ir::TableRow] = &page.table.rows[body_start..];
    for (row_index, row) in body_rows.iter().enumerate().skip(chunk_start) {
        let row_height: f64 = row.height?;
        if !row_height.is_finite() || row_height < 0.0 {
            return None;
        }
        if row_height > remaining + f64::EPSILON {
            if row_index >= requested_end {
                return Some(row_index);
            }
            remaining = regular_capacity;
            if row_height > remaining + f64::EPSILON {
                return None;
            }
        }
        remaining = (remaining - row_height).max(0.0);
    }
    Some(body_rows.len())
}

#[cfg(feature = "pdf-ops")]
fn streaming_sheet_chunk(
    page: &ir::SheetPage,
    body_range: std::ops::Range<usize>,
    is_first_chunk: bool,
) -> ir::SheetPage {
    let (heading_rows, lead_rows, title_rows, body_start) = sheet_row_sections(&page.table);
    let mut chunk: ir::SheetPage = page.clone();
    let mut rows: Vec<ir::TableRow> = Vec::with_capacity(
        heading_rows + title_rows + body_range.len() + usize::from(is_first_chunk) * lead_rows,
    );
    if is_first_chunk {
        rows.extend_from_slice(&page.table.rows[..body_start]);
    } else {
        rows.extend_from_slice(&page.table.rows[..heading_rows]);
        let title_start: usize = heading_rows + lead_rows;
        rows.extend_from_slice(&page.table.rows[title_start..title_start + title_rows]);
    }
    rows.extend_from_slice(
        &page.table.rows[body_start + body_range.start..body_start + body_range.end],
    );
    chunk.table.rows = rows;
    chunk.table.header_row_count = title_rows;
    chunk.table.non_repeating_header_row_count = if is_first_chunk { lead_rows } else { 0 };
    if !is_first_chunk {
        chunk.charts.clear();
        chunk.images.clear();
        chunk.text_boxes.clear();
        chunk.shapes.clear();
    }
    chunk
}

#[cfg(feature = "pdf-ops")]
fn streaming_chunk_document(source: &ir::Document, page: ir::SheetPage) -> ir::Document {
    ir::Document {
        metadata: source.metadata.clone(),
        pages: vec![ir::Page::Sheet(page)],
        styles: source.styles.clone(),
    }
}

#[cfg(feature = "pdf-ops")]
fn plan_streaming_sheet_chunks(
    source: &ir::Document,
    page: &ir::SheetPage,
    chunk_size: usize,
    probe_context: &mut StreamingProbeContext<'_>,
) -> Result<Vec<PlannedStreamingChunk>, ConvertError> {
    let (_, _, _, body_start) = sheet_row_sections(&page.table);
    let body_row_count: usize = page.table.rows.len().saturating_sub(body_start);
    if body_row_count == 0 {
        let document = streaming_chunk_document(source, page.clone());
        let page_count = probe_context.page_count(&document)?;
        return Ok(vec![PlannedStreamingChunk {
            document,
            page_count,
        }]);
    }

    let chunk_size: usize = chunk_size.max(1);
    let mut chunks: Vec<PlannedStreamingChunk> = Vec::new();
    let mut chunk_start: usize = 0;
    while chunk_start < body_row_count {
        let requested_end: usize = (chunk_start + chunk_size).min(body_row_count);
        let is_first_chunk: bool = chunk_start == 0;
        if let Some(aligned_end) = fixed_row_page_aligned_end(
            page,
            body_start,
            chunk_start,
            requested_end,
            probe_context.options,
        ) {
            let aligned_document = streaming_chunk_document(
                source,
                streaming_sheet_chunk(page, chunk_start..aligned_end, is_first_chunk),
            );
            let aligned_pages = probe_context.page_count(&aligned_document)?;
            let boundary_is_confirmed = if aligned_end == body_row_count {
                true
            } else {
                let next_document = streaming_chunk_document(
                    source,
                    streaming_sheet_chunk(page, chunk_start..aligned_end + 1, is_first_chunk),
                );
                probe_context.page_count(&next_document)? > aligned_pages
            };
            if boundary_is_confirmed {
                tracing::debug!(
                    sheet = page.name,
                    requested_rows = chunk_size,
                    planned_rows = aligned_end - chunk_start,
                    planned_pages = aligned_pages,
                    "aligned fixed-row streaming chunk to a PDF page boundary"
                );
                chunks.push(PlannedStreamingChunk {
                    document: aligned_document,
                    page_count: aligned_pages,
                });
                chunk_start = aligned_end;
                continue;
            }
        }

        let requested_document = streaming_chunk_document(
            source,
            streaming_sheet_chunk(page, chunk_start..requested_end, is_first_chunk),
        );
        let requested_pages = probe_context.page_count(&requested_document)?;

        if requested_end == body_row_count {
            chunks.push(PlannedStreamingChunk {
                document: requested_document,
                page_count: requested_pages,
            });
            break;
        }

        let mut last_same_end: usize = requested_end;
        let mut first_larger_end: Option<usize> = None;
        let mut step: usize = 1;
        loop {
            let probe_end: usize = requested_end.saturating_add(step).min(body_row_count);
            let probe_document = streaming_chunk_document(
                source,
                streaming_sheet_chunk(page, chunk_start..probe_end, is_first_chunk),
            );
            let probe_pages = probe_context.page_count(&probe_document)?;
            if probe_pages > requested_pages {
                first_larger_end = Some(probe_end);
                break;
            }
            last_same_end = probe_end;
            if probe_end == body_row_count {
                chunks.push(PlannedStreamingChunk {
                    document: probe_document,
                    page_count: probe_pages,
                });
                chunk_start = body_row_count;
                break;
            }
            step = step.saturating_mul(2).max(1);
        }
        if chunk_start == body_row_count {
            break;
        }

        let mut lower_end: usize = last_same_end + 1;
        let mut upper_end: usize =
            first_larger_end.expect("the page-growth search must find an upper bound");
        while lower_end < upper_end {
            let middle_end: usize = lower_end + (upper_end - lower_end) / 2;
            let middle_document = streaming_chunk_document(
                source,
                streaming_sheet_chunk(page, chunk_start..middle_end, is_first_chunk),
            );
            let middle_pages = probe_context.page_count(&middle_document)?;
            if middle_pages > requested_pages {
                upper_end = middle_end;
            } else {
                lower_end = middle_end + 1;
            }
        }

        let aligned_end: usize = lower_end - 1;
        let aligned_document = if aligned_end == requested_end {
            requested_document
        } else {
            streaming_chunk_document(
                source,
                streaming_sheet_chunk(page, chunk_start..aligned_end, is_first_chunk),
            )
        };
        tracing::debug!(
            sheet = page.name,
            requested_rows = chunk_size,
            planned_rows = aligned_end - chunk_start,
            planned_pages = requested_pages,
            "aligned streaming chunk to a PDF page boundary"
        );
        chunks.push(PlannedStreamingChunk {
            document: aligned_document,
            page_count: requested_pages,
        });
        chunk_start = aligned_end;
    }
    Ok(chunks)
}

#[cfg(feature = "pdf-ops")]
fn convert_bytes_streaming_xlsx(
    data: &[u8],
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    let total_start: Instant = Instant::now();
    let input_size_bytes = data.len() as u64;
    let mut additional_fonts = load_additional_fonts(options)?;
    let chunk_size = options
        .streaming_chunk_size
        .unwrap_or(crate::defaults::DEFAULT_STREAMING_CHUNK_SIZE)
        .max(1);

    let xlsx_parser = parser::xlsx::XlsxParser;
    let parse_start: Instant = Instant::now();
    let parse_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        xlsx_parser.parse(data, options)
    }));
    let (document, mut warnings) = match parse_result {
        Ok(result) => result?,
        Err(panic_info) => {
            return Err(ConvertError::Parse(format!(
                "upstream parser panicked: {}",
                extract_panic_message(&panic_info)
            )));
        }
    };
    let parse_duration = parse_start.elapsed();
    extend_document_fonts(&mut additional_fonts, &document);

    let needs_in_memory_font_context = !additional_fonts.is_empty()
        || effective_last_resort_family(options).is_some()
        || render::font_subst::document_requests_font_families(&document);
    #[cfg(not(target_arch = "wasm32"))]
    let font_context =
        (needs_in_memory_font_context || !options.font_paths.is_empty()).then(|| {
            render::font_context::resolve_font_search_context(&options.font_paths)
                .with_in_memory_fonts(&additional_fonts)
                .with_last_resort_family(effective_last_resort_family(options))
        });
    #[cfg(target_arch = "wasm32")]
    let font_context = needs_in_memory_font_context.then(|| {
        render::font_context::resolve_font_search_context_from_fonts(&additional_fonts)
            .with_last_resort_family(effective_last_resort_family(options))
    });

    let mut probe_context = StreamingProbeContext {
        options,
        font_context: font_context.as_ref(),
        additional_fonts: &additional_fonts,
        codegen_duration: std::time::Duration::ZERO,
        compile_duration: std::time::Duration::ZERO,
    };
    let mut planned_chunks: Vec<PlannedStreamingChunk> = Vec::new();
    for page in &document.pages {
        let ir::Page::Sheet(sheet_page) = page else {
            return Err(ConvertError::Render(
                "XLSX streaming received a non-sheet IR page".to_string(),
            ));
        };
        planned_chunks.extend(plan_streaming_sheet_chunks(
            &document,
            sheet_page,
            chunk_size,
            &mut probe_context,
        )?);
    }
    let mut codegen_duration_total = probe_context.codegen_duration;
    let mut compile_duration_total = probe_context.compile_duration;

    if planned_chunks.is_empty() {
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            &document,
            options,
            font_context.as_ref(),
        )?;
        let (pdf, page_count) = render::pdf::compile_to_pdf_with_fonts_counted(
            &output.source,
            &output.images,
            options.pdf_standard,
            font_context
                .as_ref()
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            &additional_fonts,
            options.tagged,
            options.pdf_ua,
        )?;
        let total_duration = total_start.elapsed();
        let output_size_bytes = pdf.len() as u64;
        return Ok(build_convert_result(
            pdf,
            warnings,
            Some(ConvertMetrics {
                parse_duration,
                codegen_duration: codegen_duration_total,
                compile_duration: compile_duration_total,
                total_duration,
                input_size_bytes,
                output_size_bytes,
                page_count,
            }),
        ));
    }

    let mut all_pdfs: Vec<Vec<u8>> = Vec::with_capacity(planned_chunks.len());
    let mut total_page_count: u32 = 0;
    for planned_chunk in planned_chunks {
        if let Some(font_context) = font_context.as_ref() {
            warnings.extend(
                render::font_subst::detect_missing_font_fallbacks_with_context(
                    &planned_chunk.document,
                    font_context,
                )
                .into_iter()
                .map(|(from, to)| ConvertWarning::FallbackUsed {
                    format: format_label(Format::Xlsx).to_string(),
                    from,
                    to,
                }),
            );
        }

        let codegen_start: Instant = Instant::now();
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            &planned_chunk.document,
            options,
            font_context.as_ref(),
        )?;
        codegen_duration_total += codegen_start.elapsed();

        let compile_start: Instant = Instant::now();
        let (pdf, chunk_pages) = render::pdf::compile_to_pdf_with_fonts_counted(
            &output.source,
            &output.images,
            options.pdf_standard,
            font_context
                .as_ref()
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            &additional_fonts,
            options.tagged,
            options.pdf_ua,
        )?;
        compile_duration_total += compile_start.elapsed();
        if chunk_pages != planned_chunk.page_count {
            return Err(ConvertError::Render(format!(
                "streaming pagination changed between planning and export: planned {} pages, exported {chunk_pages}",
                planned_chunk.page_count
            )));
        }
        total_page_count += chunk_pages;
        all_pdfs.push(pdf);
    }

    let final_pdf = if all_pdfs.len() == 1 {
        all_pdfs
            .into_iter()
            .next()
            .expect("all_pdfs is non-empty (len == 1)")
    } else {
        let refs: Vec<&[u8]> = all_pdfs.iter().map(|pdf| pdf.as_slice()).collect();
        crate::pdf_ops::merge(&refs)
            .map_err(|error| ConvertError::Render(format!("PDF merge failed: {error}")))?
    };

    let total_duration = total_start.elapsed();
    let output_size_bytes = final_pdf.len() as u64;
    Ok(build_convert_result(
        final_pdf,
        warnings,
        Some(ConvertMetrics {
            parse_duration,
            codegen_duration: codegen_duration_total,
            compile_duration: compile_duration_total,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count: total_page_count,
        }),
    ))
}

pub(super) fn render_document(doc: &ir::Document) -> Result<Vec<u8>, ConvertError> {
    #[cfg(not(target_arch = "wasm32"))]
    {
        let options = ConvertOptions::default();
        let mut fonts = load_additional_fonts(&options)?;
        extend_document_fonts(&mut fonts, doc);
        let font_context = resolve_font_context_with_embedded(doc, &options, None, &fonts);
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            doc,
            &options,
            font_context.as_ref(),
        )?;
        let search_paths = font_context
            .as_ref()
            .map(|context| context.search_paths())
            .unwrap_or(&[]);
        if fonts.is_empty() {
            render::pdf::compile_to_pdf(
                &output.source,
                &output.images,
                None,
                search_paths,
                false,
                false,
            )
        } else {
            render::pdf::compile_to_pdf_with_fonts(
                &output.source,
                &output.images,
                None,
                search_paths,
                &fonts,
                false,
                false,
            )
        }
    }
    #[cfg(all(target_arch = "wasm32", not(feature = "wasm-cjk-font")))]
    {
        let options = ConvertOptions::default();
        let mut fonts = load_additional_fonts(&options)?;
        extend_document_fonts(&mut fonts, doc);
        let font_context = (!fonts.is_empty()
            || render::font_subst::document_requests_font_families(doc))
        .then(|| render::font_context::resolve_font_search_context_from_fonts(&fonts));
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            doc,
            &options,
            font_context.as_ref(),
        )?;
        if fonts.is_empty() {
            // Preserve the no-font fast path for unrelated filesystem-free
            // documents; only poster requests need the bundled-face compiler.
            render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
        } else {
            render::pdf::compile_to_pdf_with_fonts(
                &output.source,
                &output.images,
                None,
                &[],
                &fonts,
                false,
                false,
            )
        }
    }
    #[cfg(all(target_arch = "wasm32", feature = "wasm-cjk-font"))]
    {
        let options = ConvertOptions::default();
        let mut fonts = load_additional_fonts(&options)?;
        extend_document_fonts(&mut fonts, doc);
        let font_context = render::font_context::resolve_font_search_context_from_fonts(&fonts)
            .with_last_resort_family(effective_last_resort_family(&options));
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            doc,
            &options,
            Some(&font_context),
        )?;
        render::pdf::compile_to_pdf_with_fonts(
            &output.source,
            &output.images,
            None,
            &[],
            &fonts,
            false,
            false,
        )
    }
}

#[cfg(all(test, feature = "pdf-ops"))]
mod streaming_chunk_tests {
    use super::*;

    fn fixed_row(height: f64) -> ir::TableRow {
        ir::TableRow {
            cant_split: false,
            cells: Vec::new(),
            height: Some(height),
            minimum_height: None,
        }
    }

    fn sheet_page(table: ir::Table) -> ir::SheetPage {
        ir::SheetPage {
            name: "Sheet1".to_string(),
            size: ir::PageSize {
                width: 100.0,
                height: 100.0,
            },
            margins: ir::Margins {
                top: 10.0,
                right: 0.0,
                bottom: 10.0,
                left: 0.0,
            },
            table,
            header: None,
            footer: None,
            charts: Vec::new(),
            images: Vec::new(),
            text_boxes: Vec::new(),
            shapes: Vec::new(),
        }
    }

    #[test]
    fn fixed_rows_extend_each_chunk_to_the_current_page_boundary() {
        let page = sheet_page(ir::Table {
            rows: (0..5).map(|_| fixed_row(30.0)).collect(),
            ..Default::default()
        });
        let options = ConvertOptions::default();

        assert_eq!(
            fixed_row_page_aligned_end(&page, 0, 0, 1, &options),
            Some(2)
        );
        assert_eq!(
            fixed_row_page_aligned_end(&page, 0, 2, 3, &options),
            Some(4)
        );
        assert_eq!(
            fixed_row_page_aligned_end(&page, 0, 4, 5, &options),
            Some(5)
        );
    }

    #[test]
    fn fixed_rows_reserve_repeating_and_first_page_only_headers() {
        let page = sheet_page(ir::Table {
            rows: std::iter::once(fixed_row(20.0))
                .chain(std::iter::once(fixed_row(10.0)))
                .chain((0..5).map(|_| fixed_row(25.0)))
                .collect(),
            header_row_count: 1,
            non_repeating_header_row_count: 1,
            ..Default::default()
        });
        let options = ConvertOptions::default();

        assert_eq!(
            fixed_row_page_aligned_end(&page, 2, 0, 1, &options),
            Some(2)
        );
        assert_eq!(
            fixed_row_page_aligned_end(&page, 2, 2, 3, &options),
            Some(4)
        );
    }

    #[test]
    fn content_driven_rows_use_typst_boundary_search() {
        let page = sheet_page(ir::Table {
            rows: vec![ir::TableRow {
                cant_split: false,
                cells: Vec::new(),
                height: None,
                minimum_height: None,
            }],
            ..Default::default()
        });

        assert_eq!(
            fixed_row_page_aligned_end(&page, 0, 0, 1, &ConvertOptions::default()),
            None
        );
    }
}
