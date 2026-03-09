use std::collections::HashSet;
use std::time::Instant;

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
pub(super) fn should_resolve_font_context(doc: &ir::Document, options: &ConvertOptions) -> bool {
    !options.font_paths.is_empty() || render::font_subst::document_requests_font_families(doc)
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_font_context_if_needed(
    doc: &ir::Document,
    options: &ConvertOptions,
) -> Option<render::font_context::FontSearchContext> {
    should_resolve_font_context(doc, options)
        .then(|| render::font_context::resolve_font_search_context(&options.font_paths))
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
    let page_count = doc.pages.len() as u32;

    #[cfg(not(target_arch = "wasm32"))]
    let font_context = resolve_font_context_if_needed(&doc, options);

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
    warnings.extend(
        render::font_subst::detect_missing_font_fallbacks(&doc, &options.font_paths)
            .into_iter()
            .map(|(from, to)| ConvertWarning::FallbackUsed {
                format: format_label(format).to_string(),
                from,
                to,
            }),
    );
    dedup_warnings(&mut warnings);

    let codegen_start: Instant = Instant::now();
    #[cfg(not(target_arch = "wasm32"))]
    let output = render::typst_gen::generate_typst_with_options_and_font_context(
        &doc,
        options,
        font_context.as_ref(),
    )?;
    #[cfg(target_arch = "wasm32")]
    let output = render::typst_gen::generate_typst_with_options(&doc, options)?;
    let codegen_duration = codegen_start.elapsed();

    let compile_start: Instant = Instant::now();
    #[cfg(not(target_arch = "wasm32"))]
    let pdf = render::pdf::compile_to_pdf(
        &output.source,
        &output.images,
        options.pdf_standard,
        font_context
            .as_ref()
            .map(|context| context.search_paths())
            .unwrap_or(&[]),
        options.tagged,
        options.pdf_ua,
    )?;
    #[cfg(target_arch = "wasm32")]
    let pdf = render::pdf::compile_to_pdf(
        &output.source,
        &output.images,
        options.pdf_standard,
        &options.font_paths,
        options.tagged,
        options.pdf_ua,
    )?;
    let compile_duration = compile_start.elapsed();

    let total_duration = total_start.elapsed();
    let output_size_bytes = pdf.len() as u64;

    Ok(ConvertResult {
        pdf,
        warnings,
        metrics: Some(ConvertMetrics {
            parse_duration,
            codegen_duration,
            compile_duration,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count,
        }),
    })
}

#[cfg(feature = "pdf-ops")]
fn convert_bytes_streaming_xlsx(
    data: &[u8],
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    let total_start: Instant = Instant::now();
    let input_size_bytes = data.len() as u64;
    let chunk_size = options.streaming_chunk_size.unwrap_or(1000);

    let xlsx_parser = parser::xlsx::XlsxParser;

    let parse_start: Instant = Instant::now();
    let parse_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        xlsx_parser.parse_streaming(data, options, chunk_size)
    }));
    let (chunk_docs, mut warnings) = match parse_result {
        Ok(result) => result?,
        Err(panic_info) => {
            return Err(ConvertError::Parse(format!(
                "upstream parser panicked: {}",
                extract_panic_message(&panic_info)
            )));
        }
    };
    let parse_duration = parse_start.elapsed();

    if chunk_docs.is_empty() {
        let empty_doc = ir::Document {
            metadata: ir::Metadata::default(),
            pages: vec![],
            styles: ir::StyleSheet::default(),
        };
        #[cfg(not(target_arch = "wasm32"))]
        let font_context = resolve_font_context_if_needed(&empty_doc, options);
        #[cfg(not(target_arch = "wasm32"))]
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            &empty_doc,
            &ConvertOptions::default(),
            font_context.as_ref(),
        )?;
        #[cfg(target_arch = "wasm32")]
        let output = render::typst_gen::generate_typst(&empty_doc)?;
        #[cfg(not(target_arch = "wasm32"))]
        let pdf = render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            None,
            font_context
                .as_ref()
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            false,
            false,
        )?;
        #[cfg(target_arch = "wasm32")]
        let pdf =
            render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)?;
        let total_duration = total_start.elapsed();
        dedup_warnings(&mut warnings);
        return Ok(ConvertResult {
            pdf,
            warnings,
            metrics: Some(ConvertMetrics {
                parse_duration,
                codegen_duration: std::time::Duration::ZERO,
                compile_duration: std::time::Duration::ZERO,
                total_duration,
                input_size_bytes,
                output_size_bytes: 0,
                page_count: 0,
            }),
        });
    }

    let mut all_pdfs: Vec<Vec<u8>> = Vec::with_capacity(chunk_docs.len());
    let mut codegen_duration_total = std::time::Duration::ZERO;
    let mut compile_duration_total = std::time::Duration::ZERO;
    let mut total_page_count: u32 = 0;

    #[cfg(not(target_arch = "wasm32"))]
    let font_context = if options.font_paths.is_empty()
        && !chunk_docs
            .iter()
            .any(render::font_subst::document_requests_font_families)
    {
        None
    } else {
        Some(render::font_context::resolve_font_search_context(
            &options.font_paths,
        ))
    };

    for chunk_doc in chunk_docs {
        total_page_count += chunk_doc.pages.len() as u32;

        let codegen_start: Instant = Instant::now();
        #[cfg(not(target_arch = "wasm32"))]
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            &chunk_doc,
            options,
            font_context.as_ref(),
        )?;
        #[cfg(target_arch = "wasm32")]
        let output = render::typst_gen::generate_typst_with_options(&chunk_doc, options)?;
        codegen_duration_total += codegen_start.elapsed();

        let compile_start: Instant = Instant::now();
        #[cfg(not(target_arch = "wasm32"))]
        let pdf = render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            options.pdf_standard,
            font_context
                .as_ref()
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            options.tagged,
            options.pdf_ua,
        )?;
        #[cfg(target_arch = "wasm32")]
        let pdf = render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            options.pdf_standard,
            &options.font_paths,
            options.tagged,
            options.pdf_ua,
        )?;
        compile_duration_total += compile_start.elapsed();

        all_pdfs.push(pdf);
    }

    let final_pdf = if all_pdfs.len() == 1 {
        all_pdfs.into_iter().next().unwrap()
    } else {
        let refs: Vec<&[u8]> = all_pdfs.iter().map(|p| p.as_slice()).collect();
        crate::pdf_ops::merge(&refs)
            .map_err(|e| ConvertError::Render(format!("PDF merge failed: {e}")))?
    };

    let total_duration = total_start.elapsed();
    let output_size_bytes = final_pdf.len() as u64;
    dedup_warnings(&mut warnings);

    Ok(ConvertResult {
        pdf: final_pdf,
        warnings,
        metrics: Some(ConvertMetrics {
            parse_duration,
            codegen_duration: codegen_duration_total,
            compile_duration: compile_duration_total,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count: total_page_count,
        }),
    })
}

pub(super) fn render_document(doc: &ir::Document) -> Result<Vec<u8>, ConvertError> {
    #[cfg(not(target_arch = "wasm32"))]
    {
        let options = ConvertOptions::default();
        let font_context = resolve_font_context_if_needed(doc, &options);
        let output = render::typst_gen::generate_typst_with_options_and_font_context(
            doc,
            &options,
            font_context.as_ref(),
        )?;
        render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            None,
            font_context
                .as_ref()
                .map(|context| context.search_paths())
                .unwrap_or(&[]),
            false,
            false,
        )
    }
    #[cfg(target_arch = "wasm32")]
    {
        let output = render::typst_gen::generate_typst(doc)?;
        render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
    }
}
