use std::collections::HashMap;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};

/// Maximum nesting depth for tables-within-tables.  Deeper nesting is silently
/// truncated to prevent stack overflow on pathological documents.
const MAX_TABLE_DEPTH: usize = 64;
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Color, Document,
    FloatingImage, FloatingTextBox, ImageData, ImageFormat, Insets, LineSpacing, Page, Paragraph,
    ParagraphStyle, Run, StyleSheet, TabAlignment, TabLeader, TabStop, Table, TableCell, TableRow,
    TextDirection, TextStyle, VerticalTextAlign,
};
use crate::parser::Parser;

#[cfg(test)]
use self::contexts::scan_table_headers;
use self::contexts::{
    BidiContext, ChartContext, DocxConversionContext, DrawingTextBoxContext, DrawingTextBoxInfo,
    MathContext, NoteContext, SmallCapsContext, TableHeaderContext, VmlTextBoxContext,
    VmlTextBoxInfo, WrapContext, build_chart_context_from_xml, build_math_context_from_xml,
    build_note_context_from_xml, build_wrap_context_from_xml,
    extract_column_layout_from_section_property, is_note_reference_run, read_zip_text,
    scan_column_layouts,
};
use self::lists::{
    NumberingMap, TaggedElement, build_numbering_map, extract_num_info, group_into_lists,
};
use self::media::{
    extract_drawing_image, extract_drawing_text_box_blocks, extract_shape_image,
    extract_vml_shape_text_box,
};
#[cfg(test)]
use self::sections::extract_page_size;
use self::sections::{
    HeaderFooterAssets, build_flow_page_from_section, build_header_footer_assets,
};
use self::styles::{
    DOC_DEFAULT_STYLE_ID, ResolvedStyle, StyleMap, TabStopOverride, apply_tab_stop_overrides,
    build_style_map, get_paragraph_style_id, merge_paragraph_style, merge_text_style,
};
use self::tables::convert_table;
use self::text::{
    extract_doc_default_text_style, extract_paragraph_style, extract_run_style, extract_run_text,
    extract_run_text_skip_column_breaks, extract_tab_stop_overrides, is_column_break,
    parse_hex_color, resolve_hyperlink_url,
};
#[cfg(test)]
use self::text::{extract_tab_stops, resolve_highlight_color};

#[path = "docx_contexts.rs"]
mod contexts;
#[path = "docx_lists.rs"]
mod lists;
#[path = "docx_media.rs"]
mod media;
#[path = "docx_sections.rs"]
mod sections;
#[path = "docx_styles.rs"]
mod styles;
#[path = "docx_tables.rs"]
mod tables;
#[path = "docx_text.rs"]
mod text;

/// Parser for DOCX (Office Open XML Word) documents.
pub struct DocxParser;

/// Map from relationship ID → PNG image bytes.
type ImageMap = HashMap<String, Vec<u8>>;

/// Map from relationship ID → hyperlink URL.
type HyperlinkMap = HashMap<String, String>;

/// Build a lookup map from the DOCX's hyperlinks (reader-populated field).
/// The reader stores hyperlinks as `(rid, url, type)` in `docx.hyperlinks`.
fn build_hyperlink_map(docx: &docx_rs::Docx) -> HyperlinkMap {
    docx.hyperlinks
        .iter()
        .map(|(rid, url, _type)| (rid.clone(), url.clone()))
        .collect()
}

/// Build a lookup map from the DOCX's embedded images.
/// docx-rs converts all images to PNG; we use the PNG bytes.
fn build_image_map(docx: &docx_rs::Docx) -> ImageMap {
    docx.images
        .iter()
        .map(|(id, _path, _image, png)| (id.clone(), png.0.clone()))
        .collect()
}

/// Convert EMU (English Metric Units) to points.
/// Kept in the parent module because multiple DOCX helper submodules use it.
fn emu_to_pt(emu: u32) -> f64 {
    emu as f64 / 12700.0
}

impl Parser for DocxParser {
    fn parse(
        &self,
        data: &[u8],
        _options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        // Open ZIP once and build all pre-parse contexts from a single pass.
        // This consolidates what was previously 5 separate ZIP opens + multiple
        // reads of word/document.xml into a single archive + single doc read.
        let (metadata, mut ctx, mut math, mut chart_ctx, column_layouts, header_footer_assets) = {
            let cursor = std::io::Cursor::new(data);
            match zip::ZipArchive::new(cursor) {
                Ok(mut archive) => {
                    let metadata = crate::parser::metadata::extract_metadata_from_zip(&mut archive);
                    let doc_xml = read_zip_text(&mut archive, "word/document.xml");
                    let notes = build_note_context_from_xml(doc_xml.as_deref(), &mut archive);
                    let wraps = build_wrap_context_from_xml(doc_xml.as_deref());
                    let drawing_text_boxes = DrawingTextBoxContext::from_xml(doc_xml.as_deref());
                    let table_headers = TableHeaderContext::from_xml(doc_xml.as_deref());
                    let vml_text_boxes = VmlTextBoxContext::from_xml(doc_xml.as_deref());
                    let math = build_math_context_from_xml(doc_xml.as_deref());
                    let chart_ctx = build_chart_context_from_xml(doc_xml.as_deref(), &mut archive);
                    let column_layouts = doc_xml
                        .as_deref()
                        .map(scan_column_layouts)
                        .unwrap_or_default();
                    let bidi = BidiContext::from_xml(doc_xml.as_deref());
                    let small_caps = SmallCapsContext::from_xml(doc_xml.as_deref());
                    let header_footer_assets = build_header_footer_assets(&mut archive);
                    let ctx = DocxConversionContext {
                        notes,
                        wraps,
                        drawing_text_boxes,
                        table_headers,
                        vml_text_boxes,
                        bidi,
                        small_caps,
                    };
                    (metadata, ctx, math, chart_ctx, column_layouts, header_footer_assets)
                }
                Err(_) => {
                    // ZIP open failed — return empty contexts; docx-rs will
                    // produce a proper parse error downstream.
                    let ctx = DocxConversionContext {
                        notes: NoteContext::empty(),
                        wraps: WrapContext::empty(),
                        drawing_text_boxes: DrawingTextBoxContext::from_xml(None),
                        table_headers: TableHeaderContext::from_xml(None),
                        vml_text_boxes: VmlTextBoxContext::from_xml(None),
                        bidi: BidiContext::from_xml(None),
                        small_caps: SmallCapsContext::from_xml(None),
                    };
                    (
                        crate::ir::Metadata::default(),
                        ctx,
                        MathContext::empty(),
                        ChartContext::empty(),
                        Vec::new(),
                        HeaderFooterAssets::default(),
                    )
                }
            }
        };

        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX (docx-rs): {e}")))?;

        // Populate locale-specific footnote/endnote style IDs from docx styles
        ctx.notes.populate_style_ids(&docx.styles);

        let images = build_image_map(&docx);
        let hyperlinks = build_hyperlink_map(&docx);
        let numberings = build_numbering_map(&docx.numberings);
        let style_map = build_style_map(&docx.styles);
        let mut warnings: Vec<ConvertWarning> = Vec::new();

        let mut elements: Vec<TaggedElement> = Vec::new();
        let mut pages: Vec<Page> = Vec::new();
        let mut section_layout_index: usize = 0;
        for (idx, child) in docx.document.children.iter().enumerate() {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| match child {
                docx_rs::DocumentChild::Paragraph(para) => {
                    let mut tagged = vec![convert_paragraph_element(
                        para,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                    )];
                    // Inject math equations for this body child
                    let eqs = math.take(idx);
                    for eq in eqs {
                        tagged.push(TaggedElement::Plain(vec![Block::MathEquation(eq)]));
                    }
                    // Inject charts for this body child
                    let chs = chart_ctx.take(idx);
                    for ch in chs {
                        tagged.push(TaggedElement::Plain(vec![Block::Chart(ch)]));
                    }
                    tagged
                }
                docx_rs::DocumentChild::Table(table) => {
                    vec![TaggedElement::Plain(vec![Block::Table(convert_table(
                        table,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                        0,
                    ))])]
                }
                docx_rs::DocumentChild::StructuredDataTag(sdt) => convert_sdt_children(
                    sdt,
                    &images,
                    &hyperlinks,
                    &style_map,
                    &ctx,
                ),
                _ => vec![TaggedElement::Plain(vec![])],
            }));

            match result {
                Ok(elems) => elements.extend(elems),
                Err(panic_info) => {
                    let detail = if let Some(s) = panic_info.downcast_ref::<String>() {
                        s.clone()
                    } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                        (*s).to_string()
                    } else {
                        "unknown panic".to_string()
                    };
                    warnings.push(ConvertWarning::ParseSkipped {
                        format: "DOCX".to_string(),
                        reason: format!(
                            "upstream panic caught (docx-rs): element at index {idx}: {detail}"
                        ),
                    });
                }
            }

            if let docx_rs::DocumentChild::Paragraph(para) = child
                && let Some(section_prop) = para.property.section_property.as_ref()
            {
                let column_layout = match column_layouts.get(section_layout_index) {
                    Some(layout) => layout.clone(),
                    None => extract_column_layout_from_section_property(section_prop),
                };
                pages.push(Page::Flow(build_flow_page_from_section(
                    section_prop,
                    std::mem::take(&mut elements),
                    &numberings,
                    &header_footer_assets,
                    column_layout,
                    &mut warnings,
                )));
                section_layout_index += 1;
            }
        }

        let final_column_layout = match column_layouts.get(section_layout_index) {
            Some(layout) => layout.clone(),
            None => extract_column_layout_from_section_property(&docx.document.section_property),
        };
        pages.push(Page::Flow(build_flow_page_from_section(
            &docx.document.section_property,
            elements,
            &numberings,
            &header_footer_assets,
            final_column_layout,
            &mut warnings,
        )));

        Ok((
            Document {
                metadata,
                pages,
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

/// Extract content from a StructuredDataTag (SDT), processing its paragraph
/// and table children through the standard conversion pipeline.
/// SDTs are used for various structured content in DOCX, including Table of Contents.
fn convert_sdt_children(
    sdt: &docx_rs::StructuredDataTag,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> Vec<TaggedElement> {
    let mut result = Vec::new();
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                result.push(convert_paragraph_element(
                    para, images, hyperlinks, style_map, ctx,
                ));
            }
            docx_rs::StructuredDataTagChild::Table(table) => {
                result.push(TaggedElement::Plain(vec![Block::Table(convert_table(
                    table, images, hyperlinks, style_map, ctx, 0,
                ))]));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                result.extend(convert_sdt_children(
                    nested, images, hyperlinks, style_map, ctx,
                ));
            }
            _ => {}
        }
    }
    result
}

/// Convert a docx-rs Paragraph into a TaggedElement.
/// If the paragraph has numbering, returns a `ListParagraph`; otherwise `Plain`.
fn convert_paragraph_element(
    para: &docx_rs::Paragraph,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> TaggedElement {
    let num_info = extract_num_info(para);

    // Build the paragraph IR
    let mut blocks = Vec::new();
    convert_paragraph_blocks(para, &mut blocks, images, hyperlinks, style_map, ctx);

    match num_info {
        Some(info) => {
            // Extract the actual Paragraph from the blocks.
            // List paragraphs may also produce page breaks and images before the paragraph.
            let mut pre_blocks = Vec::new();
            let mut paragraph = None;
            for block in blocks {
                match block {
                    Block::Paragraph(p) if paragraph.is_none() => {
                        paragraph = Some(p);
                    }
                    _ => pre_blocks.push(block),
                }
            }
            if !pre_blocks.is_empty() {
                // If there were pre-blocks (page break, images), emit them as plain first.
                // We return the plain blocks — the caller will see them before the list paragraph.
                // For simplicity, we create a combined: Plain(pre) + ListParagraph.
                // But TaggedElement is a single value, so we need to handle this differently.
                // Actually, let's just emit them as plain first. The caller handles ordering.
                // Since we can only return one TaggedElement, fold the pre-blocks into the
                // paragraph by noting that list items in a list won't have page breaks.
                // For now, treat the paragraph as a plain block if it has pre-blocks.
                pre_blocks.push(Block::Paragraph(paragraph.unwrap_or_else(|| Paragraph {
                    style: ParagraphStyle::default(),
                    runs: Vec::new(),
                })));
                TaggedElement::Plain(pre_blocks)
            } else if let Some(p) = paragraph {
                TaggedElement::ListParagraph { info, paragraph: p }
            } else {
                TaggedElement::Plain(vec![])
            }
        }
        None => TaggedElement::Plain(blocks),
    }
}

/// Convert a docx-rs Paragraph to IR blocks, handling page breaks and inline images.
/// If the paragraph has `page_break_before`, a `Block::PageBreak` is emitted first.
/// Inline images within runs are extracted as separate `Block::Image` elements.
/// Style formatting from the document's style definitions is merged with explicit formatting.
fn convert_paragraph_blocks(
    para: &docx_rs::Paragraph,
    out: &mut Vec<Block>,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) {
    // Check bidi direction for this paragraph (must be called once per XML <w:p>)
    let is_rtl = ctx.bidi.next_is_bidi();

    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    // Look up the paragraph's referenced style
    let resolved_style = get_paragraph_style_id(&para.property)
        .and_then(|id| style_map.get(id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID));

    // Collect text runs and detect inline images
    let mut runs: Vec<Run> = Vec::new();
    let mut inline_images: Vec<Block> = Vec::new();
    let mut emitted_text_box_blocks: bool = false;

    for child in &para.children {
        match child {
            docx_rs::ParagraphChild::Run(run) => {
                // Advance smallCaps cursor for every <w:r> in body
                let is_small_caps: bool = ctx.small_caps.next_is_small_caps();

                // Check for footnote/endnote reference runs
                if is_note_reference_run(run, &ctx.notes) {
                    if let Some(content) = ctx.notes.consume_next() {
                        runs.push(Run {
                            text: String::new(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: Some(content),
                        });
                    }
                    continue;
                }

                // Check for column breaks and embedded drawings in this run.
                let mut has_column_break = false;
                let mut text_box_blocks: Vec<Block> = Vec::new();
                for run_child in &run.children {
                    if let docx_rs::RunChild::Drawing(drawing) = run_child
                        && let Some(img_block) =
                            extract_drawing_image(drawing, images, &ctx.wraps)
                    {
                        inline_images.push(img_block);
                    }
                    if let docx_rs::RunChild::Drawing(drawing) = run_child {
                        text_box_blocks.extend(extract_drawing_text_box_blocks(
                            drawing, images, hyperlinks, style_map, ctx,
                        ));
                    }
                    if let docx_rs::RunChild::Shape(shape) = run_child {
                        let vml_text_box: VmlTextBoxInfo = ctx.vml_text_boxes.consume_next();
                        if let Some(floating_text_box) =
                            extract_vml_shape_text_box(shape, &vml_text_box)
                        {
                            text_box_blocks.push(Block::FloatingTextBox(floating_text_box));
                        } else {
                            text_box_blocks.extend(vml_text_box.into_blocks());
                        }

                        if let Some(img_block) = extract_shape_image(shape, images) {
                            inline_images.push(img_block);
                        }
                    }
                    if let docx_rs::RunChild::Break(br) = run_child
                        && is_column_break(br)
                    {
                        has_column_break = true;
                    }
                }

                if !text_box_blocks.is_empty() {
                    if !runs.is_empty() {
                        out.append(&mut inline_images);
                        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
                    } else if !inline_images.is_empty() {
                        out.append(&mut inline_images);
                    }
                    emitted_text_box_blocks = true;
                    out.extend(text_box_blocks);
                }

                if has_column_break {
                    // Flush current runs as a paragraph before the column break
                    if !runs.is_empty() {
                        out.append(&mut inline_images);
                        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
                    }
                    out.push(Block::ColumnBreak);

                    // Still extract any text from this run (after the break)
                    let text = extract_run_text_skip_column_breaks(run);
                    if !text.is_empty() {
                        let mut explicit_style: TextStyle = extract_run_style(&run.run_property);
                        if is_small_caps {
                            explicit_style.small_caps = Some(true);
                        }
                        runs.push(Run {
                            text,
                            style: merge_text_style(&explicit_style, resolved_style),
                            href: None,
                            footnote: None,
                        });
                    }
                } else {
                    // Extract text from the run
                    let text = extract_run_text(run);
                    if !text.is_empty() {
                        let mut explicit_style: TextStyle = extract_run_style(&run.run_property);
                        if is_small_caps {
                            explicit_style.small_caps = Some(true);
                        }
                        runs.push(Run {
                            text,
                            style: merge_text_style(&explicit_style, resolved_style),
                            href: None,
                            footnote: None,
                        });
                    }
                }
            }
            docx_rs::ParagraphChild::Hyperlink(hyperlink) => {
                // Resolve the hyperlink URL from document relationships
                let href = resolve_hyperlink_url(hyperlink, hyperlinks);

                // Extract runs from inside the hyperlink element
                for hchild in &hyperlink.children {
                    if let docx_rs::ParagraphChild::Run(run) = hchild {
                        let hl_small_caps: bool = ctx.small_caps.next_is_small_caps();
                        let text = extract_run_text(run);
                        if !text.is_empty() {
                            let mut explicit_style: TextStyle =
                                extract_run_style(&run.run_property);
                            if hl_small_caps {
                                explicit_style.small_caps = Some(true);
                            }
                            runs.push(Run {
                                text,
                                style: merge_text_style(&explicit_style, resolved_style),
                                href: href.clone(),
                                footnote: None,
                            });
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // Emit image blocks before the paragraph (inline images are block-level in our IR)
    out.extend(inline_images);

    if !runs.is_empty() || !emitted_text_box_blocks {
        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
    }
}

fn push_paragraph_from_runs(
    out: &mut Vec<Block>,
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    is_rtl: bool,
    runs: &mut Vec<Run>,
) {
    let explicit_para_style = extract_paragraph_style(&para.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&para.property.tabs);
    let mut style = merge_paragraph_style(
        &explicit_para_style,
        explicit_tab_overrides.as_deref(),
        resolved_style,
    );
    if is_rtl {
        style.direction = Some(TextDirection::Rtl);
    }
    out.push(Block::Paragraph(Paragraph {
        style,
        runs: std::mem::take(runs),
    }));
}

#[cfg(test)]
#[path = "docx_tests.rs"]
mod tests;
