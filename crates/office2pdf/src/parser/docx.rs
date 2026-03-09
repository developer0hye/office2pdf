use std::collections::HashMap;
#[cfg(test)]
use std::io::Read;

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
    BidiContext, ChartContext, DrawingTextBoxContext, DrawingTextBoxInfo, MathContext, NoteContext,
    SmallCapsContext, TableHeaderContext, VmlTextBoxContext, VmlTextBoxInfo, WrapContext,
    build_chart_context_from_xml, build_math_context_from_xml, build_note_context_from_xml,
    build_wrap_context_from_xml, extract_column_layout_from_section_property,
    is_note_reference_run, read_zip_text, scan_column_layouts,
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
        let (
            metadata,
            mut notes,
            wraps,
            drawing_text_boxes,
            table_headers,
            vml_text_boxes,
            mut math,
            mut chart_ctx,
            column_layouts,
            bidi,
            small_caps,
            header_footer_assets,
        ) = {
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
                    (
                        metadata,
                        notes,
                        wraps,
                        drawing_text_boxes,
                        table_headers,
                        vml_text_boxes,
                        math,
                        chart_ctx,
                        column_layouts,
                        bidi,
                        small_caps,
                        header_footer_assets,
                    )
                }
                Err(_) => {
                    // ZIP open failed — return empty contexts; docx-rs will
                    // produce a proper parse error downstream.
                    (
                        crate::ir::Metadata::default(),
                        NoteContext::empty(),
                        WrapContext::empty(),
                        DrawingTextBoxContext::from_xml(None),
                        TableHeaderContext::from_xml(None),
                        VmlTextBoxContext::from_xml(None),
                        MathContext::empty(),
                        ChartContext::empty(),
                        Vec::new(),
                        BidiContext::from_xml(None),
                        SmallCapsContext::from_xml(None),
                        HeaderFooterAssets::default(),
                    )
                }
            }
        };

        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX (docx-rs): {e}")))?;

        // Populate locale-specific footnote/endnote style IDs from docx styles
        notes.populate_style_ids(&docx.styles);

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
                        &notes,
                        &wraps,
                        &drawing_text_boxes,
                        &table_headers,
                        &vml_text_boxes,
                        &bidi,
                        &small_caps,
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
                        &notes,
                        &wraps,
                        &drawing_text_boxes,
                        &table_headers,
                        &vml_text_boxes,
                        &bidi,
                        &small_caps,
                        0,
                    ))])]
                }
                docx_rs::DocumentChild::StructuredDataTag(sdt) => convert_sdt_children(
                    sdt,
                    &images,
                    &hyperlinks,
                    &style_map,
                    &notes,
                    &wraps,
                    &drawing_text_boxes,
                    &table_headers,
                    &vml_text_boxes,
                    &bidi,
                    &small_caps,
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
#[allow(clippy::too_many_arguments)]
fn convert_sdt_children(
    sdt: &docx_rs::StructuredDataTag,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) -> Vec<TaggedElement> {
    let mut result = Vec::new();
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                result.push(convert_paragraph_element(
                    para,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                ));
            }
            docx_rs::StructuredDataTagChild::Table(table) => {
                result.push(TaggedElement::Plain(vec![Block::Table(convert_table(
                    table,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                    0,
                ))]));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                result.extend(convert_sdt_children(
                    nested,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                ));
            }
            _ => {}
        }
    }
    result
}

/// Convert a docx-rs Paragraph into a TaggedElement.
/// If the paragraph has numbering, returns a `ListParagraph`; otherwise `Plain`.
#[allow(clippy::too_many_arguments)]
fn convert_paragraph_element(
    para: &docx_rs::Paragraph,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) -> TaggedElement {
    let num_info = extract_num_info(para);

    // Build the paragraph IR
    let mut blocks = Vec::new();
    convert_paragraph_blocks(
        para,
        &mut blocks,
        images,
        hyperlinks,
        style_map,
        notes,
        wraps,
        drawing_text_boxes,
        table_headers,
        vml_text_boxes,
        bidi,
        small_caps,
    );

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
#[allow(clippy::too_many_arguments)]
fn convert_paragraph_blocks(
    para: &docx_rs::Paragraph,
    out: &mut Vec<Block>,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) {
    // Check bidi direction for this paragraph (must be called once per XML <w:p>)
    let is_rtl = bidi.next_is_bidi();

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
                let is_small_caps: bool = small_caps.next_is_small_caps();

                // Check for footnote/endnote reference runs
                if is_note_reference_run(run, notes) {
                    if let Some(content) = notes.consume_next() {
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
                        && let Some(img_block) = extract_drawing_image(drawing, images, wraps)
                    {
                        inline_images.push(img_block);
                    }
                    if let docx_rs::RunChild::Drawing(drawing) = run_child {
                        text_box_blocks.extend(extract_drawing_text_box_blocks(
                            drawing,
                            images,
                            hyperlinks,
                            style_map,
                            notes,
                            wraps,
                            drawing_text_boxes,
                            table_headers,
                            vml_text_boxes,
                            bidi,
                            small_caps,
                        ));
                    }
                    if let docx_rs::RunChild::Shape(shape) = run_child {
                        let vml_text_box: VmlTextBoxInfo = vml_text_boxes.consume_next();
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
                        let hl_small_caps: bool = small_caps.next_is_small_caps();
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

/// Extract paragraph-level formatting from docx-rs ParagraphProperty.
fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
    let alignment = prop.alignment.as_ref().and_then(|j| match j.val.as_str() {
        "center" => Some(Alignment::Center),
        "right" | "end" => Some(Alignment::Right),
        "left" | "start" => Some(Alignment::Left),
        "both" | "justified" => Some(Alignment::Justify),
        _ => None,
    });

    let (indent_left, indent_right, indent_first_line) = extract_indent(&prop.indent);

    let (line_spacing, space_before, space_after) = extract_line_spacing(&prop.line_spacing);

    let tab_stops = extract_tab_stops(&prop.tabs);

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        space_before,
        space_after,
        heading_level: None,
        direction: None, // Set by BidiContext after style merge
        tab_stops,
    }
}

/// Extract indentation values from docx-rs Indent.
/// All values in docx-rs are in twips; convert to points (÷20).
fn extract_indent(indent: &Option<docx_rs::Indent>) -> (Option<f64>, Option<f64>, Option<f64>) {
    let Some(indent) = indent else {
        return (None, None, None);
    };

    let left = indent.start.map(|v| v as f64 / 20.0);
    let right = indent.end.map(|v| v as f64 / 20.0);
    let first_line = indent.special_indent.map(|si| match si {
        docx_rs::SpecialIndentType::FirstLine(v) => v as f64 / 20.0,
        docx_rs::SpecialIndentType::Hanging(v) => -(v as f64 / 20.0),
    });

    (left, right, first_line)
}

/// Extract line spacing from docx-rs LineSpacing (private fields, use serde).
///
/// OOXML line spacing:
/// - Auto: `line` is in 240ths of a line (240=single, 360=1.5×, 480=double)
/// - Exact/AtLeast: `line` is in twips (÷20 → points)
/// - `before`/`after`: twips (÷20 → points)
fn extract_line_spacing(
    spacing: &Option<docx_rs::LineSpacing>,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
    let Some(spacing) = spacing else {
        return (None, None, None);
    };

    let json = match serde_json::to_value(spacing) {
        Ok(j) => j,
        Err(_) => return (None, None, None),
    };

    let space_before = json
        .get("before")
        .and_then(|v| v.as_f64())
        .map(|v| v / 20.0);
    let space_after = json.get("after").and_then(|v| v.as_f64()).map(|v| v / 20.0);

    let line_spacing = json.get("line").and_then(|line_val| {
        let line = line_val.as_f64()?;
        let rule = json.get("lineRule").and_then(|v| v.as_str());
        match rule {
            Some("exact") | Some("atLeast") => Some(LineSpacing::Exact(line / 20.0)),
            _ => {
                // Auto: 240ths of a line
                Some(LineSpacing::Proportional(line / 240.0))
            }
        }
    });

    (line_spacing, space_before, space_after)
}

/// Extract tab stops from paragraph properties.
/// docx-rs Tab has `pos` in twips and `val`/`leader` as enums.
fn extract_tab_stops(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStop>> {
    let tab_overrides = extract_tab_stop_overrides(tabs)?;
    let mut tab_stops: Vec<TabStop> = Vec::new();
    apply_tab_stop_overrides(&mut tab_stops, &tab_overrides);
    Some(tab_stops)
}

fn extract_tab_stop_overrides(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStopOverride>> {
    if tabs.is_empty() {
        return None;
    }

    Some(
        tabs.iter()
            .filter_map(|tab| {
                let position = tab.pos.map(|pos_twips| pos_twips as f64 / 20.0)?;

                if matches!(tab.val, Some(docx_rs::TabValueType::Clear)) {
                    return Some(TabStopOverride::Clear(position));
                }

                let alignment = match tab.val {
                    Some(docx_rs::TabValueType::Center) => TabAlignment::Center,
                    Some(docx_rs::TabValueType::Right) | Some(docx_rs::TabValueType::End) => {
                        TabAlignment::Right
                    }
                    Some(docx_rs::TabValueType::Decimal) => TabAlignment::Decimal,
                    _ => TabAlignment::Left,
                };

                let leader =
                    match tab.leader {
                        Some(docx_rs::TabLeaderType::Dot)
                        | Some(docx_rs::TabLeaderType::MiddleDot) => TabLeader::Dot,
                        Some(docx_rs::TabLeaderType::Hyphen)
                        | Some(docx_rs::TabLeaderType::Heavy) => TabLeader::Hyphen,
                        Some(docx_rs::TabLeaderType::Underscore) => TabLeader::Underscore,
                        _ => TabLeader::None,
                    };

                Some(TabStopOverride::Set(TabStop {
                    position,
                    alignment,
                    leader,
                }))
            })
            .collect(),
    )
}

/// Extract inline text style from a docx-rs RunProperty.
///
/// docx-rs types serialize directly as their inner value (e.g. Bold → `true`,
/// Sz → `24`, Color → `"FF0000"`), so JSON extraction works for both explicit
/// run properties and docDefaults run properties.
fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    let json = serde_json::to_value(rp).unwrap_or(serde_json::Value::Null);
    extract_run_style_from_json(&json)
}

/// Extract inline text style from a serialized RunProperty-like JSON object.
fn extract_run_style_from_json(rp: &serde_json::Value) -> TextStyle {
    let vertical_align: Option<VerticalTextAlign> =
        rp.get("vertAlign").and_then(|va| match va.as_str()? {
            "superscript" => Some(VerticalTextAlign::Superscript),
            "subscript" => Some(VerticalTextAlign::Subscript),
            _ => None,
        });

    let all_caps: Option<bool> = rp.get("caps").and_then(serde_json::Value::as_bool);

    TextStyle {
        bold: rp.get("bold").and_then(serde_json::Value::as_bool),
        italic: rp.get("italic").and_then(serde_json::Value::as_bool),
        underline: rp
            .get("underline")
            .and_then(|u| u.as_str())
            .and_then(|val| if val == "none" { None } else { Some(true) }),
        strikethrough: rp.get("strike").and_then(json_bool_or_val),
        font_size: rp
            .get("sz")
            .and_then(serde_json::Value::as_f64)
            .map(|half_points| half_points / 2.0),
        color: rp
            .get("color")
            .and_then(serde_json::Value::as_str)
            .and_then(parse_hex_color),
        font_family: rp.get("fonts").and_then(|fonts| {
            fonts
                .get("ascii")
                .or_else(|| fonts.get("hiAnsi"))
                .or_else(|| fonts.get("eastAsia"))
                .or_else(|| fonts.get("cs"))
                .and_then(serde_json::Value::as_str)
                .map(String::from)
        }),
        highlight: rp
            .get("highlight")
            .and_then(serde_json::Value::as_str)
            .and_then(resolve_highlight_color),
        vertical_align,
        all_caps,
        // smallCaps is not exposed by docx-rs; set via SmallCapsContext XML scan
        small_caps: None,
        // character_spacing is in twips (1/20 pt); convert to points
        letter_spacing: rp
            .get("characterSpacing")
            .and_then(serde_json::Value::as_i64)
            .map(|twips| twips as f64 / 20.0),
    }
}

fn json_bool_or_val(value: &serde_json::Value) -> Option<bool> {
    value
        .as_bool()
        .or_else(|| value.get("val").and_then(serde_json::Value::as_bool))
}

/// Extract document-level default text style from styles.xml docDefaults.
fn extract_doc_default_text_style(styles: &docx_rs::Styles) -> TextStyle {
    let Ok(json) = serde_json::to_value(&styles.doc_defaults) else {
        return TextStyle::default();
    };
    let Some(run_property) = json
        .get("runPropertyDefault")
        .and_then(|value| value.get("runProperty"))
    else {
        return TextStyle::default();
    };

    extract_run_style_from_json(run_property)
}

/// Map OOXML named highlight colors to RGB values.
/// The 16 named colors are defined in the ECMA-376 spec (ST_HighlightColor).
fn resolve_highlight_color(name: &str) -> Option<Color> {
    match name {
        "yellow" => Some(Color::new(255, 255, 0)),
        "green" => Some(Color::new(0, 255, 0)),
        "cyan" => Some(Color::new(0, 255, 255)),
        "magenta" => Some(Color::new(255, 0, 255)),
        "blue" => Some(Color::new(0, 0, 255)),
        "red" => Some(Color::new(255, 0, 0)),
        "darkBlue" => Some(Color::new(0, 0, 128)),
        "darkCyan" => Some(Color::new(0, 128, 128)),
        "darkGreen" => Some(Color::new(0, 128, 0)),
        "darkMagenta" => Some(Color::new(128, 0, 128)),
        "darkRed" => Some(Color::new(128, 0, 0)),
        "darkYellow" => Some(Color::new(128, 128, 0)),
        "darkGray" => Some(Color::new(128, 128, 128)),
        "lightGray" => Some(Color::new(192, 192, 192)),
        "black" => Some(Color::new(0, 0, 0)),
        "white" => Some(Color::new(255, 255, 255)),
        _ => None, // "none" or unrecognized
    }
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Resolve a hyperlink's URL from document relationships.
/// For external hyperlinks, looks up the relationship ID in the hyperlink map.
/// For anchor hyperlinks (internal bookmarks), returns None.
fn resolve_hyperlink_url(
    hyperlink: &docx_rs::Hyperlink,
    hyperlinks: &HyperlinkMap,
) -> Option<String> {
    match &hyperlink.link {
        docx_rs::HyperlinkData::External { rid, path } => {
            // First try the path (populated during writing),
            // then fall back to the relationship map (populated during reading)
            if !path.is_empty() {
                Some(path.clone())
            } else {
                hyperlinks.get(rid).cloned()
            }
        }
        docx_rs::HyperlinkData::Anchor { .. } => None, // internal bookmark, skip
    }
}

/// Check if a docx-rs Break is a column break.
/// Break.break_type is private, so we use serde to extract the value.
fn is_column_break(br: &docx_rs::Break) -> bool {
    serde_json::to_value(br)
        .ok()
        .and_then(|v| {
            v.get("breakType")
                .and_then(|bt| bt.as_str().map(|s| s == "column"))
        })
        .unwrap_or(false)
}

/// Extract text content from a docx-rs Run, skipping column breaks.
/// Column breaks are handled separately as Block::ColumnBreak.
fn extract_run_text_skip_column_breaks(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(br) => {
                if !is_column_break(br) {
                    text.push('\n');
                }
            }
            _ => {}
        }
    }
    text
}

/// Extract text content from a docx-rs Run.
fn extract_run_text(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(_) => text.push('\n'),
            _ => {}
        }
    }
    text
}

#[cfg(test)]
#[path = "docx_tests.rs"]
mod tests;
