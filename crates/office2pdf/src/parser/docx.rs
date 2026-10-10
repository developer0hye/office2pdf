use crate::parser::xml_util::OOXML_XML_VERSION;
use std::collections::HashMap;
use std::io::Read;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};

/// Maximum nesting depth for tables-within-tables.  Deeper nesting is silently
/// truncated to prevent stack overflow on pathological documents.
const MAX_TABLE_DEPTH: usize = 64;
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, Caption, CellBorder, CellVerticalAlign, Color,
    ColumnLayout, Document, DocumentLanguage, FloatingImage, FloatingImageVerticalAnchor,
    FloatingTable, FloatingTextBox, ImageData, ImageFormat, ImageParagraphSpacing, InlineTextBox,
    Insets, LineCap, LineJoin, LineSpacing, Page, PageNumbering, PairKerning, Paragraph,
    ParagraphStyle, Run, StyleSheet, TabAlignment, TabLeader, TabStop, Table, TableCell,
    TableOfContents, TableRow, TextDirection, TextStyle, VerticalTextAlign, WordCompatibilityMode,
};
use crate::parser::Parser;

#[cfg(test)]
use self::contexts::scan_table_headers;
use self::contexts::{
    BidiContext, ChartContext, ContextualSpacingContext, DocxConversionContext,
    DrawingShapeContext, DrawingTextBoxContext, DrawingTextBoxInfo, FieldContext, MathContext,
    NoteContent, NoteContext, ParagraphContextualSpacing, ParagraphMarkContext,
    ParagraphShadingContext, SmallCapsContext, TableHeaderContext, TableStyleContext,
    VmlTextBoxContext, VmlTextBoxInfo, WordWrapContext, WpgDrawingInfo, WrapContext,
    build_chart_context_from_xml, build_math_context_from_xml, build_note_context_from_xml,
    build_wrap_context_from_xml, extract_column_layout_from_section_property,
    is_note_reference_run, read_zip_text, scan_column_layouts, scan_page_numbering,
    scan_style_paragraph_shading, scan_style_word_wrap, seq_identifier, toc_caption_identifier,
    toc_heading_depth,
};
use self::lists::{
    NumberingMap, TaggedElement, build_numbering_map, extract_num_info, group_into_lists,
};
use self::media::{
    DrawingTextBoxPlacement, extract_drawing_image, extract_drawing_text_box_blocks,
    extract_shape_image, extract_vml_shape_text_box,
};
#[cfg(test)]
use self::sections::extract_page_size;
use self::sections::{
    HeaderFooterAssets, HeaderFooterStyleContext, SectionHeaderFooterInputs, SectionOverrides,
    build_flow_page_from_section, build_header_footer_assets,
};
use self::styles::{
    DOC_DEFAULT_STYLE_ID, PairKerningRules, ResolvedStyle, StyleMap, TabStopOverride,
    apply_tab_stop_overrides, build_style_map, get_paragraph_style_id, merge_paragraph_style,
    merge_text_style, resolve_doc_default_text_style,
};
use self::tables::{convert_table, convert_table_in_text_box};
use self::text::{
    ThemeFonts, extract_doc_default_paragraph_style, extract_doc_default_text_style_with_theme,
    extract_paragraph_style, extract_run_style, extract_run_style_id, extract_run_text,
    extract_tab_stop_overrides, insert_east_asian_auto_space, is_column_break, is_page_break,
    pair_kerning_from_half_points, parse_hex_color, parse_theme_fonts, resolve_hyperlink_url,
    resolve_latin_theme_font_family, resolve_theme_font_family,
};
#[cfg(test)]
use self::text::{extract_pair_kerning, extract_tab_stops, resolve_highlight_color};

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

pub(super) use sections::parse_docx_shape_gradient;

/// Parser for DOCX (Office Open XML Word) documents.
pub struct DocxParser;

/// The gap Word's built-in `Normal` opens below a paragraph that states no
/// `w:spacing w:after` anywhere in its style hierarchy — `w:after="160"`.
///
/// Measured on native Word exports of a package that states no `w:spacing` at
/// all (Malgun Gothic 10.5pt, intra-paragraph pitch 18.24pt, paragraph pitch
/// 26.16pt): patching every `w:pPr` to `w:after="160"` reproduces the
/// untouched export exactly, `w:after="0"` pulls the page up 24.00pt over the
/// three gaps and `w:after="240"` pushes it down 12.00pt. Issue #1085, probe
/// `issue-1085-space-after-declared`.
pub(super) const WORD_BUILT_IN_NORMAL_SPACE_AFTER_PT: f64 = 8.0;

/// The same gap once the document declares `w:docDefaults/w:pPrDefault`:
/// ECMA-376 leaves an unstated `w:after` at zero, and the declaration is the
/// document taking the defaults over from Word's built-in `Normal`.
///
/// The element's mere presence is the whole signal — `<w:pPrDefault/>`,
/// `<w:pPrDefault><w:pPr/></w:pPrDefault>` and a `w:pPr` carrying only
/// `w:before` all export at the same baselines an explicit `w:after="0"` does,
/// while a `Normal` style carrying its own `w:pPr` keeps the 8pt (issue #1085,
/// probes `issue-1085-space-after-pprdefault` and `-default-shape`).
pub(super) const WORD_DECLARED_DEFAULT_SPACE_AFTER_PT: f64 = 0.0;

/// The `w:spacing w:after` to fall back on, by what `styles.xml` declares.
///
/// Recording it explicitly (rather than leaving `space_after` unset) also pins
/// the paragraph block's `below`, so Typst's own 1.2em default block spacing
/// cannot leak into the gap.
///
/// Line height is left to the renderer, which derives Word's single-spacing
/// pitch from the actual font metrics (issues #354, #452).
pub(super) fn word_compatible_paragraph_space_after_pt(
    paragraph_property_defaults_are_declared: bool,
) -> f64 {
    if paragraph_property_defaults_are_declared {
        WORD_DECLARED_DEFAULT_SPACE_AFTER_PT
    } else {
        WORD_BUILT_IN_NORMAL_SPACE_AFTER_PT
    }
}

fn apply_word_compatible_paragraph_defaults(
    style: &mut ParagraphStyle,
    paragraph_property_defaults_are_declared: bool,
) {
    style
        .space_after
        .get_or_insert(word_compatible_paragraph_space_after_pt(
            paragraph_property_defaults_are_declared,
        ));
}

#[derive(Clone)]
struct DocxImageAsset {
    data: Vec<u8>,
    format: ImageFormat,
}

/// Map from relationship ID to normalized image assets.
type ImageMap = HashMap<String, DocxImageAsset>;

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
///
/// docx-rs decodes raster pictures and re-encodes them as PNG previews, and
/// we use those bytes. A format it cannot decode, such as EMF, arrives with
/// an empty preview. Those entries are left out: the metafile converter
/// supplies every EMF/WMF it can draw, and a picture nothing can render is
/// skipped rather than handed to Typst as an empty PNG, which fails the whole
/// compilation.
fn build_image_map(docx: &docx_rs::Docx) -> ImageMap {
    docx.images
        .iter()
        .filter(|(_id, _path, _image, png)| !png.0.is_empty())
        .map(|(id, _path, _image, png)| {
            (
                id.clone(),
                DocxImageAsset {
                    data: png.0.clone(),
                    format: ImageFormat::Png,
                },
            )
        })
        .collect()
}

fn build_document_metafile_image_map<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> ImageMap {
    let Some(relationships_xml) = read_zip_text(archive, "word/_rels/document.xml.rels") else {
        return ImageMap::new();
    };
    let mut reader = quick_xml::Reader::from_str(&relationships_xml);
    let mut relationships: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element))
                if element.local_name().as_ref() == b"Relationship" =>
            {
                let mut id: Option<String> = None;
                let mut target: Option<String> = None;
                let mut is_image: bool = false;
                for attribute in element.attributes().flatten() {
                    let Ok(value) = attribute.normalized_value(OOXML_XML_VERSION) else {
                        continue;
                    };
                    match attribute.key.local_name().as_ref() {
                        b"Id" => id = Some(value.to_string()),
                        b"Target" => target = Some(value.to_string()),
                        b"Type" => is_image = value.ends_with("/image"),
                        _ => {}
                    }
                }
                if is_image && let (Some(id), Some(target)) = (id, target) {
                    let lowercase_target: String = target.to_ascii_lowercase();
                    if lowercase_target.ends_with(".emf") || lowercase_target.ends_with(".wmf") {
                        relationships.push((id, target));
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    relationships
        .into_iter()
        .filter_map(|(id, target)| {
            let path = format!("word/{}", target.trim_start_matches('/'));
            let mut data: Vec<u8> = Vec::new();
            archive.by_name(&path).ok()?.read_to_end(&mut data).ok()?;
            let svg: Vec<u8> = if target.to_ascii_lowercase().ends_with(".wmf") {
                crate::parser::wmf::convert_wmf_to_svg(&data)?
            } else {
                crate::parser::emf::convert_emf_to_svg(&data)?
            };
            Some((
                id,
                DocxImageAsset {
                    data: svg,
                    format: ImageFormat::Svg,
                },
            ))
        })
        .collect()
}

/// Pre-parsed assets extracted from the DOCX ZIP archive before docx-rs parsing.
struct ZipPreParseAssets {
    metadata: crate::ir::Metadata,
    ctx: DocxConversionContext,
    math: MathContext,
    chart_ctx: ChartContext,
    column_layouts: Vec<Option<ColumnLayout>>,
    page_numbering: Vec<Option<PageNumbering>>,
    metafile_images: ImageMap,
    theme_fonts: ThemeFonts,
    default_paragraph_style_id: Option<String>,
    style_paragraph_backgrounds: HashMap<String, Color>,
    style_word_wraps: HashMap<String, bool>,
    /// Read from the raw `word/styles.xml` because docx-rs has no field for
    /// `w:kern` (issue #628).
    pair_kerning: PairKerningRules,
    /// Read from the raw `word/styles.xml` because docx-rs has no field for
    /// `w:lang`.
    default_language: Option<DocumentLanguage>,
}

/// Build all pre-parse contexts from the DOCX ZIP in a single pass.
/// Falls back to empty contexts if the ZIP cannot be opened, letting
/// docx-rs produce a proper parse error downstream.
fn build_zip_preparse_assets(data: &[u8]) -> ZipPreParseAssets {
    match crate::parser::open_zip(data) {
        Ok(mut archive) => {
            let metadata = crate::parser::metadata::extract_metadata_from_zip(&mut archive);
            let doc_xml = read_zip_text(&mut archive, "word/document.xml");
            let styles_xml = read_zip_text(&mut archive, "word/styles.xml");
            let default_paragraph_style_id = styles_xml
                .as_deref()
                .and_then(styles::scan_default_paragraph_style_id);
            let style_paragraph_backgrounds = scan_style_paragraph_shading(styles_xml.as_deref());
            let style_word_wraps = scan_style_word_wrap(styles_xml.as_deref());
            // Word writes the default paragraph style as `Normal`; docx-rs
            // writes it without the `w:default` flag.
            let default_language: Option<DocumentLanguage> =
                styles_xml.as_deref().and_then(|xml| {
                    styles::scan_default_language(
                        xml,
                        default_paragraph_style_id.as_deref().unwrap_or("Normal"),
                    )
                });
            let theme_xml = read_zip_text(&mut archive, "word/theme/theme1.xml");
            let theme_fonts = theme_xml
                .as_deref()
                .map(parse_theme_fonts)
                .unwrap_or_default();
            let notes = build_note_context_from_xml(doc_xml.as_deref(), &mut archive, &theme_fonts);
            let wraps = build_wrap_context_from_xml(doc_xml.as_deref());
            let drawing_text_boxes = DrawingTextBoxContext::from_xml_with_theme(
                doc_xml.as_deref(),
                theme_xml.as_deref(),
            );
            let drawing_shapes =
                DrawingShapeContext::from_xml_with_theme(doc_xml.as_deref(), theme_xml.as_deref());
            let table_headers = TableHeaderContext::from_xml(doc_xml.as_deref());
            let table_styles =
                TableStyleContext::from_xml(doc_xml.as_deref(), styles_xml.as_deref());
            let vml_text_boxes = VmlTextBoxContext::from_xml(doc_xml.as_deref());
            let math = build_math_context_from_xml(doc_xml.as_deref());
            let chart_ctx = build_chart_context_from_xml(doc_xml.as_deref(), &mut archive);
            let column_layouts = doc_xml
                .as_deref()
                .map(scan_column_layouts)
                .unwrap_or_default();
            let page_numbering = doc_xml
                .as_deref()
                .map(scan_page_numbering)
                .unwrap_or_default();
            let bidi = BidiContext::from_xml(doc_xml.as_deref());
            let small_caps = SmallCapsContext::from_xml(doc_xml.as_deref());
            let metafile_images = build_document_metafile_image_map(&mut archive);
            let ctx = DocxConversionContext {
                theme_fonts: theme_fonts.clone(),
                notes,
                wraps,
                drawing_text_boxes,
                drawing_shapes,
                table_headers,
                table_styles,
                vml_text_boxes,
                bidi,
                small_caps,
                paragraph_shading: ParagraphShadingContext::from_xml(doc_xml.as_deref()),
                word_wraps: WordWrapContext::from_xml(doc_xml.as_deref()),
                contextual_spacing: ContextualSpacingContext::from_xml(
                    doc_xml.as_deref(),
                    styles_xml.as_deref(),
                ),
                paragraph_marks: ParagraphMarkContext::from_xml(doc_xml.as_deref()),
                fields: FieldContext::default(),
                default_paragraph_style_is_defined: styles_xml
                    .as_deref()
                    .is_some_and(styles::scan_defines_default_paragraph_style),
                // A package with no `word/styles.xml` at all declares no
                // paragraph defaults either, so it takes the built-in gap.
                paragraph_property_defaults_are_declared: styles_xml
                    .as_deref()
                    .is_some_and(styles::scan_declares_paragraph_property_defaults),
            };
            ZipPreParseAssets {
                metadata,
                ctx,
                math,
                chart_ctx,
                column_layouts,
                page_numbering,
                metafile_images,
                theme_fonts,
                default_paragraph_style_id,
                style_paragraph_backgrounds,
                style_word_wraps,
                pair_kerning: PairKerningRules::from_styles_xml(styles_xml.as_deref()),
                default_language,
            }
        }
        Err(_) => ZipPreParseAssets {
            metadata: crate::ir::Metadata::default(),
            ctx: DocxConversionContext {
                theme_fonts: ThemeFonts::default(),
                notes: NoteContext::empty(),
                wraps: WrapContext::empty(),
                drawing_text_boxes: DrawingTextBoxContext::from_xml(None),
                drawing_shapes: DrawingShapeContext::from_xml(None),
                table_headers: TableHeaderContext::from_xml(None),
                table_styles: TableStyleContext::from_xml(None, None),
                vml_text_boxes: VmlTextBoxContext::from_xml(None),
                bidi: BidiContext::from_xml(None),
                small_caps: SmallCapsContext::from_xml(None),
                paragraph_shading: ParagraphShadingContext::from_xml(None),
                word_wraps: WordWrapContext::from_xml(None),
                contextual_spacing: ContextualSpacingContext::from_xml(None, None),
                paragraph_marks: ParagraphMarkContext::from_xml(None),
                fields: FieldContext::default(),
                default_paragraph_style_is_defined: false,
                paragraph_property_defaults_are_declared: false,
            },
            math: MathContext::empty(),
            chart_ctx: ChartContext::empty(),
            column_layouts: Vec::new(),
            page_numbering: Vec::new(),
            metafile_images: ImageMap::new(),
            theme_fonts: ThemeFonts::default(),
            default_paragraph_style_id: None,
            style_paragraph_backgrounds: HashMap::new(),
            style_word_wraps: HashMap::new(),
            pair_kerning: PairKerningRules::default(),
            default_language: None,
        },
    }
}

impl Parser for DocxParser {
    fn parse(
        &self,
        data: &[u8],
        _options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let default_tab_stop_pt: Option<f64> = extract_default_tab_stop_pt(data);
        let word_compatibility_mode: WordCompatibilityMode = extract_compatibility_mode(data);
        let hyphenates_automatically: bool = extract_auto_hyphenation(data);
        let ZipPreParseAssets {
            metadata,
            mut ctx,
            mut math,
            mut chart_ctx,
            column_layouts,
            page_numbering,
            metafile_images,
            theme_fonts,
            default_paragraph_style_id,
            style_paragraph_backgrounds,
            style_word_wraps,
            pair_kerning,
            default_language,
        } = build_zip_preparse_assets(data);

        let docx = docx_rs::read_docx(data).map_err(|e| {
            crate::parser::parse_err(format!("Failed to parse DOCX (docx-rs): {e}"))
        })?;

        // Populate locale-specific footnote/endnote style IDs from docx styles
        ctx.notes.populate_style_ids(&docx.styles);

        let mut images = build_image_map(&docx);
        images.extend(metafile_images);
        let hyperlinks = build_hyperlink_map(&docx);
        let numberings = build_numbering_map(&docx.numberings);
        let style_map = build_style_map(
            &docx.styles,
            &theme_fonts,
            default_paragraph_style_id.as_deref(),
            &style_paragraph_backgrounds,
            &style_word_wraps,
            &pair_kerning,
        );

        let header_footer_styles = HeaderFooterStyleContext {
            style_map: &style_map,
            theme_fonts: &theme_fonts,
            paragraph_property_defaults_are_declared: ctx.paragraph_property_defaults_are_declared,
        };

        // The header and footer parts are converted only now: their paragraphs
        // resolve `w:spacing w:after` through the same style cascade the body
        // takes, and that map needs the docx-rs parse above (issue #1195). The
        // archive is opened a second time for it rather than kept alive across
        // the parse; the parts themselves are still read exactly once.
        let header_footer_assets: HeaderFooterAssets = crate::parser::open_zip(data)
            .map(|mut archive| build_header_footer_assets(&mut archive, header_footer_styles))
            .unwrap_or_default();
        let mut warnings: Vec<ConvertWarning> = Vec::new();

        let mut elements: Vec<TaggedElement> = Vec::new();
        let mut pages: Vec<Page> = Vec::new();
        let mut section_layout_index: usize = 0;
        for (idx, child) in docx.document.children.iter().enumerate() {
            // A removed paragraph mark merges into the next paragraph; zero-width
            // bookmark/comment markers leave it pending. Tables, TOCs, and
            // section boundaries start separate flow blocks, so flush it here.
            if matches!(
                child,
                docx_rs::DocumentChild::Table(_)
                    | docx_rs::DocumentChild::TableOfContents(_)
                    | docx_rs::DocumentChild::Section(_)
            ) && let Some(block) = withheld_paragraph_block(&ctx)
            {
                elements.push(TaggedElement::Plain(vec![block]));
            }
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| match child {
                docx_rs::DocumentChild::Paragraph(para)
                    if is_empty_section_break_paragraph(para) =>
                {
                    // Paragraph cursors scan raw XML, so even this marker must
                    // advance them before its layout-only block is discarded.
                    let _ = convert_paragraph_element(
                        para,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                        &docx.styles,
                    );
                    vec![]
                }
                docx_rs::DocumentChild::Paragraph(para) => {
                    let mut tagged: Vec<TaggedElement> = convert_paragraph_element(
                        para,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                        &docx.styles,
                    );
                    // Inject math equations for this body child
                    let eqs = math.take(idx);
                    for eq in eqs {
                        tagged.push(TaggedElement::Plain(vec![Block::MathEquation(eq)]));
                    }
                    // Inject charts for this body child
                    let chs = chart_ctx.take(idx);
                    for ch in chs {
                        tagged.push(TaggedElement::Plain(vec![Block::Chart(Box::new(ch))]));
                    }
                    tagged
                }
                docx_rs::DocumentChild::Table(table) => {
                    vec![TaggedElement::Plain(vec![convert_body_table_block(
                        table,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                    )])]
                }
                docx_rs::DocumentChild::StructuredDataTag(sdt) => {
                    convert_sdt_children(sdt, &images, &hyperlinks, &style_map, &ctx, &docx.styles)
                }
                docx_rs::DocumentChild::BookmarkStart(_)
                | docx_rs::DocumentChild::BookmarkEnd(_)
                | docx_rs::DocumentChild::CommentStart(_)
                | docx_rs::DocumentChild::CommentEnd(_) => vec![],
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
                // The section closes here, so a paragraph still withheld has no
                // later paragraph in this flow to merge into (issue #1710).
                if let Some(block) = withheld_paragraph_block(&ctx) {
                    elements.push(TaggedElement::Plain(vec![block]));
                }
                let column_layout = match column_layouts.get(section_layout_index) {
                    Some(layout) => layout.clone(),
                    None => extract_column_layout_from_section_property(section_prop),
                };
                let previous_flow_page = pages.last().and_then(|previous| match previous {
                    Page::Flow(flow_page) | Page::FlowContinuous(flow_page) => Some(flow_page),
                    _ => None,
                });
                let flow_page: crate::ir::FlowPage = build_flow_page_from_section(
                    section_prop,
                    std::mem::take(&mut elements),
                    &numberings,
                    SectionHeaderFooterInputs {
                        assets: &header_footer_assets,
                        inherited_header: previous_flow_page.and_then(|page| page.header.as_ref()),
                        inherited_footer: previous_flow_page.and_then(|page| page.footer.as_ref()),
                    },
                    SectionOverrides {
                        column_layout,
                        page_numbering: page_numbering.get(section_layout_index).copied().flatten(),
                    },
                    header_footer_styles,
                    &mut warnings,
                );
                pages.push(
                    if section_prop.section_type == Some(docx_rs::SectionType::Continuous) {
                        Page::FlowContinuous(flow_page)
                    } else {
                        Page::Flow(flow_page)
                    },
                );
                section_layout_index += 1;
            }
        }

        // The body ends, so nothing is left to merge a withheld paragraph into.
        if let Some(block) = withheld_paragraph_block(&ctx) {
            elements.push(TaggedElement::Plain(vec![block]));
        }

        let final_column_layout = match column_layouts.get(section_layout_index) {
            Some(layout) => layout.clone(),
            None => extract_column_layout_from_section_property(&docx.document.section_property),
        };
        let previous_flow_page = pages.last().and_then(|previous| match previous {
            Page::Flow(flow_page) | Page::FlowContinuous(flow_page) => Some(flow_page),
            _ => None,
        });
        let final_flow_page: crate::ir::FlowPage = build_flow_page_from_section(
            &docx.document.section_property,
            elements,
            &numberings,
            SectionHeaderFooterInputs {
                assets: &header_footer_assets,
                inherited_header: previous_flow_page.and_then(|page| page.header.as_ref()),
                inherited_footer: previous_flow_page.and_then(|page| page.footer.as_ref()),
            },
            SectionOverrides {
                column_layout: final_column_layout,
                page_numbering: page_numbering.get(section_layout_index).copied().flatten(),
            },
            header_footer_styles,
            &mut warnings,
        );
        pages.push(
            if docx.document.section_property.section_type == Some(docx_rs::SectionType::Continuous)
            {
                Page::FlowContinuous(final_flow_page)
            } else {
                Page::Flow(final_flow_page)
            },
        );

        Ok((
            Document {
                metadata,
                pages,
                styles: StyleSheet {
                    default_tab_stop_pt,
                    default_text: Some(resolve_doc_default_text_style(
                        &docx.styles,
                        &theme_fonts,
                        &pair_kerning,
                    )),
                    word_compatibility_mode: Some(word_compatibility_mode),
                    default_language,
                    hyphenates_automatically,
                    ..StyleSheet::default()
                },
            },
            warnings,
        ))
    }
}

/// `w:defaultTabStop w:val` from `word/settings.xml`, in points. Read from
/// the raw part because docx-rs substitutes its own default when the
/// element is absent, erasing the absent-vs-explicit distinction the
/// East Asian fallback depends on (issue #393).
fn extract_default_tab_stop_pt(data: &[u8]) -> Option<f64> {
    let mut archive = crate::parser::open_zip(data).ok()?;
    let settings_xml: String = read_zip_text(&mut archive, "word/settings.xml")?;
    let element_start: usize = settings_xml.find("<w:defaultTabStop")?;
    let rest: &str = &settings_xml[element_start..];
    let value_start: usize = rest.find("w:val=\"")? + 7;
    let value_end: usize = rest[value_start..].find('"')? + value_start;
    let twips: f64 = rest[value_start..value_end].parse().ok()?;
    (twips > 0.0).then_some(twips / 20.0)
}

/// The layout engine Word lays this package out with, from the
/// `compatibilityMode` compatibility setting in `word/settings.xml`.
///
/// Read from the raw part for the same reason `w:defaultTabStop` is: docx-rs
/// models `w:compat` only as the flags it writes itself, so the setting cannot
/// be reached through its document tree.
///
/// A package that carries no setting — no `word/settings.xml` at all, or one
/// with no `compatibilityMode` — is a pre-2013 document to Word, so absence is
/// [`WordCompatibilityMode::Legacy`] rather than an unknown. The attribute
/// order is not fixed: native Word writes `w:val` first, docx-rs writes it
/// last.
fn extract_compatibility_mode(data: &[u8]) -> WordCompatibilityMode {
    let Some(mode) = declared_compatibility_mode(data) else {
        return WordCompatibilityMode::Legacy;
    };
    if mode >= 15 {
        WordCompatibilityMode::Word2013OrLater
    } else {
        WordCompatibilityMode::Legacy
    }
}

/// Whether `word/settings.xml` turns automatic hyphenation on. Read from the
/// raw part because docx-rs does not parse `w:autoHyphenation`.
///
/// TODO(docx-rs parses no hyphenation controls): the paragraph-level
/// `w:suppressAutoHyphens` and the document's `w:doNotHyphenateCaps`,
/// `w:hyphenationZone` and `w:consecutiveHyphenLimit` are not honoured yet, so
/// a document that turns hyphenation on is hyphenated wherever Typst can.
fn extract_auto_hyphenation(data: &[u8]) -> bool {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return false;
    };
    let Some(settings_xml) = read_zip_text(&mut archive, "word/settings.xml") else {
        return false;
    };
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;

    let is_word_namespace = |namespace: ResolveResult<'_>| {
        matches!(namespace, ResolveResult::Bound(namespace) if matches!(
            namespace.as_ref(),
            b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                | b"http://purl.oclc.org/ooxml/wordprocessingml/main"
        ))
    };
    let mut reader = quick_xml::reader::NsReader::from_str(&settings_xml);
    let mut depth: usize = 0;
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let has_children: bool = matches!(event, Event::Start(_));
                let (Event::Start(element) | Event::Empty(element)) = event else {
                    unreachable!("matched above");
                };
                let (namespace, name) = reader.resolver().resolve_element(element.name());
                let is_word_element: bool = is_word_namespace(namespace);
                if depth == 0 && (!is_word_element || name.as_ref() != b"settings") {
                    return false;
                }
                // A setting is a direct child, never text in a comment or
                // an identically named element in an extension subtree.
                if depth == 1 && is_word_element && name.as_ref() == b"autoHyphenation" {
                    let value = element.attributes().flatten().find_map(|attribute| {
                        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
                        (is_word_namespace(namespace) && name.as_ref() == b"val")
                            .then(|| {
                                attribute
                                    .normalized_value(OOXML_XML_VERSION)
                                    .ok()
                                    .map(|value| value.into_owned())
                            })
                            .flatten()
                    });
                    let enabled: bool = value
                        .as_deref()
                        .is_none_or(|value| matches!(value, "1" | "true" | "on"));
                    tracing::debug!(enabled, "Read DOCX automatic hyphenation setting");
                    return enabled;
                }
                if has_children {
                    depth += 1;
                }
            }
            Ok(Event::End(_)) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return false;
                }
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
    }
}

fn declared_compatibility_mode(data: &[u8]) -> Option<u32> {
    let mut archive = crate::parser::open_zip(data).ok()?;
    let settings_xml: String = read_zip_text(&mut archive, "word/settings.xml")?;
    settings_xml
        .split("<w:compatSetting")
        .skip(1)
        // Split on the tag's own end so a setting written open rather than
        // self-closing cannot swallow the ones after it.
        .filter_map(|element| element.split_once('>').map(|(attributes, _)| attributes))
        .find(|attributes| attributes.contains(r#"w:name="compatibilityMode""#))
        .and_then(|attributes| {
            let value_start: usize = attributes.find(r#"w:val=""#)? + r#"w:val=""#.len();
            let rest: &str = &attributes[value_start..];
            let value_end: usize = rest.find('"')?;
            rest[..value_end].parse().ok()
        })
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
    styles: &docx_rs::Styles,
) -> Vec<TaggedElement> {
    let mut result = Vec::new();
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                result.extend(convert_paragraph_element(
                    para, images, hyperlinks, style_map, ctx, styles,
                ));
            }
            docx_rs::StructuredDataTagChild::Table(table) => {
                result.push(TaggedElement::Plain(vec![convert_body_table_block(
                    table, images, hyperlinks, style_map, ctx,
                )]));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                result.extend(convert_sdt_children(
                    nested, images, hyperlinks, style_map, ctx, styles,
                ));
            }
            _ => {}
        }
    }
    result
}

fn convert_body_table_block(
    source_table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> Block {
    let table: Table = convert_table(source_table, images, hyperlinks, style_map, ctx, 0);
    match sections::floating_table_frame(source_table, &table) {
        Some(frame) => Block::FloatingTable(Box::new(FloatingTable { table, frame })),
        None => Block::Table(table),
    }
}

/// Convert one docx-rs paragraph into its ordered plain blocks and list paragraph.
/// A numbered paragraph can produce multiple elements when breaks or images surround its text.
fn convert_paragraph_element(
    para: &docx_rs::Paragraph,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    styles: &docx_rs::Styles,
) -> Vec<TaggedElement> {
    let num_info = extract_num_info(para, styles);

    // Build the paragraph IR
    let mut blocks = Vec::new();
    convert_paragraph_blocks(para, &mut blocks, images, hyperlinks, style_map, ctx, None);

    // A paragraph whose mark a tracked deletion or move removed is held back
    // for the paragraph it merges into, so it contributes no element of its own
    // here. An empty one would do: `group_into_lists` ends the list it is
    // grouping at any element, so the items after the origin would start a
    // second list (issue #1710).
    if ctx.paragraph_marks.is_withholding() {
        return if blocks.is_empty() {
            Vec::new()
        } else {
            vec![TaggedElement::Plain(blocks)]
        };
    }

    let Some(info) = num_info else {
        return vec![TaggedElement::Plain(blocks)];
    };

    let has_text_fragment: bool = blocks
        .iter()
        .any(|block| matches!(block, Block::Paragraph(paragraph) if !paragraph.runs.is_empty()));
    let mut tagged: Vec<TaggedElement> = Vec::new();
    let mut plain_blocks: Vec<Block> = Vec::new();
    let mut list_paragraph_emitted: bool = false;
    for block in blocks {
        let paragraph: Option<Paragraph> = match block {
            Block::Paragraph(paragraph)
                if !list_paragraph_emitted
                    && (!paragraph.runs.is_empty() || !has_text_fragment) =>
            {
                Some(paragraph)
            }
            other => {
                plain_blocks.push(other);
                None
            }
        };
        let Some(mut paragraph) = paragraph else {
            continue;
        };

        if !plain_blocks.is_empty() {
            tagged.push(TaggedElement::Plain(std::mem::take(&mut plain_blocks)));
        }
        apply_word_compatible_paragraph_defaults(
            &mut paragraph.style,
            ctx.paragraph_property_defaults_are_declared,
        );
        tagged.push(TaggedElement::ListParagraph {
            info: info.clone(),
            paragraph: Box::new(paragraph),
        });
        list_paragraph_emitted = true;
    }
    if !plain_blocks.is_empty() {
        tagged.push(TaggedElement::Plain(plain_blocks));
    }
    tagged
}

/// Build a text `Run` from extracted text, merging explicit run styling with the
/// resolved paragraph style. Returns `None` when the text is empty, so callers
/// can skip empty runs without duplicating the emptiness check.
fn build_text_run(
    text: String,
    run_property: &docx_rs::RunProperty,
    is_small_caps: bool,
    resolved_style: Option<&ResolvedStyle>,
    style_map: &StyleMap,
    theme_fonts: &ThemeFonts,
    href: Option<String>,
) -> Option<Run> {
    if text.is_empty() {
        return None;
    }
    Some(Run {
        text,
        style: resolve_run_style(
            run_property,
            is_small_caps,
            resolved_style,
            style_map,
            theme_fonts,
        ),
        href,
        footnote: None,
        inline_box: None,
    })
}

/// A run's formatting with everything it inherits already folded in: the
/// referenced character style beneath its explicit properties, then the
/// paragraph style beneath both.
///
/// Header and footer runs resolve through here too, so a `w:pStyle`'s `w:sz`
/// reaches a running head exactly as it reaches body copy (issue #1822).
fn resolve_run_style(
    run_property: &docx_rs::RunProperty,
    is_small_caps: bool,
    resolved_style: Option<&ResolvedStyle>,
    style_map: &StyleMap,
    theme_fonts: &ThemeFonts,
) -> TextStyle {
    let mut explicit_style: TextStyle = extract_run_style(run_property);
    if explicit_style.font_family.is_none()
        && let Ok(run_property_json) = serde_json::to_value(run_property)
    {
        explicit_style.font_family =
            resolve_latin_theme_font_family(&run_property_json, theme_fonts);
    }
    if is_small_caps {
        explicit_style.small_caps = Some(true);
    }
    // Layer the referenced character style (`<w:rStyle>`, e.g. a syntax
    // highlighting token) beneath the run's explicit properties so its color
    // and weight apply while explicit run formatting still wins (issue #176).
    if let Some(char_style) = extract_run_style_id(run_property).and_then(|id| style_map.get(&id)) {
        let mut combined: TextStyle = char_style.text.clone();
        combined.merge_from(&explicit_style);
        explicit_style = combined;
    }
    merge_text_style(&explicit_style, resolved_style)
}

/// Intermediate results from scanning a run segment for media and text boxes.
struct RunChildrenMedia {
    text_box_blocks: Vec<Block>,
    /// The `wp:inline` text boxes this run anchors. They belong on the anchor
    /// paragraph's own line, so they ride the run rather than becoming flow
    /// blocks beneath it (issue #1690).
    inline_text_boxes: Vec<InlineTextBox>,
}

/// The ordered pieces of a run around structural layout breaks.
enum RunLayoutPart {
    Content(Box<docx_rs::Run>),
    PageBreak,
    ColumnBreak,
}

/// Split only at page and column breaks so all content on either side keeps the
/// run's properties and is emitted on the page or column where Word places it.
fn split_run_at_layout_breaks(run: &docx_rs::Run) -> Vec<RunLayoutPart> {
    let mut parts: Vec<RunLayoutPart> = Vec::new();
    let mut content_children: Vec<docx_rs::RunChild> = Vec::new();

    for child in &run.children {
        let layout_break: Option<RunLayoutPart> = match child {
            docx_rs::RunChild::Break(br) if is_page_break(br) => Some(RunLayoutPart::PageBreak),
            docx_rs::RunChild::Break(br) if is_column_break(br) => Some(RunLayoutPart::ColumnBreak),
            _ => None,
        };

        if let Some(layout_break) = layout_break {
            if !content_children.is_empty() {
                let mut content_run: docx_rs::Run = run.clone();
                content_run.children = std::mem::take(&mut content_children);
                parts.push(RunLayoutPart::Content(Box::new(content_run)));
            }
            parts.push(layout_break);
        } else {
            content_children.push(child.clone());
        }
    }

    if !content_children.is_empty() {
        let mut content_run: docx_rs::Run = run.clone();
        content_run.children = content_children;
        parts.push(RunLayoutPart::Content(Box::new(content_run)));
    }

    parts
}

/// A blank paragraph fragment still paints these paragraph-level decorations
/// before a leading page or column break.
fn paragraph_has_visible_background(
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    flow: ParagraphFlow<'_>,
) -> bool {
    let explicit_style: ParagraphStyle = extract_paragraph_style(&para.property);
    flow.background.is_some()
        || explicit_style.background.is_some()
        || resolved_style.is_some_and(|style| style.paragraph.background.is_some())
}

/// Scan a run segment's children for drawings and VML shapes.
/// Extracted images are pushed to `inline_images`; text boxes are returned in
/// `RunChildrenMedia`.
fn extract_run_children_media(
    run: &docx_rs::Run,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    inline_images: &mut Vec<Block>,
) -> RunChildrenMedia {
    let mut text_box_blocks: Vec<Block> = Vec::new();
    let mut inline_text_boxes: Vec<InlineTextBox> = Vec::new();

    for run_child in &run.children {
        if let docx_rs::RunChild::Drawing(drawing) = run_child {
            let wpg_drawing: Option<WpgDrawingInfo> = ctx.drawing_shapes.consume_wpg_drawing();
            let canvas_image_offset: Option<(f64, f64)> =
                ctx.drawing_shapes.consume_canvas_image_offset();
            if let Some(wpg_drawing) = wpg_drawing {
                // docx-rs represents only one child from a WPG group. Use the
                // complete raw-XML group instead to avoid dropping its siblings.
                text_box_blocks.extend(convert_wpg_drawing_blocks(
                    wpg_drawing,
                    images,
                    hyperlinks,
                    style_map,
                    ctx,
                ));
            } else {
                if let Some(img_block) =
                    extract_drawing_image(drawing, images, &ctx.wraps, canvas_image_offset)
                {
                    inline_images.push(img_block);
                }
                match extract_drawing_text_box_blocks(drawing, images, hyperlinks, style_map, ctx) {
                    DrawingTextBoxPlacement::Absent => {}
                    DrawingTextBoxPlacement::Floating(blocks) => {
                        text_box_blocks.extend(blocks);
                    }
                    DrawingTextBoxPlacement::Inline(inline_box) => {
                        inline_text_boxes.push(*inline_box);
                    }
                }
                if drawing.data.is_none()
                    && let Some(shape) = ctx.drawing_shapes.consume_next()
                {
                    // docx-rs leaves geometry-only `wps:wsp` drawings unclassified.
                    text_box_blocks.push(Block::FloatingShape(shape));
                }
            }
        }
        if let docx_rs::RunChild::Shape(shape) = run_child {
            let vml_text_box: VmlTextBoxInfo = ctx.vml_text_boxes.consume_next();
            if let Some(floating_text_box) = extract_vml_shape_text_box(shape, &vml_text_box) {
                text_box_blocks.push(Block::FloatingTextBox(floating_text_box));
            } else {
                text_box_blocks.extend(vml_text_box.into_blocks());
            }

            if let Some(img_block) = extract_shape_image(shape, images) {
                inline_images.push(img_block);
            }
        }
    }

    RunChildrenMedia {
        text_box_blocks,
        inline_text_boxes,
    }
}

fn convert_wpg_drawing_blocks(
    drawing: WpgDrawingInfo,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> Vec<Block> {
    let mut result: Vec<Block> = Vec::new();
    for child in drawing.children {
        if let Some(shape) = child.shape {
            result.push(Block::FloatingShape(shape));
        }

        let mut content: Vec<Block> = Vec::new();
        for document_child in &child.content {
            match document_child {
                // A shape's text frame is its own flow, not the cell's, even
                // when the shape is anchored inside one.
                docx_rs::DocumentChild::Paragraph(paragraph) => convert_paragraph_blocks(
                    paragraph,
                    &mut content,
                    images,
                    hyperlinks,
                    style_map,
                    ctx,
                    None,
                ),
                docx_rs::DocumentChild::Table(table) => content.push(Block::Table(convert_table(
                    table, images, hyperlinks, style_map, ctx, 0,
                ))),
                _ => {}
            }
        }
        // The shape's flow ends with its own content (issue #1710).
        content.extend(withheld_paragraph_block(ctx));
        if let Some(text_color) = child.text_color {
            apply_default_text_color(&mut content, text_color);
        }
        if !content.is_empty() {
            result.push(Block::FloatingTextBox(FloatingTextBox {
                content,
                wrap_mode: child.wrap_mode,
                width: child.width,
                height: child.height,
                fill: None,
                stroke: None,
                shape_rotation_deg: child.rotation_deg,
                padding: child.padding,
                vertical_align: child.vertical_align,
                vertical_anchor: crate::ir::FrameAnchor::Text,
                vertical_position_align: None,
                horizontal_align: None,
                horizontal_anchor: crate::ir::FrameAnchor::Text,
                offset_x: child.offset_x,
                offset_y: child.offset_y,
            }));
        }
    }
    result
}

fn apply_default_text_color(blocks: &mut [Block], color: Color) {
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => {
                for run in &mut paragraph.runs {
                    run.style.color.get_or_insert(color);
                }
            }
            Block::List(list) => {
                for item in &mut list.items {
                    for paragraph in &mut item.content {
                        for run in &mut paragraph.runs {
                            run.style.color.get_or_insert(color);
                        }
                    }
                }
            }
            Block::Table(table) => {
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        apply_default_text_color(&mut cell.content, color);
                    }
                }
            }
            Block::FloatingTextBox(text_box) => {
                apply_default_text_color(&mut text_box.content, color);
            }
            _ => {}
        }
    }
}

/// The list a paragraph's `TOC` field produces, if it carries one.
///
/// A dirty `TOC` field is stored as its instruction and nothing else, so the
/// paragraph holding it has no text to render and the contents page came out
/// blank. The field becomes a block the renderer resolves against the
/// document itself instead — `\o` against its headings, `\a` against the
/// captions of one `SEQ` sequence (issue #576).
fn toc_field(para: &docx_rs::Paragraph) -> Option<TableOfContents> {
    para.children
        .iter()
        .filter_map(|child| match child {
            docx_rs::ParagraphChild::Run(run) => Some(run),
            _ => None,
        })
        .flat_map(|run| run.children.iter())
        .find_map(|child| {
            let instruction: &str = match child {
                docx_rs::RunChild::InstrText(instruction) => match instruction.as_ref() {
                    docx_rs::InstrText::Unsupported(text) => text,
                    _ => return None,
                },
                docx_rs::RunChild::InstrTextString(text) => text,
                _ => return None,
            };
            toc_caption_identifier(instruction)
                .map(|identifier| TableOfContents::Captions { identifier })
                .or_else(|| {
                    toc_heading_depth(instruction).map(|depth| TableOfContents::Headings { depth })
                })
        })
}

/// The number a run's `SEQ` field renders, if it carries one.
///
/// Word stores a caption number in the field, not in the text, so a run that
/// holds `SEQ Table` contributes the counter's next value. Text between the
/// field's `separate` and `end` is its cached result — what Word last
/// computed — and is replaced by the value computed here rather than added to
/// it (issue #577).
fn seq_field_text(
    run: &docx_rs::Run,
    fields: &FieldContext,
    seen: &mut Option<String>,
) -> Option<String> {
    let mut identifier: Option<String> = None;
    for child in &run.children {
        match child {
            docx_rs::RunChild::InstrText(instruction) => {
                if let docx_rs::InstrText::Unsupported(text) = instruction.as_ref()
                    && let Some(found) = seq_identifier(text)
                {
                    identifier = Some(found.to_string());
                }
            }
            docx_rs::RunChild::InstrTextString(text) => {
                if let Some(found) = seq_identifier(text) {
                    identifier = Some(found.to_string());
                }
            }
            _ => {}
        }
    }
    identifier.map(|identifier| {
        let number = fields.next_in_sequence(&identifier).to_string();
        *seen = Some(identifier);
        number
    })
}

/// Resolve a note's runs against the style it names.
///
/// A note is read from `footnotes.xml` before the stylesheet is, so its runs
/// arrive carrying only their own `w:rPr`. Word resolves them through the same
/// cascade as the body: the note's `w:pStyle` — `FootnoteText` and friends —
/// supplies the size, colour, and family the runs leave unstated, and falls
/// back to the document defaults when the note names no style (issue #580).
fn resolve_note_runs(content: &NoteContent, style_map: &StyleMap) -> Vec<Run> {
    let note_style = content
        .style_id
        .as_deref()
        .and_then(|style_id| style_map.get(style_id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID));

    content
        .runs
        .iter()
        .map(|note_run| Run {
            text: note_run.text.clone(),
            style: merge_text_style(&note_run.explicit, note_style),
            href: None,
            footnote: None,
            inline_box: None,
        })
        .collect()
}

/// A paragraph child once tracked changes have been resolved away.
///
/// Callers match only the variants they render; a header paragraph ignores
/// `Hyperlink`, and the body ignores the two field variants a header uses.
pub(super) enum ParagraphItem<'a> {
    Run(&'a docx_rs::Run),
    Hyperlink(&'a docx_rs::Hyperlink),
    PageNum,
    NumPages,
}

/// Resolve a paragraph's tracked changes to the final document.
///
/// Word shows two views of a document with change tracking on. The review
/// view marks up both sides; the final view — what "No Markup" shows, what
/// accepting every revision produces, and what a converter is expected to
/// render — keeps the insertions and drops the deletions.
///
/// `w:ins` and `w:del` were both falling through the paragraph child match's
/// catch-all arm, so both sides vanished. Dropping `w:del` is right; dropping
/// `w:ins` silently lost ordinary document text whose only distinction was
/// having been typed while tracking was on (issue #583).
///
/// A `w:del` nested inside a `w:ins` is text that was inserted and then
/// deleted again, so it is absent from the final document too and is dropped
/// with the rest.
///
/// A tracked move is both kinds at once: its destination (`w:moveTo`) is
/// kept like an insertion and its origin (`w:moveFrom`) is dropped like a
/// deletion.
pub(super) fn flatten_tracked_changes(
    children: &[docx_rs::ParagraphChild],
) -> Vec<ParagraphItem<'_>> {
    let mut items: Vec<ParagraphItem<'_>> = Vec::with_capacity(children.len());
    for child in children {
        match child {
            docx_rs::ParagraphChild::Run(run) => items.push(ParagraphItem::Run(run)),
            docx_rs::ParagraphChild::Hyperlink(hyperlink) => {
                items.push(ParagraphItem::Hyperlink(hyperlink))
            }
            docx_rs::ParagraphChild::PageNum(_) => items.push(ParagraphItem::PageNum),
            docx_rs::ParagraphChild::NumPages(_) => items.push(ParagraphItem::NumPages),
            docx_rs::ParagraphChild::Insert(insert) => {
                for inserted in &insert.children {
                    if let docx_rs::InsertChild::Run(run) = inserted {
                        items.push(ParagraphItem::Run(run));
                    }
                }
            }
            docx_rs::ParagraphChild::MoveTo(move_to) => {
                for moved in &move_to.children {
                    if let docx_rs::MoveToChild::Run(run) = moved {
                        items.push(ParagraphItem::Run(run));
                    }
                }
            }
            docx_rs::ParagraphChild::MoveFrom(_) => {}
            _ => {}
        }
    }
    items
}

/// Process hyperlink children, extracting text runs with the resolved URL.
fn process_hyperlink_runs(
    hyperlink: &docx_rs::Hyperlink,
    hyperlinks: &HyperlinkMap,
    resolved_style: Option<&ResolvedStyle>,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    runs: &mut Vec<Run>,
) {
    let href: Option<String> = resolve_hyperlink_url(hyperlink, hyperlinks);
    for hchild in &hyperlink.children {
        if let docx_rs::ParagraphChild::Run(run) = hchild {
            let hl_small_caps: bool = ctx.small_caps.next_is_small_caps();
            let text: String = extract_run_text(run);
            if let Some(ir_run) = build_text_run(
                text,
                &run.run_property,
                hl_small_caps,
                resolved_style,
                style_map,
                &ctx.theme_fonts,
                href.clone(),
            ) {
                runs.push(ir_run);
            }
        }
    }
}

/// Context shared by the output blocks generated from one source `<w:p>`:
/// document theme fonts used by run and paragraph-mark resolution, plus its
/// inherited direction, shading, wrapping, contextual spacing, and style state.
///
/// Resolved once per `<w:p>` because the paragraph cursors (bidi, shading,
/// `w:wordWrap`, `w:contextualSpacing`) advance on read, then reused by every
/// output paragraph the source `<w:p>` splits into.
#[derive(Clone, Copy)]
struct ParagraphFlow<'a> {
    theme_fonts: &'a ThemeFonts,
    is_rtl: bool,
    background: Option<Color>,
    /// The paragraph's own `w:wordWrap`, recovered from the raw XML — the
    /// published docx-rs does not parse it (issue #1041).
    word_wrap: Option<bool>,
    /// What `w:contextualSpacing` takes from this paragraph's `w:spacing`
    /// gaps (issue #1684).
    contextual_spacing: ParagraphContextualSpacing<'a>,
    /// Whether the style this paragraph takes its formatting from is
    /// explicitly defined in `word/styles.xml` — a resolvable `w:pStyle`, or
    /// the document's own default-style definition for a bare paragraph.
    /// False means the paragraph falls through to Word's built-in Normal,
    /// which is what suppresses the East Asian auto space (issue #732) and
    /// what breaks Hangul lines at character level rather than keeping each
    /// eojeol whole (issue #833).
    effective_style_is_defined: bool,
    /// Whether `word/styles.xml` declares `w:docDefaults/w:pPrDefault`, which
    /// decides the `w:spacing w:after` an unstated gap falls back to
    /// (issue #1085).
    paragraph_property_defaults_are_declared: bool,
    /// The active table-style paragraph after spacing for this cell, before
    /// document defaults are applied.
    table_style_space_after: Option<f64>,
}

/// Convert a docx-rs Paragraph to IR blocks, handling page breaks and inline images.
/// If the paragraph has `page_break_before`, a `Block::PageBreak` is emitted first.
/// Consecutive inline images within a paragraph are kept in one wrapping flow container.
/// Style formatting from the document's style definitions is merged with explicit formatting.
fn convert_paragraph_blocks(
    para: &docx_rs::Paragraph,
    out: &mut Vec<Block>,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    table_style_space_after: Option<f64>,
) {
    let paragraph_block_start: usize = out.len();
    // Every paragraph cursor advances here, exactly once per XML <w:p>, and
    // each checks this paragraph's `w:pStyle` against the one its scan
    // recorded at the same index (issue #1689).
    let style_id: Option<&str> = get_paragraph_style_id(&para.property);
    let flow = ParagraphFlow {
        theme_fonts: &ctx.theme_fonts,
        is_rtl: ctx.bidi.next_is_bidi(style_id),
        background: ctx.paragraph_shading.next_background(style_id),
        word_wrap: ctx.word_wraps.next_word_wrap(style_id),
        contextual_spacing: ctx.contextual_spacing.next_paragraph(style_id),
        // A `w:pStyle` naming a style the document never defines falls back
        // to the default style, the same as carrying no `w:pStyle` at all.
        effective_style_is_defined: match style_id {
            Some(id) if style_map.contains_key(id) => true,
            _ => ctx.default_paragraph_style_is_defined,
        },
        paragraph_property_defaults_are_declared: ctx.paragraph_property_defaults_are_declared,
        table_style_space_after,
    };
    // A paragraph mark a tracked deletion or move removed is not in the final
    // document, so this paragraph has no break of its own: it merges into the
    // paragraph after it (issue #1710).
    let mark_is_removed: bool = ctx.paragraph_marks.next_mark_is_removed(style_id);

    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    // A dirty `TOC` field is stored as its instruction and nothing else, so
    // the paragraph carrying it has no text of its own to render (issue #576).
    // A field Word has already computed keeps its cached entries instead:
    // those are the result, and recomputing over them would drop the numbers
    // the document shipped.
    if let Some(contents) = toc_field(para)
        && para
            .children
            .iter()
            .filter_map(|child| match child {
                docx_rs::ParagraphChild::Run(run) => Some(run),
                _ => None,
            })
            .all(|run| extract_run_text(run).trim().is_empty())
    {
        out.push(Block::TableOfContents(contents));
        return;
    }

    // Look up the paragraph's referenced style
    let resolved_style = get_paragraph_style_id(&para.property)
        .and_then(|id| style_map.get(id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID));

    // Collect text runs and detect inline images. A paragraph whose
    // predecessor's mark was removed opens with that paragraph's surviving
    // runs, because the two are one paragraph in the final view (issue #1710).
    let mut runs: Vec<Run> = ctx.paragraph_marks.take_merged_runs();
    let mut inline_images: Vec<Block> = Vec::new();
    let mut emitted_paragraph: bool = false;
    let mut emitted_media_blocks: bool = false;
    let mut emitted_floating_anchor: bool = false;
    let mut emitted_layout_break: bool = false;
    let mut empty_break_fragment_indices: Vec<usize> = Vec::new();
    let mut emitted_content_paragraph: bool = false;
    // Set by the run carrying a `SEQ` field, so the finished paragraph can be
    // wrapped as the caption a `TOC \a` list collects (issue #576).
    let mut caption_identifier: Option<String> = None;

    for child in flatten_tracked_changes(&para.children) {
        match child {
            ParagraphItem::Run(run) => {
                // Advance smallCaps cursor for every <w:r> in body
                let is_small_caps: bool = ctx.small_caps.next_is_small_caps();

                // Check for footnote/endnote reference runs
                if is_note_reference_run(run, &ctx.notes) {
                    if let Some(content) = ctx.notes.consume_next() {
                        runs.push(Run {
                            text: String::new(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: Some(resolve_note_runs(&content, style_map)),
                            inline_box: None,
                        });
                    }
                    continue;
                }

                for part in split_run_at_layout_breaks(run) {
                    match part {
                        break_part @ (RunLayoutPart::PageBreak | RunLayoutPart::ColumnBreak) => {
                            // Flush queued inline media even when no text run
                            // precedes the break; images are run children too.
                            push_inline_images(
                                out,
                                &mut inline_images,
                                paragraph_alignment(para),
                                paragraph_image_spacing(
                                    para,
                                    resolved_style,
                                    flow.contextual_spacing,
                                ),
                            );
                            if !runs.is_empty() {
                                clear_paragraph_fragment_borders(
                                    out,
                                    &empty_break_fragment_indices,
                                );
                                push_paragraph_from_runs(
                                    out,
                                    para,
                                    resolved_style,
                                    style_map,
                                    flow,
                                    &mut runs,
                                    caption_identifier.as_deref(),
                                );
                                emitted_paragraph = true;
                                emitted_content_paragraph = true;
                            } else if !emitted_content_paragraph
                                && !emitted_media_blocks
                                && paragraph_has_visible_background(para, resolved_style, flow)
                            {
                                // A break-only paragraph paints a line-sized shaded fragment
                                // before each break. Word leaves earlier fragments borderless
                                // and applies the paragraph border to the final fragment.
                                let is_continuation_fragment: bool =
                                    !empty_break_fragment_indices.is_empty();
                                if let Some(previous_fragment_index) =
                                    empty_break_fragment_indices.last().copied()
                                {
                                    clear_paragraph_fragment_border(out, previous_fragment_index);
                                }
                                let mut fragment: Block = build_paragraph_block(
                                    para,
                                    resolved_style,
                                    style_map,
                                    flow,
                                    &mut runs,
                                    caption_identifier.as_deref(),
                                );
                                let fragment_style: &mut ParagraphStyle = match &mut fragment {
                                    Block::Paragraph(paragraph) => &mut paragraph.style,
                                    Block::Caption(caption) => &mut caption.paragraph.style,
                                    _ => unreachable!("a paragraph must build a paragraph block"),
                                };
                                fragment_style.space_before = None;
                                fragment_style.decoration_before_spacing = None;
                                fragment_style.starts_after_layout_break = is_continuation_fragment;
                                out.push(fragment);
                                if !is_continuation_fragment {
                                    let fragment_index: usize = out.len() - 1;
                                    clear_paragraph_fragment_border(out, fragment_index);
                                }
                                empty_break_fragment_indices.push(out.len() - 1);
                                emitted_paragraph = true;
                            }
                            out.push(match break_part {
                                RunLayoutPart::PageBreak => Block::PageBreak,
                                RunLayoutPart::ColumnBreak => Block::ColumnBreak,
                                RunLayoutPart::Content(_) => unreachable!(),
                            });
                            emitted_layout_break = true;
                        }
                        RunLayoutPart::Content(run_segment) => {
                            let media = extract_run_children_media(
                                &run_segment,
                                images,
                                hyperlinks,
                                style_map,
                                ctx,
                                &mut inline_images,
                            );

                            // A picture is the paragraph's content, so its paragraph mark
                            // belongs to the picture rather than to a blank line. Counting
                            // only text boxes here left a picture-only paragraph emitting an
                            // empty paragraph as well, adding a full line box below every
                            // figure (issue #496).
                            emitted_media_blocks |= !inline_images.is_empty();

                            if !media.text_box_blocks.is_empty() {
                                emitted_media_blocks = true;
                                emitted_floating_anchor |=
                                    media.text_box_blocks.iter().any(|block| {
                                        matches!(
                                            block,
                                            Block::FloatingShape(_) | Block::FloatingTextBox(_)
                                        )
                                    });
                                if !runs.is_empty() {
                                    clear_paragraph_fragment_borders(
                                        out,
                                        &empty_break_fragment_indices,
                                    );
                                    push_inline_images(
                                        out,
                                        &mut inline_images,
                                        paragraph_alignment(para),
                                        paragraph_image_spacing(
                                            para,
                                            resolved_style,
                                            flow.contextual_spacing,
                                        ),
                                    );
                                    push_paragraph_from_runs(
                                        out,
                                        para,
                                        resolved_style,
                                        style_map,
                                        flow,
                                        &mut runs,
                                        caption_identifier.as_deref(),
                                    );
                                    emitted_paragraph = true;
                                    emitted_content_paragraph = true;
                                } else if !inline_images.is_empty() {
                                    push_inline_images(
                                        out,
                                        &mut inline_images,
                                        paragraph_alignment(para),
                                        paragraph_image_spacing(
                                            para,
                                            resolved_style,
                                            flow.contextual_spacing,
                                        ),
                                    );
                                }
                                out.extend(media.text_box_blocks);
                            }

                            let text: String =
                                seq_field_text(&run_segment, &ctx.fields, &mut caption_identifier)
                                    .unwrap_or_else(|| extract_run_text(&run_segment));
                            if let Some(ir_run) = build_text_run(
                                text,
                                &run_segment.run_property,
                                is_small_caps,
                                resolved_style,
                                style_map,
                                &ctx.theme_fonts,
                                None,
                            ) {
                                runs.push(ir_run);
                            }

                            // An inline text box sits on this paragraph's line, after
                            // whatever text its own `w:r` carried, so it joins the runs
                            // rather than becoming a flow block (issue #1690).
                            for inline_box in media.inline_text_boxes {
                                runs.push(Run {
                                    text: String::new(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                    inline_box: Some(Box::new(inline_box)),
                                });
                            }
                        }
                    }
                }
            }
            ParagraphItem::Hyperlink(hyperlink) => {
                process_hyperlink_runs(
                    hyperlink,
                    hyperlinks,
                    resolved_style,
                    style_map,
                    ctx,
                    &mut runs,
                );
            }
            // `w:pgNum`/`w:numPages` are header and footer fields; the body
            // resolves its page numbers through `w:fldSimple` instead.
            ParagraphItem::PageNum | ParagraphItem::NumPages => {}
        }
    }

    push_inline_images(
        out,
        &mut inline_images,
        paragraph_alignment(para),
        paragraph_image_spacing(para, resolved_style, flow.contextual_spacing),
    );

    if emitted_layout_break
        && runs.is_empty()
        && !emitted_content_paragraph
        && !emitted_media_blocks
        && paragraph_has_visible_background(para, resolved_style, flow)
        && matches!(
            out.get(paragraph_block_start..)
                .and_then(|blocks| blocks.last()),
            Some(Block::ColumnBreak)
        )
    {
        let mut continuation: Block = build_paragraph_block(
            para,
            resolved_style,
            style_map,
            flow,
            &mut runs,
            caption_identifier.as_deref(),
        );
        let continuation_style: &mut ParagraphStyle = match &mut continuation {
            Block::Paragraph(paragraph) => &mut paragraph.style,
            Block::Caption(caption) => &mut caption.paragraph.style,
            _ => unreachable!("a paragraph must build a paragraph block"),
        };
        continuation_style.starts_after_layout_break = true;
        out.push(continuation);
    }

    // A paragraph whose remaining content is just the mark left behind by a
    // page or column break is a break carrier: Word uses it only to force the
    // break, so it must not add a line box on the new page. An empty paragraph
    // with no break is a deliberate blank line and is still kept.
    let is_layout_break_carrier: bool = emitted_layout_break && runs.is_empty();

    if !is_layout_break_carrier
        && (!runs.is_empty()
            || !emitted_media_blocks
            || (emitted_floating_anchor && !emitted_paragraph))
    {
        if !runs.is_empty() {
            clear_paragraph_fragment_borders(out, &empty_break_fragment_indices);
        }
        // Keep paragraph marks for floating drawing anchors. The drawing itself
        // is positioned by offsets, but the source paragraph still contributes
        // to flow spacing between the drawing cluster and following content.
        let block: Block = build_paragraph_block(
            para,
            resolved_style,
            style_map,
            flow,
            &mut runs,
            caption_identifier.as_deref(),
        );
        match block {
            // The following paragraph's mark is the one that survives, so it
            // supplies the merged paragraph's style, its numbering and its
            // line — this one contributes only its runs (issue #1710).
            Block::Paragraph(paragraph) if mark_is_removed => {
                ctx.paragraph_marks.withhold(paragraph)
            }
            block => out.push(block),
        }
    }

    normalize_split_paragraph_styles(out, paragraph_block_start);
}

fn clear_paragraph_fragment_borders(blocks: &mut [Block], fragment_indices: &[usize]) {
    for fragment_index in fragment_indices {
        clear_paragraph_fragment_border(blocks, *fragment_index);
    }
}

fn clear_paragraph_fragment_border(blocks: &mut [Block], fragment_index: usize) {
    let style: &mut ParagraphStyle = match &mut blocks[fragment_index] {
        Block::Paragraph(paragraph) => &mut paragraph.style,
        Block::Caption(caption) => &mut caption.paragraph.style,
        _ => unreachable!("empty paragraph fragment indices point to paragraphs"),
    };
    style.border = None;
    style.border_space = None;
}

/// Keep paragraph-wide spacing and first-line indentation on the paragraph's
/// first text fragment. A hard layout break does not begin a new paragraph,
/// so its continuation must not repeat those edges or its first-line indent.
fn normalize_split_paragraph_styles(blocks: &mut [Block], start: usize) {
    let fragments: Vec<usize> = blocks[start..]
        .iter()
        .enumerate()
        .filter_map(|(offset, block)| {
            matches!(block, Block::Paragraph(_) | Block::Caption(_)).then_some(start + offset)
        })
        .collect();
    let content_fragments: Vec<usize> = fragments
        .iter()
        .copied()
        .filter(|index| match &blocks[*index] {
            Block::Paragraph(paragraph) => !paragraph.runs.is_empty(),
            Block::Caption(caption) => !caption.paragraph.runs.is_empty(),
            _ => false,
        })
        .collect();
    let Some(first_content) = content_fragments.first().copied() else {
        return;
    };
    let last_content: usize = *content_fragments.last().expect("first content exists");
    let blocks_before_first_content: &[Block] = &blocks[start..first_content];
    let starts_after_layout_break: bool = blocks_before_first_content
        .iter()
        .any(|block| matches!(block, Block::PageBreak | Block::ColumnBreak))
        && blocks_before_first_content.iter().all(|block| {
            is_empty_paragraph_block(block)
                || matches!(block, Block::PageBreak | Block::ColumnBreak)
        });

    for index in fragments {
        let (style, has_content): (&mut ParagraphStyle, bool) = match &mut blocks[index] {
            Block::Paragraph(paragraph) => (&mut paragraph.style, !paragraph.runs.is_empty()),
            Block::Caption(caption) => (
                &mut caption.paragraph.style,
                !caption.paragraph.runs.is_empty(),
            ),
            _ => unreachable!("fragment indices only include paragraphs and captions"),
        };
        if !has_content {
            style.space_before = None;
            style.space_after = None;
            style.indent_first_line = None;
            style.starts_after_layout_break = false;
            continue;
        }
        style.starts_after_layout_break = index == first_content && starts_after_layout_break;
        if index != first_content {
            style.space_before = None;
            style.indent_first_line = None;
        }
        if index != last_content {
            style.space_after = None;
        }
    }
}

fn is_empty_paragraph_block(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => paragraph.runs.is_empty(),
        Block::Caption(caption) => caption.paragraph.runs.is_empty(),
        _ => false,
    }
}

/// Emit whatever a removed paragraph mark held back, because what comes next is
/// not the paragraph it would have merged into — a table, the end of a cell, or
/// the end of the flow. Word never removes a container's last paragraph mark, so
/// this keeps malformed input from losing text rather than serving a real shape.
/// An empty withheld paragraph is dropped: that is the case issue #1710 is
/// about, and it has no text to rescue.
fn withheld_paragraph_block(ctx: &DocxConversionContext) -> Option<Block> {
    ctx.paragraph_marks
        .take_withheld()
        .filter(|paragraph| !paragraph.runs.is_empty())
        .map(Block::Paragraph)
}

/// A bare paragraph mark carrying only section properties marks the boundary;
/// Word does not lay it out as an extra blank body line.
fn is_empty_section_break_paragraph(para: &docx_rs::Paragraph) -> bool {
    if !para.children.is_empty() || para.property.section_property.is_none() {
        return false;
    }
    let mut properties: docx_rs::ParagraphProperty = para.property.clone();
    properties.section_property = None;
    properties == docx_rs::ParagraphProperty::default()
}

fn push_inline_images(
    out: &mut Vec<Block>,
    inline_images: &mut Vec<Block>,
    alignment: Option<Alignment>,
    spacing: Option<ImageParagraphSpacing>,
) {
    let mut grouped: Vec<ImageData> = Vec::new();

    for block in inline_images.drain(..) {
        match block {
            Block::Image(mut image) => {
                // Inline images inherit the containing paragraph's alignment
                // and its `w:spacing`: the picture consumes the paragraph, so
                // the gaps Word draws around it have to travel with the
                // picture instead (issue #499).
                if image.alignment.is_none() {
                    image.alignment = alignment;
                }
                if image.paragraph_spacing.is_none() {
                    image.paragraph_spacing = spacing;
                }
                grouped.push(image)
            }
            other => {
                flush_inline_image_group(out, &mut grouped);
                out.push(other);
            }
        }
    }
    flush_inline_image_group(out, &mut grouped);
}

fn flush_inline_image_group(out: &mut Vec<Block>, grouped: &mut Vec<ImageData>) {
    match grouped.len() {
        0 => {}
        1 => out.push(Block::Image(grouped.pop().expect("one inline image"))),
        _ => out.push(Block::InlineImages(std::mem::take(grouped))),
    }
}

/// The paragraph's explicit horizontal alignment, if any.
fn paragraph_alignment(para: &docx_rs::Paragraph) -> Option<Alignment> {
    extract_paragraph_style(&para.property).alignment
}

/// The `w:spacing` a picture paragraph contributes to the flow.
///
/// Resolved through the same style merge a text paragraph uses, so spacing
/// inherited from `styles.xml` counts as well as direct formatting.
fn paragraph_image_spacing(
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    contextual_spacing: ParagraphContextualSpacing<'_>,
) -> Option<ImageParagraphSpacing> {
    let style: ParagraphStyle = merge_paragraph_style(
        &extract_paragraph_style(&para.property),
        None,
        resolved_style,
    );
    let mut spacing = ImageParagraphSpacing {
        before: style.space_before,
        after: style.space_after,
    };
    contextual_spacing.apply(&mut spacing.before, &mut spacing.after);
    (spacing != ImageParagraphSpacing::default()).then_some(spacing)
}

fn push_paragraph_from_runs(
    out: &mut Vec<Block>,
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    style_map: &StyleMap,
    flow: ParagraphFlow<'_>,
    runs: &mut Vec<Run>,
    caption_identifier: Option<&str>,
) {
    out.push(build_paragraph_block(
        para,
        resolved_style,
        style_map,
        flow,
        runs,
        caption_identifier,
    ));
}

/// The block a paragraph's collected runs become: a caption when the paragraph
/// carries a `SEQ` field a `TOC \a` list collects, an ordinary paragraph
/// otherwise. Built rather than pushed so a paragraph whose mark was removed can
/// be held back for the next one instead (issue #1710).
fn build_paragraph_block(
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    style_map: &StyleMap,
    flow: ParagraphFlow<'_>,
    runs: &mut Vec<Run>,
    caption_identifier: Option<&str>,
) -> Block {
    let mut explicit_para_style = extract_paragraph_style(&para.property);
    explicit_para_style.background = flow.background;
    explicit_para_style.word_wrap = flow.word_wrap;
    let explicit_tab_overrides = extract_tab_stop_overrides(&para.property.tabs);
    let mut style = merge_paragraph_style(
        &explicit_para_style,
        explicit_tab_overrides.as_deref(),
        resolved_style,
    );
    if explicit_para_style.space_after.is_none()
        && !resolved_style.is_some_and(|resolved| resolved.paragraph_space_after_is_explicit)
        && let Some(table_space_after) = flow.table_style_space_after
    {
        style.space_after = Some(table_space_after);
    }
    if flow.is_rtl {
        style.direction = Some(TextDirection::Rtl);
    }
    apply_word_compatible_paragraph_defaults(
        &mut style,
        flow.paragraph_property_defaults_are_declared,
    );
    // After the fallback gap is in place, so a dropped `w:after` includes it.
    flow.contextual_spacing
        .apply(&mut style.space_before, &mut style.space_after);
    // Word's built-in Korean Normal — in force exactly when no document-defined
    // style resolves for the paragraph — breaks Hangul lines at character
    // level, where a document-defined style keeps each eojeol whole. Measured
    // by the #833 probe series: the same bare sentence flips between breaking
    // `표시되어야` after `표` and declining that syllable at 524.1pt of a
    // 524.45pt measure when only a default-style definition is added, a
    // referenced `ListParagraph` or `Heading6` flips it alone, and `w:numPr`
    // without a style does not. Same trigger as the auto space below
    // (issue #732); every paragraph #626 measured as eojeol-whole is
    // `ListParagraph`-styled. Recorded as an effective `w:wordWrap` default so
    // an explicit `w:val` keeps outranking it either way (issue #730).
    if style.word_wrap.is_none() && !flow.effective_style_is_defined {
        style.word_wrap = Some(false);
    }
    // Word's automatic East Asian/Latin space, applied once per paragraph so a
    // boundary falling between two runs is caught too. It goes only to
    // paragraphs whose effective style the document defines: Word's built-in
    // Korean Normal suppresses the space, and any explicit definition — the
    // paragraph's own resolvable `w:pStyle`, or a defined default style —
    // replaces that built-in and restores the spec default of on
    // (issue #732). This one predicate is what the earlier container and
    // alignment readings were each seeing a slice of: the corpus cells
    // (issue #627), its centred date line (issue #728) and its justified
    // paragraphs are all *bare* paragraphs in packages that define no default
    // style, while its widened list items and #521's all-widening probe are
    // styled or Normal-defining.
    let entry_text: Option<String> = caption_identifier.map(|_| caption_entry_text(runs));
    // Alignment is not part of the predicate. A one-factor probe that patched
    // only `w:jc` in a Normal-defining package measured left, centred,
    // justified and right at the same +2.588pt per boundary, and a stretch
    // sweep showed why the earlier justified reading looked different: Word
    // hands a line's justification demand to its word spaces first, widening
    // the auto space only once they reach half an em. Our quarter em is
    // therefore what Word draws for every line whose demand its word spaces
    // absorb, and it is the natural width Word breaks lines on either way
    // (issue #1053).
    if flow.effective_style_is_defined {
        insert_east_asian_auto_space(runs);
    }
    // Word gives a `<w:p>` with no runs a full line of its own, sized from the
    // paragraph mark's resolved run formatting — so a table row of nothing but
    // empty paragraphs is as tall as a row of text (issue #1700). Nothing
    // downstream can rebuild that resolution: the mark paints no glyph, so it
    // reaches codegen as no run at all. Resolved through the same cascade a
    // run takes, so `w:pPr/w:rPr` outranks the `w:pStyle` and both outrank
    // `w:rPrDefault`.
    if runs.is_empty() {
        style.paragraph_mark_text_style = Some(Box::new(resolve_run_style(
            &para.property.run_property,
            false,
            resolved_style,
            style_map,
            flow.theme_fonts,
        )));
    }
    let paragraph = Paragraph {
        style,
        runs: std::mem::take(runs),
    };
    match (caption_identifier, entry_text) {
        (Some(identifier), Some(entry_text)) => Block::Caption(Caption {
            identifier: identifier.to_string(),
            entry_text,
            paragraph,
        }),
        _ => Block::Paragraph(paragraph),
    }
}

/// The text a `TOC \a` list shows for a caption.
///
/// Word lists the caption without the label and the number that precede it —
/// `종전 헤드리스 변환 스택과 …`, not `그림 1  종전 헤드리스 변환 스택과 …`.
/// The number is its own run, produced by the `SEQ` field, so everything from
/// the run after it onward is the caption proper.
///
/// Read before `insert_east_asian_auto_space` rewrites the runs: those markers
/// are an instruction about the caption's own layout, and the list entry is a
/// separate piece of text that gets its own. Taking the text afterwards
/// carried them into the entry, where they rendered as stray glyphs.
fn caption_entry_text(runs: &[Run]) -> String {
    let after_number = runs
        .iter()
        .position(|run| !run.text.is_empty() && run.text.chars().all(|c| c.is_ascii_digit()))
        .map(|index| index + 1)
        .unwrap_or(0);
    runs[after_number..]
        .iter()
        .map(|run| run.text.as_str())
        .collect::<String>()
        .trim()
        .to_string()
}

#[cfg(test)]
#[path = "docx_tests.rs"]
mod tests;
