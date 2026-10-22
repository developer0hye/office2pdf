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
    FloatingTable, FloatingTextBox, HorizontalRule, ImageData, ImageFormat, ImageParagraphIndent,
    ImageParagraphSpacing, InlineTextBox, Insets, Ligatures, LineCap, LineJoin, LineSpacing, Page,
    PageNumbering, PairKerning, Paragraph, ParagraphStyle, Run, StyleSheet, TabAlignment,
    TabLeader, TabStop, Table, TableCell, TableOfContents, TableRow, TextDirection, TextFrame,
    TextStyle, VerticalTextAlign, WordCompatibilityMode,
};
use crate::parser::Parser;

#[cfg(test)]
use self::contexts::scan_table_headers;
use self::contexts::{
    BidiContext, ChartContext, ContextualSpacingContext, DocxConversionContext,
    DrawingShapeContext, DrawingTextBoxContext, DrawingTextBoxInfo, FieldContext, MathContext,
    NoteContent, NoteContext, ParagraphContextualSpacing, ParagraphMarkContext, ParagraphRunScan,
    ParagraphShadingContext, PictureCropContext, RunScan, RunScanContext, SmallCapsContext,
    TableHeaderContext, TableStyleContext, VmlTextBoxContext, VmlTextBoxInfo, WordWrapContext,
    WpgDrawingInfo, WrapContext, build_chart_context_from_xml, build_math_context_from_xml,
    build_note_context_from_xml, build_wrap_context_from_xml,
    extract_column_layout_from_section_property, is_note_reference_run, read_zip_text,
    scan_column_layouts, scan_line_numbering, scan_page_borders, scan_page_numbering,
    scan_section_alignment, scan_style_paragraph_shading, scan_style_word_wrap, seq_identifier,
    toc_caption_identifier, toc_heading_depth,
};
use self::lists::{
    NumberingMap, StyleSuppliedIndents, TaggedElement, build_numbering_map, extract_num_info,
    group_into_lists,
};
use self::media::{
    DrawingTextBoxPlacement, extract_drawing_image, extract_drawing_text_box_blocks,
    extract_shape_image, extract_vml_shape_text_box,
};
#[cfg(test)]
use self::sections::extract_page_size;
use self::sections::{
    HeaderFooterAssets, HeaderFooterStyleContext, LinkedFirstStories, SectionHeaderFooterInputs,
    SectionOverrides, body_text_frame, build_flow_page_from_section, build_header_footer_assets,
    fixed_text_frame, top_margin_frame_box,
};
#[cfg(test)]
use self::styles::parse_ligatures_value;
use self::styles::{
    DOC_DEFAULT_STYLE_ID, LigatureRules, PairKerningRules, ResolvedStyle, StyleMap,
    TabStopOverride, apply_tab_stop_overrides, build_style_map, effective_frame_property,
    get_paragraph_style_id, merge_paragraph_style, merge_text_style, paragraph_style_frames,
    resolve_doc_default_text_style,
};
use self::tables::{convert_table, convert_table_in_text_box};
use self::text::{
    ThemeFonts, WORD_UNSTATED_FONT_SIZE_PT, extract_doc_default_paragraph_style,
    extract_doc_default_text_style_with_theme, extract_indent, extract_paragraph_style,
    extract_run_style, extract_run_style_id, extract_run_text, extract_run_text_marking_clears,
    extract_tab_stop_overrides, insert_east_asian_auto_space, is_column_break, is_page_break,
    pair_kerning_from_half_points, parse_hex_color, parse_theme_fonts, resolve_hyperlink_url,
    resolve_latin_theme_font_family, resolve_theme_font_family, take_clearing_break_markers,
    word_built_in_doc_defaults,
};
#[cfg(test)]
use self::text::{extract_pair_kerning, extract_tab_stops, resolve_highlight_color};

#[path = "docx_contexts.rs"]
mod contexts;
#[path = "docx_image_assets.rs"]
mod image_assets;
#[path = "docx_lists.rs"]
mod lists;
#[path = "docx_media.rs"]
mod media;
#[path = "docx_sections.rs"]
mod sections;
#[path = "docx_smartart.rs"]
mod smartart;
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
/// PNG previews may be raw, unchecked bytes. Validate them with the renderer's
/// decoder before use; retain valid bytes and recover damaged pictures without
/// losing their declared layout extents. Empty previews remain available to the
/// separate metafile converter rather than becoming empty PNG assets.
fn build_image_map(docx: &docx_rs::Docx, warnings: &mut Vec<ConvertWarning>) -> ImageMap {
    docx.images
        .iter()
        .filter(|(_id, _path, _image, png)| !png.0.is_empty())
        .map(|(id, _path, _image, png)| {
            (
                id.clone(),
                image_assets::normalize_png_preview(&png.0, id, warnings),
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
    page_borders: Vec<Option<crate::ir::PageBorders>>,
    line_numbering: Vec<Option<crate::ir::LineNumbering>>,
    section_alignment: Vec<Option<crate::ir::PageVerticalAlignment>>,
    metafile_images: ImageMap,
    theme_fonts: ThemeFonts,
    default_paragraph_style_id: Option<String>,
    style_paragraph_backgrounds: HashMap<String, Color>,
    style_word_wraps: HashMap<String, bool>,
    /// Read from the raw `word/styles.xml` because docx-rs has no field for
    /// `w:kern` (issue #628).
    pair_kerning: PairKerningRules,
    /// Read from the raw `word/styles.xml` because docx-rs has no field for
    /// `w14:ligatures` (issue #2205).
    ligatures: LigatureRules,
    /// Read from the raw `word/styles.xml` because docx-rs has no field for
    /// `w:lang`.
    default_language: Option<DocumentLanguage>,
    /// Whether the package has a `word/styles.xml`; one without it takes
    /// Word's built-in document defaults (issue #2215).
    styles_part_is_present: bool,
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
            let page_borders = doc_xml
                .as_deref()
                .map(scan_page_borders)
                .unwrap_or_default();
            let section_alignment = doc_xml
                .as_deref()
                .map(scan_section_alignment)
                .unwrap_or_default();
            let line_numbering = doc_xml
                .as_deref()
                .map(scan_line_numbering)
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
                smartart: smartart::SmartArtContext::from_archive(
                    doc_xml.as_deref(),
                    theme_xml.as_deref(),
                    &mut archive,
                ),
                notes,
                wraps,
                picture_crops: PictureCropContext::from_xml(doc_xml.as_deref()),
                drawing_text_boxes,
                drawing_shapes,
                table_headers,
                table_styles,
                vml_text_boxes,
                bidi,
                small_caps,
                run_scan: RunScanContext::from_xml(doc_xml.as_deref()),
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
                page_borders,
                line_numbering,
                section_alignment,
                metafile_images,
                theme_fonts,
                default_paragraph_style_id,
                style_paragraph_backgrounds,
                style_word_wraps,
                pair_kerning: PairKerningRules::from_styles_xml(styles_xml.as_deref()),
                ligatures: LigatureRules::from_styles_xml(styles_xml.as_deref()),
                default_language,
                styles_part_is_present: styles_xml.is_some(),
            }
        }
        Err(_) => ZipPreParseAssets {
            metadata: crate::ir::Metadata::default(),
            ctx: DocxConversionContext {
                theme_fonts: ThemeFonts::default(),
                smartart: smartart::SmartArtContext::default(),
                notes: NoteContext::empty(),
                wraps: WrapContext::empty(),
                picture_crops: PictureCropContext::default(),
                drawing_text_boxes: DrawingTextBoxContext::from_xml(None),
                drawing_shapes: DrawingShapeContext::from_xml(None),
                table_headers: TableHeaderContext::from_xml(None),
                table_styles: TableStyleContext::from_xml(None, None),
                vml_text_boxes: VmlTextBoxContext::from_xml(None),
                bidi: BidiContext::from_xml(None),
                small_caps: SmallCapsContext::from_xml(None),
                run_scan: RunScanContext::from_xml(None),
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
            page_borders: Vec::new(),
            line_numbering: Vec::new(),
            section_alignment: Vec::new(),
            metafile_images: ImageMap::new(),
            theme_fonts: ThemeFonts::default(),
            default_paragraph_style_id: None,
            style_paragraph_backgrounds: HashMap::new(),
            style_word_wraps: HashMap::new(),
            pair_kerning: PairKerningRules::default(),
            ligatures: LigatureRules::default(),
            default_language: None,
            // docx-rs rejects a package that is not a ZIP, so nothing reads it.
            styles_part_is_present: true,
        },
    }
}

impl Parser for DocxParser {
    fn parse(
        &self,
        data: &[u8],
        _options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let (expanded, mut warnings) = super::docx_alt_chunk::expand(data)?;
        let data = expanded.as_ref();
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
            page_borders,
            line_numbering,
            section_alignment,
            metafile_images,
            theme_fonts,
            default_paragraph_style_id,
            style_paragraph_backgrounds,
            style_word_wraps,
            pair_kerning,
            ligatures,
            default_language,
            styles_part_is_present,
        } = build_zip_preparse_assets(data);

        warnings.append(&mut ctx.smartart.warnings);

        if section_alignment
            .iter()
            .flatten()
            .any(|alignment| *alignment == crate::ir::PageVerticalAlignment::Justified)
        {
            tracing::warn!(
                format = "DOCX",
                feature = "section vertical justification",
                fallback = "top",
                "Unsupported section alignment"
            );
            warnings.push(ConvertWarning::ParseSkipped {
                format: "DOCX".to_string(),
                reason: "Section vertical justification is unsupported; rendering at the top"
                    .to_string(),
            });
        }
        let mut docx = docx_rs::read_docx(data).map_err(|e| {
            crate::parser::parse_err(format!("Failed to parse DOCX (docx-rs): {e}"))
        })?;
        // docx-rs leaves `Styles` empty when it finds no styles part; the
        // comparison keeps a part it found elsewhere through its relationship.
        if !styles_part_is_present && docx.styles == docx_rs::Styles::new() {
            tracing::debug!(
                format = "DOCX",
                "No styles part; applying Word's built-in document defaults"
            );
            docx.styles.doc_defaults = word_built_in_doc_defaults(&theme_fonts);
        }

        let numbering_xml = zip::ZipArchive::new(std::io::Cursor::new(data))
            .ok()
            .and_then(|mut archive| read_zip_text(&mut archive, "word/numbering.xml"));
        if let Some(xml) = numbering_xml {
            contexts::normalize_legal_numbering(&xml, &mut docx.numberings);
        }

        // Populate locale-specific footnote/endnote style IDs from docx styles
        ctx.notes.populate_style_ids(&docx.styles);

        let mut images = build_image_map(&docx, &mut warnings);
        images.extend(metafile_images);
        let hyperlinks = build_hyperlink_map(&docx);
        let style_map = build_style_map(
            &docx.styles,
            &theme_fonts,
            default_paragraph_style_id.as_deref(),
            &style_paragraph_backgrounds,
            &style_word_wraps,
            &pair_kerning,
            &ligatures,
        );
        let numberings = build_numbering_map(&docx.numberings, &style_map, &theme_fonts);

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
            .map(|mut archive| {
                build_header_footer_assets(&mut archive, header_footer_styles, &mut warnings)
            })
            .unwrap_or_default();
        let mut elements: Vec<TaggedElement> = Vec::new();
        let mut pages: Vec<Page> = Vec::new();
        // Each section's `first` stories as the next section links to them,
        // kept apart from `pages` because a section without `w:titlePg` passes
        // them on without showing them (issue #2224).
        let mut linked_first_stories = LinkedFirstStories::default();
        let mut section_layout_index: usize = 0;
        // The body index at which the frame of the last top-margin frame box
        // opened starts, so the frame's later lines join that box.
        let mut open_frame_box: Option<usize> = None;
        // The last run of children sharing a frame at a fixed place on its
        // page, so its later children join that frame.
        let mut open_fixed_frame: Option<FixedFrameGroup> = None;
        // The body index the current section's first child sits at.
        let mut section_start: usize = 0;
        // The `w:after` of an empty section mark that took no line before a
        // continuous section, until that section's first block takes it.
        let mut section_mark_after: Option<f64> = None;
        let style_frames: HashMap<String, docx_rs::FrameProperty> =
            paragraph_style_frames(&docx.styles);
        let body_frames: Vec<Option<docx_rs::FrameProperty>> = docx
            .document
            .children
            .iter()
            .map(|child| {
                leading_frame_property(child, &style_frames, default_paragraph_style_id.as_deref())
            })
            .collect();
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
            let mut dropped_mark_after: Option<f64> = None;
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| match child {
                docx_rs::DocumentChild::Paragraph(para) => {
                    let mut tagged: Vec<TaggedElement> = convert_paragraph_element(
                        para,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &ctx,
                        &docx.styles,
                    );
                    if let Some(frame) = body_frames[idx].as_ref().and_then(fixed_text_frame) {
                        tagged = join_fixed_text_frame(
                            &mut elements,
                            &mut open_fixed_frame,
                            frame,
                            tagged,
                            &docx.document.children,
                            &body_frames,
                            idx,
                        );
                    } else if is_empty_lifted_frame(
                        &tagged,
                        &docx.document.children,
                        &body_frames,
                        idx,
                    ) {
                        tracing::debug!(
                            format = "DOCX",
                            body_index = idx,
                            "Empty framed paragraph leaves no body line"
                        );
                        let lines: Vec<TaggedElement> = std::mem::take(&mut tagged);
                        if word_compatibility_mode == WordCompatibilityMode::Word2013OrLater {
                            tagged = keep_line_in_top_margin_frame_box(
                                &mut elements,
                                &mut open_frame_box,
                                para,
                                lines,
                                &body_frames,
                                idx,
                            );
                        }
                    } else if let Some(frame) = lifted_text_frame(&tagged, &body_frames, idx) {
                        tracing::debug!(
                            format = "DOCX",
                            body_index = idx,
                            x_pt = frame.x,
                            y_pt = frame.y,
                            width_pt = frame.width,
                            height_pt = frame.height,
                            "Framed paragraph lifted out of the body flow"
                        );
                        tagged = vec![TaggedElement::Plain(vec![Block::TextFrame(Box::new(
                            frame,
                        ))])];
                    }
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
                    if para.property.section_property.is_some()
                        && body_frames[idx].is_none()
                        && is_blank_paragraph_output(&tagged)
                        && !empty_section_mark_keeps_line(
                            &docx.document.children,
                            section_start,
                            idx,
                            elements.last(),
                            word_compatibility_mode,
                            &docx.document.section_property,
                        )
                    {
                        tracing::debug!(
                            format = "DOCX",
                            body_index = idx,
                            "Empty section mark takes no line"
                        );
                        // After a page or column break the next section opens
                        // at a page or column top, which has no gap to share.
                        if next_section_starts_continuously(
                            &docx.document.children,
                            idx,
                            &docx.document.section_property,
                        ) && !ends_with_layout_break(elements.last())
                        {
                            dropped_mark_after = blank_paragraph_space_after(&tagged);
                        }
                        tagged.clear();
                    }
                    tagged
                }
                docx_rs::DocumentChild::Table(table) => {
                    let tagged: Vec<TaggedElement> =
                        vec![TaggedElement::Plain(vec![convert_body_table_block(
                            table,
                            &images,
                            &hyperlinks,
                            &style_map,
                            &ctx,
                        )])];
                    match body_frames[idx].as_ref().and_then(fixed_text_frame) {
                        Some(frame) => join_fixed_text_frame(
                            &mut elements,
                            &mut open_fixed_frame,
                            frame,
                            tagged,
                            &docx.document.children,
                            &body_frames,
                            idx,
                        ),
                        None => tagged,
                    }
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
                Ok(mut elems) => {
                    if is_body_block(child)
                        && let Some(mark_after) = section_mark_after.take()
                    {
                        pass_on_section_mark_after(&mut elems, mark_after);
                    }
                    elements.extend(elems);
                }
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
            if dropped_mark_after.is_some() {
                section_mark_after = dropped_mark_after;
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
                let (flow_page, section_first_stories) = build_flow_page_from_section(
                    section_prop,
                    std::mem::take(&mut elements),
                    &numberings,
                    SectionHeaderFooterInputs {
                        assets: &header_footer_assets,
                        inherited_header: previous_flow_page.and_then(|page| page.header.as_ref()),
                        inherited_footer: previous_flow_page.and_then(|page| page.footer.as_ref()),
                        inherited_first: &linked_first_stories,
                    },
                    SectionOverrides {
                        column_layout,
                        page_numbering: page_numbering.get(section_layout_index).copied().flatten(),
                        page_borders: page_borders.get(section_layout_index).cloned().flatten(),
                        line_numbering: line_numbering.get(section_layout_index).cloned().flatten(),
                        vertical_alignment: section_alignment
                            .get(section_layout_index)
                            .copied()
                            .flatten(),
                    },
                    header_footer_styles,
                    &mut warnings,
                );
                linked_first_stories = section_first_stories;
                section_start = idx + 1;
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
        let (final_flow_page, _) = build_flow_page_from_section(
            &docx.document.section_property,
            elements,
            &numberings,
            SectionHeaderFooterInputs {
                assets: &header_footer_assets,
                inherited_header: previous_flow_page.and_then(|page| page.header.as_ref()),
                inherited_footer: previous_flow_page.and_then(|page| page.footer.as_ref()),
                inherited_first: &linked_first_stories,
            },
            SectionOverrides {
                column_layout: final_column_layout,
                page_numbering: page_numbering.get(section_layout_index).copied().flatten(),
                page_borders: page_borders.get(section_layout_index).cloned().flatten(),
                line_numbering: line_numbering.get(section_layout_index).cloned().flatten(),
                vertical_alignment: section_alignment
                    .get(section_layout_index)
                    .copied()
                    .flatten(),
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

        // The renderer can isolate explicit pages, but has no physical-page
        // boundaries for automatic flow inside differing continuous sections.
        let alignment = |page: &Page| match page {
            Page::Flow(flow) | Page::FlowContinuous(flow) => match flow.vertical_alignment {
                Some(crate::ir::PageVerticalAlignment::Center) => {
                    crate::ir::PageVerticalAlignment::Center
                }
                Some(crate::ir::PageVerticalAlignment::Bottom) => {
                    crate::ir::PageVerticalAlignment::Bottom
                }
                _ => crate::ir::PageVerticalAlignment::Top,
            },
            _ => crate::ir::PageVerticalAlignment::Top,
        };
        if pages.windows(2).any(|pair| {
            matches!(pair[1], Page::FlowContinuous(_)) && alignment(&pair[0]) != alignment(&pair[1])
        }) {
            tracing::warn!(
                format = "DOCX",
                feature = "mixed continuous section alignment",
                fallback = "top within shared explicit segment",
                "Automatic pagination alignment is approximate"
            );
            warnings.push(ConvertWarning::ParseSkipped {
                format: "DOCX".to_string(),
                reason: "Mixed continuous section alignment is approximate during automatic pagination; explicit page breaks delimit the top-aligned shared segment".to_string(),
            });
        }

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
                        &ligatures,
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

/// Which indents a paragraph with a direct `w:numPr` takes from its style
/// chain because its own `w:ind` does not state them. None when the numbering
/// comes from the style.
fn style_supplied_indents(para: &docx_rs::Paragraph, style_map: &StyleMap) -> StyleSuppliedIndents {
    if !para.has_numbering {
        return StyleSuppliedIndents::default();
    }
    let (left, _, first_line) = extract_indent(&para.property.indent);
    let style: Option<&ParagraphStyle> = get_paragraph_style_id(&para.property)
        .and_then(|id| style_map.get(id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID))
        .map(|resolved| &resolved.paragraph);
    StyleSuppliedIndents {
        left: left.is_none() && style.is_some_and(|style| style.indent_left.is_some()),
        first_line: first_line.is_none()
            && style.is_some_and(|style| style.indent_first_line.is_some()),
    }
}

fn convert_body_table_block(
    source_table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> Block {
    // Read before the conversion consumes the table's scan entry.
    let vertical_from_text_twips: (Option<f64>, Option<f64>) =
        ctx.table_headers.next_vertical_from_text_twips();
    let table: Table = convert_table(source_table, images, hyperlinks, style_map, ctx, 0);
    match sections::floating_table_frame(source_table, &table, vertical_from_text_twips) {
        Some(frame) => Block::FloatingTable(Box::new(FloatingTable { table, frame })),
        None => Block::Table(table),
    }
}

/// The formatting a numbered paragraph's list label takes before its level's
/// own `w:rPr`: the paragraph mark's run properties, resolved through the
/// cascade a run takes, so `w:pPr/w:rPr` outranks the `w:pStyle` and both
/// outrank `w:rPrDefault`. Word formats the label this way whatever the
/// item's runs state (issue #2337).
///
/// The mark's underline, vertical position and small capitals stay off the
/// label, as LibreOffice's `IsIgnoredCharFormatForNumbering` and
/// `checkApplyParagraphMarkFormatToNumbering` record for Word.
fn numbering_label_style(
    para: &docx_rs::Paragraph,
    style_map: &StyleMap,
    theme_fonts: &ThemeFonts,
) -> TextStyle {
    let resolved_style: Option<&ResolvedStyle> = get_paragraph_style_id(&para.property)
        .and_then(|id| style_map.get(id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID));
    let mut style: TextStyle = resolve_run_style(
        &para.property.run_property,
        false,
        resolved_style,
        style_map,
        theme_fonts,
    );
    style.underline = None;
    style.vertical_align = None;
    style.baseline_shift = None;
    style.small_caps = None;
    style
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
    let num_info = extract_num_info(para, styles).map(|mut info| {
        info.style_supplied_indents = style_supplied_indents(para, style_map);
        info
    });

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
        paragraph.style.numbering_label_style = Some(Box::new(numbering_label_style(
            para,
            style_map,
            &ctx.theme_fonts,
        )));
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

/// Lay a run's own `w14:ligatures` over its resolved style. A run's direct
/// formatting outranks every style it inherits from, so this needs no merge
/// (issue #2205).
fn apply_direct_ligatures(style: &mut TextStyle, direct_ligatures: Option<Ligatures>) {
    if direct_ligatures.is_some() {
        style.ligatures = direct_ligatures;
    }
}

/// Intermediate results from scanning a run segment for media and text boxes.
struct RunChildrenMedia {
    text_box_blocks: Vec<Block>,
    /// The `wp:inline` text boxes this run anchors. They belong on the anchor
    /// paragraph's own line, so they ride the run rather than becoming flow
    /// blocks beneath it (issue #1690).
    inline_text_boxes: Vec<InlineTextBox>,
    /// The `wp:inline` pictures this run holds when they share their
    /// paragraph with text, each boxed at its extent: like the text boxes,
    /// they ride the line (issue #2227).
    inline_pictures: Vec<InlineTextBox>,
}

/// The ordered pieces of a run around structural layout breaks.
enum RunLayoutPart {
    Content(Box<docx_rs::Run>),
    PageBreak,
    ColumnBreak,
}

/// Give each `wp:inline` picture of a content part a part of its own, so a
/// picture riding its paragraph's line keeps its place among the run's text
/// (issue #2227). Other parts pass through unchanged.
fn split_content_at_inline_pictures(part: RunLayoutPart) -> Vec<RunLayoutPart> {
    let RunLayoutPart::Content(run) = part else {
        return vec![part];
    };
    let is_inline_picture = |child: &docx_rs::RunChild| {
        matches!(
            child,
            docx_rs::RunChild::Drawing(drawing)
                if matches!(
                    &drawing.data,
                    Some(docx_rs::DrawingData::Pic(pic))
                        if pic.position_type == docx_rs::DrawingPositionType::Inline
                )
        )
    };
    if run.children.len() < 2 || !run.children.iter().any(is_inline_picture) {
        return vec![RunLayoutPart::Content(run)];
    }

    let mut parts: Vec<RunLayoutPart> = Vec::new();
    let mut pending: Vec<docx_rs::RunChild> = Vec::new();
    let flush = |children: Vec<docx_rs::RunChild>, parts: &mut Vec<RunLayoutPart>| {
        if !children.is_empty() {
            let mut content_run: docx_rs::Run = (*run).clone();
            content_run.children = children;
            parts.push(RunLayoutPart::Content(Box::new(content_run)));
        }
    };
    for child in &run.children {
        if is_inline_picture(child) {
            flush(std::mem::take(&mut pending), &mut parts);
            flush(vec![child.clone()], &mut parts);
        } else {
            pending.push(child.clone());
        }
    }
    flush(pending, &mut parts);
    parts
}

/// For each stretch of the paragraph between page and column breaks, whether
/// it holds text an inline picture would share a line with: any character,
/// including a space, or a tab or line break.
fn line_text_by_layout_fragment(para: &docx_rs::Paragraph) -> Vec<bool> {
    let mut fragments: Vec<bool> = vec![false];
    let mark_text = |run: &docx_rs::Run, fragments: &mut Vec<bool>, splits: bool| {
        for child in &run.children {
            let is_line_text: bool = match child {
                docx_rs::RunChild::Text(text) => !text.text.is_empty(),
                docx_rs::RunChild::Tab(_) => true,
                docx_rs::RunChild::Break(br) if is_page_break(br) || is_column_break(br) => {
                    if splits {
                        fragments.push(false);
                    }
                    false
                }
                docx_rs::RunChild::Break(_) => true,
                _ => false,
            };
            if is_line_text {
                *fragments.last_mut().expect("one fragment at least") = true;
            }
        }
    };
    for item in flatten_tracked_changes(&para.children) {
        match item {
            // Only a paragraph's own runs split at layout breaks; a
            // hyperlink's runs stay in the fragment they start in.
            ParagraphItem::Run(run) => mark_text(run, &mut fragments, true),
            ParagraphItem::Hyperlink(hyperlink) => {
                for child in &hyperlink.children {
                    if let docx_rs::ParagraphChild::Run(run) = child {
                        mark_text(run, &mut fragments, false);
                    }
                }
            }
            ParagraphItem::PageNum | ParagraphItem::NumPages => {}
        }
    }
    fragments
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
/// Extracted images are pushed to `inline_images`, except that a sized
/// `wp:inline` picture is returned in `RunChildrenMedia` when
/// `pictures_ride_runs`; text boxes are returned there too.
fn extract_run_children_media(
    run: &docx_rs::Run,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    inline_images: &mut Vec<Block>,
    pictures_ride_runs: bool,
) -> RunChildrenMedia {
    let mut text_box_blocks: Vec<Block> = Vec::new();
    let mut inline_text_boxes: Vec<InlineTextBox> = Vec::new();
    let mut inline_pictures: Vec<InlineTextBox> = Vec::new();

    for run_child in &run.children {
        if let docx_rs::RunChild::Drawing(drawing) = run_child {
            let wpg_drawing: Option<WpgDrawingInfo> = ctx.drawing_shapes.consume_wpg_drawing();
            let canvas_image_offset: Option<(f64, f64)> =
                ctx.drawing_shapes.consume_canvas_image_offset();
            let metadata = ctx.picture_crops.next_drawing(false);
            let crop = metadata.crop;
            let emitted: [usize; 4] = [
                text_box_blocks.len(),
                inline_text_boxes.len(),
                inline_pictures.len(),
                inline_images.len(),
            ];
            if let Some(mut group) = ctx.smartart.consume_next() {
                if let Some(anchor) = group.anchor.as_mut() {
                    anchor.wrap_mode = ctx.wraps.consume_next();
                    text_box_blocks.push(Block::DrawingGroup(Box::new(group)));
                } else {
                    // Word seats an inline diagram on its paragraph's baseline,
                    // reserving its effect extents outside the drawing itself.
                    let padding: crate::ir::Insets = group.effect_extent;
                    inline_text_boxes.push(InlineTextBox {
                        width: group.width + padding.left + padding.right,
                        height: group.height + padding.top + padding.bottom,
                        content: vec![Block::DrawingGroup(Box::new(group))],
                        padding,
                        stroke: None,
                        fill: None,
                    });
                }
            } else if let Some(wpg_drawing) = wpg_drawing {
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
                match extract_drawing_image(drawing, images, &ctx.wraps, canvas_image_offset, crop)
                {
                    Some(Block::Image(image)) if pictures_ride_runs => {
                        match (image.width, image.height) {
                            (Some(width), Some(height)) => inline_pictures.push(InlineTextBox {
                                content: vec![Block::Image(image)],
                                width,
                                height,
                                padding: Insets::default(),
                                stroke: None,
                                fill: None,
                            }),
                            _ => inline_images.push(Block::Image(image)),
                        }
                    }
                    Some(img_block) => inline_images.push(img_block),
                    None => {}
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
            if metadata.hidden {
                // Converted anyway, so every side-channel cursor steps past
                // the drawing and its nested ones; then nothing of it stays,
                // not even the wrap it would reserve (issue #2299).
                text_box_blocks.truncate(emitted[0]);
                inline_text_boxes.truncate(emitted[1]);
                inline_pictures.truncate(emitted[2]);
                inline_images.truncate(emitted[3]);
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
        inline_pictures,
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
    run_scan: &ParagraphRunScan<'_>,
    runs: &mut Vec<Run>,
) {
    let href: Option<String> = resolve_hyperlink_url(hyperlink, hyperlinks);
    for hchild in &hyperlink.children {
        if let docx_rs::ParagraphChild::Run(run) = hchild {
            let hl_small_caps: bool = ctx.small_caps.next_is_small_caps();
            let scanned: RunScan<'_> = run_scan.next_run(run);
            let text: String = extract_run_text_marking_clears(
                run,
                &mut scanned.line_break_clears.iter().copied(),
            );
            if let Some(mut ir_run) = build_text_run(
                text,
                &run.run_property,
                hl_small_caps,
                resolved_style,
                style_map,
                &ctx.theme_fonts,
                href.clone(),
            ) {
                apply_direct_ligatures(&mut ir_run.style, scanned.ligatures);
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
/// Consecutive inline images with no text beside them are kept in one wrapping flow
/// container; an inline picture that shares its line with text rides the paragraph's
/// runs as `Run::inline_box` instead (issue #2227).
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
    let run_scan: ParagraphRunScan<'_> = ctx.run_scan.next_paragraph(style_id);

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
    // Word sets an inline picture on its paragraph's line like a glyph. When
    // the picture shares a stretch of the paragraph with text, it joins the
    // runs so it stays in order with that text on a shared line (issue
    // #2227). Pictures with no text beside them keep the picture-block path,
    // which carries the paragraph's alignment, spacing and indents (issues
    // #499, #2201).
    let mut line_text_fragments: Vec<bool> = line_text_by_layout_fragment(para);
    line_text_fragments[0] |= !runs.is_empty();
    let mut layout_fragment: usize = 0;
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
    // Word's horizontal lines, which docx-rs drops from the runs (issue #2251),
    // in the stretch of the paragraph since its last page or column break.
    // That stretch's blocks start at `fragment_block_start`.
    let mut horizontal_rules: Vec<HorizontalRule> = Vec::new();
    let mut fragment_block_start: usize = out.len();

    for child in flatten_tracked_changes(&para.children) {
        match child {
            ParagraphItem::Run(run) => {
                // Advance smallCaps cursor for every <w:r> in body
                let is_small_caps: bool = ctx.small_caps.next_is_small_caps();
                // Taken once per docx-rs run, before the run is split, so a
                // note reference still keeps the paragraph's runs in step.
                let scanned: RunScan<'_> = run_scan.next_run(run);
                let layout_parts: Vec<RunLayoutPart> = split_run_at_layout_breaks(run);
                // The scan does not place a rule among its own run's children,
                // so a rule in the run of a page or column break could be on
                // either side of it.
                if layout_parts.iter().any(|part| {
                    matches!(part, RunLayoutPart::PageBreak | RunLayoutPart::ColumnBreak)
                }) {
                    if !scanned.horizontal_rules.is_empty() {
                        tracing::debug!(
                            rules = scanned.horizontal_rules.len(),
                            "horizontal rule shares its run with a layout break; drawing none"
                        );
                    }
                } else {
                    horizontal_rules.extend_from_slice(scanned.horizontal_rules);
                }
                // The run's line breaks in order; its split parts keep that
                // order.
                let mut line_break_clears = scanned.line_break_clears.iter().copied();

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

                for part in layout_parts
                    .into_iter()
                    .flat_map(split_content_at_inline_pictures)
                {
                    match part {
                        break_part @ (RunLayoutPart::PageBreak | RunLayoutPart::ColumnBreak) => {
                            // Flush queued inline media even when no text run
                            // precedes the break; images are run children too.
                            push_inline_images(
                                out,
                                &mut inline_images,
                                para,
                                resolved_style,
                                flow.contextual_spacing,
                            );
                            if !runs.is_empty() {
                                clear_paragraph_fragment_borders(
                                    out,
                                    &empty_break_fragment_indices,
                                );
                                let mut fragment: Block = build_paragraph_block(
                                    para,
                                    resolved_style,
                                    style_map,
                                    flow,
                                    &mut runs,
                                    caption_identifier.as_deref(),
                                );
                                attach_horizontal_rule(&mut fragment, &horizontal_rules);
                                out.push(fragment);
                                emitted_paragraph = true;
                                emitted_content_paragraph = true;
                            } else if is_rule_line(
                                &horizontal_rules,
                                &runs,
                                out,
                                fragment_block_start,
                            ) {
                                let mut fragment: Block = build_paragraph_block(
                                    para,
                                    resolved_style,
                                    style_map,
                                    flow,
                                    &mut runs,
                                    caption_identifier.as_deref(),
                                );
                                attach_horizontal_rule(&mut fragment, &horizontal_rules);
                                out.push(fragment);
                                emitted_paragraph = true;
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
                            layout_fragment += 1;
                            horizontal_rules.clear();
                            fragment_block_start = out.len();
                        }
                        RunLayoutPart::Content(run_segment) => {
                            let media = extract_run_children_media(
                                &run_segment,
                                images,
                                hyperlinks,
                                style_map,
                                ctx,
                                &mut inline_images,
                                line_text_fragments
                                    .get(layout_fragment)
                                    .copied()
                                    .unwrap_or(false),
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
                                        para,
                                        resolved_style,
                                        flow.contextual_spacing,
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
                                        para,
                                        resolved_style,
                                        flow.contextual_spacing,
                                    );
                                }
                                out.extend(media.text_box_blocks);
                            }

                            // Extracted even when a field replaces it, so the
                            // later parts' breaks keep their clears.
                            let run_text: String = extract_run_text_marking_clears(
                                &run_segment,
                                &mut line_break_clears,
                            );
                            let text: String =
                                seq_field_text(&run_segment, &ctx.fields, &mut caption_identifier)
                                    .unwrap_or(run_text);
                            if let Some(mut ir_run) = build_text_run(
                                text,
                                &run_segment.run_property,
                                is_small_caps,
                                resolved_style,
                                style_map,
                                &ctx.theme_fonts,
                                None,
                            ) {
                                apply_direct_ligatures(&mut ir_run.style, scanned.ligatures);
                                runs.push(ir_run);
                            }

                            // An inline text box sits on this paragraph's line, after
                            // whatever text its own `w:r` carried, so it joins the runs
                            // rather than becoming a flow block (issue #1690).
                            for inline_box in media.inline_text_boxes {
                                let is_cached_drawing: bool = inline_box
                                    .content
                                    .iter()
                                    .any(|block| matches!(block, Block::DrawingGroup(_)));
                                // The cached drawing paints no host-run glyphs, but
                                // its resolved font still sizes automatic spacing.
                                let style: TextStyle = if is_cached_drawing {
                                    resolve_run_style(
                                        &run_segment.run_property,
                                        is_small_caps,
                                        resolved_style,
                                        style_map,
                                        &ctx.theme_fonts,
                                    )
                                } else {
                                    TextStyle::default()
                                };
                                runs.push(Run {
                                    text: String::new(),
                                    style,
                                    href: None,
                                    footnote: None,
                                    inline_box: Some(Box::new(inline_box)),
                                });
                            }

                            // A picture on the line is a box of its extent whose
                            // bottom edge sits on the baseline. Its own run adds
                            // no text metrics: in Word's `leading-picture-*`
                            // cases, a line holding the picture and only a
                            // trailing space has no descent below it.
                            for picture in media.inline_pictures {
                                runs.push(Run {
                                    text: String::new(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                    inline_box: Some(Box::new(picture)),
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
                    &run_scan,
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
        para,
        resolved_style,
        flow.contextual_spacing,
    );

    // The paragraph's last stretch holds only a rule, which is the line that
    // stretch lays out: after a break it is no mere break carrier.
    let ends_on_rule_line: bool = is_rule_line(&horizontal_rules, &runs, out, fragment_block_start);

    if emitted_layout_break
        && runs.is_empty()
        && !ends_on_rule_line
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
    let is_layout_break_carrier: bool =
        emitted_layout_break && runs.is_empty() && !ends_on_rule_line;

    if !is_layout_break_carrier
        && (!runs.is_empty()
            || !emitted_media_blocks
            || ends_on_rule_line
            || (emitted_floating_anchor && !emitted_paragraph))
    {
        if !runs.is_empty() {
            clear_paragraph_fragment_borders(out, &empty_break_fragment_indices);
        }
        // Keep paragraph marks for floating drawing anchors. The drawing itself
        // is positioned by offsets, but the source paragraph still contributes
        // to flow spacing between the drawing cluster and following content.
        let mut block: Block = build_paragraph_block(
            para,
            resolved_style,
            style_map,
            flow,
            &mut runs,
            caption_identifier.as_deref(),
        );
        attach_horizontal_rule(&mut block, &horizontal_rules);
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

/// Whether a stretch of a paragraph between its page and column breaks holds
/// nothing but one horizontal rule: no runs, and no block emitted since the
/// stretch began at `fragment_block_start`. That rule is the stretch's line,
/// laid out as a rule-only paragraph's (issue #2251).
///
/// TODO(Word export of a rule beside a break): no Word export of such a
/// paragraph is available yet to confirm where Word draws its rule.
fn is_rule_line(
    horizontal_rules: &[HorizontalRule],
    runs: &[Run],
    out: &[Block],
    fragment_block_start: usize,
) -> bool {
    horizontal_rules.len() == 1 && runs.is_empty() && out.len() == fragment_block_start
}

/// Give a paragraph, or the stretch of one between its page and column
/// breaks, whose only content is a horizontal rule that rule, which is drawn
/// on the paragraph's line (issue #2251).
///
/// TODO(rule beside text): Word gives a rule a line of its own inside a
/// paragraph that also holds text or another rule. That line is not laid out,
/// so such a paragraph draws no rule rather than one over its text.
///
/// Header and footer stories convert through their own path, which gives
/// their rules to their paragraphs itself (`attach_story_horizontal_rule`).
///
/// TODO(#2251 text box rules): a VML text box's paragraphs are rebuilt from
/// the raw XML, so they never reach this and their rules are not drawn yet.
fn attach_horizontal_rule(block: &mut Block, horizontal_rules: &[HorizontalRule]) {
    let [rule] = horizontal_rules else {
        if !horizontal_rules.is_empty() {
            tracing::debug!(
                rules = horizontal_rules.len(),
                "paragraph holds several horizontal rules; drawing none"
            );
        }
        return;
    };
    let paragraph: &mut Paragraph = match block {
        Block::Paragraph(paragraph) => paragraph,
        Block::Caption(caption) => &mut caption.paragraph,
        _ => return,
    };
    if paragraph.runs.is_empty() {
        paragraph.style.horizontal_rule = Some(Box::new(*rule));
    } else {
        tracing::debug!("horizontal rule shares its paragraph with text; drawing none");
    }
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
/// first text fragment, or rule line. A hard layout break does not begin a new
/// paragraph, so its continuation must not repeat those edges or its
/// first-line indent.
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
            Block::Paragraph(paragraph) => has_line_content(paragraph),
            Block::Caption(caption) => has_line_content(&caption.paragraph),
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
        let has_content: bool = content_fragments.contains(&index);
        let style: &mut ParagraphStyle = match &mut blocks[index] {
            Block::Paragraph(paragraph) => &mut paragraph.style,
            Block::Caption(caption) => &mut caption.paragraph.style,
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

/// Whether a paragraph fragment lays out a line of its own: its text, or a
/// horizontal rule standing alone in its stretch (issue #2251).
fn has_line_content(paragraph: &Paragraph) -> bool {
    !paragraph.runs.is_empty() || paragraph.style.horizontal_rule.is_some()
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

/// Whether an empty paragraph carrying a `w:sectPr` (a section mark), at body
/// index `index` of a section whose first child is at `section_start`, is
/// laid out as a line. Word decides by the mark's place in its section, not
/// by its `w:pPr`, following docxide-pdf's Word-probed rule (issue #2232):
///
/// - The first block of its section keeps its line, as Word's export of
///   docxide-pdf's `page-header-ordinary-marker` shows for a bare mark.
/// - After other blocks it takes no height: `paragraph-offset-frame-no-frame`
///   drops a 14pt mark with 4pt after, and `section-mark-after-table-float-*`
///   a 1.5-line one after a floating table.
/// - In compatibility mode 15, a mark right after an inline table keeps its
///   line when the next section starts continuously
///   (`section-mark-after-table-inline-compat15`); mode 14 drops it.
///
/// `last_output` is the converted output of the body child before the mark.
///
/// A mark that takes no line before a continuous section still has a say in
/// the gap: see [`pass_on_section_mark_after`].
fn empty_section_mark_keeps_line(
    body: &[docx_rs::DocumentChild],
    section_start: usize,
    index: usize,
    last_output: Option<&TaggedElement>,
    compatibility_mode: WordCompatibilityMode,
    final_section: &docx_rs::SectionProperty,
) -> bool {
    let Some(previous) = body[section_start..index]
        .iter()
        .rev()
        .find(|child| is_body_block(child))
    else {
        return true;
    };
    compatibility_mode == WordCompatibilityMode::Word2013OrLater
        && matches!(previous, docx_rs::DocumentChild::Table(_))
        && matches!(
            last_output,
            Some(TaggedElement::Plain(blocks)) if matches!(blocks.last(), Some(Block::Table(_)))
        )
        && next_section_starts_continuously(body, index, final_section)
}

/// A body child Word lays out as a block of its section; bookmark and
/// comment markers take no place.
fn is_body_block(child: &docx_rs::DocumentChild) -> bool {
    matches!(
        child,
        docx_rs::DocumentChild::Paragraph(_)
            | docx_rs::DocumentChild::Table(_)
            | docx_rs::DocumentChild::StructuredDataTag(_)
            | docx_rs::DocumentChild::TableOfContents(_)
    )
}

/// Whether the section after the one the mark at body index `index` closes
/// starts on the same page. That is the next section's own `w:type`, on the
/// next section mark or else on the body's closing `w:sectPr`.
fn next_section_starts_continuously(
    body: &[docx_rs::DocumentChild],
    index: usize,
    final_section: &docx_rs::SectionProperty,
) -> bool {
    let next_section: &docx_rs::SectionProperty = body[index + 1..]
        .iter()
        .find_map(|child| match child {
            docx_rs::DocumentChild::Paragraph(paragraph) => {
                paragraph.property.section_property.as_ref()
            }
            _ => None,
        })
        .unwrap_or(final_section);
    next_section.section_type == Some(docx_rs::SectionType::Continuous)
}

/// The `w:after` of a converted blank paragraph, as
/// [`is_blank_paragraph_output`] accepts it; an unstated one is zero.
fn blank_paragraph_space_after(converted: &[TaggedElement]) -> Option<f64> {
    match converted {
        [TaggedElement::Plain(blocks)] => match blocks.as_slice() {
            [Block::Paragraph(paragraph)] => Some(paragraph.style.space_after.unwrap_or(0.0)),
            _ => None,
        },
        _ => None,
    }
}

/// Whether the converted output ends at a page or column break, so a section
/// mark after it opens the next section at the top of a page or column.
fn ends_with_layout_break(output: Option<&TaggedElement>) -> bool {
    matches!(
        output,
        Some(TaggedElement::Plain(blocks))
            if matches!(blocks.last(), Some(Block::PageBreak | Block::ColumnBreak))
    )
}

/// Hand the `w:after` of an empty section mark that took no line to the first
/// paragraph of `converted`, the first block of the continuous section after
/// it. Word lays that paragraph out after the previous paragraph's whole
/// `w:after`, plus only the part of its own `w:before` beyond the mark's
/// `w:after`: `after + max(0, before − mark after)`, not the larger of the two
/// (docxide-pdf's `section_break_spacing`, from Word probes; Word's export of
/// docxide-pdf's `scraped/covid_insomnia_mental_health` starts its numbered
/// heading 12 + 10pt below the paragraph before its 0pt mark).
fn pass_on_section_mark_after(converted: &mut [TaggedElement], mark_after: f64) {
    let style: Option<&mut ParagraphStyle> = match converted.first_mut() {
        Some(TaggedElement::Plain(blocks)) => blocks.iter_mut().find_map(|block| match block {
            Block::Paragraph(paragraph) => Some(&mut paragraph.style),
            Block::Caption(caption) => Some(&mut caption.paragraph.style),
            _ => None,
        }),
        Some(TaggedElement::ListParagraph { paragraph, .. }) => Some(&mut paragraph.style),
        None => None,
    };
    if let Some(style) = style {
        style.section_mark_after = Some(mark_after);
    }
}

/// Whether a body paragraph converted to one paragraph that paints nothing:
/// no text, link, note or inline box, and no border or shading.
fn is_blank_paragraph_output(converted: &[TaggedElement]) -> bool {
    matches!(
        converted,
        [TaggedElement::Plain(blocks)] if matches!(
            blocks.as_slice(),
            [Block::Paragraph(paragraph)] if paragraph.runs.iter().all(|run| {
                run.text.is_empty() && run_shows_only_spaces(run)
            }) && paragraph.style.background.is_none()
                && paragraph.style.border.is_none()
        )
    )
}

/// A body paragraph in a `w:framePr` frame that Word lifts out of the flow,
/// whose converted paragraph is undecorated and has no text but spaces, in a
/// frame that shows nothing. Adjacent paragraphs with an identical
/// `w:framePr` share one frame, as does a table whose cell paragraphs carry
/// it, so the frame shows nothing when none of them does. Such a paragraph
/// leaves no line in the flow: Word's exports of docxide-pdf's
/// `empty-frame-control` and `cont-first-first-fits` advance the next text as
/// if it were absent (issue #2161), and those of `empty-margin-frame-14-*` do
/// the same for groups of two and three (issue #2231). An empty paragraph in a
/// frame with text or a table is a line of the frame's own content, so it
/// keeps it.
///
/// The other paragraphs of the frame are judged on their source runs, which
/// cannot tell a paragraph style's border or shading.
///
/// `frames` holds each body child's [`leading_frame_property`].
fn is_empty_lifted_frame(
    converted: &[TaggedElement],
    body: &[docx_rs::DocumentChild],
    frames: &[Option<docx_rs::FrameProperty>],
    index: usize,
) -> bool {
    let Some(frame) = frames[index].as_ref() else {
        return false;
    };
    let shares_frame = |mate: &usize| frames[*mate].as_ref() == Some(frame);
    let mut frame_mates = (0..index)
        .rev()
        .take_while(shares_frame)
        .chain((index + 1..body.len()).take_while(shares_frame))
        .map(|mate| &body[mate]);
    frame_lifts_out_of_flow(frame)
        && matches!(
            converted,
            [TaggedElement::Plain(blocks)] if matches!(
                blocks.as_slice(),
                [Block::Paragraph(paragraph)] if paragraph
                    .runs
                    .iter()
                    .all(run_shows_only_spaces)
                    && paragraph.style.background.is_none()
                    && paragraph.style.border.is_none()
            )
        )
        && frame_mates.all(|child| {
            matches!(child, docx_rs::DocumentChild::Paragraph(mate) if shows_only_spaces(mate))
        })
}

/// Keep an empty framed paragraph's line in the box Word 2013 and later keep
/// for its frame at the top margin ([`top_margin_frame_box`]), returning the
/// box when this paragraph opens it. The frame's first paragraph with a line
/// opens the box, and its later lines join it. A paragraph whose mark is
/// hidden or that ends its section adds no line: Word's export of
/// docxide-pdf's `empty-margin-frame-hidden-marks` keeps no box, and that of
/// `empty-margin-frame-15-one-empty` keeps one line for an empty paragraph
/// and the section mark after it (issue #2231).
///
/// `open_box` is the body index the frame of the box opened last starts at.
///
/// FIXME(docx-rs reads no `w:val` on `w:vanish`): a paragraph mark stating
/// `w:vanish w:val="0"` counts as hidden.
fn keep_line_in_top_margin_frame_box(
    elements: &mut [TaggedElement],
    open_box: &mut Option<usize>,
    para: &docx_rs::Paragraph,
    converted: Vec<TaggedElement>,
    frames: &[Option<docx_rs::FrameProperty>],
    index: usize,
) -> Vec<TaggedElement> {
    let Some(frame_property) = frames[index].as_ref() else {
        return Vec::new();
    };
    if para.property.section_property.is_some() || para.property.run_property.vanish.is_some() {
        return Vec::new();
    }
    let Some(mut frame) = top_margin_frame_box(frame_property) else {
        return Vec::new();
    };
    let lines: Vec<Block> = converted
        .into_iter()
        .flat_map(|element| match element {
            TaggedElement::Plain(blocks) => blocks,
            TaggedElement::ListParagraph { .. } => Vec::new(),
        })
        .collect();
    let frame_start: usize = frame_start(frames, index);
    if *open_box == Some(frame_start)
        && let Some(TaggedElement::Plain(blocks)) = elements.last_mut()
        && let [Block::TextFrame(open)] = blocks.as_mut_slice()
    {
        open.content.extend(lines);
        return Vec::new();
    }
    tracing::debug!(
        format = "DOCX",
        body_index = index,
        x_pt = frame.x,
        width_pt = frame.width,
        "Empty frame keeps its box at the top margin"
    );
    frame.content = lines;
    *open_box = Some(frame_start);
    vec![TaggedElement::Plain(vec![Block::TextFrame(Box::new(
        frame,
    ))])]
}

fn run_shows_only_spaces(run: &Run) -> bool {
    run.text.bytes().all(|byte| byte == b' ')
        && run.href.is_none()
        && run.footnote.is_none()
        && run.inline_box.is_none()
}

/// Whether a source paragraph holds nothing but spaces: runs of space-only
/// text and zero-width bookmark and comment markers. A tab, a break, a
/// picture, a field or a link is content.
fn shows_only_spaces(para: &docx_rs::Paragraph) -> bool {
    para.children.iter().all(|child| match child {
        docx_rs::ParagraphChild::Run(run) => run.children.iter().all(|run_child| {
            matches!(run_child, docx_rs::RunChild::Text(text) if text.text.bytes().all(|byte| byte == b' '))
        }),
        docx_rs::ParagraphChild::BookmarkStart(_)
        | docx_rs::ParagraphChild::BookmarkEnd(_)
        | docx_rs::ParagraphChild::CommentStart(_)
        | docx_rs::ParagraphChild::CommentEnd(_) => true,
        _ => false,
    })
}

/// A body paragraph alone in a `w:framePr` frame hung below its own place in
/// the flow, as [`body_text_frame`] describes, carrying the paragraph's
/// converted blocks. Adjacent paragraphs, or a table, that share an identical
/// frame stay in the flow, as does a list paragraph or one holding a page or
/// column break, which a placed frame cannot hold.
///
/// TODO(frames shared by several paragraphs): Word draws a group of
/// paragraphs with an identical `w:framePr` in one frame.
fn lifted_text_frame(
    converted: &[TaggedElement],
    frames: &[Option<docx_rs::FrameProperty>],
    index: usize,
) -> Option<TextFrame> {
    let frame_property: &docx_rs::FrameProperty = frames[index].as_ref()?;
    let mut frame: TextFrame = body_text_frame(frame_property)?;
    let shares_frame = |neighbour: Option<&Option<docx_rs::FrameProperty>>| {
        neighbour.and_then(Option::as_ref) == Some(frame_property)
    };
    if shares_frame(
        index
            .checked_sub(1)
            .and_then(|previous| frames.get(previous)),
    ) || shares_frame(frames.get(index + 1))
    {
        return None;
    }
    for element in converted {
        let TaggedElement::Plain(blocks) = element else {
            return None;
        };
        if blocks
            .iter()
            .any(|block| matches!(block, Block::PageBreak | Block::ColumnBreak))
        {
            return None;
        }
        frame.content.extend(blocks.iter().cloned());
    }
    (!frame.content.is_empty()).then_some(frame)
}

/// The run of adjacent body children that share a frame Word keeps at a
/// fixed place on its page, as [`join_fixed_text_frame`] last saw it.
struct FixedFrameGroup {
    /// The body index the run starts at.
    start: usize,
    /// Whether the run is drawn in its frame, or stays in the flow because one
    /// of its children holds what a placed frame cannot.
    lifted: bool,
}

/// Draw a body child in the `w:framePr` frame Word keeps at a fixed place on
/// its page ([`fixed_text_frame`]), together with the adjacent children that
/// share the frame: the first opens it, and the others join it while it is
/// still the last element. Returns what the flow keeps: the frame when this
/// child opens it, nothing when it joins one, and the child's own blocks when
/// its frame stays in the flow.
///
/// A frame whose children show nothing but spaces stays in the flow, each
/// empty paragraph keeping its line: no Word reference shows such a frame
/// that keeps text off its sides leaving no line, as [`frame_lifts_out_of_flow`]
/// notes.
///
/// A frame holds paragraphs, tables and inline pictures. A child holding
/// anything else — a list paragraph, a page or column break, a floating or
/// grouped drawing — keeps the whole frame in the flow, its earlier children
/// included, as before frames were lifted: a drawing group's zero-size
/// anchor boxes spread across a placed block's line instead of stacking at
/// their anchor.
///
/// TODO(lists and drawings in a fixed frame): Word draws them in the frame
/// too.
///
/// `frames` holds each body child's [`leading_frame_property`].
fn join_fixed_text_frame(
    elements: &mut Vec<TaggedElement>,
    group: &mut Option<FixedFrameGroup>,
    mut frame: TextFrame,
    converted: Vec<TaggedElement>,
    body: &[docx_rs::DocumentChild],
    frames: &[Option<docx_rs::FrameProperty>],
    index: usize,
) -> Vec<TaggedElement> {
    let start: usize = frame_start(frames, index);
    let fits_in_frame: bool = converted.iter().all(|element| {
        matches!(element, TaggedElement::Plain(blocks) if blocks.iter().all(|block| matches!(
            block,
            Block::Paragraph(_)
                | Block::Caption(_)
                | Block::Table(_)
                | Block::Image(_)
                | Block::InlineImages(_)
        )))
    });
    if let Some(open) = group.as_mut()
        && open.start == start
    {
        if !open.lifted {
            return converted;
        }
        let open_frame: Option<&mut TextFrame> = match elements.last_mut() {
            Some(TaggedElement::Plain(blocks)) => match blocks.as_mut_slice() {
                [Block::TextFrame(open_frame)] => Some(open_frame),
                _ => None,
            },
            _ => None,
        };
        if let Some(open_frame) = open_frame {
            if fits_in_frame {
                open_frame
                    .content
                    .extend(converted.into_iter().flat_map(plain_blocks));
                return Vec::new();
            }
            let earlier: Vec<Block> = std::mem::take(&mut open_frame.content);
            elements.pop();
            elements.push(TaggedElement::Plain(earlier));
            open.lifted = false;
            return converted;
        }
    }
    let frame_shows_something: bool = (start..body.len())
        .take_while(|&mate| frames[mate] == frames[start])
        .any(|mate| match &body[mate] {
            docx_rs::DocumentChild::Paragraph(paragraph) => !shows_only_spaces(paragraph),
            _ => true,
        });
    if !fits_in_frame || !frame_shows_something {
        *group = Some(FixedFrameGroup {
            start,
            lifted: false,
        });
        return converted;
    }
    tracing::debug!(
        format = "DOCX",
        body_index = index,
        x_pt = frame.x,
        y_pt = frame.y,
        width_pt = frame.width,
        "Framed body content drawn at its place on the page"
    );
    frame.content = converted.into_iter().flat_map(plain_blocks).collect();
    *group = Some(FixedFrameGroup {
        start,
        lifted: true,
    });
    vec![TaggedElement::Plain(vec![Block::TextFrame(Box::new(
        frame,
    ))])]
}

/// A plain element's blocks. A list paragraph has none here:
/// [`join_fixed_text_frame`] keeps those in the flow.
fn plain_blocks(element: TaggedElement) -> Vec<Block> {
    match element {
        TaggedElement::Plain(blocks) => blocks,
        TaggedElement::ListParagraph { .. } => Vec::new(),
    }
}

/// The body index at which the run of children sharing the frame of the
/// child at `index` starts.
fn frame_start(frames: &[Option<docx_rs::FrameProperty>], index: usize) -> usize {
    index
        - frames[..index]
            .iter()
            .rev()
            .take_while(|frame| *frame == &frames[index])
            .count()
}

/// The `w:framePr` a body child would sit in: a paragraph's, or a table's
/// first cell paragraph's, which is how Word frames a table, merged over its
/// paragraph style's ([`effective_frame_property`]). A paragraph that names
/// no style takes the default paragraph style's, `default_style_id`.
///
/// TODO(frames in block-level content controls): a framed paragraph inside a
/// body `w:sdt` is not a body child here, so it stays in the flow; bulk
/// `tdf148671` keeps three of its six page frames there.
fn leading_frame_property(
    child: &docx_rs::DocumentChild,
    style_frames: &HashMap<String, docx_rs::FrameProperty>,
    default_style_id: Option<&str>,
) -> Option<docx_rs::FrameProperty> {
    let paragraph: &docx_rs::Paragraph = match child {
        docx_rs::DocumentChild::Paragraph(paragraph) => paragraph,
        docx_rs::DocumentChild::Table(table) => {
            let docx_rs::TableChild::TableRow(row) = table.rows.first()?;
            let docx_rs::TableRowChild::TableCell(cell) = row.cells.first()?;
            match cell.children.first()? {
                docx_rs::TableCellContent::Paragraph(paragraph) => paragraph,
                _ => return None,
            }
        }
        _ => return None,
    };
    let style_id: Option<&str> = get_paragraph_style_id(&paragraph.property).or(default_style_id);
    effective_frame_property(
        &paragraph.property,
        style_id.and_then(|style_id| style_frames.get(style_id)),
    )
}

/// Whether Word lays an empty paragraph's `w:framePr` frame outside the body
/// flow, following docxide-pdf's Word-probed `is_empty_wrapping_frame`. A
/// frame stating only its `w:wrap` flows as a plain paragraph, and an inline
/// one sits in the flow. One that keeps text off its sides (`none`,
/// `notBeside`) or states its own height is not judged here: no reference
/// shows those leaving no line, and [`fixed_text_frame`] draws those it
/// places.
fn frame_lifts_out_of_flow(frame: &docx_rs::FrameProperty) -> bool {
    let placement = docx_rs::FrameProperty {
        wrap: None,
        ..frame.clone()
    };
    placement != docx_rs::FrameProperty::default()
        && frame.y_align.as_deref() != Some("inline")
        && !matches!(frame.wrap.as_deref(), Some("none" | "notBeside"))
        && frame.h.is_none_or(|height| height == 0)
}

fn push_inline_images(
    out: &mut Vec<Block>,
    inline_images: &mut Vec<Block>,
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    contextual_spacing: ParagraphContextualSpacing<'_>,
) {
    if inline_images.is_empty() {
        return;
    }
    let alignment: Option<Alignment> = paragraph_alignment(para);
    let style: ParagraphStyle = merge_paragraph_style(
        &extract_paragraph_style(&para.property),
        None,
        resolved_style,
    );
    let spacing: Option<ImageParagraphSpacing> =
        paragraph_image_spacing(&style, contextual_spacing);
    let indent: Option<ImageParagraphIndent> = paragraph_image_indent(&style);
    let mut grouped: Vec<ImageData> = Vec::new();

    for block in inline_images.drain(..) {
        match block {
            Block::Image(mut image) => {
                // Inline images inherit the containing paragraph's alignment,
                // `w:spacing` and `w:ind`: the picture consumes the paragraph,
                // so the gaps Word draws around it (issue #499) and the indents
                // it starts from (issue #2201) have to travel with the picture
                // instead.
                if image.alignment.is_none() {
                    image.alignment = alignment;
                }
                if image.paragraph_spacing.is_none() {
                    image.paragraph_spacing = spacing;
                }
                if image.paragraph_indent.is_none() {
                    image.paragraph_indent = indent;
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
/// Read from the same style merge a text paragraph uses, so spacing
/// inherited from `styles.xml` counts as well as direct formatting.
fn paragraph_image_spacing(
    style: &ParagraphStyle,
    contextual_spacing: ParagraphContextualSpacing<'_>,
) -> Option<ImageParagraphSpacing> {
    let mut spacing = ImageParagraphSpacing {
        before: style.space_before,
        after: style.space_after,
    };
    contextual_spacing.apply(&mut spacing.before, &mut spacing.after);
    (spacing != ImageParagraphSpacing::default()).then_some(spacing)
}

/// The `w:ind` a picture paragraph places its pictures by, from the same
/// style merge as [`paragraph_image_spacing`]: a text paragraph's indents,
/// because Word lays an inline picture out like a glyph (issue #2201).
fn paragraph_image_indent(style: &ParagraphStyle) -> Option<ImageParagraphIndent> {
    let indent = ImageParagraphIndent {
        left: style.indent_left.unwrap_or(0.0),
        right: style.indent_right.unwrap_or(0.0),
        first_line: style.indent_first_line.unwrap_or(0.0),
    };
    (indent != ImageParagraphIndent::default()).then_some(indent)
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
    // Before anything reads the runs' text, which must see only `'\n'`.
    style.clearing_breaks = take_clearing_break_markers(runs);
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

#[cfg(test)]
#[path = "docx_alt_chunk_tests.rs"]
mod alt_chunk_tests;
