use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};

use crate::error::ConvertWarning;
use crate::ir::{
    Block, BorderLineStyle, BorderSide, CellBorder, Color, ColumnLayout, FlowPage, FrameAnchor,
    HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph, Insets, LineJoin, Margins,
    PageNumbering, PageSize, PositionedTab, PositionedTabAlignment, PositionedTabRelativeTo, Run,
    TabLeader, TextDirection, TextStyle,
};

use super::contexts::WrapContext;
use super::media::extract_drawing_image;
use super::{
    DOC_DEFAULT_STYLE_ID, ImageMap, NumberingMap, ParagraphItem, ResolvedStyle, StyleMap,
    TaggedElement, extract_column_layout_from_section_property, extract_paragraph_style,
    extract_run_style, extract_tab_stop_overrides, flatten_tracked_changes, get_paragraph_style_id,
    group_into_lists, merge_paragraph_style, merge_text_style, read_zip_text,
    word_compatible_paragraph_space_after_pt,
};
use crate::parser::units::twips_to_pt;
use crate::parser::xml_util::parse_hex_color;

/// Parsed header/footer assets addressed by relationship ID.
#[derive(Default)]
pub(super) struct HeaderFooterAssets {
    headers: HashMap<String, HeaderFooter>,
    footers: HashMap<String, HeaderFooter>,
}

/// What a header or footer paragraph resolves its unstated paragraph
/// properties through.
///
/// A story's paragraphs take the same cascade the body's do, so the style map
/// has to be built — which needs the docx-rs parse — before the parts are
/// converted (issue #1195).
#[derive(Clone, Copy)]
pub(super) struct HeaderFooterStyleContext<'a> {
    pub(super) style_map: &'a StyleMap,
    /// Whether `word/styles.xml` declares `w:docDefaults/w:pPrDefault`, which
    /// decides the `w:spacing w:after` an unstated gap falls back to
    /// (issue #1085).
    pub(super) paragraph_property_defaults_are_declared: bool,
}

#[derive(Clone, Copy)]
enum SimpleFieldKind {
    PageNumber,
    TotalPages,
}

#[derive(Clone, Copy)]
struct SimpleFieldMarker {
    preceding_runs: usize,
    cached_runs: usize,
    kind: SimpleFieldKind,
}

fn scan_header_footer_relationships(
    rels_xml: &str,
) -> (HashMap<String, String>, HashMap<String, String>) {
    let mut headers: HashMap<String, String> = HashMap::new();
    let mut footers: HashMap<String, String> = HashMap::new();

    for entry in crate::parser::xml_util::parse_relationships(rels_xml) {
        let Some(relationship_type) = entry.rel_type else {
            continue;
        };

        let full_path = if let Some(stripped) = entry.target.strip_prefix('/') {
            stripped.to_string()
        } else {
            format!("word/{}", entry.target)
        };

        if relationship_type.ends_with("/header") {
            headers.insert(entry.id, full_path);
        } else if relationship_type.ends_with("/footer") {
            footers.insert(entry.id, full_path);
        }
    }

    (headers, footers)
}

pub(super) fn build_header_footer_assets<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    styles: HeaderFooterStyleContext<'_>,
) -> HeaderFooterAssets {
    let rels_xml = match read_zip_text(archive, "word/_rels/document.xml.rels") {
        Some(xml) => xml,
        None => return HeaderFooterAssets::default(),
    };
    let (header_relationships, footer_relationships) = scan_header_footer_relationships(&rels_xml);
    // A header part declares no theme of its own, so its anchored shapes'
    // `<a:schemeClr>` fills borrow the document's (issue #961).
    let theme_colors: HashMap<String, Color> = read_zip_text(archive, "word/theme/theme1.xml")
        .as_deref()
        .map(crate::parser::drawingml::parse_theme_color_scheme)
        .unwrap_or_default();
    let mut assets = HeaderFooterAssets::default();

    for (relationship_id, path) in header_relationships {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let images = build_part_image_map(archive, &path);
        let simple_fields = scan_simple_fields(&xml);
        let Ok(header) = <docx_rs::Header as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        let anchors = scan_hf_anchors(&xml, &theme_colors);
        if let Some(converted) =
            convert_docx_header(&header, &images, &simple_fields, &anchors, styles)
        {
            assets.headers.insert(relationship_id, converted);
        }
    }

    for (relationship_id, path) in footer_relationships {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let images = build_part_image_map(archive, &path);
        let bidi_paragraphs = scan_bidi_paragraphs(&xml);
        let simple_fields = scan_simple_fields(&xml);
        let Ok(footer) = <docx_rs::Footer as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        let anchors = scan_hf_anchors(&xml, &theme_colors);
        if let Some(converted) = convert_docx_footer(
            &footer,
            &images,
            &bidi_paragraphs,
            &simple_fields,
            &anchors,
            styles,
        ) {
            assets.footers.insert(relationship_id, converted);
        }
    }

    assets
}

fn scan_simple_fields(xml: &str) -> Vec<Vec<SimpleFieldMarker>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut paragraphs: Vec<Vec<SimpleFieldMarker>> = Vec::new();
    let mut paragraph_depth: usize = 0;
    let mut simple_field_depth: usize = 0;
    let mut direct_run_count: usize = 0;
    let mut fields: Vec<SimpleFieldMarker> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element)) => {
                match element.local_name().as_ref() {
                    b"p" => {
                        paragraph_depth += 1;
                        if paragraph_depth == 1 {
                            direct_run_count = 0;
                            fields.clear();
                        }
                    }
                    b"fldSimple" if paragraph_depth == 1 => {
                        if let Some(kind) = simple_field_kind(element) {
                            fields.push(SimpleFieldMarker {
                                preceding_runs: direct_run_count,
                                cached_runs: 0,
                                kind,
                            });
                        }
                        simple_field_depth += 1;
                    }
                    b"r" if paragraph_depth == 1 && simple_field_depth == 0 => {
                        direct_run_count += 1;
                    }
                    b"r" if paragraph_depth == 1 && simple_field_depth > 0 => {
                        if let Some(field) = fields.last_mut() {
                            field.cached_runs += 1;
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref element)) => {
                if element.local_name().as_ref() == b"fldSimple"
                    && paragraph_depth == 1
                    && let Some(kind) = simple_field_kind(element)
                {
                    fields.push(SimpleFieldMarker {
                        preceding_runs: direct_run_count,
                        cached_runs: 0,
                        kind,
                    });
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"fldSimple" if paragraph_depth == 1 => {
                    simple_field_depth = simple_field_depth.saturating_sub(1);
                }
                b"p" if paragraph_depth > 0 => {
                    if paragraph_depth == 1 {
                        paragraphs.push(std::mem::take(&mut fields));
                    }
                    paragraph_depth -= 1;
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    paragraphs
}

fn simple_field_kind(element: &quick_xml::events::BytesStart<'_>) -> Option<SimpleFieldKind> {
    let instruction = element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == b"instr")?
        .unescape_value()
        .ok()?;
    let field_name = instruction.split_whitespace().next()?;
    if field_name.eq_ignore_ascii_case("page") {
        Some(SimpleFieldKind::PageNumber)
    } else if field_name.eq_ignore_ascii_case("numpages") {
        Some(SimpleFieldKind::TotalPages)
    } else {
        None
    }
}

fn scan_bidi_paragraphs(xml: &str) -> Vec<bool> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut paragraphs: Vec<bool> = Vec::new();
    let mut paragraph_depth: usize = 0;
    let mut is_bidi: bool = false;
    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"p" => {
                    paragraph_depth += 1;
                    if paragraph_depth == 1 {
                        is_bidi = false;
                    }
                }
                b"bidi" if paragraph_depth == 1 => is_bidi = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element))
                if paragraph_depth == 1 && element.local_name().as_ref() == b"bidi" =>
            {
                is_bidi = true;
            }
            Ok(quick_xml::events::Event::End(ref element))
                if element.local_name().as_ref() == b"p" && paragraph_depth > 0 =>
            {
                if paragraph_depth == 1 {
                    paragraphs.push(is_bidi);
                }
                paragraph_depth -= 1;
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    paragraphs
}

fn build_part_image_map<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    part_path: &str,
) -> ImageMap {
    let Some((directory, filename)) = part_path.rsplit_once('/') else {
        return ImageMap::new();
    };
    let relationships_path = format!("{directory}/_rels/{filename}.rels");
    let Some(relationships_xml) = read_zip_text(archive, &relationships_path) else {
        return ImageMap::new();
    };
    let mut relationships: Vec<(String, String)> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(&relationships_xml);
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
                    let Ok(value) = attribute.unescape_value() else {
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
                    relationships.push((id, resolve_part_target(directory, &target)));
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    relationships
        .into_iter()
        .filter_map(|(id, path)| {
            let mut bytes: Vec<u8> = Vec::new();
            archive.by_name(&path).ok()?.read_to_end(&mut bytes).ok()?;
            let image = image::load_from_memory(&bytes).ok()?;
            let mut png = Cursor::new(Vec::new());
            image.write_to(&mut png, image::ImageFormat::Png).ok()?;
            Some((
                id,
                super::DocxImageAsset {
                    data: png.into_inner(),
                    format: crate::ir::ImageFormat::Png,
                },
            ))
        })
        .collect()
}

fn resolve_part_target(directory: &str, target: &str) -> String {
    let mut parts: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        directory
            .split('/')
            .filter(|part| !part.is_empty())
            .collect()
    };
    for part in target.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

/// The per-section values scanned out of `document.xml` directly.
///
/// Both are things docx-rs's `SectionProperty` does not carry in the form the
/// IR needs: the column layout it reports without the separator, and the page
/// numbering it reports without `w:fmt`.
#[derive(Debug, Clone, Default)]
pub(super) struct SectionOverrides {
    pub(super) column_layout: Option<ColumnLayout>,
    pub(super) page_numbering: Option<PageNumbering>,
}

pub(super) fn build_flow_page_from_section(
    section_prop: &docx_rs::SectionProperty,
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
    header_footer_assets: &HeaderFooterAssets,
    overrides: SectionOverrides,
    styles: HeaderFooterStyleContext<'_>,
    warnings: &mut Vec<ConvertWarning>,
) -> FlowPage {
    let (size, margins) = extract_page_setup(section_prop);
    let doc_default_style: Option<&ResolvedStyle> = styles.style_map.get(DOC_DEFAULT_STYLE_ID);
    let content = group_into_lists(elements, numberings);

    for block in &content {
        if let Block::Chart(chart) = block {
            let title = chart.title.as_deref().unwrap_or("untitled").to_string();
            warnings.push(ConvertWarning::FallbackUsed {
                format: "DOCX".to_string(),
                from: format!("chart ({title})"),
                to: "data table".to_string(),
            });
        }
    }

    if matches!(
        section_prop.section_type,
        Some(docx_rs::SectionType::Continuous | docx_rs::SectionType::NextColumn)
    ) {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "continuous section break".to_string(),
            to: "page-level section split".to_string(),
        });
    }

    // A `first` variant is honoured now, so only the `even` ones still collapse
    // onto the default. A first variant without `w:titlePg` is not a variant at
    // all — Word ignores it — so it does not warn either (issue #846).
    if section_prop.even_header_reference.is_some()
        || section_prop.even_footer_reference.is_some()
        || section_prop.even_header.is_some()
        || section_prop.even_footer.is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "even-page header/footer".to_string(),
            to: "the section's default header/footer".to_string(),
        });
    }

    let mut header = extract_docx_header(section_prop, header_footer_assets, styles);
    if let Some(header) = &mut header {
        header.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.header));
        apply_doc_default_text_style(header, doc_default_style);
    }
    let mut footer = extract_docx_footer(section_prop, header_footer_assets, styles);
    if let Some(footer) = &mut footer {
        footer.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.footer));
        apply_doc_default_text_style(footer, doc_default_style);
    }
    // The first-page stories take the same edge distance and default style the
    // whole-section ones do; only which story is chosen differs (issue #846).
    let mut first_header = extract_docx_first_header(section_prop, header_footer_assets, styles);
    if let Some(first_header) = &mut first_header {
        first_header.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.header));
        apply_doc_default_text_style(first_header, doc_default_style);
    }
    let mut first_footer = extract_docx_first_footer(section_prop, header_footer_assets, styles);
    if let Some(first_footer) = &mut first_footer {
        first_footer.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.footer));
        apply_doc_default_text_style(first_footer, doc_default_style);
    }

    FlowPage {
        size,
        margins,
        content,
        header,
        footer,
        first_header,
        first_footer,
        page_numbering: overrides.page_numbering,
        columns: overrides
            .column_layout
            .or_else(|| extract_column_layout_from_section_property(section_prop)),
        line_grid_pitch: extract_line_grid_pitch(section_prop),
        line_grid_snaps_lines: line_grid_snaps_lines(section_prop),
    }
}

/// The section's document-grid line pitch in points (`<w:docGrid
/// w:linePitch>`, in twips). docx-rs keeps the fields private, so read them
/// through the type's serde representation.
fn extract_line_grid_pitch(section_prop: &docx_rs::SectionProperty) -> Option<f64> {
    let grid = section_prop.doc_grid.as_ref()?;
    let value = serde_json::to_value(grid).ok()?;
    let pitch_twips = value.get("linePitch")?.as_f64()?;
    (pitch_twips > 0.0).then(|| twips_to_pt(pitch_twips as i32))
}

/// Whether the section's grid snaps body lines to that pitch.
///
/// `w:type` decides, and its default value — what an omitted attribute means —
/// is `default`, which is ECMA-376's name for *no* grid. Word writes a bare
/// `<w:docGrid w:linePitch="360"/>` into ordinary Korean documents and then
/// lays them out with no grid at all: every Korean fixture in the business
/// corpus carries that element, and none of their line advances is a multiple
/// of 18pt (issue #518).
fn line_grid_snaps_lines(section_prop: &docx_rs::SectionProperty) -> bool {
    let Some(grid) = section_prop.doc_grid.as_ref() else {
        return false;
    };
    let Ok(value) = serde_json::to_value(grid) else {
        return false;
    };
    matches!(
        value.get("gridType").and_then(serde_json::Value::as_str),
        Some("lines" | "linesAndChars" | "snapToChars")
    )
}

fn convert_docx_header(
    header: &docx_rs::Header,
    images: &ImageMap,
    simple_fields: &[Vec<SimpleFieldMarker>],
    anchors: &[HfAnchorBox],
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    let shapes = hf_anchored_shapes(anchors);
    let mut anchors = anchors.iter();
    let paragraphs = header
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::HeaderChild::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .enumerate()
        .flat_map(|(index, paragraph)| {
            let mut converted = vec![convert_hf_paragraph(
                paragraph,
                images,
                false,
                simple_fields.get(index).map(Vec::as_slice).unwrap_or(&[]),
                styles,
            )];
            converted.extend(hf_anchored_text_box_paragraphs(
                paragraph,
                &mut anchors,
                styles,
            ));
            converted
        })
        .collect::<Vec<_>>();
    // A story that draws only a decorative banner has no paragraph worth
    // keeping but is still not empty (issue #961).
    if paragraphs.is_empty() && shapes.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        distance_from_edge: None,
        sheet_print_scale: None,
        shapes,
    })
}

fn convert_docx_footer(
    footer: &docx_rs::Footer,
    images: &ImageMap,
    bidi_paragraphs: &[bool],
    simple_fields: &[Vec<SimpleFieldMarker>],
    anchors: &[HfAnchorBox],
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    let shapes = hf_anchored_shapes(anchors);
    let mut anchors = anchors.iter();
    let paragraphs = footer
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::FooterChild::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .enumerate()
        .flat_map(|(index, paragraph)| {
            let mut converted = vec![convert_hf_paragraph(
                paragraph,
                images,
                bidi_paragraphs.get(index).copied().unwrap_or(false),
                simple_fields.get(index).map(Vec::as_slice).unwrap_or(&[]),
                styles,
            )];
            converted.extend(hf_anchored_text_box_paragraphs(
                paragraph,
                &mut anchors,
                styles,
            ));
            converted
        })
        .collect::<Vec<_>>();
    // A story that draws only a decorative banner has no paragraph worth
    // keeping but is still not empty (issue #961).
    if paragraphs.is_empty() && shapes.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        distance_from_edge: None,
        sheet_print_scale: None,
        shapes,
    })
}

/// The `first` variant of a section's header, where `<w:titlePg/>` asks for one.
///
/// Unlike [`extract_docx_header`] this does not fall back to the other
/// variants: `w:titlePg` names the first-page story specifically, and standing
/// the default in for it would draw page one the same as the rest rather than
/// differently, which is the bug (issue #846).
pub(super) fn extract_docx_first_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    if !section_prop.title_pg {
        return None;
    }
    section_prop
        .first_header_reference
        .as_ref()
        .and_then(|reference| assets.headers.get(&reference.id).cloned())
        .or_else(|| {
            section_prop
                .first_header
                .as_ref()
                .and_then(|(_relationship_id, header)| {
                    convert_docx_header(header, &ImageMap::new(), &[], &[], styles)
                })
        })
}

/// The `first` variant of a section's footer, under the same rule.
pub(super) fn extract_docx_first_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    if !section_prop.title_pg {
        return None;
    }
    section_prop
        .first_footer_reference
        .as_ref()
        .and_then(|reference| assets.footers.get(&reference.id).cloned())
        .or_else(|| {
            section_prop
                .first_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| {
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[], &[], styles)
                })
        })
}

/// Extract the header for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    section_prop
        .header_reference
        .as_ref()
        .and_then(|reference| assets.headers.get(&reference.id).cloned())
        .or_else(|| {
            section_prop
                .header
                .as_ref()
                .and_then(|(_relationship_id, header)| {
                    convert_docx_header(header, &ImageMap::new(), &[], &[], styles)
                })
        })
        .or_else(|| {
            section_prop
                .first_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_header
                .as_ref()
                .and_then(|(_relationship_id, header)| {
                    convert_docx_header(header, &ImageMap::new(), &[], &[], styles)
                })
        })
        .or_else(|| {
            section_prop
                .even_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_header
                .as_ref()
                .and_then(|(_relationship_id, header)| {
                    convert_docx_header(header, &ImageMap::new(), &[], &[], styles)
                })
        })
}

/// Extract the footer for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
    styles: HeaderFooterStyleContext<'_>,
) -> Option<HeaderFooter> {
    section_prop
        .footer_reference
        .as_ref()
        .and_then(|reference| assets.footers.get(&reference.id).cloned())
        .or_else(|| {
            section_prop
                .footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| {
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[], &[], styles)
                })
        })
        .or_else(|| {
            section_prop
                .first_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| {
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[], &[], styles)
                })
        })
        .or_else(|| {
            section_prop
                .even_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| {
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[], &[], styles)
                })
        })
}

/// Resolve a header or footer's runs against `w:docDefaults/w:rPrDefault`.
///
/// Header and footer parts are read from the archive before the stylesheet is,
/// so their runs were left with only the properties they state themselves and
/// fell through to the renderer's own defaults for everything else — Libertinus
/// Serif at 11pt where the document says Calibri at 10pt. Word resolves them
/// through the same run cascade as the body: a run that names a colour and a
/// size still takes the document's family, and the run holding nothing but a
/// tab still takes its size (issue #578).
///
/// This runs at page-build time rather than at parse time because that is the
/// first point where the style map exists.
fn apply_doc_default_text_style(
    header_footer: &mut HeaderFooter,
    doc_default_style: Option<&ResolvedStyle>,
) {
    let Some(doc_default_style) = doc_default_style else {
        return;
    };
    for paragraph in &mut header_footer.paragraphs {
        for element in &mut paragraph.elements {
            match element {
                HFInline::Run(run) => {
                    run.style = merge_text_style(&run.style, Some(doc_default_style));
                }
                HFInline::PageNumber(style) | HFInline::TotalPages(style) => {
                    *style = merge_text_style(style, Some(doc_default_style));
                }
                HFInline::Image(_) | HFInline::PositionedTab(_) => {}
            }
        }
    }
}

/// The `w:spacing w:after` a header or footer paragraph resolves to, in points.
///
/// The same cascade a body paragraph takes: its own `w:pPr`, then the style it
/// names — or the document's default paragraph style where it names none — and
/// finally Word's fallback for a gap nothing states, which is the built-in
/// `Normal`'s 8pt until the package declares `w:pPrDefault` (issue #1085).
///
/// A footer needs it because Word keeps the last paragraph's gap between the
/// story's last line and `w:pgMar/@w:footer`; a header records it for the same
/// reason a body paragraph does, so the two stories resolve alike (issue #1195).
fn hf_space_after_pt(
    explicit_pt: Option<f64>,
    property: &docx_rs::ParagraphProperty,
    styles: HeaderFooterStyleContext<'_>,
) -> f64 {
    explicit_pt
        .or_else(|| {
            get_paragraph_style_id(property)
                .and_then(|style_id| styles.style_map.get(style_id))
                .or_else(|| styles.style_map.get(DOC_DEFAULT_STYLE_ID))
                .and_then(|resolved| resolved.paragraph.space_after)
        })
        .unwrap_or_else(|| {
            word_compatible_paragraph_space_after_pt(
                styles.paragraph_property_defaults_are_declared,
            )
        })
}

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE/NUMPAGES field codes within runs and emits page counter inlines.
fn convert_hf_paragraph(
    paragraph: &docx_rs::Paragraph,
    images: &ImageMap,
    is_bidi: bool,
    simple_fields: &[SimpleFieldMarker],
    styles: HeaderFooterStyleContext<'_>,
) -> HeaderFooterParagraph {
    let explicit_style = extract_paragraph_style(&paragraph.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&paragraph.property.tabs);
    let mut style = merge_paragraph_style(&explicit_style, explicit_tab_overrides.as_deref(), None);
    style.space_after = Some(hf_space_after_pt(
        style.space_after,
        &paragraph.property,
        styles,
    ));
    if is_bidi || paragraph.property.bidi == Some(true) {
        style.direction = Some(TextDirection::Rtl);
    }
    let mut elements: Vec<HFInline> = Vec::new();

    let mut processed_runs: usize = 0;
    let mut cached_runs_to_skip: usize = append_simple_fields(
        &mut elements,
        simple_fields,
        processed_runs,
        &TextStyle::default(),
    );
    // One field state for the whole paragraph: a w:fldChar field spans runs.
    let mut field_state = HeaderFieldState::default();
    for child in flatten_tracked_changes(&paragraph.children) {
        match child {
            ParagraphItem::Run(run) => {
                if cached_runs_to_skip > 0 {
                    cached_runs_to_skip -= 1;
                    continue;
                }
                let run_style = extract_run_style(&run.run_property);
                extract_hf_run_elements(&run.children, &run_style, &mut elements, &mut field_state);
                for run_child in &run.children {
                    if let docx_rs::RunChild::Drawing(drawing) = run_child
                        && let Some(block) =
                            extract_drawing_image(drawing, images, &WrapContext::empty(), None)
                    {
                        match block {
                            Block::Image(image) => elements.push(HFInline::Image(image)),
                            Block::FloatingImage(image) => {
                                elements.push(HFInline::Image(image.image));
                            }
                            _ => {}
                        }
                    }
                }
                processed_runs += 1;
                cached_runs_to_skip +=
                    append_simple_fields(&mut elements, simple_fields, processed_runs, &run_style);
            }
            ParagraphItem::PageNum => elements.push(HFInline::PageNumber(TextStyle::default())),
            ParagraphItem::NumPages => elements.push(HFInline::TotalPages(TextStyle::default())),
            // A header rarely carries a hyperlink, and when it does the URL
            // map lives on the body relationships rather than this part's.
            ParagraphItem::Hyperlink(_) => {}
        }
    }

    // A field the paragraph never closed. Per-run state used to make this
    // unreachable — the next run started clean — so leaving it unflushed would
    // silently drop text that malformed input previously rendered.
    if field_state.in_field {
        match field_state.inline.take() {
            Some(inline) => elements.push(inline),
            None => elements.extend(field_state.cached_result.drain(..).map(HFInline::Run)),
        }
    }

    HeaderFooterParagraph {
        style,
        elements,
        border: extract_hf_paragraph_border(&paragraph.property),
        border_space: extract_hf_paragraph_border_space(&paragraph.property),
        frame: extract_hf_frame(&paragraph.property),
    }
}

/// `w:pBdr` sides carry a `w:space` attribute in points that sets the gap Word
/// leaves between the paragraph text and the rule.
fn extract_hf_paragraph_border_space(property: &docx_rs::ParagraphProperty) -> Option<Insets> {
    let borders = serde_json::to_value(property.borders.as_ref()?).ok()?;
    let side_space = |key: &str| -> f64 {
        borders
            .get(key)
            .and_then(|side| side.get("space"))
            .and_then(serde_json::Value::as_f64)
            .unwrap_or(0.0)
    };
    let insets = Insets {
        top: side_space("top"),
        right: side_space("right"),
        bottom: side_space("bottom"),
        left: side_space("left"),
    };
    (insets.top > 0.0 || insets.right > 0.0 || insets.bottom > 0.0 || insets.left > 0.0)
        .then_some(insets)
}

fn extract_hf_paragraph_border(property: &docx_rs::ParagraphProperty) -> Option<CellBorder> {
    let borders = serde_json::to_value(property.borders.as_ref()?).ok()?;
    let extract_side = |key: &str| -> Option<BorderSide> {
        let side = borders.get(key)?.as_object()?;
        let border_type = side
            .get("borderType")
            .or_else(|| side.get("val"))?
            .as_str()?;
        if matches!(border_type, "none" | "nil") {
            return None;
        }
        let width = side.get("size")?.as_f64()? / 8.0;
        let color = side
            .get("color")
            .and_then(serde_json::Value::as_str)
            .filter(|value| *value != "auto")
            .and_then(parse_hex_color)
            .unwrap_or_else(Color::black);
        let style = match border_type {
            "dashed" | "dashSmallGap" => BorderLineStyle::Dashed,
            "dotted" => BorderLineStyle::Dotted,
            "dashDotStroked" | "dotDash" => BorderLineStyle::DashDot,
            "dotDotDash" => BorderLineStyle::DashDotDot,
            "double"
            | "thinThickSmallGap"
            | "thickThinSmallGap"
            | "thinThickMediumGap"
            | "thickThinMediumGap"
            | "thinThickLargeGap"
            | "thickThinLargeGap"
            | "thinThickThinSmallGap"
            | "thinThickThinMediumGap"
            | "thinThickThinLargeGap"
            | "triple" => BorderLineStyle::Double,
            _ => BorderLineStyle::Solid,
        };
        Some(BorderSide {
            width,
            color,
            style,
            join: LineJoin::Round,
        })
    };
    let border = CellBorder {
        top: extract_side("top"),
        bottom: extract_side("bottom"),
        left: extract_side("left"),
        right: extract_side("right"),
    };
    (border.top.is_some()
        || border.bottom.is_some()
        || border.left.is_some()
        || border.right.is_some())
    .then_some(border)
}

fn extract_hf_frame(property: &docx_rs::ParagraphProperty) -> Option<HeaderFooterFrame> {
    let frame = property.frame_property.as_ref()?;
    Some(HeaderFooterFrame {
        x: frame.x.map(twips_to_pt),
        y: frame.y.map(twips_to_pt),
        width: frame.w.map(|value| twips_to_pt(value as i32)),
        height: frame.h.map(|value| twips_to_pt(value as i32)),
        horizontal_anchor: frame_anchor(frame.h_anchor.as_deref()),
        vertical_anchor: frame_anchor(frame.v_anchor.as_deref()),
        horizontal_align: None,
        vertical_align: None,
        inset_left: 0.0,
        inset_top: 0.0,
        bottom_offset: None,
        // A `w:framePr` frame states no wrapping mode of its own.
        wraps_text: true,
    })
}

/// EMU per point (914400 EMU/inch / 72 pt/inch).
const EMU_PER_POINT: f64 = 12700.0;

/// `<a:bodyPr>`'s default left and right inset, 91440 EMU (ECMA-376 §20.1.2.2.5).
const DEFAULT_TEXT_INSET_PT: f64 = 7.2;

/// `<a:bodyPr>`'s default top and bottom inset, 45720 EMU.
const DEFAULT_VERTICAL_INSET_PT: f64 = 3.6;

/// One `<wp:anchor>` in a header or footer part: where its shape sits and how
/// big it is.
///
/// A header/footer story's anchored shapes never reach docx-rs' `Header`/
/// `Footer` model with their positioning intact, so the `wp:anchor` attributes
/// are scanned from the part directly and zipped with the drawings in document
/// order (issue #847).
#[derive(Debug, Clone, Default)]
struct HfAnchorBox {
    width_pt: Option<f64>,
    height_pt: Option<f64>,
    horizontal: HfAnchorAxis,
    vertical: HfAnchorAxis,
    /// `<wps:bodyPr>` left and right insets in points. The #1219 / PR #1407
    /// footer states `lIns="254000"` — 20pt — so the text box's own padding is
    /// not optional detail here. The renderer keeps Writer's separate 0.15pt
    /// page-left text-origin seat out of this value so it cannot narrow the
    /// parsed text column (issue #1487).
    left_inset_pt: f64,
    right_inset_pt: f64,
    top_inset_pt: f64,
    bottom_inset_pt: f64,
    /// `<wp:anchor behindDoc="1">` — the shape sits under the page's content.
    behind_doc: bool,
    /// The geometry and fill the anchor's `<wps:wsp>` declares, when it has
    /// any worth drawing (issue #961).
    shape: Option<crate::ir::Shape>,
    /// `<a:bodyPr anchor="b">` — the text sits at the box's bottom edge.
    seats_text_at_bottom: bool,
    /// `<a:bodyPr wrap="none">` — the paragraph stays on one line and hangs
    /// out of the text column rather than breaking (issue #967).
    wraps_text: bool,
}

#[derive(Debug, Clone, Default)]
struct HfAnchorAxis {
    relative_from_page: bool,
    align: Option<crate::ir::FrameAlign>,
    offset_pt: Option<f64>,
}

impl HfAnchorBox {
    /// The frame this anchor describes, or `None` when it is positioned
    /// relative to something this path does not model — a column or a
    /// character, say — where guessing would put the shape somewhere Word
    /// never does.
    fn to_frame(&self) -> Option<HeaderFooterFrame> {
        (self.horizontal.relative_from_page && self.vertical.relative_from_page).then(|| {
            HeaderFooterFrame {
                // The frame is the text's column, so the box's own left inset
                // moves it and both insets narrow it.
                x: self.horizontal.offset_pt,
                y: self.vertical.offset_pt,
                width: self
                    .width_pt
                    .map(|width| (width - self.left_inset_pt - self.right_inset_pt).max(0.0)),
                height: self.height_pt,
                horizontal_anchor: FrameAnchor::Page,
                vertical_anchor: FrameAnchor::Page,
                horizontal_align: self.horizontal.align,
                vertical_align: self.vertical.align,
                inset_left: self.left_inset_pt,
                inset_top: self.top_inset_pt,
                // Only a box pinned to the bottom of its reference frame can
                // resolve this without knowing where the box itself landed.
                bottom_offset: (self.seats_text_at_bottom
                    && self.vertical.align == Some(crate::ir::FrameAlign::End))
                .then_some(self.bottom_inset_pt),
                wraps_text: self.wraps_text,
            }
        })
    }

    /// The frame of the drawing box itself.
    ///
    /// [`Self::to_frame`] describes the *text column*, which the box's own
    /// padding narrows and shifts. A shape is drawn against the box, so it
    /// takes the untouched extent (issue #961).
    fn to_shape_frame(&self) -> Option<HeaderFooterFrame> {
        let mut frame = self.to_frame()?;
        frame.width = self.width_pt;
        frame.inset_left = 0.0;
        frame.inset_top = 0.0;
        frame.bottom_offset = None;
        Some(frame)
    }
}

/// The page-anchored shapes a header or footer story draws.
///
/// A shape that carries no text still carries a fill, and a decorative banner
/// is nothing else: the invoice of #841 draws its two green wedges this way,
/// as one `a:custGeom` under an `a:gradFill` (issue #961).
fn hf_anchored_shapes(anchors: &[HfAnchorBox]) -> Vec<crate::ir::HeaderFooterShape> {
    anchors
        .iter()
        .filter_map(|anchor| {
            let shape = anchor.shape.clone()?;
            let frame = anchor.to_shape_frame()?;
            Some(crate::ir::HeaderFooterShape {
                shape,
                width: anchor.width_pt?,
                height: anchor.height_pt?,
                frame,
                behind_text: anchor.behind_doc,
            })
        })
        .collect()
}

/// The framed paragraphs a paragraph's anchored text-box drawings contribute
/// to a header or footer story.
///
/// A `<wps:wsp>` in a header or footer produced nothing at all — neither its
/// fill nor the text in its `w:txbxContent` — because the story path reads
/// only inline runs and images. Its content is laid out as its own paragraphs,
/// pinned to the page by the `wp:anchor` beside it (issue #847).
fn hf_anchored_text_box_paragraphs(
    paragraph: &docx_rs::Paragraph,
    anchors: &mut std::slice::Iter<'_, HfAnchorBox>,
    styles: HeaderFooterStyleContext<'_>,
) -> Vec<HeaderFooterParagraph> {
    let mut framed: Vec<HeaderFooterParagraph> = Vec::new();
    for child in &paragraph.children {
        let docx_rs::ParagraphChild::Run(run) = child else {
            continue;
        };
        for run_child in &run.children {
            let docx_rs::RunChild::Drawing(drawing) = run_child else {
                continue;
            };
            // Every anchored drawing consumes one scanned anchor, so an image
            // beside a shape does not shift the shape onto the wrong box.
            let anchor: Option<&HfAnchorBox> = anchors.next();
            let Some(docx_rs::DrawingData::TextBox(text_box)) = &drawing.data else {
                continue;
            };
            let Some(frame) = anchor.and_then(HfAnchorBox::to_frame) else {
                continue;
            };
            for text_box_child in &text_box.children {
                let docx_rs::TextBoxContentChild::Paragraph(inner) = text_box_child else {
                    continue;
                };
                let mut converted =
                    convert_hf_paragraph(inner, &ImageMap::new(), false, &[], styles);
                if !hf_paragraph_carries_text(&converted) {
                    continue;
                }
                converted.frame = Some(frame.clone());
                framed.push(converted);
            }
        }
    }
    framed
}

/// Whether a converted story paragraph carries any text at all, so an empty
/// `w:txbxContent` line does not place a blank frame on the page.
fn hf_paragraph_carries_text(paragraph: &HeaderFooterParagraph) -> bool {
    paragraph.elements.iter().any(|element| match element {
        crate::ir::HFInline::Run(run) => !run.text.trim().is_empty(),
        _ => true,
    })
}

/// A shape with nothing stated yet: a plain rectangle with no fill, which the
/// geometry and fill handlers then fill in.
fn default_anchor_shape() -> crate::ir::Shape {
    crate::ir::Shape {
        kind: crate::ir::ShapeKind::Rectangle,
        fill: None,
        gradient_fill: None,
        pattern_fill: None,
        stroke: None,
        rotation_deg: None,
        opacity: None,
        shadow: None,
        top_bevel: None,
    }
}

/// Read a DOCX shape's `<a:gradFill>` stops and angle, resolving each scheme
/// colour against the document theme. Shared by story anchors and WPG shapes.
///
/// The reader is positioned just after the start tag and is consumed through
/// the matching end tag either way, so a gradient this cannot express still
/// leaves the caller's scan in step. `None` when fewer than two stops resolve
/// — one stop is not a gradient, and zero is not a fill.
pub(in crate::parser) fn parse_docx_shape_gradient(
    reader: &mut quick_xml::Reader<&[u8]>,
    theme_colors: &HashMap<String, Color>,
) -> Option<crate::ir::GradientFill> {
    use quick_xml::events::Event;

    let mut stops: Vec<crate::ir::GradientStop> = Vec::new();
    let mut angle: f64 = 0.0;
    let mut pending_position: Option<f64> = None;
    loop {
        let event = reader.read_event();
        let is_empty: bool = matches!(event, Ok(Event::Empty(_)));
        let element = match &event {
            Ok(Event::Start(element) | Event::Empty(element)) => element,
            Ok(Event::End(element)) if element.local_name().as_ref() == b"gradFill" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => continue,
        };
        match element.local_name().as_ref() {
            // `a:gs@pos` is in thousandths of a percent.
            b"gs" => pending_position = attribute_f64(element, b"pos").map(|pos| pos / 100_000.0),
            b"srgbClr" | b"schemeClr" | b"sysClr" => {
                // `<a:schemeClr val="accent3"><a:lumMod/><a:lumOff/></a:schemeClr>`
                // is how the invoice states both of its stops, so the nested
                // transforms are the colour, not decoration on it.
                let no_aliases: HashMap<String, String> = HashMap::new();
                let scheme = crate::parser::drawingml::SchemeColors {
                    colors: theme_colors,
                    aliases: &no_aliases,
                };
                let parsed = if is_empty {
                    crate::parser::drawingml::parse_color_from_empty(element, &scheme)
                } else {
                    crate::parser::drawingml::parse_color_from_start(reader, element, &scheme)
                };
                if let (Some(position), Some(color)) = (pending_position, parsed.color) {
                    stops.push(crate::ir::GradientStop { position, color });
                    pending_position = None;
                }
            }
            // `a:lin@ang` is in 60000ths of a degree, clockwise from the
            // positive x axis — the convention `GradientFill::angle` uses.
            b"lin" => {
                if let Some(raw) = attribute_f64(element, b"ang") {
                    angle = (raw / 60_000.0).rem_euclid(360.0);
                }
            }
            _ => {}
        }
    }
    (stops.len() >= 2).then_some(crate::ir::GradientFill { stops, angle })
}

fn attribute_f64(element: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<f64> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            std::str::from_utf8(attribute.value.as_ref())
                .ok()?
                .trim()
                .parse::<f64>()
                .ok()
        })
}

/// Read a `<wps:bodyPr>`'s padding, bottom-seating and wrap mode onto the
/// anchor being scanned.
fn read_body_insets(element: &quick_xml::events::BytesStart<'_>, anchor: Option<&mut HfAnchorBox>) {
    let Some(anchor) = anchor else {
        return;
    };
    let inset = |name: &[u8]| -> Option<f64> {
        element
            .attributes()
            .flatten()
            .find(|attribute| attribute.key.local_name().as_ref() == name)
            .and_then(|attribute| {
                std::str::from_utf8(attribute.value.as_ref())
                    .ok()?
                    .trim()
                    .parse::<f64>()
                    .ok()
            })
            .map(|emu| emu / EMU_PER_POINT)
    };
    // ECMA-376's defaults when the attribute is absent.
    anchor.left_inset_pt = inset(b"lIns").unwrap_or(DEFAULT_TEXT_INSET_PT);
    anchor.right_inset_pt = inset(b"rIns").unwrap_or(DEFAULT_TEXT_INSET_PT);
    anchor.top_inset_pt = inset(b"tIns").unwrap_or(DEFAULT_VERTICAL_INSET_PT);
    anchor.bottom_inset_pt = inset(b"bIns").unwrap_or(DEFAULT_VERTICAL_INSET_PT);
    anchor.seats_text_at_bottom = element.attributes().flatten().any(|attribute| {
        attribute.key.local_name().as_ref() == b"anchor" && attribute.value.as_ref() == b"b"
    });
    anchor.wraps_text = !element.attributes().flatten().any(|attribute| {
        attribute.key.local_name().as_ref() == b"wrap" && attribute.value.as_ref() == b"none"
    });
}

/// Scan a header or footer part for its `<wp:anchor>` boxes, in document order.
fn scan_hf_anchors(xml: &str, theme_colors: &HashMap<String, Color>) -> Vec<HfAnchorBox> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut anchors: Vec<HfAnchorBox> = Vec::new();
    let mut current: Option<HfAnchorBox> = None;
    // Which of the two `<wp:positionH>`/`<wp:positionV>` subtrees the scan is
    // inside, so `<wp:align>` and `<wp:posOffset>` land on the right axis.
    let mut axis: Option<bool> = None; // Some(true) = horizontal
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element)) => match element.local_name().as_ref() {
                b"anchor" => {
                    let behind_doc = element.attributes().flatten().any(|attribute| {
                        attribute.key.local_name().as_ref() == b"behindDoc"
                            && matches!(attribute.value.as_ref(), b"1" | b"true")
                    });
                    current = Some(HfAnchorBox {
                        behind_doc,
                        // Only `wrap="none"` turns wrapping off, so a box that
                        // states nothing wraps (issue #967).
                        wraps_text: true,
                        ..HfAnchorBox::default()
                    });
                }
                b"positionH" | b"positionV" => {
                    let horizontal = element.local_name().as_ref() == b"positionH";
                    axis = Some(horizontal);
                    let page = element.attributes().flatten().any(|attribute| {
                        attribute.key.local_name().as_ref() == b"relativeFrom"
                            && attribute.value.as_ref() == b"page"
                    });
                    if let Some(anchor) = current.as_mut() {
                        let target = if horizontal {
                            &mut anchor.horizontal
                        } else {
                            &mut anchor.vertical
                        };
                        target.relative_from_page = page;
                    }
                }
                b"bodyPr" => read_body_insets(element, current.as_mut()),
                // The geometry and the fill are separate subtrees of the same
                // `<wps:spPr>`; either can be absent, and a shape needs both
                // before it is worth drawing (issue #961).
                // `<a:xfrm rot>` is in 60000ths of a degree. The footer band
                // of #841 is the header's wedge at `rot="10800000"` — 180
                // degrees — so dropping it mirrors the band (issue #961).
                b"xfrm" => {
                    let rotation: Option<f64> = attribute_f64(element, b"rot")
                        .map(|raw| (raw / 60_000.0).rem_euclid(360.0))
                        .filter(|degrees| *degrees != 0.0);
                    if let Some(anchor) = current.as_mut()
                        && rotation.is_some()
                    {
                        anchor
                            .shape
                            .get_or_insert_with(default_anchor_shape)
                            .rotation_deg = rotation;
                    }
                }
                b"custGeom" => {
                    // `<wp:extent>` precedes the graphic, so the box a guide
                    // formula measures against is already known. Points are
                    // fine here: only the ratio survives normalization.
                    let extent = crate::parser::pptx::geometry_guides::ShapeExtent::new(
                        current
                            .as_ref()
                            .and_then(|anchor| anchor.width_pt)
                            .unwrap_or(0.0),
                        current
                            .as_ref()
                            .and_then(|anchor| anchor.height_pt)
                            .unwrap_or(0.0),
                    );
                    let subpaths = crate::parser::pptx::custom_geometry::parse_custom_geometry(
                        &mut reader,
                        extent,
                    );
                    if let Some(anchor) = current.as_mut()
                        && !subpaths.is_empty()
                    {
                        anchor.shape.get_or_insert_with(default_anchor_shape).kind =
                            crate::ir::ShapeKind::Path { subpaths };
                    }
                }
                b"gradFill" => {
                    let gradient = parse_docx_shape_gradient(&mut reader, theme_colors);
                    if let Some(anchor) = current.as_mut()
                        && gradient.is_some()
                    {
                        anchor
                            .shape
                            .get_or_insert_with(default_anchor_shape)
                            .gradient_fill = gradient;
                    }
                }
                b"align" | b"posOffset" => {
                    let is_align = element.local_name().as_ref() == b"align";
                    let name = element.name().to_owned();
                    let Ok(text) = reader.read_text(quick_xml::name::QName(name.as_ref())) else {
                        continue;
                    };
                    let (Some(anchor), Some(horizontal)) = (current.as_mut(), axis) else {
                        continue;
                    };
                    let target = if horizontal {
                        &mut anchor.horizontal
                    } else {
                        &mut anchor.vertical
                    };
                    if is_align {
                        target.align = match text.trim() {
                            "left" | "top" | "inside" => Some(crate::ir::FrameAlign::Start),
                            "center" => Some(crate::ir::FrameAlign::Center),
                            "right" | "bottom" | "outside" => Some(crate::ir::FrameAlign::End),
                            _ => None,
                        };
                    } else if let Ok(emu) = text.trim().parse::<f64>() {
                        target.offset_pt = Some(emu / EMU_PER_POINT);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bodyPr" => {
                read_body_insets(element, current.as_mut());
            }
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extent" => {
                let value = |name: &[u8]| -> Option<f64> {
                    element
                        .attributes()
                        .flatten()
                        .find(|attribute| attribute.key.local_name().as_ref() == name)
                        .and_then(|attribute| {
                            std::str::from_utf8(attribute.value.as_ref())
                                .ok()?
                                .trim()
                                .parse::<f64>()
                                .ok()
                        })
                        .map(|emu| emu / EMU_PER_POINT)
                };
                if let Some(anchor) = current.as_mut() {
                    anchor.width_pt = value(b"cx");
                    anchor.height_pt = value(b"cy");
                }
            }
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"anchor" => {
                if let Some(anchor) = current.take() {
                    anchors.push(anchor);
                }
                axis = None;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    anchors
}

fn frame_anchor(value: Option<&str>) -> FrameAnchor {
    match value {
        Some("page") => FrameAnchor::Page,
        Some("margin") => FrameAnchor::Margin,
        _ => FrameAnchor::Text,
    }
}

/// `w:fldSimple` keeps its run properties inside the element rather than on the
/// surrounding run, which docx-rs does not surface here. The adjacent run's
/// style is the closest available match and is what these headers declare.
fn append_simple_fields(
    elements: &mut Vec<HFInline>,
    simple_fields: &[SimpleFieldMarker],
    processed_runs: usize,
    style: &TextStyle,
) -> usize {
    simple_fields
        .iter()
        .filter(|field| field.preceding_runs == processed_runs)
        .map(|field| {
            elements.push(match field.kind {
                SimpleFieldKind::PageNumber => HFInline::PageNumber(style.clone()),
                SimpleFieldKind::TotalPages => HFInline::TotalPages(style.clone()),
            });
            field.cached_runs
        })
        .sum()
}

/// A `w:fldChar` field in flight.
///
/// Word writes one across several `w:r` elements — begin, the instruction,
/// separate, the cached result, end — so this cannot live inside a single
/// run's scope. Keeping it per-run meant a split PAGE field never resolved
/// and its cached number fell out as static text (issue #738).
#[derive(Default)]
struct HeaderFieldState {
    in_field: bool,
    inline: Option<HFInline>,
    past_separate: bool,
    /// The field's cached result, kept in case the instruction turns out to be
    /// one we do not model. Word shows that text, and so did this code before
    /// the state spanned runs — back then `in_field` was false by the time the
    /// result's own run was read, so it fell through as ordinary text.
    cached_result: Vec<Run>,
}

/// Extract inline elements from a run's children for header/footer use.
/// Recognizes text, tabs, and PAGE/NUMPAGES field codes.
///
/// `field` carries the enclosing paragraph's field state, so a field split
/// across runs resolves as one.
fn extract_hf_run_elements(
    children: &[docx_rs::RunChild],
    style: &TextStyle,
    elements: &mut Vec<HFInline>,
    field: &mut HeaderFieldState,
) {
    let HeaderFieldState {
        in_field,
        inline: field_inline,
        past_separate,
        cached_result,
    } = field;

    for child in children {
        match child {
            docx_rs::RunChild::FieldChar(field_char) => match field_char.field_char_type {
                docx_rs::FieldCharType::Begin => {
                    *in_field = true;
                    *field_inline = None;
                    *past_separate = false;
                    cached_result.clear();
                }
                docx_rs::FieldCharType::Separate => {
                    *past_separate = true;
                }
                docx_rs::FieldCharType::End => {
                    match field_inline.take() {
                        Some(inline) => elements.push(inline),
                        // An instruction we do not model: show what Word
                        // cached, rather than dropping the field's text.
                        None => elements.extend(cached_result.drain(..).map(HFInline::Run)),
                    }
                    cached_result.clear();
                    *in_field = false;
                    *past_separate = false;
                }
                _ => {}
            },
            docx_rs::RunChild::InstrText(instruction) => {
                if !*in_field {
                    continue;
                }
                *field_inline = match instruction.as_ref() {
                    docx_rs::InstrText::PAGE(_) => Some(HFInline::PageNumber(style.clone())),
                    docx_rs::InstrText::NUMPAGES(_) => Some(HFInline::TotalPages(style.clone())),
                    _ => field_inline.take(),
                };
            }
            docx_rs::RunChild::InstrTextString(value) => {
                if !*in_field {
                    continue;
                }
                let trimmed = value.trim();
                if trimmed.eq_ignore_ascii_case("page") {
                    *field_inline = Some(HFInline::PageNumber(style.clone()));
                } else if trimmed.eq_ignore_ascii_case("numpages") {
                    *field_inline = Some(HFInline::TotalPages(style.clone()));
                }
            }
            docx_rs::RunChild::Text(text) => {
                if *in_field && *past_separate {
                    if !text.text.is_empty() {
                        cached_result.push(Run {
                            text: text.text.clone(),
                            style: style.clone(),
                            href: None,
                            footnote: None,
                        });
                    }
                    continue;
                }
                if !*in_field && !text.text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: text.text.clone(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
            }
            docx_rs::RunChild::Tab(_) if !*in_field => {
                elements.push(HFInline::Run(Run {
                    text: "\t".to_string(),
                    style: style.clone(),
                    href: None,
                    footnote: None,
                }));
            }
            docx_rs::RunChild::PTab(tab) if !*in_field => {
                let alignment = match tab.alignment {
                    docx_rs::PositionalTabAlignmentType::Center => PositionedTabAlignment::Center,
                    docx_rs::PositionalTabAlignmentType::Right => PositionedTabAlignment::Right,
                    docx_rs::PositionalTabAlignmentType::Left => PositionedTabAlignment::Left,
                };
                let relative_to = match tab.relative_to {
                    docx_rs::PositionalTabRelativeTo::Indent => PositionedTabRelativeTo::Indent,
                    docx_rs::PositionalTabRelativeTo::Margin => PositionedTabRelativeTo::Margin,
                };
                let leader = match tab.leader {
                    docx_rs::TabLeaderType::Dot => TabLeader::Dot,
                    docx_rs::TabLeaderType::Hyphen => TabLeader::Hyphen,
                    docx_rs::TabLeaderType::Underscore => TabLeader::Underscore,
                    _ => TabLeader::None,
                };
                elements.push(HFInline::PositionedTab(PositionedTab {
                    alignment,
                    relative_to,
                    leader,
                }));
            }
            _ => {}
        }
    }
}

/// Extract page size and margins from DOCX section properties.
fn extract_page_setup(section_prop: &docx_rs::SectionProperty) -> (PageSize, Margins) {
    let size = extract_page_size(&section_prop.page_size);
    let margins = extract_margins(&section_prop.page_margin);
    (size, margins)
}

/// Extract page size from docx-rs PageSize (which has private fields).
/// Uses serde serialization to access the private `w`, `h`, and `orient` fields.
/// Values in DOCX are in twips (1/20 of a point).
/// When orient is "landscape" and width < height, dimensions are swapped to ensure
/// landscape pages have width > height.
pub(super) fn extract_page_size(page_size: &docx_rs::PageSize) -> PageSize {
    if let Ok(json) = serde_json::to_value(page_size) {
        let width_twips = json
            .get("w")
            .and_then(|value| value.as_f64())
            .unwrap_or(0.0);
        let height_twips = json
            .get("h")
            .and_then(|value| value.as_f64())
            .unwrap_or(0.0);
        let orientation = json.get("orient").and_then(|value| value.as_str());
        if width_twips > 0.0 && height_twips > 0.0 {
            let mut width = twips_to_pt(width_twips);
            let mut height = twips_to_pt(height_twips);
            if orientation == Some("landscape") && width < height {
                std::mem::swap(&mut width, &mut height);
            }
            return PageSize { width, height };
        }
    }
    PageSize::default()
}

/// Extract margins from docx-rs PageMargin.
/// PageMargin fields are public i32 values in twips.
fn extract_margins(page_margin: &docx_rs::PageMargin) -> Margins {
    Margins {
        top: twips_to_pt(page_margin.top),
        bottom: twips_to_pt(page_margin.bottom),
        left: twips_to_pt(page_margin.left),
        right: twips_to_pt(page_margin.right),
    }
}

#[cfg(test)]
mod anchor_tests {
    use super::*;

    /// The footer part of `003_FAKTURA.docx` (issue #841), trimmed to the
    /// attributes that position its `Sensitivity: Internal` shape.
    const FOOTER_ANCHOR: &str = r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0">
        <wp:simplePos x="635" y="635"/>
        <wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>
        <wp:extent cx="1089660" cy="351790"/>
        <wps:wsp><wps:bodyPr lIns="254000" tIns="0" rIns="0" bIns="190500" anchor="b"/></wps:wsp>
      </wp:anchor></w:drawing></w:r></w:p>
    </w:ftr>"#;

    /// A header/footer story's anchored shape carries its own position, size
    /// and padding, none of which reaches docx-rs' `Footer` model — so the
    /// `wp:anchor` is scanned from the part directly (issue #847).
    #[test]
    fn a_header_footer_anchor_is_scanned_from_the_part() {
        let anchors = scan_hf_anchors(FOOTER_ANCHOR, &HashMap::new());
        assert_eq!(anchors.len(), 1, "one anchored shape: {anchors:?}");
        let frame = anchors[0].to_frame().expect("a page-relative frame");

        assert_eq!(frame.horizontal_anchor, FrameAnchor::Page);
        assert_eq!(frame.vertical_anchor, FrameAnchor::Page);
        assert_eq!(frame.horizontal_align, Some(crate::ir::FrameAlign::Start));
        assert_eq!(frame.vertical_align, Some(crate::ir::FrameAlign::End));
        // 1089660 EMU is 85.8pt, less the 20pt left inset and the zero right
        // one, so the text column is 65.8pt.
        assert!(frame.width.is_some_and(|width| (width - 65.8).abs() < 0.01));
        assert!(
            (frame.inset_left - 20.0).abs() < 0.01,
            "lIns=254000 is 20pt"
        );
        // `anchor="b"` seats the text at the box's bottom edge, 15pt up.
        assert!(
            frame
                .bottom_offset
                .is_some_and(|gap| (gap - 15.0).abs() < 0.01),
            "bIns=190500 is 15pt: {:?}",
            frame.bottom_offset
        );
        assert!(
            frame.x.is_none() && frame.y.is_none(),
            "aligned, not offset"
        );
    }

    /// An anchor measured against something other than the page — a column,
    /// a margin, a character — is left alone rather than placed on a guess.
    #[test]
    fn an_anchor_relative_to_something_else_yields_no_frame() {
        let relative = FOOTER_ANCHOR.replace(r#"relativeFrom="page""#, r#"relativeFrom="column""#);
        let anchors = scan_hf_anchors(&relative, &HashMap::new());
        assert_eq!(anchors.len(), 1);
        assert!(anchors[0].to_frame().is_none());
    }

    /// The header part of `003_FAKTURA.docx` (issue #961), trimmed to the
    /// banner shape: a `custGeom` wedge under a two-stop `gradFill`, carrying
    /// no text at all.
    const HEADER_BANNER: &str = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:p><w:r><w:drawing><wp:anchor behindDoc="1" distL="114300" distR="114300">
        <wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>
        <wp:extent cx="7735824" cy="4160520"/>
        <wps:wsp><wps:spPr>
          <a:custGeom><a:pathLst><a:path w="7738110" h="2906395">
            <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
            <a:lnTo><a:pt x="7738110" y="0"/></a:lnTo>
            <a:lnTo><a:pt x="7738110" y="1896461"/></a:lnTo>
            <a:lnTo><a:pt x="0" y="2906395"/></a:lnTo>
            <a:close/>
          </a:path></a:pathLst></a:custGeom>
          <a:gradFill flip="none" rotWithShape="1"><a:gsLst>
            <a:gs pos="0"><a:schemeClr val="accent3"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="accent5"><a:lumMod val="100000"/></a:schemeClr></a:gs>
          </a:gsLst><a:lin ang="1920000" scaled="0"/><a:tileRect/></a:gradFill>
          <a:ln><a:noFill/></a:ln>
        </wps:spPr>
        <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
        <wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
        </wps:wsp>
      </wp:anchor></w:drawing></w:r></w:p>
    </w:hdr>"#;

    fn accent_theme() -> HashMap<String, Color> {
        HashMap::from([
            ("accent3".to_string(), Color::new(0xA5, 0xA5, 0xA5)),
            ("accent5".to_string(), Color::new(0x00, 0x80, 0x40)),
        ])
    }

    /// The banner carries no text, so nothing about it reaches the paragraph
    /// path — its geometry and fill have to be scanned as a shape or the
    /// header renders blank (issue #961).
    #[test]
    fn a_fill_only_header_banner_is_scanned_as_a_shape() {
        let anchors = scan_hf_anchors(HEADER_BANNER, &accent_theme());
        assert_eq!(anchors.len(), 1, "one anchored shape: {anchors:?}");
        let shapes = hf_anchored_shapes(&anchors);
        assert_eq!(shapes.len(), 1, "the banner is drawable: {shapes:?}");
        let banner = &shapes[0];

        assert!(banner.behind_text, "behindDoc=\"1\"");
        // 7735824 EMU is 609.12pt and 4160520 is 327.60pt.
        assert!((banner.width - 609.12).abs() < 0.01, "{}", banner.width);
        assert!((banner.height - 327.60).abs() < 0.01, "{}", banner.height);

        let crate::ir::ShapeKind::Path { subpaths } = &banner.shape.kind else {
            panic!("a custGeom wedge, not {:?}", banner.shape.kind);
        };
        assert_eq!(subpaths.len(), 1);
        // The wedge's right edge stops at 1896461/2906395 of the path box,
        // and the path box is stretched onto the shape's extent.
        let right_edge: f64 = subpaths[0].vertices[2].1;
        assert!((right_edge - 0.6525).abs() < 0.001, "{right_edge}");
        assert!(
            (subpaths[0].vertices[3].1 - 1.0).abs() < 0.001,
            "the left edge drops to the bottom"
        );
    }

    /// `<a:lin ang>` is in 60000ths of a degree, and each stop's scheme colour
    /// carries `lumMod`/`lumOff` that decide what green it actually is.
    #[test]
    fn the_banner_gradient_resolves_its_scheme_stops() {
        let anchors = scan_hf_anchors(HEADER_BANNER, &accent_theme());
        let shapes = hf_anchored_shapes(&anchors);
        let gradient = shapes[0]
            .shape
            .gradient_fill
            .as_ref()
            .expect("a two-stop gradient");

        assert!(
            (gradient.angle - 32.0).abs() < 0.01,
            "1920000/60000 is 32deg"
        );
        assert_eq!(gradient.stops.len(), 2);
        assert!((gradient.stops[0].position - 0.0).abs() < 1e-9);
        assert!((gradient.stops[1].position - 1.0).abs() < 1e-9);
        // accent3 at lumMod 40% + lumOff 60% is far lighter than the flat
        // theme colour, so an unresolved stop would be plainly wrong.
        assert!(
            gradient.stops[0].color != Color::new(0xA5, 0xA5, 0xA5),
            "lumMod/lumOff applied: {:?}",
            gradient.stops[0].color
        );
        assert_eq!(gradient.stops[1].color, Color::new(0x00, 0x80, 0x40));
    }

    /// An absent `<a:bodyPr>` inset falls back to ECMA-376's own default
    /// rather than to zero.
    #[test]
    fn an_unstated_inset_takes_the_schema_default() {
        let bare = FOOTER_ANCHOR.replace(
            r#"lIns="254000" tIns="0" rIns="0" bIns="190500" anchor="b""#,
            "",
        );
        let anchors = scan_hf_anchors(&bare, &HashMap::new());
        let frame = anchors[0].to_frame().expect("a frame");
        assert!((frame.inset_left - 7.2).abs() < 0.01, "91440 EMU is 7.2pt");
        assert!(
            frame.bottom_offset.is_none(),
            "no `anchor=\"b\"`, so the text is not bottom-seated"
        );
    }
}

#[cfg(test)]
mod body_pr_wrap_tests {
    use super::*;

    /// The sensitivity label of `003_FAKTURA.docx` (issue #967), trimmed to the
    /// `<wps:bodyPr>` attributes that decide whether its one line breaks.
    const LABEL: &str = r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:p><w:r><w:drawing><wp:anchor distT="0" distB="0">
        <wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>
        <wp:extent cx="1089660" cy="351790"/>
        <wps:wsp><wps:bodyPr wrap="none" horzOverflow="overflow"
          lIns="254000" tIns="0" rIns="0" bIns="190500" anchor="b"/></wps:wsp>
      </wp:anchor></w:drawing></w:r></w:p>
    </w:ftr>"#;

    /// `wrap="none"` keeps the paragraph on one line and lets it hang out of
    /// the text column; `horzOverflow="overflow"` beside it is what permits the
    /// overhang. Ignoring the attribute broke a label 1.33pt wider than its
    /// column into two lines (issue #967).
    #[test]
    fn a_body_pr_declaring_wrap_none_yields_a_non_wrapping_frame() {
        let anchors = scan_hf_anchors(LABEL, &HashMap::new());
        let frame = anchors[0].to_frame().expect("a page-relative frame");
        assert!(!frame.wraps_text);
    }

    /// `wrap="square"` and an absent attribute both wrap, which is what every
    /// text box did before the attribute was read at all.
    #[test]
    fn every_other_wrap_value_still_wraps() {
        for markup in [
            LABEL.replace(r#"wrap="none""#, r#"wrap="square""#),
            LABEL.replace(r#"wrap="none""#, ""),
        ] {
            let anchors = scan_hf_anchors(&markup, &HashMap::new());
            let frame = anchors[0].to_frame().expect("a page-relative frame");
            assert!(frame.wraps_text, "{markup}");
        }
    }
}
