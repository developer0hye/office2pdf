use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};

use crate::error::ConvertWarning;
use crate::ir::{
    Block, BorderLineStyle, BorderSide, CellBorder, Color, ColumnLayout, FlowPage, FrameAnchor,
    HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph, Margins, PageSize,
    PositionedTab, PositionedTabAlignment, PositionedTabRelativeTo, Run, TabLeader, TextDirection,
    TextStyle,
};

use super::contexts::WrapContext;
use super::media::extract_drawing_image;
use super::{
    ImageMap, NumberingMap, TaggedElement, extract_column_layout_from_section_property,
    extract_paragraph_style, extract_run_style, extract_tab_stop_overrides, group_into_lists,
    merge_paragraph_style, read_zip_text,
};
use crate::parser::units::twips_to_pt;
use crate::parser::xml_util::parse_hex_color;

/// Parsed header/footer assets addressed by relationship ID.
#[derive(Default)]
pub(super) struct HeaderFooterAssets {
    headers: HashMap<String, HeaderFooter>,
    footers: HashMap<String, HeaderFooter>,
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
) -> HeaderFooterAssets {
    let rels_xml = match read_zip_text(archive, "word/_rels/document.xml.rels") {
        Some(xml) => xml,
        None => return HeaderFooterAssets::default(),
    };
    let (header_relationships, footer_relationships) = scan_header_footer_relationships(&rels_xml);
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
        if let Some(converted) = convert_docx_header(&header, &images, &simple_fields) {
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
        if let Some(converted) =
            convert_docx_footer(&footer, &images, &bidi_paragraphs, &simple_fields)
        {
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

pub(super) fn build_flow_page_from_section(
    section_prop: &docx_rs::SectionProperty,
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
    header_footer_assets: &HeaderFooterAssets,
    column_layout: Option<ColumnLayout>,
    warnings: &mut Vec<ConvertWarning>,
) -> FlowPage {
    let (size, margins) = extract_page_setup(section_prop);
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

    if section_prop.first_header_reference.is_some()
        || section_prop.first_footer_reference.is_some()
        || section_prop.even_header_reference.is_some()
        || section_prop.even_footer_reference.is_some()
        || section_prop.first_header.is_some()
        || section_prop.first_footer.is_some()
        || section_prop.even_header.is_some()
        || section_prop.even_footer.is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "header/footer variants".to_string(),
            to: "single header/footer per section".to_string(),
        });
    }

    if section_prop
        .page_num_type
        .as_ref()
        .and_then(|page_number_type| page_number_type.start)
        .is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "section page number restart".to_string(),
            to: "global page counter".to_string(),
        });
    }

    let mut header = extract_docx_header(section_prop, header_footer_assets);
    if let Some(header) = &mut header {
        header.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.header));
    }
    let mut footer = extract_docx_footer(section_prop, header_footer_assets);
    if let Some(footer) = &mut footer {
        footer.distance_from_edge = Some(twips_to_pt(section_prop.page_margin.footer));
    }

    FlowPage {
        size,
        margins,
        content,
        header,
        footer,
        columns: column_layout
            .or_else(|| extract_column_layout_from_section_property(section_prop)),
        line_grid_pitch: extract_line_grid_pitch(section_prop),
    }
}

/// Word snaps body lines to the section's document grid; the pitch is the
/// effective single-spacing line height for grid-aligned paragraphs
/// (`<w:docGrid w:linePitch>`, in twips). docx-rs keeps the fields private,
/// so read them through the type's serde representation.
fn extract_line_grid_pitch(section_prop: &docx_rs::SectionProperty) -> Option<f64> {
    let grid = section_prop.doc_grid.as_ref()?;
    let value = serde_json::to_value(grid).ok()?;
    let pitch_twips = value.get("linePitch")?.as_f64()?;
    (pitch_twips > 0.0).then(|| twips_to_pt(pitch_twips as i32))
}

fn convert_docx_header(
    header: &docx_rs::Header,
    images: &ImageMap,
    simple_fields: &[Vec<SimpleFieldMarker>],
) -> Option<HeaderFooter> {
    let paragraphs = header
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::HeaderChild::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .enumerate()
        .map(|(index, paragraph)| {
            convert_hf_paragraph(
                paragraph,
                images,
                false,
                simple_fields.get(index).map(Vec::as_slice).unwrap_or(&[]),
            )
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        distance_from_edge: None,
    })
}

fn convert_docx_footer(
    footer: &docx_rs::Footer,
    images: &ImageMap,
    bidi_paragraphs: &[bool],
    simple_fields: &[Vec<SimpleFieldMarker>],
) -> Option<HeaderFooter> {
    let paragraphs = footer
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::FooterChild::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .enumerate()
        .map(|(index, paragraph)| {
            convert_hf_paragraph(
                paragraph,
                images,
                bidi_paragraphs.get(index).copied().unwrap_or(false),
                simple_fields.get(index).map(Vec::as_slice).unwrap_or(&[]),
            )
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        distance_from_edge: None,
    })
}

/// Extract the header for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
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
                    convert_docx_header(header, &ImageMap::new(), &[])
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
                    convert_docx_header(header, &ImageMap::new(), &[])
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
                    convert_docx_header(header, &ImageMap::new(), &[])
                })
        })
}

/// Extract the footer for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
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
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[])
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
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[])
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
                    convert_docx_footer(footer, &ImageMap::new(), &[], &[])
                })
        })
}

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE/NUMPAGES field codes within runs and emits page counter inlines.
fn convert_hf_paragraph(
    paragraph: &docx_rs::Paragraph,
    images: &ImageMap,
    is_bidi: bool,
    simple_fields: &[SimpleFieldMarker],
) -> HeaderFooterParagraph {
    let explicit_style = extract_paragraph_style(&paragraph.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&paragraph.property.tabs);
    let mut style = merge_paragraph_style(&explicit_style, explicit_tab_overrides.as_deref(), None);
    if is_bidi || paragraph.property.bidi == Some(true) {
        style.direction = Some(TextDirection::Rtl);
    }
    let mut elements: Vec<HFInline> = Vec::new();

    let mut processed_runs: usize = 0;
    let mut cached_runs_to_skip: usize =
        append_simple_fields(&mut elements, simple_fields, processed_runs);
    for child in &paragraph.children {
        match child {
            docx_rs::ParagraphChild::Run(run) => {
                if cached_runs_to_skip > 0 {
                    cached_runs_to_skip -= 1;
                    continue;
                }
                let run_style = extract_run_style(&run.run_property);
                extract_hf_run_elements(&run.children, &run_style, &mut elements);
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
                    append_simple_fields(&mut elements, simple_fields, processed_runs);
            }
            docx_rs::ParagraphChild::PageNum(_) => elements.push(HFInline::PageNumber),
            docx_rs::ParagraphChild::NumPages(_) => elements.push(HFInline::TotalPages),
            _ => {}
        }
    }

    HeaderFooterParagraph {
        style,
        elements,
        border: extract_hf_paragraph_border(&paragraph.property),
        frame: extract_hf_frame(&paragraph.property),
    }
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
    })
}

fn frame_anchor(value: Option<&str>) -> FrameAnchor {
    match value {
        Some("page") => FrameAnchor::Page,
        Some("margin") => FrameAnchor::Margin,
        _ => FrameAnchor::Text,
    }
}

fn append_simple_fields(
    elements: &mut Vec<HFInline>,
    simple_fields: &[SimpleFieldMarker],
    processed_runs: usize,
) -> usize {
    simple_fields
        .iter()
        .filter(|field| field.preceding_runs == processed_runs)
        .map(|field| {
            elements.push(match field.kind {
                SimpleFieldKind::PageNumber => HFInline::PageNumber,
                SimpleFieldKind::TotalPages => HFInline::TotalPages,
            });
            field.cached_runs
        })
        .sum()
}

/// Extract inline elements from a run's children for header/footer use.
/// Recognizes text, tabs, and PAGE/NUMPAGES field codes.
fn extract_hf_run_elements(
    children: &[docx_rs::RunChild],
    style: &TextStyle,
    elements: &mut Vec<HFInline>,
) {
    let mut in_field = false;
    let mut field_inline: Option<HFInline> = None;
    let mut past_separate = false;

    for child in children {
        match child {
            docx_rs::RunChild::FieldChar(field_char) => match field_char.field_char_type {
                docx_rs::FieldCharType::Begin => {
                    in_field = true;
                    field_inline = None;
                    past_separate = false;
                }
                docx_rs::FieldCharType::Separate => {
                    past_separate = true;
                }
                docx_rs::FieldCharType::End => {
                    if let Some(inline) = field_inline.take() {
                        elements.push(inline);
                    }
                    in_field = false;
                    past_separate = false;
                }
                _ => {}
            },
            docx_rs::RunChild::InstrText(instruction) => {
                if !in_field {
                    continue;
                }
                field_inline = match instruction.as_ref() {
                    docx_rs::InstrText::PAGE(_) => Some(HFInline::PageNumber),
                    docx_rs::InstrText::NUMPAGES(_) => Some(HFInline::TotalPages),
                    _ => field_inline,
                };
            }
            docx_rs::RunChild::InstrTextString(value) => {
                if !in_field {
                    continue;
                }
                let trimmed = value.trim();
                if trimmed.eq_ignore_ascii_case("page") {
                    field_inline = Some(HFInline::PageNumber);
                } else if trimmed.eq_ignore_ascii_case("numpages") {
                    field_inline = Some(HFInline::TotalPages);
                }
            }
            docx_rs::RunChild::Text(text) => {
                if in_field && past_separate {
                    continue;
                }
                if !in_field && !text.text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: text.text.clone(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
            }
            docx_rs::RunChild::Tab(_) if !in_field => {
                elements.push(HFInline::Run(Run {
                    text: "\t".to_string(),
                    style: style.clone(),
                    href: None,
                    footnote: None,
                }));
            }
            docx_rs::RunChild::PTab(tab) if !in_field => {
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
