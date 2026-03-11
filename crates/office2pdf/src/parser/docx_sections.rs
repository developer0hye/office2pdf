use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::error::ConvertWarning;
use crate::ir::{
    Block, ColumnLayout, FlowPage, HFInline, HeaderFooter, HeaderFooterParagraph, Margins,
    PageSize, Run, TextStyle,
};

use super::{
    NumberingMap, TaggedElement, extract_column_layout_from_section_property,
    extract_paragraph_style, extract_run_style, extract_tab_stop_overrides, group_into_lists,
    merge_paragraph_style, read_zip_text,
};
use crate::parser::units::twips_to_pt;

/// Parsed header/footer assets addressed by relationship ID.
#[derive(Default)]
pub(super) struct HeaderFooterAssets {
    headers: HashMap<String, HeaderFooter>,
    footers: HashMap<String, HeaderFooter>,
}

fn scan_header_footer_relationships(
    rels_xml: &str,
) -> (HashMap<String, String>, HashMap<String, String>) {
    let mut headers: HashMap<String, String> = HashMap::new();
    let mut footers: HashMap<String, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(rels_xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                if element.local_name().as_ref() != b"Relationship" {
                    continue;
                }

                let mut id: Option<String> = None;
                let mut target: Option<String> = None;
                let mut relationship_type: Option<String> = None;

                for attribute in element.attributes().flatten() {
                    match attribute.key.local_name().as_ref() {
                        b"Id" => {
                            if let Ok(value) = attribute.unescape_value() {
                                id = Some(value.to_string());
                            }
                        }
                        b"Target" => {
                            if let Ok(value) = attribute.unescape_value() {
                                target = Some(value.to_string());
                            }
                        }
                        b"Type" => {
                            if let Ok(value) = attribute.unescape_value() {
                                relationship_type = Some(value.to_string());
                            }
                        }
                        _ => {}
                    }
                }

                let Some(id) = id else { continue };
                let Some(target) = target else { continue };
                let Some(relationship_type) = relationship_type else {
                    continue;
                };

                let full_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("word/{target}")
                };

                if relationship_type.ends_with("/header") {
                    headers.insert(id, full_path);
                } else if relationship_type.ends_with("/footer") {
                    footers.insert(id, full_path);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
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
        let Ok(header) = <docx_rs::Header as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_header(&header) {
            assets.headers.insert(relationship_id, converted);
        }
    }

    for (relationship_id, path) in footer_relationships {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let Ok(footer) = <docx_rs::Footer as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_footer(&footer) {
            assets.footers.insert(relationship_id, converted);
        }
    }

    assets
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

    FlowPage {
        size,
        margins,
        content,
        header: extract_docx_header(section_prop, header_footer_assets),
        footer: extract_docx_footer(section_prop, header_footer_assets),
        columns: column_layout
            .or_else(|| extract_column_layout_from_section_property(section_prop)),
    }
}

fn convert_docx_header(header: &docx_rs::Header) -> Option<HeaderFooter> {
    let paragraphs = header
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::HeaderChild::Paragraph(paragraph) => Some(convert_hf_paragraph(paragraph)),
            _ => None,
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter { paragraphs })
}

fn convert_docx_footer(footer: &docx_rs::Footer) -> Option<HeaderFooter> {
    let paragraphs = footer
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::FooterChild::Paragraph(paragraph) => Some(convert_hf_paragraph(paragraph)),
            _ => None,
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter { paragraphs })
}

/// Extract the header for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .header
        .as_ref()
        .and_then(|(_relationship_id, header)| convert_docx_header(header))
        .or_else(|| {
            section_prop
                .header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_header
                .as_ref()
                .and_then(|(_relationship_id, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .first_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_header
                .as_ref()
                .and_then(|(_relationship_id, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .even_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
}

/// Extract the footer for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .footer
        .as_ref()
        .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        .or_else(|| {
            section_prop
                .footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .first_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .even_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
}

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE/NUMPAGES field codes within runs and emits page counter inlines.
fn convert_hf_paragraph(paragraph: &docx_rs::Paragraph) -> HeaderFooterParagraph {
    let explicit_style = extract_paragraph_style(&paragraph.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&paragraph.property.tabs);
    let style = merge_paragraph_style(&explicit_style, explicit_tab_overrides.as_deref(), None);
    let mut elements: Vec<HFInline> = Vec::new();

    for child in &paragraph.children {
        if let docx_rs::ParagraphChild::Run(run) = child {
            let run_style = extract_run_style(&run.run_property);
            extract_hf_run_elements(&run.children, &run_style, &mut elements);
        }
    }

    HeaderFooterParagraph { style, elements }
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
            docx_rs::RunChild::Tab(_) => {
                if !in_field {
                    elements.push(HFInline::Run(Run {
                        text: "\t".to_string(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
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
