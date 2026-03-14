use crate::parser::chart as chart_parser;

use super::*;

/// Build the .rels path for a given file path.
///
/// e.g., `ppt/slides/slide1.xml` -> `ppt/slides/_rels/slide1.xml.rels`
fn rels_path_for(path: &str) -> String {
    if let Some((dir, filename)) = path.rsplit_once('/') {
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{path}.rels")
    }
}

/// Resolve the layout and master file paths from a slide's .rels.
pub(super) fn resolve_layout_master_paths<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> (Option<String>, Option<String>) {
    let slide_dir: &str = slide_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");

    let Ok(rels_xml) = read_zip_entry(archive, &rels_path_for(slide_path)) else {
        return (None, None);
    };
    let rels: HashMap<String, String> = parse_rels_xml(&rels_xml);

    let layout_path: Option<String> = rels
        .values()
        .find(|target| target.contains("slideLayout") || target.contains("slideLayouts"))
        .map(|target| resolve_relative_path(slide_dir, target));

    let Some(layout_path_ref) = layout_path.as_ref() else {
        return (None, None);
    };

    let layout_dir: &str = layout_path_ref
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let master_path: Option<String> = read_zip_entry(archive, &rels_path_for(layout_path_ref))
        .ok()
        .and_then(|layout_rels_xml| {
            let layout_rels: HashMap<String, String> = parse_rels_xml(&layout_rels_xml);
            layout_rels
                .values()
                .find(|target| target.contains("slideMaster") || target.contains("slideMasters"))
                .map(|target| resolve_relative_path(layout_dir, target))
        });

    (layout_path, master_path)
}

/// Read a file from the ZIP archive as a UTF-8 string.
pub(super) fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String, ConvertError> {
    let mut file = archive
        .by_name(path)
        .map_err(|error| crate::parser::parse_err(format!("Missing {path} in PPTX: {error}")))?;
    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|error| crate::parser::parse_err(format!("Failed to read {path}: {error}")))?;
    Ok(content)
}

/// Load images referenced by a slide from its .rels file and the ZIP archive.
pub(super) fn load_slide_images<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> SlideImageMap {
    let mut images = SlideImageMap::new();
    let slide_rels_path: String = rels_path_for(slide_path);

    let rels_xml = match read_zip_entry(archive, &slide_rels_path) {
        Ok(xml) => xml,
        Err(_) => return images,
    };

    let slide_dir: &str = slide_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let rels: HashMap<String, Relationship> = parse_relationships_xml(&rels_xml);

    for (id, rel) in rels {
        if !is_image_relationship(rel.rel_type.as_deref(), &rel.target) {
            continue;
        }

        let image_path: String = if let Some(stripped) = rel.target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, &rel.target)
        };

        if let Ok(mut file) = archive.by_name(&image_path) {
            let mut data = Vec::new();
            if file.read_to_end(&mut data).is_ok() {
                let (data, source): (Vec<u8>, SlideImageSource) =
                    normalize_slide_image_asset(&rel.target, data);
                images.insert(
                    id,
                    SlideImageAsset {
                        path: image_path,
                        data,
                        source,
                    },
                );
            }
        }
    }

    images
}

/// Pre-load SmartArt diagram data for a slide by scanning its .rels file.
pub(super) fn load_smartart_data<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> SmartArtMap {
    let mut map = SmartArtMap::new();

    let rels_xml = match read_zip_entry(archive, &rels_path_for(slide_path)) {
        Ok(xml) => xml,
        Err(_) => return map,
    };

    let slide_dir: &str = slide_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let rels: HashMap<String, String> = parse_rels_xml(&rels_xml);
    for (id, target) in &rels {
        if !target.contains("diagrams/data") && !target.contains("diagram/data") {
            continue;
        }
        let data_path: String = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, target)
        };
        if let Ok(data_xml) = read_zip_entry(archive, &data_path) {
            let texts: Vec<SmartArtNode> = smartart::parse_smartart_data_xml(&data_xml);
            if !texts.is_empty() {
                map.insert(id.clone(), texts);
            }
        }
    }

    map
}

/// Scan slide XML for chart references within graphicFrame elements.
pub(super) fn scan_chart_refs(slide_xml: &str) -> Vec<ChartRef> {
    let mut refs = Vec::new();
    let mut reader = Reader::from_str(slide_xml);

    let mut in_graphic_frame = false;
    let mut graphic_frame_x: i64 = 0;
    let mut graphic_frame_y: i64 = 0;
    let mut graphic_frame_cx: i64 = 0;
    let mut graphic_frame_cy: i64 = 0;
    let mut in_graphic_frame_transform = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element)) => match element.local_name().as_ref() {
                b"graphicFrame" if !in_graphic_frame => {
                    in_graphic_frame = true;
                    graphic_frame_x = 0;
                    graphic_frame_y = 0;
                    graphic_frame_cx = 0;
                    graphic_frame_cy = 0;
                    in_graphic_frame_transform = false;
                }
                b"xfrm" if in_graphic_frame => {
                    in_graphic_frame_transform = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref element)) => match element.local_name().as_ref() {
                b"off" if in_graphic_frame_transform => {
                    graphic_frame_x = get_attr_i64(element, b"x").unwrap_or(0);
                    graphic_frame_y = get_attr_i64(element, b"y").unwrap_or(0);
                }
                b"ext" if in_graphic_frame_transform => {
                    graphic_frame_cx = get_attr_i64(element, b"cx").unwrap_or(0);
                    graphic_frame_cy = get_attr_i64(element, b"cy").unwrap_or(0);
                }
                b"chart" if in_graphic_frame => {
                    if let Some(chart_rid) = get_attr_str(element, b"r:id") {
                        refs.push(ChartRef {
                            x: graphic_frame_x,
                            y: graphic_frame_y,
                            cx: graphic_frame_cx,
                            cy: graphic_frame_cy,
                            chart_rid,
                        });
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref element)) => match element.local_name().as_ref() {
                b"graphicFrame" if in_graphic_frame => {
                    in_graphic_frame = false;
                }
                b"xfrm" if in_graphic_frame_transform => {
                    in_graphic_frame_transform = false;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    refs
}

/// Load chart data referenced by a slide from its .rels file and the ZIP archive.
pub(super) fn load_chart_data<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> ChartMap {
    let mut charts = ChartMap::new();

    let rels_xml = match read_zip_entry(archive, &rels_path_for(slide_path)) {
        Ok(xml) => xml,
        Err(_) => return charts,
    };

    let slide_dir: &str = slide_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let rel_map: HashMap<String, String> = parse_rels_xml(&rels_xml);

    for (id, target) in &rel_map {
        let lower_target: String = target.to_ascii_lowercase();
        if !lower_target.contains("chart")
            || lower_target.contains("chartstyle")
            || lower_target.contains("chartcolor")
        {
            continue;
        }

        let chart_path: String = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, target)
        };

        if let Ok(chart_xml) = read_zip_entry(archive, &chart_path)
            && let Some(chart) = chart_parser::parse_chart_xml(&chart_xml)
        {
            charts.insert(id.clone(), chart);
        }
    }

    charts
}

/// Parse presentation.xml to extract slide size and ordered slide relationship IDs.
pub(super) fn parse_presentation_xml(xml: &str) -> Result<(PageSize, Vec<String>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut slide_size = PageSize {
        width: 720.0,
        height: 540.0,
    };
    let mut slide_rids = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref element)) => {
                handle_presentation_element(element, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Start(ref element)) => {
                handle_presentation_element(element, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(crate::parser::parse_err(format!(
                    "XML error in presentation.xml: {error}"
                )));
            }
            _ => {}
        }
    }

    Ok((slide_size, slide_rids))
}

fn handle_presentation_element(
    element: &quick_xml::events::BytesStart,
    slide_size: &mut PageSize,
    slide_rids: &mut Vec<String>,
) {
    match element.local_name().as_ref() {
        b"sldSz" => {
            let cx: i64 = get_attr_i64(element, b"cx").unwrap_or(9_144_000);
            let cy: i64 = get_attr_i64(element, b"cy").unwrap_or(6_858_000);
            *slide_size = PageSize {
                width: emu_to_pt(cx),
                height: emu_to_pt(cy),
            };
        }
        b"sldId" => {
            if let Some(rid) = get_attr_str(element, b"r:id") {
                slide_rids.push(rid);
            }
        }
        _ => {}
    }
}

fn parse_relationships_xml(xml: &str) -> HashMap<String, Relationship> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref element)) | Ok(Event::Start(ref element)) => {
                if element.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) = (
                        get_attr_str(element, b"Id"),
                        get_attr_str(element, b"Target"),
                    )
                {
                    map.insert(
                        id,
                        Relationship {
                            target,
                            rel_type: get_attr_str(element, b"Type"),
                        },
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

pub(super) fn parse_rels_xml(xml: &str) -> HashMap<String, String> {
    parse_relationships_xml(xml)
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect()
}

pub(super) fn load_theme<R: Read + std::io::Seek>(
    rel_map: &HashMap<String, String>,
    archive: &mut ZipArchive<R>,
) -> ThemeData {
    let theme_target = rel_map.values().find(|target| target.contains("theme"));
    let Some(target) = theme_target else {
        return ThemeData::default();
    };

    let theme_path: String = if let Some(stripped) = target.strip_prefix('/') {
        stripped.to_string()
    } else {
        format!("ppt/{target}")
    };

    let Ok(theme_xml) = read_zip_entry(archive, &theme_path) else {
        return ThemeData::default();
    };

    parse_theme_xml(&theme_xml)
}

/// Load and parse `ppt/tableStyles.xml` from the archive.
/// Returns an empty map if the file is missing.
pub(super) fn load_table_styles<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> table_styles::TableStyleMap {
    let Ok(xml) = read_zip_entry(archive, "ppt/tableStyles.xml") else {
        return table_styles::TableStyleMap::new();
    };
    table_styles::parse_table_styles_xml(&xml, theme, color_map)
}

pub(super) fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    crate::parser::xml_util::resolve_relative_path(base_dir, relative)
}

fn image_format_from_ext(path: &str) -> Option<ImageFormat> {
    let lower: String = path.to_ascii_lowercase();
    if lower.ends_with(".png") {
        Some(ImageFormat::Png)
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some(ImageFormat::Jpeg)
    } else if lower.ends_with(".gif") {
        Some(ImageFormat::Gif)
    } else if lower.ends_with(".bmp") {
        Some(ImageFormat::Bmp)
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        Some(ImageFormat::Tiff)
    } else if lower.ends_with(".svg") {
        Some(ImageFormat::Svg)
    } else {
        None
    }
}

fn normalize_slide_image_asset(target: &str, data: Vec<u8>) -> (Vec<u8>, SlideImageSource) {
    if let Some(format) = image_format_from_ext(target) {
        return (data, SlideImageSource::Supported(format));
    }

    if target.to_ascii_lowercase().ends_with(".emf")
        && let Some(svg) = super::emf::convert_emf_to_svg(&data)
    {
        return (svg, SlideImageSource::Supported(ImageFormat::Svg));
    }

    (data, SlideImageSource::Unsupported)
}

fn is_image_relationship(rel_type: Option<&str>, target: &str) -> bool {
    target.to_ascii_lowercase().ends_with(".emf")
        || image_format_from_ext(target).is_some()
        || rel_type.is_some_and(|rel_type| {
            let lower: String = rel_type.to_ascii_lowercase();
            lower.contains("/image") || lower.contains("hdphoto")
        })
}
