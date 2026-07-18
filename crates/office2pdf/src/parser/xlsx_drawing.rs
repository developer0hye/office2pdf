use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::ir::Chart;
use crate::parser::chart::parse_chart_xml;
use crate::parser::xml_util;

/// Extract charts from the XLSX ZIP with their anchor positions per sheet.
///
/// Returns a map from sheet name → list of (anchor_row, Chart).
/// Charts with drawing anchors get positioned at their anchor row.
/// Charts without anchors (no drawing reference found) use `u32::MAX`
/// as a sentinel to place them at the end of the sheet.
pub(super) fn extract_charts_with_anchors(data: &[u8]) -> HashMap<String, Vec<(u32, Chart)>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    // Step 1: Read workbook.xml to get sheet name → rId mapping
    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);

    // Step 2: Read workbook rels to get rId → sheet file path
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    // Step 3: For each sheet, find its drawing and extract chart anchors
    let mut result: HashMap<String, Vec<(u32, Chart)>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        // Sheet target is relative to xl/ (e.g., "worksheets/sheet1.xml")
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");

        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        if sheet_rels_xml.is_empty() {
            continue;
        }

        // Find drawing relationship
        let drawing_targets = parse_rels_by_type(&sheet_rels_xml, "drawing");
        for drawing_target in &drawing_targets {
            // Resolve relative path from worksheets/ to drawings/
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }

            // Parse drawing for chart anchor positions
            let anchors = parse_drawing_chart_anchors(&drawing_xml);

            // Find drawing rels for chart rId resolution
            let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
            let drawing_dir = drawing_path
                .rsplit_once('/')
                .map(|(d, _)| d)
                .unwrap_or("xl/drawings");
            let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
            let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
            let drawing_rid_targets = parse_rels_targets(&drawing_rels_xml);

            for (anchor_row, chart_rid) in &anchors {
                let Some(chart_target) = drawing_rid_targets.get(chart_rid) else {
                    continue;
                };
                let chart_path = resolve_relative_xl_path(drawing_dir, chart_target);
                let chart_xml = read_zip_entry_string(&mut archive, &chart_path);
                if let Some(chart) = parse_chart_xml(&chart_xml) {
                    result
                        .entry(sheet_name.clone())
                        .or_default()
                        .push((*anchor_row, chart));
                }
            }
        }
    }

    // Step 4: Find any charts not associated with drawings (orphaned charts)
    // and assign them to the first sheet with u32::MAX sentinel
    let all_positioned_chart_paths: HashSet<String> = result
        .values()
        .flatten()
        .filter_map(|_| None::<String>) // We don't track paths, just check coverage below
        .collect();
    let _ = all_positioned_chart_paths; // consumed

    // Scan for chart XML files and check if they were captured by drawing anchors
    let chart_paths: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let entry = archive.by_index(i).ok()?;
            let name = entry.name().to_string();
            if name.starts_with("xl/charts/chart") && name.ends_with(".xml") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    // Count total positioned charts
    let positioned_count: usize = result.values().map(|v| v.len()).sum();

    if chart_paths.len() > positioned_count {
        // Some charts weren't found via drawing anchors — parse them as unanchored
        let positioned_charts: HashSet<String> = collect_positioned_chart_paths(&result, data);

        let first_sheet = sheet_rids
            .first()
            .map(|(name, _)| name.clone())
            .unwrap_or_else(|| "Sheet1".to_string());

        for path in &chart_paths {
            if positioned_charts.contains(path) {
                continue;
            }
            let chart_xml = read_zip_entry_string(&mut archive, path);
            if let Some(chart) = parse_chart_xml(&chart_xml) {
                result
                    .entry(first_sheet.clone())
                    .or_default()
                    .push((u32::MAX, chart));
            }
        }
    }

    result
}

/// Collect the set of chart XML paths that were already positioned via drawing anchors.
pub(super) fn collect_positioned_chart_paths(
    chart_map: &HashMap<String, Vec<(u32, Chart)>>,
    data: &[u8],
) -> HashSet<String> {
    // Re-trace the drawing → chart resolution to find which chart paths are covered.
    // This is intentionally conservative — if we can't determine the path, we skip.
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashSet::new();
    };
    let mut positioned = HashSet::new();

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    for (sheet_name, sheet_rid) in &sheet_rids {
        if !chart_map.contains_key(sheet_name) {
            continue;
        }
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        let drawing_targets = parse_rels_by_type(&sheet_rels_xml, "drawing");

        for drawing_target in &drawing_targets {
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            let anchors = parse_drawing_chart_anchors(&drawing_xml);
            let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
            let drawing_dir = drawing_path
                .rsplit_once('/')
                .map(|(d, _)| d)
                .unwrap_or("xl/drawings");
            let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
            let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
            let drawing_rid_targets = parse_rels_targets(&drawing_rels_xml);

            for (_row, chart_rid) in &anchors {
                if let Some(chart_target) = drawing_rid_targets.get(chart_rid) {
                    positioned.insert(resolve_relative_xl_path(drawing_dir, chart_target));
                }
            }
        }
    }

    positioned
}

/// Read a ZIP entry as a string. Returns empty string if not found.
pub(super) fn read_zip_entry_string(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    path: &str,
) -> String {
    let Ok(mut entry) = archive.by_name(path) else {
        return String::new();
    };
    let mut xml = String::new();
    let _ = std::io::Read::read_to_string(&mut entry, &mut xml);
    xml
}

/// Parse workbook.xml to extract sheet name → rId pairs (preserving order).
pub(super) fn parse_workbook_sheet_rids(xml: &str) -> Vec<(String, String)> {
    let mut result = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"sheet" {
                    let mut name = None;
                    let mut rid = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"name" => {
                                if let Ok(v) = attr.unescape_value() {
                                    name = Some(v.to_string());
                                }
                            }
                            b"id" => {
                                if let Ok(v) = attr.unescape_value() {
                                    rid = Some(v.to_string());
                                }
                            }
                            _ => {}
                        }
                    }
                    if let (Some(n), Some(r)) = (name, rid) {
                        result.push((n, r));
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

/// Parse a .rels file to get Id → Target mapping.
pub(super) fn parse_rels_targets(xml: &str) -> HashMap<String, String> {
    xml_util::parse_rels_id_target(xml)
}

/// Parse a .rels file and return targets whose Type contains the given substring.
pub(super) fn parse_rels_by_type(xml: &str, type_substring: &str) -> Vec<String> {
    let mut targets = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut target = None;
                    let mut matches_type = false;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Target" => {
                                if let Ok(v) = attr.unescape_value() {
                                    target = Some(v.to_string());
                                }
                            }
                            b"Type" => {
                                if let Ok(v) = attr.unescape_value()
                                    && v.contains(type_substring)
                                {
                                    matches_type = true;
                                }
                            }
                            _ => {}
                        }
                    }
                    if matches_type && let Some(t) = target {
                        targets.push(t);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    targets
}

/// Resolve a relative path (like `../drawings/drawing1.xml`) against a base directory.
pub(super) fn resolve_relative_xl_path(base_dir: &str, relative: &str) -> String {
    xml_util::resolve_relative_path(base_dir, relative)
}

/// Parse drawing XML for chart anchor positions.
/// Returns (anchor_row, chart_rId) pairs from `<xdr:twoCellAnchor>` elements.
pub(super) fn parse_drawing_chart_anchors(xml: &str) -> Vec<(u32, String)> {
    let mut result = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_two_cell_anchor = false;
    let mut in_from = false;
    let mut in_row = false;
    let mut anchor_row: Option<u32> = None;
    let mut chart_rid: Option<String> = None;
    let mut in_graphic_data = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
                        in_two_cell_anchor = true;
                        anchor_row = None;
                        chart_rid = None;
                    }
                    b"from" if in_two_cell_anchor => {
                        in_from = true;
                    }
                    b"row" if in_from => {
                        in_row = true;
                    }
                    b"graphicData" if in_two_cell_anchor => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"uri"
                                && let Ok(val) = attr.unescape_value()
                                && val.contains("chart")
                            {
                                in_graphic_data = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_graphic_data && local.as_ref() == b"chart" {
                    for attr in e.attributes().flatten() {
                        if (attr.key.as_ref() == b"r:id" || attr.key.local_name().as_ref() == b"id")
                            && let Ok(val) = attr.unescape_value()
                        {
                            chart_rid = Some(val.to_string());
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if in_row
                    && let Ok(s) = t.xml_content()
                    && let Ok(row) = s.trim().parse::<u32>()
                {
                    anchor_row = Some(row);
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
                        if let (Some(row), Some(rid)) = (anchor_row.take(), chart_rid.take()) {
                            result.push((row, rid));
                        }
                        in_two_cell_anchor = false;
                        in_from = false;
                        in_graphic_data = false;
                    }
                    b"from" => {
                        in_from = false;
                    }
                    b"row" => {
                        in_row = false;
                    }
                    b"graphicData" => {
                        in_graphic_data = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

// ── Drawing images ──────────────────────────────────────────────────────

/// A picture anchor from a worksheet drawing, in raw drawing coordinates.
/// Rows/columns are 0-indexed as in the XML; offsets and extents are EMU.
pub(super) struct RawImageAnchor {
    pub(super) from_row: u32,
    pub(super) from_col: u32,
    pub(super) from_col_off_emu: i64,
    pub(super) from_row_off_emu: i64,
    /// twoCellAnchor bottom-right corner: (col, col_off, row, row_off).
    pub(super) to: Option<(u32, i64, u32, i64)>,
    /// oneCellAnchor extent (cx, cy).
    pub(super) ext_emu: Option<(i64, i64)>,
    pub(super) data: Vec<u8>,
    pub(super) format: crate::ir::ImageFormat,
}

/// Extract anchored pictures per sheet from worksheet drawings.
/// Metafiles (EMF/WMF) are converted to SVG; unknown formats are skipped.
pub(super) fn extract_images_with_anchors(data: &[u8]) -> HashMap<String, Vec<RawImageAnchor>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    let mut result: HashMap<String, Vec<RawImageAnchor>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        if sheet_rels_xml.is_empty() {
            continue;
        }

        for drawing_target in &parse_rels_by_type(&sheet_rels_xml, "drawing") {
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }

            let anchors = parse_drawing_image_anchors(&drawing_xml);
            if anchors.is_empty() {
                continue;
            }

            let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
            let drawing_dir = drawing_path
                .rsplit_once('/')
                .map(|(dir, _)| dir)
                .unwrap_or("xl/drawings");
            let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
            let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
            let rid_to_media = parse_rels_targets(&drawing_rels_xml);

            for (geometry, rid) in anchors {
                let Some(media_target) = rid_to_media.get(&rid) else {
                    continue;
                };
                let media_path = resolve_relative_xl_path(drawing_dir, media_target);
                let Some(bytes) = read_zip_entry_bytes(&mut archive, &media_path) else {
                    continue;
                };
                let Some((data, format)) = decode_media(&media_path, bytes) else {
                    continue;
                };
                result
                    .entry(sheet_name.clone())
                    .or_default()
                    .push(RawImageAnchor {
                        from_row: geometry.from_row,
                        from_col: geometry.from_col,
                        from_col_off_emu: geometry.from_col_off_emu,
                        from_row_off_emu: geometry.from_row_off_emu,
                        to: geometry.to,
                        ext_emu: geometry.ext_emu,
                        data,
                        format,
                    });
            }
        }
    }

    result
}

fn read_zip_entry_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<Vec<u8>> {
    use std::io::Read;
    let mut file = archive.by_name(path).ok()?;
    let mut buffer = Vec::new();
    file.read_to_end(&mut buffer).ok()?;
    Some(buffer)
}

/// Map media bytes to a renderable (data, format) pair; metafiles are
/// converted to SVG.
fn decode_media(path: &str, bytes: Vec<u8>) -> Option<(Vec<u8>, crate::ir::ImageFormat)> {
    use crate::ir::ImageFormat;
    let extension: String = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    match extension.as_str() {
        "png" => Some((bytes, ImageFormat::Png)),
        "jpg" | "jpeg" => Some((bytes, ImageFormat::Jpeg)),
        "gif" => Some((bytes, ImageFormat::Gif)),
        "bmp" => Some((bytes, ImageFormat::Bmp)),
        "tif" | "tiff" => Some((bytes, ImageFormat::Tiff)),
        "svg" => Some((bytes, ImageFormat::Svg)),
        "emf" => crate::parser::emf::convert_emf_to_svg(&bytes).map(|svg| (svg, ImageFormat::Svg)),
        "wmf" => crate::parser::wmf::convert_wmf_to_svg(&bytes).map(|svg| (svg, ImageFormat::Svg)),
        _ => None,
    }
}

/// Geometry captured from a single pic anchor before media resolution.
pub(super) struct ImageAnchorGeometry {
    pub(super) from_row: u32,
    pub(super) from_col: u32,
    pub(super) from_col_off_emu: i64,
    pub(super) from_row_off_emu: i64,
    pub(super) to: Option<(u32, i64, u32, i64)>,
    pub(super) ext_emu: Option<(i64, i64)>,
}

/// Parse `<xdr:pic>` anchors from a worksheet drawing: anchor geometry plus
/// the blip relationship id.
pub(super) fn parse_drawing_image_anchors(xml: &str) -> Vec<(ImageAnchorGeometry, String)> {
    #[derive(Default, Clone, Copy)]
    struct Corner {
        col: u32,
        col_off: i64,
        row: u32,
        row_off: i64,
    }

    let mut result: Vec<(ImageAnchorGeometry, String)> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_anchor = false;
    let mut in_pic = false;
    let mut corner_target: Option<bool> = None; // Some(true)=from, Some(false)=to
    let mut current_field: Option<&'static str> = None;
    let mut from = Corner::default();
    let mut to: Option<Corner> = None;
    let mut ext_emu: Option<(i64, i64)> = None;
    let mut blip_rid: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    in_anchor = true;
                    in_pic = false;
                    from = Corner::default();
                    to = None;
                    ext_emu = None;
                    blip_rid = None;
                }
                b"from" if in_anchor => corner_target = Some(true),
                b"to" if in_anchor => {
                    corner_target = Some(false);
                    to = Some(Corner::default());
                }
                b"col" if corner_target.is_some() => current_field = Some("col"),
                b"colOff" if corner_target.is_some() => current_field = Some("colOff"),
                b"row" if corner_target.is_some() => current_field = Some("row"),
                b"rowOff" if corner_target.is_some() => current_field = Some("rowOff"),
                b"pic" if in_anchor => in_pic = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_anchor && local.as_ref() == b"ext" && !in_pic && to.is_none() {
                    // oneCellAnchor extent (the pic's own a:ext lives inside
                    // xfrm, which we ignore by requiring !in_pic).
                    let mut cx: i64 = 0;
                    let mut cy: i64 = 0;
                    for attr in e.attributes().flatten() {
                        let value: i64 = attr
                            .unescape_value()
                            .ok()
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        match attr.key.local_name().as_ref() {
                            b"cx" => cx = value,
                            b"cy" => cy = value,
                            _ => {}
                        }
                    }
                    ext_emu = Some((cx, cy));
                }
                if in_pic && local.as_ref() == b"blip" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"embed"
                            && let Ok(val) = attr.unescape_value()
                        {
                            blip_rid = Some(val.to_string());
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content()
                    && let Ok(number) = text.trim().parse::<i64>()
                {
                    let corner: &mut Corner = if is_from {
                        &mut from
                    } else {
                        to.as_mut().expect("to corner initialized on <to>")
                    };
                    match field {
                        "col" => corner.col = number as u32,
                        "colOff" => corner.col_off = number,
                        "row" => corner.row = number as u32,
                        "rowOff" => corner.row_off = number,
                        _ => {}
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    if in_pic && let Some(rid) = blip_rid.take() {
                        result.push((
                            ImageAnchorGeometry {
                                from_row: from.row,
                                from_col: from.col,
                                from_col_off_emu: from.col_off,
                                from_row_off_emu: from.row_off,
                                to: to.map(|c| (c.col, c.col_off, c.row, c.row_off)),
                                ext_emu,
                            },
                            rid,
                        ));
                    }
                    in_anchor = false;
                    in_pic = false;
                    corner_target = None;
                }
                b"from" | b"to" => corner_target = None,
                b"col" | b"colOff" | b"row" | b"rowOff" => current_field = None,
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

// ── Drawing text boxes ──────────────────────────────────────────────────

/// A text-box shape from a worksheet drawing, in raw drawing coordinates.
pub(super) struct RawTextBoxAnchor {
    pub(super) geometry: ImageAnchorGeometry,
    pub(super) paragraphs: Vec<crate::ir::Paragraph>,
    pub(super) fill: Option<crate::ir::Color>,
    pub(super) border: Option<crate::ir::BorderSide>,
    pub(super) vertical_center: bool,
}

/// Extract anchored text boxes per sheet from worksheet drawings.
pub(super) fn extract_text_boxes_with_anchors(
    data: &[u8],
) -> HashMap<String, Vec<RawTextBoxAnchor>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    let mut result: HashMap<String, Vec<RawTextBoxAnchor>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        if sheet_rels_xml.is_empty() {
            continue;
        }

        for drawing_target in &parse_rels_by_type(&sheet_rels_xml, "drawing") {
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }
            let boxes = parse_drawing_text_boxes(&drawing_xml);
            if !boxes.is_empty() {
                result.entry(sheet_name.clone()).or_default().extend(boxes);
            }
        }
    }

    result
}

/// Resolve a drawing color element to RGB. Scheme colors fall back to the
/// standard light/dark mapping (worksheet drawings rarely restyle them).
fn drawing_color(name: &[u8], val: &str) -> Option<crate::ir::Color> {
    use crate::ir::Color;
    match name {
        b"srgbClr" => {
            let v = u32::from_str_radix(val, 16).ok()?;
            Some(Color::new(
                ((v >> 16) & 0xFF) as u8,
                ((v >> 8) & 0xFF) as u8,
                (v & 0xFF) as u8,
            ))
        }
        b"schemeClr" => match val {
            "lt1" | "bg1" | "lt2" | "bg2" => Some(Color::new(0xFF, 0xFF, 0xFF)),
            "dk1" | "tx1" | "dk2" | "tx2" => Some(Color::new(0, 0, 0)),
            _ => None,
        },
        _ => None,
    }
}

/// Parse `<xdr:sp>` text boxes from a worksheet drawing.
pub(super) fn parse_drawing_text_boxes(xml: &str) -> Vec<RawTextBoxAnchor> {
    use crate::ir::{
        Alignment, BorderLineStyle, BorderSide, Paragraph, ParagraphStyle, Run, TextStyle,
    };

    #[derive(Default, Clone, Copy)]
    struct Corner {
        col: u32,
        col_off: i64,
        row: u32,
        row_off: i64,
    }

    let mut result: Vec<RawTextBoxAnchor> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_anchor = false;
    let mut in_sp = false;
    let mut in_tx_body = false;
    let mut in_sp_fill = false;
    let mut in_line = false;
    let mut corner_target: Option<bool> = None;
    let mut current_field: Option<&'static str> = None;
    let mut from = Corner::default();
    let mut to: Option<Corner> = None;
    let mut ext_emu: Option<(i64, i64)> = None;

    let mut paragraphs: Vec<Paragraph> = Vec::new();
    let mut current_para: Option<Paragraph> = None;
    let mut current_style = TextStyle::default();
    let mut in_run = false;
    let mut in_text = false;
    let mut fill: Option<crate::ir::Color> = None;
    let mut border_color: Option<crate::ir::Color> = None;
    let mut border_width: f64 = 0.75;
    let mut vertical_center = false;
    // Color element opened with children (e.g. <a:schemeClr><a:shade/>...):
    // committed on its End event after transforms apply.
    let mut pending_color: Option<crate::ir::Color> = None;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        in_anchor = true;
                        in_sp = false;
                        from = Corner::default();
                        to = None;
                        ext_emu = None;
                        paragraphs.clear();
                        fill = None;
                        border_color = None;
                        border_width = 0.75;
                        vertical_center = false;
                    }
                    b"from" if in_anchor => corner_target = Some(true),
                    b"to" if in_anchor => {
                        corner_target = Some(false);
                        to = Some(Corner::default());
                    }
                    b"col" if corner_target.is_some() => current_field = Some("col"),
                    b"colOff" if corner_target.is_some() => current_field = Some("colOff"),
                    b"row" if corner_target.is_some() => current_field = Some("row"),
                    b"rowOff" if corner_target.is_some() => current_field = Some("rowOff"),
                    b"sp" if in_anchor => in_sp = true,
                    b"txBody" if in_sp => in_tx_body = true,
                    b"solidFill" if in_sp && !in_tx_body && !in_line => in_sp_fill = true,
                    b"ln" if in_sp && !in_tx_body => {
                        in_line = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"w"
                                && let Ok(v) = attr.unescape_value()
                                && let Ok(w) = v.parse::<f64>()
                            {
                                border_width = w / 12_700.0;
                            }
                        }
                    }
                    b"p" if in_tx_body => {
                        current_para = Some(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: Vec::new(),
                        });
                    }
                    b"r" if current_para.is_some() => {
                        in_run = true;
                        current_style = TextStyle::default();
                    }
                    b"rPr" if in_run => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"sz" => {
                                    if let Ok(v) = attr.unescape_value()
                                        && let Ok(sz) = v.parse::<f64>()
                                    {
                                        current_style.font_size = Some(sz / 100.0);
                                    }
                                }
                                b"b" if attr.unescape_value().ok().as_deref() == Some("1") => {
                                    current_style.bold = Some(true);
                                }
                                b"i" if attr.unescape_value().ok().as_deref() == Some("1") => {
                                    current_style.italic = Some(true);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"t" if in_run => in_text = true,
                    b"srgbClr" | b"schemeClr" => {
                        let mut val = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(v) = attr.unescape_value()
                            {
                                val = v.to_string();
                            }
                        }
                        pending_color = drawing_color(local.as_ref(), &val);
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bodyPr" if in_tx_body || in_sp => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"anchor"
                                && attr.unescape_value().ok().as_deref() == Some("ctr")
                            {
                                vertical_center = true;
                            }
                        }
                    }
                    b"pPr" if current_para.is_some() => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"algn"
                                && let Ok(v) = attr.unescape_value()
                                && let Some(para) = current_para.as_mut()
                            {
                                para.style.alignment = match v.as_ref() {
                                    "ctr" => Some(Alignment::Center),
                                    "r" => Some(Alignment::Right),
                                    "just" => Some(Alignment::Justify),
                                    _ => None,
                                };
                            }
                        }
                    }
                    b"shade" => {
                        if let Some(color) = pending_color.as_mut()
                            && let Some(factor) = e.attributes().flatten().find_map(|attr| {
                                (attr.key.local_name().as_ref() == b"val")
                                    .then(|| attr.unescape_value().ok())
                                    .flatten()
                                    .and_then(|v| v.parse::<f64>().ok())
                            })
                        {
                            let shade = |channel: u8| -> u8 {
                                (channel as f64 * factor / 100_000.0)
                                    .round()
                                    .clamp(0.0, 255.0) as u8
                            };
                            *color = crate::ir::Color::new(
                                shade(color.r),
                                shade(color.g),
                                shade(color.b),
                            );
                        }
                    }
                    b"srgbClr" | b"schemeClr" => {
                        let mut val = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(v) = attr.unescape_value()
                            {
                                val = v.to_string();
                            }
                        }
                        let color = drawing_color(local.as_ref(), &val);
                        if in_run {
                            if current_style.color.is_none() {
                                current_style.color = color;
                            }
                        } else if in_line {
                            if border_color.is_none() {
                                border_color = color;
                            }
                        } else if in_sp_fill && fill.is_none() {
                            fill = color;
                        }
                    }
                    b"ext" if in_anchor && !in_sp && to.is_none() => {
                        let mut cx: i64 = 0;
                        let mut cy: i64 = 0;
                        for attr in e.attributes().flatten() {
                            let value: i64 = attr
                                .unescape_value()
                                .ok()
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0);
                            match attr.key.local_name().as_ref() {
                                b"cx" => cx = value,
                                b"cy" => cy = value,
                                _ => {}
                            }
                        }
                        ext_emu = Some((cx, cy));
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if in_text && let Ok(text) = t.xml_content() {
                    if let Some(para) = current_para.as_mut() {
                        para.runs.push(Run {
                            text: text.to_string(),
                            style: current_style.clone(),
                            href: None,
                            footnote: None,
                        });
                    }
                } else if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content()
                    && let Ok(number) = text.trim().parse::<i64>()
                {
                    let corner: &mut Corner = if is_from {
                        &mut from
                    } else {
                        to.as_mut().expect("to corner initialized on <to>")
                    };
                    match field {
                        "col" => corner.col = number as u32,
                        "colOff" => corner.col_off = number,
                        "row" => corner.row = number as u32,
                        "rowOff" => corner.row_off = number,
                        _ => {}
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        if in_sp && !paragraphs.is_empty() {
                            result.push(RawTextBoxAnchor {
                                geometry: ImageAnchorGeometry {
                                    from_row: from.row,
                                    from_col: from.col,
                                    from_col_off_emu: from.col_off,
                                    from_row_off_emu: from.row_off,
                                    to: to.map(|c| (c.col, c.col_off, c.row, c.row_off)),
                                    ext_emu,
                                },
                                paragraphs: std::mem::take(&mut paragraphs),
                                fill,
                                border: border_color.map(|color| BorderSide {
                                    width: border_width,
                                    color,
                                    style: BorderLineStyle::Solid,
                                }),
                                vertical_center,
                            });
                        }
                        in_anchor = false;
                        in_sp = false;
                        corner_target = None;
                    }
                    b"from" | b"to" => corner_target = None,
                    b"col" | b"colOff" | b"row" | b"rowOff" => current_field = None,
                    b"srgbClr" | b"schemeClr" => {
                        if let Some(color) = pending_color.take() {
                            if in_run {
                                if current_style.color.is_none() {
                                    current_style.color = Some(color);
                                }
                            } else if in_line {
                                if border_color.is_none() {
                                    border_color = Some(color);
                                }
                            } else if in_sp_fill && fill.is_none() {
                                fill = Some(color);
                            }
                        }
                    }
                    b"txBody" => in_tx_body = false,
                    b"solidFill" => in_sp_fill = false,
                    b"ln" => in_line = false,
                    b"p" => {
                        if let Some(para) = current_para.take() {
                            paragraphs.push(para);
                        }
                    }
                    b"r" => in_run = false,
                    b"t" => in_text = false,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}
