use crate::parser::xml_util::OOXML_XML_VERSION;
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::ir::Chart;
use crate::ir::LineCap;
use crate::parser::chart_style::parse_chart_with_style;
use crate::parser::drawingml::ThemeFontScheme;
use crate::parser::xml_util;

/// A chart from a worksheet drawing, in raw drawing coordinates.
pub(super) struct RawChartAnchor {
    /// The anchor's geometry, or `None` for a chart no drawing references —
    /// which flows after the sheet's rows instead of overlaying them.
    pub(super) geometry: Option<ImageAnchorGeometry>,
    pub(super) chart: Chart,
}

/// Extract charts from the XLSX ZIP with their anchor geometry per sheet.
///
/// Returns a map from sheet name → list of charts. A chart a drawing anchors
/// carries that anchor's geometry, so it can be overlaid on the grid at
/// absolute worksheet coordinates the way pictures are (issue #982). A chart
/// no drawing references carries none and is placed after the sheet's rows.
pub(super) fn extract_charts_with_anchors(data: &[u8]) -> HashMap<String, Vec<RawChartAnchor>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    // Step 1: Read workbook.xml to get sheet name → rId mapping
    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);

    // Step 2: Read workbook rels to get rId → sheet file path
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    // A series that names no fill takes its colour from the workbook's theme,
    // not from the renderer's built-in palette (issue #670).
    // The same part settles chart text's face: a chart naming no `a:latin`
    // takes the theme's minor font rather than the engine's default (#668).
    let (theme_colors, theme_fonts, theme_line_styles) =
        workbook_theme(&mut archive, &workbook_rels_xml);
    let theme_accents: Vec<crate::ir::Color> =
        crate::parser::drawingml::theme_accent_palette(&theme_colors);
    // A series that names its fill through `<a:schemeClr>` resolves it against
    // the same workbook theme (issue #876).
    let no_aliases: std::collections::HashMap<String, String> = std::collections::HashMap::new();
    let scheme = crate::parser::drawingml::SchemeColors {
        colors: &theme_colors,
        aliases: &no_aliases,
    };
    // A chart's `<c:userShapes>` part is an ordinary drawing, so its
    // `<a:schemeClr>` resolves the way a worksheet drawing's does — through
    // the background/text aliases a spreadsheet has no `<clrMap>` to state
    // (issue #1186).
    let drawing_aliases: HashMap<String, String> = xlsx_scheme_aliases();
    let drawing_scheme = crate::parser::drawingml::SchemeColors {
        colors: &theme_colors,
        aliases: &drawing_aliases,
    };

    // Step 3: For each sheet, find its drawing and extract chart anchors
    let mut result: HashMap<String, Vec<RawChartAnchor>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        // Sheet target is relative to xl/ (e.g., "worksheets/sheet1.xml")
        let sheet_dir: String = sheet_part_dir(sheet_target);
        let sheet_xml = read_zip_entry_string(&mut archive, &sheet_part_path(sheet_target));
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path(sheet_target));
        if sheet_rels_xml.is_empty() {
            continue;
        }

        // Find the drawings the sheet body names
        let drawing_targets = sheet_drawing_targets(&sheet_xml, &sheet_rels_xml);
        for drawing_target in &drawing_targets {
            // Resolve the drawing relative to the sheet's own directory
            let drawing_path = resolve_relative_xl_path(&sheet_dir, drawing_target);
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

            for (geometry, chart_rid) in anchors {
                let Some(chart_target) = drawing_rid_targets.get(&chart_rid) else {
                    continue;
                };
                let chart_path = resolve_relative_xl_path(drawing_dir, chart_target);
                let chart_xml = read_zip_entry_string(&mut archive, &chart_path);
                if let Some(mut chart) =
                    parse_chart_with_style(&mut archive, &chart_path, &chart_xml, &scheme)
                {
                    chart.theme_accent_colors = theme_accents.clone();
                    chart.host = crate::ir::ChartHost::Spreadsheet;
                    theme_fonts.resolve_chart_text_fonts(&mut chart);
                    chart.user_shapes = crate::parser::chart_drawing::load_chart_user_shapes(
                        &mut archive,
                        &chart_path,
                        &chart_xml,
                        &drawing_scheme,
                        &theme_fonts,
                        &theme_line_styles,
                    );
                    result
                        .entry(sheet_name.clone())
                        .or_default()
                        .push(RawChartAnchor {
                            geometry: Some(geometry),
                            chart,
                        });
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
        let positioned_charts: HashSet<String> = collect_drawing_reachable_chart_paths(data);

        let first_sheet = sheet_rids
            .first()
            .map(|(name, _)| name.clone())
            .unwrap_or_else(|| "Sheet1".to_string());

        for path in &chart_paths {
            if positioned_charts.contains(path) {
                continue;
            }
            let chart_xml = read_zip_entry_string(&mut archive, path);
            if let Some(mut chart) = parse_chart_with_style(&mut archive, path, &chart_xml, &scheme)
            {
                chart.theme_accent_colors = theme_accents.clone();
                chart.host = crate::ir::ChartHost::Spreadsheet;
                theme_fonts.resolve_chart_text_fonts(&mut chart);
                chart.user_shapes = crate::parser::chart_drawing::load_chart_user_shapes(
                    &mut archive,
                    path,
                    &chart_xml,
                    &drawing_scheme,
                    &theme_fonts,
                    &theme_line_styles,
                );
                result
                    .entry(first_sheet.clone())
                    .or_default()
                    .push(RawChartAnchor {
                        geometry: None,
                        chart,
                    });
            }
        }
    }

    result
}

/// Every chart XML path a sheet's drawing part reaches, whether or not the
/// sheet prints that drawing.
///
/// The unanchored fallback above exists for a chart part no drawing reaches at
/// all. A chart behind a drawing whose relationship the sheet body never names
/// is a different case: Excel prints nothing for it (issue #1158), so it must
/// count as reached here rather than reappear after the grid.
pub(super) fn collect_drawing_reachable_chart_paths(data: &[u8]) -> HashSet<String> {
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

    for (_, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_dir: String = sheet_part_dir(sheet_target);
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path(sheet_target));
        let drawing_targets = parse_rels_by_type(&sheet_rels_xml, "drawing");

        for drawing_target in &drawing_targets {
            let drawing_path = resolve_relative_xl_path(&sheet_dir, drawing_target);
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

            for (_geometry, chart_rid) in &anchors {
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
                                if let Ok(v) = attr.normalized_value(OOXML_XML_VERSION) {
                                    name = Some(v.to_string());
                                }
                            }
                            b"id" => {
                                if let Ok(v) = attr.normalized_value(OOXML_XML_VERSION) {
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
pub(in crate::parser) fn parse_rels_targets(xml: &str) -> HashMap<String, String> {
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
                                if let Ok(v) = attr.normalized_value(OOXML_XML_VERSION) {
                                    target = Some(v.to_string());
                                }
                            }
                            b"Type" => {
                                if let Ok(v) = attr.normalized_value(OOXML_XML_VERSION)
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

/// The directory holding a sheet part, given the target `xl/workbook.xml.rels`
/// declares for it. Drawing relationships are relative to this, not to
/// `xl/worksheets`.
///
/// Not every sheet is a worksheet: a chartsheet's part lives under
/// `xl/chartsheets/`, and assuming the worksheet directory both missed its
/// drawing outright and — where a worksheet part shares its filename, which
/// Excel's `sheet1.xml` numbering makes the common case — resolved it to the
/// *worksheet's* rels, so the chartsheet drew a second copy of the worksheet's
/// chart (issue #1099).
pub(super) fn sheet_part_dir(sheet_target: &str) -> String {
    let full_path: String = resolve_relative_xl_path("xl", sheet_target);
    full_path
        .rsplit_once('/')
        .map(|(dir, _)| dir.to_string())
        .unwrap_or_else(|| "xl".to_string())
}

/// The `_rels` part accompanying a sheet part, in the sheet's own directory.
pub(super) fn sheet_rels_path(sheet_target: &str) -> String {
    let full_path: String = resolve_relative_xl_path("xl", sheet_target);
    let filename: &str = full_path.rsplit('/').next().unwrap_or(&full_path);
    format!("{}/_rels/{filename}.rels", sheet_part_dir(sheet_target))
}

/// The sheet part itself, given the target `xl/workbook.xml.rels` declares.
pub(super) fn sheet_part_path(sheet_target: &str) -> String {
    resolve_relative_xl_path("xl", sheet_target)
}

/// The drawing parts one sheet prints, as relationship targets in the order
/// its body names them.
///
/// A drawing is attached to a sheet by its `<drawing r:id="...">` element; the
/// relationship on its own attaches nothing.
/// `tests/fixtures/xlsx/issue_1158_unreferenced_drawing_rel.xlsx` is
/// `issue_1066_blip_effect_picture.xlsx` with that one element removed and the
/// relationship, `xl/drawings/drawing1.xml`, `xl/media/image1.png` and the
/// content types all left in place: Excel for Mac prints it as the grid alone,
/// where reading the relationships instead still drew the picture (issue
/// #1158).
///
/// Going through the element also stops `drawingHF` coming through, whose
/// relationship type contains `drawing` as a substring: a header/footer
/// drawing is named by its own `<drawingHF>` element and is not part of the
/// grid.
pub(super) fn sheet_drawing_targets(sheet_xml: &str, sheet_rels_xml: &str) -> Vec<String> {
    let rid_to_target: HashMap<String, String> = parse_rels_targets(sheet_rels_xml);
    parse_sheet_drawing_rids(sheet_xml)
        .into_iter()
        .filter_map(|rid| rid_to_target.get(&rid).cloned())
        .collect()
}

/// The relationship ids a sheet body's `<drawing>` elements name, in order.
///
/// `CT_Worksheet` and `CT_Chartsheet` both carry the element, so a worksheet
/// and a chartsheet resolve their drawings the same way.
fn parse_sheet_drawing_rids(sheet_xml: &str) -> Vec<String> {
    let mut rids: Vec<String> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(sheet_xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                // `legacyDrawing`, `legacyDrawingHF` and `drawingHF` are
                // elements of their own, so the local name has to match
                // exactly rather than by prefix.
                if e.local_name().as_ref() == b"drawing"
                    && let Some(rid) = xml_util::get_attr_str(e, b"id")
                {
                    rids.push(rid);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    rids
}

/// Parse `<xdr:graphicFrame>` chart anchors from a worksheet drawing: the
/// anchor's geometry plus the chart relationship id.
///
/// The whole `from`/`to` span is read, not just the anchor row: Excel floats
/// the chart over the cells at those coordinates and sizes it to them, so the
/// renderer needs the same geometry a picture anchor already carries (issue
/// #982).
pub(super) fn parse_drawing_chart_anchors(xml: &str) -> Vec<(ImageAnchorGeometry, String)> {
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
    let mut in_graphic_frame = false;
    let mut in_chart_graphic_data = false;
    let mut corner_target: Option<bool> = None; // Some(true)=from, Some(false)=to
    let mut current_field: Option<&'static str> = None;
    let mut from = Corner::default();
    let mut to: Option<Corner> = None;
    let mut ext_emu: Option<(i64, i64)> = None;
    let mut chart_rid: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    in_anchor = true;
                    in_graphic_frame = false;
                    in_chart_graphic_data = false;
                    from = Corner::default();
                    to = None;
                    ext_emu = None;
                    chart_rid = None;
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
                b"graphicFrame" if in_anchor => in_graphic_frame = true,
                b"graphicData" if in_graphic_frame => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"uri"
                            && let Ok(val) = attr.normalized_value(OOXML_XML_VERSION)
                            && val.contains("chart")
                        {
                            in_chart_graphic_data = true;
                        }
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_anchor && local.as_ref() == b"ext" && !in_graphic_frame && to.is_none() {
                    // The anchor's own extent (a oneCellAnchor's or an
                    // absoluteAnchor's). The frame's `a:ext` lives inside
                    // `xdr:xfrm`, which `!in_graphic_frame` excludes.
                    let mut cx: i64 = 0;
                    let mut cy: i64 = 0;
                    for attr in e.attributes().flatten() {
                        let value: i64 = attr
                            .normalized_value(OOXML_XML_VERSION)
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
                if in_anchor && local.as_ref() == b"pos" && !in_graphic_frame {
                    // An absoluteAnchor measures from the sheet origin, which
                    // is column 0 / row 0 plus the position as the offset.
                    for attr in e.attributes().flatten() {
                        let value: i64 = attr
                            .normalized_value(OOXML_XML_VERSION)
                            .ok()
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        match attr.key.local_name().as_ref() {
                            b"x" => from.col_off = value,
                            b"y" => from.row_off = value,
                            _ => {}
                        }
                    }
                }
                if in_chart_graphic_data && local.as_ref() == b"chart" {
                    for attr in e.attributes().flatten() {
                        if (attr.key.as_ref() == b"r:id" || attr.key.local_name().as_ref() == b"id")
                            && let Ok(val) = attr.normalized_value(OOXML_XML_VERSION)
                        {
                            chart_rid = Some(val.to_string());
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content(OOXML_XML_VERSION)
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
                    if let Some(rid) = chart_rid.take() {
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
                    in_graphic_frame = false;
                    in_chart_graphic_data = false;
                    corner_target = None;
                }
                b"from" | b"to" => corner_target = None,
                b"col" | b"colOff" | b"row" | b"rowOff" => current_field = None,
                b"graphicFrame" => in_graphic_frame = false,
                b"graphicData" => in_chart_graphic_data = false,
                _ => {}
            },
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
/// Raster parts are sniffed and fully decoded; invalid rasters and unknown
/// formats are omitted with an unsupported-element warning. Metafiles
/// (EMF/WMF) are converted to SVG, with conversion failures omitted likewise.
pub(super) fn extract_images_with_anchors(
    data: &[u8],
    warnings: &mut Vec<crate::error::ConvertWarning>,
) -> HashMap<String, Vec<RawImageAnchor>> {
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
        let sheet_dir: String = sheet_part_dir(sheet_target);
        let sheet_xml = read_zip_entry_string(&mut archive, &sheet_part_path(sheet_target));
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path(sheet_target));
        if sheet_rels_xml.is_empty() {
            continue;
        }

        for drawing_target in &sheet_drawing_targets(&sheet_xml, &sheet_rels_xml) {
            let drawing_path = resolve_relative_xl_path(&sheet_dir, drawing_target);
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

            for (geometry, blip) in anchors {
                let Some(media_target) = rid_to_media.get(&blip.rid) else {
                    continue;
                };
                let media_path = resolve_relative_xl_path(drawing_dir, media_target);
                let Some(bytes) = read_zip_entry_bytes(&mut archive, &media_path) else {
                    continue;
                };
                let Some((data, format)) = decode_media(&media_path, bytes) else {
                    warnings.push(crate::error::ConvertWarning::UnsupportedElement {
                        format: "XLSX".to_string(),
                        element: format!("image omitted: {media_path}"),
                    });
                    continue;
                };
                // Excel composites a picture carrying `<a:alphaModFix>` onto
                // the worksheet at that strength. Typst has no per-image
                // opacity, so the factor is baked into the pixels the way the
                // pptx picture path already bakes it (issue #1103).
                let (data, format) = match blip.alpha {
                    Some(alpha) if alpha < 1.0 => {
                        crate::parser::drawingml::apply_image_alpha(&data, alpha)
                            .unwrap_or((data, format))
                    }
                    _ => (data, format),
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
        "png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" => {
            crate::parser::drawingml::validated_raster_format(&bytes).map(|format| (bytes, format))
        }
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

/// A `<xdr:pic>`'s `<a:blip>`: the media it embeds and how solidly Excel
/// draws it.
pub(super) struct PictureBlip {
    /// The `r:embed` relationship id naming the media part.
    pub(super) rid: String,
    /// `<a:alphaModFix amt>` as a 0.0-1.0 factor, when the picture carries a
    /// transparency effect.
    pub(super) alpha: Option<f64>,
}

/// Relationship id in an `<a:blip>`'s `r:embed`, if it carries one.
///
/// Read from the `a:blip` element itself and never from a descendant: the
/// `a14:imgProps` extension Excel nests inside an effect-bearing blip declares
/// its own `r:embed` for the high-definition layer, which is not the picture
/// Excel draws.
fn blip_embed_rid(element: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    element.attributes().flatten().find_map(|attr| {
        if attr.key.local_name().as_ref() == b"embed" {
            attr.normalized_value(OOXML_XML_VERSION)
                .ok()
                .map(|value| value.to_string())
        } else {
            None
        }
    })
}

/// Parse `<xdr:pic>` anchors from a worksheet drawing: anchor geometry plus
/// the picture's blip.
pub(super) fn parse_drawing_image_anchors(xml: &str) -> Vec<(ImageAnchorGeometry, PictureBlip)> {
    #[derive(Default, Clone, Copy)]
    struct Corner {
        col: u32,
        col_off: i64,
        row: u32,
        row_off: i64,
    }

    let mut result: Vec<(ImageAnchorGeometry, PictureBlip)> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_anchor = false;
    let mut in_pic = false;
    let mut corner_target: Option<bool> = None; // Some(true)=from, Some(false)=to
    let mut current_field: Option<&'static str> = None;
    let mut from = Corner::default();
    let mut to: Option<Corner> = None;
    let mut ext_emu: Option<(i64, i64)> = None;
    let mut blip_rid: Option<String> = None;
    let mut blip_alpha: Option<f64> = None;
    // `<a:alphaModFix>` counts only as the blip's own child. The `a:extLst`
    // Excel nests inside an effect-bearing blip describes the high-definition
    // layer, whose effects are not the ones applied to the picture — the same
    // rule `blip_embed_rid` follows for `r:embed`.
    let mut in_blip = false;
    let mut ext_lst_depth: u32 = 0;

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
                    blip_alpha = None;
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
                // A picture carrying an alpha or recolour effect spells its
                // blip as a start element with children (issue #1066).
                b"blip" if in_pic => {
                    in_blip = true;
                    if let Some(rid) = blip_embed_rid(e) {
                        blip_rid = Some(rid);
                    }
                }
                b"extLst" if in_blip => ext_lst_depth += 1,
                // `a:alphaModFix` holds nothing but attributes, so it reaches
                // the reader as a start/end pair whenever a writer spells its
                // end tag out.
                b"alphaModFix" if in_blip && ext_lst_depth == 0 => {
                    if let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e) {
                        blip_alpha = Some(alpha);
                    }
                }
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
                            .normalized_value(OOXML_XML_VERSION)
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
                if in_pic
                    && local.as_ref() == b"blip"
                    && let Some(rid) = blip_embed_rid(e)
                {
                    blip_rid = Some(rid);
                }
                // Excel honours the transparency an `<a:alphaModFix>` declares
                // by drawing the picture through a soft mask (issue #1103).
                if in_blip
                    && ext_lst_depth == 0
                    && local.as_ref() == b"alphaModFix"
                    && let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e)
                {
                    blip_alpha = Some(alpha);
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content(OOXML_XML_VERSION)
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
                            PictureBlip {
                                rid,
                                alpha: blip_alpha.take(),
                            },
                        ));
                    }
                    in_anchor = false;
                    in_pic = false;
                    corner_target = None;
                }
                b"blip" => in_blip = false,
                b"extLst" if ext_lst_depth > 0 => ext_lst_depth -= 1,
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
    let (theme_colors, theme_fonts, theme_line_styles) =
        workbook_theme(&mut archive, &workbook_rels_xml);

    let mut result: HashMap<String, Vec<RawTextBoxAnchor>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_dir: String = sheet_part_dir(sheet_target);
        let sheet_xml = read_zip_entry_string(&mut archive, &sheet_part_path(sheet_target));
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path(sheet_target));
        if sheet_rels_xml.is_empty() {
            continue;
        }

        for drawing_target in &sheet_drawing_targets(&sheet_xml, &sheet_rels_xml) {
            let drawing_path = resolve_relative_xl_path(&sheet_dir, drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }
            let boxes = parse_drawing_text_boxes_with_theme_lines(
                &drawing_xml,
                &theme_colors,
                &theme_fonts,
                &theme_line_styles,
            );
            if !boxes.is_empty() {
                result.entry(sheet_name.clone()).or_default().extend(boxes);
            }
        }
    }

    result
}

/// Spreadsheets have no `<clrMap>` part; background/text scheme names map
/// onto the light/dark theme slots directly.
fn xlsx_scheme_aliases() -> HashMap<String, String> {
    [
        ("bg1", "lt1"),
        ("tx1", "dk1"),
        ("bg2", "lt2"),
        ("tx2", "dk2"),
    ]
    .into_iter()
    .map(|(from, to)| (from.to_string(), to.to_string()))
    .collect()
}

/// Historical fallback for workbooks without a readable theme part: light
/// slots render white, dark slots black, everything else is dropped.
fn legacy_scheme_fallback(val: &str) -> Option<crate::ir::Color> {
    use crate::ir::Color;
    match val {
        "lt1" | "bg1" | "lt2" | "bg2" => Some(Color::new(0xFF, 0xFF, 0xFF)),
        "dk1" | "tx1" | "dk2" | "tx2" => Some(Color::new(0, 0, 0)),
        _ => None,
    }
}

/// Resolve a parsed drawing color, falling back to the legacy light/dark
/// mapping for scheme colors the theme could not resolve.
pub(in crate::parser) fn resolved_or_legacy(
    parsed: Option<crate::ir::Color>,
    element_name: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
) -> Option<crate::ir::Color> {
    parsed.or_else(|| {
        if element_name == b"schemeClr" {
            crate::parser::xml_util::get_attr_str(e, b"val")
                .and_then(|val| legacy_scheme_fallback(&val))
        } else {
            None
        }
    })
}

/// Apply an `<a:rPr>` element's size and emphasis attributes to a run
/// style. Called from both the `Start` and `Empty` arms of the drawing
/// reader: Excel writes shape-label run properties as a self-closing
/// element, which quick-xml reports as `Empty`, and reading them on `Start`
/// alone dropped every attribute (issue #466).
pub(in crate::parser) fn apply_run_properties(
    style: &mut crate::ir::TextStyle,
    element: &quick_xml::events::BytesStart<'_>,
) {
    for attr in element.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"sz" => {
                if let Ok(value) = attr.normalized_value(OOXML_XML_VERSION)
                    && let Ok(hundredths_pt) = value.parse::<f64>()
                {
                    style.font_size = Some(hundredths_pt / 100.0);
                }
            }
            b"b" if attr.normalized_value(OOXML_XML_VERSION).ok().as_deref() == Some("1") => {
                style.bold = Some(true);
            }
            b"i" if attr.normalized_value(OOXML_XML_VERSION).ok().as_deref() == Some("1") => {
                style.italic = Some(true);
            }
            _ => {}
        }
    }
}

/// Parse `<xdr:sp>` text boxes from a worksheet drawing, resolving scheme
/// colors against the workbook theme palette and run typefaces against its
/// font scheme.
#[cfg(test)]
pub(super) fn parse_drawing_text_boxes(
    xml: &str,
    theme_colors: &HashMap<String, crate::ir::Color>,
    theme_fonts: &ThemeFontScheme,
) -> Vec<RawTextBoxAnchor> {
    parse_drawing_text_boxes_with_theme_lines(xml, theme_colors, theme_fonts, &[])
}

fn parse_drawing_text_boxes_with_theme_lines(
    xml: &str,
    theme_colors: &HashMap<String, crate::ir::Color>,
    theme_fonts: &ThemeFontScheme,
    theme_line_styles: &[crate::parser::pptx::ThemeLineStyle],
) -> Vec<RawTextBoxAnchor> {
    use crate::ir::{
        Alignment, BorderLineStyle, BorderSide, LineJoin, Paragraph, ParagraphStyle, Run, TextStyle,
    };
    use crate::parser::drawingml::{self, SchemeColors};

    let aliases = xlsx_scheme_aliases();
    let scheme = SchemeColors {
        colors: theme_colors,
        aliases: &aliases,
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
    let mut border_cap: Option<LineCap> = None;
    let mut border_join: Option<LineJoin> = None;
    let mut line_reference_index: Option<usize> = None;
    let mut vertical_center = false;
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
                        border_cap = None;
                        border_join = None;
                        line_reference_index = None;
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
                        border_cap = drawingml::line_cap(e).or(border_cap);
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"w"
                                && let Ok(v) = attr.normalized_value(OOXML_XML_VERSION)
                                && let Ok(w) = v.parse::<f64>()
                            {
                                border_width = w / 12_700.0;
                            }
                        }
                    }
                    b"lnRef" if in_sp && !in_tx_body => {
                        line_reference_index = xml_util::get_attr_str(e, b"idx")
                            .and_then(|value| value.parse::<usize>().ok())
                            .filter(|index| *index > 0);
                    }
                    b"round" | b"bevel" | b"miter" if in_line => {
                        border_join = crate::parser::pptx::drawingml_line_join(local.as_ref())
                            .or(border_join);
                    }
                    b"p" if in_tx_body => {
                        current_para = Some(Paragraph {
                            // A worksheet text box's paragraphs stack with no
                            // gap of their own unless the body asks for one.
                            // This parser reads no `a:spcBef`/`a:spcAft`, and
                            // leaving the fields unset let the renderer's
                            // default block spacing in on top of the line
                            // height, roughly doubling the pitch (issue #656).
                            style: ParagraphStyle {
                                space_before: Some(0.0),
                                space_after: Some(0.0),
                                ..ParagraphStyle::default()
                            },
                            runs: Vec::new(),
                        });
                    }
                    b"r" if current_para.is_some() => {
                        in_run = true;
                        current_style = TextStyle::default();
                    }
                    b"rPr" if in_run => apply_run_properties(&mut current_style, e),
                    b"t" if in_run => in_text = true,
                    b"srgbClr" | b"schemeClr" => {
                        let parsed =
                            drawingml::parse_color_from_start(&mut reader, e, &scheme).color;
                        let color = resolved_or_legacy(parsed, local.as_ref(), e);
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
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bodyPr" if in_tx_body || in_sp => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"anchor"
                                && attr.normalized_value(OOXML_XML_VERSION).ok().as_deref()
                                    == Some("ctr")
                            {
                                vertical_center = true;
                            }
                        }
                    }
                    b"latin" if in_run => {
                        if let Some(typeface) = xml_util::get_attr_str(e, b"typeface")
                            && let Some(family) = theme_fonts.resolve_typeface(&typeface)
                        {
                            current_style.font_family = Some(family);
                        }
                    }
                    b"rPr" if in_run => apply_run_properties(&mut current_style, e),
                    b"ln" if in_sp && !in_tx_body => {
                        border_cap = drawingml::line_cap(e).or(border_cap);
                        if let Some(width) = xml_util::get_attr_str(e, b"w")
                            && let Ok(emu) = width.parse::<f64>()
                        {
                            border_width = emu / 12_700.0;
                        }
                    }
                    b"lnRef" if in_sp && !in_tx_body => {
                        line_reference_index = xml_util::get_attr_str(e, b"idx")
                            .and_then(|value| value.parse::<usize>().ok())
                            .filter(|index| *index > 0);
                    }
                    b"round" | b"bevel" | b"miter" if in_line => {
                        border_join = crate::parser::pptx::drawingml_line_join(local.as_ref())
                            .or(border_join);
                    }
                    b"pPr" if current_para.is_some() => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"algn"
                                && let Ok(v) = attr.normalized_value(OOXML_XML_VERSION)
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
                    b"srgbClr" | b"schemeClr" => {
                        let parsed = drawingml::parse_color_from_empty(e, &scheme).color;
                        let color = resolved_or_legacy(parsed, local.as_ref(), e);
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
                                .normalized_value(OOXML_XML_VERSION)
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
                if in_text && let Ok(text) = t.xml_content(OOXML_XML_VERSION) {
                    if let Some(para) = current_para.as_mut() {
                        let mut style = current_style.clone();
                        // An `<a:rPr>` without `<a:latin>` - how Excel writes
                        // plain shape labels - resolves to the theme's minor
                        // Latin font in DrawingML, not to a renderer default
                        // (issue #461).
                        if style.font_family.is_none() {
                            style.font_family = theme_fonts.minor_latin.clone();
                        }
                        para.runs.push(Run {
                            text: text.to_string(),
                            style,
                            href: None,
                            footnote: None,
                            inline_box: None,
                        });
                    }
                } else if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content(OOXML_XML_VERSION)
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
                                    join: border_join
                                        .or_else(|| {
                                            line_reference_index
                                                .and_then(|index| index.checked_sub(1))
                                                .and_then(|index| theme_line_styles.get(index))
                                                .and_then(|style| style.join)
                                        })
                                        .unwrap_or(LineJoin::Round),
                                    cap: border_cap
                                        .or_else(|| {
                                            line_reference_index
                                                .and_then(|index| index.checked_sub(1))
                                                .and_then(|index| theme_line_styles.get(index))
                                                .and_then(|style| style.cap)
                                        })
                                        .unwrap_or(LineCap::Flat),
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

// ── Drawing line shapes ─────────────────────────────────────────────────

/// A line-geometry shape from a worksheet drawing, in raw drawing
/// coordinates: an `xdr:cxnSp` connector or an `xdr:sp` whose preset is a
/// line. Its extent comes from the anchor once the sheet's printed grid is
/// known, so only the direction is recorded here (issue #1566).
pub(super) struct RawLineAnchor {
    pub(super) geometry: ImageAnchorGeometry,
    pub(super) stroke: crate::ir::BorderSide,
    /// `a:xfrm flipH`: the line runs from the anchor's right edge to its left.
    pub(super) flip_h: bool,
    /// `a:xfrm flipV`: the line runs from the anchor's bottom edge to its top.
    pub(super) flip_v: bool,
}

/// Extract anchored line shapes per sheet from worksheet drawings.
pub(super) fn extract_line_shapes_with_anchors(data: &[u8]) -> HashMap<String, Vec<RawLineAnchor>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);
    let (theme_colors, _theme_fonts, theme_line_styles) =
        workbook_theme(&mut archive, &workbook_rels_xml);

    let mut result: HashMap<String, Vec<RawLineAnchor>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_dir: String = sheet_part_dir(sheet_target);
        let sheet_xml = read_zip_entry_string(&mut archive, &sheet_part_path(sheet_target));
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path(sheet_target));
        if sheet_rels_xml.is_empty() {
            continue;
        }

        for drawing_target in &sheet_drawing_targets(&sheet_xml, &sheet_rels_xml) {
            let drawing_path = resolve_relative_xl_path(&sheet_dir, drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }
            let lines = parse_drawing_line_shapes_with_theme_lines(
                &drawing_xml,
                &theme_colors,
                &theme_line_styles,
            );
            if !lines.is_empty() {
                result.entry(sheet_name.clone()).or_default().extend(lines);
            }
        }
    }

    result
}

/// Whether a DrawingML preset geometry is a straight line.
fn is_line_preset(preset: &str) -> bool {
    matches!(preset, "line" | "straightConnector1")
}

/// Parse the line shapes of one worksheet drawing: every `xdr:cxnSp` or
/// `xdr:sp` whose `a:prstGeom` is a straight line and whose outline paints.
///
/// The outline is the shape's own `a:ln`; when that names no fill, the
/// `xdr:style`'s `a:lnRef` colour stands in, as it does for a slide's
/// connector. `a:noFill` inside `a:ln`, a hidden shape (`cNvPr hidden="1"`),
/// and a shape with no colour at all draw nothing. Excel prints the budget
/// workbook's two `a:ln w="12700"` "Chart border" connectors as 0.78pt
/// #D9D9D9 lines on its fitted page (issue #1566).
#[cfg(test)]
pub(super) fn parse_drawing_line_shapes(
    xml: &str,
    theme_colors: &HashMap<String, crate::ir::Color>,
) -> Vec<RawLineAnchor> {
    parse_drawing_line_shapes_with_theme_lines(xml, theme_colors, &[])
}

fn parse_drawing_line_shapes_with_theme_lines(
    xml: &str,
    theme_colors: &HashMap<String, crate::ir::Color>,
    theme_line_styles: &[crate::parser::pptx::ThemeLineStyle],
) -> Vec<RawLineAnchor> {
    use crate::ir::{BorderLineStyle, BorderSide, LineJoin};
    use crate::parser::drawingml::{self, SchemeColors};

    let aliases = xlsx_scheme_aliases();
    let scheme = SchemeColors {
        colors: theme_colors,
        aliases: &aliases,
    };

    #[derive(Default, Clone, Copy)]
    struct Corner {
        col: u32,
        col_off: i64,
        row: u32,
        row_off: i64,
    }

    /// DrawingML's default outline width when `a:ln` names none.
    const DEFAULT_LINE_WIDTH_PT: f64 = 0.75;

    /// What the current anchor's shape has declared so far.
    struct PendingLine {
        is_line_geometry: bool,
        is_hidden: bool,
        flip_h: bool,
        flip_v: bool,
        width_pt: f64,
        /// The `a:ln` fill colour.
        color: Option<crate::ir::Color>,
        /// The `xdr:style`/`a:lnRef` colour, used when `a:ln` names none.
        style_color: Option<crate::ir::Color>,
        cap: Option<LineCap>,
        join: Option<LineJoin>,
        line_reference_index: Option<usize>,
        has_no_fill: bool,
    }

    impl Default for PendingLine {
        fn default() -> Self {
            Self {
                is_line_geometry: false,
                is_hidden: false,
                flip_h: false,
                flip_v: false,
                width_pt: DEFAULT_LINE_WIDTH_PT,
                color: None,
                style_color: None,
                cap: None,
                join: None,
                line_reference_index: None,
                has_no_fill: false,
            }
        }
    }

    fn is_true(value: Option<String>) -> bool {
        value.is_some_and(|value| value == "1" || value == "true")
    }

    let mut result: Vec<RawLineAnchor> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_anchor = false;
    let mut in_shape = false;
    let mut in_shape_properties = false;
    let mut in_line = false;
    let mut in_style_line_ref = false;
    let mut corner_target: Option<bool> = None;
    let mut current_field: Option<&'static str> = None;
    let mut from = Corner::default();
    let mut to: Option<Corner> = None;
    let mut ext_emu: Option<(i64, i64)> = None;
    let mut pending = PendingLine::default();

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        in_anchor = true;
                        in_shape = false;
                        from = Corner::default();
                        to = None;
                        ext_emu = None;
                        pending = PendingLine::default();
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
                    b"sp" | b"cxnSp" if in_anchor => in_shape = true,
                    b"cNvPr" if in_shape => {
                        pending.is_hidden = is_true(xml_util::get_attr_str(e, b"hidden"));
                    }
                    b"spPr" if in_shape => in_shape_properties = true,
                    b"xfrm" if in_shape_properties => {
                        pending.flip_h = is_true(xml_util::get_attr_str(e, b"flipH"));
                        pending.flip_v = is_true(xml_util::get_attr_str(e, b"flipV"));
                    }
                    b"prstGeom" if in_shape_properties => {
                        pending.is_line_geometry = xml_util::get_attr_str(e, b"prst")
                            .is_some_and(|preset| is_line_preset(&preset));
                    }
                    b"ln" if in_shape_properties => {
                        in_line = true;
                        pending.cap = drawingml::line_cap(e).or(pending.cap);
                        if let Some(width) = xml_util::get_attr_str(e, b"w")
                            && let Ok(emu) = width.parse::<f64>()
                        {
                            pending.width_pt = emu / 12_700.0;
                        }
                    }
                    b"lnRef" if in_shape && !in_shape_properties => {
                        in_style_line_ref = true;
                        pending.line_reference_index = xml_util::get_attr_str(e, b"idx")
                            .and_then(|value| value.parse::<usize>().ok())
                            .filter(|index| *index > 0);
                    }
                    b"round" | b"bevel" | b"miter" if in_line => {
                        pending.join = crate::parser::pptx::drawingml_line_join(local.as_ref())
                            .or(pending.join);
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_line || in_style_line_ref => {
                        let parsed =
                            drawingml::parse_color_from_start(&mut reader, e, &scheme).color;
                        let color = resolved_or_legacy(parsed, local.as_ref(), e);
                        if in_line {
                            if pending.color.is_none() {
                                pending.color = color;
                            }
                        } else if pending.style_color.is_none() {
                            pending.style_color = color;
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"cNvPr" if in_shape => {
                        pending.is_hidden = is_true(xml_util::get_attr_str(e, b"hidden"));
                    }
                    b"xfrm" if in_shape_properties => {
                        pending.flip_h = is_true(xml_util::get_attr_str(e, b"flipH"));
                        pending.flip_v = is_true(xml_util::get_attr_str(e, b"flipV"));
                    }
                    b"prstGeom" if in_shape_properties => {
                        pending.is_line_geometry = xml_util::get_attr_str(e, b"prst")
                            .is_some_and(|preset| is_line_preset(&preset));
                    }
                    b"ln" if in_shape_properties => {
                        // An empty `<a:ln/>` still carries a width; its colour
                        // falls to the style reference.
                        if let Some(width) = xml_util::get_attr_str(e, b"w")
                            && let Ok(emu) = width.parse::<f64>()
                        {
                            pending.width_pt = emu / 12_700.0;
                        }
                        pending.cap = drawingml::line_cap(e).or(pending.cap);
                    }
                    b"lnRef" if in_shape && !in_shape_properties => {
                        pending.line_reference_index = xml_util::get_attr_str(e, b"idx")
                            .and_then(|value| value.parse::<usize>().ok())
                            .filter(|index| *index > 0);
                    }
                    b"round" | b"bevel" | b"miter" if in_line => {
                        pending.join = crate::parser::pptx::drawingml_line_join(local.as_ref())
                            .or(pending.join);
                    }
                    b"noFill" if in_line => pending.has_no_fill = true,
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_line || in_style_line_ref => {
                        let parsed = drawingml::parse_color_from_empty(e, &scheme).color;
                        let color = resolved_or_legacy(parsed, local.as_ref(), e);
                        if in_line {
                            if pending.color.is_none() {
                                pending.color = color;
                            }
                        } else if pending.style_color.is_none() {
                            pending.style_color = color;
                        }
                    }
                    b"ext" if in_anchor && !in_shape && to.is_none() => {
                        let mut cx: i64 = 0;
                        let mut cy: i64 = 0;
                        for attr in e.attributes().flatten() {
                            let value: i64 = attr
                                .normalized_value(OOXML_XML_VERSION)
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
                if let (Some(is_from), Some(field)) = (corner_target, current_field)
                    && let Ok(text) = t.xml_content(OOXML_XML_VERSION)
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
                        let color: Option<crate::ir::Color> = if pending.has_no_fill {
                            None
                        } else {
                            pending.color.or(pending.style_color)
                        };
                        if in_shape
                            && pending.is_line_geometry
                            && !pending.is_hidden
                            && let Some(color) = color
                        {
                            result.push(RawLineAnchor {
                                geometry: ImageAnchorGeometry {
                                    from_row: from.row,
                                    from_col: from.col,
                                    from_col_off_emu: from.col_off,
                                    from_row_off_emu: from.row_off,
                                    to: to.map(|c| (c.col, c.col_off, c.row, c.row_off)),
                                    ext_emu,
                                },
                                stroke: BorderSide {
                                    width: pending.width_pt,
                                    color,
                                    style: BorderLineStyle::Solid,
                                    join: pending
                                        .join
                                        .or_else(|| {
                                            pending
                                                .line_reference_index
                                                .and_then(|index| index.checked_sub(1))
                                                .and_then(|index| theme_line_styles.get(index))
                                                .and_then(|style| style.join)
                                        })
                                        .unwrap_or(LineJoin::Round),
                                    cap: pending
                                        .cap
                                        .or_else(|| {
                                            pending
                                                .line_reference_index
                                                .and_then(|index| index.checked_sub(1))
                                                .and_then(|index| theme_line_styles.get(index))
                                                .and_then(|style| style.cap)
                                        })
                                        .unwrap_or(LineCap::Flat),
                                },
                                flip_h: pending.flip_h,
                                flip_v: pending.flip_v,
                            });
                        }
                        in_anchor = false;
                        in_shape = false;
                        corner_target = None;
                    }
                    b"from" | b"to" => corner_target = None,
                    b"col" | b"colOff" | b"row" | b"rowOff" => current_field = None,
                    b"spPr" => in_shape_properties = false,
                    b"ln" => in_line = false,
                    b"lnRef" => in_style_line_ref = false,
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

/// Read the workbook theme's color palette and Latin font scheme
/// (`xl/theme/theme1.xml`, or the rels-declared target). A missing or
/// unreadable theme yields an empty palette, which downgrades scheme
/// resolution to the legacy light/dark fallback, and an empty font scheme,
/// which leaves font selection to the renderer.
fn workbook_theme(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    workbook_rels_xml: &str,
) -> (
    HashMap<String, crate::ir::Color>,
    ThemeFontScheme,
    Vec<crate::parser::pptx::ThemeLineStyle>,
) {
    let theme_path = parse_rels_by_type(workbook_rels_xml, "theme")
        .first()
        .map(|target| resolve_relative_xl_path("xl", target))
        .unwrap_or_else(|| "xl/theme/theme1.xml".to_string());
    let theme_xml = read_zip_entry_string(archive, &theme_path);
    let line_styles: Vec<crate::parser::pptx::ThemeLineStyle> =
        crate::parser::pptx::parse_theme_line_styles(&theme_xml);
    (
        crate::parser::drawingml::parse_theme_color_scheme(&theme_xml),
        crate::parser::drawingml::parse_theme_font_scheme(&theme_xml),
        line_styles,
    )
}

#[cfg(test)]
#[path = "xlsx_drawing_tests.rs"]
mod tests;
