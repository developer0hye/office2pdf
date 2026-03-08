use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, Chart, Color, Document, HFInline,
    HeaderFooter, HeaderFooterParagraph, Margins, Metadata, Page, PageSize, Paragraph,
    ParagraphStyle, Run, StyleSheet, Table, TableCell, TablePage, TableRow, TextStyle,
};
use crate::parser::Parser;
use crate::parser::chart::parse_chart_xml;
use crate::parser::cond_fmt::build_cond_fmt_overrides;

/// Parser for XLSX (Office Open XML Excel) spreadsheets.
pub struct XlsxParser;

/// Default column width in Excel character units.
const DEFAULT_COLUMN_WIDTH: f64 = 8.43;

/// Convert Excel column width (character units) to points.
/// Excel character width ≈ 7 pixels at 96 DPI, 1 point = 96/72 pixels.
/// Empirically: width_pt ≈ char_width * 7.0 (approximate, close to Excel's rendering).
fn column_width_to_pt(char_width: f64) -> f64 {
    char_width * 7.0
}

/// Parse an ARGB hex string (e.g. "FFFF0000") into an IR Color.
/// Returns None if the string is too short or invalid.
fn parse_argb_color(argb: &str) -> Option<Color> {
    if argb.len() < 8 {
        return None;
    }
    let r = u8::from_str_radix(&argb[2..4], 16).ok()?;
    let g = u8::from_str_radix(&argb[4..6], 16).ok()?;
    let b = u8::from_str_radix(&argb[6..8], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Map Excel border style name to width in points.
fn border_style_to_width(style: &str) -> Option<f64> {
    match style {
        "hair" => Some(0.25),
        "thin" | "dashed" | "dotted" | "dashDot" | "dashDotDot" => Some(0.5),
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "double"
        | "slantDashDot" => Some(1.0),
        "thick" => Some(2.0),
        _ => None, // "none" or unknown
    }
}

/// Extract font styling from a cell's style into an IR TextStyle.
fn extract_cell_text_style(cell: &umya_spreadsheet::Cell) -> TextStyle {
    let style = cell.get_style();
    let Some(font) = style.get_font() else {
        return TextStyle::default();
    };

    let bold = if *font.get_bold() { Some(true) } else { None };
    let italic = if *font.get_italic() { Some(true) } else { None };
    let underline = match font.get_underline() {
        "none" | "" => None,
        _ => Some(true),
    };
    let strikethrough = if *font.get_strikethrough() {
        Some(true)
    } else {
        None
    };

    // Font name: skip default "Calibri" (Excel default) — only set if explicitly customized
    let font_name = font.get_name();
    let font_family = if font_name.is_empty() || font_name == "Calibri" {
        None
    } else {
        Some(font_name.to_string())
    };

    // Font size: skip default 11.0 (Excel default)
    let raw_size = *font.get_size();
    let font_size = if (raw_size - 11.0).abs() < 0.01 {
        None
    } else {
        Some(raw_size)
    };

    // Font color
    let color_argb = font.get_color().get_argb();
    let color = if color_argb.is_empty() || color_argb == "FF000000" {
        // Default black — skip
        None
    } else {
        parse_argb_color(color_argb)
    };

    TextStyle {
        font_family,
        font_size,
        bold,
        italic,
        underline,
        strikethrough,
        color,
        highlight: None,
        vertical_align: None,
        all_caps: None,
        small_caps: None,
        letter_spacing: None,
    }
}

/// Extract background color from a cell's style.
fn extract_cell_background(cell: &umya_spreadsheet::Cell) -> Option<Color> {
    let bg = cell.get_style().get_background_color()?;
    parse_argb_color(bg.get_argb())
}

/// Map Excel border style name to `BorderLineStyle`.
fn border_style_to_line_style(style: &str) -> BorderLineStyle {
    match style {
        "dashed" | "mediumDashed" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "dashDot" | "mediumDashDot" | "slantDashDot" => BorderLineStyle::DashDot,
        "dashDotDot" | "mediumDashDotDot" => BorderLineStyle::DashDotDot,
        "double" => BorderLineStyle::Double,
        _ => BorderLineStyle::Solid,
    }
}

/// Extract a single border side from an umya Border object.
fn extract_border_side(border: &umya_spreadsheet::Border) -> Option<BorderSide> {
    let border_style_str = border.get_border_style();
    let width = border_style_to_width(border_style_str)?;
    let color = parse_argb_color(border.get_color().get_argb()).unwrap_or(Color::black());
    let style = border_style_to_line_style(border_style_str);
    Some(BorderSide {
        width,
        color,
        style,
    })
}

/// Extract cell border properties.
fn extract_cell_borders(cell: &umya_spreadsheet::Cell) -> Option<CellBorder> {
    let borders = cell.get_style().get_borders()?;
    let top = extract_border_side(borders.get_top());
    let bottom = extract_border_side(borders.get_bottom());
    let left = extract_border_side(borders.get_left());
    let right = extract_border_side(borders.get_right());
    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }
    Some(CellBorder {
        top,
        bottom,
        left,
        right,
    })
}

/// A cell range within a sheet (1-indexed, inclusive).
#[derive(Debug, Clone, Copy)]
struct CellRange {
    start_col: u32,
    start_row: u32,
    end_col: u32,
    end_row: u32,
}

/// Parse an Excel column letter string (e.g., "A", "B", "AA") into a 1-indexed column number.
fn parse_column_letters(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }
    let mut col: u32 = 0;
    for c in s.chars() {
        if !c.is_ascii_uppercase() {
            return None;
        }
        col = col * 26 + (c as u32 - b'A' as u32 + 1);
    }
    Some(col)
}

/// Parse a cell reference like "$A$1", "A1", "$B$10" into (col, row), both 1-indexed.
fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    // Strip dollar signs
    let s = s.replace('$', "");
    // Split into letter part and number part
    let split_pos = s.find(|c: char| c.is_ascii_digit())?;
    let col_str = &s[..split_pos];
    let row_str = &s[split_pos..];
    let col = parse_column_letters(col_str)?;
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
}

/// Parse a print area address string (e.g., "Sheet1!$A$1:$C$10") into a CellRange.
fn parse_print_area_range(address: &str) -> Option<CellRange> {
    // Strip optional sheet prefix (everything up to and including '!')
    let range_part = if let Some(pos) = address.rfind('!') {
        &address[pos + 1..]
    } else {
        address
    };

    let (start_str, end_str) = range_part.split_once(':')?;
    let (start_col, start_row) = parse_cell_ref(start_str)?;
    let (end_col, end_row) = parse_cell_ref(end_str)?;
    Some(CellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    })
}

/// Look up the print area for a given sheet from its defined names.
fn find_print_area(sheet: &umya_spreadsheet::Worksheet) -> Option<CellRange> {
    for dn in sheet.get_defined_names() {
        if dn.get_name() == "_xlnm.Print_Area" {
            let addr = dn.get_address();
            if let Some(range) = parse_print_area_range(&addr) {
                return Some(range);
            }
        }
    }
    None
}

/// Collect sorted manual row page break positions from a sheet.
fn collect_row_breaks(sheet: &umya_spreadsheet::Worksheet) -> Vec<u32> {
    let mut breaks: Vec<u32> = sheet
        .get_row_breaks()
        .get_break_list()
        .iter()
        .filter(|b| *b.get_manual_page_break())
        .map(|b| *b.get_id())
        .collect();
    breaks.sort_unstable();
    breaks.dedup();
    breaks
}

/// Parse an Excel header/footer format string into IR HeaderFooter.
///
/// Excel format strings use `&L`, `&C`, `&R` to define left/center/right sections,
/// `&P` for current page number, and `&N` for total page count.
/// Returns `None` if the format string is empty.
fn parse_hf_format_string(format_str: &str) -> Option<HeaderFooter> {
    let s = format_str.trim();
    if s.is_empty() {
        return None;
    }

    // Split into left/center/right sections
    let mut left = String::new();
    let mut center = String::new();
    let mut right = String::new();
    let mut current = &mut center; // Default section is center if no &L/&C/&R prefix

    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '&' && i + 1 < chars.len() {
            match chars[i + 1] {
                'L' => {
                    current = &mut left;
                    i += 2;
                }
                'C' => {
                    current = &mut center;
                    i += 2;
                }
                'R' => {
                    current = &mut right;
                    i += 2;
                }
                'P' => {
                    current.push('\x01'); // Sentinel for page number
                    i += 2;
                }
                'N' => {
                    current.push('\x02'); // Sentinel for total pages
                    i += 2;
                }
                '&' => {
                    // Escaped ampersand: && → &
                    current.push('&');
                    i += 2;
                }
                '"' => {
                    // Font name: &"FontName" — skip to closing quote
                    i += 2; // skip &"
                    while i < chars.len() && chars[i] != '"' {
                        i += 1;
                    }
                    if i < chars.len() {
                        i += 1; // skip closing "
                    }
                }
                c if c.is_ascii_digit() => {
                    // Font size: &NN — skip digits
                    i += 1; // skip &
                    while i < chars.len() && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                }
                _ => {
                    // Unknown code — skip it
                    i += 2;
                }
            }
        } else {
            current.push(chars[i]);
            i += 1;
        }
    }

    let mut paragraphs = Vec::new();

    // Build paragraph for each non-empty section
    let sections = [
        (&left, Alignment::Left),
        (&center, Alignment::Center),
        (&right, Alignment::Right),
    ];

    for (text, alignment) in &sections {
        if text.is_empty() {
            continue;
        }
        let elements = build_hf_elements(text);
        if !elements.is_empty() {
            paragraphs.push(HeaderFooterParagraph {
                style: ParagraphStyle {
                    alignment: Some(*alignment),
                    ..ParagraphStyle::default()
                },
                elements,
            });
        }
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter { paragraphs })
    }
}

/// Build HFInline elements from a section string, replacing sentinel chars.
fn build_hf_elements(section: &str) -> Vec<HFInline> {
    let mut elements = Vec::new();
    let mut current_text = String::new();

    for ch in section.chars() {
        match ch {
            '\x01' => {
                // Page number sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::PageNumber);
            }
            '\x02' => {
                // Total pages sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::TotalPages);
            }
            _ => {
                current_text.push(ch);
            }
        }
    }

    if !current_text.is_empty() {
        elements.push(HFInline::Run(Run {
            text: current_text,
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }));
    }

    elements
}

/// A (column, row) coordinate pair (1-indexed).
type CellPos = (u32, u32);

/// Info about a merged cell region, keyed by its top-left coordinate.
struct MergeInfo {
    col_span: u32,
    row_span: u32,
}

/// Build a lookup of merge info from the sheet's merged cell ranges.
///
/// Returns two structures:
/// - `top_left_map`: top-left coordinate → MergeInfo for each merge
/// - `skip_set`: set of coordinates that are inside a merge but NOT the top-left
fn build_merge_maps(
    sheet: &umya_spreadsheet::Worksheet,
) -> (HashMap<CellPos, MergeInfo>, HashSet<CellPos>) {
    let mut top_left_map: HashMap<CellPos, MergeInfo> = HashMap::new();
    let mut skip_set: HashSet<CellPos> = HashSet::new();

    for range in sheet.get_merge_cells() {
        let start_col = range
            .get_coordinate_start_col()
            .map(|c| *c.get_num())
            .unwrap_or(1);
        let start_row = range
            .get_coordinate_start_row()
            .map(|r| *r.get_num())
            .unwrap_or(1);
        let end_col = range
            .get_coordinate_end_col()
            .map(|c| *c.get_num())
            .unwrap_or(start_col);
        let end_row = range
            .get_coordinate_end_row()
            .map(|r| *r.get_num())
            .unwrap_or(start_row);

        let col_span = end_col.saturating_sub(start_col) + 1;
        let row_span = end_row.saturating_sub(start_row) + 1;

        top_left_map.insert((start_col, start_row), MergeInfo { col_span, row_span });

        // Mark all other cells in the range as skip
        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if r != start_row || c != start_col {
                    skip_set.insert((c, r));
                }
            }
        }
    }

    (top_left_map, skip_set)
}

/// Extract charts from the XLSX ZIP with their anchor positions per sheet.
///
/// Returns a map from sheet name → list of (anchor_row, Chart).
/// Charts with drawing anchors get positioned at their anchor row.
/// Charts without anchors (no drawing reference found) use `u32::MAX`
/// as a sentinel to place them at the end of the sheet.
fn extract_charts_with_anchors(data: &[u8]) -> HashMap<String, Vec<(u32, Chart)>> {
    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(data)) else {
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
fn collect_positioned_chart_paths(
    chart_map: &HashMap<String, Vec<(u32, Chart)>>,
    data: &[u8],
) -> HashSet<String> {
    // Re-trace the drawing → chart resolution to find which chart paths are covered.
    // This is intentionally conservative — if we can't determine the path, we skip.
    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(data)) else {
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
fn read_zip_entry_string(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, path: &str) -> String {
    let Ok(mut entry) = archive.by_name(path) else {
        return String::new();
    };
    let mut xml = String::new();
    let _ = std::io::Read::read_to_string(&mut entry, &mut xml);
    xml
}

/// Parse workbook.xml to extract sheet name → rId pairs (preserving order).
fn parse_workbook_sheet_rids(xml: &str) -> Vec<(String, String)> {
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
fn parse_rels_targets(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Id" => {
                                if let Ok(v) = attr.unescape_value() {
                                    id = Some(v.to_string());
                                }
                            }
                            b"Target" => {
                                if let Ok(v) = attr.unescape_value() {
                                    target = Some(v.to_string());
                                }
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Parse a .rels file and return targets whose Type contains the given substring.
fn parse_rels_by_type(xml: &str, type_substring: &str) -> Vec<String> {
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
fn resolve_relative_xl_path(base_dir: &str, relative: &str) -> String {
    if relative.starts_with('/') {
        return relative.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for segment in relative.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            _ => parts.push(segment),
        }
    }
    parts.join("/")
}

/// Parse drawing XML for chart anchor positions.
/// Returns (anchor_row, chart_rId) pairs from `<xdr:twoCellAnchor>` elements.
fn parse_drawing_chart_anchors(xml: &str) -> Vec<(u32, String)> {
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

/// Shared context for processing a single XLSX sheet.
struct SheetContext {
    col_start: u32,
    col_end: u32,
    num_cols: usize,
    column_widths: Vec<f64>,
    merge_tops: HashMap<(u32, u32), MergeInfo>,
    merge_skips: HashSet<(u32, u32)>,
    cond_fmt_overrides: HashMap<(u32, u32), crate::parser::cond_fmt::CondFmtOverride>,
}

/// Build TableRows for a range of rows in a sheet.
fn build_rows_for_range(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    row_start: u32,
    row_end: u32,
) -> Vec<TableRow> {
    let num_rows = (row_end - row_start + 1) as usize;
    let mut rows = Vec::with_capacity(num_rows);
    for row_idx in row_start..=row_end {
        let mut cells = Vec::with_capacity(ctx.num_cols);
        for col_idx in ctx.col_start..=ctx.col_end {
            // Skip cells that are part of a merge but not the top-left
            if ctx.merge_skips.contains(&(col_idx, row_idx)) {
                continue;
            }

            // umya-spreadsheet tuple is (column, row), both 1-indexed
            let umya_cell = sheet.get_cell((col_idx, row_idx));
            let value = umya_cell
                .map(|cell| cell.get_formatted_value())
                .unwrap_or_default();

            // Extract formatting from the cell
            let mut text_style = umya_cell.map(extract_cell_text_style).unwrap_or_default();
            let mut background = umya_cell.and_then(extract_cell_background);
            let border = umya_cell.and_then(extract_cell_borders);

            // Apply conditional formatting overrides
            let mut data_bar = None;
            let mut icon_text = None;
            if let Some(ovr) = ctx.cond_fmt_overrides.get(&(col_idx, row_idx)) {
                if ovr.background.is_some() {
                    background = ovr.background;
                }
                if ovr.font_color.is_some() {
                    text_style.color = ovr.font_color;
                }
                if let Some(bold) = ovr.bold {
                    text_style.bold = Some(bold);
                }
                data_bar = ovr.data_bar.clone();
                icon_text = ovr.icon_text.clone();
            }

            let content = if value.is_empty() {
                Vec::new()
            } else {
                vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: value,
                        style: text_style,
                        href: None,
                        footnote: None,
                    }],
                })]
            };

            let (col_span, row_span) = if let Some(info) = ctx.merge_tops.get(&(col_idx, row_idx)) {
                (info.col_span, info.row_span)
            } else {
                (1, 1)
            };

            cells.push(TableCell {
                content,
                col_span,
                row_span,
                border,
                background,
                data_bar,
                icon_text,
                vertical_align: None,
                padding: None,
            });
        }

        // Extract row height if custom
        let height = sheet
            .get_row_dimension(&row_idx)
            .filter(|r| *r.get_custom_height())
            .map(|r| *r.get_height());

        rows.push(TableRow { cells, height });
    }
    rows
}

/// Prepare the shared context for processing a sheet (dimensions, merges, styles, etc.).
/// Returns (SheetContext, col_start, col_end, row_start, row_end) or None if the sheet is empty.
fn prepare_sheet_context(sheet: &umya_spreadsheet::Worksheet) -> Option<(SheetContext, u32, u32)> {
    let (mut max_col, mut max_row) = sheet.get_highest_column_and_row();
    if max_col == 0 || max_row == 0 {
        return None;
    }

    // Expand grid to include the extent of all merged ranges
    for range in sheet.get_merge_cells() {
        if let Some(c) = range.get_coordinate_end_col() {
            max_col = max_col.max(*c.get_num());
        }
        if let Some(r) = range.get_coordinate_end_row() {
            max_row = max_row.max(*r.get_num());
        }
    }

    // Check for print area — limit to that range if defined
    let print_area = find_print_area(sheet);
    let (col_start, col_end, row_start, row_end) = if let Some(pa) = print_area {
        (pa.start_col, pa.end_col, pa.start_row, pa.end_row)
    } else {
        (1, max_col, 1, max_row)
    };

    let column_widths: Vec<f64> = (col_start..=col_end)
        .map(|col| {
            sheet
                .get_column_dimension_by_number(&col)
                .map(|c| column_width_to_pt(*c.get_width()))
                .unwrap_or_else(|| column_width_to_pt(DEFAULT_COLUMN_WIDTH))
        })
        .collect();

    let (merge_tops, merge_skips) = build_merge_maps(sheet);
    let cond_fmt_overrides = build_cond_fmt_overrides(sheet);
    let num_cols = (col_end - col_start + 1) as usize;

    Some((
        SheetContext {
            col_start,
            col_end,
            num_cols,
            column_widths,
            merge_tops,
            merge_skips,
            cond_fmt_overrides,
        },
        row_start,
        row_end,
    ))
}

impl XlsxParser {
    /// Parse XLSX in streaming mode, returning one `Document` per chunk of rows.
    ///
    /// Each chunk contains a single `TablePage` with at most `chunk_size` rows.
    /// This allows the caller to compile each chunk independently, bounding peak
    /// memory during Typst compilation.
    pub fn parse_streaming(
        &self,
        data: &[u8],
        options: &ConvertOptions,
        chunk_size: usize,
    ) -> Result<(Vec<Document>, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).map_err(|e| {
            ConvertError::Parse(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        let metadata = extract_xlsx_metadata(&book);
        let mut chart_map = extract_charts_with_anchors(data);

        let mut chunks = Vec::new();
        let mut warnings = Vec::new();

        for sheet in book.get_sheet_collection() {
            // Filter by sheet name if specified
            if let Some(ref names) = options.sheet_names
                && !names.iter().any(|n| n == sheet.get_name())
            {
                continue;
            }

            let Some((ctx, row_start, row_end)) = prepare_sheet_context(sheet) else {
                continue;
            };

            let sheet_name = sheet.get_name().to_string();

            // Extract sheet header/footer
            let hf = sheet.get_header_footer();
            let sheet_header = parse_hf_format_string(hf.get_odd_header().get_value());
            let sheet_footer = parse_hf_format_string(hf.get_odd_footer().get_value());

            // Pull charts for this sheet
            let mut sheet_charts = chart_map.remove(&sheet_name).unwrap_or_default();
            for (_, chart) in &sheet_charts {
                let title = chart.title.as_deref().unwrap_or("untitled").to_string();
                warnings.push(ConvertWarning::FallbackUsed {
                    format: "XLSX".to_string(),
                    from: format!("chart ({title})"),
                    to: "data table".to_string(),
                });
            }
            sheet_charts.sort_by_key(|(row, _)| *row);

            // Process rows in chunks
            let mut chunk_start = row_start;
            let mut first_chunk = true;
            while chunk_start <= row_end {
                let chunk_end = (chunk_start + chunk_size as u32 - 1).min(row_end);

                let rows = build_rows_for_range(sheet, &ctx, chunk_start, chunk_end);

                let doc = Document {
                    metadata: metadata.clone(),
                    pages: vec![Page::Table(TablePage {
                        name: sheet_name.clone(),
                        size: PageSize::default(),
                        margins: Margins::default(),
                        table: Table {
                            rows,
                            column_widths: ctx.column_widths.clone(),
                            header_row_count: 0,
                            alignment: None,
                            default_cell_padding: None,
                            use_content_driven_row_heights: false,
                        },
                        header: sheet_header.clone(),
                        footer: sheet_footer.clone(),
                        charts: if first_chunk {
                            first_chunk = false;
                            std::mem::take(&mut sheet_charts)
                        } else {
                            vec![]
                        },
                    })],
                    styles: StyleSheet::default(),
                };

                chunks.push(doc);
                chunk_start = chunk_end + 1;
            }
        }

        Ok((chunks, warnings))
    }
}

impl Parser for XlsxParser {
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).map_err(|e| {
            ConvertError::Parse(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        // Extract metadata from umya-spreadsheet properties
        let metadata = extract_xlsx_metadata(&book);

        // Extract charts with anchor positions per sheet
        let mut chart_map = extract_charts_with_anchors(data);

        let sheet_count = book.get_sheet_collection().len();
        let mut pages = Vec::with_capacity(sheet_count);
        let mut warnings = Vec::new();

        for sheet in book.get_sheet_collection() {
            // Filter by sheet name if specified
            if let Some(ref names) = options.sheet_names
                && !names.iter().any(|n| n == sheet.get_name())
            {
                continue;
            }

            let Some((ctx, row_start, row_end)) = prepare_sheet_context(sheet) else {
                continue;
            };

            let rows = build_rows_for_range(sheet, &ctx, row_start, row_end);

            // Collect row page breaks and split rows into page segments
            let row_breaks = collect_row_breaks(sheet);
            let sheet_name = sheet.get_name().to_string();

            // Extract sheet header/footer
            let hf = sheet.get_header_footer();
            let sheet_header = parse_hf_format_string(hf.get_odd_header().get_value());
            let sheet_footer = parse_hf_format_string(hf.get_odd_footer().get_value());

            // Pull charts for this sheet (if any)
            let mut sheet_charts = chart_map.remove(&sheet_name).unwrap_or_default();
            for (_, chart) in &sheet_charts {
                let title = chart.title.as_deref().unwrap_or("untitled").to_string();
                warnings.push(ConvertWarning::FallbackUsed {
                    format: "XLSX".to_string(),
                    from: format!("chart ({title})"),
                    to: "data table".to_string(),
                });
            }
            // Sort by anchor row
            sheet_charts.sort_by_key(|(row, _)| *row);

            if row_breaks.is_empty() {
                // No page breaks — single page
                pages.push(Page::Table(TablePage {
                    name: sheet_name,
                    size: PageSize::default(),
                    margins: Margins::default(),
                    table: Table {
                        rows,
                        column_widths: ctx.column_widths,
                        header_row_count: 0,
                        alignment: None,
                        default_cell_padding: None,
                        use_content_driven_row_heights: false,
                    },
                    header: sheet_header.clone(),
                    footer: sheet_footer.clone(),
                    charts: sheet_charts,
                }));
            } else {
                // Split rows at break points
                // Breaks are 1-indexed row numbers; break after that row
                let mut segments: Vec<Vec<TableRow>> = Vec::new();
                let mut current_segment: Vec<TableRow> = Vec::new();
                let mut break_idx = 0;

                for (i, row) in rows.into_iter().enumerate() {
                    let actual_row = row_start + i as u32; // 1-indexed row number
                    current_segment.push(row);

                    // Check if this row is a break point
                    if break_idx < row_breaks.len() && actual_row == row_breaks[break_idx] {
                        segments.push(std::mem::take(&mut current_segment));
                        break_idx += 1;
                    }
                }
                // Push remaining rows as the last segment
                if !current_segment.is_empty() {
                    segments.push(current_segment);
                }

                // For page-break segments, attach all charts to the first segment
                let mut first_segment = true;
                for segment in segments {
                    pages.push(Page::Table(TablePage {
                        name: sheet_name.clone(),
                        size: PageSize::default(),
                        margins: Margins::default(),
                        table: Table {
                            rows: segment,
                            column_widths: ctx.column_widths.clone(),
                            header_row_count: 0,
                            alignment: None,
                            default_cell_padding: None,
                            use_content_driven_row_heights: false,
                        },
                        header: sheet_header.clone(),
                        footer: sheet_footer.clone(),
                        charts: if first_segment {
                            first_segment = false;
                            std::mem::take(&mut sheet_charts)
                        } else {
                            vec![]
                        },
                    }));
                }
            }
        }

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

/// Extract metadata from umya-spreadsheet Properties.
/// Empty strings are converted to None.
fn extract_xlsx_metadata(book: &umya_spreadsheet::Spreadsheet) -> Metadata {
    let props = book.get_properties();
    let non_empty = |s: &str| {
        if s.is_empty() {
            None
        } else {
            Some(s.to_string())
        }
    };
    Metadata {
        title: non_empty(props.get_title()),
        author: non_empty(props.get_creator()),
        subject: non_empty(props.get_subject()),
        description: non_empty(props.get_description()),
        created: non_empty(props.get_created()),
        modified: non_empty(props.get_modified()),
    }
}

#[cfg(test)]
#[path = "xlsx_tests.rs"]
mod tests;
