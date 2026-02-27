use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, Block, BorderSide, CellBorder, Color, Document, HFInline, HeaderFooter,
    HeaderFooterParagraph, Margins, Metadata, Page, PageSize, Paragraph, ParagraphStyle, Run,
    StyleSheet, Table, TableCell, TablePage, TableRow, TextStyle,
};
use crate::parser::Parser;

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
    }
}

/// Extract background color from a cell's style.
fn extract_cell_background(cell: &umya_spreadsheet::Cell) -> Option<Color> {
    let bg = cell.get_style().get_background_color()?;
    parse_argb_color(bg.get_argb())
}

/// Extract a single border side from an umya Border object.
fn extract_border_side(border: &umya_spreadsheet::Border) -> Option<BorderSide> {
    let width = border_style_to_width(border.get_border_style())?;
    let color = parse_argb_color(border.get_color().get_argb()).unwrap_or(Color::black());
    Some(BorderSide { width, color })
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

impl Parser for XlsxParser {
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse XLSX: {e}")))?;

        let mut pages = Vec::new();
        let warnings = Vec::new();

        for sheet in book.get_sheet_collection() {
            // Filter by sheet name if specified
            if let Some(ref names) = options.sheet_names
                && !names.iter().any(|n| n == sheet.get_name())
            {
                continue;
            }

            let (mut max_col, mut max_row) = sheet.get_highest_column_and_row();
            if max_col == 0 || max_row == 0 {
                continue; // skip empty sheets
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

            let mut rows = Vec::new();
            for row_idx in row_start..=row_end {
                let mut cells = Vec::new();
                for col_idx in col_start..=col_end {
                    // Skip cells that are part of a merge but not the top-left
                    if merge_skips.contains(&(col_idx, row_idx)) {
                        continue;
                    }

                    // umya-spreadsheet tuple is (column, row), both 1-indexed
                    let umya_cell = sheet.get_cell((col_idx, row_idx));
                    let value = umya_cell
                        .map(|cell| cell.get_formatted_value())
                        .unwrap_or_default();

                    // Extract formatting from the cell
                    let text_style = umya_cell.map(extract_cell_text_style).unwrap_or_default();
                    let background = umya_cell.and_then(extract_cell_background);
                    let border = umya_cell.and_then(extract_cell_borders);

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

                    let (col_span, row_span) =
                        if let Some(info) = merge_tops.get(&(col_idx, row_idx)) {
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
                    });
                }

                // Extract row height if custom
                let height = sheet
                    .get_row_dimension(&row_idx)
                    .filter(|r| *r.get_custom_height())
                    .map(|r| *r.get_height());

                rows.push(TableRow { cells, height });
            }

            // Collect row page breaks and split rows into page segments
            let row_breaks = collect_row_breaks(sheet);
            let sheet_name = sheet.get_name().to_string();

            // Extract sheet header/footer
            let hf = sheet.get_header_footer();
            let sheet_header = parse_hf_format_string(hf.get_odd_header().get_value());
            let sheet_footer = parse_hf_format_string(hf.get_odd_footer().get_value());

            if row_breaks.is_empty() {
                // No page breaks — single page
                pages.push(Page::Table(TablePage {
                    name: sheet_name,
                    size: PageSize::default(),
                    margins: Margins::default(),
                    table: Table {
                        rows,
                        column_widths,
                    },
                    header: sheet_header.clone(),
                    footer: sheet_footer.clone(),
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

                for segment in segments {
                    pages.push(Page::Table(TablePage {
                        name: sheet_name.clone(),
                        size: PageSize::default(),
                        margins: Margins::default(),
                        table: Table {
                            rows: segment,
                            column_widths: column_widths.clone(),
                        },
                        header: sheet_header.clone(),
                        footer: sheet_footer.clone(),
                    }));
                }
            }
        }

        Ok((
            Document {
                metadata: Metadata::default(),
                pages,
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;

    /// Helper: build a minimal XLSX as bytes with a single sheet.
    fn build_xlsx_bytes(sheet_name: &str, cells: &[(&str, &str)]) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name(sheet_name);
            for &(coord, value) in cells {
                sheet.get_cell_mut(coord).set_value(value);
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build XLSX with multiple sheets.
    fn build_xlsx_multi_sheet(sheets: &[(&str, &[(&str, &str)])]) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        // Remove the default sheet first
        for (i, &(name, cells)) in sheets.iter().enumerate() {
            if i == 0 {
                let sheet = book.get_sheet_mut(&0).unwrap();
                sheet.set_name(name);
                for &(coord, value) in cells {
                    sheet.get_cell_mut(coord).set_value(value);
                }
            } else {
                let mut sheet = umya_spreadsheet::Worksheet::default();
                sheet.set_name(name);
                for &(coord, value) in cells {
                    sheet.get_cell_mut(coord).set_value(value);
                }
                book.add_sheet(sheet).unwrap();
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: extract TablePage from Document by index.
    fn get_table_page(doc: &Document, idx: usize) -> &TablePage {
        match &doc.pages[idx] {
            Page::Table(tp) => tp,
            _ => panic!("Expected TablePage at index {idx}"),
        }
    }

    /// Helper: get cell text from a TableCell.
    fn cell_text(cell: &TableCell) -> String {
        cell.content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect::<Vec<_>>()
            .join("")
    }

    // ----- Basic parsing tests -----

    #[test]
    fn test_parse_single_cell() {
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Hello")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.name, "Sheet1");
        assert_eq!(tp.table.rows.len(), 1);
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Hello");
    }

    #[test]
    fn test_parse_multiple_cells() {
        let data = build_xlsx_bytes(
            "Data",
            &[("A1", "Name"), ("B1", "Age"), ("A2", "Alice"), ("B2", "30")],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 2);
        assert_eq!(tp.table.rows[0].cells.len(), 2);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Name");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "Age");
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Alice");
        assert_eq!(cell_text(&tp.table.rows[1].cells[1]), "30");
    }

    #[test]
    fn test_parse_empty_cells_in_grid() {
        // A1 filled, B1 empty, A2 empty, B2 filled → 2x2 grid with gaps
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Top-Left"), ("B2", "Bottom-Right")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 2);
        assert_eq!(tp.table.rows[0].cells.len(), 2);
        // A1 has content
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Top-Left");
        // B1 is empty
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "");
        // A2 is empty
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "");
        // B2 has content
        assert_eq!(cell_text(&tp.table.rows[1].cells[1]), "Bottom-Right");
    }

    #[test]
    fn test_parse_numbers() {
        let data = build_xlsx_bytes("Numbers", &[("A1", "42"), ("B1", "3.14")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "42");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "3.14");
    }

    #[test]
    fn test_parse_dates_as_text() {
        let data = build_xlsx_bytes("Dates", &[("A1", "2024-01-15"), ("A2", "December 25")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "2024-01-15");
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "December 25");
    }

    // ----- Sheet name tests -----

    #[test]
    fn test_sheet_name_preserved() {
        let data = build_xlsx_bytes("Financial Report", &[("A1", "Revenue")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.name, "Financial Report");
    }

    // ----- Multi-sheet tests -----

    #[test]
    fn test_parse_multiple_sheets() {
        let data = build_xlsx_multi_sheet(&[
            ("Sheet1", &[("A1", "Data1")]),
            ("Sheet2", &[("A1", "Data2")]),
        ]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 2);
        let tp1 = get_table_page(&doc, 0);
        let tp2 = get_table_page(&doc, 1);
        assert_eq!(tp1.name, "Sheet1");
        assert_eq!(tp2.name, "Sheet2");
        assert_eq!(cell_text(&tp1.table.rows[0].cells[0]), "Data1");
        assert_eq!(cell_text(&tp2.table.rows[0].cells[0]), "Data2");
    }

    // ----- Column width tests -----

    #[test]
    fn test_column_widths_default() {
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Hello"), ("B1", "World")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.column_widths.len(), 2);
        // Default column width is ~8.43 chars * 7.0 ≈ 59 pt
        // umya-spreadsheet may use a slightly different default; allow 1pt tolerance
        for w in &tp.table.column_widths {
            assert!(
                *w > 50.0 && *w < 70.0,
                "Expected default width in 50-70pt range, got {w}"
            );
        }
    }

    // ----- Page size and margins defaults -----

    #[test]
    fn test_page_size_defaults() {
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Test")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let default_size = PageSize::default();
        assert!((tp.size.width - default_size.width).abs() < 0.01);
        assert!((tp.size.height - default_size.height).abs() < 0.01);
    }

    // ----- Table structure tests -----

    #[test]
    fn test_table_row_column_consistency() {
        // 3x3 grid, only some cells filled
        let data = build_xlsx_bytes(
            "Grid",
            &[("A1", "1"), ("C1", "3"), ("B2", "5"), ("C3", "9")],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 3, "Expected 3 rows");
        // All rows should have same number of columns
        for row in &tp.table.rows {
            assert_eq!(row.cells.len(), 3, "Expected 3 columns per row");
        }
    }

    // ----- Error handling -----

    #[test]
    fn test_parse_invalid_data_returns_error() {
        let parser = XlsxParser;
        let result = parser.parse(b"not an xlsx file", &ConvertOptions::default());
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ConvertError::Parse(_)),
            "Expected Parse error, got {err:?}"
        );
    }

    // ----- Empty cell content -----

    #[test]
    fn test_empty_cells_have_no_content() {
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Only A1"), ("C1", "Only C1")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        // B1 should be empty (no paragraphs)
        assert!(
            tp.table.rows[0].cells[1].content.is_empty(),
            "Expected empty cell content for B1"
        );
    }

    #[test]
    fn test_cell_default_span_values() {
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Test")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let cell = &tp.table.rows[0].cells[0];
        assert_eq!(cell.col_span, 1);
        assert_eq!(cell.row_span, 1);
        assert!(cell.border.is_none());
        assert!(cell.background.is_none());
    }

    // ----- Cell merging tests (US-015) -----

    /// Helper: build XLSX with merge ranges.
    fn build_xlsx_with_merges(
        sheet_name: &str,
        cells: &[(&str, &str)],
        merges: &[&str],
    ) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name(sheet_name);
            for &(coord, value) in cells {
                sheet.get_cell_mut(coord).set_value(value);
            }
            for &merge_range in merges {
                sheet.add_merge_cells(merge_range);
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_merge_colspan_basic() {
        // A1:B1 merged → colspan=2 on A1, B1 is skipped
        let data = build_xlsx_with_merges("Sheet1", &[("A1", "Merged")], &["A1:B1"]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(
            tp.table.rows[0].cells.len(),
            1,
            "Merged cells should produce 1 cell"
        );
        assert_eq!(tp.table.rows[0].cells[0].col_span, 2);
        assert_eq!(tp.table.rows[0].cells[0].row_span, 1);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Merged");
    }

    #[test]
    fn test_merge_rowspan_basic() {
        // A1:A2 merged → rowspan=2 on A1, A2 is skipped
        let data = build_xlsx_with_merges("Sheet1", &[("A1", "Tall")], &["A1:A2"]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        // Row 0: one cell with rowspan 2
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(tp.table.rows[0].cells[0].row_span, 2);
        assert_eq!(tp.table.rows[0].cells[0].col_span, 1);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Tall");
        // Row 1: no cells (the merged cell from row 0 spans here)
        assert_eq!(tp.table.rows[1].cells.len(), 0);
    }

    #[test]
    fn test_merge_colspan_and_rowspan() {
        // A1:B2 merged → colspan=2, rowspan=2 on A1
        let data = build_xlsx_with_merges(
            "Sheet1",
            &[("A1", "Big"), ("C1", "Right"), ("C2", "Below")],
            &["A1:B2"],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        // Row 0: merged cell (A1:B2) + C1
        assert_eq!(tp.table.rows[0].cells.len(), 2);
        assert_eq!(tp.table.rows[0].cells[0].col_span, 2);
        assert_eq!(tp.table.rows[0].cells[0].row_span, 2);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Big");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "Right");
        // Row 1: only C2 (A1:B2 merge continues from row 0)
        assert_eq!(tp.table.rows[1].cells.len(), 1);
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Below");
    }

    #[test]
    fn test_merge_content_in_top_left_only() {
        // Merge A1:B1, content only in A1
        let data = build_xlsx_with_merges(
            "Sheet1",
            &[("A1", "TopLeft"), ("B1", "should be ignored")],
            &["A1:B1"],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "TopLeft");
    }

    #[test]
    fn test_merge_multiple_ranges() {
        // Two merges: A1:B1, A2:A3
        let data = build_xlsx_with_merges(
            "Sheet1",
            &[("A1", "Wide"), ("A2", "Tall"), ("B2", "B2"), ("B3", "B3")],
            &["A1:B1", "A2:A3"],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        // Row 0: A1:B1 merged (colspan=2)
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(tp.table.rows[0].cells[0].col_span, 2);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Wide");
        // Row 1: A2:A3 (rowspan=2) + B2
        assert_eq!(tp.table.rows[1].cells.len(), 2);
        assert_eq!(tp.table.rows[1].cells[0].row_span, 2);
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Tall");
        assert_eq!(cell_text(&tp.table.rows[1].cells[1]), "B2");
        // Row 2: only B3 (A2:A3 continues)
        assert_eq!(tp.table.rows[2].cells.len(), 1);
        assert_eq!(cell_text(&tp.table.rows[2].cells[0]), "B3");
    }

    #[test]
    fn test_merge_no_merges_unchanged() {
        // No merges: cells should have default span values
        let data = build_xlsx_bytes("Sheet1", &[("A1", "X"), ("B1", "Y")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows[0].cells.len(), 2);
        for cell in &tp.table.rows[0].cells {
            assert_eq!(cell.col_span, 1);
            assert_eq!(cell.row_span, 1);
        }
    }

    #[test]
    fn test_merge_wide_colspan() {
        // A1:D1 merged → colspan=4
        let data = build_xlsx_with_merges("Sheet1", &[("A1", "Title")], &["A1:D1"]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(tp.table.rows[0].cells[0].col_span, 4);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Title");
    }

    // ----- US-027: Cell formatting tests -----

    /// Helper: build XLSX with formatted cells.
    fn build_xlsx_formatted(setup: impl FnOnce(&mut umya_spreadsheet::Worksheet)) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name("Sheet1");
            setup(sheet);
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: extract the first run's TextStyle from a cell.
    fn first_run_style(cell: &TableCell) -> &TextStyle {
        match &cell.content[0] {
            Block::Paragraph(p) => &p.runs[0].style,
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_cell_bold_text() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Bold");
            cell.get_style_mut().get_font_mut().set_bold(true);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(style.bold, Some(true));
    }

    #[test]
    fn test_cell_italic_text() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Italic");
            cell.get_style_mut().get_font_mut().set_italic(true);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(style.italic, Some(true));
    }

    #[test]
    fn test_cell_font_color() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Red");
            cell.get_style_mut()
                .get_font_mut()
                .get_color_mut()
                .set_argb("FFFF0000");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(style.color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_cell_font_name_and_size() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Styled");
            let font = cell.get_style_mut().get_font_mut();
            font.set_name("Arial");
            font.set_size(14.0);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(style.font_family.as_deref(), Some("Arial"));
        assert_eq!(style.font_size, Some(14.0));
    }

    #[test]
    fn test_cell_background_fill() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Yellow BG");
            cell.get_style_mut().set_background_color("FFFFFF00");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let cell = &tp.table.rows[0].cells[0];
        assert_eq!(cell.background, Some(Color::new(255, 255, 0)));
    }

    #[test]
    fn test_cell_borders() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Bordered");
            let borders = cell.get_style_mut().get_borders_mut();
            borders
                .get_bottom_mut()
                .set_border_style(umya_spreadsheet::Border::BORDER_MEDIUM);
            borders
                .get_bottom_mut()
                .get_color_mut()
                .set_argb("FF000000");
            borders
                .get_top_mut()
                .set_border_style(umya_spreadsheet::Border::BORDER_THIN);
            borders.get_top_mut().get_color_mut().set_argb("FFFF0000");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let cell = &tp.table.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected border");
        // Bottom: medium → ~1pt, black
        let bottom = border.bottom.as_ref().expect("Expected bottom border");
        assert!((bottom.width - 1.0).abs() < 0.01);
        assert_eq!(bottom.color, Color::new(0, 0, 0));
        // Top: thin → ~0.5pt, red
        let top = border.top.as_ref().expect("Expected top border");
        assert!((top.width - 0.5).abs() < 0.01);
        assert_eq!(top.color, Color::new(255, 0, 0));
    }

    #[test]
    fn test_row_height() {
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_cell_mut("A1").set_value("Tall row");
            sheet.get_row_dimension_mut(&1).set_height(30.0);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let row = &tp.table.rows[0];
        assert_eq!(row.height, Some(30.0));
    }

    #[test]
    fn test_cell_no_formatting_defaults() {
        // Plain cell with no explicit formatting → default TextStyle, no border, no background
        let data = build_xlsx_bytes("Sheet1", &[("A1", "Plain")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let cell = &tp.table.rows[0].cells[0];
        let style = first_run_style(cell);
        // No explicit formatting → all None
        assert!(style.bold.is_none() || style.bold == Some(false));
        assert!(style.italic.is_none() || style.italic == Some(false));
        assert!(cell.border.is_none());
        assert!(cell.background.is_none());
    }

    // ----- US-028: Number format tests -----

    #[test]
    fn test_number_format_currency() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(1234.56f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_CURRENCY_USD_SIMPLE);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        // Should contain $ and formatted number, not raw "1234.56"
        assert!(
            text.contains('$') && text.contains("1,234.56"),
            "Expected currency format with $ and 1,234.56, got: {text}"
        );
    }

    #[test]
    fn test_number_format_percentage() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(0.456f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_PERCENTAGE);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        // 0.456 with "0%" format → "46%" (rounded)
        assert!(
            text.contains('%'),
            "Expected percentage format with %, got: {text}"
        );
    }

    #[test]
    fn test_number_format_percentage_with_decimals() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(0.5f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_PERCENTAGE_00);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        // 0.5 with "0.00%" format → "50.00%"
        assert!(
            text.contains('%') && text.contains("50.00"),
            "Expected 50.00%, got: {text}"
        );
    }

    #[test]
    fn test_number_format_date() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            // Excel serial number for a date (e.g., 45306 = 2024-01-15 approximately)
            cell.set_value_number(45306f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_DATE_YYYYMMDD);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        // Should be a date string like "2024-01-05" (exact date depends on serial), NOT "45306"
        assert!(
            text.contains('-') && !text.contains("45306"),
            "Expected date format yyyy-mm-dd, got: {text}"
        );
    }

    #[test]
    fn test_number_format_thousands_separator() {
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(1234567f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code("#,##0");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        assert_eq!(text, "1,234,567", "Expected thousands separator formatting");
    }

    #[test]
    fn test_number_format_general_unchanged() {
        // General format should not change the display of simple numbers
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_cell_mut("A1").set_value("42");
            sheet.get_cell_mut("B1").set_value("3.14");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "42");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "3.14");
    }

    #[test]
    fn test_number_format_builtin_id() {
        // Use a built-in format ID (ID 4 = "#,##0.00")
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(1234.5f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_number_format_id(4);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        // Format ID 4 = "#,##0.00" → should have thousands separator and decimals
        assert!(
            text.contains("1,234") && text.contains("50"),
            "Expected #,##0.00 formatting via ID 4, got: {text}"
        );
    }

    #[test]
    fn test_number_format_custom_format_string() {
        // Custom format: display with 3 decimal places
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(std::f64::consts::PI);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code("0.000");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let text = cell_text(&tp.table.rows[0].cells[0]);
        assert_eq!(text, "3.142", "Expected 3 decimal places formatting");
    }

    #[test]
    fn test_cell_combined_formatting() {
        // Cell with font + background + border all at once
        let data = build_xlsx_formatted(|sheet| {
            let cell = sheet.get_cell_mut("A1");
            cell.set_value("Full");
            let style = cell.get_style_mut();
            let font = style.get_font_mut();
            font.set_bold(true);
            font.set_size(16.0);
            font.set_name("Helvetica");
            font.get_color_mut().set_argb("FF0000FF"); // Blue text
            style.set_background_color("FFFFCC00"); // Orange BG
            let borders = style.get_borders_mut();
            borders
                .get_left_mut()
                .set_border_style(umya_spreadsheet::Border::BORDER_THICK);
            borders.get_left_mut().get_color_mut().set_argb("FF00FF00"); // Green border
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let cell = &tp.table.rows[0].cells[0];
        let style = first_run_style(cell);
        assert_eq!(style.bold, Some(true));
        assert_eq!(style.font_size, Some(16.0));
        assert_eq!(style.font_family.as_deref(), Some("Helvetica"));
        assert_eq!(style.color, Some(Color::new(0, 0, 255)));
        assert_eq!(cell.background, Some(Color::new(255, 204, 0)));
        let border = cell.border.as_ref().expect("Expected border");
        let left = border.left.as_ref().expect("Expected left border");
        assert!((left.width - 2.0).abs() < 0.01);
        assert_eq!(left.color, Color::new(0, 255, 0));
    }

    // ----- US-029: Sheet selection tests -----

    #[test]
    fn test_sheet_filter_single_sheet() {
        let data = build_xlsx_multi_sheet(&[
            ("Sales", &[("A1", "Revenue")]),
            ("Expenses", &[("A1", "Cost")]),
            ("Summary", &[("A1", "Total")]),
        ]);
        let parser = XlsxParser;
        let opts = ConvertOptions {
            sheet_names: Some(vec!["Expenses".to_string()]),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 1, "Should only include 1 sheet");
        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.name, "Expenses");
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Cost");
    }

    #[test]
    fn test_sheet_filter_multiple_sheets() {
        let data = build_xlsx_multi_sheet(&[
            ("Sales", &[("A1", "Revenue")]),
            ("Expenses", &[("A1", "Cost")]),
            ("Summary", &[("A1", "Total")]),
        ]);
        let parser = XlsxParser;
        let opts = ConvertOptions {
            sheet_names: Some(vec!["Sales".to_string(), "Summary".to_string()]),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 2, "Should include 2 sheets");
        let tp0 = get_table_page(&doc, 0);
        let tp1 = get_table_page(&doc, 1);
        assert_eq!(tp0.name, "Sales");
        assert_eq!(tp1.name, "Summary");
    }

    #[test]
    fn test_sheet_filter_none_includes_all() {
        let data =
            build_xlsx_multi_sheet(&[("Sheet1", &[("A1", "A")]), ("Sheet2", &[("A1", "B")])]);
        let parser = XlsxParser;
        let opts = ConvertOptions {
            sheet_names: None,
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 2, "None should include all sheets");
    }

    #[test]
    fn test_sheet_filter_nonexistent_name() {
        let data =
            build_xlsx_multi_sheet(&[("Sheet1", &[("A1", "A")]), ("Sheet2", &[("A1", "B")])]);
        let parser = XlsxParser;
        let opts = ConvertOptions {
            sheet_names: Some(vec!["DoesNotExist".to_string()]),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(
            doc.pages.len(),
            0,
            "No matching sheets should produce empty document"
        );
    }

    // ----- US-035: Print area and page breaks tests -----

    /// Helper: build XLSX with a print area defined name.
    fn build_xlsx_with_print_area(cells: &[(&str, &str)], print_area: &str) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name("Sheet1");
            for &(coord, value) in cells {
                sheet.get_cell_mut(coord).set_value(value);
            }
            sheet
                .add_defined_name("_xlnm.Print_Area", print_area)
                .unwrap();
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build XLSX with row page breaks.
    fn build_xlsx_with_row_breaks(cells: &[(&str, &str)], break_rows: &[u32]) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name("Sheet1");
            for &(coord, value) in cells {
                sheet.get_cell_mut(coord).set_value(value);
            }
            for &row in break_rows {
                let mut brk = umya_spreadsheet::Break::default();
                brk.set_id(row);
                brk.set_manual_page_break(true);
                sheet.get_row_breaks_mut().add_break_list(brk);
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_print_area_limits_output() {
        // Sheet has data in A1:D4, but print area is A1:B2
        let data = build_xlsx_with_print_area(
            &[
                ("A1", "In"),
                ("B1", "In"),
                ("C1", "Out"),
                ("D1", "Out"),
                ("A2", "In"),
                ("B2", "In"),
                ("C2", "Out"),
                ("A3", "Out"),
                ("B3", "Out"),
                ("A4", "Out"),
            ],
            "Sheet1!$A$1:$B$2",
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let tp = get_table_page(&doc, 0);
        // Only rows 1-2, columns A-B
        assert_eq!(tp.table.rows.len(), 2, "Should have 2 rows from print area");
        assert_eq!(
            tp.table.rows[0].cells.len(),
            2,
            "Should have 2 columns from print area"
        );
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "In");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "In");
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "In");
        assert_eq!(cell_text(&tp.table.rows[1].cells[1]), "In");
        // Column widths should only have 2 entries
        assert_eq!(tp.table.column_widths.len(), 2);
    }

    #[test]
    fn test_print_area_without_dollar_signs() {
        // Print area without dollar signs should also work
        let data = build_xlsx_with_print_area(
            &[("A1", "X"), ("B1", "Y"), ("A2", "Z"), ("B2", "W")],
            "Sheet1!A1:A2",
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 2);
        assert_eq!(tp.table.rows[0].cells.len(), 1, "Only column A");
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "X");
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Z");
    }

    #[test]
    fn test_no_print_area_includes_all() {
        // Without print area, all data should be included (existing behavior)
        let data = build_xlsx_bytes("Sheet1", &[("A1", "All"), ("C3", "Data")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 3);
        assert_eq!(tp.table.rows[0].cells.len(), 3);
    }

    #[test]
    fn test_row_page_breaks_split_into_pages() {
        // 5 rows of data, page break after row 2
        let data = build_xlsx_with_row_breaks(
            &[
                ("A1", "R1"),
                ("A2", "R2"),
                ("A3", "R3"),
                ("A4", "R4"),
                ("A5", "R5"),
            ],
            &[2], // break after row 2
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        // Should produce 2 pages: rows 1-2 and rows 3-5
        assert_eq!(doc.pages.len(), 2, "Break should split into 2 pages");
        let tp0 = get_table_page(&doc, 0);
        let tp1 = get_table_page(&doc, 1);

        assert_eq!(tp0.table.rows.len(), 2, "First page: rows 1-2");
        assert_eq!(cell_text(&tp0.table.rows[0].cells[0]), "R1");
        assert_eq!(cell_text(&tp0.table.rows[1].cells[0]), "R2");

        assert_eq!(tp1.table.rows.len(), 3, "Second page: rows 3-5");
        assert_eq!(cell_text(&tp1.table.rows[0].cells[0]), "R3");
        assert_eq!(cell_text(&tp1.table.rows[1].cells[0]), "R4");
        assert_eq!(cell_text(&tp1.table.rows[2].cells[0]), "R5");
    }

    #[test]
    fn test_multiple_row_page_breaks() {
        // 6 rows, breaks after rows 2 and 4
        let data = build_xlsx_with_row_breaks(
            &[
                ("A1", "R1"),
                ("A2", "R2"),
                ("A3", "R3"),
                ("A4", "R4"),
                ("A5", "R5"),
                ("A6", "R6"),
            ],
            &[2, 4], // breaks after row 2 and row 4
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        // Should produce 3 pages: rows 1-2, rows 3-4, rows 5-6
        assert_eq!(doc.pages.len(), 3, "Two breaks should produce 3 pages");
        let tp0 = get_table_page(&doc, 0);
        let tp1 = get_table_page(&doc, 1);
        let tp2 = get_table_page(&doc, 2);

        assert_eq!(tp0.table.rows.len(), 2);
        assert_eq!(tp1.table.rows.len(), 2);
        assert_eq!(tp2.table.rows.len(), 2);

        assert_eq!(cell_text(&tp0.table.rows[0].cells[0]), "R1");
        assert_eq!(cell_text(&tp1.table.rows[0].cells[0]), "R3");
        assert_eq!(cell_text(&tp2.table.rows[0].cells[0]), "R5");
    }

    #[test]
    fn test_no_page_breaks_single_page() {
        // No page breaks → single page per sheet (existing behavior)
        let data = build_xlsx_bytes("Sheet1", &[("A1", "R1"), ("A2", "R2"), ("A3", "R3")]);
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows.len(), 3);
    }

    #[test]
    fn test_page_break_column_widths_preserved() {
        // Page breaks should preserve column widths across all pages
        let data = build_xlsx_with_row_breaks(
            &[("A1", "R1"), ("B1", "R1B"), ("A2", "R2"), ("B2", "R2B")],
            &[1],
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 2);
        let tp0 = get_table_page(&doc, 0);
        let tp1 = get_table_page(&doc, 1);
        assert_eq!(tp0.table.column_widths.len(), 2);
        assert_eq!(tp1.table.column_widths.len(), 2);
        // Same column widths on both pages
        assert_eq!(tp0.table.column_widths, tp1.table.column_widths);
    }

    // --- US-036: Sheet headers and footers ---

    #[test]
    fn test_parse_hf_format_string_empty() {
        assert!(parse_hf_format_string("").is_none());
        assert!(parse_hf_format_string("   ").is_none());
    }

    #[test]
    fn test_parse_hf_format_string_center_only() {
        // No section prefix → defaults to center
        let hf = parse_hf_format_string("My Report").unwrap();
        assert_eq!(hf.paragraphs.len(), 1);
        assert_eq!(hf.paragraphs[0].style.alignment, Some(Alignment::Center));
        assert_eq!(hf.paragraphs[0].elements.len(), 1);
        match &hf.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "My Report"),
            _ => panic!("Expected Run"),
        }
    }

    #[test]
    fn test_parse_hf_format_string_left_center_right() {
        let hf = parse_hf_format_string("&LLeft Text&CCenter Text&RRight Text").unwrap();
        assert_eq!(hf.paragraphs.len(), 3);

        // Left section
        assert_eq!(hf.paragraphs[0].style.alignment, Some(Alignment::Left));
        match &hf.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Left Text"),
            _ => panic!("Expected Run"),
        }

        // Center section
        assert_eq!(hf.paragraphs[1].style.alignment, Some(Alignment::Center));
        match &hf.paragraphs[1].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Center Text"),
            _ => panic!("Expected Run"),
        }

        // Right section
        assert_eq!(hf.paragraphs[2].style.alignment, Some(Alignment::Right));
        match &hf.paragraphs[2].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Right Text"),
            _ => panic!("Expected Run"),
        }
    }

    #[test]
    fn test_parse_hf_format_string_page_numbers() {
        // Footer with "Page X of Y"
        let hf = parse_hf_format_string("&CPage &P of &N").unwrap();
        assert_eq!(hf.paragraphs.len(), 1);
        let elems = &hf.paragraphs[0].elements;
        assert_eq!(elems.len(), 4);
        match &elems[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Page "),
            _ => panic!("Expected Run"),
        }
        assert!(matches!(elems[1], HFInline::PageNumber));
        match &elems[2] {
            HFInline::Run(r) => assert_eq!(r.text, " of "),
            _ => panic!("Expected Run"),
        }
        assert!(matches!(elems[3], HFInline::TotalPages));
    }

    #[test]
    fn test_parse_hf_format_string_escaped_ampersand() {
        let hf = parse_hf_format_string("&CA && B").unwrap();
        assert_eq!(hf.paragraphs.len(), 1);
        match &hf.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "A & B"),
            _ => panic!("Expected Run"),
        }
    }

    #[test]
    fn test_parse_hf_format_string_font_codes_skipped() {
        // Font name and size codes should be skipped
        let hf = parse_hf_format_string(r#"&C&"Arial"&12Hello"#).unwrap();
        assert_eq!(hf.paragraphs.len(), 1);
        match &hf.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Hello"),
            _ => panic!("Expected Run"),
        }
    }

    /// Helper: build an XLSX with a custom header on the sheet.
    fn build_xlsx_with_header(header_str: &str) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.get_cell_mut("A1").set_value("Data");
            sheet
                .get_header_footer_mut()
                .get_odd_header_mut()
                .set_value(header_str);
        }
        let mut buf = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut buf).unwrap();
        buf.into_inner()
    }

    /// Helper: build an XLSX with a custom footer on the sheet.
    fn build_xlsx_with_footer(footer_str: &str) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.get_cell_mut("A1").set_value("Data");
            sheet
                .get_header_footer_mut()
                .get_odd_footer_mut()
                .set_value(footer_str);
        }
        let mut buf = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut buf).unwrap();
        buf.into_inner()
    }

    #[test]
    fn test_xlsx_sheet_with_custom_header() {
        let data = build_xlsx_with_header("&CMonthly Report");
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let header = tp.header.as_ref().expect("Expected header");
        assert_eq!(header.paragraphs.len(), 1);
        assert_eq!(
            header.paragraphs[0].style.alignment,
            Some(Alignment::Center)
        );
        match &header.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "Monthly Report"),
            _ => panic!("Expected Run"),
        }
    }

    #[test]
    fn test_xlsx_sheet_with_page_number_footer() {
        let data = build_xlsx_with_footer("&CPage &P of &N");
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_table_page(&doc, 0);
        let footer = tp.footer.as_ref().expect("Expected footer");
        assert_eq!(footer.paragraphs.len(), 1);
        let elems = &footer.paragraphs[0].elements;
        assert_eq!(elems.len(), 4); // "Page ", PageNumber, " of ", TotalPages
        assert!(matches!(elems[1], HFInline::PageNumber));
        assert!(matches!(elems[3], HFInline::TotalPages));
    }
}
