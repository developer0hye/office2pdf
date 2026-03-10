use std::collections::{HashMap, HashSet};

use crate::ir::{Block, Paragraph, ParagraphStyle, Run, TableRow};
use crate::parser::cond_fmt::build_cond_fmt_overrides;

use super::xlsx_style::{extract_cell_background, extract_cell_borders, extract_cell_text_style};
use crate::ir::TableCell;

/// A cell range within a sheet (1-indexed, inclusive).
#[derive(Debug, Clone, Copy)]
pub(crate) struct CellRange {
    pub(crate) start_col: u32,
    pub(crate) start_row: u32,
    pub(crate) end_col: u32,
    pub(crate) end_row: u32,
}

/// A (column, row) coordinate pair (1-indexed).
pub(crate) type CellPos = (u32, u32);

/// Info about a merged cell region, keyed by its top-left coordinate.
pub(super) struct MergeInfo {
    pub(super) col_span: u32,
    pub(super) row_span: u32,
}

/// Default column width in Excel character units.
pub(super) const DEFAULT_COLUMN_WIDTH: f64 = 8.43;

/// Convert Excel column width (character units) to points.
/// Excel character width ≈ 7 pixels at 96 DPI, 1 point = 96/72 pixels.
/// Empirically: width_pt ≈ char_width * 7.0 (approximate, close to Excel's rendering).
pub(super) fn column_width_to_pt(char_width: f64) -> f64 {
    char_width * 7.0
}

/// Parse an Excel column letter string (e.g., "A", "B", "AA") into a 1-indexed column number.
pub(super) fn parse_column_letters(s: &str) -> Option<u32> {
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
pub(crate) fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
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
pub(super) fn parse_print_area_range(address: &str) -> Option<CellRange> {
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
pub(super) fn find_print_area(sheet: &umya_spreadsheet::Worksheet) -> Option<CellRange> {
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
pub(super) fn collect_row_breaks(sheet: &umya_spreadsheet::Worksheet) -> Vec<u32> {
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

/// Build a lookup of merge info from the sheet's merged cell ranges.
///
/// Returns two structures:
/// - `top_left_map`: top-left coordinate → MergeInfo for each merge
/// - `skip_set`: set of coordinates that are inside a merge but NOT the top-left
pub(super) fn build_merge_maps(
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

/// Shared context for processing a single XLSX sheet.
pub(super) struct SheetContext {
    pub(super) col_start: u32,
    pub(super) col_end: u32,
    pub(super) num_cols: usize,
    pub(super) column_widths: Vec<f64>,
    pub(super) merge_tops: HashMap<(u32, u32), MergeInfo>,
    pub(super) merge_skips: HashSet<(u32, u32)>,
    pub(super) cond_fmt_overrides: HashMap<(u32, u32), crate::parser::cond_fmt::CondFmtOverride>,
}

/// Build TableRows for a range of rows in a sheet.
pub(super) fn build_rows_for_range(
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
/// Returns (SheetContext, row_start, row_end) or None if the sheet is empty.
pub(super) fn prepare_sheet_context(
    sheet: &umya_spreadsheet::Worksheet,
) -> Option<(SheetContext, u32, u32)> {
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
