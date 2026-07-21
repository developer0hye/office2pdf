use std::collections::{HashMap, HashSet};

use crate::ir::{Block, Paragraph, ParagraphStyle, Run, TableRow};
use crate::parser::cond_fmt::build_cond_fmt_overrides;

use super::xlsx_style::{
    apply_rich_run_font, extract_cell_alignment, extract_cell_background, extract_cell_borders,
    extract_cell_text_style,
};
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
const DEFAULT_MAX_DIGIT_WIDTH_PX: f64 = 7.0;

/// Convert Excel column width (character units) to points.
/// OOXML widths are expressed relative to the maximum digit width (MDW) of
/// the worksheet's Normal font. The stored width already incorporates Excel's
/// cell padding adjustment, so print geometry must not add another 5 pixels.
pub(super) fn column_width_to_pt(char_width: f64, max_digit_width_px: f64) -> f64 {
    char_width * max_digit_width_px * 0.75
}

/// Read the workbook's Normal font (the first `<font>` in `xl/styles.xml`)
/// straight from the archive; umya does not expose the stylesheet. Excel
/// derives all column print metrics from this font, not from cell fonts.
pub(super) fn extract_normal_font(data: &[u8]) -> Option<(String, f64)> {
    use quick_xml::events::Event;
    use std::io::Read;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(data)).ok()?;
    let mut file = archive.by_name("xl/styles.xml").ok()?;
    let mut xml = String::new();
    file.read_to_string(&mut xml).ok()?;

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut in_first_font = false;
    let mut name: Option<String> = None;
    let mut size: Option<f64> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                in_first_font = true;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Empty(ref e)) if in_first_font => {
                let val = e
                    .try_get_attribute("val")
                    .ok()
                    .flatten()
                    .and_then(|a| String::from_utf8(a.value.into_owned()).ok());
                match e.local_name().as_ref() {
                    b"name" => name = val,
                    b"sz" => size = val.and_then(|v| v.parse::<f64>().ok()),
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Some((name?, size.unwrap_or(11.0)))
}

/// Excel pixel-ceils the Normal font's max digit width at 96 DPI to derive
/// column print metrics. Digit advances: Calibri/Carlito 0.5066 em,
/// Arial/Helvetica/Liberation Sans 0.556 em, Malgun Gothic ≈0.529 em.
/// Calibri 11 gives ceil(7.43) = 8 px, which reproduces the native Excel
/// print pagination of the audit fixtures (issues #330, #366).
pub(super) fn max_digit_width_px_for_normal_font(family: &str, size_pt: f64) -> f64 {
    let digit_em: f64 = match family.to_ascii_lowercase().as_str() {
        "arial" | "helvetica" | "liberation sans" => 0.556,
        "malgun gothic" | "맑은 고딕" => 0.529,
        _ => 0.5066,
    };
    (digit_em * size_pt * (96.0 / 72.0)).ceil()
}

/// Fallback when `xl/styles.xml` is unreadable: infer the metric from the
/// dominant cell font. umya resolves each cell's effective style while
/// reading, so the dominant family is a stable approximation.
pub(super) fn sheet_max_digit_width_px(sheet: &umya_spreadsheet::Worksheet) -> f64 {
    let mut family_counts: HashMap<String, usize> = HashMap::new();
    for cell in sheet.get_cell_collection() {
        let Some(font) = cell.get_style().get_font() else {
            continue;
        };
        let family = font.get_name().trim();
        if !family.is_empty() {
            *family_counts
                .entry(family.to_ascii_lowercase())
                .or_default() += 1;
        }
    }

    let dominant_family = family_counts
        .into_iter()
        .max_by(|(family_a, count_a), (family_b, count_b)| {
            count_a.cmp(count_b).then_with(|| family_b.cmp(family_a))
        })
        .map(|(family, _)| family);

    match dominant_family.as_deref() {
        // Excel's macOS print output for Carlito 11 uses an 8px MDW. This
        // yields the fixture's native 26/20/24-char widths of 156/120/144pt.
        Some("carlito") => 8.0,
        _ => DEFAULT_MAX_DIGIT_WIDTH_PX,
    }
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

/// Print-title ranges from `_xlnm.Print_Titles`: rows and/or columns that
/// Excel repeats on every printed page (1-indexed, inclusive).
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct PrintTitles {
    pub(super) rows: Option<(u32, u32)>,
    pub(super) cols: Option<(u32, u32)>,
}

/// Look up the sheet's print titles. The defined name holds one or two
/// comma-separated parts like `Sheet4!$A:$B,Sheet4!$2:$3`. Sheet-scoped
/// names (localSheetId) land on the worksheet; names the reader could not
/// scope stay at the workbook level, so both are consulted.
pub(super) fn find_print_titles(
    book: &umya_spreadsheet::Spreadsheet,
    sheet: &umya_spreadsheet::Worksheet,
) -> PrintTitles {
    let mut titles = PrintTitles::default();
    for dn in sheet.get_defined_names() {
        if dn.get_name() == "_xlnm.Print_Titles" {
            parse_print_title_address(&dn.get_address(), &mut titles);
        }
    }
    if titles.rows.is_none() && titles.cols.is_none() {
        let plain_prefix: String = format!("{}!", sheet.get_name());
        let quoted_prefix: String = format!("'{}'!", sheet.get_name());
        for dn in book.get_defined_names() {
            let address: String = dn.get_address();
            if dn.get_name() == "_xlnm.Print_Titles"
                && (address.contains(&plain_prefix) || address.contains(&quoted_prefix))
            {
                parse_print_title_address(&address, &mut titles);
            }
        }
    }
    titles
}

fn parse_print_title_address(address: &str, titles: &mut PrintTitles) {
    for part in address.split(',') {
        let range_part: String = part
            .rsplit('!')
            .next()
            .unwrap_or(part)
            .replace('$', "")
            .trim()
            .to_string();
        let Some((start_str, end_str)) = range_part.split_once(':') else {
            continue;
        };
        if let (Ok(row_start), Ok(row_end)) = (start_str.parse::<u32>(), end_str.parse::<u32>()) {
            titles.rows = Some((row_start.min(row_end), row_start.max(row_end)));
        } else if let (Some(col_start), Some(col_end)) = (
            parse_column_letters(start_str),
            parse_column_letters(end_str),
        ) {
            titles.cols = Some((col_start.min(col_end), col_start.max(col_end)));
        }
    }
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
    pub(super) max_digit_width_px: f64,
    pub(super) merge_tops: HashMap<(u32, u32), MergeInfo>,
    pub(super) merge_skips: HashSet<(u32, u32)>,
    pub(super) cond_fmt_overrides: HashMap<(u32, u32), crate::parser::cond_fmt::CondFmtOverride>,
}

/// First strong bidi direction of a character: Some(true) for right-to-left
/// scripts (Hebrew, Arabic and its supplements), Some(false) for Latin-like
/// letters, None for neutral characters (digits, punctuation, spaces).
fn strong_direction(c: char) -> Option<bool> {
    match c as u32 {
        // Hebrew, Arabic, Syriac, Thaana, and Arabic presentation forms.
        0x0590..=0x08FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF => Some(true),
        _ if c.is_alphabetic() => Some(false),
        _ => None,
    }
}

/// Map ASCII digits (and separators) to Arabic-Indic digits, as Excel does
/// for number formats carrying a native-digit locale prefix like
/// `[$-3000401]`.
fn to_arabic_indic_digits(value: &str) -> String {
    value
        .chars()
        .map(|c| match c {
            '0'..='9' => char::from_u32(0x0660 + (c as u32 - '0' as u32)).unwrap_or(c),
            '.' => '\u{066B}',
            ',' => '\u{066C}',
            _ => c,
        })
        .collect()
}

/// Excel number formats may carry a locale prefix `[$-XXXXXXXX]` whose high
/// byte selects digit shaping (>= 2 substitutes national digits). Arabic
/// primary language (low byte 0x01) then prints Arabic-Indic digits.
fn uses_native_arabic_digits(format_code: &str) -> bool {
    let Some(rest) = format_code.strip_prefix("[$-") else {
        return false;
    };
    let Some(end) = rest.find(']') else {
        return false;
    };
    let Ok(locale) = u64::from_str_radix(&rest[..end], 16) else {
        return false;
    };
    let digit_substitution: u64 = locale >> 24;
    let language_id: u64 = locale & 0xFF;
    digit_substitution >= 2 && language_id == 0x01
}

/// Rough single-line text width estimate in points: ASCII glyphs average
/// about half the font size in Calibri-class fonts, CJK glyphs are full-width.
fn estimate_text_width_pt(runs: &[Run]) -> f64 {
    runs.iter()
        .map(|run| {
            let font_size: f64 = run.style.font_size.unwrap_or(11.0);
            run.text
                .chars()
                .map(|c| {
                    if c.is_ascii() {
                        0.55 * font_size
                    } else {
                        1.05 * font_size
                    }
                })
                .sum::<f64>()
        })
        .sum()
}

/// Excel lets an unwrapped left-aligned text overflow into consecutive empty
/// cells to its right without growing the row. Returns the total width the
/// content may paint across (own column + empty neighbors), or None when the
/// text fits, wraps, or has no empty neighbor to spill into.
#[allow(clippy::too_many_arguments)]
fn compute_spill_width(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    col_idx: u32,
    row_idx: u32,
    runs: &[Run],
    cell_alignment: Option<crate::ir::Alignment>,
    col_span: u32,
    umya_cell: Option<&umya_spreadsheet::Cell>,
) -> Option<f64> {
    if runs.is_empty() {
        return None;
    }
    // Only Excel's "general"/left horizontal alignment spills to the right.
    if !matches!(cell_alignment, None | Some(crate::ir::Alignment::Left)) {
        return None;
    }
    // Explicit wrapText wraps inside the cell instead.
    let has_wrap_text: bool = umya_cell
        .and_then(|cell| cell.get_style().get_alignment().cloned())
        .map(|alignment| *alignment.get_wrap_text())
        .unwrap_or(false);
    if has_wrap_text {
        return None;
    }
    // Embedded line breaks always wrap.
    if runs.iter().any(|run| run.text.contains('\n')) {
        return None;
    }

    // A merged cell never paints past the merge edge: Excel keeps unwrapped
    // text on one line and clips it at the merged width. Apply this even when
    // the text fits — column pagination may clamp the merge to fewer columns
    // on a page, and Excel still paints the single line across the page edge
    // rather than wrapping it.
    if col_span > 1 {
        let merged_width: f64 = (col_idx..col_idx + col_span)
            .map(|c| {
                ctx.column_widths
                    .get((c - ctx.col_start) as usize)
                    .copied()
                    .unwrap_or(0.0)
            })
            .sum();
        return Some(merged_width);
    }

    let own_width: f64 = *ctx.column_widths.get((col_idx - ctx.col_start) as usize)?;
    // Leave room for the ~4pt total horizontal cell inset.
    if estimate_text_width_pt(runs) <= own_width - 4.0 {
        return None;
    }

    let mut total_width: f64 = own_width;
    let mut has_empty_neighbor = false;
    let mut blocked = false;
    for neighbor_col in (col_idx + 1)..=ctx.col_end {
        // Merged regions block the spill like occupied cells do.
        if ctx.merge_skips.contains(&(neighbor_col, row_idx))
            || ctx.merge_tops.contains_key(&(neighbor_col, row_idx))
        {
            blocked = true;
            break;
        }
        let neighbor_is_empty: bool = sheet
            .get_cell((neighbor_col, row_idx))
            .map(|cell| cell.get_formatted_value().is_empty())
            .unwrap_or(true);
        if !neighbor_is_empty {
            blocked = true;
            break;
        }
        total_width += *ctx
            .column_widths
            .get((neighbor_col - ctx.col_start) as usize)
            .unwrap_or(&0.0);
        has_empty_neighbor = true;
    }

    // Every used cell to the right is empty: Excel keeps painting across
    // the virtual empty cells beyond the used range toward the page edge.
    // Give the text the width it needs; the page boundary clips the rest.
    if !blocked {
        let needed_width: f64 = estimate_text_width_pt(runs) + 4.0;
        if needed_width > total_width {
            total_width = needed_width;
            has_empty_neighbor = true;
        }
    }

    has_empty_neighbor.then_some(total_width)
}

/// Excel's fallback row height when the sheet declares none (Calibri 11).
const EXCEL_DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;

/// The height a row prints at. A recorded `ht` is the actual current height
/// even when `customHeight` is false; rows without one use the sheet's
/// defaultRowHeight. Exception: auto-sized rows (customHeight=false) that
/// contain wrapped cells stay content-driven — our text metrics differ
/// slightly from Excel's and a fixed height could clip a wrapped line.
fn printed_row_height(
    sheet: &umya_spreadsheet::Worksheet,
    row_idx: u32,
    row_has_wrapping_cell: &dyn Fn() -> bool,
) -> Option<f64> {
    let row_dimension = sheet.get_row_dimension(&row_idx);
    let is_custom_height: bool = row_dimension
        .map(|row| *row.get_custom_height())
        .unwrap_or(false);
    if !is_custom_height && row_has_wrapping_cell() {
        return None;
    }
    let declared_height: Option<f64> = row_dimension
        .map(|row| *row.get_height())
        .filter(|height| *height > 0.0);
    declared_height.or_else(|| {
        let sheet_default: f64 = *sheet.get_sheet_format_properties().get_default_row_height();
        if sheet_default > 0.0 {
            Some(sheet_default)
        } else {
            Some(EXCEL_DEFAULT_ROW_HEIGHT_PT)
        }
    })
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
            let mut value = umya_cell
                .map(|cell| cell.get_formatted_value())
                .unwrap_or_default();
            if let Some(cell) = umya_cell
                && let Some(number_format) = cell.get_style().get_number_format()
                && uses_native_arabic_digits(number_format.get_format_code())
            {
                value = to_arabic_indic_digits(&value);
            }

            // Extract formatting from the cell
            let mut text_style = umya_cell.map(extract_cell_text_style).unwrap_or_default();
            let (cell_alignment, cell_vertical_align) = umya_cell
                .map(extract_cell_alignment)
                .unwrap_or((None, None));
            let mut background = umya_cell.and_then(extract_cell_background);
            let border = umya_cell.and_then(extract_cell_borders);

            // Apply conditional formatting overrides
            let mut data_bar = None;
            let mut icon_text = None;
            let mut icon_color = None;
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
                icon_color = ovr.icon_color;
            }

            // Rich-text shared strings carry per-run formatting (bold labels,
            // per-run fonts/colors) that the cell's single xf style loses —
            // emit one IR run per rich run instead of flattening.
            let rich_text: Option<umya_spreadsheet::RichText> =
                umya_cell.and_then(|cell| cell.get_cell_value().get_raw_value().get_rich_text());
            let runs: Vec<Run> = if let Some(rich_text) = rich_text {
                rich_text
                    .get_rich_text_elements()
                    .iter()
                    .filter(|element| !element.get_text().is_empty())
                    .map(|element| Run {
                        text: element.get_text().to_string(),
                        style: element
                            .get_run_properties()
                            .map(|font| apply_rich_run_font(&text_style, font))
                            .unwrap_or_else(|| text_style.clone()),
                        href: None,
                        footnote: None,
                    })
                    .collect()
            } else if value.is_empty() {
                Vec::new()
            } else {
                vec![Run {
                    text: value,
                    style: text_style,
                    href: None,
                    footnote: None,
                }]
            };

            // Excel's "general" horizontal alignment follows the text
            // direction: cells whose text starts with a right-to-left script
            // print right-aligned.
            let cell_alignment: Option<crate::ir::Alignment> = cell_alignment.or_else(|| {
                runs.iter()
                    .flat_map(|run| run.text.chars())
                    .find_map(strong_direction)
                    .filter(|is_rtl| *is_rtl)
                    .map(|_| crate::ir::Alignment::Right)
            });
            let paragraph_alignment = cell_alignment.or_else(|| {
                umya_cell
                    .and_then(|cell| cell.get_value_number())
                    .map(|_| crate::ir::Alignment::Right)
            });

            let (col_span, row_span) = if let Some(info) = ctx.merge_tops.get(&(col_idx, row_idx)) {
                (info.col_span, info.row_span)
            } else {
                (1, 1)
            };

            let spill_width: Option<f64> = compute_spill_width(
                sheet,
                ctx,
                col_idx,
                row_idx,
                &runs,
                paragraph_alignment,
                col_span,
                umya_cell,
            );

            let content = if runs.is_empty() {
                Vec::new()
            } else {
                vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle {
                        alignment: paragraph_alignment,
                        ..ParagraphStyle::default()
                    },
                    runs,
                })]
            };

            cells.push(TableCell {
                content,
                col_span,
                row_span,
                border,
                background,
                data_bar,
                icon_text,
                icon_color,
                spill_width,
                vertical_align: cell_vertical_align,
                padding: None,
            });
        }

        let row_has_wrapping_cell = || {
            (ctx.col_start..=ctx.col_end).any(|col| {
                sheet
                    .get_cell((col, row_idx))
                    .and_then(|cell| cell.get_style().get_alignment().cloned())
                    .map(|alignment| *alignment.get_wrap_text())
                    .unwrap_or(false)
            })
        };
        let height: Option<f64> = printed_row_height(sheet, row_idx, &row_has_wrapping_cell);

        rows.push(TableRow { cells, height });
    }
    rows
}

/// Prepare the shared context for processing a sheet (dimensions, merges, styles, etc.).
/// Returns (SheetContext, row_start, row_end) or None if the sheet is empty.
pub(super) fn prepare_sheet_context(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font_mdw: Option<f64>,
    raw_cond_fmt_hints: Option<&super::cond_fmt_raw::RawCondFmtHints>,
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

    let max_digit_width_px = normal_font_mdw.unwrap_or_else(|| sheet_max_digit_width_px(sheet));
    let column_widths: Vec<f64> = (col_start..=col_end)
        .map(|col| {
            sheet
                .get_column_dimension_by_number(&col)
                .map(|c| column_width_to_pt(*c.get_width(), max_digit_width_px))
                .unwrap_or_else(|| column_width_to_pt(DEFAULT_COLUMN_WIDTH, max_digit_width_px))
        })
        .collect();

    let (merge_tops, merge_skips) = build_merge_maps(sheet);
    let cond_fmt_overrides = build_cond_fmt_overrides(sheet, raw_cond_fmt_hints);
    let num_cols = (col_end - col_start + 1) as usize;

    Some((
        SheetContext {
            col_start,
            col_end,
            num_cols,
            column_widths,
            max_digit_width_px,
            merge_tops,
            merge_skips,
            cond_fmt_overrides,
        },
        row_start,
        row_end,
    ))
}
