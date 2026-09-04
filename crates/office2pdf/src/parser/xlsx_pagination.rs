//! Column-wise pagination for sheets wider than the printable page.
//!
//! Excel prints columns that overflow the page width on subsequent pages
//! (default order: down, then over). office2pdf previously clipped them at
//! the right page edge, silently losing content.
//!
//! A drawing-only sheet has no columns to split on, so its page-columns come
//! from the drawings' extents instead ([`split_drawing_only_page`],
//! issue #713).

use crate::ir::{Block, HFInline, HeaderFooter, SheetPage, Table, TableCell, TableRow};

/// Upper bound on overflow pages per sheet chunk. Pathological sheets (used
/// ranges thousands of columns wide) would otherwise explode into thousands
/// of pages and blow the Typst compiler's stack; columns beyond the cap stay
/// on the last page (clipped, the pre-pagination behavior).
const MAX_COLUMN_GROUPS: usize = 12;

/// What one sheet's `<pageSetUpPr fitToPage="1"/>` asks pagination to scale it
/// onto. Both directions are bounded separately and Excel obeys the tighter of
/// the two.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub(super) struct SheetFit {
    /// `fitToWidth` when it binds; `None` leaves the column direction free.
    pub(super) pages_wide: Option<u32>,
    /// `fitToHeight` when it binds; `None` leaves the row direction free.
    pub(super) pages_tall: Option<u32>,
    /// The printed grid height of the *whole* sheet in points, which is what
    /// the row bound is measured against. The page handed to pagination may
    /// carry one streaming chunk or one explicit-break segment of that sheet,
    /// so its own rows cannot supply the total. Unread when `pages_tall` is
    /// `None`.
    pub(super) sheet_height_pt: f64,
}

/// Split a sheet page into column groups that each fit the printable width.
/// Returns the page unchanged when everything fits. `title_columns` is the
/// 0-based inclusive-exclusive range of print-title columns (from
/// `_xlnm.Print_Titles`) repeated at the left of every overflow page.
pub(super) fn split_sheet_page_by_width(
    page: SheetPage,
    title_columns: Option<(usize, usize)>,
    fit: SheetFit,
    header_footer_scales_with_doc: bool,
) -> Vec<SheetPage> {
    let page: SheetPage = fit_page_to_pages(page, fit, header_footer_scales_with_doc);
    let printable_width: f64 = page.size.width - page.margins.left - page.margins.right;
    let total_width: f64 = page.table.column_widths.iter().sum();
    if total_width <= printable_width || page.table.column_widths.len() <= 1 {
        return vec![page];
    }

    let title_columns: Option<(usize, usize)> = title_columns
        .map(|(start, end)| (start, end.min(page.table.column_widths.len())))
        .filter(|(start, end)| start < end);
    // Reserve the repeated title width so overflow groups still fit the
    // page. The first group holds the title columns physically (they never
    // get prepended to it), so it packs against the full printable width —
    // reserving there too underpacked page 1 by the title width (issue #623
    // adversarial review, finding 3).
    let title_width: f64 = title_columns
        .map(|(start, end)| page.table.column_widths[start..end].iter().sum())
        .unwrap_or(0.0);
    let widest_column: f64 = page.table.column_widths.iter().cloned().fold(0.0, f64::max);
    let first_group_packing_width: f64 = printable_width.max(widest_column);
    let overflow_packing_width: f64 = (printable_width - title_width).max(widest_column);

    let mut groups: Vec<(usize, usize)> = column_groups(
        &page.table.column_widths,
        first_group_packing_width,
        overflow_packing_width,
    );
    if groups.len() <= 1 {
        return vec![page];
    }
    if groups.len() > MAX_COLUMN_GROUPS {
        let column_count = page.table.column_widths.len();
        groups.truncate(MAX_COLUMN_GROUPS);
        if let Some(last) = groups.last_mut() {
            last.1 = column_count;
        }
    }

    let title_table: Option<Table> =
        title_columns.map(|(start, end)| slice_table_columns(&page.table, start, end));

    let mut result: Vec<SheetPage> = Vec::with_capacity(groups.len());
    for (index, &(start, end)) in groups.iter().enumerate() {
        let mut table: Table = slice_table_columns(&page.table, start, end);
        // Excel repeats title columns on pages that no longer show them.
        if let (Some(title_table), Some((title_start, _))) = (title_table.as_ref(), title_columns)
            && start > title_start
        {
            table = prepend_title_columns(title_table, table);
        }
        result.push(SheetPage {
            name: page.name.clone(),
            size: page.size,
            margins: page.margins,
            table,
            header: page.header.clone(),
            footer: page.footer.clone(),
            // Charts and images anchor to rows of the first column group only.
            charts: if index == 0 {
                page.charts.clone()
            } else {
                Vec::new()
            },
            images: if index == 0 {
                page.images.clone()
            } else {
                Vec::new()
            },
            text_boxes: if index == 0 {
                page.text_boxes.clone()
            } else {
                Vec::new()
            },
        });
    }
    result
}

/// Concatenate the repeated title columns before a column group's table.
/// Shrink a sheet until it fits the pages `fitToWidth` and `fitToHeight`
/// allow.
///
/// A sheet with `<pageSetUpPr fitToPage="1"/>` and `fitToWidth="1"` asks Excel
/// to scale it onto one page wide rather than to spill the overflow onto a
/// second strip. Reading neither attribute printed the repository workbook on
/// 53 pages where Excel prints 23 (issue #530).
///
/// `fitToHeight` bounds the row direction the same way, and ECMA-376 defaults
/// it to one page just as it defaults `fitToWidth`, so a sheet naming neither
/// is asking to be squeezed onto a single page both ways. Excel obeys the
/// tighter of the two bounds: the reported college-budget workbook fits A3's
/// width at 0.89 and its height at 0.78, and its native export is one page at
/// 0.78 (issue #1181).
///
/// Excel scales the whole sheet, not the columns alone, so the row heights and
/// the type scale with the widths — the audited sheet's 10pt body text prints
/// at 7.50pt, the same 0.75 the columns take.
///
/// Excel never scales *up* to fill a page, so a sheet that already fits is
/// left alone.
fn fit_page_to_pages(
    page: SheetPage,
    fit: SheetFit,
    header_footer_scales_with_doc: bool,
) -> SheetPage {
    let printable_width: f64 = page.size.width - page.margins.left - page.margins.right;
    let total_width: f64 = page.table.column_widths.iter().sum();
    let printable_height: f64 = page.size.height - page.margins.top - page.margins.bottom;
    let Some(scale) = [
        fit_scale(fit.pages_wide, printable_width, total_width),
        fit_scale(fit.pages_tall, printable_height, fit.sheet_height_pt),
    ]
    .into_iter()
    .flatten()
    .reduce(f64::min) else {
        return page;
    };
    if scale >= 1.0 {
        return page;
    }
    scale_sheet_page(page, scale, header_footer_scales_with_doc)
}

/// The scale that fits `total_pt` of sheet into `pages` pages of `printable_pt`,
/// or `None` when that direction is unconstrained — either unbounded by the
/// file (a declared zero is Excel's "as many pages as it takes") or unmeasurable.
///
/// Excel's auto-fit scale is a whole percent, truncated rather than rounded so
/// the content is guaranteed to fit. Keeping the raw ratio leaves every derived
/// type size a fraction of a point off the printed sheet — the audited sheet
/// came out at 7.55pt against Excel's 7.50pt.
fn fit_scale(pages: Option<u32>, printable_pt: f64, total_pt: f64) -> Option<f64> {
    let pages: u32 = pages.filter(|pages| *pages > 0)?;
    if printable_pt <= 0.0 || total_pt <= 0.0 {
        return None;
    }
    let exact_scale: f64 = (printable_pt * f64::from(pages)) / total_pt;
    Some((exact_scale * 100.0).floor() / 100.0)
}

/// Multiply a sheet's widths, heights, type sizes, cell padding, and anchored
/// drawings by `scale`.
///
/// Padding has to scale with the rest: it is a fixed per-row overhead, so
/// leaving it at full size while the rows shrink costs a constant slice of
/// every row and accumulates into whole extra pages over a long sheet.
///
/// The header and footer scale too, unless the sheet opts out.
/// `headerFooter/@scaleWithDoc` defaults to 1 (ECMA-376 §18.3.1.46), so Excel
/// shrinks them with the sheet; leaving them at full size printed the Gantt
/// template's 8pt `&8` run beside 5.85pt body text (issue #940).
fn scale_sheet_page(
    mut page: SheetPage,
    scale: f64,
    header_footer_scales_with_doc: bool,
) -> SheetPage {
    if header_footer_scales_with_doc {
        for header_footer in [page.header.as_mut(), page.footer.as_mut()]
            .into_iter()
            .flatten()
        {
            scale_header_footer_with_sheet(header_footer, scale);
        }
    }
    // Every size below is multiplied by the scale outright. The factor itself
    // rides on the table because a rule Excel evaluates at the declared size
    // and scales afterwards cannot be recovered from the products — the
    // wrapped-line advance of issue #1163 is one.
    page.table.print_scale = Some(page.table.print_scale.unwrap_or(1.0) * scale);
    for width in &mut page.table.column_widths {
        *width *= scale;
    }
    // An anchored chart is measured against the sheet's own columns and rows,
    // so the scale that shrinks those has to shrink the chart with them —
    // otherwise the fitted grid slides out from under a full-size chart. On
    // the reported workbook the 0.82 scale left the chart 183pt wider than
    // the band it is anchored to, spilling past the printable edge (#982).
    //
    // The scale rides on the placement rather than being folded into its frame
    // because Excel shrinks the whole drawing, not just the box around it: the
    // chart's tick labels, category labels and legend scale with it. Shrinking
    // the frame alone printed them at the size the chart XML declares, about
    // 22% larger than the native export's (#1069).
    for placement in page
        .charts
        .iter_mut()
        .filter_map(|chart| chart.placement.as_mut())
    {
        placement.x_offset_pt *= scale;
        placement.y_offset_pt *= scale;
        placement.print_scale *= scale;
    }
    // A picture is anchored to the same columns and rows, so it shrinks with
    // them too. Scaling the grid alone printed the reported workbook's photo
    // at 234.95 x 171.05pt against the native export's 192.66 x 140.26 —
    // 1/0.82 in each axis, the print scale never reaching it — and 84.83pt
    // further down the page (#1111).
    //
    // The scale goes into the frame rather than riding beside it as the
    // chart's does: a picture carries no text of its own, so there is nothing
    // in it that a plain resize would leave at full size.
    for image in &mut page.images {
        image.x_offset_pt *= scale;
        image.y_offset_pt *= scale;
        if let Some(width) = image.image.width.as_mut() {
            *width *= scale;
        }
        if let Some(height) = image.image.height.as_mut() {
            *height *= scale;
        }
    }
    for row in &mut page.table.rows {
        if let Some(height) = row.height.as_mut() {
            *height *= scale;
        }
        for cell in &mut row.cells {
            if let Some(padding) = cell.padding.as_mut() {
                padding.top *= scale;
                padding.right *= scale;
                padding.bottom *= scale;
                padding.left *= scale;
            }
            for block in &mut cell.content {
                scale_block_font_sizes(block, scale);
            }
        }
    }
    page
}

/// Scale every run of a header or footer and retain the same factor for its
/// horizontal sheet-coordinate box.
///
/// A run that states no size takes the renderer's default rather than being
/// left alone: it is the size the run actually prints at, and skipping it left
/// the Gantt template's leading `_x000D_` at 11pt while everything around it
/// shrank (issue #940).
///
/// Excel applies the fit factor to the story's coordinates too. Retaining it
/// separately lets the renderer reconstruct the scaled horizontal seat without
/// scaling these already-adjusted font sizes a second time (issue #1510).
fn scale_header_footer_with_sheet(header_footer: &mut HeaderFooter, scale: f64) {
    header_footer.sheet_print_scale = Some(header_footer.sheet_print_scale.unwrap_or(1.0) * scale);
    for paragraph in &mut header_footer.paragraphs {
        for element in &mut paragraph.elements {
            if let HFInline::Run(run) = element {
                let size_pt: f64 = run
                    .style
                    .font_size
                    .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);
                run.style.font_size = Some(size_pt * scale);
            }
        }
    }
}

fn scale_block_font_sizes(block: &mut Block, scale: f64) {
    match block {
        Block::Paragraph(paragraph) => {
            for run in &mut paragraph.runs {
                if let Some(size) = run.style.font_size.as_mut() {
                    *size *= scale;
                }
            }
        }
        Block::Table(table) => {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    for nested in &mut cell.content {
                        scale_block_font_sizes(nested, scale);
                    }
                }
            }
        }
        _ => {}
    }
}

fn prepend_title_columns(title_table: &Table, group_table: Table) -> Table {
    let mut column_widths: Vec<f64> = title_table.column_widths.clone();
    column_widths.extend(group_table.column_widths.iter().copied());

    let rows: Vec<TableRow> = title_table
        .rows
        .iter()
        .zip(group_table.rows)
        .map(|(title_row, group_row)| {
            let mut cells: Vec<TableCell> = title_row.cells.clone();
            cells.extend(group_row.cells);
            TableRow {
                minimum_height: None,
                cells,
                height: group_row.height,
            }
        })
        .collect();

    Table {
        rows,
        column_widths,
        ..group_table
    }
}

/// Greedily pack columns left-to-right into groups whose summed width fits
/// their capacity; every group holds at least one column. The first group
/// packs against `first_group_width` (the full printable width — it shows
/// the title columns in place); later groups pack against
/// `overflow_group_width`, which reserves room for the prepended titles.
fn column_groups(
    column_widths: &[f64],
    first_group_width: f64,
    overflow_group_width: f64,
) -> Vec<(usize, usize)> {
    let mut groups: Vec<(usize, usize)> = Vec::new();
    let mut start: usize = 0;
    let mut acc: f64 = 0.0;
    for (index, width) in column_widths.iter().enumerate() {
        let capacity: f64 = if groups.is_empty() {
            first_group_width
        } else {
            overflow_group_width
        };
        if index > start && acc + width > capacity {
            groups.push((start, index));
            start = index;
            acc = 0.0;
        }
        acc += width;
    }
    groups.push((start, column_widths.len()));
    groups
}

/// Build a table containing only columns `[start, end)`, truncating cell
/// spans at the group boundary. A merged cell that starts before the group
/// keeps its geometry (background/border) but blanks its content.
///
/// That blanking is a stopgap, not a match for how a spreadsheet application
/// prints the continuation. A LibreOffice render of
/// `tests/fixtures/xlsx/merged_row_overflows_page_column.xlsx` redraws the
/// merge's line on the following page-column at a negative x so its tail lands
/// there, rather than leaving the cell empty. Reproducing that is #631; no
/// native Excel export has been measured yet, so the exact geometry is
/// corroborated rather than settled.
fn slice_table_columns(table: &Table, start: usize, end: usize) -> Table {
    let column_count: usize = table.column_widths.len();
    // Tracks rows still covered by a row-spanning cell, per column.
    let mut rowspan_remaining: Vec<usize> = vec![0; column_count];

    let mut rows: Vec<TableRow> = Vec::with_capacity(table.rows.len());
    for row in &table.rows {
        let mut column_cursor: usize = 0;
        let mut cells: Vec<TableCell> = Vec::new();

        for cell in &row.cells {
            while column_cursor < column_count && rowspan_remaining[column_cursor] > 0 {
                rowspan_remaining[column_cursor] -= 1;
                column_cursor += 1;
            }
            if column_cursor >= column_count {
                break;
            }

            let span: usize = cell.col_span.max(1) as usize;
            let cell_start: usize = column_cursor;
            let cell_end: usize = (column_cursor + span).min(column_count);

            if cell.row_span > 1 {
                for occupied in rowspan_remaining.iter_mut().take(cell_end).skip(cell_start) {
                    *occupied = (cell.row_span - 1) as usize;
                }
            }

            let overlap_start: usize = cell_start.max(start);
            let overlap_end: usize = cell_end.min(end);
            if overlap_start < overlap_end {
                let mut sliced: TableCell = cell.clone();
                sliced.col_span = (overlap_end - overlap_start) as u32;
                if cell_start < start {
                    // Continuation of a merge that began on an earlier page.
                    sliced.content = Vec::new();
                    sliced.spill_width = None;
                } else if let Some(spill) = sliced.spill_width {
                    // The spill width was measured against the whole sheet, so
                    // it can reach far past the columns this page actually
                    // carries — on a sheet wide enough to split, past the paper
                    // edge, losing the ink entirely (#631). Clamp it to what
                    // remains of the group from this cell's left edge.
                    let available: f64 = table.column_widths[overlap_start..end].iter().sum();
                    sliced.spill_width = Some(spill.min(available));
                }
                cells.push(sliced);
            }

            column_cursor = cell_end;
        }

        // Columns occupied only by rowspans still need their counters advanced.
        while column_cursor < column_count {
            if rowspan_remaining[column_cursor] > 0 {
                rowspan_remaining[column_cursor] -= 1;
            }
            column_cursor += 1;
        }

        rows.push(TableRow {
            minimum_height: None,
            cells,
            height: row.height,
        });
    }

    Table {
        rows,
        column_widths: table.column_widths[start..end].to_vec(),
        header_row_count: table.header_row_count,
        non_repeating_header_row_count: table.non_repeating_header_row_count,
        alignment: table.alignment,
        default_cell_padding: table.default_cell_padding,
        use_content_driven_row_heights: table.use_content_driven_row_heights,
        default_vertical_align: table.default_vertical_align,
        seats_bottom_aligned_text_on_descender: table.seats_bottom_aligned_text_on_descender,
        bottom_aligned_descent_floor_pt: table.bottom_aligned_descent_floor_pt,
        border_paint_model: table.border_paint_model,
        prints_gridlines: table.prints_gridlines,
        prints_headings: table.prints_headings,
        centers_between_print_margins: table.centers_between_print_margins,
        print_scale: table.print_scale,
    }
}

/// Split a drawing-only sheet into page-columns at printable-width
/// boundaries.
///
/// Excel prints a drawing that crosses the printable edge clipped there and
/// continues it on the next page-column. The empty-sheet branch previously
/// emitted one page and let the pictures overflow the right margin — the
/// tomcat of `WithDrawing.xlsx` ran 36pt past the printable edge on a single
/// page where Excel prints two (issue #713). The table is empty, so
/// [`split_sheet_page_by_width`] has no column widths to split on; the
/// drawings' extents drive the paging instead.
///
/// Every image on a split page carries [`crate::ir::SheetImage::clip_width_pt`]
/// so the renderer clips it to its page-column window; a continued copy also
/// carries a negative `x_offset_pt`. Charts and text boxes stay on the first
/// page-column, like the column splitter keeps them on its first group.
pub(super) fn split_drawing_only_page(page: SheetPage) -> Vec<SheetPage> {
    let printable_width: f64 = page.size.width - page.margins.left - page.margins.right;
    if printable_width <= 0.0 {
        return vec![page];
    }
    let right_extent: f64 = page
        .images
        .iter()
        .map(|image| image.x_offset_pt + image.image.width.unwrap_or(0.0))
        .fold(0.0, f64::max);
    if right_extent <= printable_width {
        return vec![page];
    }
    let group_count: usize =
        ((right_extent / printable_width).ceil() as usize).clamp(2, MAX_COLUMN_GROUPS);

    (0..group_count)
        .map(|group| {
            let window_left: f64 = group as f64 * printable_width;
            let mut paged: SheetPage = page.clone();
            paged.images = page
                .images
                .iter()
                .filter(|image| {
                    let width: f64 = image.image.width.unwrap_or(0.0);
                    image.x_offset_pt + width > window_left
                        && image.x_offset_pt < window_left + printable_width
                })
                .map(|image| {
                    let mut paged_image = image.clone();
                    paged_image.x_offset_pt -= window_left;
                    paged_image.clip_width_pt = Some(printable_width);
                    paged_image
                })
                .collect();
            if group > 0 {
                paged.charts = Vec::new();
                paged.text_boxes = Vec::new();
            }
            paged
        })
        .collect()
}

#[cfg(test)]
#[path = "xlsx_pagination_tests.rs"]
mod tests;
