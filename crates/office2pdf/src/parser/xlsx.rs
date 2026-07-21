use std::io::Cursor;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Chart, Document, ImageData, Margins, Metadata, Page, PageSize, SheetPage, StyleSheet, Table,
    TableRow,
};
use crate::parser::Parser;

#[path = "xlsx_cond_fmt_raw.rs"]
pub(crate) mod cond_fmt_raw;
#[path = "xlsx_cells.rs"]
mod xlsx_cells;
#[path = "xlsx_drawing.rs"]
mod xlsx_drawing;
#[path = "xlsx_hf.rs"]
mod xlsx_hf;
#[path = "xlsx_pagination.rs"]
mod xlsx_pagination;
#[path = "xlsx_style.rs"]
mod xlsx_style;

use self::xlsx_cells::*;
use self::xlsx_drawing::*;
use self::xlsx_hf::*;

// Re-export cell address types for cond_fmt module.
pub(crate) use self::xlsx_cells::{CellPos, CellRange, parse_cell_ref};

/// Parser for XLSX (Office Open XML Excel) spreadsheets.
/// Print margins for a sheet: the worksheet's explicit `<pageMargins>` when
/// present, otherwise Excel's defaults (0.7" left/right, 0.75" top/bottom).
/// umya leaves absent margin attributes at 0.0, which is not a value Excel
/// ever writes, so ≤0 means "not specified".
fn sheet_print_margins(sheet: &umya_spreadsheet::Worksheet) -> Margins {
    let page_margins = sheet.get_page_margins();
    let inches_to_pt = |inches: f64, default_pt: f64| -> f64 {
        if inches > 0.0 {
            inches * 72.0
        } else {
            default_pt
        }
    };
    Margins {
        top: inches_to_pt(*page_margins.get_top(), 54.0),
        bottom: inches_to_pt(*page_margins.get_bottom(), 54.0),
        left: inches_to_pt(*page_margins.get_left(), 50.4),
        right: inches_to_pt(*page_margins.get_right(), 50.4),
    }
}

/// Map an OOXML worksheet paper-size code to portrait dimensions in points.
/// Unknown or omitted codes keep the renderer's A4 default.
fn worksheet_paper_size(code: u32) -> PageSize {
    let (width, height) = match code {
        1 | 2 => (612.0, 792.0),    // Letter / Letter Small
        3 => (792.0, 1224.0),       // Tabloid
        4 => (1224.0, 792.0),       // Ledger
        5 => (612.0, 1008.0),       // Legal
        6 => (396.0, 612.0),        // Statement
        7 => (522.0, 756.0),        // Executive
        8 => (841.89, 1190.55),     // A3
        9 | 10 => (595.28, 841.89), // A4 / A4 Small
        11 => (419.53, 595.28),     // A5
        12 => (728.50, 1031.81),    // B4 (JIS)
        13 => (515.91, 728.50),     // B5 (JIS)
        _ => return PageSize::default(),
    };
    PageSize { width, height }
}

/// Preserve a worksheet's paper size and landscape orientation in the IR.
fn sheet_page_size(sheet: &umya_spreadsheet::Worksheet) -> PageSize {
    let page_setup = sheet.get_page_setup();
    let size = worksheet_paper_size(*page_setup.get_paper_size());
    if matches!(
        page_setup.get_orientation(),
        umya_spreadsheet::structs::OrientationValues::Landscape
    ) {
        PageSize {
            width: size.height,
            height: size.width,
        }
    } else {
        size
    }
}

/// Convert absolute print-title columns to 0-based indices within the
/// rendered column range, half-open. None when the titles fall outside it.
fn title_column_indices(print_titles: PrintTitles, ctx: &SheetContext) -> Option<(usize, usize)> {
    let (col_start, col_end) = print_titles.cols?;
    if col_end < ctx.col_start || col_start > ctx.col_end {
        return None;
    }
    let start_idx = col_start.max(ctx.col_start) - ctx.col_start;
    let end_idx = col_end.min(ctx.col_end) - ctx.col_start + 1;
    Some((start_idx as usize, end_idx as usize))
}

/// Convert a raw drawing anchor into a render-ready image: 1-indexed anchor
/// row plus a size in points resolved against the sheet's column widths and
/// row heights (twoCellAnchor) or the declared extent (oneCellAnchor).
fn anchored_image(
    anchor: xlsx_drawing::RawImageAnchor,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
) -> crate::ir::SheetImage {
    const EMU_PER_PT: f64 = 12_700.0;

    let column_width_at = |col_zero_based: u32| -> f64 {
        let col: u32 = col_zero_based + 1;
        if col >= ctx.col_start && col <= ctx.col_end {
            ctx.column_widths
                .get((col - ctx.col_start) as usize)
                .copied()
                .unwrap_or(0.0)
        } else {
            column_width_to_pt(DEFAULT_COLUMN_WIDTH, ctx.max_digit_width_px)
        }
    };
    let row_height_at = |row_zero_based: u32| -> f64 {
        sheet
            .get_row_dimension(&(row_zero_based + 1))
            .map(|row| *row.get_height())
            .filter(|height| *height > 0.0)
            .unwrap_or(15.0)
    };

    let (width, height): (f64, f64) =
        if let Some((to_col, to_col_off, to_row, to_row_off)) = anchor.to {
            let width: f64 = (anchor.from_col..to_col).map(column_width_at).sum::<f64>()
                - anchor.from_col_off_emu as f64 / EMU_PER_PT
                + to_col_off as f64 / EMU_PER_PT;
            let height: f64 = (anchor.from_row..to_row).map(row_height_at).sum::<f64>()
                - anchor.from_row_off_emu as f64 / EMU_PER_PT
                + to_row_off as f64 / EMU_PER_PT;
            (width.max(1.0), height.max(1.0))
        } else if let Some((cx, cy)) = anchor.ext_emu {
            (
                (cx as f64 / EMU_PER_PT).max(1.0),
                (cy as f64 / EMU_PER_PT).max(1.0),
            )
        } else {
            (100.0, 100.0)
        };

    let x_offset_pt: f64 = (0..anchor.from_col).map(column_width_at).sum::<f64>()
        + anchor.from_col_off_emu as f64 / EMU_PER_PT;

    let image = ImageData {
        data: anchor.data,
        format: anchor.format,
        width: Some(width),
        height: Some(height),
        crop: None,
        stroke: None,
        alignment: None,
        clip_shape: None,
        shadow: None,
    };
    crate::ir::SheetImage {
        anchor_row: anchor.from_row + 1,
        x_offset_pt,
        image,
    }
}

/// Context stand-in for sheets with no used cells, so drawing anchors can
/// still resolve against default column widths and row heights.
fn empty_sheet_context() -> SheetContext {
    SheetContext {
        col_start: 1,
        col_end: 0,
        num_cols: 0,
        column_widths: Vec::new(),
        max_digit_width_px: 7.0,
        merge_tops: std::collections::HashMap::new(),
        merge_skips: std::collections::HashSet::new(),
        cond_fmt_overrides: std::collections::HashMap::new(),
    }
}

/// Convert a raw text-box anchor into a render-ready box, sized like images.
fn anchored_text_box(
    anchor: xlsx_drawing::RawTextBoxAnchor,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
) -> crate::ir::SheetTextBox {
    let placed = anchored_image(
        xlsx_drawing::RawImageAnchor {
            from_row: anchor.geometry.from_row,
            from_col: anchor.geometry.from_col,
            from_col_off_emu: anchor.geometry.from_col_off_emu,
            from_row_off_emu: anchor.geometry.from_row_off_emu,
            to: anchor.geometry.to,
            ext_emu: anchor.geometry.ext_emu,
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
        },
        sheet,
        ctx,
    );
    crate::ir::SheetTextBox {
        anchor_row: placed.anchor_row,
        x_offset_pt: placed.x_offset_pt,
        width: placed.image.width.unwrap_or(100.0),
        height: placed.image.height.unwrap_or(50.0),
        paragraphs: anchor.paragraphs,
        fill: anchor.fill,
        border: anchor.border,
        vertical_center: anchor.vertical_center,
    }
}

pub struct XlsxParser;

impl XlsxParser {
    /// Parse XLSX in streaming mode, returning one `Document` per chunk of rows.
    ///
    /// Each chunk contains a single `SheetPage` with at most `chunk_size` rows.
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
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        let metadata = extract_xlsx_metadata(&book);
        let cond_fmt_hints = cond_fmt_raw::extract_cond_fmt_hints(data);
        // Excel derives every column print metric from the workbook Normal
        // font; cell fonts do not participate (issue #366).
        let normal_font_mdw: Option<f64> = extract_normal_font(data)
            .map(|(family, size)| max_digit_width_px_for_normal_font(&family, size));

        let mut chart_map = extract_charts_with_anchors(data);
        let mut image_map = extract_images_with_anchors(data);
        let mut text_box_map = extract_text_boxes_with_anchors(data);

        let mut chunks = Vec::new();
        let mut warnings = Vec::new();

        for sheet in book.get_sheet_collection() {
            // Filter by sheet name if specified
            if let Some(ref names) = options.sheet_names
                && !names.iter().any(|n| n == sheet.get_name())
            {
                continue;
            }

            let Some((ctx, row_start, row_end)) =
                prepare_sheet_context(sheet, normal_font_mdw, cond_fmt_hints.get(sheet.get_name()))
            else {
                // A sheet without used cells can still carry drawings; give
                // its images a page instead of dropping them.
                let sheet_name = sheet.get_name().to_string();
                let raw_images = image_map.remove(&sheet_name);
                let raw_text_boxes = text_box_map.remove(&sheet_name);
                let raw_charts = chart_map.remove(&sheet_name);
                if raw_images.is_some() || raw_text_boxes.is_some() || raw_charts.is_some() {
                    let stub_ctx = empty_sheet_context();
                    let images: Vec<crate::ir::SheetImage> = raw_images
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_image(anchor, sheet, &stub_ctx))
                        .collect();
                    let text_boxes: Vec<crate::ir::SheetTextBox> = raw_text_boxes
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_text_box(anchor, sheet, &stub_ctx))
                        .collect();
                    let charts: Vec<(u32, Chart)> = raw_charts.unwrap_or_default();
                    if !images.is_empty() || !text_boxes.is_empty() || !charts.is_empty() {
                        chunks.push(Document {
                            metadata: metadata.clone(),
                            pages: vec![Page::Sheet(SheetPage {
                                name: sheet_name,
                                size: sheet_page_size(sheet),
                                margins: sheet_print_margins(sheet),
                                table: Table::default(),
                                header: None,
                                footer: None,
                                charts,
                                images,
                                text_boxes,
                            })],
                            styles: StyleSheet::default(),
                        });
                    }
                }
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
            let mut sheet_images: Vec<crate::ir::SheetImage> = image_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_image(anchor, sheet, &ctx))
                .collect();
            sheet_images.sort_by_key(|sheet_image| sheet_image.anchor_row);
            let mut sheet_text_boxes: Vec<crate::ir::SheetTextBox> = text_box_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_text_box(anchor, sheet, &ctx))
                .collect();
            sheet_text_boxes.sort_by_key(|text_box| text_box.anchor_row);

            let print_titles = find_print_titles(&book, sheet);
            let title_columns: Option<(usize, usize)> = title_column_indices(print_titles, &ctx);

            // Process rows in chunks
            let mut chunk_start = row_start;
            let mut first_chunk = true;
            while chunk_start <= row_end {
                let chunk_end = (chunk_start + chunk_size as u32 - 1).min(row_end);

                let mut rows = build_rows_for_range(sheet, &ctx, chunk_start, chunk_end);
                let mut header_row_count: usize = 0;
                if let Some((title_start, title_end)) = print_titles.rows
                    && title_end < chunk_start
                {
                    // Later chunks don't contain the title rows — prepend them.
                    let mut title_rows = build_rows_for_range(sheet, &ctx, title_start, title_end);
                    header_row_count = title_rows.len();
                    title_rows.append(&mut rows);
                    rows = title_rows;
                } else if let Some((_, title_end)) = print_titles.rows
                    && title_end >= chunk_start
                    && title_end <= chunk_end
                {
                    header_row_count = (title_end - chunk_start + 1) as usize;
                }

                let doc = Document {
                    metadata: metadata.clone(),
                    pages: xlsx_pagination::split_sheet_page_by_width(
                        SheetPage {
                            name: sheet_name.clone(),
                            size: sheet_page_size(sheet),
                            margins: sheet_print_margins(sheet),
                            table: Table {
                                rows,
                                column_widths: ctx.column_widths.clone(),
                                header_row_count,
                                alignment: None,
                                default_cell_padding: None,
                                use_content_driven_row_heights: false,
                                default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
                            },
                            header: sheet_header.clone(),
                            footer: sheet_footer.clone(),
                            charts: if first_chunk {
                                std::mem::take(&mut sheet_charts)
                            } else {
                                vec![]
                            },
                            images: if first_chunk {
                                std::mem::take(&mut sheet_images)
                            } else {
                                vec![]
                            },
                            text_boxes: if first_chunk {
                                first_chunk = false;
                                std::mem::take(&mut sheet_text_boxes)
                            } else {
                                vec![]
                            },
                        },
                        title_columns,
                    )
                    .into_iter()
                    .map(Page::Sheet)
                    .collect(),
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
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        // Extract metadata from umya-spreadsheet properties
        let metadata = extract_xlsx_metadata(&book);
        let cond_fmt_hints = cond_fmt_raw::extract_cond_fmt_hints(data);
        // Excel derives every column print metric from the workbook Normal
        // font; cell fonts do not participate (issue #366).
        let normal_font_mdw: Option<f64> = extract_normal_font(data)
            .map(|(family, size)| max_digit_width_px_for_normal_font(&family, size));

        // Extract charts with anchor positions per sheet
        let mut chart_map = extract_charts_with_anchors(data);
        let mut image_map = extract_images_with_anchors(data);
        let mut text_box_map = extract_text_boxes_with_anchors(data);

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

            let Some((ctx, row_start, row_end)) =
                prepare_sheet_context(sheet, normal_font_mdw, cond_fmt_hints.get(sheet.get_name()))
            else {
                // A sheet without used cells can still carry drawings; give
                // its images a page instead of dropping them.
                let sheet_name = sheet.get_name().to_string();
                let raw_images = image_map.remove(&sheet_name);
                let raw_text_boxes = text_box_map.remove(&sheet_name);
                let raw_charts = chart_map.remove(&sheet_name);
                if raw_images.is_some() || raw_text_boxes.is_some() || raw_charts.is_some() {
                    let stub_ctx = empty_sheet_context();
                    let images: Vec<crate::ir::SheetImage> = raw_images
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_image(anchor, sheet, &stub_ctx))
                        .collect();
                    let text_boxes: Vec<crate::ir::SheetTextBox> = raw_text_boxes
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_text_box(anchor, sheet, &stub_ctx))
                        .collect();
                    let charts: Vec<(u32, Chart)> = raw_charts.unwrap_or_default();
                    if !images.is_empty() || !text_boxes.is_empty() || !charts.is_empty() {
                        pages.push(Page::Sheet(SheetPage {
                            name: sheet_name,
                            size: sheet_page_size(sheet),
                            margins: sheet_print_margins(sheet),
                            table: Table::default(),
                            header: None,
                            footer: None,
                            charts,
                            images,
                            text_boxes,
                        }));
                    }
                }
                continue;
            };

            let rows = build_rows_for_range(sheet, &ctx, row_start, row_end);

            let print_titles = find_print_titles(&book, sheet);
            let title_columns: Option<(usize, usize)> = title_column_indices(print_titles, &ctx);
            // Rows from the sheet top through the end of the title range
            // repeat as the table header on every page. Excel repeats only
            // the title rows themselves; when they don't start at the top
            // this over-repeats the few rows above them, which reads better
            // than not repeating at all.
            let header_row_count: usize = print_titles
                .rows
                .filter(|(_, title_end)| *title_end >= row_start)
                .map(|(_, title_end)| (title_end.min(row_end) - row_start + 1) as usize)
                .unwrap_or(0);

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
            let mut sheet_images: Vec<crate::ir::SheetImage> = image_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_image(anchor, sheet, &ctx))
                .collect();
            sheet_images.sort_by_key(|sheet_image| sheet_image.anchor_row);
            let mut sheet_text_boxes: Vec<crate::ir::SheetTextBox> = text_box_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_text_box(anchor, sheet, &ctx))
                .collect();
            sheet_text_boxes.sort_by_key(|text_box| text_box.anchor_row);

            if row_breaks.is_empty() {
                // No page breaks — single page
                pages.extend(
                    xlsx_pagination::split_sheet_page_by_width(
                        SheetPage {
                            name: sheet_name,
                            size: sheet_page_size(sheet),
                            margins: sheet_print_margins(sheet),
                            table: Table {
                                rows,
                                column_widths: ctx.column_widths,
                                header_row_count,
                                alignment: None,
                                default_cell_padding: None,
                                use_content_driven_row_heights: false,
                                default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
                            },
                            header: sheet_header.clone(),
                            footer: sheet_footer.clone(),
                            charts: sheet_charts,
                            images: sheet_images,
                            text_boxes: sheet_text_boxes,
                        },
                        title_columns,
                    )
                    .into_iter()
                    .map(Page::Sheet),
                );
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
                for mut segment in segments {
                    let mut segment_header_rows: usize = 0;
                    if first_segment {
                        segment_header_rows = header_row_count.min(segment.len());
                    } else if let Some((title_start, title_end)) = print_titles.rows
                        && title_end >= row_start
                    {
                        // Later segments don't contain the title rows — prepend.
                        let mut title_rows = build_rows_for_range(
                            sheet,
                            &ctx,
                            title_start.max(row_start),
                            title_end,
                        );
                        segment_header_rows = title_rows.len();
                        title_rows.append(&mut segment);
                        segment = title_rows;
                    }
                    pages.extend(
                        xlsx_pagination::split_sheet_page_by_width(
                            SheetPage {
                                name: sheet_name.clone(),
                                size: sheet_page_size(sheet),
                                margins: sheet_print_margins(sheet),
                                table: Table {
                                    rows: segment,
                                    column_widths: ctx.column_widths.clone(),
                                    header_row_count: segment_header_rows,
                                    alignment: None,
                                    default_cell_padding: None,
                                    use_content_driven_row_heights: false,
                                    default_vertical_align: Some(
                                        crate::ir::CellVerticalAlign::Bottom,
                                    ),
                                },
                                header: sheet_header.clone(),
                                footer: sheet_footer.clone(),
                                charts: if first_segment {
                                    std::mem::take(&mut sheet_charts)
                                } else {
                                    vec![]
                                },
                                images: if first_segment {
                                    std::mem::take(&mut sheet_images)
                                } else {
                                    vec![]
                                },
                                text_boxes: if first_segment {
                                    first_segment = false;
                                    std::mem::take(&mut sheet_text_boxes)
                                } else {
                                    vec![]
                                },
                            },
                            title_columns,
                        )
                        .into_iter()
                        .map(Page::Sheet),
                    );
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
