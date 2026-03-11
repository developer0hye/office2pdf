use std::io::Cursor;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Document, Margins, Metadata, Page, PageSize, StyleSheet, Table, TablePage, TableRow,
};
use crate::parser::Parser;

#[path = "xlsx_cells.rs"]
mod xlsx_cells;
#[path = "xlsx_drawing.rs"]
mod xlsx_drawing;
#[path = "xlsx_hf.rs"]
mod xlsx_hf;
#[path = "xlsx_style.rs"]
mod xlsx_style;

use self::xlsx_cells::*;
use self::xlsx_drawing::*;
use self::xlsx_hf::*;

// Re-export cell address types for cond_fmt module.
pub(crate) use self::xlsx_cells::{CellPos, CellRange, parse_cell_ref};

/// Parser for XLSX (Office Open XML Excel) spreadsheets.
pub struct XlsxParser;

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
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
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
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
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
