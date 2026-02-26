use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::error::ConvertError;
use crate::ir::{
    Block, Document, Margins, Metadata, Page, PageSize, Paragraph, ParagraphStyle, Run, StyleSheet,
    Table, TableCell, TablePage, TableRow, TextStyle,
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
    fn parse(&self, data: &[u8]) -> Result<Document, ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse XLSX: {e}")))?;

        let mut pages = Vec::new();

        for sheet in book.get_sheet_collection() {
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

            let column_widths: Vec<f64> = (1..=max_col)
                .map(|col| {
                    sheet
                        .get_column_dimension_by_number(&col)
                        .map(|c| column_width_to_pt(*c.get_width()))
                        .unwrap_or_else(|| column_width_to_pt(DEFAULT_COLUMN_WIDTH))
                })
                .collect();

            let (merge_tops, merge_skips) = build_merge_maps(sheet);

            let mut rows = Vec::new();
            for row_idx in 1..=max_row {
                let mut cells = Vec::new();
                for col_idx in 1..=max_col {
                    // Skip cells that are part of a merge but not the top-left
                    if merge_skips.contains(&(col_idx, row_idx)) {
                        continue;
                    }

                    // umya-spreadsheet tuple is (column, row), both 1-indexed
                    let value = sheet
                        .get_cell((col_idx, row_idx))
                        .map(|cell| cell.get_value().to_string())
                        .unwrap_or_default();

                    let content = if value.is_empty() {
                        Vec::new()
                    } else {
                        vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: value,
                                style: TextStyle::default(),
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
                        ..TableCell::default()
                    });
                }
                rows.push(TableRow {
                    cells,
                    height: None,
                });
            }

            pages.push(Page::Table(TablePage {
                name: sheet.get_name().to_string(),
                size: PageSize::default(),
                margins: Margins::default(),
                table: Table {
                    rows,
                    column_widths,
                },
            }));
        }

        Ok(Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        })
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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "42");
        assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "3.14");
    }

    #[test]
    fn test_parse_dates_as_text() {
        let data = build_xlsx_bytes("Dates", &[("A1", "2024-01-15"), ("A2", "December 25")]);
        let parser = XlsxParser;
        let doc = parser.parse(&data).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "2024-01-15");
        assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "December 25");
    }

    // ----- Sheet name tests -----

    #[test]
    fn test_sheet_name_preserved() {
        let data = build_xlsx_bytes("Financial Report", &[("A1", "Revenue")]);
        let parser = XlsxParser;
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let result = parser.parse(b"not an xlsx file");
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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

        let tp = get_table_page(&doc, 0);
        assert_eq!(tp.table.rows[0].cells.len(), 1);
        assert_eq!(tp.table.rows[0].cells[0].col_span, 4);
        assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Title");
    }
}
