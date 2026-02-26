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

impl Parser for XlsxParser {
    fn parse(&self, data: &[u8]) -> Result<Document, ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse XLSX: {e}")))?;

        let mut pages = Vec::new();

        for sheet in book.get_sheet_collection() {
            let (max_col, max_row) = sheet.get_highest_column_and_row();
            if max_col == 0 || max_row == 0 {
                continue; // skip empty sheets
            }

            let column_widths: Vec<f64> = (1..=max_col)
                .map(|col| {
                    sheet
                        .get_column_dimension_by_number(&col)
                        .map(|c| column_width_to_pt(*c.get_width()))
                        .unwrap_or_else(|| column_width_to_pt(DEFAULT_COLUMN_WIDTH))
                })
                .collect();

            let mut rows = Vec::new();
            for row_idx in 1..=max_row {
                let mut cells = Vec::new();
                for col_idx in 1..=max_col {
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

                    cells.push(TableCell {
                        content,
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
}
