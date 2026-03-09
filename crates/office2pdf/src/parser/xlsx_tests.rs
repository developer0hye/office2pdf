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
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
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

#[test]
fn test_parse_error_includes_library_name() {
    let parser = XlsxParser;
    let result = parser.parse(b"not an xlsx file", &ConvertOptions::default());
    let err = result.unwrap_err();
    let msg = err.to_string();
    assert!(
        msg.contains("umya-spreadsheet"),
        "Parse error should include upstream library name 'umya-spreadsheet', got: {msg}"
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
fn build_xlsx_with_merges(sheet_name: &str, cells: &[(&str, &str)], merges: &[&str]) -> Vec<u8> {
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
fn test_cell_border_styles() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Styled borders");
        let borders = cell.get_style_mut().get_borders_mut();
        // Dashed top
        borders
            .get_top_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DASHED);
        borders.get_top_mut().get_color_mut().set_argb("FF000000");
        // Dotted bottom
        borders
            .get_bottom_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DOTTED);
        borders
            .get_bottom_mut()
            .get_color_mut()
            .set_argb("FF000000");
        // DashDot left
        borders
            .get_left_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DASHDOT);
        borders.get_left_mut().get_color_mut().set_argb("FF000000");
        // Double right
        borders
            .get_right_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DOUBLE);
        borders.get_right_mut().get_color_mut().set_argb("FF000000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_table_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");

    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Dashed, "Top should be dashed");

    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert_eq!(
        bottom.style,
        BorderLineStyle::Dotted,
        "Bottom should be dotted"
    );

    let left = border.left.as_ref().expect("Expected left border");
    assert_eq!(
        left.style,
        BorderLineStyle::DashDot,
        "Left should be dashDot"
    );

    let right = border.right.as_ref().expect("Expected right border");
    assert_eq!(
        right.style,
        BorderLineStyle::Double,
        "Right should be double"
    );
}

#[test]
fn test_cell_border_medium_dashed() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("MedDash");
        let borders = cell.get_style_mut().get_borders_mut();
        borders
            .get_top_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_MEDIUMDASHED);
        borders.get_top_mut().get_color_mut().set_argb("FF000000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_table_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");
    let top = border.top.as_ref().expect("Expected top border");
    // mediumDashed maps to Dashed style with medium width (1.0pt)
    assert_eq!(top.style, BorderLineStyle::Dashed);
    assert!((top.width - 1.0).abs() < 0.01);
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
    let data = build_xlsx_multi_sheet(&[("Sheet1", &[("A1", "A")]), ("Sheet2", &[("A1", "B")])]);
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
    let data = build_xlsx_multi_sheet(&[("Sheet1", &[("A1", "A")]), ("Sheet2", &[("A1", "B")])]);
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

#[path = "xlsx_condfmt_tests.rs"]
mod condfmt_tests;

#[path = "xlsx_chart_tests.rs"]
mod chart_tests;

// ── Metadata extraction tests ──────────────────────────────────────

#[test]
fn test_parse_xlsx_extracts_metadata() {
    let mut book = umya_spreadsheet::new_file();
    {
        let props = book.get_properties_mut();
        props.set_title("My XLSX Title");
        props.set_creator("XLSX Author");
        props.set_subject("XLSX Subject");
        props.set_description("XLSX description text");
        props.set_created("2024-01-10T07:00:00Z");
        props.set_modified("2024-02-20T15:45:00Z");
    }
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sheet1");
        sheet.get_cell_mut("A1").set_value("Hello");
    }

    let mut buf = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut buf).unwrap();
    let data = buf.into_inner();

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.metadata.title.as_deref(), Some("My XLSX Title"));
    assert_eq!(doc.metadata.author.as_deref(), Some("XLSX Author"));
    assert_eq!(doc.metadata.subject.as_deref(), Some("XLSX Subject"));
    assert_eq!(
        doc.metadata.description.as_deref(),
        Some("XLSX description text")
    );
    assert_eq!(
        doc.metadata.created.as_deref(),
        Some("2024-01-10T07:00:00Z")
    );
    assert_eq!(
        doc.metadata.modified.as_deref(),
        Some("2024-02-20T15:45:00Z")
    );
}

#[test]
fn test_parse_xlsx_without_metadata_no_crash() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "test")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    // Should not crash; default metadata has no values
    // (umya-spreadsheet defaults may have empty strings)
    let _ = doc.metadata;
}

#[path = "xlsx_streaming_tests.rs"]
mod streaming_tests;
