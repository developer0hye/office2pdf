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

/// Helper: extract SheetPage from Document by index.
fn get_sheet_page(doc: &Document, idx: usize) -> &SheetPage {
    match &doc.pages[idx] {
        Page::Sheet(sp) => sp,
        _ => panic!("Expected SheetPage at index {idx}"),
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

/// Helper: extract the first run's TextStyle from a cell.
fn first_run_style(cell: &TableCell) -> &TextStyle {
    match &cell.content[0] {
        Block::Paragraph(p) => &p.runs[0].style,
        _ => panic!("Expected Paragraph"),
    }
}

// ----- Basic parsing tests -----

#[test]
fn test_parse_single_cell() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Hello")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "42");
    assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "3.14");
}

#[test]
fn test_parse_dates_as_text() {
    let data = build_xlsx_bytes("Dates", &[("A1", "2024-01-15"), ("A2", "December 25")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "2024-01-15");
    assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "December 25");
}

// ----- Sheet name tests -----

#[test]
fn test_sheet_name_preserved() {
    let data = build_xlsx_bytes("Financial Report", &[("A1", "Revenue")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
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
    let tp1 = get_sheet_page(&doc, 0);
    let tp2 = get_sheet_page(&doc, 1);
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

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.column_widths.len(), 2);
    // Print geometry uses the stored character width and the Normal font's
    // 8px MDW (Calibri 11, pixel-ceiled) without adding screen-only cell
    // padding a second time: 8.43 chars × 8px × 0.75 ≈ 50.6pt.
    for w in &tp.table.column_widths {
        assert!(
            *w > 50.0 && *w < 51.0,
            "Expected default print width around 50.6pt, got {w}"
        );
    }
}

#[test]
fn test_carlito_column_widths_match_native_print_metrics() {
    assert_eq!(column_width_to_pt(26.0, 8.0), 156.0);
    assert_eq!(column_width_to_pt(20.0, 8.0), 120.0);
    assert_eq!(column_width_to_pt(24.0, 8.0), 144.0);
}

#[test]
fn test_sheet_uses_dominant_carlito_font_for_column_metrics() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_cell_mut("A1")
        .set_value("Header")
        .get_style_mut()
        .get_font_mut()
        .set_name("Carlito");
    sheet
        .get_cell_mut("A2")
        .set_value("Body")
        .get_style_mut()
        .get_font_mut()
        .set_name("Carlito");

    assert_eq!(sheet_max_digit_width_px(sheet), 8.0);
}

#[test]
fn test_normal_font_max_digit_width_pixel_ceils_at_96_dpi() {
    // Excel pixel-ceils the Normal font's max digit width: Calibri 11 is
    // 0.5066em × 11pt × 96/72 ≈ 7.43px → 8px, which is what native Excel
    // print pagination of the audit fixtures requires (issue #366).
    assert_eq!(max_digit_width_px_for_normal_font("Calibri", 11.0), 8.0);
    assert_eq!(max_digit_width_px_for_normal_font("Carlito", 11.0), 8.0);
    assert_eq!(max_digit_width_px_for_normal_font("Arial", 10.0), 8.0);
    assert_eq!(
        max_digit_width_px_for_normal_font("Malgun Gothic", 11.0),
        8.0
    );
    // Smaller Normal fonts shrink the metric.
    assert_eq!(max_digit_width_px_for_normal_font("Calibri", 8.0), 6.0);
}

#[test]
fn test_extract_normal_font_reads_first_styles_font() {
    let mut book = umya_spreadsheet::new_file();
    book.get_sheet_mut(&0)
        .unwrap()
        .get_cell_mut("A1")
        .set_value("x");
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    let data = cursor.into_inner();

    let normal_font = extract_normal_font(&data).expect("styles.xml has a Normal font");
    assert_eq!(normal_font.family, "Calibri");
    assert_eq!(normal_font.size_pt, 11.0);
}

#[test]
fn test_column_overflow_splits_to_second_page_like_excel() {
    // Quotation-style layout: A4 portrait with 0.75in side margins leaves a
    // 487pt printable width. Columns of 5+30+16+8+14+16 = 89 chars under the
    // Calibri-11 Normal font are 534pt at Excel's 8px MDW, so the last
    // column overflows onto page 2 — exactly how Excel paginates the audit
    // fixture (issue #366).
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sheet1");
        for (index, (col, width)) in [
            ("A", 5.0),
            ("B", 30.0),
            ("C", 16.0),
            ("D", 8.0),
            ("E", 14.0),
            ("F", 16.0),
        ]
        .iter()
        .enumerate()
        {
            sheet.get_column_dimension_mut(col).set_width(*width);
            let cell_ref = format!("{}1", col);
            sheet
                .get_cell_mut(cell_ref.as_str())
                .set_value(format!("Col {}", index + 1));
        }
        let margins = sheet.get_page_margins_mut();
        margins.set_left(0.75);
        margins.set_right(0.75);
        margins.set_top(1.0);
        margins.set_bottom(1.0);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let parser = XlsxParser;
    let (doc, _warnings) = parser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();
    assert_eq!(
        doc.pages.len(),
        2,
        "the sixth column must overflow onto its own page like Excel"
    );
}

// ----- Page size and margins defaults -----

#[test]
fn test_page_size_defaults() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Test")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let default_size = PageSize::default();
    assert!((tp.size.width - default_size.width).abs() < 0.01);
    assert!((tp.size.height - default_size.height).abs() < 0.01);
}

/// Build a workbook whose only sheet has no cells, carrying a paper size and a
/// header/footer. LibreOffice writes exactly this shape for a workbook saved
/// with nothing typed into it.
///
/// The header and footer are declared so the tests can assert they are *not*
/// carried onto the blank page, not because the page renders them.
fn build_empty_sheet_xlsx(paper_size: u32, header: &str, footer: &str) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_page_setup_mut().set_paper_size(paper_size);
        sheet
            .get_header_footer_mut()
            .get_odd_header_mut()
            .set_value(header);
        sheet
            .get_header_footer_mut()
            .get_odd_footer_mut()
            .set_value(footer);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// A workbook whose only sheet has no cells still prints one page, and that
/// page is the size the sheet asks for.
///
/// The sheet loop skips a sheet with no used range, so a single-sheet workbook
/// reached codegen with no pages at all and the compiler's own default supplied
/// a blank A4 — the file's `<pageSetup paperSize="1"/>` never reached the
/// renderer (issue #632).
#[test]
fn test_empty_sheet_keeps_its_declared_paper_size() {
    // 1 = Letter.
    let data = build_empty_sheet_xlsx(1, "&CReport", "&CPage &P");
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1, "an empty sheet still prints one page");
    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.size.width - 612.0).abs() < 0.01 && (page.size.height - 792.0).abs() < 0.01,
        "expected Letter, got {:?}",
        page.size
    );
}

/// Triangulation: a different paper code must produce that code's size, so the
/// page cannot be a hardcoded Letter.
#[test]
fn test_empty_sheet_keeps_a_non_letter_paper_size() {
    // 5 = Legal.
    let data = build_empty_sheet_xlsx(5, "&CReport", "&CPage &P");
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.size.width - 612.0).abs() < 0.01 && (page.size.height - 1008.0).abs() < 0.01,
        "expected Legal, got {:?}",
        page.size
    );
}

/// The page an empty sheet prints stays blank.
///
/// The ground truth for a sheet with no used range is a blank page — Excel
/// declines to print one at all — so nothing is invented to fill it. Only the
/// paper the file asks for is restored.
#[test]
fn test_empty_sheet_page_stays_blank() {
    let data = build_empty_sheet_xlsx(1, "&CQuarterly", "&CPage &P");
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(page.header.is_none(), "no header on a page with no cells");
    assert!(page.footer.is_none(), "no footer on a page with no cells");
    assert!(page.table.rows.is_empty());
    assert!(page.images.is_empty() && page.charts.is_empty());
}

/// The page still carries the sheet's own print margins, not the renderer's.
#[test]
fn test_empty_sheet_page_keeps_its_print_margins() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_page_setup_mut().set_paper_size(1);
        sheet.get_page_margins_mut().set_left(1.25);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let (doc, _warnings) = XlsxParser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.margins.left - 90.0).abs() < 0.01,
        "expected 1.25in = 90pt, got {}",
        page.margins.left
    );
}

/// A sheet that does have cells keeps deciding the page count on its own — an
/// empty *second* sheet must not add a blank page, which is what Excel does.
#[test]
fn test_empty_sheet_alongside_a_used_sheet_adds_no_page() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        sheet.get_cell_mut("A1").set_value("Value");
    }
    book.new_sheet("Blank").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let (doc, _warnings) = XlsxParser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();

    assert_eq!(doc.pages.len(), 1, "the blank sheet contributes no page");
    assert_eq!(get_sheet_page(&doc, 0).name, "Data");
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    assert_eq!(cell.col_span, 1);
    assert_eq!(cell.row_span, 1);
    assert!(cell.border.is_none());
    assert!(cell.background.is_none());
}

#[path = "xlsx_cell_format_tests.rs"]
mod cell_format_tests;

#[path = "xlsx_page_feature_tests.rs"]
mod page_feature_tests;

#[path = "xlsx_condfmt_tests.rs"]
mod condfmt_tests;

#[path = "xlsx_chart_tests.rs"]
mod chart_tests;

#[path = "xlsx_streaming_tests.rs"]
mod streaming_tests;

/// The style of the first run of the first cell a parsed workbook produces.
fn first_cell_text_style(data: &[u8]) -> crate::ir::TextStyle {
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("workbook should parse");
    let Page::Sheet(sheet) = &doc.pages[0] else {
        panic!("expected a sheet page");
    };
    for row in &sheet.table.rows {
        for cell in &row.cells {
            for block in &cell.content {
                if let Block::Paragraph(paragraph) = block
                    && let Some(run) = paragraph.runs.first()
                {
                    return run.style.clone();
                }
            }
        }
    }
    panic!("expected a cell run");
}

#[test]
fn test_unstyled_cell_carries_the_workbook_normal_font() {
    // A cell with no `s` attribute uses cellXfs[0], whose font is the
    // workbook's Normal font. umya reports no font for such a cell, so the
    // style path has to fall back to styles.xml itself or the renderer picks
    // its own default family and size (issue #462).
    let data = build_xlsx_with_normal_font("Malgun Gothic", 12.0);
    let style = first_cell_text_style(&data);
    assert_eq!(style.font_family.as_deref(), Some("Malgun Gothic"));
    assert_eq!(style.font_size, Some(12.0));
}

#[test]
fn test_unstyled_cell_keeps_a_calibri_normal_font() {
    // Triangulation: Calibri is the most common Normal font and used to be
    // dropped on the grounds that it was "the default". It has to survive
    // like any other family, or Calibri workbooks render in the renderer's
    // serif default (issue #462).
    let data = build_xlsx_with_normal_font("Calibri", 11.0);
    let style = first_cell_text_style(&data);
    assert_eq!(style.font_family.as_deref(), Some("Calibri"));
    assert_eq!(style.font_size, Some(11.0));
}

#[test]
fn test_explicit_cell_font_overrides_the_workbook_normal_font() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("styled");
        cell.get_style_mut().get_font_mut().set_name("Georgia");
        cell.get_style_mut().get_font_mut().set_size(20.0);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    let style = first_cell_text_style(&cursor.into_inner());
    assert_eq!(style.font_family.as_deref(), Some("Georgia"));
    assert_eq!(style.font_size, Some(20.0));
}

/// A one-cell workbook whose Normal font (styles.xml font 0) is `family` at
/// `size_pt`, with the cell itself left unstyled.
fn build_xlsx_with_normal_font(family: &str, size_pt: f64) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("title");
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    rewrite_first_styles_font(&cursor.into_inner(), family, size_pt)
}

/// Rewrite the first `<font>` of `xl/styles.xml` in place. umya always
/// writes Calibri 11 there, so the fixture has to patch the part directly to
/// exercise a different Normal font.
fn rewrite_first_styles_font(data: &[u8], family: &str, size_pt: f64) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).expect("readable zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("readable entry");
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("readable entry body");
        if name == "xl/styles.xml" {
            let xml = String::from_utf8(bytes).expect("styles.xml is utf-8");
            let start = xml.find("<font>").expect("styles.xml has a font");
            let end = xml[start..].find("</font>").expect("font is closed") + start;
            let replacement =
                format!("<font><sz val=\"{size_pt}\"/><name val=\"{family}\"/></font>");
            bytes = format!(
                "{}{}{}",
                &xml[..start],
                replacement,
                &xml[end + "</font>".len()..]
            )
            .into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("writable entry");
        std::io::Write::write_all(&mut out, &bytes).expect("writable entry body");
    }
    out.finish().expect("finished zip").into_inner()
}
