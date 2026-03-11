use super::*;

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
    let tp = get_sheet_page(&doc, 0);
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
    let tp0 = get_sheet_page(&doc, 0);
    let tp1 = get_sheet_page(&doc, 1);
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
    let tp = get_sheet_page(&doc, 0);
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
    assert_eq!(tp.table.column_widths.len(), 2);
}

#[test]
fn test_print_area_without_dollar_signs() {
    let data = build_xlsx_with_print_area(
        &[("A1", "X"), ("B1", "Y"), ("A2", "Z"), ("B2", "W")],
        "Sheet1!A1:A2",
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows.len(), 2);
    assert_eq!(tp.table.rows[0].cells.len(), 1, "Only column A");
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "X");
    assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Z");
}

#[test]
fn test_no_print_area_includes_all() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "All"), ("C3", "Data")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows.len(), 3);
    assert_eq!(tp.table.rows[0].cells.len(), 3);
}

#[test]
fn test_row_page_breaks_split_into_pages() {
    let data = build_xlsx_with_row_breaks(
        &[
            ("A1", "R1"),
            ("A2", "R2"),
            ("A3", "R3"),
            ("A4", "R4"),
            ("A5", "R5"),
        ],
        &[2],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2, "Break should split into 2 pages");
    let tp0 = get_sheet_page(&doc, 0);
    let tp1 = get_sheet_page(&doc, 1);

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
    let data = build_xlsx_with_row_breaks(
        &[
            ("A1", "R1"),
            ("A2", "R2"),
            ("A3", "R3"),
            ("A4", "R4"),
            ("A5", "R5"),
            ("A6", "R6"),
        ],
        &[2, 4],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 3, "Two breaks should produce 3 pages");
    let tp0 = get_sheet_page(&doc, 0);
    let tp1 = get_sheet_page(&doc, 1);
    let tp2 = get_sheet_page(&doc, 2);

    assert_eq!(tp0.table.rows.len(), 2);
    assert_eq!(tp1.table.rows.len(), 2);
    assert_eq!(tp2.table.rows.len(), 2);

    assert_eq!(cell_text(&tp0.table.rows[0].cells[0]), "R1");
    assert_eq!(cell_text(&tp1.table.rows[0].cells[0]), "R3");
    assert_eq!(cell_text(&tp2.table.rows[0].cells[0]), "R5");
}

#[test]
fn test_no_page_breaks_single_page() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "R1"), ("A2", "R2"), ("A3", "R3")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows.len(), 3);
}

#[test]
fn test_page_break_column_widths_preserved() {
    let data = build_xlsx_with_row_breaks(
        &[("A1", "R1"), ("B1", "R1B"), ("A2", "R2"), ("B2", "R2B")],
        &[1],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2);
    let tp0 = get_sheet_page(&doc, 0);
    let tp1 = get_sheet_page(&doc, 1);
    assert_eq!(tp0.table.column_widths.len(), 2);
    assert_eq!(tp1.table.column_widths.len(), 2);
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

    assert_eq!(hf.paragraphs[0].style.alignment, Some(Alignment::Left));
    match &hf.paragraphs[0].elements[0] {
        HFInline::Run(r) => assert_eq!(r.text, "Left Text"),
        _ => panic!("Expected Run"),
    }

    assert_eq!(hf.paragraphs[1].style.alignment, Some(Alignment::Center));
    match &hf.paragraphs[1].elements[0] {
        HFInline::Run(r) => assert_eq!(r.text, "Center Text"),
        _ => panic!("Expected Run"),
    }

    assert_eq!(hf.paragraphs[2].style.alignment, Some(Alignment::Right));
    match &hf.paragraphs[2].elements[0] {
        HFInline::Run(r) => assert_eq!(r.text, "Right Text"),
        _ => panic!("Expected Run"),
    }
}

#[test]
fn test_parse_hf_format_string_page_numbers() {
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
    let footer = tp.footer.as_ref().expect("Expected footer");
    assert_eq!(footer.paragraphs.len(), 1);
    let elems = &footer.paragraphs[0].elements;
    assert_eq!(elems.len(), 4);
    assert!(matches!(elems[1], HFInline::PageNumber));
    assert!(matches!(elems[3], HFInline::TotalPages));
}

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

    let _ = doc.metadata;
}
