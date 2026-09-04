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

// ----- Hidden worksheet tests (issue #1065) -----

/// A workbook that hides its lookup sheet prints only the two visible ones.
///
/// Modelled on the audited gift-budget workbook of issue #982: two visible
/// sheets in front of a `state="hidden"` data sheet. Excel and LibreOffice
/// both export two pages from it; office2pdf paged the hidden sheet too.
#[test]
fn test_hidden_sheet_is_not_printed() {
    let data = build_xlsx_multi_sheet_with_states(&[
        (
            "Start",
            umya_spreadsheet::SheetStateValues::Visible,
            &[("A1", "Welcome")],
        ),
        (
            "Gift budget and tracker",
            umya_spreadsheet::SheetStateValues::Visible,
            &[("A1", "Gift"), ("B1", "Budget")],
        ),
        (
            "Data",
            umya_spreadsheet::SheetStateValues::Hidden,
            &[("A1", "Lookup")],
        ),
    ]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let names: Vec<&str> = doc
        .pages
        .iter()
        .map(|page| match page {
            Page::Sheet(sheet) => sheet.name.as_str(),
            _ => panic!("expected a sheet page"),
        })
        .collect();
    assert_eq!(names, vec!["Start", "Gift budget and tracker"]);
}

/// `veryHidden` is the state a sheet hidden from the unhide dialog carries;
/// Excel prints it no more than a plain hidden one.
#[test]
fn test_very_hidden_sheet_is_not_printed() {
    let data = build_xlsx_multi_sheet_with_states(&[
        (
            "Report",
            umya_spreadsheet::SheetStateValues::Visible,
            &[("A1", "Quarterly revenue")],
        ),
        (
            "Config",
            umya_spreadsheet::SheetStateValues::VeryHidden,
            &[("A1", "Region codes")],
        ),
    ]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    assert_eq!(get_sheet_page(&doc, 0).name, "Report");
}

/// `--sheets Data` asks for the hidden sheet by name, which is the one way to
/// print it: the caller has overridden the workbook's own visibility.
#[test]
fn test_hidden_sheet_prints_when_named_explicitly() {
    let data = build_xlsx_multi_sheet_with_states(&[
        (
            "Start",
            umya_spreadsheet::SheetStateValues::Visible,
            &[("A1", "Welcome")],
        ),
        (
            "Data",
            umya_spreadsheet::SheetStateValues::Hidden,
            &[("A1", "Lookup")],
        ),
    ]);
    let parser = XlsxParser;
    let opts = ConvertOptions {
        sheet_names: Some(vec!["Data".to_string()]),
        ..Default::default()
    };
    let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

    assert_eq!(doc.pages.len(), 1);
    let page = get_sheet_page(&doc, 0);
    assert_eq!(page.name, "Data");
    assert_eq!(cell_text(&page.table.rows[0].cells[0]), "Lookup");
}

/// The blank page a workbook with no used cells prints comes from the first
/// sheet that would actually print, not from a hidden one preceding it.
#[test]
fn test_empty_workbook_page_skips_a_hidden_first_sheet() {
    let data = build_xlsx_multi_sheet_with_states(&[
        ("Macros", umya_spreadsheet::SheetStateValues::Hidden, &[]),
        ("Blank", umya_spreadsheet::SheetStateValues::Visible, &[]),
    ]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    assert_eq!(get_sheet_page(&doc, 0).name, "Blank");
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

/// Parse a header/footer string against a fixed sheet name, discarding
/// warnings. Tests that care about either pass them explicitly instead.
fn parse_hf(format_str: &str) -> Option<HeaderFooter> {
    parse_hf_format_string(format_str, "Sheet1", None, &mut Vec::new())
}

/// The concatenated text of every section, in left/center/right order.
fn hf_section_texts(hf: &HeaderFooter) -> Vec<String> {
    hf.paragraphs
        .iter()
        .map(|p| {
            p.elements
                .iter()
                .filter_map(|e| match e {
                    HFInline::Run(run) => Some(run.text.clone()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect()
}

/// `&A` prints the worksheet name (issue #690).
///
/// This is the exact `<oddHeader>` of
/// `tests/fixtures/xlsx/libreoffice/page_scale.xlsx`. `&A` was its only
/// content, so the catch-all arm left the section empty and the header
/// paragraph was dropped entirely — Excel prints "Sheet1".
#[test]
fn test_sheet_name_code_resolves_to_the_worksheet_name() {
    let hf = parse_hf_format_string(
        r#"&C&"Times New Roman,Regular"&12&A"#,
        "Sheet1",
        None,
        &mut Vec::new(),
    )
    .expect("a header whose only content is &A still produces a paragraph");

    assert_eq!(hf_section_texts(&hf), vec!["Sheet1"]);
    assert_eq!(hf.paragraphs[0].style.alignment, Some(Alignment::Center));
}

/// Triangulation: the name comes from the sheet, not from a constant.
#[test]
fn test_sheet_name_code_uses_the_actual_sheet_name() {
    let hf =
        parse_hf_format_string("&C&A", "Q3 Budget", None, &mut Vec::new()).expect("header parsed");

    assert_eq!(hf_section_texts(&hf), vec!["Q3 Budget"]);
}

/// `&A` composes with surrounding literal text and with other field codes,
/// in whichever section it appears.
#[test]
fn test_sheet_name_code_composes_with_text_and_other_fields() {
    let hf = parse_hf_format_string("&LSheet: &A&RPage &P", "Summary", None, &mut Vec::new())
        .expect("header parsed");

    assert_eq!(hf_section_texts(&hf), vec!["Sheet: Summary", "Page "]);
    assert_eq!(hf.paragraphs[0].style.alignment, Some(Alignment::Left));
    assert_eq!(hf.paragraphs[1].style.alignment, Some(Alignment::Right));
    assert!(
        hf.paragraphs[1]
            .elements
            .iter()
            .any(|e| matches!(e, HFInline::PageNumber(_))),
        "&P must still resolve alongside &A"
    );
}

/// Codes naming data the parser does not hold are reported rather than
/// silently discarded by a catch-all (issue #690).
///
/// `&F`/`&Z` need the source path, which never reaches `Parser::parse`; `&D`
/// and `&T` are Excel's print date/time; `&G` is a picture.
#[test]
fn test_unresolvable_field_codes_warn_instead_of_vanishing() {
    for (code, description) in [
        ("&F", "&F (file name)"),
        ("&Z", "&Z (file path)"),
        ("&D", "&D (print date)"),
        ("&T", "&T (print time)"),
        ("&G", "&G (picture)"),
    ] {
        let mut warnings: Vec<ConvertWarning> = Vec::new();
        parse_hf_format_string(&format!("&C{code}"), "Sheet1", None, &mut warnings);

        assert!(
            warnings.iter().any(|w| matches!(
                w,
                ConvertWarning::UnsupportedElement { format, element }
                    if format == "XLSX" && element.contains(description)
            )),
            "{code} must be reported, got {warnings:?}"
        );
    }
}

/// A code that resolves must not also warn.
#[test]
fn test_resolved_field_codes_do_not_warn() {
    let mut warnings: Vec<ConvertWarning> = Vec::new();
    parse_hf_format_string("&C&A &P of &N", "Sheet1", None, &mut warnings);

    assert!(
        warnings.is_empty(),
        "expected no warnings, got {warnings:?}"
    );
}

#[test]
fn test_parse_hf_format_string_empty() {
    assert!(parse_hf("").is_none());
    assert!(parse_hf("   ").is_none());
}

#[test]
fn test_parse_hf_format_string_center_only() {
    let hf = parse_hf("My Report").unwrap();
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
    let hf = parse_hf("&LLeft Text&CCenter Text&RRight Text").unwrap();
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
    let hf = parse_hf("&CPage &P of &N").unwrap();
    assert_eq!(hf.paragraphs.len(), 1);
    let elems = &hf.paragraphs[0].elements;
    assert_eq!(elems.len(), 4);
    match &elems[0] {
        HFInline::Run(r) => assert_eq!(r.text, "Page "),
        _ => panic!("Expected Run"),
    }
    assert!(matches!(elems[1], HFInline::PageNumber(_)));
    match &elems[2] {
        HFInline::Run(r) => assert_eq!(r.text, " of "),
        _ => panic!("Expected Run"),
    }
    assert!(matches!(elems[3], HFInline::TotalPages(_)));
}

#[test]
fn test_parse_hf_format_string_escaped_ampersand() {
    let hf = parse_hf("&CA && B").unwrap();
    assert_eq!(hf.paragraphs.len(), 1);
    match &hf.paragraphs[0].elements[0] {
        HFInline::Run(r) => assert_eq!(r.text, "A & B"),
        _ => panic!("Expected Run"),
    }
}

#[test]
fn test_parse_hf_format_string_font_codes_skipped() {
    let hf = parse_hf(r#"&C&"Arial"&12Hello"#).unwrap();
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

/// Helper: build an XLSX whose sheet states `<pageMargins>` alongside a footer.
fn build_xlsx_with_footer_margins(footer_str: &str, footer_in: f64, bottom_in: f64) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Data");
        sheet
            .get_header_footer_mut()
            .get_odd_footer_mut()
            .set_value(footer_str);
        let margins = sheet.get_page_margins_mut();
        margins.set_footer(footer_in);
        margins.set_bottom(bottom_in);
    }
    let mut buf = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut buf).unwrap();
    buf.into_inner()
}

/// Excel measures a printed footer up from the page's bottom edge, through
/// `<pageMargins>/@footer`, and leaves a further 2pt below the text's line box
/// (issue #1142).
///
/// The 2pt is measured: on Excel-for-Mac exports of one-factor variants of
/// `tests/fixtures/xlsx/headerFooterTest.xlsx`, a 12pt Calibri footer over a
/// 0.5in footer margin puts its baseline 41pt above the page's bottom edge —
/// 36pt of margin, Calibri's 3.22pt `hhea` descent, and 2pt between them. The
/// same series holds at 6, 8, 14, 20, 40 and 80pt, and across Arial, Verdana,
/// Times New Roman and Aptos.
#[test]
fn a_sheet_footer_is_seated_from_the_page_bottom_edge() {
    let data = build_xlsx_with_footer_margins("&LSensitivity: Internal", 0.3, 0.75);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let footer = get_sheet_page(&doc, 0)
        .footer
        .as_ref()
        .expect("the sheet states a footer");
    assert_eq!(
        footer.distance_from_edge,
        Some(23.0),
        "0.3in floors to 21pt of footer margin, plus Excel's 2pt band inset"
    );
}

/// Triangulation for the seat: it tracks `@footer`, and nothing else on the
/// page moves it (issue #1142).
#[test]
fn a_sheet_footer_seat_follows_its_own_margin_not_the_bottom_one() {
    let parser = XlsxParser;
    let seat_of = |footer_in: f64, bottom_in: f64| -> Option<f64> {
        let data = build_xlsx_with_footer_margins("&LSensitivity: Internal", footer_in, bottom_in);
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        get_sheet_page(&doc, 0)
            .footer
            .as_ref()
            .expect("the sheet states a footer")
            .distance_from_edge
    };

    assert_eq!(
        seat_of(0.5, 0.75),
        Some(38.0),
        "0.5in is 36pt plus the inset"
    );
    assert_eq!(
        seat_of(0.5, 1.5),
        Some(38.0),
        "doubling the bottom margin must not move the footer"
    );
    assert_eq!(
        seat_of(1.0, 1.5),
        Some(74.0),
        "1.0in is 72pt plus the inset"
    );
}

/// A sheet that states no `<pageMargins>` takes Excel's own 0.3in default.
#[test]
fn a_sheet_footer_without_page_margins_takes_excels_default() {
    let data = build_xlsx_with_footer("&LSensitivity: Internal");
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let footer = get_sheet_page(&doc, 0)
        .footer
        .as_ref()
        .expect("the sheet states a footer");
    assert_eq!(footer.distance_from_edge, Some(23.0));
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
    assert!(matches!(elems[1], HFInline::PageNumber(_)));
    assert!(matches!(elems[3], HFInline::TotalPages(_)));
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

#[test]
fn test_hf_font_color_code_is_stripped() {
    // Excel emits &KRRGGBB for header colors; the six hex digits must not
    // leak into the text ("000000top center").
    let hf = parse_hf(
        r#"&L&"Calibri,Regular"&K000000top left&C&"Calibri,Regular"&K000000top center&R&"Calibri,Regular"&K000000top right"#,
    )
    .expect("header parsed");
    let texts: Vec<String> = hf
        .paragraphs
        .iter()
        .map(|p| {
            p.elements
                .iter()
                .filter_map(|e| match e {
                    HFInline::Run(run) => Some(run.text.clone()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect();
    assert_eq!(texts, vec!["top left", "top center", "top right"]);
}

// --- Header/footer font and size codes (issue #633) ---

/// The style of the first run of the first section.
fn hf_first_style(hf: &HeaderFooter) -> TextStyle {
    hf.paragraphs
        .iter()
        .flat_map(|p| p.elements.iter())
        .find_map(|element| match element {
            HFInline::Run(run) => Some(run.style.clone()),
            _ => None,
        })
        .expect("a run in the header")
}

/// `&"Font,Style"` selects the face for the runs after it, and `&<n>` the size.
/// Both used to be parsed and discarded, so every header printed in the
/// fallback face at the document default size.
#[test]
fn test_hf_font_and_size_codes_reach_the_runs() {
    let hf = parse_hf("&L&\"Calibri,Bold\"&12left").expect("header parsed");
    let style = hf_first_style(&hf);

    assert_eq!(style.font_family.as_deref(), Some("Calibri"));
    assert_eq!(style.east_asian_font_family.as_deref(), Some("Calibri"));
    assert_eq!(style.font_size, Some(12.0));
    assert_eq!(style.bold, Some(true));
    assert_eq!(
        style.italic, None,
        "no italic word, so the default survives"
    );
}

/// A section with no codes keeps every field unset, so the renderer's own
/// defaults apply rather than being pinned by the parser.
#[test]
fn test_hf_without_codes_carries_no_style() {
    let style = hf_first_style(&parse_hf("&Lplain").expect("header parsed"));

    assert_eq!(style.font_family, None);
    assert_eq!(style.font_size, None);
    assert_eq!(style.bold, None);
    assert_eq!(style.italic, None);
}

/// Excel writes `&"-,Bold"` to change the style while keeping the face, and
/// `Regular` to turn the flags back off.
#[test]
fn test_hf_font_code_dash_keeps_the_face_and_regular_clears_the_style() {
    let hf = parse_hf("&L&\"Calibri,Bold\"bold&\"-,Regular\"plain").expect("header parsed");
    let styles: Vec<TextStyle> = hf
        .paragraphs
        .iter()
        .flat_map(|p| p.elements.iter())
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.style.clone()),
            _ => None,
        })
        .collect();

    assert_eq!(styles.len(), 2, "one run per style change, got {styles:?}");
    assert_eq!(styles[0].bold, Some(true));
    assert_eq!(styles[1].bold, None, "Regular turns bold back off");
    assert_eq!(
        styles[1].font_family.as_deref(),
        Some("Calibri"),
        "`-` keeps the face the previous code set"
    );
}

/// A code partway through a section applies only from that point on, so the
/// text before it keeps the earlier face.
#[test]
fn test_hf_code_midway_splits_the_section() {
    let hf = parse_hf("&Lplain&\"Calibri,Regular\"styled").expect("header parsed");
    let texts: Vec<(String, Option<String>)> = hf
        .paragraphs
        .iter()
        .flat_map(|p| p.elements.iter())
        .filter_map(|element| match element {
            HFInline::Run(run) => Some((run.text.clone(), run.style.font_family.clone())),
            _ => None,
        })
        .collect();

    assert_eq!(
        texts,
        vec![
            ("plain".to_string(), None),
            ("styled".to_string(), Some("Calibri".to_string())),
        ]
    );
}

/// A header/footer run before the format string's first `&"Font"` code takes
/// the workbook's Normal font, as an unstyled cell does. Leaving it unset sent
/// it to the renderer's ambient default — a serif — so the Gantt template of
/// issue #841 printed the run ahead of its `&"Aptos"` code in a serif nothing
/// else on the sheet used (issue #951).
///
/// `_x005F_x000D_` decodes to the literal text `_x000D_`, exactly as it does
/// in the pinned native Excel probe from `Gift Budget and Tracker1.xlsx`.
#[test]
fn a_header_footer_run_before_any_font_code_takes_the_normal_font() {
    let normal_font = NormalFont {
        family: "Corbel".to_string(),
        size_pt: 11.0,
        color: Some(Color::new(0x44, 0x54, 0x6A)),
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };
    let hf = parse_hf_format_string(
        r#"&L_x005F_x000D_&1#&"Aptos"&8&K000000 Sensitivity: Internal"#,
        "Sheet1",
        Some(&normal_font),
        &mut Vec::new(),
    )
    .expect("footer parsed");
    let styles: Vec<(Option<String>, Option<f64>, Option<Color>)> = hf.paragraphs[0]
        .elements
        .iter()
        .filter_map(|element| match element {
            HFInline::Run(run) => Some((
                run.style.font_family.clone(),
                run.style.font_size,
                run.style.color,
            )),
            _ => None,
        })
        .collect();

    assert_eq!(
        styles,
        vec![
            (
                Some("Corbel".to_string()),
                Some(11.0),
                Some(Color::new(0x44, 0x54, 0x6A)),
            ),
            (
                Some("Corbel".to_string()),
                Some(1.0),
                Some(Color::new(0x44, 0x54, 0x6A)),
            ),
            (Some("Aptos".to_string()), Some(8.0), Some(Color::black()),),
        ],
        "the run before the code takes the whole Normal font; later codes keep their own sizes"
    );
}

/// A workbook whose Normal font could not be read leaves the family unstated,
/// exactly as before, rather than inventing one (issue #951).
#[test]
fn a_header_footer_without_a_normal_font_states_no_family() {
    let hf =
        parse_hf_format_string("&LPlain", "Sheet1", None, &mut Vec::new()).expect("footer parsed");
    let HFInline::Run(run) = &hf.paragraphs[0].elements[0] else {
        panic!("expected a run");
    };
    assert_eq!(run.style.font_family, None);
}

/// `_xHHHH_` is ECMA-376's escape for a character a spreadsheet string cannot
/// carry directly (§22.9.2.19, `ST_Xstring`). Excel writes a footer's line
/// break as `_x000D_`; we printed the seven characters (issue #929).
#[test]
fn a_header_footer_decodes_xstring_escapes() {
    let hf = parse_hf_format_string(
        r#"&L_x000D_&1#&"Aptos"&8&K000000 Sensitivity: Internal"#,
        "Sheet1",
        None,
        &mut Vec::new(),
    )
    .expect("footer parsed");
    let text: String = hf_section_texts(&hf).join("");
    assert!(
        !text.contains("_x000D_") && !text.contains("x000D"),
        "the escape is decoded, not printed: {text:?}"
    );
    assert_eq!(text, "# Sensitivity: Internal");
}

/// The escape is case-insensitive per the schema, and `_x005F_` is the way a
/// literal underscore is written — decoding it first is what lets a genuine
/// `_x000D_` in the *text* survive as seven characters (issue #929).
#[test]
fn a_header_footer_escape_is_case_insensitive_and_underscore_safe() {
    let upper = parse_hf_format_string("&Ca_x0062_c", "Sheet1", None, &mut Vec::new())
        .expect("header parsed");
    assert_eq!(hf_section_texts(&upper), vec!["abc"]);

    // `_x005F_x000D_` is a literal `_` followed by `x000D_`, not a carriage
    // return: the underscore escape is consumed and its output not rescanned.
    let literal = parse_hf_format_string("&C_x005F_x000D_", "Sheet1", None, &mut Vec::new())
        .expect("header parsed");
    assert_eq!(hf_section_texts(&literal), vec!["_x000D_"]);
}

/// A decoded carriage return breaks the section into lines rather than
/// vanishing: text after it starts a new footer line (issue #929).
#[test]
fn a_decoded_carriage_return_breaks_the_section_into_lines() {
    let hf = parse_hf_format_string("&Ltop_x000D_bottom", "Sheet1", None, &mut Vec::new())
        .expect("footer parsed");
    assert_eq!(hf_section_texts(&hf), vec!["top", "bottom"]);
    for paragraph in &hf.paragraphs {
        assert_eq!(paragraph.style.alignment, Some(Alignment::Left));
    }
}

/// A character the escape names but that no text can carry — a control code
/// other than the line break — is dropped rather than emitted raw, so it
/// cannot reach the renderer as an unprintable glyph (issue #929).
#[test]
fn a_header_footer_drops_a_decoded_control_character() {
    let hf = parse_hf_format_string("&Ca_x0007_b", "Sheet1", None, &mut Vec::new())
        .expect("header parsed");
    assert_eq!(hf_section_texts(&hf), vec!["ab"]);
}
