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

/// One sheet for `build_xlsx_multi_sheet_with_states`: its name, the
/// visibility state to declare for it, and its cells as (coordinate, value).
type SheetWithState<'a> = (
    &'a str,
    umya_spreadsheet::SheetStateValues,
    &'a [(&'a str, &'a str)],
);

/// Helper: build XLSX with multiple sheets, each declaring the visibility
/// state `xl/workbook.xml` carries on its `<sheet>` entry.
fn build_xlsx_multi_sheet_with_states(sheets: &[SheetWithState<'_>]) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    for (index, (name, state, cells)) in sheets.iter().enumerate() {
        if index > 0 {
            book.add_sheet(umya_spreadsheet::Worksheet::default())
                .unwrap();
        }
        let sheet = book.get_sheet_mut(&index).unwrap();
        sheet.set_name(*name);
        sheet.set_state(state.clone());
        for &(coord, value) in *cells {
            sheet.get_cell_mut(coord).set_value(value);
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
    // A1 filled, B1 empty, A2 empty, B2 filled → 2-row grid with gaps.
    // "Bottom-Right" overflows B2's default-width column, so the printed
    // range extends one column past the used range, as Excel prints an
    // unwrapped overflow (issue #718) — a third, empty column follows.
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Top-Left"), ("B2", "Bottom-Right")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows.len(), 2);
    assert_eq!(tp.table.rows[0].cells.len(), 3);
    assert_eq!(cell_text(&tp.table.rows[0].cells[2]), "");
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
    // umya's writer auto-emits `<col min="1" max="2" width="8.38"
    // customWidth="1"/>` for used columns, so this fixture takes the
    // declared-width path: round(8.38 × 6pt Calibri-11 unit) = 50pt on the
    // integer point grid (issue #621; the old pixel model printed 50.28pt).
    for w in &tp.table.column_widths {
        assert_eq!(
            *w, 50.0,
            "Expected declared 8.38-unit width of 50pt, got {w}"
        );
    }
}

#[test]
fn test_carlito_column_widths_match_native_print_metrics() {
    // Carlito 11 has a 6pt column unit (issue #621), so the pr_186 fixture's
    // native 26/20/24-unit columns print 156/120/144pt.
    assert_eq!(column_width_to_pt(26.0, 6.0), 156.0);
    assert_eq!(column_width_to_pt(20.0, 6.0), 120.0);
    assert_eq!(column_width_to_pt(24.0, 6.0), 144.0);
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

    // Styles-unreadable fallback: the dominant Carlito face at the assumed
    // 11pt Normal size gives the same 6pt unit as a declared Carlito-11
    // Normal font (issue #621).
    assert_eq!(sheet_column_unit_pt(sheet), 6.0);
}

/// The column character-unit is an INTEGER POINT count: round-half-up of the
/// Normal font's max digit advance in points. Measured on 17 one-factor
/// native Excel-for-Mac probes (issue #621): each family/size pair below is a
/// discriminator — Calibri 10 → 5pt kills every integer-96dpi-pixel model
/// (the old ceil gave 7px = 5.25pt), Times New Roman 13 (exactly 6.500pt)
/// rounds UP to 7 (kills half-even), Calibri 9 and Verdana 11 kill
/// truncation, Calibri 10 and Verdana 10 kill ceiling.
#[test]
fn test_column_unit_pt_is_integer_points_from_digit_advance() {
    assert_eq!(column_unit_pt("Calibri", 9.0), 5.0);
    assert_eq!(column_unit_pt("Calibri", 10.0), 5.0);
    assert_eq!(column_unit_pt("Calibri", 11.0), 6.0);
    assert_eq!(column_unit_pt("Calibri", 12.0), 6.0);
    assert_eq!(column_unit_pt("Arial", 10.0), 6.0);
    assert_eq!(column_unit_pt("Arial", 12.0), 7.0);
    assert_eq!(column_unit_pt("Verdana", 10.0), 6.0);
    assert_eq!(column_unit_pt("Verdana", 11.0), 7.0);
    assert_eq!(column_unit_pt("Times New Roman", 12.0), 6.0);
    assert_eq!(column_unit_pt("Times New Roman", 13.0), 7.0);
    assert_eq!(column_unit_pt("Courier New", 10.0), 6.0);
    assert_eq!(column_unit_pt("Courier New", 12.0), 7.0);
    assert_eq!(column_unit_pt("Malgun Gothic", 10.0), 6.0);
    assert_eq!(column_unit_pt("Malgun Gothic", 11.0), 6.0);
    assert_eq!(column_unit_pt("Segoe UI", 10.0), 5.0);
}

/// The reference digit advances are the real `hmtx` maxima over U+0030..=0039
/// of the faces Excel itself resolves (read from Excel's own DFonts/system
/// faces by the issue #621 probe tooling). They pin the wasm/font-less arm so
/// output stays deterministic, and they outrank live font resolution so a
/// machine substituting a digit-incompatible face (Calibri → Liberation Sans
/// is 0.556em against Calibri's 0.5068) cannot shift column geometry.
#[test]
fn test_reference_digit_advance_em_pins_excel_face_metrics() {
    let calibri: f64 = reference_digit_advance_em("Calibri").unwrap();
    assert!((calibri - 0.506836).abs() < 1e-6);
    assert_eq!(
        reference_digit_advance_em("Carlito"),
        reference_digit_advance_em("Calibri"),
        "Carlito is metrically identical to Calibri"
    );
    let arial: f64 = reference_digit_advance_em("Arial").unwrap();
    assert!((arial - 0.556152).abs() < 1e-6);
    let verdana: f64 = reference_digit_advance_em("Verdana").unwrap();
    assert!((verdana - 0.635742).abs() < 1e-6);
    let times: f64 = reference_digit_advance_em("Times New Roman").unwrap();
    assert!((times - 0.500000).abs() < 1e-6);
    let courier: f64 = reference_digit_advance_em("Courier New").unwrap();
    assert!((courier - 0.600098).abs() < 1e-6);
    // The repo's previous 0.529em Malgun estimate was wrong: the real face
    // advances 0.550781em (issue #621 probe artifacts).
    let malgun: f64 = reference_digit_advance_em("Malgun Gothic").unwrap();
    assert!((malgun - 0.550781).abs() < 1e-6);
    assert_eq!(
        reference_digit_advance_em("맑은 고딕"),
        reference_digit_advance_em("Malgun Gothic"),
        "the localized Malgun name must map to the same face"
    );
    let segoe: f64 = reference_digit_advance_em("Segoe UI").unwrap();
    assert!((segoe - 0.5390625).abs() < 1e-6);
    assert_eq!(
        reference_digit_advance_em("Selawik"),
        Some(segoe),
        "Microsoft's Selawik replacement is metric-compatible with Segoe UI"
    );
    assert_eq!(
        reference_digit_advance_em("Definitely Not A Font"),
        None,
        "unknown families fall through to live font resolution"
    );
}

/// Families outside the reference table resolve their digit advance from the
/// real face `hmtx`, exactly as Excel measures the face it resolves. The
/// embedded Libertinus Serif face makes this deterministic on every target:
/// its digit advance is 465/1000 em.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_max_digit_advance_em_reads_real_face_hmtx() {
    let advance: f64 = crate::render::pdf::max_digit_advance_em("Libertinus Serif")
        .expect("the embedded Libertinus Serif face must resolve");
    assert!(
        (advance - 0.465).abs() < 1e-6,
        "Libertinus Serif digit advance should be 0.465em, got {advance}"
    );
    // And the column metric consumes it: round(0.465 × 11pt) = 5pt.
    assert_eq!(column_unit_pt("Libertinus Serif", 11.0), 5.0);
}

/// The single-line width estimate prices each ASCII character against the
/// family's own digit advance, so a line of narrow letters costs a fraction of
/// a line of capitals. The flat half-em-per-character rule it replaced put a
/// realistic sentence a third over its real advance, which walked the printed
/// range past the page and split a one-page sheet in two (issue #1054).
///
/// Ground truth is the real `hmtx` advance sum of each named face over the
/// literal string, read from those faces' own tables: 41.6772em of ArialMT,
/// 22.9761em of Verdana, 8.6875em of Calibri.
#[test]
fn test_estimate_line_width_tracks_real_face_advances() {
    let cases: [(&str, &str, f64, f64); 3] = [
        (
            "Arial",
            "Blue = target (input). Achievement = Actual / Target; \
             for 'lower is better' metrics read inversely.",
            9.0,
            41.6772,
        ),
        (
            "Verdana",
            "Fractions, Tricky With Much Longer Text Here",
            10.0,
            22.9761,
        ),
        ("Calibri", "Monthly Active Users", 11.0, 8.6875),
    ];
    for (family, text, size_pt, truth_em) in cases {
        let truth_pt: f64 = truth_em * size_pt;
        let estimate: f64 = estimate_line_width_pt(text, Some(family), size_pt);
        let error: f64 = (estimate - truth_pt) / truth_pt;
        assert!(
            error.abs() < 0.05,
            "{family} {size_pt}pt: estimate {estimate:.2}pt is {:.1}% off the \
             face's own {truth_pt:.2}pt",
            error * 100.0
        );
    }

    // Triangulation: the estimate must follow the glyphs, not the character
    // count. Ten Arial 'i' advance 2.2217em against ten 'W' at 9.4385em, a
    // ratio the flat per-character rule collapsed to 1.
    let narrow: f64 = estimate_line_width_pt(&"i".repeat(10), Some("Arial"), 10.0);
    let wide: f64 = estimate_line_width_pt(&"W".repeat(10), Some("Arial"), 10.0);
    assert!(
        (wide / narrow - 9.4385 / 2.2217).abs() < 0.5,
        "narrow-to-wide ratio {:.2} should track the face's 4.25",
        wide / narrow
    );

    // A cell naming no family is priced on Excel's default Normal font, the
    // same last resort `column_unit_pt` falls back to.
    assert_eq!(
        estimate_line_width_pt("Monthly Active Users", None, 11.0),
        estimate_line_width_pt("Monthly Active Users", Some("Calibri"), 11.0),
    );
}

/// A declared column width prints as an integer point count: Excel quantizes
/// `width × unit` per column. Probe calibri11frac (issue #621): width 10.6 at
/// the 6pt Calibri-11 unit prints 64pt, not 63.6pt.
#[test]
fn test_declared_column_width_quantizes_to_integer_points() {
    assert_eq!(column_width_to_pt(10.6, 6.0), 64.0);
    // Whole-unit widths land on exact multiples — the pr_186 fixture's
    // Carlito-11 26/20/24-unit columns stay 156/120/144pt.
    assert_eq!(column_width_to_pt(26.0, 6.0), 156.0);
    assert_eq!(column_width_to_pt(20.0, 6.0), 120.0);
    assert_eq!(column_width_to_pt(24.0, 6.0), 144.0);
}

/// A column with no `<col>` entry and no declared `defaultColWidth` prints at
/// `baseColWidth × unit + 5` points — NOT 8.43 character units — where
/// `baseColWidth` defaults to 8 when `sheetFormatPr` does not declare it.
/// Verified by the issue #621 probes: at the 6pt Calibri-11 unit,
/// baseColWidth 10 → 65pt and 12 → 77pt (round-3 probes calibri11base10/12),
/// absent → 53pt; units 5/7 with no baseColWidth → 45/61pt. A declared
/// `defaultColWidth` outranks `baseColWidth` and goes through the
/// declared-units quantization instead.
#[test]
fn test_default_column_width_is_base_col_width_units_plus_five_points() {
    assert_eq!(default_column_width_pt(None, None, 5.0), 45.0);
    assert_eq!(default_column_width_pt(None, None, 6.0), 53.0);
    assert_eq!(default_column_width_pt(None, None, 7.0), 61.0);
    // Measured baseColWidth probes (no defaultColWidth): 10 → 65, 12 → 77.
    assert_eq!(default_column_width_pt(None, Some(10), 6.0), 65.0);
    assert_eq!(default_column_width_pt(None, Some(12), 6.0), 77.0);
    // Declared defaultColWidth quantizes like any declared width and
    // outranks baseColWidth.
    assert_eq!(default_column_width_pt(Some(10.6), None, 6.0), 64.0);
    assert_eq!(default_column_width_pt(Some(10.6), Some(12), 6.0), 64.0);
}

/// `declared_base_column_width` surfaces `sheetFormatPr@baseColWidth` only
/// when the file declares one — umya reports 0 for an absent attribute,
/// a width Excel never writes.
#[test]
fn test_declared_base_column_width_reads_sheet_format_pr() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    assert_eq!(declared_base_column_width(sheet), None);
    sheet
        .get_sheet_format_properties_mut()
        .set_base_column_width(10);
    assert_eq!(declared_base_column_width(sheet), Some(10));
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

    let normal_font =
        extract_normal_font(&data, Some(book.get_theme())).expect("styles.xml has a Normal font");
    assert_eq!(normal_font.family, "Calibri");
    assert_eq!(normal_font.size_pt, 11.0);
}

/// The pinned Excel-for-Mac footer probe behind issue #1210 uses a Segoe UI
/// 10pt Normal font whose colour is theme slot 3 (`dk2`). The header/footer
/// parser needs the resolved colour, not the unresolved stylesheet index.
#[test]
fn test_extract_normal_font_resolves_its_theme_colour() {
    let data = build_xlsx_with_normal_font("Segoe UI", 10.0);
    let data =
        rewrite_first_styles_font_with_color(&data, "Segoe UI", 10.0, r#"<color theme="3"/>"#);
    let book = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(&data), true)
        .expect("readable workbook");

    let normal_font =
        extract_normal_font(&data, Some(book.get_theme())).expect("styles.xml has a Normal font");
    assert_eq!(normal_font.family, "Segoe UI");
    assert_eq!(normal_font.size_pt, 10.0);
    assert_eq!(normal_font.color, Some(Color::new(0x44, 0x54, 0x6A)));
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

/// A worksheet with no print metadata follows the converter's A4 default.
///
/// The customer fixture carries neither `<pageSetup>` nor `<pageMargins>`.
/// Excel-for-Mac therefore uses its current paper selection, A4 on the
/// reference machine, instead of applying the OOXML default that belongs to
/// an initialised but silent `<pageSetup>` (issue #1382).
#[test]
fn test_pristine_worksheet_uses_the_converter_a4_default() {
    let data = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/xlsx/100-customers.xlsx"
    ));
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("the customer fixture should parse");

    assert!(!doc.pages.is_empty(), "the populated sheet should print");
    for page in &doc.pages {
        let Page::Sheet(sheet) = page else {
            panic!("an XLSX page should be a sheet page");
        };
        assert!(
            (sheet.size.width - 595.28).abs() < 0.01 && (sheet.size.height - 841.89).abs() < 0.01,
            "expected A4, got {:?}",
            sheet.size
        );
    }
}

/// A sheet whose `<pageMargins>` has initialised print state but which has no
/// `<pageSetup>` prints on the schema default, Letter (issue #717).
///
/// Confirmed against a native Excel export of this exact shape: the margin
/// element makes the absent paper setup an initialised default rather than the
/// fully pristine application-default case from issue #1382.
#[test]
fn test_page_margins_without_page_setup_default_to_letter() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Test")]);
    let mut archive = zip::ZipArchive::new(Cursor::new(&data)).expect("readable workbook");
    let mut worksheet_xml = String::new();
    std::io::Read::read_to_string(
        &mut archive
            .by_name("xl/worksheets/sheet1.xml")
            .expect("first worksheet part"),
        &mut worksheet_xml,
    )
    .expect("worksheet XML");
    assert!(worksheet_xml.contains("<pageMargins"));
    assert!(!worksheet_xml.contains("<pageSetup"));

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert!(
        (tp.size.width - 612.0).abs() < 0.01 && (tp.size.height - 792.0).abs() < 0.01,
        "expected Letter, got {:?}",
        tp.size
    );
}

/// Build a workbook whose only sheet declares `<pageSetup>` with an orientation
/// but no `paperSize`. This is the exact shape Excel writes for a sheet left on
/// the default paper, e.g. `<pageSetup orientation="portrait" .../>`.
fn build_xlsx_with_page_setup_lacking_paper_size(
    orientation: umya_spreadsheet::structs::OrientationValues,
) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Test");
        sheet.get_page_setup_mut().set_orientation(orientation);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// A `<pageSetup>` that omits `paperSize` resolves to the schema default,
/// Letter, rather than falling through to the renderer's A4 (issue #717).
///
/// A4 left such a workbook 16.7pt narrow and 49.9pt tall, which repaginated it:
/// the audited sheet collapsed from Excel's 8 printed pages to 6.
#[test]
fn test_omitted_paper_size_resolves_to_letter() {
    let data = build_xlsx_with_page_setup_lacking_paper_size(
        umya_spreadsheet::structs::OrientationValues::Portrait,
    );
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.size.width - 612.0).abs() < 0.01 && (page.size.height - 792.0).abs() < 0.01,
        "expected Letter, got {:?}",
        page.size
    );
}

/// The omitted-`paperSize` default is still just a default: a sheet that names
/// its paper keeps it. Triangulation against a hardcoded Letter.
#[test]
fn test_declared_a4_paper_size_survives_the_letter_default() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Test");
        // 9 = A4.
        sheet.get_page_setup_mut().set_paper_size(9);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let (doc, _warnings) = XlsxParser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.size.width - 595.28).abs() < 0.01 && (page.size.height - 841.89).abs() < 0.01,
        "expected A4, got {:?}",
        page.size
    );
}

/// The defaulted paper still rotates with `orientation="landscape"`, so the
/// default feeds the same path a declared code does.
#[test]
fn test_omitted_paper_size_still_honours_landscape() {
    let data = build_xlsx_with_page_setup_lacking_paper_size(
        umya_spreadsheet::structs::OrientationValues::Landscape,
    );
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&doc, 0);
    assert!(
        (page.size.width - 792.0).abs() < 0.01 && (page.size.height - 612.0).abs() < 0.01,
        "expected landscape Letter, got {:?}",
        page.size
    );
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

#[path = "xlsx_indent_tests.rs"]
mod indent_tests;

#[path = "xlsx_cell_inset_tests.rs"]
mod cell_inset_tests;

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
    rewrite_first_styles_font_with_color(data, family, size_pt, "")
}

/// Variant of [`rewrite_first_styles_font`] that includes one raw OOXML
/// `<color>` element in the first font.
fn rewrite_first_styles_font_with_color(
    data: &[u8],
    family: &str,
    size_pt: f64,
    color_xml: &str,
) -> Vec<u8> {
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
                format!("<font><sz val=\"{size_pt}\"/>{color_xml}<name val=\"{family}\"/></font>");
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

/// Replace every worksheet's `<sheetFormatPr .../>` with `replacement`. umya
/// models only `defaultRowHeight`, so a fixture that needs `customHeight` or a
/// hint umya would not write has to patch the part directly.
fn rewrite_sheet_format_pr(data: &[u8], replacement: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).expect("readable zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("readable entry");
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("readable entry body");
        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
            let xml = String::from_utf8(bytes).expect("worksheet part is utf-8");
            let start = xml
                .find("<sheetFormatPr")
                .expect("worksheet has sheetFormatPr");
            let end = xml[start..].find("/>").expect("sheetFormatPr is closed") + start + 2;
            bytes = format!("{}{}{}", &xml[..start], replacement, &xml[end..]).into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("writable entry");
        std::io::Write::write_all(&mut out, &bytes).expect("writable entry body");
    }
    out.finish().expect("finished zip").into_inner()
}

// ----- Drawing-only sheets (issue #620) -----

/// A workbook whose only sheet has no cells but carries one picture anchored
/// C1:F9 (cols 2..5, rows 0..8, zero offsets). umya cannot author drawings,
/// so the drawing parts are spliced into the zip it writes.
fn build_drawing_only_sheet_xlsx() -> Vec<u8> {
    let book = umya_spreadsheet::new_file();
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    splice_picture_drawing(&cursor.into_inner(), &one_pixel_png())
}

fn one_pixel_png() -> Vec<u8> {
    let image = image::DynamicImage::new_rgba8(1, 1);
    let mut encoded = Cursor::new(Vec::new());
    image
        .write_to(&mut encoded, image::ImageFormat::Png)
        .expect("1x1 PNG encodes");
    encoded.into_inner()
}

/// Splice `xl/drawings/drawing1.xml` (one twoCellAnchor picture), its rels,
/// and the supplied PNG media part into a workbook zip, wiring the first
/// worksheet to it.
fn splice_picture_drawing(data: &[u8], media: &[u8]) -> Vec<u8> {
    const DRAWING_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
    const DRAWING_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#;
    const SHEET_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).expect("readable zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut has_sheet_rels = false;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("readable entry");
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("readable entry body");
        match name.as_str() {
            "xl/worksheets/sheet1.xml" => {
                let xml = String::from_utf8(bytes).expect("sheet1.xml is utf-8");
                bytes = xml
                    .replace(
                        "</worksheet>",
                        r#"<drawing r:id="rIdDrawing1"/></worksheet>"#,
                    )
                    .into_bytes();
            }
            "xl/worksheets/_rels/sheet1.xml.rels" => {
                has_sheet_rels = true;
                let xml = String::from_utf8(bytes).expect("sheet rels is utf-8");
                bytes = xml
                    .replace(
                        "</Relationships>",
                        r#"<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
                    )
                    .into_bytes();
            }
            "[Content_Types].xml" => {
                let xml = String::from_utf8(bytes).expect("content types is utf-8");
                bytes = xml
                    .replace(
                        "</Types>",
                        r#"<Default Extension="png" ContentType="image/png"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"#,
                    )
                    .into_bytes();
            }
            _ => {}
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("writable entry");
        std::io::Write::write_all(&mut out, &bytes).expect("writable entry body");
    }
    if !has_sheet_rels {
        out.start_file(
            "xl/worksheets/_rels/sheet1.xml.rels",
            zip::write::FileOptions::default(),
        )
        .expect("writable sheet rels");
        std::io::Write::write_all(&mut out, SHEET_RELS.as_bytes()).expect("writable sheet rels");
    }
    for (path, body) in [
        ("xl/drawings/drawing1.xml", DRAWING_XML.as_bytes()),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            DRAWING_RELS.as_bytes(),
        ),
        ("xl/media/image1.png", media),
    ] {
        out.start_file(path, zip::write::FileOptions::default())
            .expect("writable drawing part");
        std::io::Write::write_all(&mut out, body).expect("writable drawing part body");
    }
    out.finish().expect("finished zip").into_inner()
}

#[test]
fn malformed_raster_image_is_omitted_with_warning_and_workbook_still_converts() {
    let workbook = build_xlsx_bytes("Sheet1", &[("A1", "survives")]);
    let data = splice_picture_drawing(&workbook, b"not a png");

    let (doc, warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("a malformed picture must not abort XLSX parsing");
    let page = get_sheet_page(&doc, 0);
    assert!(
        page.images.is_empty(),
        "the malformed picture must not reach the renderer"
    );
    assert!(
        warnings.iter().any(|warning| matches!(
            warning,
            ConvertWarning::UnsupportedElement { format, element }
                if format == "XLSX"
                    && element.contains("image omitted")
                    && element.contains("xl/media/image1.png")
        )),
        "expected an omission warning naming the malformed media part, got {warnings:?}"
    );

    let result = crate::convert_bytes(
        &data,
        crate::config::Format::Xlsx,
        &ConvertOptions::default(),
    )
    .expect("the workbook must still produce a PDF");
    assert!(result.pdf.starts_with(b"%PDF-"));
    assert!(result.warnings.iter().any(|warning| matches!(
        warning,
        ConvertWarning::UnsupportedElement { format, element }
            if format == "XLSX"
                && element.contains("image omitted")
                && element.contains("xl/media/image1.png")
    )));
}

/// A sheet with no cells must resolve its drawing anchors against the
/// workbook Normal font, producing the same column metric as a populated
/// sheet (issue #620). umya writes Calibri 11 as the Normal font, whose 6pt
/// unit prices an undeclared default column at 8 × 6 + 5 = 53pt (issue #621
/// probes); the legacy hardcoded 7px metric produced 44.2575pt.
#[test]
fn test_drawing_only_sheet_resolves_anchors_with_normal_font_metric() {
    let data = build_drawing_only_sheet_xlsx();
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&doc, 0);
    assert_eq!(
        page.images.len(),
        1,
        "the spliced picture must survive parse"
    );
    let image: &crate::ir::SheetImage = &page.images[0];

    let column_pt: f64 = default_column_width_pt(None, None, 6.0);
    assert_eq!(
        column_pt, 53.0,
        "Calibri-11 default column must be 53pt, got {column_pt}"
    );
    // Anchor spans cols 2..5 with zero offsets: x = 2 columns, width = 3.
    assert!(
        (image.x_offset_pt - 2.0 * column_pt).abs() < 0.01,
        "x_offset_pt {} != 2 x {column_pt}",
        image.x_offset_pt
    );
    let width: f64 = image.image.width.expect("twoCellAnchor resolves a width");
    assert!(
        (width - 3.0 * column_pt).abs() < 0.01,
        "width {width} != 3 x {column_pt}"
    );
    // Negative: neither the legacy hardcoded-7px metric (44.2575pt columns)
    // nor the pre-#621 8.43-character model (50.58pt) may resurface.
    for stale_column_pt in [44.2575_f64, 50.58_f64] {
        assert!(
            (width - 3.0 * stale_column_pt).abs() > 1.0,
            "width {width} still matches the stale {stale_column_pt}pt column metric"
        );
    }
}

// ── Overrunning `to` column offsets (issue #1149) ────────────────────

/// The width `anchored_image` resolves for the anchor of
/// `tests/fixtures/xlsx/issue_1066_blip_effect_picture.xlsx` — `from` col 1 at
/// 11880 EMU (0.935pt), `to` col 3 — on a sheet whose columns all measure the
/// requested width. Calibri 11 prices a character unit at 6pt, the same unit
/// that workbook's Calibri 12 Normal font resolves to, so a declared width of
/// `W` characters lands on the `6W` points Excel reported for it.
fn overrun_sweep_picture_width_pt(
    declared_width_chars: Option<f64>,
    base_column_width_chars: Option<u32>,
    to_col_off_emu: i64,
) -> f64 {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        if let Some(base_width) = base_column_width_chars {
            sheet
                .get_sheet_format_properties_mut()
                .set_base_column_width(base_width);
        }
        // Cover the whole span the anchor touches, `to` column included: an
        // undeclared column falls through to the default width instead.
        if let Some(width) = declared_width_chars {
            for col in 1..=4u32 {
                sheet
                    .get_column_dimension_by_number_mut(&col)
                    .set_width(width);
            }
        }
    }
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    let normal_font = NormalFont {
        family: "Calibri".to_string(),
        size_pt: 11.0,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };
    let ctx: SheetContext = empty_sheet_context(sheet, Some(&normal_font), None, None);
    let placed: crate::ir::SheetImage = anchored_image(
        xlsx_drawing::RawImageAnchor {
            from_row: 6,
            from_col: 1,
            from_col_off_emu: 11_880,
            from_row_off_emu: 0,
            to: Some((3, to_col_off_emu, 13, 0)),
            ext_emu: None,
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
        },
        sheet,
        &ctx,
    );
    placed.image.width.expect("a two-cell anchor sizes")
}

/// Sweeping the column width under a fixed 963720 EMU (75.883pt) `to` offset.
/// Excel adds the offset as written only from 78pt up — once the column is
/// wider than the offset; below that the picture comes out `3W - 1.052` wide,
/// the offset behaving as if it stopped just short of the column's far edge.
/// Every width here is `width of picture 1` read back through AppleScript from
/// an Excel for Mac export of a one-factor variant (issue #1149).
#[test]
fn the_to_column_offset_stops_growing_once_it_overruns_its_column() {
    // (declared `<col width>`, `baseColWidth`, column pt, Excel picture width)
    let measured: [(Option<f64>, Option<u32>, f64, f64); 9] = [
        (Some(5.0), None, 30.0, 88.948),
        (Some(8.0), None, 48.0, 142.948),
        (None, Some(10), 65.0, 193.948),
        (Some(12.0), None, 72.0, 214.948),
        (Some(12.5), None, 75.0, 223.948),
        (Some(13.0), None, 78.0, 230.948),
        (Some(16.0), None, 96.0, 266.948),
        (Some(20.0), None, 120.0, 314.948),
        (None, Some(20), 125.0, 324.948),
    ];
    for (declared, base_width, column_pt, expected) in measured {
        let width: f64 = overrun_sweep_picture_width_pt(declared, base_width, 963_720);
        assert!(
            (width - expected).abs() < 0.001,
            "{column_pt}pt columns: expected {expected}, got {width}"
        );
    }
}

/// Sweeping the offset instead, with the columns held at 65pt. The offset is
/// taken as written while it stays inside the column, and past that edge the
/// effective offset stays within half a point of the column width however far
/// the anchor overruns — 2000000 EMU is 157.480pt of offset and still prints
/// 194.545pt wide. Excel for Mac measurements again (issue #1149).
#[test]
fn an_overrunning_to_column_offset_keeps_only_the_fractional_overrun() {
    // (`xdr:colOff` EMU, Excel picture width over 65pt columns)
    let measured: [(i64, f64); 16] = [
        (0, 129.065),
        (400_000, 160.561),
        (800_000, 192.057),
        (823_000, 193.868),
        (825_500, 194.065),
        (830_000, 194.419),
        (840_000, 194.206),
        (850_000, 193.994),
        (860_000, 193.781),
        (870_000, 193.568),
        (880_000, 194.356),
        (900_000, 193.931),
        (963_720, 193.948),
        (1_100_000, 193.679),
        (1_500_000, 194.175),
        (2_000_000, 194.545),
    ];
    for (col_off_emu, expected) in measured {
        let width: f64 = overrun_sweep_picture_width_pt(None, Some(10), col_off_emu);
        assert!(
            (width - expected).abs() < 0.001,
            "colOff {col_off_emu}: expected {expected}, got {width}"
        );
    }
}

/// Triangulation for issue #620: the empty-sheet context must derive its
/// metric from whatever Normal font it is given — not a hardcoded value —
/// and fall back to the legacy 5.25pt unit only when no Normal font is
/// readable. The carried `normal_font` keeps the stub structurally
/// consistent with a populated-sheet context; nothing on the drawing-only
/// path reads it today (text boxes take their fonts from DrawingML run
/// properties and the theme).
#[test]
fn test_empty_sheet_context_derives_metric_from_normal_font() {
    let book = umya_spreadsheet::new_file();
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();

    let calibri_11 = NormalFont {
        family: "Calibri".to_string(),
        size_pt: 11.0,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };
    let calibri_ctx = empty_sheet_context(sheet, Some(&calibri_11), None, None);
    assert_eq!(resolve_column_unit_pt(sheet, Some(&calibri_11)), 6.0);
    assert_eq!(calibri_ctx.default_column_width_pt, 53.0);
    assert_eq!(calibri_ctx.normal_font, Some(calibri_11));

    // A smaller Normal font must shrink the metric with it:
    // round(0.506836 × 8pt) = 4pt unit → 8 × 4 + 5 = 37pt default columns
    // (issue #621 model).
    let calibri_8 = NormalFont {
        family: "Calibri".to_string(),
        size_pt: 8.0,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };
    assert_eq!(resolve_column_unit_pt(sheet, Some(&calibri_8)), 4.0);
    assert_eq!(
        empty_sheet_context(sheet, Some(&calibri_8), None, None).default_column_width_pt,
        37.0
    );

    // No readable Normal font: the shared cell-font fallback finds no cells
    // on an empty sheet and keeps the legacy 5.25pt unit (7px × 0.75); the
    // #621 probes never covered a stylesheet-less workbook.
    let fallback_ctx = empty_sheet_context(sheet, None, None, None);
    assert_eq!(resolve_column_unit_pt(sheet, None), 5.25);
    assert_eq!(fallback_ctx.default_column_width_pt, 8.0 * 5.25 + 5.0);
    assert_eq!(fallback_ctx.normal_font, None);
}

/// Rewrite the workbook's worksheet parts, inserting `insertion` before each
/// closing `</worksheet>` tag. umya's writer does not model
/// `printOptions@gridLines`, so the attribute is injected into the archive
/// the way Excel writes it — after `sheetData`.
fn inject_before_worksheet_close(xlsx: &[u8], insertion: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(xlsx.to_vec())).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).unwrap();
        let name: String = file.name().to_string();
        let mut contents: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut contents).unwrap();
        if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
            let text: String = String::from_utf8(contents).unwrap();
            contents = text
                .replace("</worksheet>", &format!("{insertion}</worksheet>"))
                .into_bytes();
        }
        writer
            .start_file(name, zip::write::FileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut writer, &contents).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn test_print_options_grid_lines_flags_the_sheet_table() {
    let plain = build_xlsx_bytes("Sheet1", &[("A1", "x"), ("B2", "y")]);
    let flagged = inject_before_worksheet_close(&plain, r#"<printOptions gridLines="1"/>"#);

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&flagged, &ConvertOptions::default()).unwrap();
    let table = &get_sheet_page(&doc, 0).table;
    assert!(
        table.prints_gridlines,
        "printOptions gridLines must set the table's gridline flag"
    );
    assert!(
        table.border_paint_model == TableBorderPaintModel::ExcelBoundaryBands,
        "the gridline flag rides on the boundary-band regime"
    );

    let (doc, _warnings) = parser.parse(&plain, &ConvertOptions::default()).unwrap();
    assert!(
        !get_sheet_page(&doc, 0).table.prints_gridlines,
        "a sheet without printOptions must not print gridlines"
    );
}

#[test]
fn test_print_options_horizontal_centered_flags_the_sheet_table() {
    let plain = build_xlsx_bytes("Sheet1", &[("A1", "x"), ("B2", "y")]);
    let flagged =
        inject_before_worksheet_close(&plain, r#"<printOptions horizontalCentered="1"/>"#);

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&flagged, &ConvertOptions::default()).unwrap();
    let table = &get_sheet_page(&doc, 0).table;
    assert!(
        table.centers_between_print_margins,
        "printOptions horizontalCentered must set the table's centering flag"
    );
    assert!(
        !table.prints_gridlines && !table.prints_headings,
        "centering alone must not turn gridlines or headings on"
    );

    let (doc, _warnings) = parser.parse(&plain, &ConvertOptions::default()).unwrap();
    assert!(
        !get_sheet_page(&doc, 0).table.centers_between_print_margins,
        "a sheet without printOptions must print flush to the left margin"
    );
}

#[test]
fn test_print_options_headings_prepends_gutter_column_and_letter_strip() {
    let plain = build_xlsx_bytes("Sheet1", &[("A1", "x"), ("B2", "y")]);
    let flagged = inject_before_worksheet_close(&plain, r#"<printOptions headings="1"/>"#);

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&flagged, &ConvertOptions::default()).unwrap();
    let table = &get_sheet_page(&doc, 0).table;
    assert!(
        table.prints_headings,
        "printOptions headings must set the table's heading flag"
    );
    // GT geometry (issue #623): 23pt gutter track, 13pt letter-strip track.
    assert_eq!(table.column_widths[0], 23.0);
    assert_eq!(table.rows[0].height, Some(13.0));

    // Strip row: empty corner + letters covering the printed columns.
    let strip = &table.rows[0];
    assert!(strip.cells[0].content.is_empty());
    assert_eq!(cell_text(&strip.cells[1]), "A");
    assert_eq!(cell_text(&strip.cells[2]), "B");

    // Gutter cells carry the sheet row numbers; data follows one column later.
    assert_eq!(cell_text(&table.rows[1].cells[0]), "1");
    assert_eq!(cell_text(&table.rows[2].cells[0]), "2");
    assert_eq!(cell_text(&table.rows[1].cells[1]), "x");
    assert_eq!(cell_text(&table.rows[2].cells[2]), "y");

    let (doc, _warnings) = parser.parse(&plain, &ConvertOptions::default()).unwrap();
    let table = &get_sheet_page(&doc, 0).table;
    assert!(
        !table.prints_headings,
        "a sheet without printOptions must not print headings"
    );
    assert_eq!(
        cell_text(&table.rows[0].cells[0]),
        "x",
        "an unflagged sheet must keep its grid unshifted"
    );
}

#[test]
fn test_print_headings_row_numbers_continue_across_manual_page_breaks() {
    // A row break after row 2 splits the sheet; the second segment's gutter
    // must continue at the actual sheet row number, not restart at 1.
    let plain = {
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.set_name("Sheet1");
            for (coord, value) in [("A1", "r1"), ("A2", "r2"), ("A3", "r3")] {
                sheet.get_cell_mut(coord).set_value(value);
            }
            let mut brk = umya_spreadsheet::Break::default();
            brk.set_id(2);
            brk.set_manual_page_break(true);
            sheet.get_row_breaks_mut().add_break_list(brk);
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    };
    let flagged = inject_before_worksheet_close(&plain, r#"<printOptions headings="1"/>"#);

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&flagged, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.pages.len(), 2);

    let first = &get_sheet_page(&doc, 0).table;
    assert!(first.prints_headings);
    assert_eq!(cell_text(&first.rows[1].cells[0]), "1");
    assert_eq!(cell_text(&first.rows[2].cells[0]), "2");

    let second = &get_sheet_page(&doc, 1).table;
    assert!(second.prints_headings);
    assert_eq!(cell_text(&second.rows[0].cells[1]), "A");
    assert_eq!(cell_text(&second.rows[1].cells[0]), "3");
    assert_eq!(cell_text(&second.rows[1].cells[1]), "r3");
}

/// A sheet with no used cells can still declare `<cols>`, and a drawing
/// anchored to those columns is placed against their widths (issue #714).
///
/// The window used to be left empty, so every column fell through to the
/// default and an anchored picture was priced at the 8.43-char width whatever
/// the file declared.
#[test]
fn test_empty_sheet_context_reads_declared_column_widths() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        for col in 1..=3u32 {
            sheet
                .get_column_dimension_by_number_mut(&col)
                .set_width(20.0);
        }
    }
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    let calibri_11 = NormalFont {
        family: "Calibri".to_string(),
        size_pt: 11.0,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    };

    let ctx = empty_sheet_context(sheet, Some(&calibri_11), None, None);

    assert_eq!((ctx.col_start, ctx.col_end), (1, 3));
    assert_eq!(ctx.column_widths.len(), 3);
    for width in &ctx.column_widths {
        assert!(
            *width > ctx.default_column_width_pt,
            "a declared width of 20 is wider than the 8.43-char default, got \
             {width} against {}",
            ctx.default_column_width_pt
        );
    }
}

/// A sheet declaring no `<cols>` at all keeps the empty window, so every
/// column still falls through to the default width.
#[test]
fn test_empty_sheet_context_without_cols_keeps_the_default_window() {
    let book = umya_spreadsheet::new_file();
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();

    let ctx = empty_sheet_context(sheet, None, None, None);

    assert!(ctx.column_widths.is_empty());
    assert_eq!(ctx.num_cols, 0);
    assert!(ctx.col_start > ctx.col_end, "an empty window");
}

// ── Dimension-less row heights (issues #715, #1047) ──────────────────

/// A Normal font declared at `size_pt` that resolves through the theme's
/// minor font scheme, as every Excel-authored stylesheet writes it.
fn theme_scheme_normal_font(size_pt: f64) -> NormalFont {
    NormalFont {
        family: "Calibri".to_string(),
        size_pt,
        color: None,
        uses_theme_scheme: true,
        theme_declares_script_faces: true,
    }
}

/// Excel recomputes the default row height from the Normal font for rows
/// that record no height, ignoring the `defaultRowHeight` hint unless the
/// sheet marks it `customHeight` (issue #715).
#[test]
fn a_dimensionless_row_takes_the_normal_font_default() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_sheet_format_properties_mut()
        .set_default_row_height(15.0);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(sheet, Some(&theme_scheme_normal_font(11.0))),
        17.0
    );
}

/// The declared hint is not a floor or a starting point — it is ignored.
/// Probe-measured (issue #1047): declaring 9, 15, 20 or 30 exports the same
/// 17pt rows, so a hint that disagrees with the font changes nothing.
#[test]
fn a_declared_default_the_normal_font_disagrees_with_is_ignored() {
    for declared in [9.0, 20.0, 30.0] {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(declared);
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&theme_scheme_normal_font(11.0))
            ),
            17.0,
            "declared {declared} must not reach the recomputed default"
        );
    }
}

#[test]
fn a_custom_height_default_is_honoured_as_declared() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    let properties = sheet.get_sheet_format_properties_mut();
    properties.set_default_row_height(30.0);
    properties.set_custom_height(true);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(sheet, Some(&theme_scheme_normal_font(11.0))),
        30.0
    );
}

/// Every size measured for a Normal font that resolves through the theme's
/// per-script face list. The customer workbooks in issue #1226 exercise 12pt,
/// but the native sweep measured one complete column, so the lookup must not
/// leave the other seven added sizes on the declared hint either.
#[test]
fn the_theme_scheme_recomputes_from_the_full_ui_script_face_series() {
    let sizes = [
        8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 20.0, 22.0, 24.0,
    ];
    let heights = [
        13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 26.0, 27.0, 30.0, 32.0, 35.0,
    ];
    for (size_pt, expected) in sizes.into_iter().zip(heights) {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(40.0);
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&theme_scheme_normal_font(size_pt))
            ),
            expected,
            "{size_pt}pt Normal font"
        );
    }
}

/// A Normal font that names its own face instead of deferring to the theme
/// scheme is a different measurement, not an absent one: Excel substitutes
/// unavailable Calibri and Aptos independently on the reference machine and
/// recomputes against the resolved face, ignoring the declared hint exactly
/// as the scheme path does (issues #1102, #1225).
fn substituted_face_normal_font(family: &str, size_pt: f64) -> NormalFont {
    NormalFont {
        family: family.to_string(),
        size_pt,
        color: None,
        uses_theme_scheme: false,
        theme_declares_script_faces: false,
    }
}

#[test]
fn a_scheme_less_substituted_face_recomputes_over_the_declared_default() {
    for declared in [15.0, 18.0, 30.0] {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(declared);
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&substituted_face_normal_font("Calibri", 12.0))
            ),
            16.0,
            "declared {declared} must not reach a 12pt Calibri row"
        );
    }
}

/// Every size the scheme-less sweep measured for Calibri and Aptos. The two
/// agree at nine sizes but separate at 9, 13, 16, 20 and 24pt, so neither may
/// borrow the other face's series (issue #1225).
#[test]
fn calibri_and_aptos_recompute_from_their_own_full_size_series() {
    let sizes = [
        8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 20.0, 22.0, 24.0,
    ];
    for (family, heights) in [
        (
            "Calibri",
            [
                11.0, 12.0, 14.0, 15.0, 16.0, 17.0, 19.0, 20.0, 21.0, 23.0, 24.0, 26.0, 29.0, 31.0,
            ],
        ),
        (
            "Aptos",
            [
                11.0, 13.0, 14.0, 15.0, 16.0, 18.0, 19.0, 20.0, 22.0, 23.0, 24.0, 27.0, 29.0, 32.0,
            ],
        ),
    ] {
        for (size_pt, expected) in sizes.iter().zip(heights) {
            let mut book = umya_spreadsheet::new_file();
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet
                .get_sheet_format_properties_mut()
                .set_default_row_height(18.0);
            assert_eq!(
                xlsx_cells::worksheet_default_row_height_pt(
                    sheet,
                    Some(&substituted_face_normal_font(family, *size_pt))
                ),
                expected,
                "{size_pt}pt {family} Normal font"
            );
        }
    }
}

/// A family no sweep has measured keeps the declared hint — as does a size
/// its family's series skips, and a workbook with no readable stylesheet
/// (issues #1047, #1150).
#[test]
fn an_unmeasured_normal_font_keeps_the_declared_default() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_sheet_format_properties_mut()
        .set_default_row_height(15.0);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(sheet, Some(&theme_scheme_normal_font(19.0))),
        15.0,
        "a size between the theme scheme's measured points keeps the declared hint"
    );
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("Wingdings", 12.0))
        ),
        15.0
    );
    for unmeasured_variant in ["Calibri Light", "Aptos Narrow", "Aptos Display"] {
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&substituted_face_normal_font(unmeasured_variant, 13.0))
            ),
            15.0,
            "{unmeasured_variant} does not borrow the exact Calibri/Aptos recompute"
        );
    }
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("Calibri", 19.0))
        ),
        15.0
    );
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("Arial", 10.5))
        ),
        15.0,
        "a size between Arial's measured points has no series entry"
    );
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(sheet, None),
        15.0
    );
}

/// Every family the issue #1150 sweep measured, at every size it measured.
///
/// One `xl/styles.xml` `<font>` per variant over
/// `issue_1066_blip_effect_picture.xlsx`, reading `standard height of
/// worksheet 1` back through AppleScript
/// (`scripts/measure_excel_row_height.py`). Whole columns, not spot checks:
/// no family here follows another's curve, so a size sampled from one proves
/// nothing about the rest.
#[test]
fn every_measured_face_recomputes_its_own_series() {
    let measured: [(&str, [f64; 14]); 8] = [
        (
            "Arial",
            [
                11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0,
            ],
        ),
        (
            "Times New Roman",
            [
                11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0,
            ],
        ),
        (
            "Verdana",
            [
                11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0,
            ],
        ),
        (
            "Tahoma",
            [
                11.0, 12.0, 13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0,
            ],
        ),
        (
            "Georgia",
            [
                11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 21.0, 22.0, 23.0, 25.0, 28.0, 30.0,
            ],
        ),
        (
            "Helvetica",
            [
                11.0, 12.0, 13.0, 15.0, 16.0, 17.0, 18.0, 19.0, 21.0, 22.0, 23.0, 26.0, 28.0, 31.0,
            ],
        ),
        (
            "Courier New",
            [
                11.0, 13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 21.0, 22.0, 23.0, 24.0, 27.0, 30.0, 32.0,
            ],
        ),
        (
            "Segoe UI",
            [
                11.0, 13.0, 14.0, 16.0, 16.0, 20.0, 21.0, 23.0, 23.0, 25.0, 26.0, 28.0, 31.0, 33.0,
            ],
        ),
    ];
    let sizes: [f64; 14] = [
        8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 20.0, 22.0, 24.0,
    ];
    for (family, heights) in measured {
        for (size_pt, expected) in sizes.iter().zip(heights) {
            let mut book = umya_spreadsheet::new_file();
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet
                .get_sheet_format_properties_mut()
                .set_default_row_height(18.0);
            assert_eq!(
                xlsx_cells::worksheet_default_row_height_pt(
                    sheet,
                    Some(&substituted_face_normal_font(family, *size_pt))
                ),
                expected,
                "{size_pt}pt {family} Normal font"
            );
        }
    }
}

/// The Korean faces the same sweep measured. `나눔명조` answered Helvetica's
/// column at all fourteen sizes, and `맑은 고딕` the theme scheme's — the
/// reading that the scheme resolves to that face on the reference machine
/// (issues #1047, #1150).
#[test]
fn a_normal_font_naming_a_korean_face_recomputes_its_measured_series() {
    for (family, size_pt, expected) in [
        ("나눔명조", 11.0, 15.0),
        ("나눔명조", 12.0, 16.0),
        ("나눔명조", 16.0, 21.0),
        ("맑은 고딕", 10.0, 15.0),
        ("맑은 고딕", 11.0, 17.0),
        ("Malgun Gothic", 14.0, 20.0),
        ("Malgun Gothic", 18.0, 27.0),
    ] {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(18.0);
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&substituted_face_normal_font(family, size_pt))
            ),
            expected,
            "{size_pt}pt {family} Normal font"
        );
    }
}

/// Two spellings of one family are aliased only where the sweep read them as
/// one column. `Malgun Gothic` answers `맑은 고딕` at all fourteen sizes;
/// `NanumMyeongjo` answers a column of its own — 16 at 12pt against 15 at
/// 13pt, reproduced across two runs — so it keeps the declared hint instead
/// of borrowing `나눔명조`'s series (issue #1150).
#[test]
fn an_ascii_spelling_is_aliased_only_where_it_was_measured() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_sheet_format_properties_mut()
        .set_default_row_height(15.0);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("NanumMyeongjo", 12.0))
        ),
        15.0
    );
}

/// A family name is matched whole and case-insensitively, never as a prefix:
/// `Arial Narrow` and `Arial Black` are different faces with row heights of
/// their own, and nothing has measured them.
#[test]
fn a_measured_family_matches_the_whole_name_only() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_sheet_format_properties_mut()
        .set_default_row_height(15.0);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("arial", 10.0))
        ),
        13.0,
        "the corpus writes the family lower-cased too"
    );
    for unmeasured in ["Arial Narrow", "Arial Black", "Arial Unicode MS"] {
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&substituted_face_normal_font(unmeasured, 10.0))
            ),
            15.0,
            "{unmeasured} is not Arial"
        );
    }
}

/// The declared hint loses to a measured family exactly as it loses to
/// Calibri, and `customHeight` restores it exactly as it does there
/// (issue #1150).
#[test]
fn a_measured_family_overrides_the_declared_default_unless_custom() {
    for declared in [12.75, 15.0, 18.0, 30.0] {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(declared);
        assert_eq!(
            xlsx_cells::worksheet_default_row_height_pt(
                sheet,
                Some(&substituted_face_normal_font("Arial", 10.0))
            ),
            13.0,
            "declared {declared} must not reach a 10pt Arial row"
        );
    }
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    let properties = sheet.get_sheet_format_properties_mut();
    properties.set_default_row_height(30.0);
    properties.set_custom_height(true);
    assert_eq!(
        xlsx_cells::worksheet_default_row_height_pt(
            sheet,
            Some(&substituted_face_normal_font("Arial", 10.0))
        ),
        30.0
    );
}

/// The printed grid recomputes exactly as anchors do, and prints the
/// recomputed height unscaled — the `native_excel_pdf_row_height`
/// compaction is calibrated for declared heights, not for this one
/// (issue #1047: `Formatting.xlsx` exports 17pt rows against our 14pt).
#[test]
fn a_dimensionless_printed_row_takes_the_recomputed_default() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "one"), ("A2", "two")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let table = &get_sheet_page(&doc, 0).table;
    for row in &table.rows {
        assert_eq!(row.height, Some(17.0));
    }
}

#[test]
fn a_dimensionless_printed_row_ignores_a_disagreeing_declared_default() {
    let data = rewrite_sheet_format_pr(
        &build_xlsx_bytes("Sheet1", &[("A1", "one"), ("A2", "two")]),
        r#"<sheetFormatPr defaultRowHeight="30"/>"#,
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let table = &get_sheet_page(&doc, 0).table;
    for row in &table.rows {
        assert_eq!(row.height, Some(17.0));
    }
}

// ── Auto-sized rows take their own cells' font (issue #1140) ─────────

/// A cell that carries an explicit font, so the row it sits in is auto-sized
/// against that font rather than against the workbook's Normal one.
fn sheet_with_one_cell_font(row_idx: u32, size_pt: f64) -> umya_spreadsheet::Spreadsheet {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet
        .get_sheet_format_properties_mut()
        .set_default_row_height(15.0);
    let cell = sheet.get_cell_mut((1u32, row_idx));
    cell.set_value("가나다");
    cell.get_style_mut().get_font_mut().set_size(size_pt);
    book
}

/// A row that records no `ht` is auto-sized, and Excel sizes it from the
/// tallest font its own cells carry — not from the workbook's Normal font.
///
/// Measured on `issue_1060_sheet_row_line_box_probe.xlsx`, whose Normal font
/// is a theme-scheme Calibri 11 that none of its cells use. Native
/// Excel-for-Mac export; AppleScript `row height of row N` and the
/// `mutool draw -F trace` baseline pitch agree (issue #1140):
///
/// | rows | cell font | Excel row |
/// | --- | --- | ---: |
/// | 1-6, 8-13 | Malgun Gothic 14 | 20.00pt |
/// | 15-17 | Malgun Gothic 24 | 35.00pt |
#[test]
fn an_auto_row_is_sized_by_its_own_cells_tallest_font() {
    for (cell_size_pt, expected) in [(14.0, 20.0), (18.0, 27.0), (24.0, 35.0)] {
        let book = sheet_with_one_cell_font(1, cell_size_pt);
        let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
        assert_eq!(
            xlsx_cells::printed_grid_row_height_pt(
                sheet,
                1,
                Some(&theme_scheme_normal_font(11.0)),
                None,
            ),
            expected,
            "a row of {cell_size_pt}pt cells"
        );
    }
}

/// Only the row's own cells count. A tall cell one row down leaves this row
/// on the sheet's recomputed default.
#[test]
fn an_auto_row_ignores_another_rows_cell_font() {
    let book = sheet_with_one_cell_font(2, 24.0);
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    assert_eq!(
        xlsx_cells::printed_grid_row_height_pt(
            sheet,
            1,
            Some(&theme_scheme_normal_font(11.0)),
            None,
        ),
        17.0
    );
}

/// The row-cell term only ever raises the track: what Excel gives a row whose
/// cells are *smaller* than the Normal font is unmeasured, so such a row keeps
/// the sheet's recomputed default rather than shrinking on a guess.
#[test]
fn an_auto_row_of_smaller_cells_keeps_the_recomputed_default() {
    let book = sheet_with_one_cell_font(1, 8.0);
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    assert_eq!(
        xlsx_cells::printed_grid_row_height_pt(
            sheet,
            1,
            Some(&theme_scheme_normal_font(11.0)),
            None,
        ),
        17.0
    );
}

/// A size the face's series does not measure keeps the recomputed default
/// too — the measured points are too irregular to interpolate between.
#[test]
fn an_auto_row_whose_cell_size_is_unmeasured_keeps_the_recomputed_default() {
    let book = sheet_with_one_cell_font(1, 19.0);
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    assert_eq!(
        xlsx_cells::printed_grid_row_height_pt(
            sheet,
            1,
            Some(&theme_scheme_normal_font(11.0)),
            None,
        ),
        17.0
    );
}

/// A recorded `ht` is the row's current worksheet height whatever its cells
/// hold, so it still outranks the auto-size recompute.
#[test]
fn a_recorded_row_height_outranks_its_cells_font() {
    let mut book = sheet_with_one_cell_font(1, 24.0);
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.get_row_dimension_mut(&1).set_height(36.0);
    let sheet: &umya_spreadsheet::Worksheet = book.get_sheet(&0).unwrap();
    assert_eq!(
        xlsx_cells::printed_grid_row_height_pt(
            sheet,
            1,
            Some(&theme_scheme_normal_font(11.0)),
            None,
        ),
        36.0
    );
}

/// The whole probe workbook end to end: its twelve 14pt auto rows print
/// Excel's 20pt track and its three 24pt ones Excel's 35pt track, where every
/// one of them used to print the Normal font's 17pt (issue #1140).
#[test]
fn the_probe_workbooks_auto_rows_print_their_own_cell_tracks() {
    let heights: Vec<Option<f64>> = printed_row_heights_over_all_pages(THEME_SCHEME_PROBE);

    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(20.0))
            .count(),
        12,
        "12 auto rows of Malgun Gothic 14 print a 20pt track, got {heights:?}"
    );
    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(35.0))
            .count(),
        3,
        "3 auto rows of Malgun Gothic 24 print a 35pt track, got {heights:?}"
    );
    assert!(
        !heights.contains(&Some(17.0)),
        "no row keeps the Normal font's own track, got {heights:?}"
    );
}

/// `customHeight` marks the declared default as user-set, so it stays in play
/// as a declared height and maps like one: 30 → 28 under a Normal font whose
/// grid compacts (issue #1047).
///
/// The Normal font has to be a compacting one for that to be the answer. The
/// same `defaultRowHeight="30" customHeight="1"` over umya's own scheme
/// Calibri 11 and Office theme exports 30.00pt tracks natively, one factor
/// against a byte-identical re-zip control — no compaction at all (#1094).
#[test]
fn a_custom_height_printed_default_maps_through_the_grid() {
    let data = rewrite_first_styles_font(
        &rewrite_sheet_format_pr(
            &build_xlsx_bytes("Sheet1", &[("A1", "one"), ("A2", "two")]),
            r#"<sheetFormatPr defaultRowHeight="30" customHeight="1"/>"#,
        ),
        "Calibri",
        11.0,
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let table = &get_sheet_page(&doc, 0).table;
    for row in &table.rows {
        assert_eq!(row.height, Some(28.0));
    }
}

/// A scheme-less Normal font recomputes too, and its recompute is the
/// *worksheet* height rather than the printed one, so it still maps through
/// the compaction: `functions-excel-2010.xlsx` (Calibri 11, no `scheme`,
/// `defaultRowHeight="15"`) exports 14pt rows, and 11pt Calibri recomputes to
/// the same 15 its sheet declares.
#[test]
fn a_scheme_less_dimensionless_printed_row_compacts_the_recomputed_default() {
    let data = rewrite_sheet_format_pr(
        &build_xlsx_with_normal_font("Calibri", 11.0),
        r#"<sheetFormatPr defaultRowHeight="15"/>"#,
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let table = &get_sheet_page(&doc, 0).table;
    assert_eq!(table.rows[0].height, Some(14.0));
}

/// And where the recompute disagrees with the hint, the hint is gone: 12pt
/// Calibri recomputes a 16pt worksheet row whatever `defaultRowHeight` says,
/// which its compacting grid prints as `round(16 x 0.92) = 15` — the track
/// `issue_1066_blip_effect_picture.xlsx`'s native export measures, against
/// the 17 the declared 18 would have given (issue #1102).
#[test]
fn a_scheme_less_dimensionless_printed_row_ignores_a_disagreeing_declared_default() {
    let data = rewrite_sheet_format_pr(
        &build_xlsx_with_normal_font("Calibri", 12.0),
        r#"<sheetFormatPr defaultRowHeight="18"/>"#,
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let table = &get_sheet_page(&doc, 0).table;
    assert_eq!(table.rows[0].height, Some(15.0));
}

/// Aptos diverges from Calibri at five sizes in the issue #1225 worksheet
/// recompute sweep. Each recomputed height then maps through the separately
/// measured Calibri/Aptos printed-grid rule. The 13pt case is also the visual
/// probe: 18pt in the worksheet becomes a 17pt printed track.
#[test]
fn aptos_dimensionless_rows_print_from_the_aptos_recompute() {
    for (size_pt, expected_printed) in [
        (9.0, 13.0),
        (13.0, 17.0),
        (16.0, 20.0),
        (20.0, 25.0),
        (24.0, 29.0),
    ] {
        let data = rewrite_sheet_format_pr(
            &build_xlsx_with_normal_font("Aptos", size_pt),
            r#"<sheetFormatPr defaultRowHeight="18"/>"#,
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(
            get_sheet_page(&doc, 0).table.rows[0].height,
            Some(expected_printed),
            "{size_pt}pt Aptos Normal font"
        );
    }
}

/// Printed-grid compaction belongs to the resolved face, not to the
/// Calibri/Aptos substitution set. Native Excel-for-Mac exports of
/// `issue_1066_blip_effect_picture.xlsx` measure a 15pt printed track for
/// both Arial 12 (16pt worksheet row) and Courier New 12 (17pt worksheet
/// row), while Verdana 12 keeps its 16pt worksheet row whole (issue #1224).
#[test]
fn named_faces_map_dimensionless_rows_through_their_own_printed_grid() {
    for (family, expected) in [("Arial", 15.0), ("Courier New", 15.0), ("Verdana", 16.0)] {
        let data = rewrite_sheet_format_pr(
            &build_xlsx_with_normal_font(family, 12.0),
            r#"<sheetFormatPr defaultRowHeight="18"/>"#,
        );
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(
            get_sheet_page(&doc, 0).table.rows[0].height,
            Some(expected),
            "{family} 12"
        );
    }
}

// ── Declared row tracks and the Normal font (issue #1068) ────────────

/// A workbook whose Normal font names `family` at `size_pt` outright, with
/// one `customHeight` row per entry of `heights`.
fn build_xlsx_with_normal_font_and_row_heights(
    family: &str,
    size_pt: f64,
    heights: &[f64],
) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        for (index, height) in heights.iter().enumerate() {
            let row: u32 = index as u32 + 1;
            sheet
                .get_cell_mut(format!("A{row}").as_str())
                .set_value("x");
            let dimension = sheet.get_row_dimension_mut(&row);
            dimension.set_height(*height);
            dimension.set_custom_height(true);
        }
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    rewrite_first_styles_font(&cursor.into_inner(), family, size_pt)
}

fn printed_row_heights(data: &[u8]) -> Vec<Option<f64>> {
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    get_sheet_page(&doc, 0)
        .table
        .rows
        .iter()
        .map(|row| row.height)
        .collect()
}

/// A Normal font the machine renders itself leaves the printed grid alone:
/// the row prints its declared height, truncated to the whole point Excel's
/// grid lands on. Probe-measured on the reported workbook (Segoe UI 10
/// Normal), one factor per export: 30 -> 30, 25.5 -> 25, 49.5 -> 49
/// (issue #1068).
#[test]
fn a_kept_normal_font_prints_a_declared_row_height_whole() {
    assert_eq!(
        printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
            "Segoe UI",
            10.0,
            &[30.0, 25.5, 49.5]
        )),
        vec![Some(30.0), Some(25.0), Some(49.0)]
    );
}

/// The same declared heights through a Normal font Excel remaps compact
/// instead: 30 -> 28, 25.5 -> 23, 49.5 -> 46 (probe-measured, Calibri 11).
#[test]
fn a_remapped_normal_font_compacts_a_declared_row_height() {
    assert_eq!(
        printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
            "Calibri",
            11.0,
            &[30.0, 25.5, 49.5]
        )),
        vec![Some(28.0), Some(23.0), Some(46.0)]
    );
}

/// The native issue #1224 face sweep gives a different whole-point series per
/// face. Every point below was exported independently from three equal fixed
/// rows, with the PDF baseline pitch agreeing across both row pairs.
#[test]
fn named_faces_map_declared_rows_through_their_own_printed_grid() {
    let heights = [
        12.0, 15.0, 16.0, 17.0, 18.0, 20.0, 25.5, 30.0, 36.0, 40.0, 49.5,
    ];
    for (family, expected) in [
        (
            "Arial",
            vec![
                11.0, 14.0, 15.0, 16.0, 17.0, 18.0, 24.0, 28.0, 33.0, 37.0, 46.0,
            ],
        ),
        (
            "Courier New",
            vec![
                10.0, 13.0, 14.0, 15.0, 16.0, 17.0, 22.0, 26.0, 31.0, 35.0, 43.0,
            ],
        ),
        (
            "Verdana",
            vec![
                12.0, 15.0, 16.0, 17.0, 18.0, 20.0, 25.0, 30.0, 36.0, 40.0, 49.0,
            ],
        ),
        (
            "맑은 고딕",
            vec![
                12.0, 15.0, 16.0, 17.0, 18.0, 20.0, 25.0, 30.0, 36.0, 40.0, 49.0,
            ],
        ),
    ] {
        assert_eq!(
            printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
                family, 12.0, &heights
            )),
            expected.into_iter().map(Some).collect::<Vec<_>>(),
            "{family} 12"
        );
    }
}

/// The half-point Arial sweep disproves a scale even within one face: these
/// adjacent steps cannot all be derived by multiplying and rounding. Keep the
/// measured staircase explicit (issue #1224).
#[test]
fn arial_12_uses_its_measured_half_point_staircase() {
    assert_eq!(
        printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
            "Arial",
            12.0,
            &[10.5, 11.5, 19.0, 19.5, 24.0]
        )),
        vec![Some(9.0), Some(10.0), Some(17.0), Some(18.0), Some(22.0)]
    );
}

/// Do not turn the measured table into interpolation or a family-prefix
/// guess. An unmeasured height, size, or distinct Arial face keeps the
/// conservative whole-point path until a native sweep covers it.
#[test]
fn named_face_printed_grid_measurements_are_not_extrapolated() {
    for (family, size, height, expected) in [
        ("Arial", 12.0, 24.25, 24.0),
        ("Arial", 11.0, 16.0, 16.0),
        ("Arial Narrow", 12.0, 16.0, 16.0),
        ("Courier New", 12.0, 16.5, 16.0),
    ] {
        assert_eq!(
            printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
                family,
                size,
                &[height]
            )),
            vec![Some(expected)],
            "{family} {size} at {height}pt"
        );
    }

    let scheme_font = NormalFont {
        family: "Arial".to_string(),
        size_pt: 12.0,
        color: None,
        uses_theme_scheme: true,
        theme_declares_script_faces: false,
    };
    assert_eq!(
        xlsx_cells::native_excel_pdf_row_height(16.0, Some(&scheme_font)),
        16.0,
        "a scheme font does not name Arial outright"
    );
}

fn table_bottom_aligned_descent_floor_pt(data: &[u8]) -> f64 {
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    get_sheet_page(&doc, 0)
        .table
        .bottom_aligned_descent_floor_pt
}

/// The script-face theme family in the native floor probe matrix holds a
/// bottom-aligned cell's baseline 4pt clear of the row boundary
/// (issue #1097). This test does not generalize that floor to every face whose
/// row track remains whole.
#[test]
fn a_kept_normal_font_floors_the_bottom_aligned_seat_at_four_points() {
    assert_eq!(
        table_bottom_aligned_descent_floor_pt(&build_xlsx_with_normal_font_and_row_heights(
            "Segoe UI",
            10.0,
            &[40.0]
        )),
        4.0
    );
}

/// The remapped Calibri/Aptos family in that floor probe matrix seats it a
/// point lower, measured over an eleven-size sweep of a ruled native
/// re-export (issue #1199).
#[test]
fn a_remapped_normal_font_floors_the_bottom_aligned_seat_at_three_points() {
    assert_eq!(
        table_bottom_aligned_descent_floor_pt(&build_xlsx_with_normal_font_and_row_heights(
            "Calibri",
            11.0,
            &[40.0]
        )),
        3.0
    );
}

/// The compaction belongs to the face Excel substitutes, not to the family
/// name: the same Calibri below 11pt prints whole (probe: ht=40 exports a
/// 40pt track at 9pt and 10pt, a 37pt one at 11pt).
#[test]
fn a_remapped_family_below_the_measured_size_prints_whole() {
    assert_eq!(
        printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
            "Calibri",
            10.0,
            &[40.0]
        )),
        vec![Some(40.0)]
    );
}

/// Aptos is the other face the standard theme resolves to, and it compacts
/// exactly as Calibri does (probe: ht=40 exports a 37pt track for both
/// "Aptos" and "Aptos Narrow" at 11pt).
#[test]
fn the_other_remapped_family_compacts_too() {
    assert_eq!(
        printed_row_heights(&build_xlsx_with_normal_font_and_row_heights(
            "Aptos Narrow",
            11.0,
            &[40.0]
        )),
        vec![Some(37.0)]
    );
}

// ── Thick row-border tracks (issue #1228) ────────────────────────────

/// The customer workbook shared by the print-range and thick-row probes.
const SH107_FORMATTED_TABLE: &[u8] =
    include_bytes!("../../../../tests/fixtures/xlsx/SH107-9-x-9-Formatted-Table.xlsx");

/// SH107 declares A1:K9 and writes quote-prefix-only J/K cells in two rows,
/// but its native Excel-for-Mac export stops at the nine value-bearing columns
/// A:I. The two value-less columns paint no fill, border or text and therefore
/// must not claim a second horizontal page (issue #1229).
#[test]
fn value_less_unpainted_trailing_cells_do_not_extend_the_printed_grid() {
    let (doc, _warnings) = XlsxParser
        .parse(SH107_FORMATTED_TABLE, &ConvertOptions::default())
        .unwrap();

    assert_eq!(doc.pages.len(), 1, "the native export is one page");
    assert_eq!(
        get_sheet_page(&doc, 0).table.column_widths.len(),
        9,
        "the printed grid stops at value-bearing column I"
    );
}

/// SH107's I1 has a pale-yellow direct fill and a regular theme-dark cell
/// font, while its `TableStyleMedium2` header prints the `Col9` run white and
/// bold. Excel composes those two sources: the direct fill stays yellow, but
/// the table header's font colour still reaches the run (issue #1230).
#[test]
fn a_table_header_text_color_overrides_the_cell_xf_font_color() {
    let (doc, _warnings) = XlsxParser
        .parse(SH107_FORMATTED_TABLE, &ConvertOptions::default())
        .unwrap();
    let cell = &get_sheet_page(&doc, 0).table.rows[0].cells[8];

    assert_eq!(cell_text(cell), "Col9");
    assert_eq!(cell.background, Some(Color::new(255, 255, 204)));
    assert_eq!(first_run_style(cell).color, Some(Color::white()));
    assert_eq!(first_run_style(cell).bold, Some(true));
}

/// A non-default direct font colour is a real override, unlike the un-tinted
/// theme-dark default copied into SH107's I1 XF. Keep that red while fixing
/// the table header's missing white default (issue #1230).
#[test]
fn a_direct_non_default_header_text_color_still_wins() {
    let cursor = Cursor::new(SH107_FORMATTED_TABLE);
    let mut book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).unwrap();
    book.get_sheet_mut(&0)
        .unwrap()
        .get_cell_mut("I1")
        .get_style_mut()
        .get_font_mut()
        .get_color_mut()
        .set_argb("FFFF0000");
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let (doc, _warnings) = XlsxParser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();
    let cell = &get_sheet_page(&doc, 0).table.rows[0].cells[8];

    assert_eq!(cell_text(cell), "Col9");
    assert_eq!(first_run_style(cell).color, Some(Color::new(255, 0, 0)));
    assert_eq!(first_run_style(cell).bold, Some(true));
}

/// Trimming is about paint, not merely the absence of a value: a fill on an
/// empty trailing cell is visible and keeps that column in the printed grid.
#[test]
fn a_painted_value_less_trailing_cell_still_extends_the_printed_grid() {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.get_cell_mut("A1").set_value("Value");
    sheet
        .get_cell_mut("C1")
        .get_style_mut()
        .set_background_color("FFFF0000");
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let (doc, _warnings) = XlsxParser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();
    let page = get_sheet_page(&doc, 0);

    assert_eq!(page.table.column_widths.len(), 3);
    assert_eq!(
        page.table.rows[0].cells[2].background,
        Some(Color::new(255, 0, 0))
    );
}

/// Rows 1, 3 and 4 of SH107 end at a thick border, while rows 1, 2, 4 and 5
/// start at one.
/// A row receives one printed point when it begins with `thickTop`, plus one
/// when the preceding row ends with `thickBot`. Native Excel-for-Mac PDF
/// baselines on the probe are 68, 87, 104, 123, 142, 159, 176, 193 and 210pt:
/// pitches of 19/17/19/19/17/17/17/17 against a bare 17pt track, while the
/// first row's top term places its own baseline. A separate 8-24pt Normal-font
/// sweep kept each individual term at exactly 1pt rather than scaling it with
/// the font.
#[test]
fn thick_row_borders_expand_the_printed_boundary() {
    assert_eq!(
        printed_row_heights(SH107_FORMATTED_TABLE),
        vec![
            Some(18.0),
            Some(19.0),
            Some(17.0),
            Some(19.0),
            Some(19.0),
            Some(17.0),
            Some(17.0),
            Some(17.0),
            Some(17.0),
        ]
    );
}

/// A custom track already states the full printed boundary. The thick flags
/// in `issue_1181_fit_to_height.xlsx` sit on custom-height rows, and adding
/// their reservations again grows the native 161.87pt chart area to 163.43pt.
#[test]
fn custom_row_heights_already_include_thick_border_reservations() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let points = row_boundaries::extract_row_boundary_points(data);

    assert!(!points.contains_key("Monthly college budget"));
}

// ── Theme-resolved scheme Normal fonts (issue #1094) ─────────────────

/// The probe workbook of issue #1094: a `<scheme val="minor"/>` Normal font
/// nominally naming Calibri 11, over a full Office theme.
const THEME_SCHEME_PROBE: &[u8] =
    include_bytes!("../../../../tests/fixtures/xlsx/issue_1060_sheet_row_line_box_probe.xlsx");

/// Every declared row track a workbook prints, across all of its sheet pages.
fn printed_row_heights_over_all_pages(data: &[u8]) -> Vec<Option<f64>> {
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    doc.pages
        .iter()
        .filter_map(|page| match page {
            Page::Sheet(sheet_page) => Some(sheet_page),
            _ => None,
        })
        .flat_map(|sheet_page| sheet_page.table.rows.iter().map(|row| row.height))
        .collect()
}

/// Drop the `<a:font script="..."/>` faces from a theme's minor font scheme,
/// leaving its `<a:latin>` typeface alone — the difference between an Office
/// theme and the bare one LibreOffice writes.
fn strip_theme_minor_font_script_faces(data: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).expect("readable zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("readable entry");
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("readable entry body");
        if name.starts_with("xl/theme/") && name.ends_with(".xml") {
            let xml = String::from_utf8(bytes).expect("theme is utf-8");
            let start: usize = xml.find("<a:minorFont>").expect("theme has a minor font");
            let end: usize = xml[start..]
                .find("</a:minorFont>")
                .expect("minor font is closed")
                + start;
            let stripped: String = strip_script_font_elements(&xml[start..end]);
            bytes = format!("{}{}{}", &xml[..start], stripped, &xml[end..]).into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .expect("writable entry");
        std::io::Write::write_all(&mut out, &bytes).expect("writable entry body");
    }
    out.finish().expect("finished zip").into_inner()
}

/// Remove every `<a:font script=... />` element from one XML fragment.
fn strip_script_font_elements(fragment: &str) -> String {
    let mut kept = String::with_capacity(fragment.len());
    let mut rest: &str = fragment;
    while let Some(start) = rest.find("<a:font script=") {
        kept.push_str(&rest[..start]);
        let end: usize = rest[start..].find("/>").expect("script font is closed") + start + 2;
        rest = &rest[end..];
    }
    kept.push_str(rest);
    kept
}

/// A Normal font that defers its face to the theme scheme is laid out
/// against whatever that scheme resolves to, and a full Office theme gives
/// the minor scheme a per-script face list Excel resolves through — not the
/// Calibri its `<a:latin>` names. The grid then keeps every declared height.
///
/// Native Excel-for-Mac export of the probe workbook, baselines read with
/// `mutool draw -F trace`: its 16 `ht="36"` rows print a 36.00pt track and
/// its 5 `ht="12"` rows a 12.00pt one (issue #1094).
#[test]
fn a_theme_resolved_scheme_normal_font_prints_declared_heights_whole() {
    let heights: Vec<Option<f64>> = printed_row_heights_over_all_pages(THEME_SCHEME_PROBE);

    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(36.0))
            .count(),
        16,
        "16 ht=36 rows print their declared track, got {heights:?}"
    );
    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(12.0))
            .count(),
        5,
        "5 ht=12 rows print their declared track, got {heights:?}"
    );
}

/// The theme is what makes the difference, not the scheme flag on its own.
/// Stripping the same workbook's per-script faces — one factor, nothing else
/// touched — leaves the minor scheme on its Calibri `<a:latin>`, and the
/// export compacts every track: 36 -> 33 and 12 -> 11 (issue #1094).
#[test]
fn a_scheme_normal_font_over_a_theme_without_script_faces_compacts() {
    let bare_theme: Vec<u8> = strip_theme_minor_font_script_faces(THEME_SCHEME_PROBE);

    let heights: Vec<Option<f64>> = printed_row_heights_over_all_pages(&bare_theme);

    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(33.0))
            .count(),
        16,
        "16 ht=36 rows compact to 33pt, got {heights:?}"
    );
    assert_eq!(
        heights
            .iter()
            .filter(|height| **height == Some(11.0))
            .count(),
        5,
        "5 ht=12 rows compact to 11pt, got {heights:?}"
    );
}
