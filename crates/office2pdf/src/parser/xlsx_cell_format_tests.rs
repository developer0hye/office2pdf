use super::*;
use crate::parser::xlsx::xlsx_style::resolve_style_color;

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
    let data = build_xlsx_with_merges("Sheet1", &[("A1", "Merged")], &["A1:B1"]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
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
    let data = build_xlsx_with_merges("Sheet1", &[("A1", "Tall")], &["A1:A2"]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells.len(), 1);
    assert_eq!(tp.table.rows[0].cells[0].row_span, 2);
    assert_eq!(tp.table.rows[0].cells[0].col_span, 1);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Tall");
    assert_eq!(tp.table.rows[1].cells.len(), 0);
}

#[test]
fn test_merge_colspan_and_rowspan() {
    let data = build_xlsx_with_merges(
        "Sheet1",
        &[("A1", "Big"), ("C1", "Right"), ("C2", "Below")],
        &["A1:B2"],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells.len(), 2);
    assert_eq!(tp.table.rows[0].cells[0].col_span, 2);
    assert_eq!(tp.table.rows[0].cells[0].row_span, 2);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Big");
    assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "Right");
    assert_eq!(tp.table.rows[1].cells.len(), 1);
    assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Below");
}

#[test]
fn test_merge_content_in_top_left_only() {
    let data = build_xlsx_with_merges(
        "Sheet1",
        &[("A1", "TopLeft"), ("B1", "should be ignored")],
        &["A1:B1"],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells.len(), 1);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "TopLeft");
}

#[test]
fn test_merge_multiple_ranges() {
    let data = build_xlsx_with_merges(
        "Sheet1",
        &[("A1", "Wide"), ("A2", "Tall"), ("B2", "B2"), ("B3", "B3")],
        &["A1:B1", "A2:A3"],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells.len(), 1);
    assert_eq!(tp.table.rows[0].cells[0].col_span, 2);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "Wide");
    assert_eq!(tp.table.rows[1].cells.len(), 2);
    assert_eq!(tp.table.rows[1].cells[0].row_span, 2);
    assert_eq!(cell_text(&tp.table.rows[1].cells[0]), "Tall");
    assert_eq!(cell_text(&tp.table.rows[1].cells[1]), "B2");
    assert_eq!(tp.table.rows[2].cells.len(), 1);
    assert_eq!(cell_text(&tp.table.rows[2].cells[0]), "B3");
}

#[test]
fn test_merge_no_merges_unchanged() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "X"), ("B1", "Y")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells.len(), 2);
    for cell in &tp.table.rows[0].cells {
        assert_eq!(cell.col_span, 1);
        assert_eq!(cell.row_span, 1);
    }
}

#[test]
fn test_merge_wide_colspan() {
    let data = build_xlsx_with_merges("Sheet1", &[("A1", "Title")], &["A1:D1"]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
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

/// The same workbook with a Normal font that names Calibri 11 outright, as
/// every business golden mock's stylesheet does.
///
/// That is the configuration whose printed grid compacts, and the one the
/// native measurements in the row-track tests below were taken from. umya
/// writes a `<scheme val="minor"/>` font over a full Office theme instead,
/// which Excel resolves by script to a face that keeps every declared height
/// whole (issue #1094).
fn build_xlsx_formatted_over_a_compacting_grid(
    setup: impl FnOnce(&mut umya_spreadsheet::Worksheet),
) -> Vec<u8> {
    rewrite_first_styles_font(&build_xlsx_formatted(setup), "Calibri", 11.0)
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");
    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    // `medium` prints a 2pt boundary-anchored band on a native Excel 16.111
    // probe (issue #619), superseding the 1.75pt centred-stroke calibration
    // of #487.
    assert!((bottom.width - 2.0).abs() < 0.01);
    assert_eq!(bottom.color, Color::new(0, 0, 0));
    let top = border.top.as_ref().expect("Expected top border");
    // `thin`: a 1pt band, 2px at 150 DPI on a native Excel export.
    assert!((top.width - 1.0).abs() < 0.01);
    assert_eq!(top.color, Color::new(255, 0, 0));
}

#[test]
fn test_cell_border_styles() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Styled borders");
        let borders = cell.get_style_mut().get_borders_mut();
        borders
            .get_top_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DASHED);
        borders.get_top_mut().get_color_mut().set_argb("FF000000");
        borders
            .get_bottom_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DOTTED);
        borders
            .get_bottom_mut()
            .get_color_mut()
            .set_argb("FF000000");
        borders
            .get_left_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DASHDOT);
        borders.get_left_mut().get_color_mut().set_argb("FF000000");
        borders
            .get_right_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_DOUBLE);
        borders.get_right_mut().get_color_mut().set_argb("FF000000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
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

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");
    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Dashed);
    // `mediumDashed` shares the `medium` band weight (issues #487, #619).
    assert!((top.width - 2.0).abs() < 0.01);
}

#[test]
fn test_cell_border_hair_and_thick_weights() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Hair/Thick");
        let borders = cell.get_style_mut().get_borders_mut();
        borders
            .get_top_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_HAIR);
        borders.get_top_mut().get_color_mut().set_argb("FF000000");
        borders
            .get_bottom_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_THICK);
        borders
            .get_bottom_mut()
            .get_color_mut()
            .set_argb("FF000000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected border");
    // `hair` prints the same 1pt band as `thin`, with a dotted texture
    // (native Excel 16.111 probe, issue #619).
    let top = border.top.as_ref().expect("Expected top border");
    assert!((top.width - 1.0).abs() < 0.01);
    assert_eq!(top.style, BorderLineStyle::Dotted);
    // `thick` prints a 3pt band on the same probe.
    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert!((bottom.width - 3.0).abs() < 0.01);
    assert_eq!(bottom.style, BorderLineStyle::Solid);
}

#[test]
fn test_row_height() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("Tall row");
        sheet.get_row_dimension_mut(&1).set_height(30.0);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let row = &tp.table.rows[0];
    assert_eq!(row.height, Some(28.0));
}

#[test]
fn test_cell_no_formatting_defaults() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Plain")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let style = first_run_style(cell);
    assert!(style.bold.is_none() || style.bold == Some(false));
    assert!(style.italic.is_none() || style.italic == Some(false));
    assert!(cell.border.is_none());
    assert!(cell.background.is_none());
}

// ----- US-028: Number format tests -----

#[test]
fn test_number_format_currency() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1234.56f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_CURRENCY_USD_SIMPLE);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert!(
        text.contains('$') && text.contains("1,234.56"),
        "Expected currency format with $ and 1,234.56, got: {text}"
    );
}

#[test]
fn accounting_number_format_separates_currency_prefix_from_value() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1234.56f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code(r#"_("$"* #,##0.00_);_("$"* \(#,##0.00_);_("$"* "-"??_);_(@_)"#);
    });
    let parser = XlsxParser;
    let (document, _) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = get_sheet_page(&document, 0);
    let Block::Paragraph(paragraph) = &page.table.rows[0].cells[0].content[0] else {
        unreachable!("a formatted numeric cell emits one paragraph")
    };
    assert_eq!(
        paragraph.runs.len(),
        2,
        "accounting format should keep its left currency prefix separate from the right-aligned value: {:?}",
        paragraph
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>()
    );
    assert_eq!(paragraph.runs[0].text, "$");
    assert_eq!(paragraph.runs[1].text, "1,234.56");
    assert_eq!(
        paragraph
            .style
            .sheet_number_format_accounting_leading_reserved_glyphs
            .as_deref(),
        Some("(")
    );
    assert_eq!(
        paragraph
            .style
            .sheet_number_format_reserved_glyphs
            .as_deref(),
        Some(")")
    );
}

#[test]
fn accounting_number_format_splits_negative_and_zero_sections() {
    let format_code: &str = r#"_("$"* #,##0.00_);_("$"* \(#,##0.00_);_("$"* "-"??_);_(@_)"#;
    for (value, expected_value) in [(-1234.56, "(1,234.56"), (0.0, "-")] {
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(value);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        });
        let parser = XlsxParser;
        let (document, _) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = get_sheet_page(&document, 0);
        let Block::Paragraph(paragraph) = &page.table.rows[0].cells[0].content[0] else {
            unreachable!("a formatted numeric cell emits one paragraph")
        };
        assert_eq!(paragraph.runs.len(), 2, "value {value}");
        assert_eq!(paragraph.runs[0].text, "$", "value {value}");
        assert!(
            paragraph.runs[1].text.contains(expected_value),
            "value {value} should remain intact after the accounting prefix is split: {:?}",
            paragraph.runs[1].text
        );
        assert_eq!(
            paragraph
                .style
                .sheet_number_format_accounting_leading_reserved_glyphs
                .as_deref(),
            Some("("),
            "value {value}"
        );
    }
}

#[test]
fn test_number_format_keeps_quoted_currency_suffix() {
    // The quoted euro literal after the digits was dropped, printing
    // "1,240.00" instead of Excel's "1,240.00 €" (issue #365).
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1240.0f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("#,##0.00\" €\"");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "1,240.00 €");
}

#[test]
fn a_single_section_prefixes_a_negative_sign_before_its_leading_literal() {
    for (value, format_code, expected) in [
        (-123.0, r#""ABC"0.00"#, "-ABC123.00"),
        (-42.5, r#""USD "0.00"#, "-USD 42.50"),
        (-123.0, "0.00", "-123.00"),
        (-123.0, r#""ABC"0.00;0.00"#, "123.00"),
    ] {
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(value);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        });
        let parser = XlsxParser;
        let (document, _) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = get_sheet_page(&document, 0);
        assert_eq!(
            cell_text(&page.table.rows[0].cells[0]),
            expected,
            "format {format_code:?} with value {value}"
        );
    }
}

#[test]
fn a_selected_negative_section_keeps_its_leading_quoted_literal() {
    for (value, format_code, expected) in [
        (-123.0, r#"0.00;"[Red]"0.00"#, "[Red]123.00"),
        (-42.5, r#"0.00;"USD "0.00"#, "USD 42.50"),
        (-42.5, r#""positive "0.00;0.00"#, "42.50"),
        (42.5, r#""positive "0.00;0.00"#, "positive 42.50"),
    ] {
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(value);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        });
        let parser = XlsxParser;
        let (document, _) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = get_sheet_page(&document, 0);
        assert_eq!(
            cell_text(&page.table.rows[0].cells[0]),
            expected,
            "format {format_code:?} with value {value}"
        );
    }
}

#[test]
fn test_number_format_rounds_half_away_from_zero() {
    // Excel rounds display values; the formatter truncated 107310.6 with
    // #,##0 to 107,310 (issue #363).
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(107310.6f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("#,##0");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "107,311");
}

#[test]
fn test_number_format_pads_short_fractions_to_two_decimals() {
    // 39.1 with a two-decimal format rendered "39.100" (issue #364).
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(39.1f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("0.00");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "39.10");
}

#[test]
fn test_percentage_format_rounds_decimal_tie_like_excel() {
    // 21,300 / 20,000 = 1.065 displays as 107% in Excel: the display value
    // is rounded as a decimal, not as the binary double 106.4999… (#363).
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1.065f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_PERCENTAGE);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "107%");
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

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
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

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert!(
        text.contains('%') && text.contains("50.00"),
        "Expected 50.00%, got: {text}"
    );
}

#[test]
fn test_number_format_date() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(45306f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code(umya_spreadsheet::NumberingFormat::FORMAT_DATE_YYYYMMDD);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert!(
        text.contains('-') && !text.contains("45306"),
        "Expected date format yyyy-mm-dd, got: {text}"
    );
}

#[test]
fn test_number_format_thousands_separator() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1234567f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("#,##0");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert_eq!(text, "1,234,567", "Expected thousands separator formatting");
}

/// Excel replaces a formatted numeric value with hashes when its measured
/// ink plus the number format's skip-width padding cannot fit the cell. The
/// native issue #1263 workbook keeps four-digit `7,500` visible in O32 but
/// prints five hashes for the five-digit totals in O37 and O72.
#[test]
fn a_fixed_numeric_format_that_does_not_fit_prints_hashes() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();

    let sheet = get_sheet_page(&doc, 1);
    assert_eq!(cell_text(&sheet.table.rows[31].cells[14]), "7,500");
    assert_eq!(cell_text(&sheet.table.rows[36].cells[14]), "#####");
    assert_eq!(cell_text(&sheet.table.rows[71].cells[14]), "#####");
}

/// Every colour Excel's number-format grammar names paints the fully
/// saturated primary, whatever section selects it.
///
/// Measured on a native Excel 16 for Mac export (2026-09-28) of a one-factor
/// probe workbook: ten cells differing only in their selected section's
/// bracketed colour control. The traces print Red `1 0 0`, Green `0 1 0`,
/// Blue `0 0 1`, Cyan `0 1 1`, Magenta `1 0 1`, Yellow `1 1 0`, White
/// `1 1 1` and Black `0 0 0` (issue #1776).
#[test]
fn every_named_number_format_colour_paints_its_primary() {
    let named_colours: [(&str, Color); 8] = [
        ("Red", Color::new(255, 0, 0)),
        ("Green", Color::new(0, 255, 0)),
        ("Blue", Color::new(0, 0, 255)),
        ("Cyan", Color::new(0, 255, 255)),
        ("Magenta", Color::new(255, 0, 255)),
        ("Yellow", Color::new(255, 255, 0)),
        ("White", Color::new(255, 255, 255)),
        ("Black", Color::new(0, 0, 0)),
    ];
    for (name, expected) in named_colours {
        let format_code = format!("0.00;[{name}]0.00");
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(-123.0f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(&format_code);
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_sheet_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(
            style.color,
            Some(expected),
            "[{name}] must paint its primary"
        );
    }
}

/// The colour a section names wins over the cell font's own colour: the probe
/// prints `[Red]` over an explicit `FF0000FF` font as `1 0 0`, and `[Black]`
/// over that same blue font as `0 0 0` (issue #1776).
#[test]
fn a_number_format_section_colour_overrides_the_cell_font_colour() {
    for (name, expected) in [
        ("Red", Color::new(255, 0, 0)),
        ("Black", Color::new(0, 0, 0)),
    ] {
        let format_code = format!("0.00;[{name}]0.00");
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
            let cell = sheet.get_cell_mut("A1");
            cell.set_value_number(-123.0f64);
            let style = cell.get_style_mut();
            style.get_number_format_mut().set_format_code(&format_code);
            style.get_font_mut().get_color_mut().set_argb("FF0000FF");
        });
        let parser = XlsxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let tp = get_sheet_page(&doc, 0);
        let style = first_run_style(&tp.table.rows[0].cells[0]);
        assert_eq!(
            style.color,
            Some(expected),
            "[{name}] must override the cell font's blue"
        );
    }
}

/// Excel maps `[Color N]` to indexed palette entry `N + 7`; named color
/// controls and the indexed legacy palette are separate number-format forms
/// (issue #1923).
#[test]
fn indexed_number_format_colors_override_the_cell_font_color() {
    let cases = [
        ("A1", "0.00;[Color3]0.00", Color::new(255, 0, 0)),
        ("B1", "0.00;[Color10]0.00", Color::new(0, 128, 0)),
        ("C1", "0.00;[Color46]0.00", Color::new(255, 102, 0)),
        ("D1", "0.00;[Color56]0.00", Color::new(51, 51, 51)),
    ];
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        for (coordinate, format_code, _) in cases {
            let cell = sheet.get_cell_mut(coordinate);
            cell.set_value_number(-123.0f64);
            let style = cell.get_style_mut();
            style.get_number_format_mut().set_format_code(format_code);
            style.get_font_mut().get_color_mut().set_argb("FF0000FF");
        }
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let row = &get_sheet_page(&doc, 0).table.rows[0];
    for (index, (coordinate, _, expected)) in cases.iter().enumerate() {
        assert_eq!(
            first_run_style(&row.cells[index]).color,
            Some(*expected),
            "{coordinate} must use its selected indexed palette color"
        );
    }

    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 1)
        .unwrap();
    assert_eq!(
        first_run_style(&get_sheet_page(&chunks[0], 0).table.rows[0].cells[0]).color,
        Some(Color::new(255, 0, 0)),
        "streaming parsing must use the same indexed palette"
    );
}

/// Once a workbook declares an indexed palette, Excel resolves `[Color N]`
/// through that table. A custom entry replaces the corresponding default
/// color rather than acting as an additive override (issue #1923).
#[test]
fn indexed_number_format_color_uses_a_custom_workbook_palette() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(-123.0f64);
        let style = cell.get_style_mut();
        style
            .get_number_format_mut()
            .set_format_code("0.00;[Color3]0.00");
        style.get_font_mut().get_color_mut().set_argb("FF0000FF");
    });
    let data = rewrite_zip_parts(
        &data,
        |name| name == "xl/styles.xml",
        |xml| {
            let indexed_colors: String = (0..64)
                .map(|index| {
                    let argb = if index == 10 { "FF123456" } else { "FF000000" };
                    format!("<rgbColor rgb=\"{argb}\"/>")
                })
                .collect();
            let colors =
                format!("<colors><indexedColors>{indexed_colors}</indexedColors></colors>");
            let end = xml
                .find("</styleSheet>")
                .expect("styles.xml closes its root element");
            format!("{}{colors}{}", &xml[..end], &xml[end..])
        },
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let cell = &get_sheet_page(&doc, 0).table.rows[0].cells[0];
    assert_eq!(
        first_run_style(cell).color,
        Some(Color::new(18, 52, 86)),
        "[Color3] must resolve to the custom indexedColors[10] value"
    );
}

/// Only the section Excel selects contributes its colour: the same
/// `0.00;[Red]0.00` that reddens a negative value leaves a positive one on the
/// ink a control cell under a colourless `0.00` prints (issue #1776).
#[test]
fn a_colourless_selected_section_keeps_the_cell_font_colour() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        for (coordinate, value, format_code) in [
            ("A1", -123.0f64, "0.00;[Red]0.00"),
            ("A2", 123.0f64, "0.00;[Red]0.00"),
            ("A3", 123.0f64, "0.00"),
        ] {
            let cell = sheet.get_cell_mut(coordinate);
            cell.set_value_number(value);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        }
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let rows = &get_sheet_page(&doc, 0).table.rows;
    let control_ink = first_run_style(&rows[2].cells[0]).color;
    assert_eq!(
        first_run_style(&rows[0].cells[0]).color,
        Some(Color::new(255, 0, 0)),
        "the negative section names red"
    );
    assert_eq!(
        first_run_style(&rows[1].cells[0]).color,
        control_ink,
        "the positive section names no colour, so the cell keeps its own ink"
    );
}

/// A one-section format's colour applies to every value, since that section is
/// the one Excel selects for all of them (issue #1776).
#[test]
fn a_single_section_colour_paints_a_positive_value() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(20.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(123.0f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("[Blue]0.00");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.color, Some(Color::new(0, 0, 255)));
}

/// A quoted colour name is literal text Excel prints, not a colour control,
/// so the value keeps the ink a colourless control cell prints (issue #1776).
#[test]
fn a_quoted_colour_name_is_literal_text_not_an_ink() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_column_dimension_by_number_mut(&1).set_width(30.0);
        for (coordinate, format_code) in [("A1", "0.00;\"[Red]\"0.00"), ("A2", "0.00;0.00")] {
            let cell = sheet.get_cell_mut(coordinate);
            cell.set_value_number(-123.0f64);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        }
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let rows = &get_sheet_page(&doc, 0).table.rows;
    assert_eq!(
        first_run_style(&rows[0].cells[0]).color,
        first_run_style(&rows[1].cells[0]).color,
        "a quoted name names no colour"
    );
}

/// The reported worksheet cell: `-123` under built-in format 8,
/// `"$"#,##0.00_);[Red]("$"#,##0.00)`, inside a `TableStyleMedium2` table.
/// The native Excel export paints its `($123.00)` glyphs `1 0 0` while the
/// converter printed them in the sheet's default black (issue #1776).
#[test]
fn the_reported_table_cell_negative_currency_prints_red() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/SH008-Table-With-Tall-Row.xlsx");
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();

    let sheet = get_sheet_page(&doc, 0);
    let negative = &sheet.table.rows[1].cells[0];
    assert_eq!(cell_text(negative), "($123.00)");
    assert_eq!(
        first_run_style(negative).color,
        Some(Color::new(255, 0, 0)),
        "the [Red] negative section paints the cell red"
    );
}

#[test]
fn test_number_format_general_unchanged() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("42");
        sheet.get_cell_mut("B1").set_value("3.14");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "42");
    assert_eq!(cell_text(&tp.table.rows[0].cells[1]), "3.14");
}

#[test]
fn test_number_format_builtin_id() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(1234.5f64);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_number_format_id(4);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert!(
        text.contains("1,234") && text.contains("50"),
        "Expected #,##0.00 formatting via ID 4, got: {text}"
    );
}

#[test]
fn test_number_format_custom_format_string() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(std::f64::consts::PI);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("0.000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let text = cell_text(&tp.table.rows[0].cells[0]);
    assert_eq!(text, "3.142", "Expected 3 decimal places formatting");
}

#[test]
fn test_cell_combined_formatting() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Full");
        let style = cell.get_style_mut();
        let font = style.get_font_mut();
        font.set_bold(true);
        font.set_size(16.0);
        font.set_name("Helvetica");
        font.get_color_mut().set_argb("FF0000FF");
        style.set_background_color("FFFFCC00");
        let borders = style.get_borders_mut();
        borders
            .get_left_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_THICK);
        borders.get_left_mut().get_color_mut().set_argb("FF00FF00");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let style = first_run_style(cell);
    assert_eq!(style.bold, Some(true));
    assert_eq!(style.font_size, Some(16.0));
    assert_eq!(style.font_family.as_deref(), Some("Helvetica"));
    assert_eq!(style.color, Some(Color::new(0, 0, 255)));
    assert_eq!(cell.background, Some(Color::new(255, 204, 0)));
    let border = cell.border.as_ref().expect("Expected border");
    let left = border.left.as_ref().expect("Expected left border");
    // `thick` prints a 3pt boundary-anchored band (issue #619 probe),
    // superseding the 2.5pt centred-stroke calibration of #487.
    assert!((left.width - 3.0).abs() < 0.01);
    assert_eq!(left.color, Color::new(0, 255, 0));
}

#[test]
fn test_cell_without_underline_style_is_not_underlined() {
    // A font entry with other properties (e.g. bold) but no <u> element must
    // not inherit a spurious underline from the library's enum default.
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Plain");
        cell.get_style_mut().get_font_mut().set_bold(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.underline, None);
}

#[test]
fn test_cell_explicit_underline_is_applied() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Underlined");
        cell.get_style_mut().get_font_mut().set_underline("single");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.underline, Some(true));
}

#[test]
fn test_cell_underline_none_is_not_underlined() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("NoUnderline");
        cell.get_style_mut().get_font_mut().set_underline("none");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.underline, None);
}

#[test]
fn test_cell_horizontal_center_alignment_applied() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Centered");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Center);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    let Block::Paragraph(paragraph) = &cell.content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(paragraph.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_cell_horizontal_right_alignment_applied() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Right");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Right);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(paragraph) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(paragraph.style.alignment, Some(Alignment::Right));
}

#[test]
fn test_cell_vertical_center_alignment_applied() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("Middle");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_vertical(umya_spreadsheet::VerticalAlignmentValues::Center);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].cells[0].vertical_align,
        Some(CellVerticalAlign::Center)
    );
}

#[test]
fn test_cell_without_alignment_keeps_default() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Plain");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(paragraph) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(paragraph.style.alignment, None);
    assert_eq!(tp.table.rows[0].cells[0].vertical_align, None);
}

#[test]
fn test_percent_format_keeps_decimal_precision() {
    // A cached formula ratio formatted as "0.0%" must not round to an
    // integer first (0.17309... rendered "17.0%" instead of "17.3%").
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(0.1730909090909091);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("0.0%");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(paragraph) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(paragraph.runs[0].text, "17.3%");
}

// ----- In-cell rich text runs (issue #275) -----

/// Helper: build a rich text value like the classified workbook's headings —
/// a bold label run followed by a plain continuation run.
fn build_rich_text_cell(setup: impl FnOnce(&mut umya_spreadsheet::Worksheet)) -> Vec<u8> {
    build_xlsx_formatted(setup)
}

#[test]
fn test_rich_text_runs_keep_per_run_formatting() {
    let data = build_rich_text_cell(|sheet| {
        let mut rich = umya_spreadsheet::RichText::default();

        let mut bold_run = umya_spreadsheet::TextElement::default();
        bold_run.set_text("지원율 ");
        {
            let font = bold_run.get_run_properties_mut();
            font.set_bold(true);
            font.set_size(14.0);
            font.get_color_mut().set_argb("FFC00000");
            font.set_name("Arial");
        }
        rich.add_rich_text_elements(bold_run);

        let mut plain_run = umya_spreadsheet::TextElement::default();
        plain_run.set_text("(최근 3년)");
        rich.add_rich_text_elements(plain_run);

        sheet
            .get_cell_mut("A1")
            .get_cell_value_mut()
            .set_rich_text(rich);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(paragraph) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(
        paragraph.runs.len(),
        2,
        "each rich text run must become its own IR run"
    );

    let bold_run = &paragraph.runs[0];
    assert_eq!(bold_run.text, "지원율 ");
    assert_eq!(bold_run.style.bold, Some(true));
    assert_eq!(bold_run.style.font_size, Some(14.0));
    assert_eq!(bold_run.style.font_family.as_deref(), Some("Arial"));
    assert_eq!(
        bold_run.style.color,
        Some(Color {
            r: 0xC0,
            g: 0x00,
            b: 0x00
        })
    );

    let plain_run = &paragraph.runs[1];
    assert_eq!(plain_run.text, "(최근 3년)");
    assert_eq!(plain_run.style.bold, None, "unstyled run stays regular");
    // The run keeps the cell's inherited workbook Normal font rather than the
    // styled run's 14pt; that font is now carried explicitly (issue #462).
    assert_eq!(plain_run.style.font_size, Some(11.0));
}

#[test]
fn test_rich_text_unstyled_run_inherits_cell_style() {
    // Cell-level style is 12pt green italic; a rich run without its own
    // properties must inherit it, while a styled run overrides per-property.
    let data = build_rich_text_cell(|sheet| {
        let mut rich = umya_spreadsheet::RichText::default();

        let mut styled_run = umya_spreadsheet::TextElement::default();
        styled_run.set_text("34.8%");
        // Excel writes minimal <rPr> with only the changed property — build the
        // font from empty instead of get_run_properties_mut(), which seeds the
        // library's full default font (explicit sz=11/Calibri).
        let mut bold_only_font = umya_spreadsheet::Font::default();
        bold_only_font.set_bold(true);
        styled_run.set_run_properties(bold_only_font);
        rich.add_rich_text_elements(styled_run);

        let mut plain_run = umya_spreadsheet::TextElement::default();
        plain_run.set_text(" 달성");
        rich.add_rich_text_elements(plain_run);

        let cell = sheet.get_cell_mut("B2");
        cell.get_cell_value_mut().set_rich_text(rich);
        let font = cell.get_style_mut().get_font_mut();
        font.set_size(12.0);
        font.set_italic(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = tp
        .table
        .rows
        .iter()
        .flat_map(|r| r.cells.iter())
        .find(|c| !c.content.is_empty())
        .expect("cell with content");
    let Block::Paragraph(paragraph) = &cell.content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(paragraph.runs.len(), 2);

    let styled_run = &paragraph.runs[0];
    assert_eq!(styled_run.style.bold, Some(true));
    assert_eq!(
        styled_run.style.font_size,
        Some(12.0),
        "run without explicit size keeps the cell size"
    );
    assert_eq!(styled_run.style.italic, Some(true));

    let plain_run = &paragraph.runs[1];
    assert_eq!(plain_run.style.font_size, Some(12.0));
    assert_eq!(plain_run.style.italic, Some(true));
    assert_eq!(plain_run.style.bold, None);
}

// ----- Text spill into adjacent empty cells (issue #293) -----

#[test]
fn test_long_text_spills_over_empty_neighbors() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value(
            "전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공하는 기관과 제휴",
        );
        // B1..C1 empty, D1 occupied — the spill must stop before D1.
        sheet.get_cell_mut("D1").set_value("차단");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let cell = &tp.table.rows[0].cells[0];
    let spill_width = cell
        .spill_width
        .expect("long unwrapped text with empty right neighbors should spill");
    let own_width = tp.table.column_widths[0];
    let three_columns: f64 = tp.table.column_widths[..3].iter().sum();
    assert!(
        (spill_width - three_columns).abs() < 0.5,
        "spill should cover A..C ({three_columns}pt), got {spill_width}pt (own {own_width}pt)"
    );
}

#[test]
fn test_short_text_does_not_spill() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("짧음");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells[0].spill_width, None);
}

#[test]
fn test_wrap_text_disables_spill() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value(
            "전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공하는 기관과 제휴",
        );
        cell.get_style_mut().get_alignment_mut().set_wrap_text(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].cells[0].spill_width, None,
        "explicit wrapText must wrap inside the cell, not spill"
    );
}

// ----- Unwrapped text clips rather than wraps (issue #615) -----

/// An occupied neighbour leaves nowhere to paint, so the line is clipped at the
/// cell's own edge. It still does not wrap: `wrapText="false"` means the text
/// never moves to a second line, whatever is beside it (issue #615).
#[test]
fn test_occupied_neighbor_clips_at_the_cell_edge_instead_of_wrapping() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value(
            "전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공하는 기관과 제휴",
        );
        sheet.get_cell_mut("B1").set_value("옆");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let own_width = tp.table.column_widths[0];
    let spill_width = tp.table.rows[0].cells[0]
        .spill_width
        .expect("an unwrapped cell stays on one line even with nowhere to spill");
    assert!(
        (spill_width - own_width).abs() < 0.5,
        "clip width should be the cell's own {own_width}pt, got {spill_width}pt"
    );
}

/// A centred cell whose text overruns its column is clipped on one line, the
/// way Excel prints it — it must not fall through to wrapping.
///
/// Probed against Excel 16.0: a centred `wrapText="false"` cell with occupied
/// cells on both sides prints one clipped line. This is the shape that made
/// `코드 근거 미확인` wrap to two lines on the repository workbook (issue #615).
#[test]
fn test_centered_overflowing_cell_clips_instead_of_wrapping() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("왼쪽");
        let cell = sheet.get_cell_mut("B1");
        cell.set_value("전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Center);
        sheet.get_cell_mut("C1").set_value("오른쪽");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let own_width = tp.table.column_widths[1];
    let spill_width = tp.table.rows[0].cells[1]
        .spill_width
        .expect("a centred unwrapped cell stays on one line");
    assert!(
        (spill_width - own_width).abs() < 0.5,
        "a centred cell clips at its own {own_width}pt, got {spill_width}pt"
    );
}

/// Triangulation: right alignment behaves the same as centre, so the rule is
/// "unwrapped never wraps" rather than a special case for one alignment.
#[test]
fn test_right_aligned_overflowing_cell_clips_instead_of_wrapping() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("B1");
        cell.set_value("전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Right);
        sheet.get_cell_mut("C1").set_value("오른쪽");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let own_width = tp.table.column_widths[1];
    let spill_width = tp.table.rows[0].cells[1]
        .spill_width
        .expect("a right-aligned unwrapped cell stays on one line");
    assert!(
        (spill_width - own_width).abs() < 0.5,
        "a right-aligned cell clips at its own {own_width}pt, got {spill_width}pt"
    );
}

/// A centred cell whose text fits needs no clip box at all, so short centred
/// text is unaffected by the rule above.
#[test]
fn test_short_centered_cell_does_not_clip() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("B1");
        cell.set_value("짧음");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Center);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].cells[1].spill_width, None);
}

#[test]
fn test_merged_cell_clips_at_merge_edge_instead_of_wrapping() {
    let data = build_xlsx_with_merges(
        "Sheet1",
        &[(
            "A1",
            "전 직원 통일 방식으로 운영 시, 최소 주 2회 이상의 수업을 제공하는 기관과 제휴",
        )],
        &["A1:C1"],
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let cell = &tp.table.rows[0].cells[0];
    let merged_width: f64 = tp.table.column_widths[..3].iter().sum();
    let spill_width = cell
        .spill_width
        .expect("overflowing unwrapped text in a merge should clip, not wrap");
    assert!(
        (spill_width - merged_width).abs() < 0.5,
        "clip width should equal the merged width {merged_width}pt, got {spill_width}pt"
    );
}

// ----- Default bottom vertical alignment (issue #298) -----

#[test]
fn test_sheet_table_defaults_to_bottom_vertical_alignment() {
    // Excel's default cell vertical alignment is bottom; sheets must carry
    // that down to the renderer so text sits at the row bottom like Excel
    // prints it.
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("바닥 정렬");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.default_vertical_align,
        Some(crate::ir::CellVerticalAlign::Bottom)
    );
}

// ----- Explicit print margins (issue #300) -----

/// An explicitly declared zero margin is a margin, not a missing one.
///
/// Every body edge is exercised because the presence test is per attribute:
/// a fix that only consulted the left edge would leave the other three
/// silently defaulted (issue #1812).
#[test]
fn explicit_zero_print_margins_are_preserved_on_every_body_edge() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("여백 없음");
        let margins = sheet.get_page_margins_mut();
        margins.set_top(0.0);
        margins.set_bottom(0.0);
        margins.set_left(0.0);
        margins.set_right(0.0);
    });
    let (document, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = get_sheet_page(&document, 0);
    assert_eq!(page.margins.top, 0.0, "declared zero top is not absent");
    assert_eq!(page.margins.bottom, 0.0);
    assert_eq!(page.margins.left, 0.0);
    assert_eq!(page.margins.right, 0.0);
}

/// A zero edge beside declared positive edges keeps both: the presence of one
/// attribute says nothing about the others (issue #1812).
#[test]
fn a_zero_margin_beside_positive_margins_keeps_each_declared_value() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("혼합 여백");
        let margins = sheet.get_page_margins_mut();
        margins.set_top(0.0);
        margins.set_bottom(1.0);
        margins.set_left(0.75);
        margins.set_right(0.0);
    });
    let (document, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = get_sheet_page(&document, 0);
    assert_eq!(page.margins.top, 0.0);
    assert_eq!(page.margins.bottom, 72.0, "1in bottom stays 72pt");
    assert_eq!(page.margins.left, 54.0, "0.75in left stays 54pt");
    assert_eq!(page.margins.right, 0.0);
}

/// The synthetic workbook filed with issue #1812: `pageMargins/@left="0"`
/// beside three positive edges, as Excel writes a borderless left margin.
#[test]
fn a_declared_zero_left_margin_seats_the_sheet_on_the_paper_edge() {
    let data: &[u8] = include_bytes!("../../../../tests/visual_audits/issue-1812/source.xlsx");
    let (document, _warnings) = XlsxParser.parse(data, &ConvertOptions::default()).unwrap();
    let page = get_sheet_page(&document, 0);
    assert_eq!(
        page.margins.left, 0.0,
        "explicit zero is not an absent margin"
    );
    assert_eq!(page.margins.right, 50.0, "0.7in right prints against 50");
    assert_eq!(page.margins.top, 82.0, "1.15in top prints against 82");
    assert_eq!(page.margins.bottom, 72.0, "1in bottom prints against 72");
}

#[test]
fn test_explicit_page_margins_are_used() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("여백 테스트");
        let margins = sheet.get_page_margins_mut();
        margins.set_top(1.0);
        margins.set_bottom(1.0);
        margins.set_left(0.75);
        margins.set_right(0.75);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.margins.top, 72.0, "1in top margin must be honored");
    assert_eq!(tp.margins.bottom, 72.0);
    assert_eq!(tp.margins.left, 54.0);
    assert_eq!(tp.margins.right, 54.0);
}

#[test]
fn test_absent_page_margins_fall_back_to_excel_defaults() {
    let data = remove_page_margins(&build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("기본 여백");
    }));
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.margins.top, 54.0, "Excel default 0.75in top");
    assert_eq!(
        tp.margins.left, 50.0,
        "Excel default 0.7in left, on the whole point it prints against"
    );
}

// ----- Printed page laid out on whole device points (issue #1127) -----

#[test]
fn test_a_fractional_print_margin_snaps_to_the_whole_point_below_it() {
    // Excel lays a printed sheet out on whole device points, so the 2cm
    // margin a metric Excel writes — 0.7874in, 56.69pt — puts the sheet's
    // first grid boundary on 56, and every row boundary below it follows from
    // there. Printing against the exact margin left the whole grid up to 1pt
    // low.
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("여백 스냅");
        let margins = sheet.get_page_margins_mut();
        margins.set_top(0.787_401_575);
        margins.set_bottom(0.787_401_575);
        margins.set_left(0.7);
        margins.set_right(0.7);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.margins.top, 56.0, "56.69pt top prints against 56");
    assert_eq!(tp.margins.bottom, 56.0);
    assert_eq!(tp.margins.left, 50.0, "50.4pt left prints against 50");
    assert_eq!(tp.margins.right, 50.0);
}

#[test]
fn test_print_margins_snap_downwards_rather_than_to_the_nearest_point() {
    // A fraction over a half still prints against the point below it: the
    // rule is a floor, not a round. A margin already on a whole point is left
    // exactly where the file puts it.
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("여백 내림");
        let margins = sheet.get_page_margins_mut();
        margins.set_top(1.25); // 90pt, already whole
        margins.set_bottom(0.99); // 71.28pt
        margins.set_left(1.3); // 93.6pt
        margins.set_right(0.4); // 28.8pt
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.margins.top, 90.0, "a whole margin is left alone");
    assert_eq!(tp.margins.bottom, 71.0, "71.28pt prints against 71");
    assert_eq!(tp.margins.left, 93.0, "93.6pt prints against 93");
    assert_eq!(tp.margins.right, 28.0, "28.8pt prints against 28");
}

// ----- Row heights without customHeight (issues #303, #1151) -----

#[test]
fn test_recorded_ht_without_custom_height_is_recomputed_from_the_normal_font() {
    // An `ht` a row records without `customHeight` is a cached auto-height,
    // and Excel recomputes the row from the Normal font on load rather than
    // printing the cached number.
    //
    // Measured on `issue_1066_blip_effect_picture.xlsx`, whose rows 1, 3, 4
    // and 5 carry `ht="16"` and no `customHeight` while row 7 carries no
    // dimension at all. Sweeping the Normal font size one variant per export
    // and reading `height of row 1` and `height of row 7` back through
    // AppleScript answers the same number for both at every size — 11, 12, 14,
    // 15, 16, 17, 19, 21, 24 for Calibri 8, 9, 10, 11, 12, 13, 14, 16, 18 —
    // and neither is 16 except where the recompute happens to land there
    // (issue #1151).
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("행 높이");
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(20.0);
        row.set_custom_height(false);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(14.0),
        "the cached ht is inert: 11pt Calibri recomputes a 15pt worksheet row, \
         which this compacting grid prints as a 14pt track"
    );
}

#[test]
fn test_recorded_ht_with_custom_height_stays_the_declared_track() {
    // The same height with the flag set is a fixed track, and still reaches
    // the page calibrated to the native PDF grid (issue #303).
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("행 높이");
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(20.0);
        row.set_custom_height(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(18.0),
        "a customHeight row keeps its declared 20pt, compacted to an 18pt track"
    );
}

#[test]
fn test_a_row_auto_grown_by_an_unmeasured_font_size_keeps_its_cached_ht() {
    // A row Excel auto-grew around a large font is written as an `ht` with no
    // `customHeight`, and Excel recomputes very nearly the same height on
    // load. Every face series stops at 24pt, so a 26pt row has no modelled
    // track of its own, and falling back to the Normal font's 15pt one would
    // print it at half the height its own text needs. The row's cached `ht`
    // is what Excel last measured for that text, so it stands in.
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("분기 실적 요약");
        cell.get_style_mut().get_font_mut().set_size(26.0);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(34.0);
        row.set_custom_height(false);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(31.0),
        "no series covers a 26pt cell, so the cached 34pt stands and this \
         compacting grid prints it as a 31pt track"
    );
}

#[test]
fn test_recorded_ht_survives_where_the_normal_font_recompute_is_unmeasured() {
    // No sweep covers this face, so nothing here knows what Excel recomputes
    // for it. The row's own cached `ht` is Excel's last recompute of that very
    // row, which is a closer hint than the sheet's `defaultRowHeight` — and
    // keeping the declared number where nothing is measured is the same
    // convention the face tables follow for a size their series skips.
    let data = rewrite_first_styles_font(
        &build_xlsx_formatted(|sheet| {
            sheet.get_cell_mut("A1").set_value("행 높이");
            sheet
                .get_sheet_format_properties_mut()
                .set_default_row_height(30.0);
            let row = sheet.get_row_dimension_mut(&1);
            row.set_height(20.0);
            row.set_custom_height(false);
        }),
        "Comic Sans MS",
        11.0,
    );
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(20.0),
        "an unmeasured face keeps the cached ht rather than falling back to the \
         sheet's declared defaultRowHeight"
    );
}

#[test]
fn test_row_without_dimension_takes_the_normal_font_recompute() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("첫째 줄");
        sheet.get_cell_mut("A2").set_value("둘째 줄");
        sheet
            .get_sheet_format_properties_mut()
            .set_default_row_height(18.0);
        sheet.get_row_dimension_mut(&1).set_height(24.0);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].height, Some(22.0));
    assert_eq!(
        tp.table.rows[1].height,
        Some(14.0),
        "a row without its own ht is recomputed from the Normal font, not from \
         defaultRowHeight: 11pt Calibri gives a 15pt worksheet row over this \
         sheet's declared 18, which the compacting grid prints as 14"
    );
}

#[test]
fn test_wrapping_row_without_custom_height_stays_auto() {
    // Excel auto-grows rows containing wrapped cells unless customHeight is
    // set; our text metrics differ slightly from Excel's, so a fixed height
    // could clip a line — keep those rows content-driven.
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("줄바꿈이 있는 긴 텍스트가 이 셀에 들어 있습니다");
        cell.get_style_mut().get_alignment_mut().set_wrap_text(true);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(30.0);
        row.set_custom_height(false);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height, None,
        "auto-sized wrapping rows stay content-driven"
    );
}

#[test]
fn test_wrapping_row_whose_text_fits_takes_the_native_track() {
    // `wrapText` is a property of the cell, not evidence that anything wraps.
    // Every data cell in the business mocks carries it, so treating the flag
    // alone as "content-driven" left every ht=15 auto row growing to its own
    // content box: 15.00pt against Excel's 14.00pt across the six Latin
    // workbooks (issue #710), and 22.32pt against 15.00pt across the Korean
    // ones (issue #709). A single short word cannot wrap, so the row prints
    // a mapped track like any other — Excel's own recompute for a wrapped
    // Calibri 11 cell under a Calibri 11 Normal font, which the native
    // reading puts one point over the unwrapped 15 (issue #709), compacted
    // to the 15pt track. What must not happen is the `None` of a
    // content-driven row.
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("OK");
        cell.get_style_mut().get_alignment_mut().set_wrap_text(true);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(15.0);
        row.set_custom_height(false);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(15.0),
        "a wrapText cell that fits its column still prints Excel's mapped track"
    );
}

#[test]
fn test_wrapping_row_with_custom_height_stays_fixed() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("줄바꿈이 있는 긴 텍스트가 이 셀에 들어 있습니다");
        cell.get_style_mut().get_alignment_mut().set_wrap_text(true);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(30.0);
        row.set_custom_height(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.rows[0].height, Some(28.0));
}

#[test]
fn test_native_excel_print_height_calibrates_custom_title_rows() {
    // Native Excel PDF output measures a 25.5pt worksheet row as a 23pt
    // printed track. Keeping the raw OOXML value makes both the title row
    // and the table header too tall and pushes the data block down.
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("Dashboard title");
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(25.5);
        row.set_custom_height(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[0].height,
        Some(23.0),
        "25.5pt custom rows must match Excel's 23pt PDF track"
    );
}

#[test]
fn test_native_excel_print_height_calibrates_default_rows() {
    // The same native print path measures a declared 15pt default row as a
    // 14pt track. This matters for blank spacer rows because they have
    // no content from which Typst could derive the native height.
    //
    // `customHeight` is what keeps the declared 15 in play: without it
    // Excel ignores the hint and recomputes the default from the Normal
    // font instead (issue #1047).
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_cell_mut("A1").set_value("Title");
        sheet.get_cell_mut("A3").set_value("Header");
        let properties = sheet.get_sheet_format_properties_mut();
        properties.set_default_row_height(15.0);
        properties.set_custom_height(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.rows[1].height,
        Some(14.0),
        "blank default rows must match Excel's 14pt PDF track"
    );
}

#[test]
fn test_xlsx_auto_row_padding_matches_native_14pt_track() {
    // Arial 10's full hhea line plus Excel's asymmetric 1pt/1.5pt vertical
    // insets forms the 14pt single-line track measured in native output.
    // Keeping only 1pt below makes every later row drift upward by ~0.5pt.
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("North America");
        cell.get_style_mut().get_alignment_mut().set_wrap_text(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let padding = get_sheet_page(&doc, 0)
        .table
        .default_cell_padding
        .expect("XLSX table padding");
    assert_eq!(padding.top, 1.0);
    assert_eq!(padding.bottom, 1.5);
}

// ----- Print titles (issue #234) -----

/// Helper: build a workbook whose sheet declares `_xlnm.Print_Titles`.
fn build_xlsx_with_print_titles(
    address: &str,
    setup: impl FnOnce(&mut umya_spreadsheet::Worksheet),
) -> Vec<u8> {
    build_xlsx_formatted(|sheet| {
        setup(sheet);
        sheet
            .add_defined_name("_xlnm.Print_Titles", address)
            .unwrap();
    })
}

#[test]
fn test_print_title_rows_become_repeating_header() {
    let data = build_xlsx_with_print_titles("Sheet1!$1:$2", |sheet| {
        sheet.get_cell_mut("A1").set_value("제목 1");
        sheet.get_cell_mut("A2").set_value("제목 2");
        sheet.get_cell_mut("A3").set_value("데이터");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.header_row_count, 2,
        "title rows $1:$2 must repeat as the table header"
    );
}

/// Excel repeats only the rows named by `_xlnm.Print_Titles`. When the range
/// starts below the sheet top, the rows above it print once.
#[test]
fn test_print_titles_below_sheet_top_do_not_repeat_the_rows_above() {
    let data = build_xlsx_with_print_titles("Sheet1!$3:$3", |sheet| {
        sheet
            .get_cell_mut("A1")
            .set_value("Warehouse Inventory Snapshot");
        sheet.get_cell_mut("A3").set_value("SKU");
        sheet.get_cell_mut("A4").set_value("SKU-1000");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(
        tp.table.header_row_count, 1,
        "only row 3 repeats on later pages"
    );
    assert_eq!(
        tp.table.non_repeating_header_row_count, 2,
        "rows 1-2 lead the table but must not repeat"
    );
}

/// A title range starting at the sheet top has nothing above it to hold back.
#[test]
fn test_print_titles_at_sheet_top_have_no_non_repeating_rows() {
    let data = build_xlsx_with_print_titles("Sheet1!$1:$2", |sheet| {
        sheet.get_cell_mut("A1").set_value("Title");
        sheet.get_cell_mut("A2").set_value("Header");
        sheet.get_cell_mut("A3").set_value("Data");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.header_row_count, 2);
    assert_eq!(tp.table.non_repeating_header_row_count, 0);
}

#[test]
fn test_no_print_titles_means_no_header() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("데이터");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.header_row_count, 0);
}

#[test]
fn test_print_title_columns_repeat_on_overflow_pages() {
    // Column A is a title column; enough wide columns follow to force
    // column pagination. Every overflow page must start with column A.
    let data = build_xlsx_with_print_titles("Sheet1!$A:$A", |sheet| {
        sheet.get_cell_mut("A1").set_value("이름");
        for col in 2..=12u32 {
            let cell = sheet.get_cell_mut((col, 1));
            cell.set_value(format!("값{col}"));
        }
        for col in 1..=12u32 {
            sheet
                .get_column_dimension_by_number_mut(&col)
                .set_width(30.0);
        }
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(doc.pages.len() >= 2, "wide sheet must paginate by columns");

    for (page_idx, page) in doc.pages.iter().enumerate().skip(1) {
        let Page::Sheet(sp) = page else {
            panic!("expected sheet page");
        };
        let first_cell_text = cell_text(&sp.table.rows[0].cells[0]);
        assert_eq!(
            first_cell_text, "이름",
            "page {page_idx} must repeat the title column"
        );
    }
}

#[test]
fn test_print_titles_with_both_rows_and_columns() {
    // Mirrors `Sheet4!$A:$B,Sheet4!$2:$3`-style definitions with two parts.
    let data = build_xlsx_with_print_titles("Sheet1!$A:$A,Sheet1!$1:$1", |sheet| {
        sheet.get_cell_mut("A1").set_value("이름");
        for col in 2..=12u32 {
            sheet.get_cell_mut((col, 1)).set_value(format!("값{col}"));
        }
        sheet.get_cell_mut("A2").set_value("둘째");
        for col in 1..=12u32 {
            sheet
                .get_column_dimension_by_number_mut(&col)
                .set_width(30.0);
        }
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    assert_eq!(tp.table.header_row_count, 1, "row titles parsed");
    assert!(doc.pages.len() >= 2);
    let Page::Sheet(sp) = &doc.pages[1] else {
        panic!("expected sheet page");
    };
    assert_eq!(
        cell_text(&sp.table.rows[0].cells[0]),
        "이름",
        "column titles parsed from the multi-part address"
    );
}

// ----- RTL text presentation (issue #236) -----

#[test]
fn test_rtl_text_right_aligns_under_general_alignment() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("نص");
        sheet.get_cell_mut("A2").set_value("עִבְרִית");
        sheet.get_cell_mut("A3").set_value("text");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let alignment_of = |row: usize| match &tp.table.rows[row].cells[0].content[0] {
        Block::Paragraph(p) => p.style.alignment,
        _ => panic!("expected paragraph"),
    };
    assert_eq!(
        alignment_of(0),
        Some(crate::ir::Alignment::Right),
        "Arabic text under general alignment renders right-aligned in Excel"
    );
    assert_eq!(alignment_of(1), Some(crate::ir::Alignment::Right));
    assert_eq!(alignment_of(2), None, "Latin text keeps the default");
}

#[test]
fn test_explicit_alignment_wins_over_rtl_text() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("نص");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Left);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(p) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(p.style.alignment, Some(crate::ir::Alignment::Left));
}

#[test]
fn test_native_digit_locale_format_renders_arabic_indic_digits() {
    let data = build_xlsx_formatted(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value_number(123.0);
        cell.get_style_mut()
            .get_number_format_mut()
            .set_format_code("[$-3000401]0");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);
    let Block::Paragraph(p) = &tp.table.rows[0].cells[0].content[0] else {
        panic!("expected paragraph");
    };
    assert_eq!(
        p.runs[0].text, "١٢٣",
        "native-digit locale prefix must map to Arabic-Indic digits"
    );
}

// ----- Spill past the used range (issue #309) -----

#[test]
fn test_spill_extends_past_used_range_over_virtual_cells() {
    // Single used column: Excel paints the text across the virtual empty
    // cells to the right instead of wrapping inside column A.
    let data = build_xlsx_formatted(|sheet| {
        sheet
            .get_cell_mut("A1")
            .set_value("text with comment stretching well past column A");
        sheet.get_cell_mut("A2").set_value("둘째");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let cell = &tp.table.rows[0].cells[0];
    let own_width = tp.table.column_widths[0];
    let spill_width = cell
        .spill_width
        .expect("text must spill past the used range");
    assert!(
        spill_width > own_width * 2.0,
        "spill ({spill_width}pt) must extend well past the single {own_width}pt column"
    );
}

/// An occupied neighbour still stops the text painting past column A — the
/// clip box is the cell's own width, not the width the text wants. The line
/// stays a line either way (issue #615).
#[test]
fn test_spill_still_blocked_by_occupied_neighbor() {
    let data = build_xlsx_formatted(|sheet| {
        sheet
            .get_cell_mut("A1")
            .set_value("text with comment stretching well past column A");
        sheet.get_cell_mut("B1").set_value("차단");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let own_width = tp.table.column_widths[0];
    let spill_width = tp.table.rows[0].cells[0]
        .spill_width
        .expect("an unwrapped cell still renders one clipped line");
    assert!(
        (spill_width - own_width).abs() < 0.5,
        "an occupied neighbor holds the clip to column A's {own_width}pt, got {spill_width}pt"
    );
}

/// `<color theme="N"/>` names a slot in the workbook's colour scheme and
/// carries no `rgb`, so reading colours through `get_argb()` alone left every
/// themed run, fill and border to the renderer's default (issue #853).
/// umya's map is `[lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink]`, the
/// ECMA-376 order with the light/dark swap on slots 0 and 1; the Office theme
/// umya builds by default puts accent1 at 4472C4 and accent2 at ED7D31.
#[test]
fn test_font_theme_color_resolves_through_the_workbook_scheme() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Accent one");
        sheet
            .get_cell_mut("A1")
            .get_style_mut()
            .get_font_mut()
            .get_color_mut()
            .set_theme_index(4);
    });
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.color, Some(Color::new(0x44, 0x72, 0xC4)));
}

/// A second slot, so nothing passes by resolving every theme index to one
/// colour.
#[test]
fn test_a_second_theme_slot_resolves_to_its_own_colour() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Accent two");
        sheet
            .get_cell_mut("A1")
            .get_style_mut()
            .get_font_mut()
            .get_color_mut()
            .set_theme_index(5);
    });
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let style = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style.color, Some(Color::new(0xED, 0x7D, 0x31)));
}

/// `tint` shifts the resolved colour's luminance. Slot 1 is the scheme's
/// black, so a positive tint has to lighten it to a neutral grey — resolving
/// the slot but dropping the tint would leave it black.
#[test]
fn test_a_tint_lightens_the_resolved_theme_colour() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Tinted");
        let color = sheet
            .get_cell_mut("A1")
            .get_style_mut()
            .get_font_mut()
            .get_color_mut();
        color.set_theme_index(1);
        color.set_tint(0.5);
    });
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let color = first_run_style(&tp.table.rows[0].cells[0])
        .color
        .expect("a tinted theme colour must resolve");
    assert_eq!(color.r, color.g, "a tinted black stays neutral");
    assert_eq!(color.g, color.b, "a tinted black stays neutral");
    assert!(
        color.r > 0x40 && color.r < 0xC0,
        "tint 0.5 on black must land mid-grey, got {:#04x}",
        color.r
    );
}

/// Excel applies SpreadsheetML `tint` in its 240-step integer HLS space.
/// The Gift Budget and Tracker fixture from issue #1394 supplies accent 2 as
/// `DAB6BA` and a body fill at `tint="0.7999"`; native Excel prints `F8EFF0`.
/// A floating-point 255-step conversion instead lands one RGB step lighter at
/// `F8F0F1`.
#[test]
fn test_positive_theme_tint_matches_native_excel_hls_quantization() {
    let mut book = umya_spreadsheet::new_file();
    let mut accent_two = umya_spreadsheet::structs::drawing::RgbColorModelHex::default();
    accent_two.set_val("DAB6BA");
    book.get_theme_mut()
        .get_theme_elements_mut()
        .get_color_scheme_mut()
        .get_accent2_mut()
        .set_rgb_color_model_hex(accent_two);

    let mut source = umya_spreadsheet::Color::default();
    source.set_theme_index(5).set_tint(0.7999);

    assert_eq!(
        resolve_style_color(&source, Some(book.get_theme())),
        Some(Color::new(0xF8, 0xEF, 0xF0))
    );
}

/// A non-solid `patternFill` paints its foreground over its background at
/// the pattern's ink coverage; it is not a solid foreground (issue #926).
///
/// The legend of `004_Gantt-prosjektplanlegger1.xlsx` (attached to #841) is
/// the measurement: `lightUp` with `fgColor` `735773` over an omitted
/// background prints `DCD5DC`, which is white blended one quarter of the way
/// to `735773` on all three channels. Four of its five swatches were coming
/// out as the same `735773`.
#[test]
fn test_non_solid_pattern_fill_blends_foreground_over_background() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Swatch");
        sheet
            .get_cell_mut("A1")
            .get_style_mut()
            .get_fill_mut()
            .get_pattern_fill_mut()
            .set_pattern_type(umya_spreadsheet::PatternValues::LightUp)
            .get_foreground_color_mut()
            .set_argb("FF735773");
    });
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(
        tp.table.rows[0].cells[0].background,
        Some(Color::new(0xDC, 0xD5, 0xDC)),
        "lightUp over an omitted (white) background is a quarter of the way to the foreground"
    );
}

/// Triangulation for #926: a different pattern with a different coverage, and
/// a stated background, so neither a fixed quarter nor a white background can
/// pass. `mediumGray` covers half the cell, so `000000` over `FFFFFF` is the
/// mid-grey `808080`, and `darkGray` covers three quarters of it.
#[test]
fn test_pattern_fill_coverage_follows_the_pattern_type() {
    for (pattern, expected) in [
        (umya_spreadsheet::PatternValues::MediumGray, 0x80_u8),
        (umya_spreadsheet::PatternValues::DarkGray, 0x40_u8),
        (umya_spreadsheet::PatternValues::Gray125, 0xE0_u8),
    ] {
        let data = build_xlsx_formatted(|sheet| {
            sheet.get_cell_mut("A1").set_value("Swatch");
            let fill = sheet
                .get_cell_mut("A1")
                .get_style_mut()
                .get_fill_mut()
                .get_pattern_fill_mut();
            fill.set_pattern_type(pattern.clone());
            fill.get_foreground_color_mut().set_argb("FF000000");
            fill.get_background_color_mut().set_argb("FFFFFFFF");
        });
        let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let tp = get_sheet_page(&doc, 0);
        let background = tp.table.rows[0].cells[0].background.expect("a fill");

        assert!(
            background.r.abs_diff(expected) <= 1
                && background.g.abs_diff(expected) <= 1
                && background.b.abs_diff(expected) <= 1,
            "{pattern:?} must land near {expected:#04x}, got {background:?}"
        );
    }
}

/// `patternType="none"` — every workbook's `fillId` 0 — paints nothing, and
/// the new coverage path must not invent a colour for it.
///
/// The combination `none` + `fgColor` is NOT covered here: umya promotes it to
/// `solid` in `PatternFill::auto_set_pattern_type`, before the parser can see
/// it, so no assertion at this level can pin it down.
#[test]
fn test_pattern_type_none_paints_no_background() {
    const STYLES_WITH_A_BARE_NONE_FILL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1"/></cellXfs>
</styleSheet>"#;
    let data = build_xlsx_with_style_xfs(
        STYLES_WITH_A_BARE_NONE_FILL,
        r#"<c r="A1" s="0" t="inlineStr"><is><t>Bare</t></is></c>"#,
    );
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(tp.table.rows[0].cells[0].background, None);
}

/// The same resolution has to reach a cell's fill, not only its text.
#[test]
fn test_fill_theme_color_resolves_through_the_workbook_scheme() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Filled");
        let style = sheet.get_cell_mut("A1").get_style_mut();
        style
            .get_fill_mut()
            .get_pattern_fill_mut()
            .set_pattern_type(umya_spreadsheet::PatternValues::Solid);
        style
            .get_fill_mut()
            .get_pattern_fill_mut()
            .get_foreground_color_mut()
            .set_theme_index(4);
    });
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(
        tp.table.rows[0].cells[0].background,
        Some(Color::new(0x44, 0x72, 0xC4))
    );
}

/// Build a one-cell workbook from raw parts so the `cellStyleXfs` /
/// `cellXfs` / `xfId` relationship can be stated exactly — umya's builder
/// cannot express a Normal style that switches a format category off.
fn build_xlsx_with_style_xfs(styles_xml: &str, cell_xml: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();
    let mut write = |path: &str, body: &str| {
        zip.start_file(path, options).unwrap();
        std::io::Write::write_all(&mut zip, body.as_bytes()).unwrap();
    };

    write(
        "[Content_Types].xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
    );
    write(
        "_rels/.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );
    write(
        "xl/workbook.xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
    );
    write(
        "xl/_rels/workbook.xml.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );
    write("xl/styles.xml", styles_xml);
    write(
        "xl/worksheets/sheet1.xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">{cell_xml}</row></sheetData>
</worksheet>"#
        ),
    );

    zip.finish().unwrap().into_inner()
}

/// A `styles.xml` whose Normal style (`cellStyleXfs[0]`) switches number
/// formatting off, and whose `cellXfs[1]` inherits from `cellStyleXfs[1]`
/// instead. Every cell resolved against entry 0 while `xfId` went unread, so
/// the whole workbook lost its number formats (issue #851).
const STYLES_WITH_APPLY_NUMBER_FORMAT_OFF: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="0"/>
<xf numFmtId="9" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="1"/>
</cellXfs>
</styleSheet>"#;

#[test]
fn test_builtin_percent_format_survives_a_normal_style_that_disables_it() {
    let data = build_xlsx_with_style_xfs(
        STYLES_WITH_APPLY_NUMBER_FORMAT_OFF,
        r#"<c r="A1" s="1"><v>0.25</v></c>"#,
    );
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "25%");
}

/// A different value and a different builtin id, so nothing can pass by
/// returning a fixed string.
#[test]
fn test_builtin_two_decimal_percent_resolves_through_its_own_xf() {
    let styles = STYLES_WITH_APPLY_NUMBER_FORMAT_OFF.replace(r#"numFmtId="9""#, r#"numFmtId="10""#);
    let data = build_xlsx_with_style_xfs(&styles, r#"<c r="A1" s="1"><v>0.5</v></c>"#);
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "50.00%");
}

/// A cell that really does point at the Normal style keeps taking its veto,
/// so the fix cannot be "ignore applyNumberFormat".
#[test]
fn test_a_cell_pointing_at_the_normal_style_still_takes_its_veto() {
    let styles = STYLES_WITH_APPLY_NUMBER_FORMAT_OFF.replace(
        r#"<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="1"/>"#,
        r#"<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#,
    );
    let data = build_xlsx_with_style_xfs(&styles, r#"<c r="A1" s="1"><v>0.25</v></c>"#);
    let (doc, _warnings) = XlsxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert_eq!(cell_text(&tp.table.rows[0].cells[0]), "0.25");
}

/// Excel composes a merged range's outline from the cells on each edge — the
/// bottom from the range's bottom row, the right from its right column — while
/// we built the whole cell from the top-left member alone. The Gantt template
/// of issue #841 merges each header label down one row and declares the header
/// rule only on the bottom members, so the rule vanished across every merged
/// column while surviving on the unmerged ones beside them (issue #939).
#[test]
fn merged_range_takes_its_bottom_border_from_the_bottom_row() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("AKTIVITET");
        // The bottom member of the merge carries the rule, as Excel writes it.
        let bottom_member = sheet.get_cell_mut("A2");
        let borders = bottom_member.get_style_mut().get_borders_mut();
        borders
            .get_bottom_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_THIN);
        borders
            .get_bottom_mut()
            .get_color_mut()
            .set_argb("FF7F5F7F");
        sheet.add_merge_cells("A1:A2");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    assert_eq!(cell.row_span, 2, "the merge is still one cell");
    let border = cell
        .border
        .as_ref()
        .expect("the merged range keeps the rule its bottom row declares");
    let bottom = border
        .bottom
        .as_ref()
        .expect("the bottom side comes from the range's bottom row");
    assert_eq!(bottom.color, Color::new(0x7F, 0x5F, 0x7F));
    assert!((bottom.width - 1.0).abs() < 0.01);
}

/// The same composition on the horizontal axis: a right border declared on the
/// range's right-hand member has to reach the merged cell (issue #939).
#[test]
fn merged_range_takes_its_right_border_from_the_right_column() {
    let data = build_xlsx_formatted(|sheet| {
        sheet.get_cell_mut("A1").set_value("Wide");
        let right_member = sheet.get_cell_mut("C1");
        let borders = right_member.get_style_mut().get_borders_mut();
        borders
            .get_right_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_MEDIUM);
        borders.get_right_mut().get_color_mut().set_argb("FF0000FF");
        sheet.add_merge_cells("A1:C1");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let cell = &tp.table.rows[0].cells[0];
    assert_eq!(cell.col_span, 3);
    let right = cell
        .border
        .as_ref()
        .and_then(|border| border.right.as_ref())
        .expect("the right side comes from the range's right column");
    assert_eq!(right.color, Color::new(0, 0, 255));
}

/// The top-left member still states the top and left sides, and a range whose
/// members declare nothing keeps no border at all (issue #939).
#[test]
fn merged_range_keeps_the_top_left_members_own_sides() {
    let data = build_xlsx_formatted(|sheet| {
        let top_left = sheet.get_cell_mut("A1");
        top_left.set_value("Corner");
        let borders = top_left.get_style_mut().get_borders_mut();
        borders
            .get_top_mut()
            .set_border_style(umya_spreadsheet::Border::BORDER_THIN);
        borders.get_top_mut().get_color_mut().set_argb("FF00FF00");
        sheet.get_cell_mut("E1").set_value("Plain");
        sheet.add_merge_cells("A1:B2");
        sheet.add_merge_cells("E1:F1");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_sheet_page(&doc, 0);
    let corner = &tp.table.rows[0].cells[0];
    let border = corner.border.as_ref().expect("the top side survives");
    assert_eq!(
        border.top.as_ref().expect("top side").color,
        Color::new(0, 255, 0)
    );
    assert!(border.bottom.is_none(), "no member declares a bottom side");
    let plain = tp.table.rows[0]
        .cells
        .iter()
        .find(|cell| cell_text(cell) == "Plain")
        .expect("the unbordered merge is still emitted");
    assert!(
        plain.border.is_none(),
        "a merge whose members declare nothing keeps no border"
    );
}

// ── A wrapped single-line cell lifts its auto row (issue #709) ────────

/// The four Korean business mocks, whose Normal font names Calibri 11
/// outright, so their printed grid compacts (issue #1094).
const PAYROLL_KO_MOCK: &[u8] =
    include_bytes!("../../../../tests/golden_mocks/business/sources/xlsx/04_payroll_ko.xlsx");
const QUOTATION_KO_MOCK: &[u8] =
    include_bytes!("../../../../tests/golden_mocks/business/sources/xlsx/01_quotation_ko.xlsx");

/// One auto row over the compacting Calibri 11 grid holding a single cell of
/// `text` in `family` at `size_pt`, with or without `wrapText`, and no `ht`
/// at all — Excel's own recompute decides the whole row.
fn printed_auto_row_height(family: &str, size_pt: f64, text: &str, wraps: bool) -> Option<f64> {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        sheet.get_column_dimension_mut("A").set_width(40.0);
        let cell = sheet.get_cell_mut("A1");
        cell.set_value(text);
        let style = cell.get_style_mut();
        if wraps {
            style.get_alignment_mut().set_wrap_text(true);
        }
        let font = style.get_font_mut();
        font.set_name(family);
        font.set_size(size_pt);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    get_sheet_page(&doc, 0).table.rows[0].height
}

/// A wrapped cell whose text fits one line still re-measures its auto row:
/// Excel gives it one point more than the row its face recomputes at that
/// size, and the compacting grid prints that.
///
/// Native Excel-for-Mac readings of `row height of row N` over a Calibri 11
/// Normal font (the business mocks' configuration), one wrapped single-line
/// cell per auto row, no `ht`:
///
/// | cell font | unwrapped | wrapped | printed (x0.92) |
/// | --- | ---: | ---: | ---: |
/// | Malgun Gothic 10 | 15 | 16 | 15 |
/// | Malgun Gothic 14 | 20 | 21 | 19 |
/// | Arial 12 | 16 | 17 | 16 |
/// | Calibri 11 | 15 | 16 | 15 |
///
/// The Malgun Gothic 10 row is the one every Korean mock's data rows print:
/// 15.00pt in all four golden exports where the unlifted track prints 14.
#[test]
fn test_a_wrapped_single_line_cell_lifts_its_auto_row_one_point_over_its_face_recompute() {
    for (family, size_pt, text, expected) in [
        ("Malgun Gothic", 10.0, "김민준", 15.0),
        ("Malgun Gothic", 14.0, "김민준", 19.0),
        ("Arial", 12.0, "Quarterly total", 16.0),
        ("Calibri", 11.0, "Quarterly total", 15.0),
    ] {
        assert_eq!(
            printed_auto_row_height(family, size_pt, text, true),
            Some(expected),
            "a wrapped {family} {size_pt}pt cell"
        );
    }
}

/// The lift needs the face's recompute to reach the row's default: a face
/// that recomputes a shorter row than the Normal font's leaves the default
/// standing, wrapped or not. Measured the same way: Arial 10 (13pt own row)
/// and Malgun Gothic 9 (14pt) both read 15 under Calibri 11, wrapped, which
/// prints 14 — the track every Latin business mock's wrapped Arial 10 data
/// rows print (issue #710).
#[test]
fn test_a_wrapped_cell_whose_face_recompute_sits_under_the_default_row_keeps_the_default() {
    for (family, size_pt, text) in [
        ("Arial", 10.0, "Quarterly total"),
        ("Malgun Gothic", 9.0, "김민준"),
        ("Calibri", 10.0, "Quarterly total"),
    ] {
        assert_eq!(
            printed_auto_row_height(family, size_pt, text, true),
            Some(14.0),
            "a wrapped {family} {size_pt}pt cell"
        );
    }
}

/// Without `wrapText` the same cell sizes its row by the face's bare
/// recompute: quotation rows 11-13 hold unwrapped Malgun Gothic 10 totals and
/// print 14.00pt in the native export.
#[test]
fn test_an_unwrapped_cell_takes_no_wrap_point() {
    assert_eq!(
        printed_auto_row_height("Malgun Gothic", 10.0, "김민준", false),
        Some(14.0)
    );
}

/// A wrapped cell with nothing in it has no line to re-measure: the probe's
/// empty wrapped Malgun Gothic 10 cell read the bare 15, printed 14.
#[test]
fn test_an_empty_wrapped_cell_takes_no_wrap_point() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let empty = sheet.get_cell_mut("A1");
        empty.set_value("");
        let style = empty.get_style_mut();
        style.get_alignment_mut().set_wrap_text(true);
        let font = style.get_font_mut();
        font.set_name("Malgun Gothic");
        font.set_size(10.0);
        // An unwrapped neighbour so the row exists to be measured.
        let neighbour = sheet.get_cell_mut("B1");
        neighbour.set_value("E-1021");
        neighbour.get_style_mut().get_font_mut().set_name("Arial");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(get_sheet_page(&doc, 0).table.rows[0].height, Some(14.0));
}

/// A face with no measured recompute contributes no lift either; the probe
/// cannot say what Excel does with it, so the row keeps its default rather
/// than borrowing another face's point.
#[test]
fn test_a_wrapped_cell_in_an_unmeasured_face_takes_no_wrap_point() {
    assert_eq!(
        printed_auto_row_height("Noto Sans CJK SC", 10.0, "김민준", true),
        Some(14.0)
    );
}

/// A wrapped cell inside a merged range does not re-measure the row: the
/// probe's `A:B` merge of a wrapped Malgun Gothic 10 cell read the bare 15,
/// and the payroll mock's merged, wrapped `합계` label prints its row at 14.
#[test]
fn test_a_merged_wrapped_cell_takes_no_wrap_point() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("공급가액 합계");
        let style = cell.get_style_mut();
        style.get_alignment_mut().set_wrap_text(true);
        let font = style.get_font_mut();
        font.set_name("Malgun Gothic");
        font.set_size(10.0);
        sheet.add_merge_cells("A1:B1");
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(get_sheet_page(&doc, 0).table.rows[0].height, Some(14.0));
}

/// A `customHeight` track is the user's, not Excel's recompute: the payroll
/// mock's row 9 carries `ht=15 customHeight="true"` and prints 14.00pt.
#[test]
fn test_a_fixed_track_ignores_the_wrap_point() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("클라우드 인프라");
        let style = cell.get_style_mut();
        style.get_alignment_mut().set_wrap_text(true);
        let font = style.get_font_mut();
        font.set_name("Malgun Gothic");
        font.set_size(10.0);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(15.0);
        row.set_custom_height(true);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(get_sheet_page(&doc, 0).table.rows[0].height, Some(14.0));
}

/// An `ht` recorded without `customHeight` is a cache Excel discards
/// (issue #1151), so the wrapped cell's lift still applies over it — exactly
/// the shape of payroll rows 4-8 and attendance rows 4-9.
#[test]
fn test_a_cached_auto_height_still_takes_the_wrap_point() {
    let data = build_xlsx_formatted_over_a_compacting_grid(|sheet| {
        let cell = sheet.get_cell_mut("A1");
        cell.set_value("E-1021");
        let style = cell.get_style_mut();
        style.get_alignment_mut().set_wrap_text(true);
        let font = style.get_font_mut();
        font.set_name("Malgun Gothic");
        font.set_size(10.0);
        let row = sheet.get_row_dimension_mut(&1);
        row.set_height(15.0);
        row.set_custom_height(false);
    });
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert_eq!(get_sheet_page(&doc, 0).table.rows[0].height, Some(15.0));
}

/// The payroll mock end to end against its native export's row bands, read
/// off the horizontal rules with `mutool draw -F trace`: header row 3 at
/// 23.00pt, the five wrapped Malgun Gothic 10 data rows at 15.00pt each, and
/// the fixed total row 9 at 14.00pt.
#[test]
fn test_the_payroll_mocks_korean_data_rows_print_excels_fifteen_point_track() {
    let heights: Vec<Option<f64>> = printed_row_heights_over_all_pages(PAYROLL_KO_MOCK);
    assert_eq!(heights[2], Some(23.0), "header row 3, got {heights:?}");
    assert_eq!(
        &heights[3..8],
        &[Some(15.0); 5],
        "data rows 4-8, got {heights:?}"
    );
    assert_eq!(heights[8], Some(14.0), "total row 9, got {heights:?}");
}

/// The quotation mock carries every gate at once: unwrapped Korean rows 3-4
/// (14.00pt), wrapped Malgun data rows 7-10 (15.00pt), and rows 11-13 whose
/// only unmerged Malgun cell is an unwrapped total (14.00pt) — all read off
/// the native export's rules.
#[test]
fn test_the_quotation_mocks_rows_take_the_lift_only_where_a_wrapped_cell_sits_unmerged() {
    let heights: Vec<Option<f64>> = printed_row_heights_over_all_pages(QUOTATION_KO_MOCK);
    assert_eq!(
        &heights[2..4],
        &[Some(14.0); 2],
        "rows 3-4, got {heights:?}"
    );
    assert_eq!(
        &heights[6..10],
        &[Some(15.0); 4],
        "rows 7-10, got {heights:?}"
    );
    assert_eq!(
        &heights[10..13],
        &[Some(14.0); 3],
        "rows 11-13, got {heights:?}"
    );
}
