use super::*;

fn build_xlsx_with_cond_fmt(setup: impl FnOnce(&mut umya_spreadsheet::Worksheet)) -> Vec<u8> {
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

#[test]
fn test_cond_fmt_greater_than_background() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(10.0);
        sheet.get_cell_mut("A2").set_value_number(60.0);
        sheet.get_cell_mut("A3").set_value_number(50.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::CellIs);
        rule.set_operator(umya_spreadsheet::ConditionalFormattingOperatorValues::GreaterThan);
        rule.set_priority(1);
        let mut style = umya_spreadsheet::Style::default();
        style.set_background_color("FFFF0000");
        rule.set_style(style);
        let mut formula = umya_spreadsheet::Formula::default();
        formula.set_string_value("50");
        rule.set_formula(formula);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A3");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert!(
        tp.table.rows[0].cells[0].background.is_none(),
        "A1 (10) should NOT match >50"
    );
    assert_eq!(
        tp.table.rows[1].cells[0].background,
        Some(Color::new(255, 0, 0)),
        "A2 (60) should match >50 and get red bg"
    );
    assert!(
        tp.table.rows[2].cells[0].background.is_none(),
        "A3 (50) should NOT match >50"
    );
}

#[test]
fn test_cond_fmt_less_than_font_color() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(15.0);
        sheet.get_cell_mut("A2").set_value_number(25.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::CellIs);
        rule.set_operator(umya_spreadsheet::ConditionalFormattingOperatorValues::LessThan);
        rule.set_priority(1);
        let mut style = umya_spreadsheet::Style::default();
        style.get_font_mut().get_color_mut().set_argb("FF0000FF");
        rule.set_style(style);
        let mut formula = umya_spreadsheet::Formula::default();
        formula.set_string_value("20");
        rule.set_formula(formula);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A2");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let style_a1 = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(
        style_a1.color,
        Some(Color::new(0, 0, 255)),
        "A1 (15) should match <20 and get blue color"
    );
    let style_a2 = first_run_style(&tp.table.rows[1].cells[0]);
    assert!(style_a2.color.is_none(), "A2 (25) should NOT match <20");
}

#[test]
fn test_cond_fmt_equal_bold() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(100.0);
        sheet.get_cell_mut("A2").set_value_number(99.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::CellIs);
        rule.set_operator(umya_spreadsheet::ConditionalFormattingOperatorValues::Equal);
        rule.set_priority(1);
        let mut style = umya_spreadsheet::Style::default();
        style.get_font_mut().set_bold(true);
        rule.set_style(style);
        let mut formula = umya_spreadsheet::Formula::default();
        formula.set_string_value("100");
        rule.set_formula(formula);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A2");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let style_a1 = first_run_style(&tp.table.rows[0].cells[0]);
    assert_eq!(style_a1.bold, Some(true), "A1 (100) should be bold");
    let style_a2 = first_run_style(&tp.table.rows[1].cells[0]);
    assert!(
        style_a2.bold.is_none() || style_a2.bold == Some(false),
        "A2 (99) should NOT be bold"
    );
}

#[test]
fn test_cond_fmt_between() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(5.0);
        sheet.get_cell_mut("A2").set_value_number(20.0);
        sheet.get_cell_mut("A3").set_value_number(35.0);
        sheet.get_cell_mut("A4").set_value_number(10.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::CellIs);
        rule.set_operator(umya_spreadsheet::ConditionalFormattingOperatorValues::Between);
        rule.set_priority(1);
        let mut style = umya_spreadsheet::Style::default();
        style.set_background_color("FF00FF00");
        rule.set_style(style);
        let mut formula = umya_spreadsheet::Formula::default();
        formula.set_string_value("10");
        rule.set_formula(formula);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A4");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert!(tp.table.rows[0].cells[0].background.is_none());
    assert_eq!(
        tp.table.rows[1].cells[0].background,
        Some(Color::new(0, 255, 0))
    );
    assert_eq!(
        tp.table.rows[2].cells[0].background,
        Some(Color::new(0, 255, 0))
    );
    assert_eq!(
        tp.table.rows[3].cells[0].background,
        Some(Color::new(0, 255, 0))
    );
}

#[test]
fn test_cond_fmt_color_scale_two_color() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(0.0);
        sheet.get_cell_mut("A2").set_value_number(50.0);
        sheet.get_cell_mut("A3").set_value_number(100.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::ColorScale);
        rule.set_priority(1);

        let mut cs = umya_spreadsheet::ColorScale::default();

        let mut cfvo_min = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo_min.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Min);
        cs.add_cfvo_collection(cfvo_min);

        let mut cfvo_max = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo_max.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Max);
        cs.add_cfvo_collection(cfvo_max);

        let mut color_min = umya_spreadsheet::Color::default();
        color_min.set_argb("FFFFFFFF");
        cs.add_color_collection(color_min);

        let mut color_max = umya_spreadsheet::Color::default();
        color_max.set_argb("FFFF0000");
        cs.add_color_collection(color_max);

        rule.set_color_scale(cs);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A3");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let bg_a1 = tp.table.rows[0].cells[0]
        .background
        .expect("A1 should have color scale bg");
    assert_eq!(bg_a1, Color::new(255, 255, 255));

    let bg_a3 = tp.table.rows[2].cells[0]
        .background
        .expect("A3 should have color scale bg");
    assert_eq!(bg_a3, Color::new(255, 0, 0));

    let bg_a2 = tp.table.rows[1].cells[0]
        .background
        .expect("A2 should have color scale bg");
    assert_eq!(bg_a2.r, 255);
    assert!(
        bg_a2.g > 100 && bg_a2.g < 150,
        "Expected ~128, got {}",
        bg_a2.g
    );
    assert!(
        bg_a2.b > 100 && bg_a2.b < 150,
        "Expected ~128, got {}",
        bg_a2.b
    );
}

#[test]
fn test_cond_fmt_no_rules_unchanged() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "42")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert!(tp.table.rows[0].cells[0].background.is_none());
}

#[test]
fn test_cond_fmt_non_numeric_cell_skipped() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value("hello");
        sheet.get_cell_mut("A2").set_value_number(60.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::CellIs);
        rule.set_operator(umya_spreadsheet::ConditionalFormattingOperatorValues::GreaterThan);
        rule.set_priority(1);
        let mut style = umya_spreadsheet::Style::default();
        style.set_background_color("FFFF0000");
        rule.set_style(style);
        let mut formula = umya_spreadsheet::Formula::default();
        formula.set_string_value("50");
        rule.set_formula(formula);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A2");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    assert!(tp.table.rows[0].cells[0].background.is_none());
    assert_eq!(
        tp.table.rows[1].cells[0].background,
        Some(Color::new(255, 0, 0))
    );
}

#[test]
fn test_cond_fmt_data_bar() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(10.0);
        sheet.get_cell_mut("A2").set_value_number(50.0);
        sheet.get_cell_mut("A3").set_value_number(100.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::DataBar);
        rule.set_priority(1);

        let mut db = umya_spreadsheet::DataBar::default();
        let mut cfvo_min = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo_min.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Min);
        let mut cfvo_max = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo_max.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Max);
        db.add_cfvo_collection(cfvo_min);
        db.add_cfvo_collection(cfvo_max);
        let mut bar_color = umya_spreadsheet::Color::default();
        bar_color.set_argb("FF638EC6");
        db.add_color_collection(bar_color);
        rule.set_data_bar(db);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A3");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let db1 = tp.table.rows[0].cells[0]
        .data_bar
        .as_ref()
        .expect("A1 should have data_bar");
    assert!(
        db1.fill_pct < 0.01,
        "Min value should have ~0% fill, got {}",
        db1.fill_pct
    );

    let db2 = tp.table.rows[1].cells[0]
        .data_bar
        .as_ref()
        .expect("A2 should have data_bar");
    assert!(
        (db2.fill_pct - 100.0 * (50.0 - 10.0) / (100.0 - 10.0)).abs() < 1.0,
        "Mid value should have ~44% fill, got {}",
        db2.fill_pct
    );

    let db3 = tp.table.rows[2].cells[0]
        .data_bar
        .as_ref()
        .expect("A3 should have data_bar");
    assert!(
        (db3.fill_pct - 100.0).abs() < 0.01,
        "Max value should have 100% fill, got {}",
        db3.fill_pct
    );

    assert_eq!(db1.color, Color::new(0x63, 0x8E, 0xC6));
}

#[test]
fn test_cond_fmt_icon_set() {
    let data = build_xlsx_with_cond_fmt(|sheet| {
        sheet.get_cell_mut("A1").set_value_number(10.0);
        sheet.get_cell_mut("A2").set_value_number(50.0);
        sheet.get_cell_mut("A3").set_value_number(90.0);

        let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
        rule.set_type(umya_spreadsheet::ConditionalFormatValues::IconSet);
        rule.set_priority(1);

        let mut is = umya_spreadsheet::IconSet::default();
        let mut cfvo0 = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo0.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Percent);
        cfvo0.set_val("0");
        let mut cfvo1 = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo1.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Percent);
        cfvo1.set_val("33");
        let mut cfvo2 = umya_spreadsheet::ConditionalFormatValueObject::default();
        cfvo2.set_type(umya_spreadsheet::ConditionalFormatValueObjectValues::Percent);
        cfvo2.set_val("67");
        is.add_cfvo_collection(cfvo0);
        is.add_cfvo_collection(cfvo1);
        is.add_cfvo_collection(cfvo2);
        rule.set_icon_set(is);

        let mut seq = umya_spreadsheet::SequenceOfReferences::default();
        seq.set_sqref("A1:A3");
        let mut cf = umya_spreadsheet::ConditionalFormatting::default();
        cf.set_sequence_of_references(seq);
        cf.add_conditional_collection(rule);
        sheet.set_conditional_formatting_collection(vec![cf]);
    });

    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let tp = get_sheet_page(&doc, 0);

    let icon1 = tp.table.rows[0].cells[0]
        .icon_text
        .as_ref()
        .expect("A1 should have icon_text");
    assert_eq!(icon1, "↓", "Low value should get down arrow");

    let icon2 = tp.table.rows[1].cells[0]
        .icon_text
        .as_ref()
        .expect("A2 should have icon_text");
    assert_eq!(icon2, "→", "Mid value should get right arrow");

    let icon3 = tp.table.rows[2].cells[0]
        .icon_text
        .as_ref()
        .expect("A3 should have icon_text");
    assert_eq!(icon3, "↑", "High value should get up arrow");
}
