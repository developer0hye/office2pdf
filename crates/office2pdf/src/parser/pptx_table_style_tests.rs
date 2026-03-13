use super::*;
use table_styles::{PptxTableProps, PptxTableStyleDef, TableCellRegionStyle, TableStyleMap};

// ── Unit tests: parse_table_styles_xml ─────────────────────────────────

fn make_table_style_xml(styles: &[(&str, &str)]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">"#,
    );
    for (style_id, body) in styles {
        xml.push_str(&format!(
            r#"<a:tblStyle styleId="{style_id}" styleName="Test">{body}</a:tblStyle>"#
        ));
    }
    xml.push_str("</a:tblStyleLst>");
    xml
}

fn test_theme() -> ThemeData {
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    parse_theme_xml(&theme_xml)
}

fn test_color_map() -> ColorMapData {
    default_color_map()
}

#[test]
fn test_parse_table_style_with_whole_table_fill() {
    let body = r#"<a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>"#;
    let xml = make_table_style_xml(&[("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", body)]);
    let theme: ThemeData = test_theme();
    let color_map: ColorMapData = test_color_map();

    let styles: TableStyleMap =
        table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let style: &PptxTableStyleDef = styles
        .get("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        .expect("style not found");
    let whole = style.whole_table.as_ref().expect("wholeTbl missing");
    assert_eq!(whole.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_parse_table_style_with_first_row_scheme_color() {
    // firstRow with accent1 fill and white bold text
    let body = concat!(
        r#"<a:firstRow>"#,
        r#"<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle>"#,
        r#"</a:firstRow>"#,
    );
    let xml = make_table_style_xml(&[("style1", body)]);
    let theme: ThemeData = test_theme();
    let color_map: ColorMapData = test_color_map();

    let styles: TableStyleMap =
        table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let style: &PptxTableStyleDef = styles.get("style1").expect("style not found");
    let first_row = style.first_row.as_ref().expect("firstRow missing");
    assert_eq!(first_row.fill, Some(Color::new(0x44, 0x72, 0xC4)));
    assert_eq!(first_row.text_color, Some(Color::new(0xFF, 0xFF, 0xFF)));
    assert_eq!(first_row.text_bold, Some(true));
}

#[test]
fn test_parse_table_style_banded_rows() {
    let body = concat!(
        r#"<a:band1H><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill></a:fill></a:tcStyle></a:band1H>"#,
        r#"<a:band2H><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:fill></a:tcStyle></a:band2H>"#,
    );
    let xml = make_table_style_xml(&[("bandtest", body)]);
    let theme: ThemeData = test_theme();
    let color_map: ColorMapData = test_color_map();

    let styles: TableStyleMap =
        table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let style: &PptxTableStyleDef = styles.get("bandtest").expect("style not found");
    assert_eq!(
        style.band1_h.as_ref().unwrap().fill,
        Some(Color::new(0xDD, 0xDD, 0xDD))
    );
    assert_eq!(
        style.band2_h.as_ref().unwrap().fill,
        Some(Color::new(0xFF, 0xFF, 0xFF))
    );
}

#[test]
fn test_parse_table_style_with_color_transforms() {
    // accent1=#4472C4 with tint 40% → blend toward white
    let body = r#"<a:band1H><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>"#;
    let xml = make_table_style_xml(&[("tinttest", body)]);
    let theme: ThemeData = test_theme();
    let color_map: ColorMapData = test_color_map();

    let styles: TableStyleMap =
        table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let style: &PptxTableStyleDef = styles.get("tinttest").expect("style not found");
    let band = style.band1_h.as_ref().expect("band1H missing");
    // accent1 = (68, 114, 196). tint 40%: channel = 255 - (255-ch)*0.4
    // r = 255 - 187*0.4 = 255 - 74.8 = 180.2 → 180
    // g = 255 - 141*0.4 = 255 - 56.4 = 198.6 → 199
    // b = 255 - 59*0.4 = 255 - 23.6 = 231.4 → 231
    assert_eq!(band.fill, Some(Color::new(180, 199, 231)));
}

// ── Unit tests: apply_table_style ──────────────────────────────────────

#[test]
fn test_apply_table_style_first_row_gets_header_fill_and_text_color() {
    let mut styles: TableStyleMap = HashMap::new();
    styles.insert(
        "style1".to_string(),
        PptxTableStyleDef {
            first_row: Some(TableCellRegionStyle {
                fill: Some(Color::new(0x44, 0x72, 0xC4)),
                text_color: Some(Color::new(255, 255, 255)),
                text_bold: Some(true),
            }),
            ..Default::default()
        },
    );
    let props = PptxTableProps {
        style_id: Some("style1".to_string()),
        first_row: true,
        ..Default::default()
    };

    // Build a simple 2-row table with no explicit fills
    let mut table = Table {
        rows: vec![
            TableRow {
                cells: vec![TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Header".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    col_span: 1,
                    row_span: 1,
                    border: None,
                    background: None,
                    data_bar: None,
                    icon_text: None,
                    vertical_align: None,
                    padding: None,
                }],
                height: Some(30.0),
            },
            TableRow {
                cells: vec![TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Data".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    col_span: 1,
                    row_span: 1,
                    border: None,
                    background: None,
                    data_bar: None,
                    icon_text: None,
                    vertical_align: None,
                    padding: None,
                }],
                height: Some(30.0),
            },
        ],
        column_widths: vec![200.0],
        header_row_count: 1,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    // Header row cell should have blue background and white bold text
    let header_cell = &table.rows[0].cells[0];
    assert_eq!(header_cell.background, Some(Color::new(0x44, 0x72, 0xC4)));
    let header_run = match &header_cell.content[0] {
        Block::Paragraph(p) => &p.runs[0],
        _ => panic!("Expected paragraph"),
    };
    assert_eq!(header_run.style.color, Some(Color::new(255, 255, 255)));
    assert_eq!(header_run.style.bold, Some(true));

    // Data row should be unaffected
    let data_cell = &table.rows[1].cells[0];
    assert_eq!(data_cell.background, None);
}

#[test]
fn test_apply_table_style_banded_rows_skip_first_row() {
    let mut styles: TableStyleMap = HashMap::new();
    styles.insert(
        "bandstyle".to_string(),
        PptxTableStyleDef {
            band1_h: Some(TableCellRegionStyle {
                fill: Some(Color::new(0xDD, 0xEE, 0xFF)),
                text_color: None,
                text_bold: None,
            }),
            ..Default::default()
        },
    );
    let props = PptxTableProps {
        style_id: Some("bandstyle".to_string()),
        first_row: true,
        band_row: true,
        ..Default::default()
    };

    let make_row = |text: &str| -> TableRow {
        TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: text.to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                col_span: 1,
                row_span: 1,
                border: None,
                background: None,
                data_bar: None,
                icon_text: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }
    };

    let mut table = Table {
        rows: vec![make_row("Header"), make_row("Row1"), make_row("Row2"), make_row("Row3")],
        column_widths: vec![200.0],
        header_row_count: 1,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    // Header row (row 0) excluded from banding
    assert_eq!(table.rows[0].cells[0].background, None);
    // Row 1 (data row index 0) = band1 → fill applied
    assert_eq!(
        table.rows[1].cells[0].background,
        Some(Color::new(0xDD, 0xEE, 0xFF))
    );
    // Row 2 (data row index 1) = band2 → no fill (band2 not defined)
    assert_eq!(table.rows[2].cells[0].background, None);
    // Row 3 (data row index 2) = band1 → fill applied
    assert_eq!(
        table.rows[3].cells[0].background,
        Some(Color::new(0xDD, 0xEE, 0xFF))
    );
}

#[test]
fn test_apply_table_style_explicit_cell_fill_not_overridden() {
    let mut styles: TableStyleMap = HashMap::new();
    styles.insert(
        "override".to_string(),
        PptxTableStyleDef {
            whole_table: Some(TableCellRegionStyle {
                fill: Some(Color::new(0xAA, 0xBB, 0xCC)),
                text_color: None,
                text_bold: None,
            }),
            ..Default::default()
        },
    );
    let props = PptxTableProps {
        style_id: Some("override".to_string()),
        ..Default::default()
    };

    let mut table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Explicit".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                col_span: 1,
                row_span: 1,
                border: None,
                background: Some(Color::new(0xFF, 0x00, 0x00)),
                data_bar: None,
                icon_text: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }],
        column_widths: vec![200.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    // Explicit cell fill should be preserved, not overridden by wholeTbl
    assert_eq!(
        table.rows[0].cells[0].background,
        Some(Color::new(0xFF, 0x00, 0x00))
    );
}

#[test]
fn test_apply_table_style_missing_style_id_is_noop() {
    let styles: TableStyleMap = HashMap::new();
    let props = PptxTableProps {
        style_id: None,
        ..Default::default()
    };

    let mut table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![],
                col_span: 1,
                row_span: 1,
                border: None,
                background: None,
                data_bar: None,
                icon_text: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }],
        column_widths: vec![200.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    assert_eq!(table.rows[0].cells[0].background, None);
}
