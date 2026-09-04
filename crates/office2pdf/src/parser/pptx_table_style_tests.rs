use super::*;
use crate::test_support::{make_table_graphic_frame, make_table_row};
use std::io::Write;
use table_styles::{PptxTableProps, PptxTableStyleDef, TableCellRegionStyle, TableStyleMap};

// ── Helpers ────────────────────────────────────────────────────────────

fn table_element(elem: &FixedElement) -> &Table {
    match &elem.kind {
        FixedElementKind::Table(table) => table,
        _ => panic!("Expected Table, got {:?}", elem.kind),
    }
}

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

    let styles: TableStyleMap = table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let style: &PptxTableStyleDef = styles
        .get("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        .expect("style not found");
    let whole = style.whole_table.as_ref().expect("wholeTbl missing");
    assert_eq!(whole.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn a_table_style_fill_preserves_its_alpha() {
    let body = r#"<a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="95000"/><a:alpha val="35000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>"#;
    let xml = make_table_style_xml(&[("translucent", body)]);
    let theme = test_theme();
    let color_map = test_color_map();

    let styles = table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

    let whole = styles["translucent"]
        .whole_table
        .as_ref()
        .expect("wholeTbl missing");
    assert_eq!(whole.fill_alpha, Some(0.35));
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

    let styles: TableStyleMap = table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

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

    let styles: TableStyleMap = table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

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

    let styles: TableStyleMap = table_styles::parse_table_styles_xml(&xml, &theme, &color_map);

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
                fill_alpha: Some(0.4),
                text_font_family: None,
                text_color: Some(Color::new(255, 255, 255)),
                text_bold: Some(true),
                borders: Default::default(),
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
                minimum_height: None,
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
                    background_alpha: None,
                    data_bar: None,
                    sparkline: None,
                    icon_text: None,
                    icon_color: None,
                    icon_shading: None,
                    spill_width: None,
                    vertical_align: None,
                    padding: None,
                }],
                height: Some(30.0),
            },
            TableRow {
                minimum_height: None,
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
                    background_alpha: None,
                    data_bar: None,
                    sparkline: None,
                    icon_text: None,
                    icon_color: None,
                    icon_shading: None,
                    spill_width: None,
                    vertical_align: None,
                    padding: None,
                }],
                height: Some(30.0),
            },
        ],
        column_widths: vec![200.0],
        header_row_count: 1,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    // Header row cell should have blue background and white bold text
    let header_cell = &table.rows[0].cells[0];
    assert_eq!(header_cell.background, Some(Color::new(0x44, 0x72, 0xC4)));
    assert_eq!(header_cell.background_alpha, Some(0.4));
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
                fill_alpha: None,
                text_font_family: None,
                text_color: None,
                text_bold: None,
                borders: Default::default(),
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
            minimum_height: None,
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
                background_alpha: None,
                data_bar: None,
                sparkline: None,
                icon_text: None,
                icon_color: None,
                icon_shading: None,
                spill_width: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }
    };

    let mut table = Table {
        rows: vec![
            make_row("Header"),
            make_row("Row1"),
            make_row("Row2"),
            make_row("Row3"),
        ],
        column_widths: vec![200.0],
        header_row_count: 1,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
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
                fill_alpha: None,
                text_font_family: None,
                text_color: None,
                text_bold: None,
                borders: Default::default(),
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
            minimum_height: None,
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
                background_alpha: None,
                data_bar: None,
                sparkline: None,
                icon_text: None,
                icon_color: None,
                icon_shading: None,
                spill_width: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }],
        column_widths: vec![200.0],
        header_row_count: 0,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
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
            minimum_height: None,
            cells: vec![TableCell {
                content: vec![],
                col_span: 1,
                row_span: 1,
                border: None,
                background: None,
                background_alpha: None,
                data_bar: None,
                sparkline: None,
                icon_text: None,
                icon_color: None,
                icon_shading: None,
                spill_width: None,
                vertical_align: None,
                padding: None,
            }],
            height: Some(30.0),
        }],
        column_widths: vec![200.0],
        header_row_count: 0,
        non_repeating_header_row_count: 0,
        alignment: None,
        default_cell_padding: None,
        use_content_driven_row_heights: true,
        default_vertical_align: None,
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::CenteredStroke,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    assert_eq!(table.rows[0].cells[0].background, None);
}

// ── Integration tests: end-to-end PPTX with table styles ──────────────

/// Build a PPTX with theme and tableStyles.xml included.
/// `master_xml` is the slide master, or `None` to omit the part entirely.
/// Only the master carries a `<p:clrMap>`, so a deck built without one cannot
/// resolve `bg1`/`tx1` in a table style.
fn build_test_pptx_with_table_styles(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slide_xmls: &[String],
    theme_xml: &str,
    table_styles_xml: &str,
    master_xml: Option<&str>,
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    for i in 0..slide_xmls.len() {
        ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            i + 1
        ));
    }
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
        slide_cx_emu, slide_cy_emu
    );
    for i in 0..slide_xmls.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    pres_rels.push_str(
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
    );
    for i in 0..slide_xmls.len() {
        pres_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
            2 + i,
            1 + i
        ));
    }
    if master_xml.is_some() {
        pres_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#,
            2 + slide_xmls.len()
        ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(theme_xml.as_bytes()).unwrap();

    zip.start_file("ppt/tableStyles.xml", opts).unwrap();
    zip.write_all(table_styles_xml.as_bytes()).unwrap();

    if let Some(master) = master_xml {
        zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
            .unwrap();
        zip.write_all(master.as_bytes()).unwrap();
    }

    for (i, slide_xml) in slide_xmls.iter().enumerate() {
        zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_pptx_table_with_style_applies_header_fill_and_text_color() {
    // Table style: firstRow has accent1 fill and white bold text, band1H has light tint
    let table_styles_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5940675A-B579-460E-94D1-54222C63F5DA}">"#,
        r#"<a:tblStyle styleId="{5940675A-B579-460E-94D1-54222C63F5DA}" styleName="Test">"#,
        r#"<a:firstRow>"#,
        r#"<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle>"#,
        r#"</a:firstRow>"#,
        r#"<a:band1H>"#,
        r#"<a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle>"#,
        r#"</a:band1H>"#,
        r#"</a:tblStyle>"#,
        r#"</a:tblStyleLst>"#,
    );

    // Table with tblPr firstRow=1 bandRow=1 and a tableStyleId
    let table_xml = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/>"#,
        r#"<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>"#,
        r#"<p:nvPr/></p:nvGraphicFramePr>"#,
        r#"<p:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#,
        r#"<a:tbl>"#,
        r#"<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr>"#,
        r#"<a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>"#,
        // Header row with white text (schemeClr bg1 = lt1 = white)
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Model</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>GPU</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
        // Data row 1
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>YOLOv8n</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>RTX 4090</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
        // Data row 2
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>YOLOv8s</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>A100</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );

    let slide = make_slide_xml(&[table_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml,
        table_styles_xml,
        None,
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    // Header row should get accent1 (#4472C4) background from firstRow style
    assert_eq!(
        table.rows[0].cells[0].background,
        Some(Color::new(0x44, 0x72, 0xC4))
    );
    assert_eq!(
        table.rows[0].cells[1].background,
        Some(Color::new(0x44, 0x72, 0xC4))
    );

    // Header row text: explicit white (bg1→lt1→#FFFFFF) preserved, bold from style
    let header_run = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(p) => &p.runs[0],
        _ => panic!("Expected paragraph"),
    };
    assert_eq!(header_run.text, "Model");
    assert_eq!(header_run.style.color, Some(Color::new(0xFF, 0xFF, 0xFF)));
    assert_eq!(header_run.style.bold, Some(true));

    // Data row 1 (band index 0 → band1H) should get tinted accent1
    // accent1=(68,114,196) with tint 40%: (180,199,231)
    assert_eq!(
        table.rows[1].cells[0].background,
        Some(Color::new(180, 199, 231))
    );

    // Data row 2 (band index 1 → band2H, not defined) → no fill
    assert_eq!(table.rows[2].cells[0].background, None);

    // header_row_count should be 1
    assert_eq!(table.header_row_count, 1);
}

#[test]
fn table_fonts_resolve_style_refs_fallbacks_and_direct_overrides() {
    let table_styles_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="font-style">"#,
        r#"<a:tblStyle styleId="font-style" styleName="Font Style">"#,
        r#"<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef></a:tcTxStyle><a:tcStyle/></a:wholeTbl>"#,
        r#"<a:firstRow><a:tcTxStyle><a:fontRef idx="major"><a:prstClr val="black"/></a:fontRef></a:tcTxStyle><a:tcStyle/></a:firstRow>"#,
        r#"</a:tblStyle></a:tblStyleLst>"#,
    );
    let table_xml = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>"#,
        r#"<a:tblPr firstRow="1"><a:tableStyleId>font-style</a:tableStyleId></a:tblPr>"#,
        r#"<a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Specific</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Whole</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"><a:latin typeface=" Lato "/></a:rPr><a:t>Direct</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );
    let slide = make_slide_xml(&[table_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), " Gill Sans MT ", " Arial ");
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml,
        table_styles_xml,
        None,
    );

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let table = table_element(&first_fixed_page(&doc).elements[0]);
    let font_of = |row: usize| match &table.rows[row].cells[0].content[0] {
        Block::Paragraph(paragraph) => paragraph.runs[0].style.font_family.as_deref(),
        other => panic!("Expected paragraph, got {other:?}"),
    };

    assert_eq!(
        font_of(0),
        Some("Gill Sans MT"),
        "the specific first-row major fontRef beats the whole-table style"
    );
    assert_eq!(
        font_of(1),
        Some("Arial"),
        "the whole-table minor fontRef resolves the trimmed theme face"
    );
    assert_eq!(
        font_of(2),
        Some("Lato"),
        "a trimmed direct run typeface beats every table-style declaration"
    );
}

#[test]
fn unstyled_table_run_inherits_trimmed_theme_minor_font() {
    let table_xml = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>"#,
        r#"<a:tblPr/><a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Fallback</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );
    let slide = make_slide_xml(&[table_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Heading", " Arial ");
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        None,
    );

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let table = table_element(&first_fixed_page(&doc).elements[0]);
    let run = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("Expected paragraph, got {other:?}"),
    };
    assert_eq!(run.style.font_family.as_deref(), Some("Arial"));
}

#[test]
fn explicit_cell_no_fill_suppresses_table_style_fill_only() {
    let table_styles_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="style1">"#,
        r#"<a:tblStyle styleId="style1" styleName="Test">"#,
        r#"<a:wholeTbl><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>"#,
        r#"</a:tblStyle></a:tblStyleLst>"#,
    );
    let table_xml = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/>"#,
        r#"<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>"#,
        r#"<a:tblPr><a:tableStyleId>style1</a:tableStyleId></a:tblPr>"#,
        r#"<a:tblGrid><a:gridCol w="1828800"/></a:tblGrid><a:tr h="370840"><a:tc>"#,
        r#"<a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Transparent</a:t></a:r></a:p></a:txBody>"#,
        r#"<a:tcPr><a:noFill/></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );
    let slide = make_slide_xml(&[table_xml.to_string()]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml,
        table_styles_xml,
        None,
    );

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let table = table_element(&first_fixed_page(&doc).elements[0]);
    let cell = &table.rows[0].cells[0];
    assert_eq!(
        cell.background, None,
        "explicit noFill beats the style fill"
    );
    let run = match &cell.content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("Expected paragraph, got {other:?}"),
    };
    assert_eq!(
        run.style.bold,
        Some(true),
        "noFill does not suppress independent style text properties"
    );
}

/// A direct `<a:noFill/>` on one cell-border side suppresses only that side
/// of the table-style border. Each edge is independent: an absent left rule,
/// for example, must not discard the style's top, right, or bottom rules.
#[test]
fn direct_border_no_fill_suppresses_each_style_side_independently() {
    let table_styles_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="bordered">"#,
        r#"<a:tblStyle styleId="bordered" styleName="Bordered"><a:wholeTbl><a:tcStyle><a:tcBdr>"#,
        r#"<a:left><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:left>"#,
        r#"<a:right><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:right>"#,
        r#"<a:top><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:top>"#,
        r#"<a:bottom><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:bottom>"#,
        r#"<a:insideV><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:insideV>"#,
        r#"</a:tcBdr></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>"#,
    );
    let cell = |label: &str, side: &str| {
        format!(
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{label}</a:t></a:r></a:p></a:txBody><a:tcPr><a:{side} w="12700"><a:noFill/></a:{side}></a:tcPr></a:tc>"#,
        )
    };
    let table_xml = format!(
        concat!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/>"#,
            r#"<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>"#,
            r#"<a:tblPr><a:tableStyleId>bordered</a:tableStyleId></a:tblPr>"#,
            r#"<a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>"#,
            r#"<a:tr h="370840">{}{}{}{}</a:tr>"#,
            r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        ),
        cell("left", "lnL"),
        cell("top", "lnT"),
        cell("bottom", "lnB"),
        cell("right", "lnR"),
    );
    let slide = make_slide_xml(&[table_xml]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml,
        table_styles_xml,
        None,
    );

    let (doc, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let table = table_element(&first_fixed_page(&doc).elements[0]);
    let borders: Vec<&CellBorder> = table.rows[0]
        .cells
        .iter()
        .map(|cell| cell.border.as_ref().expect("unsuppressed sides remain"))
        .collect();

    assert!(borders[0].left.is_none(), "direct lnL noFill beats style");
    assert!(borders[1].top.is_none(), "direct lnT noFill beats style");
    assert!(borders[2].bottom.is_none(), "direct lnB noFill beats style");
    assert!(borders[3].right.is_none(), "direct lnR noFill beats style");
    for (index, border) in borders.iter().enumerate() {
        let remaining = [
            border.left.as_ref(),
            border.right.as_ref(),
            border.top.as_ref(),
            border.bottom.as_ref(),
        ]
        .into_iter()
        .flatten()
        .count();
        assert_eq!(remaining, 3, "cell {index} lost unrelated style borders");
    }
}

#[test]
fn test_pptx_table_without_table_styles_xml_still_works() {
    // Regular PPTX without tableStyles.xml should work fine
    let rows = format!(
        "{}{}",
        make_table_row(&["A1", "B1"]),
        make_table_row(&["A2", "B2"]),
    );
    let table_frame = make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows);
    let slide = make_slide_xml(&[table_frame]);
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells[0].background, None);
}

// ── Built-in table styles (issue #224) ─────────────────────────────────

#[test]
fn test_builtin_medium_style_2_no_accent() {
    let mut styles = TableStyleMap::new();
    table_styles::add_builtin_table_styles(&mut styles, &test_theme(), &test_color_map());

    let def = styles
        .get("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
        .expect("built-in Medium Style 2 must be generated");
    let first_row = def.first_row.as_ref().expect("firstRow region");
    assert_eq!(
        first_row.fill,
        Some(Color::new(0, 0, 0)),
        "solid dk1 header"
    );
    assert_eq!(
        first_row.text_color,
        Some(Color::new(255, 255, 255)),
        "lt1 header text"
    );
    let whole = def.whole_table.as_ref().expect("wholeTbl region");
    assert_eq!(
        whole.fill,
        Some(Color::new(230, 230, 230)),
        "dk1 tint 10% body fill"
    );
    let band = def.band1_h.as_ref().expect("band1H region");
    assert_eq!(
        band.fill,
        Some(Color::new(204, 204, 204)),
        "dk1 tint 20% band fill"
    );
    let inside_h = whole
        .borders
        .inside_h
        .as_ref()
        .expect("wholeTbl insideH border");
    assert_eq!(inside_h.color, Color::new(255, 255, 255), "lt1 grid lines");
    assert_eq!(inside_h.width, 1.0);
}

#[test]
fn test_builtin_medium_style_2_accent1() {
    let mut styles = TableStyleMap::new();
    table_styles::add_builtin_table_styles(&mut styles, &test_theme(), &test_color_map());

    let def = styles
        .get("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        .expect("built-in Medium Style 2 Accent 1 must be generated");
    let first_row = def.first_row.as_ref().unwrap();
    assert_eq!(first_row.fill, Some(Color::new(0x44, 0x72, 0xC4)));
    let whole = def.whole_table.as_ref().unwrap();
    assert_eq!(
        whole.fill,
        Some(Color::new(236, 241, 249)),
        "accent1 tint 10%"
    );
}

#[test]
fn test_file_defined_style_wins_over_builtin() {
    let mut styles = TableStyleMap::new();
    let custom = PptxTableStyleDef {
        first_row: Some(TableCellRegionStyle {
            fill: Some(Color::new(1, 2, 3)),
            ..TableCellRegionStyle::default()
        }),
        ..PptxTableStyleDef::default()
    };
    styles.insert("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}".to_string(), custom);
    table_styles::add_builtin_table_styles(&mut styles, &test_theme(), &test_color_map());

    let def = &styles["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"];
    assert_eq!(
        def.first_row.as_ref().unwrap().fill,
        Some(Color::new(1, 2, 3)),
        "a style defined in the file's tableStyles.xml must not be overwritten"
    );
}

#[test]
fn test_builtin_style_borders_applied_to_cells() {
    let mut styles = TableStyleMap::new();
    table_styles::add_builtin_table_styles(&mut styles, &test_theme(), &test_color_map());

    let mut table = Table {
        rows: (0..3)
            .map(|_| TableRow {
                minimum_height: None,
                cells: (0..3).map(|_| TableCell::default()).collect(),
                height: None,
            })
            .collect(),
        column_widths: vec![100.0, 100.0, 100.0],
        ..Table::default()
    };
    let props = PptxTableProps {
        style_id: Some("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}".to_string()),
        first_row: true,
        band_row: true,
        ..PptxTableProps::default()
    };
    table_styles::apply_table_style(&mut table, &props, &styles);

    // Interior cell gets white grid borders on all sides.
    let middle = table.rows[1].cells[1]
        .border
        .as_ref()
        .expect("interior cell should get style borders");
    for side in [&middle.top, &middle.bottom, &middle.left, &middle.right] {
        let side = side.as_ref().expect("all interior sides bordered");
        assert_eq!(side.color, Color::new(255, 255, 255));
    }
    // Header cells keep the firstRow solid fill.
    assert_eq!(table.rows[0].cells[0].background, Some(Color::new(0, 0, 0)));
    // Band row (first data row) uses the 20% tint.
    assert_eq!(
        table.rows[1].cells[0].background,
        Some(Color::new(204, 204, 204))
    );
    // Second data row falls back to the wholeTbl 10% tint — and still gets
    // the style grid borders.
    assert_eq!(
        table.rows[2].cells[0].background,
        Some(Color::new(230, 230, 230))
    );
    assert!(
        table.rows[2].cells[1].border.is_some(),
        "fallback (band2) cells must still get wholeTbl grid borders"
    );
}

/// A banded row's own `tcBdr` draws the rule between rows (issue #764).
///
/// Shape taken from "Light Style 2 - Accent 1": `wholeTbl` outlines the table
/// and switches the interior off, and `band1H` puts a rule on each banded row's
/// top and bottom. Only the band region supplies that rule, so dropping region
/// borders loses it and the rows run together.
#[test]
fn band_region_border_draws_the_rule_between_rows() {
    let rule = BorderSide {
        width: 0.5,
        color: Color::new(0x44, 0x72, 0xC4),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
    };
    let mut styles = TableStyleMap::new();
    styles.insert(
        "banded".to_string(),
        PptxTableStyleDef {
            whole_table: Some(TableCellRegionStyle {
                borders: table_styles::RegionBorders {
                    left: Some(rule.clone()),
                    right: Some(rule.clone()),
                    top: Some(rule.clone()),
                    bottom: Some(rule.clone()),
                    inside_h: None,
                    inside_v: None,
                },
                ..Default::default()
            }),
            band1_h: Some(TableCellRegionStyle {
                borders: table_styles::RegionBorders {
                    top: Some(rule.clone()),
                    bottom: Some(rule.clone()),
                    ..Default::default()
                },
                ..Default::default()
            }),
            ..Default::default()
        },
    );
    // Three rows so the banded row has an interior edge on both sides.
    let mut table = Table {
        rows: (0..3)
            .map(|_| TableRow {
                minimum_height: None,
                cells: vec![TableCell::default(), TableCell::default()],
                height: None,
            })
            .collect(),
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let props = PptxTableProps {
        style_id: Some("banded".to_string()),
        first_row: true,
        band_row: true,
        ..PptxTableProps::default()
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    let banded = table.rows[1].cells[0]
        .border
        .as_ref()
        .expect("the banded row is bordered");
    assert!(
        banded.top.is_some(),
        "band1H's top is the rule below the header"
    );
    assert_eq!(banded.top.as_ref().unwrap().color, rule.color);
    assert!(banded.bottom.is_some(), "band1H's bottom is drawn too");
    // insideV stays off: band1H names no left/right, so the vertical edge
    // between the two columns keeps wholeTbl's absent insideV. That edge is
    // the first cell's right and the second cell's left — the outer left and
    // right stay bordered by the grid, so those are not what this checks.
    assert!(
        banded.right.is_none(),
        "a band's horizontal rule must not switch the interior vertical on"
    );
    let banded_second = table.rows[1].cells[1]
        .border
        .as_ref()
        .expect("the second column is bordered too");
    assert!(
        banded_second.left.is_none(),
        "the same interior vertical edge, seen from the other cell"
    );
    assert!(
        banded.left.is_some() && banded_second.right.is_some(),
        "the table's outer left and right still come from the grid"
    );
}

#[test]
fn test_parse_region_borders_from_tc_bdr() {
    let body = r#"<a:wholeTbl><a:tcStyle><a:tcBdr><a:bottom><a:ln w="25400" cmpd="sng"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:bottom><a:insideH><a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:insideH></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>"#;
    let xml = make_table_style_xml(&[("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", body)]);
    let styles = table_styles::parse_table_styles_xml(&xml, &test_theme(), &test_color_map());
    let def = &styles["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"];
    let whole = def.whole_table.as_ref().unwrap();

    let bottom = whole.borders.bottom.as_ref().expect("bottom border parsed");
    assert_eq!(bottom.color, Color::new(0xFF, 0, 0));
    assert_eq!(bottom.width, 2.0, "25400 EMU = 2pt");

    let inside_h = whole.borders.inside_h.as_ref().expect("insideH parsed");
    assert_eq!(inside_h.color, Color::new(0, 0xFF, 0));
    assert_eq!(inside_h.width, 1.0);

    assert_eq!(
        whole.fill,
        Some(Color::new(0, 0, 0xFF)),
        "fill still parsed alongside borders"
    );
}

// ── Style references: fillRef / lnRef / tcTxStyle color (issue #674) ───
//
// The built-in styles PowerPoint ships express a region's fill and border
// through `<a:fillRef>`/`<a:lnRef>` into the theme's `fmtScheme`, not through
// the literal `<a:fill>`/`<a:ln>` the tests above use. The XML in this section
// is copied from `tests/fixtures/pptx/poi/table-with-theme.pptx`.

/// `make_theme_xml` emits no `<a:fmtScheme>`, so `line_styles` comes back
/// empty and no `lnRef` can resolve. The three widths below are the ones the
/// fixture's own theme declares — 6350, 12700 and 19050 EMU (0.5pt, 1pt,
/// 1.5pt).
fn theme_xml_with_line_styles() -> String {
    let mut color_xml = String::new();
    for (name, hex) in standard_theme_colors() {
        if name == "dk1" || name == "lt1" {
            color_xml.push_str(&format!(
                r#"<a:{name}><a:sysClr val="windowText" lastClr="{hex}"/></a:{name}>"#
            ));
        } else {
            color_xml.push_str(&format!(r#"<a:{name}><a:srgbClr val="{hex}"/></a:{name}>"#));
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Test">{color_xml}</a:clrScheme><a:fontScheme name="Test"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
    )
}

fn theme_with_line_styles() -> ThemeData {
    parse_theme_xml(&theme_xml_with_line_styles())
}

/// `default_color_map` maps every key onto itself, so `bg1` never reaches a
/// theme color. Real decks get the mapping from the slide master; this mirrors
/// the `<p:clrMap>` of the fixture these tests are drawn from.
fn fixture_color_map() -> ColorMapData {
    let aliases = [
        ("bg1", "lt1"),
        ("tx1", "dk1"),
        ("bg2", "lt2"),
        ("tx2", "dk2"),
        ("accent1", "accent1"),
        ("accent2", "accent2"),
        ("accent3", "accent3"),
        ("accent4", "accent4"),
        ("accent5", "accent5"),
        ("accent6", "accent6"),
        ("hlink", "hlink"),
        ("folHlink", "folHlink"),
    ]
    .into_iter()
    .map(|(key, target)| (key.to_string(), target.to_string()))
    .collect();
    ColorMapData { aliases }
}

fn first_row_of(body: &str) -> TableCellRegionStyle {
    let xml = make_table_style_xml(&[("{69012ECD-51FC-41F1-AA8D-1B2483CD663E}", body)]);
    let styles: TableStyleMap =
        table_styles::parse_table_styles_xml(&xml, &theme_with_line_styles(), &fixture_color_map());
    styles
        .get("{69012ECD-51FC-41F1-AA8D-1B2483CD663E}")
        .expect("style not found")
        .first_row
        .clone()
        .expect("firstRow missing")
}

#[test]
fn test_first_row_fill_ref_resolves_against_theme() {
    // Verbatim from the fixture: the header fill exists only as a fillRef.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="bg1"/></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:tcBdr/><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></a:tcStyle></a:firstRow>"#,
    ));

    assert_eq!(
        first_row.fill,
        Some(Color::new(0x44, 0x72, 0xC4)),
        "fillRef child color (accent1) is the header fill"
    );
}

#[test]
fn test_first_row_text_color_prefers_direct_child_over_font_ref() {
    // The fontRef carries its own black; the direct child is the real text
    // color. Taking the fontRef's leaves the header black-on-blue.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="bg1"/></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:tcBdr/><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></a:tcStyle></a:firstRow>"#,
    ));

    assert_eq!(
        first_row.text_color,
        Some(Color::new(0xFF, 0xFF, 0xFF)),
        "bg1 resolves to white, overriding the fontRef color"
    );
    assert_eq!(first_row.text_bold, Some(true), "bold still parsed");
}

#[test]
fn test_font_ref_color_still_used_when_no_direct_child() {
    // Triangulation: the fontRef color must remain the fallback, not be dropped.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:srgbClr val="00FF00"/></a:fontRef></a:tcTxStyle>"#,
        r#"<a:tcStyle/></a:firstRow>"#,
    ));

    assert_eq!(
        first_row.text_color,
        Some(Color::new(0, 0xFF, 0)),
        "fontRef color is the fallback when tcTxStyle has no direct color"
    );
}

#[test]
fn test_border_ln_ref_takes_width_from_theme_line_style() {
    // lnRef idx=1 → lnStyleLst[0] = 6350 EMU = 0.5pt. Without this the border
    // falls back to a flat 1pt, twice what PowerPoint draws.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcStyle><a:tcBdr>"#,
        r#"<a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left>"#,
        r#"</a:tcBdr></a:tcStyle></a:firstRow>"#,
    ));

    let left = first_row.borders.left.expect("left border missing");
    assert_eq!(left.width, 0.5, "6350 EMU is 0.5pt");
    assert_eq!(left.color, Color::new(0x44, 0x72, 0xC4));
}

#[test]
fn test_border_ln_ref_indexes_into_the_line_style_list() {
    // Triangulation: idx=3 must reach the third entry (19050 EMU = 1.5pt), so
    // the width cannot be a constant that happens to match idx=1.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcStyle><a:tcBdr>"#,
        r#"<a:top><a:lnRef idx="3"><a:schemeClr val="accent2"/></a:lnRef></a:top>"#,
        r#"</a:tcBdr></a:tcStyle></a:firstRow>"#,
    ));

    let top = first_row.borders.top.expect("top border missing");
    assert_eq!(top.width, 1.5, "19050 EMU is 1.5pt");
    assert_eq!(top.color, Color::new(0xED, 0x7D, 0x31), "accent2");
}

#[test]
fn test_explicit_ln_width_still_wins_over_ln_ref_absence() {
    // Triangulation: a literal <a:ln w=> must keep working unchanged.
    let first_row = first_row_of(concat!(
        r#"<a:firstRow><a:tcStyle><a:tcBdr>"#,
        r#"<a:bottom><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:bottom>"#,
        r#"</a:tcBdr></a:tcStyle></a:firstRow>"#,
    ));

    let bottom = first_row.borders.bottom.expect("bottom border missing");
    assert_eq!(bottom.width, 2.0, "25400 EMU is 2pt");
    assert_eq!(bottom.color, Color::new(0xFF, 0, 0));
}

#[test]
fn test_built_in_style_header_resolves_through_the_master_color_map() {
    // End-to-end for issue #674: the header's fill and text color exist only
    // as a fillRef and a `bg1` scheme name. `bg1` is not a theme color — it
    // reaches `lt1` through the master's <p:clrMap>, so a deck parsed without
    // that map renders the header black on white instead of white on blue.
    let table_styles_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{69012ECD-51FC-41F1-AA8D-1B2483CD663E}">"#,
        r#"<a:tblStyle styleId="{69012ECD-51FC-41F1-AA8D-1B2483CD663E}" styleName="Themed Style">"#,
        r#"<a:wholeTbl><a:tcStyle><a:tcBdr>"#,
        r#"<a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left>"#,
        r#"<a:right><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:right>"#,
        r#"<a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top>"#,
        r#"<a:bottom><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:bottom>"#,
        r#"</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>"#,
        r#"<a:firstRow>"#,
        r#"<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="bg1"/></a:tcTxStyle>"#,
        r#"<a:tcStyle><a:tcBdr/><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></a:tcStyle>"#,
        r#"</a:firstRow>"#,
        r#"</a:tblStyle></a:tblStyleLst>"#,
    );

    let table_xml = concat!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/>"#,
        r#"<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>"#,
        r#"<p:nvPr/></p:nvGraphicFramePr>"#,
        r#"<p:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#,
        r#"<a:tbl>"#,
        r#"<a:tblPr firstRow="1"><a:tableStyleId>{69012ECD-51FC-41F1-AA8D-1B2483CD663E}</a:tableStyleId></a:tblPr>"#,
        r#"<a:tblGrid><a:gridCol w="3657600"/></a:tblGrid>"#,
        // Neither run states a color; both must come from the style.
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Abc</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );

    // The clrMap of the fixture's master: bg1 → lt1.
    let master_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
        r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>"#,
        r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#,
        r#"</p:sldMaster>"#,
    );

    let slide = make_slide_xml(&[table_xml.to_string()]);
    let data = build_test_pptx_with_table_styles(
        SLIDE_CX,
        SLIDE_CY,
        &[slide],
        &theme_xml_with_line_styles(),
        table_styles_xml,
        Some(master_xml),
    );

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.rows[0].cells[0].background,
        Some(Color::new(0x44, 0x72, 0xC4)),
        "header fill comes from the fillRef"
    );
    let header_run = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        other => panic!("Expected paragraph, got {other:?}"),
    };
    assert_eq!(
        header_run.style.color,
        Some(Color::new(0xFF, 0xFF, 0xFF)),
        "bg1 reaches white only through the master color map"
    );

    let border = table.rows[0].cells[0]
        .border
        .as_ref()
        .expect("header cell has no border");
    let top = border.top.as_ref().expect("no top border");
    assert_eq!(top.width, 0.5, "lnRef idx=1 is 6350 EMU = 0.5pt");
    assert_eq!(top.color, Color::new(0x44, 0x72, 0xC4));
}

/// DrawingML composes the table style's regions: `wholeTbl` covers every cell
/// and a more specific region overrides only the properties it actually states.
/// Every stock "Medium Style 2" declares its off-band as
/// `<a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>` — a region that is
/// present but names no fill — and picking one region instead of composing them
/// left those rows transparent, showing the slide through between two banded
/// ones (issue #941).
#[test]
fn a_region_stating_no_fill_falls_through_to_whole_table() {
    let whole_fill = Color::new(0xD6, 0xF0, 0xF9);
    let band1_fill = Color::new(0xAD, 0xE1, 0xF2);
    let header_fill = Color::new(0x32, 0xB5, 0xDF);
    let mut styles: TableStyleMap = HashMap::new();
    styles.insert(
        "medium2".to_string(),
        PptxTableStyleDef {
            whole_table: Some(TableCellRegionStyle {
                fill: Some(whole_fill),
                fill_alpha: None,
                text_font_family: None,
                text_color: Some(Color::new(0x00, 0x00, 0x00)),
                text_bold: None,
                borders: Default::default(),
            }),
            first_row: Some(TableCellRegionStyle {
                fill: Some(header_fill),
                fill_alpha: None,
                text_font_family: None,
                text_color: Some(Color::new(0xFF, 0xFF, 0xFF)),
                text_bold: Some(true),
                borders: Default::default(),
            }),
            band1_h: Some(TableCellRegionStyle {
                fill: Some(band1_fill),
                fill_alpha: None,
                text_font_family: None,
                text_color: None,
                text_bold: None,
                borders: Default::default(),
            }),
            // The off-band: defined, but states nothing.
            band2_h: Some(TableCellRegionStyle::default()),
            ..Default::default()
        },
    );
    let props = PptxTableProps {
        style_id: Some("medium2".to_string()),
        first_row: true,
        band_row: true,
        ..Default::default()
    };
    let make_row = |text: &str| -> TableRow {
        TableRow {
            minimum_height: None,
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
                ..TableCell::default()
            }],
            height: Some(48.07),
        }
    };
    let mut table = Table {
        rows: vec![
            make_row("NØKKELPROSJEKTER"),
            make_row("Europium"),
            make_row("Bravo"),
            make_row("Gullfisk"),
        ],
        column_widths: vec![234.7],
        header_row_count: 1,
        ..Table::default()
    };

    table_styles::apply_table_style(&mut table, &props, &styles);

    assert_eq!(
        table.rows[0].cells[0].background,
        Some(header_fill),
        "firstRow states its own fill and keeps it"
    );
    assert_eq!(table.rows[1].cells[0].background, Some(band1_fill));
    assert_eq!(
        table.rows[2].cells[0].background,
        Some(whole_fill),
        "band2H states no fill, so wholeTbl's stands"
    );
    assert_eq!(table.rows[3].cells[0].background, Some(band1_fill));

    let run_color = |row: usize| -> Option<Color> {
        match &table.rows[row].cells[0].content[0] {
            Block::Paragraph(paragraph) => paragraph.runs[0].style.color,
            _ => panic!("expected a paragraph"),
        }
    };
    assert_eq!(
        run_color(0),
        Some(Color::new(0xFF, 0xFF, 0xFF)),
        "firstRow's text colour beats wholeTbl's"
    );
    assert_eq!(
        run_color(1),
        Some(Color::new(0x00, 0x00, 0x00)),
        "a band that names no text colour inherits wholeTbl's"
    );
}
