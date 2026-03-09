use super::*;

fn make_table_graphic_frame(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    col_widths_emu: &[i64],
    rows_xml: &str,
) -> String {
    let mut grid = String::new();
    for width in col_widths_emu {
        grid.push_str(&format!(r#"<a:gridCol w="{width}"/>"#));
    }
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{rows_xml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

fn make_table_row(cells: &[&str]) -> String {
    let mut xml = String::from(r#"<a:tr h="370840">"#);
    for text in cells {
        xml.push_str(&format!(
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#
        ));
    }
    xml.push_str("</a:tr>");
    xml
}

fn table_element(elem: &FixedElement) -> &Table {
    match &elem.kind {
        FixedElementKind::Table(table) => table,
        _ => panic!("Expected Table, got {:?}", elem.kind),
    }
}

#[test]
fn test_slide_with_basic_table() {
    let rows = format!(
        "{}{}",
        make_table_row(&["A1", "B1"]),
        make_table_row(&["A2", "B2"]),
    );
    let table_frame =
        make_table_graphic_frame(914400, 914400, 3657600, 1828800, &[1828800, 1828800], &rows);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let elem = &page.elements[0];
    assert!((elem.x - 72.0).abs() < 0.1);
    assert!((elem.y - 72.0).abs() < 0.1);

    let table = table_element(elem);
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 144.0).abs() < 0.1);

    let first_cell = &table.rows[0].cells[0];
    assert_eq!(first_cell.content.len(), 1);
    match &first_cell.content[0] {
        Block::Paragraph(paragraph) => assert_eq!(paragraph.runs[0].text, "A1"),
        _ => panic!("Expected paragraph in cell"),
    }

    match &table.rows[1].cells[1].content[0] {
        Block::Paragraph(paragraph) => assert_eq!(paragraph.runs[0].text, "B2"),
        _ => panic!("Expected paragraph in cell"),
    }
}

#[test]
fn test_slide_table_scales_geometry_to_graphic_frame_extent() {
    let rows_xml = format!(
        "{}{}",
        make_table_row(&["A1", "B1"]),
        make_table_row(&["A2", "B2"]),
    );
    let table_frame = make_table_graphic_frame(
        914400,
        914400,
        3_657_600,
        1_483_360,
        &[914_400, 914_400],
        &rows_xml,
    );
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];
    let table = table_element(elem);

    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 144.0).abs() < 0.1);
    assert!((table.column_widths.iter().sum::<f64>() - elem.width).abs() < 0.1);
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].height, Some(58.4));
    assert_eq!(table.rows[1].height, Some(58.4));
    assert!(
        (table
            .rows
            .iter()
            .map(|row| row.height.unwrap_or(0.0))
            .sum::<f64>()
            - elem.height)
            .abs()
            < 0.1
    );
}

#[test]
fn test_slide_table_reads_column_widths_from_gridcol_with_extensions() {
    let rows_xml = make_table_row(&["A1", "B1"]);
    let table_frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="370840"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="1"/></a:ext></a:extLst></a:gridCol><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="2"/></a:ext></a:extLst></a:gridCol></a:tblGrid>"#.to_string()
        + &rows_xml
        + r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(table.column_widths.len(), 2);
    assert!((table.column_widths[0] - 72.0).abs() < 0.1);
    assert!((table.column_widths[1] - 72.0).abs() < 0.1);
}

#[test]
fn test_slide_table_cell_anchor_maps_to_vertical_alignment() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Centered</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bottom</a:t></a:r></a:p></a:txBody><a:tcPr anchor="b"/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame =
        make_table_graphic_frame(0, 0, 1_828_800, 370_840, &[914_400, 914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.rows[0].cells[0].vertical_align,
        Some(crate::ir::CellVerticalAlign::Center)
    );
    assert_eq!(
        table.rows[0].cells[1].vertical_align,
        Some(crate::ir::CellVerticalAlign::Bottom)
    );
}

#[test]
fn test_slide_table_cell_margins_map_to_padding() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Padded</a:t></a:r></a:p></a:txBody><a:tcPr marL="76200" marR="76200" marT="38100" marB="38100"/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.rows[0].cells[0].padding,
        Some(crate::ir::Insets {
            top: 3.0,
            right: 6.0,
            bottom: 3.0,
            left: 6.0,
        })
    );
}

#[test]
fn test_slide_table_uses_powerpoint_default_cell_padding() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>DefaultPadding</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(
        table.default_cell_padding,
        Some(crate::ir::Insets {
            top: 3.6,
            right: 7.2,
            bottom: 3.6,
            left: 7.2,
        })
    );
    assert_eq!(table.rows[0].cells[0].padding, None);
    assert!(table.use_content_driven_row_heights);
}

#[test]
fn test_slide_table_coalesces_adjacent_runs_with_same_style() {
    let rows_xml = concat!(
        r#"<a:tr h="370840">"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t>YOLOv8n + </a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100" err="1"><a:latin typeface="Arial"/></a:rPr><a:t>topk filtering on gpu(</a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>K</a:t></a:r>"#,
        r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t> = 100)</a:t></a:r>"#,
        r#"</a:p></a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    let paragraph = match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected paragraph, got {other:?}"),
    };

    assert_eq!(paragraph.runs.len(), 3);
    assert_eq!(paragraph.runs[0].text, "YOLOv8n + topk filtering on gpu(");
    assert_eq!(paragraph.runs[1].text, "K");
    assert_eq!(paragraph.runs[2].text, "\u{00A0}= 100)");
    assert_eq!(paragraph.runs[1].style.italic, Some(true));
}

#[test]
fn test_slide_table_cell_bulleted_paragraphs_group_into_list() {
    let rows_xml = concat!(
        r#"<a:tr h="740000">"#,
        r#"<a:tc><a:txBody><a:bodyPr/>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>First bullet</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Second bullet</a:t></a:r></a:p>"#,
        r#"</a:txBody><a:tcPr/></a:tc>"#,
        r#"</a:tr>"#,
    );
    let table_frame = make_table_graphic_frame(0, 0, 914_400, 740_000, &[914_400], rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    assert_eq!(table.rows[0].cells[0].content.len(), 1);

    let list = match &table.rows[0].cells[0].content[0] {
        Block::List(list) => list,
        other => panic!("Expected List block, got {other:?}"),
    };
    assert_eq!(list.kind, crate::ir::ListKind::Unordered);
    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
    assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
}

#[test]
fn test_slide_table_with_merged_cells() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");
    rows_xml.push_str(&make_table_row(&["C1", "C2"]));

    let table_frame =
        make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(table.rows[0].cells.len(), 2);
    assert_eq!(table.rows[0].cells[0].col_span, 2);
    assert_eq!(table.rows[0].cells[1].col_span, 0);
    assert_eq!(table.rows[1].cells[0].col_span, 1);
    assert_eq!(table.rows[1].cells[1].col_span, 1);
}

#[test]
fn test_slide_table_with_vertical_merge() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>VMerged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc vMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame =
        make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    assert_eq!(table.rows[0].cells[0].row_span, 2);
    assert_eq!(table.rows[1].cells[0].row_span, 0);
}

#[test]
fn test_slide_table_with_formatted_text() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" b="1" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Bold Red</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    match &table.rows[0].cells[0].content[0] {
        Block::Paragraph(paragraph) => {
            assert_eq!(paragraph.runs[0].text, "Bold Red");
            assert_eq!(paragraph.runs[0].style.bold, Some(true));
            assert_eq!(paragraph.runs[0].style.font_size, Some(18.0));
            assert_eq!(paragraph.runs[0].style.color, Some(Color::new(255, 0, 0)));
        }
        _ => panic!("Expected paragraph in cell"),
    }
}

#[test]
fn test_slide_table_with_cell_background() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Filled</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    assert_eq!(
        table.rows[0].cells[0].background,
        Some(Color::new(0, 255, 0))
    );
}

#[test]
fn test_slide_table_with_cell_borders() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bordered</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB></a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);

    let border = table.rows[0].cells[0]
        .border
        .as_ref()
        .expect("Expected border");
    assert!(border.left.is_some());
    assert!(border.right.is_some());
    assert!(border.top.is_some());
    assert!(border.bottom.is_some());
    let left = border.left.as_ref().unwrap();
    assert!((left.width - 1.0).abs() < 0.1);
    assert_eq!(left.color, Color::new(0, 0, 0));
}

#[test]
fn test_slide_table_cell_border_dash_styles() {
    let mut rows_xml = String::new();
    rows_xml.push_str(r#"<a:tr h="370840">"#);
    rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Dashed</a:t></a:r></a:p></a:txBody><a:tcPr>"#);
    rows_xml.push_str(r#"<a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:lnT>"#);
    rows_xml.push_str(r#"<a:lnB w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dot"/></a:lnB>"#);
    rows_xml.push_str(r#"</a:tcPr></a:tc>"#);
    rows_xml.push_str("</a:tr>");

    let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
    let slide = make_slide_xml(&[table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let table = table_element(&page.elements[0]);
    let border = table.rows[0].cells[0]
        .border
        .as_ref()
        .expect("Expected border");
    assert_eq!(
        border.top.as_ref().expect("Expected top border").style,
        BorderLineStyle::Dashed
    );
    assert_eq!(
        border
            .bottom
            .as_ref()
            .expect("Expected bottom border")
            .style,
        BorderLineStyle::Dotted
    );
}

#[test]
fn test_slide_table_coexists_with_shapes() {
    let text_box = make_text_box(0, 0, 914400, 457200, "Header");
    let rows = make_table_row(&["Cell"]);
    let table_frame = make_table_graphic_frame(0, 914400, 914400, 370840, &[914400], &rows);
    let slide = make_slide_xml(&[text_box, table_frame]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2);
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Table(_)));
}
