use super::*;

#[test]
fn test_table_simple_2x2() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A1")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 3000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 2);
    assert_eq!(t.rows[0].cells.len(), 2);
    assert_eq!(t.rows[1].cells.len(), 2);

    let cell_text = |row: usize, col: usize| -> String {
        t.rows[row].cells[col]
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect::<String>()
    };
    assert_eq!(cell_text(0, 0), "A1");
    assert_eq!(cell_text(0, 1), "B1");
    assert_eq!(cell_text(1, 0), "A2");
    assert_eq!(cell_text(1, 1), "B2");
}

#[test]
fn test_table_column_widths_from_grid() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A"))),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B"))),
    ])])
    .set_grid(vec![2000, 3000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 150.0).abs() < 0.1,
        "Expected 150pt, got {}",
        t.column_widths[1]
    );
}

#[test]
fn test_table_tiny_autofit_grid_is_ignored() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Left column text")),
        ),
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Right column text")),
        ),
    ])])
    .set_grid(vec![100, 100]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert!(
        t.column_widths.is_empty(),
        "Expected tiny autofit grid widths to be ignored, got {:?}",
        t.column_widths
    );
}

#[test]
fn test_table_tiny_non_auto_grid_is_normalized_to_table_width() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A"))),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B"))),
    ])])
    .set_grid(vec![100, 100])
    .width(400, docx_rs::WidthType::Dxa);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!((t.column_widths[0] - 10.0).abs() < 0.1);
    assert!((t.column_widths[1] - 10.0).abs() < 0.1);
}

#[test]
fn test_table_single_column_placeholder_grid_uses_table_width() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Long single cell")),
        ),
    ])])
    .set_grid(vec![100])
    .width(9360, docx_rs::WidthType::Dxa);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 1);
    assert!(
        (t.column_widths[0] - 468.0).abs() < 0.2,
        "Expected table width normalization to 468pt, got {}",
        t.column_widths[0]
    );
}

#[test]
fn test_table_column_widths_from_cell_widths_without_grid() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A")))
            .width(2000, docx_rs::WidthType::Dxa),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B")))
            .width(3000, docx_rs::WidthType::Dxa),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 2);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 150.0).abs() < 0.1,
        "Expected 150pt, got {}",
        t.column_widths[1]
    );
}

#[test]
fn test_table_column_widths_from_spanned_cell_widths_without_grid() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
            )
            .grid_span(2)
            .width(4000, docx_rs::WidthType::Dxa),
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C")))
            .width(2000, docx_rs::WidthType::Dxa),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.column_widths.len(), 3);
    assert!(
        (t.column_widths[0] - 100.0).abs() < 0.1,
        "Expected first merged column to be 100pt, got {}",
        t.column_widths[0]
    );
    assert!(
        (t.column_widths[1] - 100.0).abs() < 0.1,
        "Expected second merged column to be 100pt, got {}",
        t.column_widths[1]
    );
    assert!(
        (t.column_widths[2] - 100.0).abs() < 0.1,
        "Expected final column to be 100pt, got {}",
        t.column_widths[2]
    );
}

#[test]
fn test_scan_table_headers_counts_only_leading_rows() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>D1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Ignored</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Only body</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
</w:document>"#;

    let headers = scan_table_headers(document_xml);

    assert_eq!(headers.len(), 2);
    assert_eq!(headers[0].repeat_rows, 2);
    assert_eq!(headers[1].repeat_rows, 0);
}

#[test]
fn test_table_header_rows_from_raw_docx_xml() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tblPr/>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Body A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Body B</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.header_row_count, 1);
}

#[test]
fn test_table_default_cell_margins_from_table_property() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell"))),
    ])])
    .margins(docx_rs::TableCellMargins::new().margin(40, 60, 20, 80));

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(
        t.default_cell_padding,
        Some(Insets {
            top: 2.0,
            right: 3.0,
            bottom: 1.0,
            left: 4.0,
        })
    );
    assert!(t.rows[0].cells[0].padding.is_none());
}

#[test]
fn test_table_cell_margins_override_table_defaults() {
    let mut cell = docx_rs::TableCell::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")));
    cell.property = docx_rs::TableCellProperty::new()
        .margin_top(100, docx_rs::WidthType::Dxa)
        .margin_left(120, docx_rs::WidthType::Dxa);

    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![cell])])
        .margins(docx_rs::TableCellMargins::new().margin(20, 40, 60, 80));

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(
        t.default_cell_padding,
        Some(Insets {
            top: 1.0,
            right: 2.0,
            bottom: 3.0,
            left: 4.0,
        })
    );
    assert_eq!(
        t.rows[0].cells[0].padding,
        Some(Insets {
            top: 5.0,
            right: 2.0,
            bottom: 3.0,
            left: 6.0,
        })
    );
}

#[test]
fn test_table_alignment_from_table_property() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
        ),
    ])])
    .align(docx_rs::TableAlignmentType::Center);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.alignment, Some(Alignment::Center));
}

#[test]
fn test_table_cell_with_formatted_text() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bold").bold())
                .add_run(docx_rs::Run::new().add_text(" and italic").italic()),
        ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let para = match &cell.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph in cell"),
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Bold");
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert_eq!(para.runs[1].text, " and italic");
    assert_eq!(para.runs[1].style.italic, Some(true));
}

#[test]
fn test_table_colspan_via_grid_span() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
                )
                .grid_span(2),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows[0].cells.len(), 1);
    assert_eq!(t.rows[0].cells[0].col_span, 2);
    assert_eq!(t.rows[1].cells.len(), 2);
    assert_eq!(t.rows[1].cells[0].col_span, 1);
}

#[test]
fn test_table_rowspan_via_vmerge() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tall")),
                )
                .vertical_merge(docx_rs::VMergeType::Restart),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 3);
    let tall_cell = &t.rows[0].cells[0];
    assert_eq!(tall_cell.row_span, 3);
    assert_eq!(t.rows[1].cells.len(), 1);
    assert_eq!(t.rows[2].cells.len(), 1);
}

#[test]
fn test_table_exact_row_height_and_cell_vertical_align() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
                )
                .vertical_align(docx_rs::VAlignType::Center),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Peer")),
            ),
        ])
        .row_height(36.0)
        .height_rule(docx_rs::HeightRule::Exact),
    ])
    .set_grid(vec![2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows[0].height, Some(36.0));
    assert_eq!(
        t.rows[0].cells[0].vertical_align,
        Some(CellVerticalAlign::Center)
    );
}

#[test]
fn test_table_cell_background_color() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Red bg")),
            )
            .shading(docx_rs::Shading::new().fill("FF0000")),
        docx_rs::TableCell::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("No bg")),
        ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows[0].cells[0].background, Some(Color::new(255, 0, 0)));
    assert!(t.rows[0].cells[1].background.is_none());
}

#[test]
fn test_table_cell_borders() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bordered")),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("FF0000"),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                    .size(8)
                    .color("0000FF"),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");
    let top = border.top.as_ref().expect("Expected top border");
    assert!(
        (top.width - 2.0).abs() < 0.01,
        "Expected 2pt, got {}",
        top.width
    );
    assert_eq!(top.color, Color::new(255, 0, 0));

    let bottom = border.bottom.as_ref().expect("Expected bottom border");
    assert!(
        (bottom.width - 1.0).abs() < 0.01,
        "Expected 1pt, got {}",
        bottom.width
    );
    assert_eq!(bottom.color, Color::new(0, 0, 255));
}

#[test]
fn test_table_cell_border_styles() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Styled borders")),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("000000")
                    .border_type(docx_rs::BorderType::Dashed),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                    .size(8)
                    .color("0000FF")
                    .border_type(docx_rs::BorderType::Dotted),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Left)
                    .size(12)
                    .color("FF0000")
                    .border_type(docx_rs::BorderType::DotDash),
            )
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Right)
                    .size(16)
                    .color("00FF00")
                    .border_type(docx_rs::BorderType::Double),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");
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
fn test_table_cell_solid_border_default_style() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Solid")))
            .set_border(
                docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                    .size(16)
                    .color("000000"),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);
    let cell = &t.rows[0].cells[0];
    let border = cell.border.as_ref().expect("Expected cell border");
    let top = border.top.as_ref().expect("Expected top border");
    assert_eq!(top.style, BorderLineStyle::Solid, "Single -> Solid");
}

#[test]
fn test_table_cell_with_multiple_paragraphs() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 1")),
            )
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 2")),
            ),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let cell = &t.rows[0].cells[0];
    let paras: Vec<&str> = cell
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) if !p.runs.is_empty() => Some(p.runs[0].text.as_str()),
            _ => None,
        })
        .collect();
    assert!(paras.contains(&"Para 1"), "Expected 'Para 1' in cell");
    assert!(paras.contains(&"Para 2"), "Expected 'Para 2' in cell");
}

#[test]
fn test_table_with_paragraph_before_and_after() {
    let data = {
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before")),
            )
            .add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")),
                ),
            ])]))
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After")),
            );
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    };

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let blocks = all_blocks(&doc);

    assert!(
        blocks.len() >= 3,
        "Expected at least 3 blocks, got {}",
        blocks.len()
    );
    assert!(matches!(&blocks[0], Block::Paragraph(_)));
    let has_table = blocks.iter().any(|b| matches!(b, Block::Table(_)));
    assert!(has_table, "Expected a Table block");
}

#[test]
fn test_table_colspan_and_rowspan_combined() {
    let table = docx_rs::Table::new(vec![
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Big")),
                )
                .grid_span(2)
                .vertical_merge(docx_rs::VMergeType::Restart),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C1")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new())
                .grid_span(2)
                .vertical_merge(docx_rs::VMergeType::Continue),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C2")),
            ),
        ]),
        docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A3")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C3")),
            ),
        ]),
    ])
    .set_grid(vec![2000, 2000, 2000]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    let big_cell = &t.rows[0].cells[0];
    assert_eq!(big_cell.col_span, 2, "Expected colspan=2");
    assert_eq!(big_cell.row_span, 2, "Expected rowspan=2");
    assert_eq!(t.rows[1].cells.len(), 1);
    assert_eq!(t.rows[2].cells.len(), 3);
}

#[test]
fn test_table_empty_cells() {
    let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
        docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
    ])]);

    let data = build_docx_with_table(table);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let t = first_table(&doc);

    assert_eq!(t.rows.len(), 1);
    assert_eq!(t.rows[0].cells.len(), 2);
    for cell in &t.rows[0].cells {
        assert_eq!(cell.col_span, 1);
        assert_eq!(cell.row_span, 1);
    }
}
