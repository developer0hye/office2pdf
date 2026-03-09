use super::*;

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
