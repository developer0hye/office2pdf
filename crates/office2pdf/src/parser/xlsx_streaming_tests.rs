use super::*;

fn build_xlsx_with_rows(sheet_name: &str, num_rows: u32, num_cols: u32) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name(sheet_name);
    for row in 1..=num_rows {
        for col in 1..=num_cols {
            sheet
                .get_cell_mut((col, row))
                .set_value(format!("R{row}C{col}"));
        }
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_parse_streaming_creates_chunks() {
    let data = build_xlsx_with_rows("Sheet1", 5, 2);
    let parser = XlsxParser;
    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 2)
        .unwrap();

    assert_eq!(
        chunks.len(),
        3,
        "5 rows with chunk_size=2 should yield 3 chunks"
    );

    let tp0 = get_table_page(&chunks[0], 0);
    assert_eq!(tp0.table.rows.len(), 2);
    assert_eq!(cell_text(&tp0.table.rows[0].cells[0]), "R1C1");
    assert_eq!(cell_text(&tp0.table.rows[1].cells[0]), "R2C1");

    let tp1 = get_table_page(&chunks[1], 0);
    assert_eq!(tp1.table.rows.len(), 2);
    assert_eq!(cell_text(&tp1.table.rows[0].cells[0]), "R3C1");

    let tp2 = get_table_page(&chunks[2], 0);
    assert_eq!(tp2.table.rows.len(), 1);
    assert_eq!(cell_text(&tp2.table.rows[0].cells[0]), "R5C1");
}

#[test]
fn test_parse_streaming_single_chunk_for_small_sheet() {
    let data = build_xlsx_with_rows("Data", 3, 1);
    let parser = XlsxParser;
    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 10)
        .unwrap();

    assert_eq!(
        chunks.len(),
        1,
        "3 rows with chunk_size=10 should yield 1 chunk"
    );
    let tp = get_table_page(&chunks[0], 0);
    assert_eq!(tp.table.rows.len(), 3);
}

#[test]
fn test_parse_streaming_preserves_column_widths() {
    let data = build_xlsx_with_rows("Sheet1", 4, 3);
    let parser = XlsxParser;
    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 2)
        .unwrap();

    let tp0 = get_table_page(&chunks[0], 0);
    let tp1 = get_table_page(&chunks[1], 0);
    assert_eq!(tp0.table.column_widths.len(), tp1.table.column_widths.len());
    assert_eq!(tp0.table.column_widths, tp1.table.column_widths);
}

#[test]
fn test_parse_streaming_respects_sheet_filter() {
    let data = build_xlsx_multi_sheet(&[
        ("Sheet1", &[("A1", "s1")]),
        ("Sheet2", &[("A1", "s2"), ("A2", "s2b")]),
    ]);
    let parser = XlsxParser;
    let opts = ConvertOptions {
        sheet_names: Some(vec!["Sheet2".to_string()]),
        ..Default::default()
    };
    let (chunks, _warnings) = parser.parse_streaming(&data, &opts, 10).unwrap();

    assert_eq!(chunks.len(), 1, "Only Sheet2 should be included");
    let tp = get_table_page(&chunks[0], 0);
    assert_eq!(tp.name, "Sheet2");
}

#[test]
fn test_parse_streaming_multi_sheet() {
    let data = build_xlsx_multi_sheet(&[
        ("Sheet1", &[("A1", "a"), ("A2", "b"), ("A3", "c")]),
        ("Sheet2", &[("A1", "x"), ("A2", "y")]),
    ]);
    let parser = XlsxParser;
    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 2)
        .unwrap();

    assert_eq!(chunks.len(), 3, "Sheet1→2 chunks + Sheet2→1 chunk");
}

#[test]
fn test_parse_streaming_empty_sheet_skipped() {
    let data = build_xlsx_bytes("Empty", &[]);
    let parser = XlsxParser;
    let (chunks, _warnings) = parser
        .parse_streaming(&data, &ConvertOptions::default(), 10)
        .unwrap();

    assert_eq!(chunks.len(), 0, "Empty sheet should be skipped");
}
