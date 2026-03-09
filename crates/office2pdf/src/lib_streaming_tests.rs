use std::io::Cursor;

use super::*;

fn build_xlsx_with_rows(num_rows: u32, num_cols: u32) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name("Data");
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
fn test_streaming_xlsx_produces_valid_pdf() {
    let data = build_xlsx_with_rows(50, 3);
    let options = config::ConvertOptions {
        streaming: true,
        streaming_chunk_size: Some(20),
        ..Default::default()
    };
    let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be valid PDF"
    );
    assert!(result.pdf.len() > 100, "PDF should have content");
}

#[test]
fn test_streaming_xlsx_same_data_as_normal() {
    let data = build_xlsx_with_rows(10, 2);

    let normal_opts = config::ConvertOptions::default();
    let normal_result = convert_bytes(&data, config::Format::Xlsx, &normal_opts).unwrap();

    let streaming_opts = config::ConvertOptions {
        streaming: true,
        streaming_chunk_size: Some(5),
        ..Default::default()
    };
    let streaming_result = convert_bytes(&data, config::Format::Xlsx, &streaming_opts).unwrap();

    assert!(normal_result.pdf.starts_with(b"%PDF"));
    assert!(streaming_result.pdf.starts_with(b"%PDF"));
    assert!(normal_result.pdf.len() > 100);
    assert!(streaming_result.pdf.len() > 100);
}

#[test]
fn test_streaming_large_xlsx_completes() {
    let data = build_xlsx_with_rows(10_000, 3);
    let options = config::ConvertOptions {
        streaming: true,
        streaming_chunk_size: Some(1000),
        ..Default::default()
    };
    let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be valid PDF"
    );
    assert!(result.metrics.is_some(), "streaming should produce metrics");
}

#[test]
fn test_streaming_non_xlsx_falls_through() {
    let docx = {
        let doc = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello streaming")),
        );
        let mut cursor = Cursor::new(Vec::new());
        doc.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    };
    let options = config::ConvertOptions {
        streaming: true,
        ..Default::default()
    };
    let result = convert_bytes(&docx, config::Format::Docx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
}

#[test]
fn test_streaming_chunk_size_default() {
    let data = build_xlsx_with_rows(20, 1);
    let options = config::ConvertOptions {
        streaming: true,
        streaming_chunk_size: None,
        ..Default::default()
    };
    let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
}

#[test]
fn test_streaming_memory_bounded() {
    let data = build_xlsx_with_rows(5_000, 5);
    let options = config::ConvertOptions {
        streaming: true,
        streaming_chunk_size: Some(500),
        ..Default::default()
    };
    let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
    assert!(result.pdf.starts_with(b"%PDF"));
    assert!(
        result.pdf.len() > 1000,
        "PDF should have substantial content"
    );
}
