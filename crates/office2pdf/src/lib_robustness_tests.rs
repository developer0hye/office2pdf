use super::test_support::build_test_docx;
use super::*;
use crate::ir::*;

fn build_pptx_with_broken_slide() -> Vec<u8> {
    use std::io::{Cursor, Write};

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst></p:presentation>"#,
    )
    .unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Valid slide content</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    )
    .unwrap();

    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_pptx_broken_slide_emits_warning_and_produces_pdf() {
    let pptx_bytes = build_pptx_with_broken_slide();
    let result = convert_bytes(&pptx_bytes, Format::Pptx, &ConvertOptions::default()).unwrap();

    assert!(
        !result.pdf.is_empty(),
        "Should produce PDF despite broken slide"
    );
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Output should be valid PDF"
    );
    assert!(
        !result.warnings.is_empty(),
        "Should emit warning for broken slide"
    );
    let warning_text = result.warnings[0].to_string();
    assert!(
        warning_text.contains("slide") || warning_text.contains("Slide"),
        "Warning should mention the problematic slide: {warning_text}"
    );
}

#[test]
fn test_edge_empty_docx_produces_valid_pdf() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new();
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Empty DOCX should produce valid PDF"
    );
}

#[test]
fn test_edge_empty_xlsx_produces_valid_pdf() {
    use std::io::Cursor;

    let book = umya_spreadsheet::new_file();
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Empty XLSX should produce valid PDF"
    );
}

#[test]
fn test_edge_empty_pptx_produces_valid_pdf() {
    use std::io::{Cursor, Write};

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::FileOptions::default();
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
    )
    .unwrap();
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst/></p:presentation>"#,
    )
    .unwrap();
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    )
    .unwrap();
    let data = zip.finish().unwrap().into_inner();
    let result = convert_bytes(&data, Format::Pptx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Empty PPTX should produce valid PDF"
    );
}

#[test]
fn test_edge_long_paragraph_no_panic() {
    use std::io::Cursor;

    let long_text: String = "Lorem ipsum dolor sit amet. ".repeat(400);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(&long_text)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Long paragraph should produce valid PDF"
    );
}

#[test]
fn test_edge_large_table_no_panic() {
    use std::io::Cursor;

    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        for row in 1..=100u32 {
            for col in 1..=20u32 {
                let coord = format!("{}{}", (b'A' + ((col - 1) % 26) as u8) as char, row);
                sheet
                    .get_cell_mut(coord.as_str())
                    .set_value(format!("R{row}C{col}"));
            }
        }
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Large table should produce valid PDF"
    );
}

#[test]
fn test_edge_corrupted_docx_returns_error() {
    let data = b"not a valid ZIP file at all";
    let result = convert_bytes(data, Format::Docx, &ConvertOptions::default());
    assert!(result.is_err(), "Corrupted DOCX should return an error");
    let err = result.unwrap_err();
    match err {
        ConvertError::Parse(msg) => {
            assert!(!msg.is_empty(), "Error message should not be empty");
        }
        _ => panic!("Expected Parse error for corrupted DOCX, got {err:?}"),
    }
}

#[test]
fn test_edge_corrupted_xlsx_returns_error() {
    let data = b"this is not an xlsx file";
    let result = convert_bytes(data, Format::Xlsx, &ConvertOptions::default());
    assert!(result.is_err(), "Corrupted XLSX should return an error");
}

#[test]
fn test_edge_corrupted_pptx_returns_error() {
    let data = b"garbage data that is not a pptx";
    let result = convert_bytes(data, Format::Pptx, &ConvertOptions::default());
    assert!(result.is_err(), "Corrupted PPTX should return an error");
}

#[test]
fn test_edge_truncated_zip_returns_error() {
    let full_data = build_test_docx();
    let truncated = &full_data[..full_data.len() / 2];
    let result = convert_bytes(truncated, Format::Docx, &ConvertOptions::default());
    assert!(result.is_err(), "Truncated DOCX should return an error");
}

#[test]
fn test_edge_unicode_cjk_text() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("中文测试 日本語テスト 한국어 테스트")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "CJK text should produce valid PDF"
    );
}

#[test]
fn test_edge_unicode_emoji_text() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello 🌍🎉💡 World")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Emoji text should produce valid PDF"
    );
}

#[test]
fn test_edge_unicode_rtl_text() {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("مرحبا بالعالم")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "RTL text should produce valid PDF"
    );
}

#[test]
fn test_bidi_docx_rtl_direction_in_typst() {
    use crate::parser::Parser;
    use std::io::Cursor;

    let mut bidi_para =
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("مرحبا بالعالم"));
    bidi_para.property = docx_rs::ParagraphProperty::new().bidi(true);
    let docx = docx_rs::Docx::new().add_paragraph(bidi_para);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = crate::parser::docx::DocxParser;
    let (doc, _) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let flow = match &doc.pages[0] {
        Page::Flow(flow_page) => flow_page,
        _ => panic!("Expected FlowPage"),
    };
    let para = flow.content.iter().find_map(|block| match block {
        Block::Paragraph(paragraph) => Some(paragraph),
        _ => None,
    });
    assert_eq!(
        para.unwrap().style.direction,
        Some(TextDirection::Rtl),
        "Bidi DOCX paragraph should parse as RTL"
    );

    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "RTL bidi text should produce valid PDF"
    );
}

#[test]
fn test_bidi_mixed_rtl_ltr_docx() {
    use std::io::Cursor;

    let mut bidi_para =
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("العدد 42 والعدد 100"));
    bidi_para.property = docx_rs::ParagraphProperty::new().bidi(true);
    let docx = docx_rs::Docx::new().add_paragraph(bidi_para).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("This is English")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "Mixed RTL/LTR DOCX should produce valid PDF"
    );
}

#[test]
fn test_edge_image_only_docx() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Image(ImageData {
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                width: Some(100.0),
                height: Some(100.0),
                crop: None,
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let _result = render_document(&doc);
}

#[test]
fn test_is_ole2_with_magic_bytes() {
    let ole2_magic: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut data = ole2_magic.to_vec();
    data.extend_from_slice(&[0x00; 100]);
    assert!(is_ole2(&data));
}

#[test]
fn test_is_ole2_with_zip_bytes() {
    let zip_data = [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
    assert!(!is_ole2(&zip_data));
}

#[test]
fn test_is_ole2_with_short_data() {
    let short = [0xD0, 0xCF, 0x11];
    assert!(!is_ole2(&short));
}

#[test]
fn test_is_ole2_with_empty_data() {
    assert!(!is_ole2(&[]));
}

#[test]
fn test_ole2_bytes_return_unsupported_encryption() {
    let ole2_magic: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut data = ole2_magic.to_vec();
    data.extend_from_slice(&[0x00; 100]);

    let err = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap_err();
    assert!(
        matches!(err, ConvertError::UnsupportedEncryption),
        "Expected UnsupportedEncryption, got: {err:?}"
    );
}

#[test]
fn test_ole2_bytes_return_unsupported_encryption_xlsx() {
    let ole2_magic: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut data = ole2_magic.to_vec();
    data.extend_from_slice(&[0x00; 100]);

    let err = convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap_err();
    assert!(
        matches!(err, ConvertError::UnsupportedEncryption),
        "Expected UnsupportedEncryption, got: {err:?}"
    );
}

#[test]
fn test_ole2_bytes_return_unsupported_encryption_pptx() {
    let ole2_magic: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut data = ole2_magic.to_vec();
    data.extend_from_slice(&[0x00; 100]);

    let err = convert_bytes(&data, Format::Pptx, &ConvertOptions::default()).unwrap_err();
    assert!(
        matches!(err, ConvertError::UnsupportedEncryption),
        "Expected UnsupportedEncryption, got: {err:?}"
    );
}
