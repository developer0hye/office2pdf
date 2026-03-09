use crate::ir::*;

pub(super) fn make_simple_document(text: &str) -> Document {
    Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    }
}

pub(super) fn build_docx_with_title(title: &str) -> Vec<u8> {
    use std::io::{Cursor, Write};

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    Write::write_all(&mut zip, br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    Write::write_all(&mut zip, br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
    )
    .unwrap();

    zip.start_file("docProps/core.xml", options).unwrap();
    let core_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>{title}</dc:title>
</cp:coreProperties>"#
    );
    Write::write_all(&mut zip, core_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

pub(super) fn make_test_png() -> Vec<u8> {
    fn png_crc32(chunk_type: &[u8], data: &[u8]) -> u32 {
        let mut crc: u32 = 0xFFFF_FFFF;
        for &byte in chunk_type.iter().chain(data.iter()) {
            crc ^= byte as u32;
            for _ in 0..8 {
                if crc & 1 != 0 {
                    crc = (crc >> 1) ^ 0xEDB8_8320;
                } else {
                    crc >>= 1;
                }
            }
        }
        crc ^ 0xFFFF_FFFF
    }

    let mut png = Vec::new();
    png.extend_from_slice(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
    let ihdr_data: [u8; 13] = [
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00,
    ];
    let ihdr_type = b"IHDR";
    png.extend_from_slice(&(ihdr_data.len() as u32).to_be_bytes());
    png.extend_from_slice(ihdr_type);
    png.extend_from_slice(&ihdr_data);
    png.extend_from_slice(&png_crc32(ihdr_type, &ihdr_data).to_be_bytes());
    let idat_data: [u8; 15] = [
        0x78, 0x01, 0x01, 0x04, 0x00, 0xFB, 0xFF, 0x00, 0xFF, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
    ];
    let idat_type = b"IDAT";
    png.extend_from_slice(&(idat_data.len() as u32).to_be_bytes());
    png.extend_from_slice(idat_type);
    png.extend_from_slice(&idat_data);
    png.extend_from_slice(&png_crc32(idat_type, &idat_data).to_be_bytes());
    let iend_type = b"IEND";
    png.extend_from_slice(&0u32.to_be_bytes());
    png.extend_from_slice(iend_type);
    png.extend_from_slice(&png_crc32(iend_type, &[]).to_be_bytes());
    png
}

pub(super) fn build_test_docx() -> Vec<u8> {
    use std::io::Cursor;

    let docx = docx_rs::Docx::new()
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello from DOCX")),
        )
        .add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Second paragraph").bold()),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

pub(super) fn build_test_xlsx() -> Vec<u8> {
    use std::io::Cursor;

    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Name");
        sheet.get_cell_mut("B1").set_value("Value");
        sheet.get_cell_mut("A2").set_value("Item 1");
        sheet.get_cell_mut("B2").set_value("100");
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

pub(super) fn build_test_pptx() -> Vec<u8> {
    use std::io::{Cursor, Write};

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
    ).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    ).unwrap();

    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    ).unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
    ).unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello from PPTX</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    ).unwrap();

    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    ).unwrap();

    zip.finish().unwrap().into_inner()
}

pub(super) fn make_test_docx_bytes() -> Vec<u8> {
    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello metrics")),
    );
    let mut buf = std::io::Cursor::new(Vec::new());
    docx.build().pack(&mut buf).unwrap();
    buf.into_inner()
}
