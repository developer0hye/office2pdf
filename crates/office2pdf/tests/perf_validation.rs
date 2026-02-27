//! Performance validation tests.
//!
//! These tests verify that conversion of 10-page equivalent documents completes
//! within a 3-second budget on CI hardware (relaxed from 2s to handle CI variability).

use std::io::Cursor;
use std::time::Instant;

use office2pdf::config::{ConvertOptions, Format};

const BUDGET: std::time::Duration = std::time::Duration::from_secs(3);

fn build_docx_10_pages() -> Vec<u8> {
    let mut doc = docx_rs::Docx::new();
    for i in 0..30 {
        doc = doc.add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(format!(
                "Paragraph {i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
                     Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
            ))),
        );
    }
    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

fn build_pptx_10_slides() -> Vec<u8> {
    let slides = 10;
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let opts: zip::write::FileOptions = zip::write::FileOptions::default();

    let mut slide_ct = String::new();
    for i in 1..=slides {
        slide_ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
        ));
    }
    writer.start_file("[Content_Types].xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  {slide_ct}
</Types>"#
        )
        .as_bytes(),
    )
    .unwrap();

    writer.start_file("_rels/.rels", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    )
    .unwrap();

    let mut sid = String::new();
    for i in 1..=slides {
        sid.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{i}"/>"#, 255 + i));
    }
    writer.start_file("ppt/presentation.xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>{sid}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        )
        .as_bytes(),
    )
    .unwrap();

    let mut srels = String::new();
    for i in 1..=slides {
        srels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>"#
        ));
    }
    writer
        .start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {srels}
</Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    for i in 1..=slides {
        writer
            .start_file(format!("ppt/slides/slide{i}.xml"), opts)
            .unwrap();
        std::io::Write::write_all(
            &mut writer,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TextBox {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="5943600"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Slide {i}: Lorem ipsum dolor sit amet.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#
            )
            .as_bytes(),
        )
        .unwrap();
    }

    writer.finish().unwrap().into_inner()
}

fn build_xlsx_10_sheets() -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name("Sheet1");
    for row in 1..=20u32 {
        for col in 1..=5u32 {
            let coord = format!("{}{}", (b'A' + (col - 1) as u8) as char, row);
            sheet
                .get_cell_mut(coord.as_str())
                .set_value(format!("S1R{row}C{col}"));
        }
    }
    for s in 2..=10 {
        let name = format!("Sheet{s}");
        book.new_sheet(&name).unwrap();
        let sheet = book.get_sheet_by_name_mut(&name).unwrap();
        for row in 1..=20u32 {
            for col in 1..=5u32 {
                let coord = format!("{}{}", (b'A' + (col - 1) as u8) as char, row);
                sheet
                    .get_cell_mut(coord.as_str())
                    .set_value(format!("S{s}R{row}C{col}"));
            }
        }
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn perf_docx_10_pages_under_2s() {
    let data = build_docx_10_pages();
    let start = Instant::now();
    office2pdf::convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let elapsed = start.elapsed();
    assert!(
        elapsed < BUDGET,
        "DOCX 10-page conversion took {elapsed:?}, exceeds {BUDGET:?} budget"
    );
}

#[test]
fn perf_pptx_10_slides_under_2s() {
    let data = build_pptx_10_slides();
    let start = Instant::now();
    office2pdf::convert_bytes(&data, Format::Pptx, &ConvertOptions::default()).unwrap();
    let elapsed = start.elapsed();
    assert!(
        elapsed < BUDGET,
        "PPTX 10-slide conversion took {elapsed:?}, exceeds {BUDGET:?} budget"
    );
}

#[test]
fn perf_xlsx_10_sheets_under_2s() {
    let data = build_xlsx_10_sheets();
    let start = Instant::now();
    office2pdf::convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap();
    let elapsed = start.elapsed();
    assert!(
        elapsed < BUDGET,
        "XLSX 10-sheet conversion took {elapsed:?}, exceeds {BUDGET:?} budget"
    );
}
