use std::io::Cursor;

use criterion::{Criterion, criterion_group, criterion_main};
use office2pdf::config::{ConvertOptions, Format};

/// Build a DOCX with the given number of paragraphs.
fn build_docx(paragraphs: usize) -> Vec<u8> {
    let mut doc = docx_rs::Docx::new();
    for i in 0..paragraphs {
        doc = doc.add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text(format!("Paragraph {i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."))),
        );
    }
    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

/// Build a PPTX with the given number of slides, each containing a text box.
fn build_pptx(slides: usize) -> Vec<u8> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut writer = zip::ZipWriter::new(cursor);
    let opts: zip::write::FileOptions = zip::write::FileOptions::default();

    // [Content_Types].xml
    let mut slide_content_types = String::new();
    for i in 1..=slides {
        slide_content_types.push_str(&format!(
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
  {slide_content_types}
</Types>"#
        )
        .as_bytes(),
    )
    .unwrap();

    // _rels/.rels
    writer.start_file("_rels/.rels", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    )
    .unwrap();

    // ppt/presentation.xml
    let mut slide_id_list = String::new();
    for i in 1..=slides {
        slide_id_list.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{i}"/>"#, 255 + i));
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
  <p:sldIdLst>{slide_id_list}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        )
        .as_bytes(),
    )
    .unwrap();

    // ppt/_rels/presentation.xml.rels
    let mut slide_rels = String::new();
    for i in 1..=slides {
        slide_rels.push_str(&format!(
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
  {slide_rels}
</Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    // Slides
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
          <a:p><a:r><a:t>Slide {i}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.</a:t></a:r></a:p>
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

/// Build an XLSX with the given number of rows.
fn build_xlsx(rows: usize) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name("Data");
    for row in 1..=rows {
        for col in 1..=5u32 {
            let coord = format!("{}{}", (b'A' + (col - 1) as u8) as char, row);
            sheet
                .get_cell_mut(coord.as_str())
                .set_value(format!("R{row}C{col}"));
        }
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

fn bench_docx_conversion(c: &mut Criterion) {
    let data_10 = build_docx(30); // ~10 pages worth
    let data_100 = build_docx(300); // ~100 pages worth

    let mut group = c.benchmark_group("docx");
    group.sample_size(10);

    group.bench_function("10_pages", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_10, Format::Docx, &ConvertOptions::default()).unwrap()
        })
    });

    group.bench_function("100_pages", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_100, Format::Docx, &ConvertOptions::default()).unwrap()
        })
    });

    group.finish();
}

fn bench_pptx_conversion(c: &mut Criterion) {
    let data_10 = build_pptx(10);
    let data_100 = build_pptx(100);

    let mut group = c.benchmark_group("pptx");
    group.sample_size(10);

    group.bench_function("10_slides", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_10, Format::Pptx, &ConvertOptions::default()).unwrap()
        })
    });

    group.bench_function("100_slides", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_100, Format::Pptx, &ConvertOptions::default()).unwrap()
        })
    });

    group.finish();
}

fn bench_xlsx_conversion(c: &mut Criterion) {
    let data_small = build_xlsx(50);
    let data_large = build_xlsx(500);

    let mut group = c.benchmark_group("xlsx");
    group.sample_size(10);

    group.bench_function("50_rows", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_small, Format::Xlsx, &ConvertOptions::default())
                .unwrap()
        })
    });

    group.bench_function("500_rows", |b| {
        b.iter(|| {
            office2pdf::convert_bytes(&data_large, Format::Xlsx, &ConvertOptions::default())
                .unwrap()
        })
    });

    group.finish();
}

criterion_group!(
    benches,
    bench_docx_conversion,
    bench_pptx_conversion,
    bench_xlsx_conversion
);
criterion_main!(benches);
