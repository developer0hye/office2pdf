use super::*;
use std::io::{Cursor, Read};

pub(super) fn make_test_bmp() -> Vec<u8> {
    let mut bmp = Vec::new();
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&58u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&54u32.to_le_bytes());
    bmp.extend_from_slice(&40u32.to_le_bytes());
    bmp.extend_from_slice(&1i32.to_le_bytes());
    bmp.extend_from_slice(&1i32.to_le_bytes());
    bmp.extend_from_slice(&1u16.to_le_bytes());
    bmp.extend_from_slice(&24u16.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&4u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);
    bmp
}

fn build_docx_with_image(width_px: u32, height_px: u32) -> Vec<u8> {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data).size(width_px * 9525, height_px * 9525);
    let paragraph = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic));
    let docx = docx_rs::Docx::new().add_paragraph(paragraph);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

fn build_docx_with_custom_image_document(document_xml: &str) -> Vec<u8> {
    build_docx_with_custom_media_document(document_xml, "media/image1.bmp", &make_test_bmp())
}

/// A DOCX whose `word/document.xml` is `document_xml` and whose only media
/// part, `word/<media_target>`, is referenced as relationship `rIdImage1`.
pub(super) fn build_docx_with_custom_media_document(
    document_xml: &str,
    media_target: &str,
    media_bytes: &[u8],
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    std::io::Write::write_all(
        &mut zip,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    std::io::Write::write_all(
        &mut zip,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{media_target}"/>
</Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

    zip.start_file(format!("word/{media_target}"), options)
        .unwrap();
    std::io::Write::write_all(&mut zip, media_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

fn find_images(doc: &Document) -> Vec<&ImageData> {
    let page = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .collect()
}

fn find_floating_images(doc: &Document) -> Vec<&FloatingImage> {
    let page = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|block| match block {
            Block::FloatingImage(image) => Some(image),
            _ => None,
        })
        .collect()
}

#[test]
fn test_docx_image_inline_basic() {
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(images.len(), 1, "Expected exactly one image block");
    assert!(!images[0].data.is_empty(), "Image data should not be empty");
}

#[test]
fn test_docx_image_format_is_png() {
    let data = build_docx_with_image(50, 50);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(images[0].format, ImageFormat::Png);
}

#[test]
fn test_docx_vml_shape_image_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="VMLImage1" style="width:72pt;height:36pt">
                        <v:imagedata r:id="rIdImage1"/>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_custom_image_document(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    assert_eq!(images.len(), 1, "Expected one VML image");
    assert_eq!(images[0].format, ImageFormat::Png);
    assert_eq!(images[0].width, Some(72.0));
    assert_eq!(images[0].height, Some(36.0));
}

#[test]
fn test_docx_image_dimensions() {
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    let width = images[0].width.expect("Expected width");
    let height = images[0].height.expect("Expected height");

    assert!((width - 75.0).abs() < 0.1);
    assert!((height - 60.0).abs() < 0.1);
}

#[test]
fn test_docx_image_with_text_paragraphs() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data);
    let docx = docx_rs::Docx::new()
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before image")),
        )
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)))
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After image")),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    assert!(
        page.content
            .iter()
            .any(|block| matches!(block, Block::Image(_)))
    );
    assert!(page.content.iter().any(|block| match block {
        Block::Paragraph(paragraph) => paragraph.runs.iter().any(|run| run.text.contains("Before")),
        _ => false,
    }));
    assert!(page.content.iter().any(|block| match block {
        Block::Paragraph(paragraph) => paragraph.runs.iter().any(|run| run.text.contains("After")),
        _ => false,
    }));
}

#[test]
fn test_docx_multiple_images() {
    let bmp_data = make_test_bmp();
    let pic1 = docx_rs::Pic::new(&bmp_data).size(100 * 9525, 100 * 9525);
    let pic2 = docx_rs::Pic::new(&bmp_data).size(200 * 9525, 150 * 9525);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic1)))
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic2)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(find_images(&doc).len() >= 2);
}

#[test]
fn test_docx_image_data_contains_png_header() {
    let data = build_docx_with_image(50, 50);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    let img_data = &images[0].data;
    assert!(img_data.len() >= 8 && img_data[0..4] == [0x89, 0x50, 0x4E, 0x47]);
}

#[test]
fn test_docx_no_images_produces_no_image_blocks() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    let image_count = page
        .content
        .iter()
        .filter(|block| matches!(block, Block::Image(_)))
        .count();
    assert_eq!(image_count, 0);
}

#[test]
fn test_docx_image_with_custom_emu_size() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data).size(2_540_000, 1_270_000);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let images = find_images(&doc);
    let width = images[0].width.expect("Expected width");
    let height = images[0].height.expect("Expected height");
    assert!((width - 200.0).abs() < 0.1);
    assert!((height - 100.0).abs() < 0.1);
}

#[test]
fn test_docx_floating_image_square_wrap() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(2_540_000, 1_270_000)
        .floating()
        .offset_x(914_400)
        .offset_y(457_200);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1);
    assert_eq!(floating[0].wrap_mode, WrapMode::Square);
    assert!(!floating[0].image.data.is_empty());
    assert!((floating[0].image.width.expect("Expected width") - 200.0).abs() < 0.5);
    assert!((floating[0].image.height.expect("Expected height") - 100.0).abs() < 0.5);
}

#[test]
fn test_docx_picture_with_unconvertible_emf_is_skipped_not_fatal() {
    // docx-rs lists an EMF picture with an empty PNG preview and leaves the
    // conversion to us. When the metafile converter cannot read it either,
    // the picture must be dropped: handing Typst the empty preview as a PNG
    // failed the whole document ("failed to decode image").
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r><w:t xml:space="preserve">Quarterly chart: </w:t></w:r>
            <w:r><w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                    <wp:extent cx="914400" cy="457200"/>
                    <wp:docPr id="1" name="Chart 1"/>
                    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic>
                            <pic:nvPicPr><pic:cNvPr id="1" name="image1.emf"/><pic:cNvPicPr/></pic:nvPicPr>
                            <pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                            <pic:spPr>
                                <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData></a:graphic>
                </wp:inline>
            </w:drawing></w:r>
            <w:r><w:t xml:space="preserve"> see the appendix.</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;
    // Truncated inside its header record: no metafile reader can draw it.
    let truncated_emf: &[u8] = &[1, 0, 0, 0, 108, 0, 0, 0, 0, 0];
    let data =
        build_docx_with_custom_media_document(document_xml, "media/image1.emf", truncated_emf);

    let result = crate::convert_bytes(&data, crate::Format::Docx, &ConvertOptions::default())
        .expect("an unreadable picture must not fail the document");
    let pdf_text: String =
        pdf_extract::extract_text_from_mem(&result.pdf).expect("PDF text should extract");
    assert!(
        pdf_text.contains("Quarterly chart:") && pdf_text.contains("see the appendix."),
        "the paragraph around the picture still renders: {pdf_text:?}"
    );
}

#[test]
fn docx_floating_picture_rotation_reaches_ir_and_typst() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p><w:r><w:drawing>
            <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                       relativeHeight="1" behindDoc="1" locked="0"
                       layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                <wp:extent cx="4366365" cy="4366365"/>
                <wp:wrapNone/>
                <wp:docPr id="1" name="Picture 1"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                        <pic:nvPicPr><pic:cNvPr id="1" name="image1.bmp"/><pic:cNvPicPr/></pic:nvPicPr>
                        <pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                        <pic:spPr>
                            <a:xfrm rot="20933656"><a:off x="0" y="0"/><a:ext cx="4366365" cy="4366365"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        </pic:spPr>
                    </pic:pic>
                </a:graphicData></a:graphic>
            </wp:anchor>
        </w:drawing></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_custom_image_document(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let floating = find_floating_images(&doc);

    assert_eq!(floating.len(), 1);
    assert_eq!(floating[0].image.rotation_deg, Some(349.0));

    let source = crate::render::typst_gen::generate_typst(&doc)
        .unwrap()
        .source;
    assert!(
        source.contains("#move(dx:")
            && source.contains("#rotate(349deg, origin: top + left)[#image("),
        "floating picture rotation was not emitted: {source}"
    );
}

#[test]
fn test_docx_floating_image_top_and_bottom_wrap() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000)
        .floating()
        .overlapping();
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = patch_docx_wrap_type(&cursor.into_inner(), "wp:wrapNone", "wp:wrapTopAndBottom");

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1);
    assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
}

#[test]
fn test_docx_floating_image_behind_wrap() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000)
        .floating()
        .overlapping();
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = patch_docx_behind_doc(&cursor.into_inner());

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1);
    assert_eq!(floating[0].wrap_mode, WrapMode::Behind);
}

#[test]
fn test_docx_floating_image_position_offset() {
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data)
        .size(1_270_000, 1_270_000)
        .floating()
        .offset_x(914_400)
        .offset_y(457_200);
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let floating = find_floating_images(&doc);
    assert_eq!(floating.len(), 1);
    assert!((floating[0].offset_x - 72.0).abs() < 0.5);
    assert!((floating[0].offset_y - 36.0).abs() < 0.5);
}

#[test]
fn test_docx_inline_image_not_floating() {
    let data = build_docx_with_image(100, 80);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(find_floating_images(&doc).len(), 0);
    assert_eq!(find_images(&doc).len(), 1);
}

fn patch_docx_wrap_type(data: &[u8], old_wrap: &str, new_wrap: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::FileOptions::default();
        new_zip.start_file(&name, options).unwrap();

        let mut contents = Vec::new();
        file.read_to_end(&mut contents).unwrap();

        if name == "word/document.xml" {
            let xml = String::from_utf8(contents).unwrap();
            let xml = xml
                .replace(&format!("<{old_wrap} />"), &format!("<{new_wrap} />"))
                .replace(&format!("<{old_wrap}/>"), &format!("<{new_wrap}/>"));
            std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
        } else {
            std::io::Write::write_all(&mut new_zip, &contents).unwrap();
        }
    }

    new_zip.finish().unwrap().into_inner()
}

fn patch_docx_behind_doc(data: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::FileOptions::default();
        new_zip.start_file(&name, options).unwrap();

        let mut contents = Vec::new();
        file.read_to_end(&mut contents).unwrap();

        if name == "word/document.xml" {
            let xml = String::from_utf8(contents).unwrap();
            let xml = xml.replace("behindDoc=\"0\"", "behindDoc=\"1\"");
            std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
        } else {
            std::io::Write::write_all(&mut new_zip, &contents).unwrap();
        }
    }

    new_zip.finish().unwrap().into_inner()
}

#[test]
fn test_inline_image_inherits_paragraph_center_alignment() {
    // Word centers an inline image via the paragraph's <w:jc w:val="center"/>.
    let bmp_data = make_test_bmp();
    let pic = docx_rs::Pic::new(&bmp_data).size(100 * 9525, 50 * 9525);
    let paragraph = docx_rs::Paragraph::new()
        .add_run(docx_rs::Run::new().add_image(pic))
        .align(docx_rs::AlignmentType::Center);
    let docx = docx_rs::Docx::new().add_paragraph(paragraph);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();

    let parser = DocxParser;
    let (doc, _warnings) = parser
        .parse(&cursor.into_inner(), &ConvertOptions::default())
        .unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    let image = page
        .content
        .iter()
        .find_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .expect("expected an inline image block");
    assert_eq!(image.alignment, Some(Alignment::Center));
}
