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

pub(super) fn build_docx_with_custom_image_document(document_xml: &str) -> Vec<u8> {
    build_docx_with_custom_media_document(document_xml, "media/image1.bmp", &make_test_bmp())
}

/// A DOCX whose `word/document.xml` is `document_xml` and whose only media
/// part, `word/<media_target>`, is referenced as relationship `rIdImage1`.
fn build_docx_with_custom_media_document(
    document_xml: &str,
    media_target: &str,
    media_bytes: &[u8],
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();

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

/// A VML picture in the line takes its paragraph's `w:ind` as a DrawingML
/// inline does, but one whose style positions it is placed by its own
/// offsets, so the paragraph's indent must not move it (issue #2201).
#[test]
fn only_vml_pictures_in_the_line_take_the_paragraph_indent() {
    let cases: [(&str, Option<ImageParagraphIndent>); 4] = [
        (
            "width:72pt;height:36pt",
            Some(ImageParagraphIndent {
                left: 36.0,
                right: 0.0,
                first_line: 9.0,
            }),
        ),
        (
            "position:absolute;margin-left:0;margin-top:0;width:72pt;height:36pt",
            Some(ImageParagraphIndent::default()),
        ),
        (
            "position:absolute;left:0;margin-left:35.25pt;width:72pt;height:36pt;mso-position-horizontal-relative:page",
            Some(ImageParagraphIndent::default()),
        ),
        (
            "margin-left:12pt;width:72pt;height:36pt",
            Some(ImageParagraphIndent::default()),
        ),
    ];
    for (shape_style, expected) in cases {
        let document_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:pPr><w:ind w:left="720" w:firstLine="180"/></w:pPr>
            <w:r>
                <w:pict>
                    <v:shape id="VMLImage1" style="{shape_style}">
                        <v:imagedata r:id="rIdImage1"/>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#
        );
        let data = build_docx_with_custom_image_document(&document_xml);
        let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let images = find_images(&doc);
        assert_eq!(images.len(), 1, "{shape_style}: one VML picture");
        assert_eq!(images[0].paragraph_indent, expected, "{shape_style}");
    }
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
        let name = file.name().expect("entry name should decode").into_owned();
        let options = zip::write::SimpleFileOptions::default();
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
        let name = file.name().expect("entry name should decode").into_owned();
        let options = zip::write::SimpleFileOptions::default();
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

/// A picture paragraph's `w:ind` reaches its pictures resolved the way a text
/// paragraph's does: each attribute it states directly over its style's, and a
/// hanging indent as a negative first line (issue #2201).
#[test]
fn inline_pictures_carry_their_paragraphs_resolved_indent() {
    use docx_rs::SpecialIndentType::{FirstLine, Hanging};

    let bmp_data = make_test_bmp();
    let picture = || docx_rs::Run::new().add_image(docx_rs::Pic::new(&bmp_data).size(9525, 9525));
    let figure_style = docx_rs::Style::new("Figure", docx_rs::StyleType::Paragraph)
        .name("Figure")
        .indent(Some(1440), Some(Hanging(360)), Some(240), None);
    let indent = |left: f64, right: f64, first_line: f64| {
        Some(ImageParagraphIndent {
            left,
            right,
            first_line,
        })
    };
    let cases: [(&str, docx_rs::Paragraph, Option<ImageParagraphIndent>); 5] = [
        ("no indent", docx_rs::Paragraph::new(), None),
        (
            "direct first line",
            docx_rs::Paragraph::new().indent(None, Some(FirstLine(709)), None, None),
            indent(0.0, 0.0, 35.45),
        ),
        (
            "direct left, hanging and right",
            docx_rs::Paragraph::new().indent(Some(720), Some(Hanging(360)), Some(360), None),
            indent(36.0, 18.0, -18.0),
        ),
        (
            "style indent",
            docx_rs::Paragraph::new().style("Figure"),
            indent(72.0, 12.0, -18.0),
        ),
        (
            // The builder always writes `w:left` and `w:right`, so all
            // three attributes are direct here.
            "direct indent over the style's",
            docx_rs::Paragraph::new().style("Figure").indent(
                Some(360),
                Some(FirstLine(720)),
                None,
                None,
            ),
            indent(18.0, 0.0, 36.0),
        ),
    ];
    for (case, paragraph, expected) in cases {
        for picture_count in [1, 2] {
            let paragraph = (0..picture_count).fold(paragraph.clone(), |paragraph, _| {
                paragraph.add_run(picture())
            });
            let docx = docx_rs::Docx::new()
                .add_style(figure_style.clone())
                .add_paragraph(paragraph);
            let mut cursor = Cursor::new(Vec::new());
            docx.build().pack(&mut cursor).unwrap();
            let (doc, _warnings) = DocxParser
                .parse(&cursor.into_inner(), &ConvertOptions::default())
                .unwrap();
            let Page::Flow(page) = &doc.pages[0] else {
                panic!("Expected FlowPage");
            };
            let indents: Vec<Option<ImageParagraphIndent>> = page
                .content
                .iter()
                .flat_map(|block| match block {
                    Block::Image(image) => vec![image.paragraph_indent],
                    Block::InlineImages(images) => {
                        images.iter().map(|image| image.paragraph_indent).collect()
                    }
                    _ => Vec::new(),
                })
                .collect();
            assert_eq!(
                indents,
                vec![expected; picture_count],
                "{case}, {picture_count} picture(s)"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_source_crops_preserve_signed_offsets_in_body_header_and_footer() {
    let fixture: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
    let mut archive = zip::ZipArchive::new(Cursor::new(fixture)).unwrap();
    let mut png: Vec<u8> = Vec::new();
    archive
        .by_name("word/media/bands.png")
        .unwrap()
        .read_to_end(&mut png)
        .unwrap();
    let picture_paragraph = || {
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new().add_image(docx_rs::Pic::new(&png).size(1_714_500, 952_500)),
        )
    };
    let docx = docx_rs::Docx::new()
        .header(docx_rs::Header::new().add_paragraph(picture_paragraph()))
        .footer(docx_rs::Footer::new().add_paragraph(picture_paragraph()))
        .add_paragraph(picture_paragraph());
    let mut packed = Cursor::new(Vec::new());
    docx.build().pack(&mut packed).unwrap();
    let base: Vec<u8> = packed.into_inner();
    let archive = zip::ZipArchive::new(Cursor::new(&base)).unwrap();
    let parts: Vec<String> = archive
        .file_names()
        .map(|name| name.expect("entry name should decode").into_owned())
        .filter(|name| {
            name == "word/document.xml"
                || ((name.starts_with("word/header") || name.starts_with("word/footer"))
                    && name.ends_with(".xml"))
        })
        .collect();
    assert_eq!(parts.len(), 3);
    for (rectangle, expected) in [
        (
            r#"<draw:srcRect xmlns:draw="http://schemas.openxmlformats.org/drawingml/2006/main" l="12.5%" r="12500" b="-2.25%"/>"#,
            Some(ImageCrop {
                left: 0.125,
                top: 0.0,
                right: 0.125,
                bottom: -0.0225,
            }),
        ),
        (
            r#"<a:srcRect t="4604" r="1295" b="6879"/>"#,
            Some(ImageCrop {
                left: 0.0,
                top: 0.04604,
                right: 0.01295,
                bottom: 0.06879,
            }),
        ),
        (
            r#"<a:extLst><a:ext uri="urn:unrelated"><a:srcRect l="75000"/></a:ext></a:extLst><a:srcRect l="12500" foreign:l="75000" xmlns:foreign="urn:foreign"/>"#,
            Some(ImageCrop {
                left: 0.125,
                top: 0.0,
                right: 0.0,
                bottom: 0.0,
            }),
        ),
        (r#"<a:srcRect l="0" t="0" r="0" b="0"/>"#, None),
        (
            r#"<foreign:srcRect xmlns:foreign="urn:foreign" l="75000"/>"#,
            None,
        ),
    ] {
        let mut data: Vec<u8> = base.clone();
        for part in &parts {
            data = rewrite_docx_part(&data, part, |xml| {
                assert!(xml.contains("<a:srcRect />") || xml.contains("<a:srcRect/>"));
                xml.replace("<a:srcRect />", rectangle)
                    .replace("<a:srcRect/>", rectangle)
            });
        }
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let images = find_images(&document);
        assert_eq!(images.len(), 1);
        assert_eq!(
            images[0].crop, expected,
            "body source rectangle {rectangle}"
        );
        let Page::Flow(page) = &document.pages[0] else {
            panic!("flow page")
        };
        for story in [page.header.as_ref().unwrap(), page.footer.as_ref().unwrap()] {
            let image = story
                .paragraphs
                .iter()
                .flat_map(|p| &p.elements)
                .find_map(|element| match element {
                    HFInline::Image(image) => Some(image),
                    _ => None,
                })
                .expect("story picture");
            assert_eq!(
                image.crop, expected,
                "header/footer source rectangle {rectangle}"
            );
            assert_eq!((image.width, image.height), (Some(135.0), Some(75.0)));
        }
    }
}

#[test]
fn docx_source_crops_stay_with_shared_pictures_across_story_tables() {
    let fixture: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
    let mut archive = zip::ZipArchive::new(Cursor::new(fixture)).unwrap();
    let mut png: Vec<u8> = Vec::new();
    archive
        .by_name("word/media/bands.png")
        .unwrap()
        .read_to_end(&mut png)
        .unwrap();
    let picture = docx_rs::Pic::new(&png).size(1_714_500, 952_500);
    let paragraph =
        || docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(picture.clone()));
    let table = || {
        docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(paragraph()),
        ])])
    };
    let docx = docx_rs::Docx::new()
        .header(
            docx_rs::Header::new()
                .add_paragraph(paragraph())
                .add_table(table())
                .add_paragraph(paragraph()),
        )
        .footer(
            docx_rs::Footer::new()
                .add_paragraph(paragraph())
                .add_table(table())
                .add_paragraph(paragraph()),
        )
        .add_paragraph(docx_rs::Paragraph::new());
    let mut packed = Cursor::new(Vec::new());
    docx.build().pack(&mut packed).unwrap();
    let mut data: Vec<u8> = packed.into_inner();
    let archive = zip::ZipArchive::new(Cursor::new(&data)).unwrap();
    let parts: Vec<String> = archive
        .file_names()
        .map(|name| name.expect("entry name should decode").into_owned())
        .filter(|name| {
            (name.starts_with("word/header") || name.starts_with("word/footer"))
                && name.ends_with(".xml")
        })
        .collect();
    for part in parts {
        data = rewrite_docx_part(&data, &part, |xml| {
            let mut result = xml.replace("<a:srcRect />", "<a:srcRect/>");
            for left in [12500, 25000, 37500] {
                result = result.replacen("<a:srcRect/>", &format!("<a:srcRect l=\"{left}\"/>"), 1);
            }
            assert!(!result.contains("<a:srcRect/>"));
            result
        });
    }
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("flow page")
    };
    for story in [page.header.as_ref().unwrap(), page.footer.as_ref().unwrap()] {
        let paragraph_crops: Vec<f64> = story
            .paragraphs
            .iter()
            .flat_map(|paragraph| &paragraph.elements)
            .filter_map(|element| match element {
                HFInline::Image(image) => Some(image.crop.expect("paragraph crop").left),
                _ => None,
            })
            .collect();
        assert_eq!(paragraph_crops, vec![0.125, 0.375]);
        let table = story
            .shapes
            .iter()
            .find_map(|shape| match &shape.content {
                crate::ir::HeaderFooterShapeContent::Table(table) => Some(table),
                _ => None,
            })
            .expect("story table");
        let table_crop = table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.content)
            .find_map(|block| match block {
                Block::Image(image) => image.crop,
                Block::InlineImages(images) => images.first().and_then(|image| image.crop),
                _ => None,
            })
            .expect("table crop");
        assert_eq!(table_crop.left, 0.25);
    }
}

/// A paragraph's runs in order: a text run as its text, and a picture riding
/// the line as `[width x height]`, read from both the box and its picture.
fn line_items(paragraph: &Paragraph) -> Vec<String> {
    paragraph
        .runs
        .iter()
        .map(|run| match run.inline_box.as_deref() {
            Some(inline_box) => match inline_box.content.as_slice() {
                [Block::Image(image)]
                    if image.width == Some(inline_box.width)
                        && image.height == Some(inline_box.height) =>
                {
                    format!("[{}x{}]", inline_box.width, inline_box.height)
                }
                _ => "[other box]".to_string(),
            },
            None => run.text.clone(),
        })
        .collect()
}

/// Every picture block and every paragraph in `blocks`, descending into
/// table cells.
fn picture_blocks_and_paragraphs(blocks: &[Block]) -> (usize, Vec<&Paragraph>) {
    let mut picture_blocks: usize = 0;
    let mut paragraphs: Vec<&Paragraph> = Vec::new();
    for block in blocks {
        match block {
            Block::Image(_) | Block::InlineImages(_) => picture_blocks += 1,
            Block::Paragraph(paragraph) => paragraphs.push(paragraph),
            Block::Table(table) => {
                for cell in table.rows.iter().flat_map(|row| &row.cells) {
                    let (cell_pictures, cell_paragraphs) =
                        picture_blocks_and_paragraphs(&cell.content);
                    picture_blocks += cell_pictures;
                    paragraphs.extend(cell_paragraphs);
                }
            }
            _ => {}
        }
    }
    (picture_blocks, paragraphs)
}

/// An inline picture that shares its paragraph with text is a glyph-like
/// object on that text's line, so it rides the paragraph's runs in source
/// order as a box of its `wp:extent` (issue #2227). Lifting it out as an image
/// block drew it above the paragraph and printed text that comes before it
/// after it. Whitespace is text too: Word keeps leading spaces before the
/// picture. A paragraph holding only pictures keeps the picture-paragraph path
/// (issues #499, #2201).
#[test]
fn an_inline_picture_sharing_its_paragraph_with_text_rides_the_runs() {
    let bmp_data = make_test_bmp();
    // 72x36pt, the `wp:extent` of docxide-pdf's `cell-inline-*` cases.
    let pic = || docx_rs::Pic::new(&bmp_data).size(914_400, 457_200);
    let picture = || docx_rs::Run::new().add_image(pic());
    let text = |text: &str| docx_rs::Run::new().add_text(text);
    let cases: [(&str, docx_rs::Paragraph, &[&str]); 5] = [
        (
            "text on either side",
            docx_rs::Paragraph::new()
                .add_run(text("Left "))
                .add_run(picture())
                .add_run(text(" End")),
            &["Left ", "[72x36]", " End"],
        ),
        (
            "leading spaces only",
            docx_rs::Paragraph::new()
                .add_run(text("   "))
                .add_run(picture()),
            &["   ", "[72x36]"],
        ),
        (
            "picture first",
            docx_rs::Paragraph::new()
                .add_run(picture())
                .add_run(text("Caption")),
            &["[72x36]", "Caption"],
        ),
        (
            "text and picture in one run",
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Left ")
                    .add_image(pic())
                    .add_text(" End"),
            ),
            &["Left ", "[72x36]", " End"],
        ),
        (
            "tab before the picture",
            docx_rs::Paragraph::new()
                .add_run(text("Left ").add_tab())
                .add_run(picture())
                .add_run(text(" End")),
            &["Left \t", "[72x36]", " End"],
        ),
    ];
    for (case, paragraph, expected) in cases {
        let in_body = docx_rs::Docx::new().add_paragraph(paragraph.clone());
        let in_cell =
            docx_rs::Docx::new().add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(
                vec![docx_rs::TableCell::new().add_paragraph(paragraph)],
            )]));
        for (place, docx) in [("body", in_body), ("table cell", in_cell)] {
            let mut cursor = Cursor::new(Vec::new());
            docx.build().pack(&mut cursor).unwrap();
            let (doc, _warnings) = DocxParser
                .parse(&cursor.into_inner(), &ConvertOptions::default())
                .unwrap();
            let Page::Flow(page) = &doc.pages[0] else {
                panic!("Expected FlowPage");
            };
            let (picture_blocks, paragraphs) = picture_blocks_and_paragraphs(&page.content);
            let lines: Vec<Vec<String>> = paragraphs.into_iter().map(line_items).collect();
            assert_eq!(picture_blocks, 0, "{case} in {place}: {lines:?}");
            assert_eq!(lines, vec![expected.to_vec()], "{case} in {place}");
        }
    }

    // The controls: a picture alone in its paragraph is still a picture
    // block, and so is one a page break separates from the paragraph's text,
    // while a picture beside text past the break rides that text's line.
    let alone = docx_rs::Paragraph::new().add_run(picture());
    let split_by_break = docx_rs::Paragraph::new()
        .add_run(picture())
        .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page))
        .add_run(text("Next "))
        .add_run(picture());
    for (case, paragraph, expected_picture_blocks, expected_lines) in [
        ("picture alone", alone, 1, Vec::new()),
        (
            "picture split from the text by a page break",
            split_by_break,
            1,
            vec![vec!["Next ".to_string(), "[72x36]".to_string()]],
        ),
    ] {
        let docx = docx_rs::Docx::new().add_paragraph(paragraph);
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let (doc, _warnings) = DocxParser
            .parse(&cursor.into_inner(), &ConvertOptions::default())
            .unwrap();
        let Page::Flow(page) = &doc.pages[0] else {
            panic!("Expected FlowPage");
        };
        let (picture_blocks, paragraphs) = picture_blocks_and_paragraphs(&page.content);
        let lines: Vec<Vec<String>> = paragraphs
            .into_iter()
            .map(line_items)
            .filter(|line| !line.is_empty())
            .collect();
        assert_eq!(picture_blocks, expected_picture_blocks, "{case}: {lines:?}");
        assert_eq!(lines, expected_lines, "{case}");
    }
}
