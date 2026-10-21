//! A floating DrawingML object whose `wp:docPr` states `hidden` is neither
//! drawn nor wrapped around (issue #2299).
use super::*;
use std::io::Cursor;

const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="wps""#;

/// One page-anchored drawing. `hidden` is the raw `wp:docPr` attribute text,
/// or `None` to leave the attribute out.
fn anchored_drawing(
    id: u32,
    hidden: Option<&str>,
    (x, y): (i64, i64),
    (cx, cy): (i64, i64),
    wrap: &str,
    graphic: &str,
) -> String {
    let hidden: String = hidden
        .map(|value| format!(r#" hidden="{value}""#))
        .unwrap_or_default();
    format!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
            relativeHeight="{id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>
            <wp:extent cx="{cx}" cy="{cy}"/>
            {wrap}
            <wp:docPr id="{id}" name="Drawing {id}"{hidden}/>
            <a:graphic>{graphic}</a:graphic>
        </wp:anchor></w:drawing></w:r>"#
    )
}

/// A picture of the package's one image, outlined by a 1pt red `a:ln`.
fn picture_graphic(id: u32, (cx, cy): (i64, i64)) -> String {
    format!(
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
                <pic:nvPicPr><pic:cNvPr id="{id}" name="image1.bmp"/><pic:cNvPicPr/></pic:nvPicPr>
                <pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
                </pic:spPr>
            </pic:pic>
        </a:graphicData>"#
    )
}

fn text_box_graphic(text: &str) -> String {
    format!(
        r#"<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
                <wps:cNvSpPr txBox="1"/>
                <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                <wps:txbx><w:txbxContent>
                    <w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t>{text}</w:t></w:r></w:p>
                </w:txbxContent></wps:txbx>
                <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
            </wps:wsp>
        </a:graphicData>"#
    )
}

/// The anchored drawings every test story carries, in order: a hidden picture
/// and a hidden text box, each followed by a visible one of a different size
/// and wrap, so a drawing that skips its side-channel cursors hands the next
/// drawing the wrong metadata. `false` and `0` are the shown values.
fn story_drawings() -> String {
    [
        anchored_drawing(
            1,
            Some("1"),
            (12_700, 25_400),
            (1_270_000, 635_000),
            r#"<wp:wrapSquare wrapText="bothSides"/>"#,
            &picture_graphic(1, (1_270_000, 635_000)),
        ),
        anchored_drawing(
            2,
            None,
            (38_100, 50_800),
            (2_540_000, 1_270_000),
            "<wp:wrapTopAndBottom/>",
            &picture_graphic(2, (2_540_000, 1_270_000)),
        ),
        anchored_drawing(
            3,
            Some("true"),
            (63_500, 76_200),
            (914_400, 457_200),
            "<wp:wrapSquare wrapText=\"bothSides\"/>",
            &text_box_graphic("Hidden box"),
        ),
        anchored_drawing(
            4,
            Some("false"),
            (88_900, 101_600),
            (1_828_800, 914_400),
            "<wp:wrapNone/>",
            &text_box_graphic("Shown box"),
        ),
        anchored_drawing(
            5,
            Some("0"),
            (114_300, 127_000),
            (635_000, 635_000),
            r#"<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>"#,
            &picture_graphic(5, (635_000, 635_000)),
        ),
    ]
    .concat()
}

fn body_paragraph() -> &'static str {
    r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t>Body text</w:t></w:r></w:p>"#
}

/// A package whose body, and optionally whose default header, is
/// `body`/`header`, with relationship `rIdImage1` naming one BMP from either.
fn build_package(body: &str, header: Option<&str>) -> Vec<u8> {
    let image: Vec<u8> = super::image_tests::make_test_bmp();
    let header_reference: &str = if header.is_some() {
        r#"<w:headerReference w:type="default" r:id="rIdHeader1"/>"#
    } else {
        ""
    };
    let document: String = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NAMESPACES}><w:body>{body}<w:sectPr>{header_reference}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    );
    let image_relationship: &str = r#"<Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bmp"/>"#;
    let header_relationship: &str = if header.is_some() {
        r#"<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>"#
    } else {
        ""
    };
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"#
                .to_vec(),
        ),
        (
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#
                .to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_relationship}{header_relationship}</Relationships>"#
            )
            .into_bytes(),
        ),
        ("word/document.xml".to_string(), document.into_bytes()),
        ("word/media/image1.bmp".to_string(), image),
    ];
    if let Some(header) = header {
        parts.push((
            "word/header1.xml".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr {NAMESPACES}>{header}</w:hdr>"#
            )
            .into_bytes(),
        ));
        parts.push((
            "word/_rels/header1.xml.rels".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_relationship}</Relationships>"#
            )
            .into_bytes(),
        ));
    }

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, &content).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn first_flow_page(document: &Document) -> &FlowPage {
    match &document.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected FlowPage"),
    }
}

fn block_text(blocks: &[Block]) -> String {
    blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

#[test]
fn hidden_floating_drawings_in_the_body_are_dropped_and_the_rest_keep_their_wraps() {
    let data: Vec<u8> = build_package(
        &format!("<w:p>{}</w:p>{}", story_drawings(), body_paragraph()),
        None,
    );
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page: &FlowPage = first_flow_page(&document);

    let pictures: Vec<(f64, f64, WrapMode)> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::FloatingImage(image) => Some((
                image.image.width.unwrap(),
                image.image.height.unwrap(),
                image.wrap_mode,
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        pictures,
        vec![
            (200.0, 100.0, WrapMode::TopAndBottom),
            (50.0, 50.0, WrapMode::Tight)
        ],
        "only the shown pictures are drawn, each with its own wrap"
    );

    let text_boxes: Vec<(f64, f64, WrapMode, String)> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::FloatingTextBox(text_box) => Some((
                text_box.width,
                text_box.height,
                text_box.wrap_mode,
                block_text(&text_box.content),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        text_boxes,
        vec![(144.0, 72.0, WrapMode::None, "Shown box".to_string())],
        "only the shown text box is drawn, with its own size and wrap"
    );
    assert!(
        !format!("{document:?}").contains("Hidden box"),
        "a hidden text box's content reaches no flow"
    );
}

#[test]
fn hidden_floating_drawings_in_a_header_are_dropped() {
    let data: Vec<u8> = build_package(
        body_paragraph(),
        Some(&format!(
            "<w:p>{}</w:p><w:p><w:r><w:t>Header text</w:t></w:r></w:p>",
            story_drawings()
        )),
    );
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let header: &HeaderFooter = first_flow_page(&document)
        .header
        .as_ref()
        .expect("default header");

    let pictures: Vec<(Option<f64>, Option<f64>)> = header
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.elements)
        .filter_map(|element| match element {
            HFInline::Image(image) => Some((image.width, image.height)),
            _ => None,
        })
        .collect();
    assert_eq!(
        pictures,
        vec![(Some(200.0), Some(100.0)), (Some(50.0), Some(50.0))],
        "only the shown header pictures are drawn"
    );

    let outlines: Vec<(f64, f64)> = header
        .shapes
        .iter()
        .filter(|shape| matches!(shape.content, HeaderFooterShapeContent::Shape(_)))
        .map(|shape| (shape.width, shape.height))
        .collect();
    assert_eq!(
        outlines,
        vec![(200.0, 100.0), (50.0, 50.0)],
        "a hidden picture's outline is not drawn either"
    );

    let framed: Vec<(String, Option<f64>, Option<f64>)> = header
        .paragraphs
        .iter()
        .filter_map(|paragraph| {
            let frame = paragraph.frame.as_ref()?;
            let text: String = paragraph
                .elements
                .iter()
                .filter_map(|element| match element {
                    HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
                .collect();
            Some((text, frame.x, frame.y))
        })
        .collect();
    assert_eq!(
        framed,
        vec![("Shown box".to_string(), Some(7.0), Some(8.0))],
        "only the shown text box is framed, at its own offsets"
    );
}
