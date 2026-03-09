use super::*;
use crate::ir::ImageCrop;
use std::io::{Cursor, Write};
use zip::write::FileOptions;

/// Create a minimal valid BMP (1×1 pixel, red) for test images.
fn make_test_bmp() -> Vec<u8> {
    let mut bmp = Vec::new();
    // BMP header (14 bytes)
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&70u32.to_le_bytes()); // file size
    bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
    bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset
    // DIB header (40 bytes)
    bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
    bmp.extend_from_slice(&1i32.to_le_bytes()); // width
    bmp.extend_from_slice(&1i32.to_le_bytes()); // height
    bmp.extend_from_slice(&1u16.to_le_bytes()); // planes
    bmp.extend_from_slice(&24u16.to_le_bytes()); // bpp
    bmp.extend_from_slice(&0u32.to_le_bytes()); // compression
    bmp.extend_from_slice(&16u32.to_le_bytes()); // image size
    bmp.extend_from_slice(&2835u32.to_le_bytes()); // h resolution
    bmp.extend_from_slice(&2835u32.to_le_bytes()); // v resolution
    bmp.extend_from_slice(&0u32.to_le_bytes()); // colors
    bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors
    // Pixel data: 1 pixel (BGR) + 1 byte padding to align to 4 bytes
    bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);
    bmp
}

/// Create a minimal valid SVG image for test images.
fn make_test_svg() -> Vec<u8> {
    br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1" viewBox="0 0 1 1"><rect width="1" height="1" fill="#ff0000"/></svg>"##.to_vec()
}

/// Create a picture XML element referencing an image via relationship ID.
fn make_pic_xml(x: i64, y: i64, cx: i64, cy: i64, r_embed: &str) -> String {
    make_custom_pic_xml(
        x,
        y,
        cx,
        cy,
        &format!(r#"<a:blip r:embed="{r_embed}"/><a:stretch><a:fillRect/></a:stretch>"#),
    )
}

/// Create a picture XML element with custom `<p:blipFill>` contents.
fn make_custom_pic_xml(x: i64, y: i64, cx: i64, cy: i64, blip_fill_xml: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip_fill_xml}</p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr></p:pic>"#
    )
}

/// Slide image for the test PPTX builder.
struct TestSlideImage {
    rid: String,
    path: String,
    data: Vec<u8>,
    relationship_type: Option<String>,
}

/// Build a PPTX file with slides that have image relationships.
fn build_test_pptx_with_images(
    slide_cx_emu: i64,
    slide_cy_emu: i64,
    slides: &[(String, Vec<TestSlideImage>)],
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    // [Content_Types].xml
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    ct.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    ct.push_str(r#"<Default Extension="png" ContentType="image/png"/>"#);
    ct.push_str(r#"<Default Extension="bmp" ContentType="image/bmp"/>"#);
    ct.push_str(r#"<Default Extension="jpeg" ContentType="image/jpeg"/>"#);
    ct.push_str(r#"<Default Extension="svg" ContentType="image/svg+xml"/>"#);
    for i in 0..slides.len() {
        ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            i + 1
        ));
    }
    ct.push_str("</Types>");
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    // ppt/presentation.xml
    let mut pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
        slide_cx_emu, slide_cy_emu
    );
    for i in 0..slides.len() {
        pres.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            2 + i
        ));
    }
    pres.push_str("</p:sldIdLst></p:presentation>");
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    // ppt/_rels/presentation.xml.rels
    let mut pres_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for i in 0..slides.len() {
        pres_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
            2 + i,
            1 + i
        ));
    }
    pres_rels.push_str("</Relationships>");
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(pres_rels.as_bytes()).unwrap();

    // Slides and their .rels files
    for (i, (slide_xml, slide_images)) in slides.iter().enumerate() {
        let slide_num = i + 1;

        zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
            .unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        if !slide_images.is_empty() {
            let mut rels = String::from(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            );
            for img in slide_images {
                rels.push_str(&format!(
                    r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                    img.rid,
                    img.relationship_type.as_deref().unwrap_or(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                    ),
                    img.path
                ));
            }
            rels.push_str("</Relationships>");
            zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
                .unwrap();
            zip.write_all(rels.as_bytes()).unwrap();

            for img in slide_images {
                let media_path = resolve_relative_path("ppt/slides", &img.path);
                zip.start_file(media_path, opts).unwrap();
                zip.write_all(&img.data).unwrap();
            }
        }
    }

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

fn get_image(elem: &FixedElement) -> &ImageData {
    match &elem.kind {
        FixedElementKind::Image(img) => img,
        other => panic!("Expected Image, got {other:?}"),
    }
}

#[test]
fn test_image_basic_extraction() {
    let bmp_data = make_test_bmp();
    let pic = make_pic_xml(1_000_000, 500_000, 3_000_000, 2_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let elem = &page.elements[0];
    assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
    assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
    assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
    assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

    let img = get_image(elem);
    assert!(!img.data.is_empty(), "Image data should not be empty");
    assert_eq!(img.data, bmp_data);
}

#[test]
fn test_image_format_detection() {
    let bmp_data = make_test_bmp();
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Bmp);
}

#[test]
fn test_svg_image_extraction() {
    let svg_data = make_test_svg();
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.svg".to_string(),
        data: svg_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Svg);
    assert_eq!(img.data, svg_data);
}

#[test]
fn test_image_blip_start_tag_with_children_is_extracted() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data.clone(),
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 image element");

    let img = get_image(&page.elements[0]);
    assert_eq!(img.data, bmp_data);
}

#[test]
fn test_svg_blip_is_preferred_over_base_raster() {
    let bmp_data = make_test_bmp();
    let svg_data = make_test_svg();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId4"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.svg".to_string(),
            data: svg_data.clone(),
            relationship_type: None,
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(img.format, ImageFormat::Svg);
    assert_eq!(img.data, svg_data);
}

#[test]
fn test_src_rect_crop_is_extracted() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        2_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"/><a:srcRect l="25000" t="10000" r="5000" b="20000"/><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    assert_eq!(
        img.crop,
        Some(ImageCrop {
            left: 0.25,
            top: 0.10,
            right: 0.05,
            bottom: 0.20,
        })
    );
}

#[test]
fn test_unsupported_img_layer_emits_partial_warning_but_keeps_base_image() {
    let bmp_data = make_test_bmp();
    let pic = make_custom_pic_xml(
        0,
        0,
        1_000_000,
        1_000_000,
        r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId4"/></a14:imgProps></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
    );
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.wdp".to_string(),
            data: vec![0x00, 0x01, 0x02],
            relationship_type: Some(
                "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
            ),
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Base image should still render");
    assert_eq!(get_image(&page.elements[0]).data, bmp_data);
    assert!(
        warnings.iter().any(|warning| matches!(
            warning,
            ConvertWarning::PartialElement { format, element, detail }
                if format == "PPTX"
                    && element.contains("slide 1")
                    && detail.contains("image layer")
                    && detail.contains("image2.wdp")
        )),
        "Expected partial warning for unsupported image layer, got: {warnings:?}"
    );
}

#[test]
fn test_wdp_only_picture_emits_unsupported_warning() {
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.wdp".to_string(),
        data: vec![0x00, 0x01, 0x02],
        relationship_type: Some(
            "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
        ),
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        0,
        "Unsupported WDP image should be omitted"
    );
    assert!(
        warnings.iter().any(|warning| matches!(
            warning,
            ConvertWarning::UnsupportedElement { format, element }
                if format == "PPTX"
                    && element.contains("slide 1")
                    && element.contains("image1.wdp")
        )),
        "Expected unsupported warning for WDP-only picture, got: {warnings:?}"
    );
}

#[test]
fn test_image_dimensions_preserved() {
    let bmp_data = make_test_bmp();
    let pic = make_pic_xml(0, 0, 2_540_000, 1_270_000, "rId3");
    let slide_xml = make_slide_xml(&[pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let img = get_image(&page.elements[0]);
    let w = img.width.expect("Expected width");
    let h = img.height.expect("Expected height");
    assert!((w - 200.0).abs() < 0.1, "Expected ~200pt, got {w}");
    assert!((h - 100.0).abs() < 0.1, "Expected ~100pt, got {h}");
}

#[test]
fn test_image_with_shapes_and_text() {
    let bmp_data = make_test_bmp();
    let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Title");
    let rect = make_shape(
        0,
        600_000,
        1_000_000,
        500_000,
        "rect",
        Some("AABBCC"),
        None,
        None,
    );
    let pic = make_pic_xml(2_000_000, 600_000, 1_500_000, 1_000_000, "rId3");
    let slide_xml = make_slide_xml(&[text_box, rect, pic]);
    let slide_images = vec![TestSlideImage {
        rid: "rId3".to_string(),
        path: "../media/image1.bmp".to_string(),
        data: bmp_data,
        relationship_type: None,
    }];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 3, "Expected 3 elements");
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
    assert!(matches!(&page.elements[2].kind, FixedElementKind::Image(_)));
}

#[test]
fn test_image_missing_rid_ignored() {
    let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId99");
    let slide_xml = make_slide_xml(&[pic]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        0,
        "Missing image ref should be skipped"
    );
}

#[test]
fn test_multiple_images_on_slide() {
    let bmp_data = make_test_bmp();
    let pic1 = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
    let pic2 = make_pic_xml(2_000_000, 0, 1_500_000, 1_000_000, "rId4");
    let slide_xml = make_slide_xml(&[pic1, pic2]);
    let slide_images = vec![
        TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        },
        TestSlideImage {
            rid: "rId4".to_string(),
            path: "../media/image2.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        },
    ];
    let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 image elements");
    assert!(matches!(&page.elements[0].kind, FixedElementKind::Image(_)));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Image(_)));
}
