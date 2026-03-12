use super::image_tests::{
    TestSlideImage, build_test_pptx_with_images, get_image, make_pic_xml, make_test_bmp,
};
use super::*;

#[allow(clippy::too_many_arguments)]
fn make_group_shape(
    off_x: i64,
    off_y: i64,
    ext_cx: i64,
    ext_cy: i64,
    ch_off_x: i64,
    ch_off_y: i64,
    ch_ext_cx: i64,
    ch_ext_cy: i64,
    children: &[String],
) -> String {
    let mut xml = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/><a:chOff x="{ch_off_x}" y="{ch_off_y}"/><a:chExt cx="{ch_ext_cx}" cy="{ch_ext_cy}"/></a:xfrm></p:grpSpPr>"#
    );
    for child in children {
        xml.push_str(child);
    }
    xml.push_str("</p:grpSp>");
    xml
}

fn make_shape_rect(x: i64, y: i64, cx: i64, cy: i64, fill_hex: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill></p:spPr></p:sp>"#
    )
}

#[test]
fn test_group_shape_two_text_boxes() {
    let child_a = make_text_box(0, 0, 2_000_000, 1_000_000, "Shape A");
    let child_b = make_text_box(2_000_000, 1_000_000, 2_000_000, 1_000_000, "Shape B");
    let group = make_group_shape(
        1_000_000,
        500_000,
        4_000_000,
        2_000_000,
        0,
        0,
        4_000_000,
        2_000_000,
        &[child_a, child_b],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2);

    let first = &page.elements[0];
    assert!((first.x - emu_to_pt(1_000_000)).abs() < 0.1);
    assert!((first.y - emu_to_pt(500_000)).abs() < 0.1);

    let second = &page.elements[1];
    assert!((second.x - emu_to_pt(3_000_000)).abs() < 0.1);
    assert!((second.y - emu_to_pt(1_500_000)).abs() < 0.1);

    let paragraph = match &text_box_blocks(first)[0] {
        Block::Paragraph(paragraph) => paragraph,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(paragraph.runs[0].text, "Shape A");
}

#[test]
fn test_group_shape_with_scaling() {
    let child = make_text_box(0, 0, 4_000_000, 2_000_000, "Scaled");
    let group = make_group_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        0,
        0,
        4_000_000,
        2_000_000,
        &[child],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let element = &page.elements[0];
    assert!((element.width - emu_to_pt(2_000_000)).abs() < 0.1);
    assert!((element.height - emu_to_pt(1_000_000)).abs() < 0.1);
}

#[test]
fn test_nested_group_shapes() {
    let inner_child = make_text_box(0, 0, 1_000_000, 1_000_000, "Nested");
    let inner_group = make_group_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        0,
        0,
        2_000_000,
        2_000_000,
        &[inner_child],
    );
    let outer_group = make_group_shape(
        1_000_000,
        1_000_000,
        4_000_000,
        4_000_000,
        0,
        0,
        4_000_000,
        4_000_000,
        &[inner_group],
    );
    let slide = make_slide_xml(&[outer_group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let element = &page.elements[0];
    assert!((element.x - emu_to_pt(1_000_000)).abs() < 0.1);
    assert!((element.y - emu_to_pt(1_000_000)).abs() < 0.1);
    assert_eq!(element.width, emu_to_pt(1_000_000));
    assert_eq!(element.height, emu_to_pt(1_000_000));
}

#[test]
fn test_group_shape_mixed_element_types() {
    let text = make_text_box(0, 0, 2_000_000, 1_000_000, "Text");
    let rect = make_shape_rect(2_000_000, 0, 2_000_000, 1_000_000, "FF0000");
    let group = make_group_shape(
        0,
        0,
        4_000_000,
        2_000_000,
        0,
        0,
        4_000_000,
        2_000_000,
        &[text, rect],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2);
    assert!(matches!(
        &page.elements[0].kind,
        FixedElementKind::TextBox(_)
    ));
    assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
    assert!((page.elements[1].x - emu_to_pt(2_000_000)).abs() < 0.1);
}

#[test]
fn test_group_shape_with_nonzero_child_offset() {
    let child = make_text_box(1_000_000, 1_000_000, 2_000_000, 1_000_000, "Offset");
    let group = make_group_shape(
        500_000,
        500_000,
        4_000_000,
        2_000_000,
        1_000_000,
        1_000_000,
        4_000_000,
        2_000_000,
        &[child],
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    assert!((page.elements[0].x - emu_to_pt(500_000)).abs() < 0.1);
    assert!((page.elements[0].y - emu_to_pt(500_000)).abs() < 0.1);
}

#[test]
fn test_group_shape_scales_image_dimensions() {
    // Group scales child space by 0.5x, 0.5y:
    // ext = 2_000_000 x 1_000_000, chExt = 4_000_000 x 2_000_000
    let bmp_data = make_test_bmp();
    let pic = make_pic_xml(0, 0, 4_000_000, 2_000_000, "rId3");
    let group = make_group_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        0,
        0,
        4_000_000,
        2_000_000,
        &[pic],
    );
    let slide_xml = make_slide_xml(&[group]);
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
    assert_eq!(page.elements.len(), 1);

    let elem = &page.elements[0];
    let expected_w: f64 = emu_to_pt(2_000_000);
    let expected_h: f64 = emu_to_pt(1_000_000);
    // FixedElement dimensions should be scaled
    assert!((elem.width - expected_w).abs() < 0.1);
    assert!((elem.height - expected_h).abs() < 0.1);

    // Inner ImageData dimensions must also be scaled by the group transform
    let img = get_image(elem);
    let img_w: f64 = img.width.expect("Expected width");
    let img_h: f64 = img.height.expect("Expected height");
    assert!(
        (img_w - expected_w).abs() < 0.1,
        "ImageData.width should be {expected_w}, got {img_w}"
    );
    assert!(
        (img_h - expected_h).abs() < 0.1,
        "ImageData.height should be {expected_h}, got {img_h}"
    );
}
