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

#[test]
fn test_group_rotation_orbits_children_and_adds_shape_rotation() {
    // A group rotated 90° cw: a child shape at the group's right edge must
    // move to the bottom edge, and the child's own rotation gains the
    // group's angle.
    let child = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="1800000"><a:off x="1828800" y="762000"/><a:ext cx="457200" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#;
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="914400" y="914400"/><a:ext cx="2286000" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="2286000" cy="1828800"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"#
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];

    // Group center: (162, 144) pt. Child center maps to (234, 144) pt,
    // i.e. +72pt right of center; rotated 90° cw -> +72pt below -> (162, 216).
    let center_x = elem.x + elem.width / 2.0;
    let center_y = elem.y + elem.height / 2.0;
    assert!((center_x - 162.0).abs() < 0.5, "center x {center_x}");
    assert!((center_y - 216.0).abs() < 0.5, "center y {center_y}");

    let FixedElementKind::Shape(ref shape) = elem.kind else {
        panic!("expected shape");
    };
    // Child's own 30° plus the group's 90°.
    assert!((shape.rotation_deg.unwrap() - 120.0).abs() < 0.01);
}

/// A text box inside a rotated group composes the group's angle the same way
/// a shape does (issue #894). Before `TextBoxData` could hold a rotation, a
/// grouped label was orbited into place and then drawn horizontally.
#[test]
fn test_group_rotation_adds_to_a_child_text_box() {
    let child = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="1800000"><a:off x="1828800" y="762000"/><a:ext cx="457200" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Label</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="914400" y="914400"/><a:ext cx="2286000" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="2286000" cy="1828800"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"#
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::TextBox(ref text_box) = page.elements[0].kind else {
        panic!("expected a text box");
    };
    // The child's own 30° plus the group's 90°.
    let rotation = text_box
        .shape_rotation_deg
        .expect("a grouped text box must carry the composed rotation");
    assert!(
        (rotation - 120.0).abs() < 0.01,
        "expected 120 degrees, got {rotation}"
    );
}

#[test]
fn test_group_scaling_stretches_line_endpoints() {
    // Hairline axes bake their geometry in child-space points; without
    // endpoint scaling a 2x group left them as sub-pixel stubs.
    let child = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="L"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="500000"/><a:ext cx="1000000" cy="1"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dashDot"/></a:ln></p:spPr></p:sp>"#;
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="2000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"#
    );
    let slide = make_slide_xml(&[group]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref shape) = page.elements[0].kind else {
        panic!("expected shape");
    };
    let ShapeKind::Line { x2, .. } = shape.kind else {
        panic!("expected line, got {:?}", shape.kind);
    };
    // Child 1_000_000 EMU wide scaled 2x -> 2_000_000 EMU = 157.48pt.
    assert!((x2 - 157.48).abs() < 0.1, "x2 {x2}");
}

#[test]
fn group_transform_applies_anisotropic_scale_after_line_rotation() {
    fn to_page_point(element: &FixedElement, rotation_radians: f64, x: f64, y: f64) -> (f64, f64) {
        let center_x: f64 = element.width / 2.0;
        let center_y: f64 = element.height / 2.0;
        let dx: f64 = x - center_x;
        let dy: f64 = y - center_y;
        let (sin, cos): (f64, f64) = rotation_radians.sin_cos();
        (
            element.x + center_x + dx * cos - dy * sin,
            element.y + center_y + dx * sin + dy * cos,
        )
    }

    let child: String = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rotated line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="5400000"><a:off x="500000" y="500000"/><a:ext cx="1000000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="63500"><a:solidFill><a:srgbClr val="00AEEF"/></a:solidFill><a:round/></a:ln></p:spPr></p:sp>"#.to_string();
    let group: String = make_group_shape(
        100_000,
        200_000,
        2_000_000,
        3_000_000,
        0,
        0,
        1_000_000,
        1_000_000,
        &[child],
    );
    let slide_xml: String = make_slide_xml(&[group]);
    let data: Vec<u8> = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);

    let (document, _warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page: &FixedPage = first_fixed_page(&document);
    let element: &FixedElement = &page.elements[0];
    let FixedElementKind::Shape(shape) = &element.kind else {
        panic!("expected a transformed connector shape");
    };
    let ShapeKind::Line { x1, y1, x2, y2, .. } = &shape.kind else {
        panic!("expected a line, got {:?}", shape.kind);
    };

    let rotation_radians: f64 = shape.rotation_deg.unwrap_or(0.0).to_radians();
    let endpoints: [(f64, f64); 2] = [
        to_page_point(element, rotation_radians, *x1, *y1),
        to_page_point(element, rotation_radians, *x2, *y2),
    ];
    let (top, bottom): ((f64, f64), (f64, f64)) = if endpoints[0].1 <= endpoints[1].1 {
        (endpoints[0], endpoints[1])
    } else {
        (endpoints[1], endpoints[0])
    };

    assert!(
        (top.0 - emu_to_pt(2_100_000)).abs() < 0.1,
        "top x {}",
        top.0
    );
    assert!(
        (bottom.0 - emu_to_pt(2_100_000)).abs() < 0.1,
        "bottom x {}",
        bottom.0
    );
    assert!((top.1 - emu_to_pt(200_000)).abs() < 0.1, "top y {}", top.1);
    assert!(
        (bottom.1 - emu_to_pt(3_200_000)).abs() < 0.1,
        "bottom y {}",
        bottom.1
    );
}
