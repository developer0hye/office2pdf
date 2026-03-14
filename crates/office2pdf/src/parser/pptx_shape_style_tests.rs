use super::*;

#[test]
fn test_shape_outline_dash_style() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape_elem = &page.elements[0];
    if let FixedElementKind::Shape(ref s) = shape_elem.kind {
        let stroke = s.stroke.as_ref().expect("Expected stroke");
        assert_eq!(
            stroke.style,
            BorderLineStyle::Dashed,
            "Shape stroke should be dashed"
        );
    } else {
        panic!("Expected Shape element");
    }
}

// ── Shape style (rotation, transparency) test helpers ────────────────

#[allow(clippy::too_many_arguments)]
fn make_styled_shape(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    fill_hex: Option<&str>,
    rot: Option<i64>,
    alpha_thousandths: Option<i64>,
) -> String {
    let rot_attr = rot.map(|r| format!(r#" rot="{r}""#)).unwrap_or_default();

    let fill_xml = match (fill_hex, alpha_thousandths) {
        (Some(h), Some(a)) => format!(
            r#"<a:solidFill><a:srgbClr val="{h}"><a:alpha val="{a}"/></a:srgbClr></a:solidFill>"#
        ),
        (Some(h), None) => {
            format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#)
        }
        _ => String::new(),
    };

    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}</p:spPr></p:sp>"#
    )
}

// ── Shape style tests (US-034) ──────────────────────────────────────

#[test]
fn test_shape_rotation() {
    let shape = make_styled_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "rect",
        Some("FF0000"),
        Some(5_400_000),
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(s.rotation_deg.is_some(), "Expected rotation_deg to be set");
    assert!(
        (s.rotation_deg.unwrap() - 90.0).abs() < 0.01,
        "Expected 90°, got {}",
        s.rotation_deg.unwrap()
    );
}

#[test]
fn test_shape_transparency() {
    let shape = make_styled_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "rect",
        Some("00FF00"),
        None,
        Some(50_000),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(s.opacity.is_some(), "Expected opacity to be set");
    assert!(
        (s.opacity.unwrap() - 0.5).abs() < 0.01,
        "Expected 0.5 opacity, got {}",
        s.opacity.unwrap()
    );
}

#[test]
fn test_shape_rotation_and_transparency() {
    let shape = make_styled_shape(
        1_000_000,
        500_000,
        3_000_000,
        2_000_000,
        "ellipse",
        Some("0000FF"),
        Some(2_700_000),
        Some(75_000),
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let s = get_shape(&page.elements[0]);
    assert!(
        (s.rotation_deg.unwrap() - 45.0).abs() < 0.01,
        "Expected 45°, got {}",
        s.rotation_deg.unwrap()
    );
    assert!(
        (s.opacity.unwrap() - 0.75).abs() < 0.01,
        "Expected 0.75 opacity, got {}",
        s.opacity.unwrap()
    );
    assert!(matches!(s.kind, ShapeKind::Ellipse));
}

// ── Gradient background tests (US-050) ──────────────────────────────

#[test]
fn test_gradient_background_two_stops() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient background");
    assert_eq!(gradient.stops.len(), 2);
    assert!((gradient.stops[0].position - 0.0).abs() < 0.001);
    assert_eq!(gradient.stops[0].color, Color::new(255, 0, 0));
    assert!((gradient.stops[1].position - 1.0).abs() < 0.001);
    assert_eq!(gradient.stops[1].color, Color::new(0, 0, 255));
    assert!((gradient.angle - 90.0).abs() < 0.001);
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_gradient_background_three_stops() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert_eq!(gradient.stops.len(), 3);
    assert!((gradient.stops[1].position - 0.5).abs() < 0.001);
    assert_eq!(gradient.stops[1].color, Color::new(0, 255, 0));
    assert!((gradient.angle - 0.0).abs() < 0.001);
}

#[test]
fn test_gradient_background_with_scheme_colors() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert_eq!(gradient.stops.len(), 2);
    assert!((gradient.angle - 45.0).abs() < 0.001);
}

#[test]
fn test_gradient_filled_shape_keeps_following_siblings() {
    let before = make_text_box(0, 0, 2_000_000, 600_000, "Before");
    let gradient_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="30" name="GradientShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="0" y="800000"/><a:ext cx="3000000" cy="1200000"/></a:xfrm>"#,
        r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
        r#"<a:gradFill flip="none" rotWithShape="1"><a:gsLst>"#,
        r#"<a:gs pos="0"><a:srgbClr val="367482"/></a:gs>"#,
        r#"<a:gs pos="100000"><a:srgbClr val="306572"/></a:gs>"#,
        r#"</a:gsLst><a:lin ang="5400000" scaled="1"/><a:tileRect/></a:gradFill>"#,
        r#"</p:spPr></p:sp>"#
    )
    .to_string();
    let after = make_text_box(0, 2_400_000, 2_500_000, 600_000, "After");
    let slide = make_slide_xml(&[before, gradient_shape, after]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        3,
        "Gradient-filled shapes must not consume later siblings: {:#?}",
        page.elements
    );

    let last_text = match &page.elements[2].kind {
        FixedElementKind::TextBox(text_box) => match &text_box.content[0] {
            Block::Paragraph(paragraph) => paragraph.runs[0].text.clone(),
            other => panic!("Expected paragraph block, got {other:?}"),
        },
        other => panic!("Expected final sibling text box, got {other:?}"),
    };
    assert_eq!(last_text, "After");
}

#[test]
fn test_gradient_text_shape_with_style_keeps_following_siblings() {
    let before = make_text_box(0, 0, 2_000_000, 600_000, "Before");
    let gradient_text_shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="31" name="StyledGradientShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm flipV="1"><a:off x="0" y="800000"/><a:ext cx="3000000" cy="200000"/></a:xfrm>"#,
        r#"<a:prstGeom prst="trapezoid"><a:avLst/></a:prstGeom>"#,
        r#"<a:gradFill flip="none" rotWithShape="1"><a:gsLst>"#,
        r#"<a:gs pos="0"><a:srgbClr val="FFFFFF"><a:alpha val="70000"/></a:srgbClr></a:gs>"#,
        r#"<a:gs pos="76000"><a:srgbClr val="FFFFFF"><a:alpha val="29000"/></a:srgbClr></a:gs>"#,
        r#"<a:gs pos="92000"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs>"#,
        r#"</a:gsLst><a:lin ang="16200000" scaled="1"/><a:tileRect/></a:gradFill><a:ln><a:noFill/></a:ln></p:spPr>"#,
        r#"<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>"#,
        r#"<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#
    )
    .to_string();
    let after = make_text_box(0, 1_400_000, 2_500_000, 600_000, "After");
    let slide = make_slide_xml(&[before, gradient_text_shape, after]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        3,
        "Styled gradient text shapes must not consume later siblings: {:#?}",
        page.elements
    );

    let last_text = match &page.elements[2].kind {
        FixedElementKind::TextBox(text_box) => match &text_box.content[0] {
            Block::Paragraph(paragraph) => paragraph.runs[0].text.clone(),
            other => panic!("Expected paragraph block, got {other:?}"),
        },
        other => panic!("Expected final sibling text box, got {other:?}"),
    };
    assert_eq!(last_text, "After");
}

#[test]
fn test_solid_background_no_gradient() {
    let bg_xml =
        r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    assert!(
        page.background_gradient.is_none(),
        "Solid fill should not produce gradient"
    );
    assert_eq!(page.background_color, Some(Color::new(255, 204, 0)));
}

#[test]
fn test_gradient_shape_fill() {
    let shape_xml =
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></p:spPr></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    assert_eq!(page.elements.len(), 1);
    let shape = get_shape(&page.elements[0]);
    let gf = shape
        .gradient_fill
        .as_ref()
        .expect("Expected gradient fill on shape");
    assert_eq!(gf.stops.len(), 2);
    assert_eq!(gf.stops[0].color, Color::new(255, 0, 0));
    assert_eq!(gf.stops[1].color, Color::new(0, 255, 0));
    assert!((gf.angle - 90.0).abs() < 0.001);
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_shape_solid_fill_no_gradient() {
    let shape_xml =
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    assert!(
        shape.gradient_fill.is_none(),
        "Solid fill shape should have no gradient"
    );
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_gradient_background_no_angle() {
    let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill></p:bgPr></p:bg>"#;
    let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let gradient = page
        .background_gradient
        .as_ref()
        .expect("Expected gradient");
    assert!(
        (gradient.angle - 0.0).abs() < 0.001,
        "Default angle should be 0"
    );
}

// ── Shadow / effects tests ─────────────────────────────────────────

#[test]
fn test_shape_outer_shadow_parsed() {
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    let shadow = shape.shadow.as_ref().expect("Expected shadow");
    assert!(
        (shadow.blur_radius - 4.0).abs() < 0.01,
        "Expected blur_radius ~4.0, got {}",
        shadow.blur_radius
    );
    assert!(
        (shadow.distance - 3.0).abs() < 0.01,
        "Expected distance ~3.0, got {}",
        shadow.distance
    );
    assert!(
        (shadow.direction - 45.0).abs() < 0.01,
        "Expected direction ~45.0, got {}",
        shadow.direction
    );
    assert_eq!(shadow.color, Color::new(0, 0, 0));
    assert!(
        (shadow.opacity - 0.5).abs() < 0.01,
        "Expected opacity ~0.5, got {}",
        shadow.opacity
    );
}

#[test]
fn test_shape_no_effects_no_shadow() {
    let shape_xml = make_shape(
        100_000,
        200_000,
        500_000,
        300_000,
        "rect",
        Some("00FF00"),
        None,
        None,
    );
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    assert!(
        shape.shadow.is_none(),
        "Shape without effectLst should have no shadow"
    );
}

#[test]
fn test_shape_shadow_default_opacity() {
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="25400" dist="12700" dir="5400000"><a:srgbClr val="333333"/></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    let shadow = shape.shadow.as_ref().expect("Expected shadow");
    assert!(
        (shadow.blur_radius - 2.0).abs() < 0.01,
        "Expected blur ~2.0, got {}",
        shadow.blur_radius
    );
    assert!(
        (shadow.distance - 1.0).abs() < 0.01,
        "Expected dist ~1.0, got {}",
        shadow.distance
    );
    assert!(
        (shadow.direction - 90.0).abs() < 0.01,
        "Expected dir ~90.0, got {}",
        shadow.direction
    );
    assert_eq!(shadow.color, Color::new(0x33, 0x33, 0x33));
    assert!(
        (shadow.opacity - 1.0).abs() < 0.01,
        "Expected opacity ~1.0 (default), got {}",
        shadow.opacity
    );
}

// ── fillRef style fallback tests ─────────────────────────────────

#[test]
fn test_shape_fill_from_style_fill_ref() {
    // Shape with no explicit fill, but <p:style><a:fillRef> referencing accent1.
    // accent1 = #4472C4 in standard theme.
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    // accent1 = #4472C4
    assert_eq!(
        shape.fill,
        Some(Color::new(0x44, 0x72, 0xC4)),
        "Shape should get fill from fillRef accent1"
    );
}

#[test]
fn test_shape_explicit_fill_overrides_fill_ref() {
    // Shape with explicit solidFill AND <p:style><a:fillRef>.
    // Explicit fill should win.
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(
        shape.fill,
        Some(Color::new(255, 0, 0)),
        "Explicit solidFill should override fillRef"
    );
}

#[test]
fn test_shape_no_fill_overrides_fill_ref() {
    // Shape with explicit <a:noFill/> AND <p:style><a:fillRef>.
    // noFill should prevent style fallback.
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(
        shape.fill, None,
        "noFill should prevent style fillRef fallback"
    );
}

#[test]
fn test_textbox_fill_from_style_fill_ref() {
    // TextBox with roundRect (non-rectangular shape) and text gets split into
    // two elements: Shape background (with fill) + transparent TextBox overlay.
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    // First element: Shape background with geometry and fill
    assert_eq!(page.elements.len(), 2, "Expected Shape + TextBox pair");
    let shape = get_shape(&page.elements[0]);
    assert_eq!(
        shape.fill,
        Some(Color::new(0x44, 0x72, 0xC4)),
        "Shape background should get fill from fillRef accent1"
    );
    assert!(matches!(shape.kind, ShapeKind::RoundedRectangle { .. }));
    // Second element: Transparent text overlay
    let tb = text_box_data(&page.elements[1]);
    assert_eq!(tb.fill, None, "Text overlay should have no fill");
}

#[test]
fn test_split_textbox_preserves_alignment() {
    // roundRect with centered text, solidFill, and bodyPr anchor="ctr".
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1696720" cy="650158"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="003481"/></a:solidFill></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>Random Sample</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    // Should be split into Shape + TextBox
    assert_eq!(page.elements.len(), 2, "Expected Shape + TextBox pair");

    // TextBox overlay should preserve vertical and horizontal alignment
    let tb = text_box_data(&page.elements[1]);
    assert_eq!(
        tb.vertical_align,
        TextBoxVerticalAlign::Center,
        "Vertical align should be Center"
    );
    // Check paragraph alignment
    let para = match &tb.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(
        para.style.alignment,
        Some(Alignment::Center),
        "Paragraph alignment should be Center"
    );
    assert_eq!(
        para.runs[0].text, "Random Sample",
        "Text content should be preserved"
    );

    // Verify Typst output contains #align(center)
    let typst_output = crate::render::typst_gen::generate_typst(&doc).unwrap();
    assert!(
        typst_output.source.contains("#set align(center)"),
        "Typst output should contain #set align(center) for centered paragraph, got:\n{}",
        typst_output.source,
    );
}
