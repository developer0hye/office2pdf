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

/// Every `a:prstDash` preset survives parsing as its own style (issue #758).
///
/// Ten presets used to reach the IR as four styles, so a deck's `lgDash` and
/// `sysDash` lines became one line. Asserting distinctness rather than naming
/// each variant keeps this about the observable outcome: whichever styles the
/// IR grows, two different presets must never arrive as the same one.
#[test]
fn each_preset_dash_reaches_the_ir_as_a_distinct_style() {
    const PRESETS: [&str; 10] = [
        "dot",
        "sysDot",
        "dash",
        "sysDash",
        "lgDash",
        "dashDot",
        "sysDashDot",
        "lgDashDot",
        "lgDashDotDot",
        "sysDashDotDot",
    ];
    let shapes: Vec<String> = PRESETS
        .iter()
        .enumerate()
        .map(|(index, preset)| {
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{preset}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{y}"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="{preset}"/></a:ln></p:spPr></p:sp>"#,
                id = index + 2,
                y = index as i64 * 457200,
            )
        })
        .collect();
    let slide = make_slide_xml(&shapes);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let (doc, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the probe deck parses");

    let page = first_fixed_page(&doc);
    let styles: Vec<BorderLineStyle> = page
        .elements
        .iter()
        .filter_map(|element| match element.kind {
            FixedElementKind::Shape(ref shape) => Some(shape.stroke.as_ref()?.style),
            _ => None,
        })
        .collect();
    assert_eq!(styles.len(), PRESETS.len(), "one stroked line per preset");
    for (index, style) in styles.iter().enumerate() {
        assert!(
            !styles[..index].contains(style),
            "{} and an earlier preset both parse as {style:?}",
            PRESETS[index],
        );
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

/// The page-9 title rule in the public fixture is an open custom-geometry
/// line on the top edge of a vertically flipped shape box (issue #1418).
/// The shape transform must move that line to the bottom edge without
/// changing its normalized horizontal span.
#[test]
fn a_custom_geometry_path_follows_shape_flips() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="400000" cy="200000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="200000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="200000" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="F0CDA1"/></a:solidFill></a:ln></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the custom line deck parses");
    let page = first_fixed_page(&document);
    let FixedElementKind::Shape(parsed_shape) = &page.elements[0].kind else {
        panic!("expected a custom Shape, got {:?}", page.elements[0].kind);
    };
    let ShapeKind::Path { subpaths } = &parsed_shape.kind else {
        panic!("expected a custom path, got {:?}", parsed_shape.kind);
    };

    assert_eq!(subpaths.len(), 1);
    let vertices = &subpaths[0].vertices;
    assert_eq!(vertices, &vec![(0.0, 1.0), (1.0, 1.0)]);
}

/// Triangulation: horizontal and vertical flips are independent shape-box
/// transforms, not a title-rule-specific vertical offset.
#[test]
fn a_custom_geometry_horizontal_flip_mirrors_only_x() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Diagonal"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm flipH="true"><a:off x="0" y="0"/><a:ext cx="400000" cy="200000"/></a:xfrm><a:custGeom><a:pathLst><a:path w="400000" h="200000"><a:moveTo><a:pt x="100000" y="50000"/></a:moveTo><a:lnTo><a:pt x="300000" y="150000"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the custom line deck parses");
    let page = first_fixed_page(&document);
    let FixedElementKind::Shape(parsed_shape) = &page.elements[0].kind else {
        panic!("expected a custom Shape, got {:?}", page.elements[0].kind);
    };
    let ShapeKind::Path { subpaths } = &parsed_shape.kind else {
        panic!("expected a custom path, got {:?}", parsed_shape.kind);
    };

    assert_eq!(subpaths[0].vertices, vec![(0.75, 0.25), (0.25, 0.75)]);
}

/// A shape box may be zero-height. The page-8 title rule of the deck on
/// issue #1447 is `<a:ext cx="3708000" cy="0"/>` around an open two-point
/// custom geometry, and its `<a:path w=…>` declares no `h`. That subpath used
/// to be discarded for want of a vertical coordinate space, leaving the
/// rectangle fallback to stroke a closed zero-height box whose round corner
/// joins printed as semicircular ends where PowerPoint draws flat ones.
#[test]
fn a_zero_height_custom_geometry_rule_stays_an_open_line() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="942535" y="1337304"/><a:ext cx="3708000" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="3218815"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="3218395" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="54863"><a:solidFill><a:srgbClr val="F0CDA1"/></a:solidFill></a:ln></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the zero-height rule deck parses");
    let page = first_fixed_page(&document);
    let FixedElementKind::Shape(parsed_shape) = &page.elements[0].kind else {
        panic!("expected a custom Shape, got {:?}", page.elements[0].kind);
    };
    let ShapeKind::Path { subpaths } = &parsed_shape.kind else {
        panic!(
            "a zero-height rule must keep its open path, got {:?}",
            parsed_shape.kind
        );
    };

    assert_eq!(subpaths.len(), 1, "got {subpaths:?}");
    assert!(
        !subpaths[0].closed,
        "the rule declares no a:close, so its stroke must not join back"
    );
    assert_eq!(subpaths[0].vertices.len(), 2, "got {:?}", subpaths[0]);
    assert_eq!(subpaths[0].vertices[0], (0.0, 0.0));
    assert!(
        (subpaths[0].vertices[1].0 - 3_218_395.0 / 3_218_815.0).abs() < 1e-9,
        "the horizontal span comes from the path's own space, got {:?}",
        subpaths[0].vertices[1]
    );
    assert_eq!(subpaths[0].vertices[1].1, 0.0);
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

// ── Theme shape fill-reference tests ───────────────────────────────

fn make_theme_xml_with_shape_fill_styles() -> String {
    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
    let format_scheme = concat!(
        r#"<a:fmtScheme name="Test">"#,
        r#"<a:fillStyleLst>"#,
        r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
        r#"<a:gradFill><a:gsLst>"#,
        r#"<a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>"#,
        r#"<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs>"#,
        r#"</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>"#,
        r#"<a:gradFill rotWithShape="1"><a:gsLst>"#,
        r#"<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>"#,
        r#"<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>"#,
        r#"<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>"#,
        r#"</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>"#,
        r#"</a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/>"#,
        r#"</a:fmtScheme>"#,
    );
    theme_xml.replacen(
        "</a:themeElements>",
        &format!("{format_scheme}</a:themeElements>"),
        1,
    )
}

#[test]
fn shape_fill_ref_selects_theme_gradient_and_resolves_placeholder_color() {
    let shape_xml = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Gradient style"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>"#,
        r#"<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>"#,
        r#"</p:sp>"#,
    )
    .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx_with_theme(
        SLIDE_CX,
        SLIDE_CY,
        &[slide_xml],
        &make_theme_xml_with_shape_fill_styles(),
    );

    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the style-matrix probe parses");
    let page = first_fixed_page(&document);
    let shape = get_shape(&page.elements[0]);
    let gradient = shape
        .gradient_fill
        .as_ref()
        .expect("fillRef idx=3 must select the third theme fill style");

    assert_eq!(gradient.stops.len(), 3);
    assert!((gradient.angle - 270.0).abs() < 0.001);
    assert!(
        gradient
            .stops
            .windows(2)
            .any(|pair| pair[0].color != pair[1].color),
        "phClr transforms must produce a visible ramp"
    );
    assert_ne!(gradient.stops[0].color, Color::new(0x44, 0x72, 0xC4));
}

#[test]
fn theme_gradient_fill_on_text_rectangle_is_rendered_behind_its_text() {
    let shape_xml = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Gradient title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="5000000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>"#,
        r#"<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>"#,
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Gradient title</a:t></a:r></a:p></p:txBody></p:sp>"#,
    )
    .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx_with_theme(
        SLIDE_CX,
        SLIDE_CY,
        &[slide_xml],
        &make_theme_xml_with_shape_fill_styles(),
    );

    let (document, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("the styled text rectangle parses");
    let page = first_fixed_page(&document);

    assert_eq!(page.elements.len(), 2, "background plus text overlay");
    let background = get_shape(&page.elements[0]);
    let gradient = background
        .gradient_fill
        .as_ref()
        .expect("the theme gradient must reach the shape renderer");
    assert_eq!(gradient.stops.len(), 2);
    assert!((gradient.angle - 90.0).abs() < 0.001);
    match &page.elements[1].kind {
        FixedElementKind::TextBox(text_box) => assert!(text_box.fill.is_none()),
        other => panic!("expected transparent text overlay, got {other:?}"),
    }
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
fn test_pattern_shape_fill_preserves_preset_and_colors() {
    let shape_xml =
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="PatternBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:pattFill prst="ltUpDiag"><a:fgClr><a:srgbClr val="0000FF"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></p:spPr></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    let shape = get_shape(&page.elements[0]);
    let pattern = shape
        .pattern_fill
        .as_ref()
        .expect("Expected pattern fill on shape");
    assert_eq!(pattern.preset, PatternPreset::LightUpwardDiagonal);
    assert_eq!(pattern.foreground, Color::new(0, 0, 255));
    assert_eq!(pattern.background, Color::new(255, 255, 255));
    assert!(shape.fill.is_none());
    assert!(shape.gradient_fill.is_none());
    assert!(shape.stroke.is_none());
}

#[test]
fn test_pattern_filled_rectangle_keeps_text_as_overlay() {
    let shape_xml =
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="PatternBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:pattFill prst="cross"><a:fgClr><a:srgbClr val="0000FF"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Pattern text</a:t></a:r></a:p></p:txBody></p:sp>"#
            .to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);

    assert_eq!(page.elements.len(), 2);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(
        shape.pattern_fill.as_ref().map(|fill| fill.preset),
        Some(PatternPreset::Cross)
    );
    assert!(matches!(
        page.elements[1].kind,
        FixedElementKind::TextBox(_)
    ));
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
fn test_shape_extension_hidden_line_does_not_override_visible_fill() {
    let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Ellipse"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst><a:outerShdw blurRad="63500" sx="102000" sy="102000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="39999"/></a:srgbClr></a:outerShdw></a:effectLst><a:extLst><a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}"><a16:hiddenLine xmlns:a16="http://schemas.microsoft.com/office/drawing/2010/main" w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/><a:headEnd/><a:tailEnd/></a16:hiddenLine></a:ext></a:extLst></p:spPr></p:sp>"#.to_string();
    let slide_xml = make_slide_xml(&[shape_xml]);

    let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert_eq!(
        shape.fill,
        Some(Color::new(0xFF, 0xFF, 0xFF)),
        "Vendor extension hiddenLine should not override the visible shape fill"
    );
    let shadow = shape.shadow.as_ref().expect("Expected shadow");
    assert_eq!(shadow.color, Color::new(0, 0, 0));
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

#[test]
fn test_shape_style_lnref_outline_resolves_width_and_shaded_color() {
    // A shape whose outline comes only from <p:style><a:lnRef idx=..> with a
    // shaded scheme color (a Start-event color) must render a stroke: width
    // from the theme lnStyleLst, color from the resolved scheme (issue #318).
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X"><a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    </a:lnStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref s) = page.elements[0].kind else {
        panic!("expected shape");
    };
    let stroke = s.stroke.as_ref().expect("lnRef must produce a stroke");
    // idx=2 → theme lnStyleLst[1] = 12700 EMU = 1pt.
    assert!(
        (stroke.width - 1.0).abs() < 0.01,
        "outline width from theme lnStyleLst idx 2, got {}",
        stroke.width
    );
    // accent1 (4472C4) shaded 50% in linear light (issue #667).
    assert_eq!(stroke.color, Color::new(0x2F, 0x52, 0x8F));
}

// ── `<a:ln><a:noFill/></a:ln>` disables the outline (issue #516) ─────

#[test]
fn test_no_fill_line_suppresses_style_outline_fallback() {
    // PowerPoint writes this exact construct for "No line"; the
    // `<p:style><a:lnRef>` fallback must not repaint the outline.
    let shape = concat!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
        r#"<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2794000" cy="1651000"/></a:xfrm>"#,
        r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
        r#"<a:solidFill><a:srgbClr val="80B0D8"/></a:solidFill>"#,
        r#"<a:ln w="19050" cap="flat"><a:noFill/><a:prstDash val="solid"/></a:ln></p:spPr>"#,
        r#"<p:style><a:lnRef idx="2"><a:srgbClr val="000000"/></a:lnRef>"#,
        r#"<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>"#,
        r#"<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>"#,
        r#"<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>"#
    )
    .to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = match &page.elements[0].kind {
        FixedElementKind::Shape(s) => s,
        other => panic!("Expected shape, got {other:?}"),
    };
    assert!(
        shape
            .stroke
            .as_ref()
            .is_some_and(|stroke| stroke.style == BorderLineStyle::None),
        "a:noFill inside a:ln must remain distinct from an omitted a:ln, got {:?}",
        shape.stroke
    );
    assert!(shape.fill.is_some(), "the shape fill must be unaffected");
}

/// A rotated shape that carries text keeps its rotation (issue #894).
///
/// A plain rectangle with text becomes a `TextBoxData` rather than a `Shape`,
/// and `TextBoxData` had nowhere to put `a:xfrm/@rot`, so the rotation was
/// dropped: LibreOffice writes a `0 24 -24 0` text matrix for a 270° box where
/// we wrote the unrotated `24 0 0 -24`.
#[test]
fn a_rotated_text_shape_keeps_its_rotation() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="16200000"><a:off x="1000000" y="1000000"/><a:ext cx="4000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>ROTATED</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let FixedElementKind::TextBox(ref text_box) = page.elements[0].kind else {
        panic!("expected a TextBox, got {:?}", page.elements[0].kind);
    };
    let rotation = text_box
        .shape_rotation_deg
        .expect("a rotated text shape must carry its rotation");
    assert!(
        (rotation - 270.0).abs() < 0.01,
        "expected 270 degrees, got {rotation}"
    );
    assert_eq!(
        text_box.text_rotation_deg, None,
        "the box rotates, the glyphs inside it do not"
    );
}

/// Triangulation: a different `rot` must produce that angle, not a fixed one,
/// and an unrotated text shape must carry no rotation at all.
#[test]
fn a_text_shape_rotation_follows_the_declared_angle() {
    for (rot_attr, expected) in [(Some(2_700_000_i64), Some(45.0_f64)), (None, None)] {
        let xfrm = match rot_attr {
            Some(rot) => format!(r#"<a:xfrm rot="{rot}">"#),
            None => "<a:xfrm>".to_string(),
        };
        let shape = format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>{xfrm}<a:off x="0" y="0"/><a:ext cx="2000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Tilted</a:t></a:r></a:p></p:txBody></p:sp>"#
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);
        let FixedElementKind::TextBox(ref text_box) = page.elements[0].kind else {
            panic!("expected a TextBox");
        };
        match (text_box.shape_rotation_deg, expected) {
            (None, None) => {}
            (Some(actual), Some(want)) => assert!(
                (actual - want).abs() < 0.01,
                "expected {want} degrees, got {actual}"
            ),
            (actual, want) => panic!("expected {want:?}, got {actual:?}"),
        }
    }
}

// ── `a:ln` corner join (issue #1090) ─────────────────────────────────

/// Build a one-shape deck whose outline is the given `<a:ln>` body and return
/// the parsed stroke.
fn shape_stroke_for_line(ln_xml: &str) -> BorderSide {
    let shape = format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/>{ln_xml}</a:ln></p:spPr></p:sp>"#
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref s) = page.elements[0].kind else {
        panic!("expected a Shape");
    };
    s.stroke.as_ref().expect("expected a stroke").clone()
}

/// An `a:ln` naming none of `a:round`, `a:bevel` or `a:miter` rounds its
/// corners. Measured on a native macOS PowerPoint export of `customGeo.pptx`
/// page 46, whose banner takes a theme line with no join child: the export
/// traces as `linejoin="1"`, PDF's round join (issue #1090).
#[test]
fn a_line_without_a_join_child_rounds_its_corners() {
    assert_eq!(shape_stroke_for_line("").join, LineJoin::Round);
}

/// Triangulation: each of the three join children is carried through as
/// itself, so the default above cannot be a constant.
#[test]
fn each_stated_join_child_reaches_the_stroke() {
    assert_eq!(
        shape_stroke_for_line(r#"<a:miter lim="800000"/>"#).join,
        LineJoin::Miter
    );
    assert_eq!(shape_stroke_for_line("<a:bevel/>").join, LineJoin::Bevel);
    assert_eq!(shape_stroke_for_line("<a:round/>").join, LineJoin::Round);
}

/// A theme `<a:lnStyleLst>` entry carries its join to any shape that takes its
/// outline from `<a:lnRef idx>`. PowerPoint's own themes state
/// `<a:miter lim="800000"/>` there, so treating the reference as joinless
/// would round every outline in a deck built on one.
#[test]
fn a_theme_line_style_carries_its_join_through_lnref() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X"><a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:miter lim="800000"/></a:ln>
      <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:miter lim="800000"/></a:ln>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:bevel/></a:ln>
    </a:lnStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let styled_shape = |idx: u32, ln: &str| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{idx}" name="S{idx}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{}"/><a:ext cx="914400" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>{ln}</p:spPr><p:style><a:lnRef idx="{idx}"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#,
            idx as i64 * 600_000
        )
    };
    let slide = make_slide_xml(&[
        // idx 2 → theme miter, no `a:ln` of its own.
        styled_shape(2, ""),
        // idx 3 → theme bevel, no `a:ln` of its own.
        styled_shape(3, ""),
    ]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let joins: Vec<LineJoin> = page
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::Shape(s) => s.stroke.as_ref().map(|stroke| stroke.join),
            _ => None,
        })
        .collect();
    assert_eq!(joins, vec![LineJoin::Miter, LineJoin::Bevel]);
}

/// A shape's own `a:ln` overrides the join it would otherwise inherit from the
/// theme line its `<a:lnRef>` names.
#[test]
fn a_stated_join_overrides_the_theme_line_style() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X"><a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:miter lim="800000"/></a:ln>
      <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:miter lim="800000"/></a:ln>
    </a:lnStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref s) = page.elements[0].kind else {
        panic!("expected a Shape");
    };
    assert_eq!(
        s.stroke.as_ref().expect("expected a stroke").join,
        LineJoin::Round
    );
}

// ── `a:ln/@cap` end geometry (issue #1682) ───────────────────────────

/// Build a one-shape deck whose outline declares `attributes` on its `<a:ln>`
/// and return the parsed stroke.
fn shape_stroke_for_line_attributes(attributes: &str) -> BorderSide {
    let shape = format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln {attributes}><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></p:spPr></p:sp>"#
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref s) = page.elements[0].kind else {
        panic!("expected a Shape");
    };
    s.stroke.as_ref().expect("expected a stroke").clone()
}

/// An `a:ln` naming no `cap` ends its stroke flat, and so does `cap="flat"`.
///
/// Measured on a native macOS PowerPoint export of `1-slide.pptx`'s 3.5pt
/// straight connector: both the unpatched deck and a `cap="flat"` variant trace
/// as PDF `linecap="0,0,0"`, butt ends (issue #1682).
#[test]
fn an_omitted_or_flat_cap_ends_the_stroke_flat() {
    assert_eq!(
        shape_stroke_for_line_attributes(r#"w="38100""#).cap,
        LineCap::Flat
    );
    assert_eq!(
        shape_stroke_for_line_attributes(r#"w="38100" cap="flat""#).cap,
        LineCap::Flat
    );
}

/// Triangulation: `rnd` and `sq` are their own end geometries at any width, so
/// the flat default above cannot be a constant. The same export traces
/// `linecap="1,1,1"` for `cap="rnd"` and `linecap="2,2,2"` for `cap="sq"`.
#[test]
fn each_stated_cap_reaches_the_stroke() {
    for width_emu in ["12700", "38100", "54863"] {
        assert_eq!(
            shape_stroke_for_line_attributes(&format!(r#"w="{width_emu}" cap="rnd""#)).cap,
            LineCap::Round,
            "cap=rnd at w={width_emu}"
        );
        assert_eq!(
            shape_stroke_for_line_attributes(&format!(r#"w="{width_emu}" cap="sq""#)).cap,
            LineCap::Square,
            "cap=sq at w={width_emu}"
        );
    }
}

/// A `cap` DrawingML does not define states nothing, so the flat default holds
/// rather than the attribute being read positionally.
#[test]
fn an_unknown_cap_value_keeps_the_flat_default() {
    assert_eq!(
        shape_stroke_for_line_attributes(r#"w="38100" cap="rounded""#).cap,
        LineCap::Flat
    );
}

/// A theme `<a:lnStyleLst>` entry carries its `cap` to any shape that takes its
/// outline from `<a:lnRef idx>`.
///
/// Measured: patching only `1-slide.pptx`'s theme `lnStyleLst` entry 1 from
/// `cap="flat"` to `cap="rnd"` turns every stroke whose `<a:lnRef idx="1">`
/// reaches it — the slide's two connectors and the layout's two — from
/// `linecap="0,0,0"` to `linecap="1,1,1"` in the native export, even though no
/// shape states a cap of its own (issue #1682).
#[test]
fn a_theme_line_style_carries_its_cap_through_lnref() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X"><a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="12700" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="19050" cap="sq"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    </a:lnStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let styled_shape = |idx: u32| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{idx}" name="S{idx}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{}"/><a:ext cx="914400" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr><p:style><a:lnRef idx="{idx}"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#,
            idx as i64 * 600_000
        )
    };
    // idx 1 → theme entry with no cap, idx 2 → `rnd`, idx 3 → `sq`.
    let slide = make_slide_xml(&[styled_shape(1), styled_shape(2), styled_shape(3)]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let caps: Vec<LineCap> = page
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::Shape(s) => s.stroke.as_ref().map(|stroke| stroke.cap),
            _ => None,
        })
        .collect();
    assert_eq!(
        caps,
        vec![LineCap::Flat, LineCap::Round, LineCap::Square],
        "each theme line's cap must reach the shape that references it"
    );
}

/// A shape's own `a:ln/@cap` overrides the cap it would otherwise inherit from
/// the theme line its `<a:lnRef>` names.
#[test]
fn a_stated_cap_overrides_the_theme_line_style() {
    let theme_xml = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="X">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="X"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="X"><a:lnStyleLst>
      <a:ln w="6350" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="12700" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    </a:lnStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="25400" cap="sq"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], theme_xml);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let FixedElementKind::Shape(ref s) = page.elements[0].kind else {
        panic!("expected a Shape");
    };
    assert_eq!(
        s.stroke.as_ref().expect("expected a stroke").cap,
        LineCap::Square
    );
}
