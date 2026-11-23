use super::*;

#[test]
fn test_shape_triangle() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "triangle",
        Some("FF0000"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 3, "Triangle should have 3 vertices");
            assert!((vertices[0].0 - 0.5).abs() < 0.01);
            assert!(vertices[0].1.abs() < 0.01);
        }
        other => panic!("Expected Polygon for triangle, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_shape_right_triangle() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "rtTriangle",
        Some("00FF00"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 3, "Right triangle should have 3 vertices");
            assert!(vertices[0].0.abs() < 0.01);
            assert!(vertices[0].1.abs() < 0.01);
        }
        other => panic!("Expected Polygon for rtTriangle, got {other:?}"),
    }
}

#[test]
fn test_shape_round_rect() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        1_000_000,
        "roundRect",
        Some("0000FF"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::RoundedRectangle { radius_fraction } => {
            assert!(*radius_fraction > 0.0, "Radius fraction should be positive");
        }
        other => panic!("Expected RoundedRectangle for roundRect, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(0, 0, 255)));
}

#[test]
fn test_shape_diamond() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "diamond",
        Some("FFFF00"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 4),
        other => panic!("Expected Polygon for diamond, got {other:?}"),
    }
}

#[test]
fn test_shape_pentagon() {
    let shape = make_shape(0, 0, 2_000_000, 2_000_000, "pentagon", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 5),
        other => panic!("Expected Polygon for pentagon, got {other:?}"),
    }
}

#[test]
fn test_shape_hexagon() {
    let shape = make_shape(0, 0, 2_000_000, 2_000_000, "hexagon", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 6),
        other => panic!("Expected Polygon for hexagon, got {other:?}"),
    }
}

#[test]
fn test_shape_octagon() {
    let shape = make_shape(0, 0, 2_000_000, 2_000_000, "octagon", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 8),
        other => panic!("Expected Polygon for octagon, got {other:?}"),
    }
}

#[test]
fn test_shape_right_arrow() {
    let shape = make_shape(
        0,
        0,
        3_000_000,
        1_500_000,
        "rightArrow",
        Some("FF8800"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 7);
            let rightmost = vertices
                .iter()
                .map(|vertex| vertex.0)
                .fold(f64::NEG_INFINITY, f64::max);
            assert!((rightmost - 1.0).abs() < 0.01);
        }
        other => panic!("Expected Polygon for rightArrow, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(255, 136, 0)));
}

#[test]
fn test_shape_right_arrow_uses_short_side_for_default_head_length() {
    let shape = make_shape(0, 0, 3_000_000, 1_500_000, "rightArrow", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let ShapeKind::Polygon { vertices } = &shape.kind else {
        panic!("Expected Polygon for rightArrow, got {:?}", shape.kind);
    };

    assert_eq!(vertices.len(), 7);
    assert!((vertices[1].0 - 0.75).abs() < 1e-9, "{vertices:?}");
    assert!((vertices[2].0 - 0.75).abs() < 1e-9, "{vertices:?}");
    assert!((vertices[0].1 - 0.25).abs() < 1e-9, "{vertices:?}");
    assert!((vertices[6].1 - 0.75).abs() < 1e-9, "{vertices:?}");
}

#[test]
fn test_shape_right_arrow_applies_adjustment_guides() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1500000"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst><a:gd name="adj1" fmla="val 25000"/><a:gd name="adj2" fmla="val 25000"/></a:avLst></a:prstGeom></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let ShapeKind::Polygon { vertices } = &shape.kind else {
        panic!("Expected Polygon for rightArrow, got {:?}", shape.kind);
    };

    assert!((vertices[1].0 - 0.875).abs() < 1e-9, "{vertices:?}");
    assert!((vertices[0].1 - 0.375).abs() < 1e-9, "{vertices:?}");
    assert!((vertices[6].1 - 0.625).abs() < 1e-9, "{vertices:?}");
}

#[test]
fn test_shape_left_arrow() {
    let shape = make_shape(0, 0, 3_000_000, 1_500_000, "leftArrow", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 7);
            let leftmost = vertices
                .iter()
                .map(|vertex| vertex.0)
                .fold(f64::INFINITY, f64::min);
            assert!(leftmost.abs() < 0.01);
            assert!((vertices[1].0 - 0.25).abs() < 1e-9, "{vertices:?}");
        }
        other => panic!("Expected Polygon for leftArrow, got {other:?}"),
    }
}

#[test]
fn test_shape_up_arrow() {
    let shape = make_shape(0, 0, 1_500_000, 3_000_000, "upArrow", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 7);
            assert!((vertices[1].1 - 0.25).abs() < 1e-9, "{vertices:?}");
        }
        other => panic!("Expected Polygon for upArrow, got {other:?}"),
    }
}

#[test]
fn test_shape_down_arrow() {
    let shape = make_shape(0, 0, 1_500_000, 3_000_000, "downArrow", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 7);
            assert!((vertices[1].1 - 0.75).abs() < 1e-9, "{vertices:?}");
        }
        other => panic!("Expected Polygon for downArrow, got {other:?}"),
    }
}

#[test]
fn test_shape_star5() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "star5",
        Some("FFD700"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_vertices_approx(
            vertices,
            &[
                (0.000_001_058, 0.381_965_041),
                (0.381_968_119, 0.381_967_729),
                (0.5, 0.0),
                (0.618_031_881, 0.381_967_729),
                (0.999_998_942, 0.381_965_041),
                (0.690_979_596, 0.618_031_392),
                (0.809_016_341, 0.999_997_459),
                (0.5, 0.763_926_759),
                (0.190_983_659, 0.999_997_459),
                (0.309_020_404, 0.618_031_392),
            ],
        ),
        other => panic!("Expected Polygon for star5, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(255, 215, 0)));
}

#[test]
fn test_shape_star5_applies_adjustment_guide() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Star"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="2000000"/></a:xfrm><a:prstGeom prst="star5"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></p:spPr></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let ShapeKind::Polygon { vertices } = &shape.kind else {
        panic!("Expected Polygon for star5, got {:?}", shape.kind);
    };
    assert_vertices_approx(
        vertices,
        &[
            (0.000_001_058, 0.381_965_041),
            (0.345_491_830, 0.329_178_770),
            (0.5, 0.0),
            (0.654_508_170, 0.329_178_770),
            (0.999_998_942, 0.381_965_041),
            (0.749_999_471, 0.638_194_980),
            (0.809_016_341, 0.999_997_459),
            (0.5, 0.829_177_500),
            (0.190_983_659, 0.999_997_459),
            (0.250_000_529, 0.638_194_980),
        ],
    );
}

#[test]
fn test_shape_star4() {
    let shape = make_shape(0, 0, 2_000_000, 2_000_000, "star4", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_vertices_approx(
            vertices,
            &[
                (0.0, 0.5),
                (0.411_611_652, 0.411_611_652),
                (0.5, 0.0),
                (0.588_388_348, 0.411_611_652),
                (1.0, 0.5),
                (0.588_388_348, 0.588_388_348),
                (0.5, 1.0),
                (0.411_611_652, 0.588_388_348),
            ],
        ),
        other => panic!("Expected Polygon for star4, got {other:?}"),
    }
}

#[test]
fn test_shape_star6() {
    let shape = make_shape(0, 0, 2_000_000, 2_000_000, "star6", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => assert_vertices_approx(
            vertices,
            &[
                (0.000_000_233, 0.25),
                (0.333_330_602, 0.249_995_786),
                (0.5, 0.0),
                (0.666_669_398, 0.249_995_786),
                (0.999_999_767, 0.25),
                (0.833_338_796, 0.5),
                (0.999_999_767, 0.75),
                (0.666_669_398, 0.750_004_214),
                (0.5, 1.0),
                (0.333_330_602, 0.750_004_214),
                (0.000_000_233, 0.75),
                (0.166_661_204, 0.5),
            ],
        ),
        other => panic!("Expected Polygon for star6, got {other:?}"),
    }
}

fn assert_vertices_approx(actual: &[(f64, f64)], expected: &[(f64, f64)]) {
    assert_eq!(actual.len(), expected.len(), "{actual:?}");
    for (index, ((actual_x, actual_y), (expected_x, expected_y))) in
        actual.iter().zip(expected).enumerate()
    {
        assert!(
            (actual_x - expected_x).abs() < 1e-8 && (actual_y - expected_y).abs() < 1e-8,
            "vertex {index}: expected ({expected_x}, {expected_y}), got ({actual_x}, {actual_y}); all vertices: {actual:?}"
        );
    }
}

#[test]
fn test_shape_home_plate() {
    // homePlate: pentagon arrow shape (rect with pointed right edge)
    // Wide shape: cx=1980000 (wider than tall), cy=584391
    let shape = make_shape(
        0,
        0,
        1_980_000,
        584_391,
        "homePlate",
        Some("00259A"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 5, "homePlate should have 5 vertices");
            // First vertex is top-left (0, 0)
            assert!(vertices[0].0.abs() < 0.01);
            assert!(vertices[0].1.abs() < 0.01);
            // Last vertex is bottom-left (0, 1)
            assert!(vertices[4].0.abs() < 0.01);
            assert!((vertices[4].1 - 1.0).abs() < 0.01);
            // Middle vertex is the rightmost point at (1.0, 0.5)
            assert!((vertices[2].0 - 1.0).abs() < 0.01);
            assert!((vertices[2].1 - 0.5).abs() < 0.01);
            // Arrow notch vertices should be between 0 and 1 on x
            assert!(vertices[1].0 > 0.5 && vertices[1].0 < 1.0);
            assert!(vertices[3].0 > 0.5 && vertices[3].0 < 1.0);
        }
        other => panic!("Expected Polygon for homePlate, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(0, 37, 154)));
}

#[test]
fn test_shape_home_plate_square() {
    // Square bounding box: the notch should be at x = 0.5
    let shape = make_shape(0, 0, 1_000_000, 1_000_000, "homePlate", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 5);
            // For square with default adj=50000: notch_x = 1.0 - 0.5 = 0.5
            assert!((vertices[1].0 - 0.5).abs() < 0.01);
        }
        other => panic!("Expected Polygon for homePlate square, got {other:?}"),
    }
}

#[test]
fn test_unsupported_preset_falls_back_to_rectangle() {
    let shape = make_shape(
        0,
        0,
        2_000_000,
        2_000_000,
        "cloudCallout",
        Some("AABBCC"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    assert!(matches!(shape.kind, ShapeKind::Rectangle));
}

#[test]
fn test_regular_polygons_fill_their_bounding_box() {
    // Preset geometries span the full shape box (issue #319): an inscribed
    // pentagon leaves ~10% slack at the bottom, printing the shape short.
    for prst in ["pentagon", "hexagon", "octagon"] {
        let kind = prst_to_shape_kind(
            prst,
            100.0,
            100.0,
            false,
            false,
            ArrowHead::None,
            ArrowHead::None,
            &[],
        );
        let ShapeKind::Polygon { vertices } = kind else {
            panic!("{prst} should be a polygon");
        };
        let min_x = vertices.iter().map(|v| v.0).fold(f64::MAX, f64::min);
        let max_x = vertices.iter().map(|v| v.0).fold(f64::MIN, f64::max);
        let min_y = vertices.iter().map(|v| v.1).fold(f64::MAX, f64::min);
        let max_y = vertices.iter().map(|v| v.1).fold(f64::MIN, f64::max);
        assert!(
            min_x.abs() < 1e-9 && (max_x - 1.0).abs() < 1e-9,
            "{prst} x span {min_x}..{max_x}"
        );
        assert!(
            min_y.abs() < 1e-9 && (max_y - 1.0).abs() < 1e-9,
            "{prst} y span {min_y}..{max_y}"
        );
    }
}

#[test]
fn test_shape_chevron() {
    // chevron: arrow band with a pointed right edge and a matching notch cut
    // into the left edge; rendered as a plain rectangle before (issue #358).
    let shape = make_shape(
        0,
        0,
        1_980_000,
        584_391,
        "chevron",
        Some("00259A"),
        None,
        None,
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Polygon { vertices } => {
            assert_eq!(vertices.len(), 6, "chevron should have 6 vertices");
            // Pointed right edge at (1.0, 0.5)
            assert!((vertices[2].0 - 1.0).abs() < 0.01);
            assert!((vertices[2].1 - 0.5).abs() < 0.01);
            // Notch cut into the left edge, between x=0 and the point
            assert!(vertices[5].0 > 0.0 && vertices[5].0 < 0.5);
            assert!((vertices[5].1 - 0.5).abs() < 0.01);
            // Top-left corner starts at x=0
            assert!(vertices[0].0.abs() < 0.01);
        }
        other => panic!("Expected Polygon for chevron, got {other:?}"),
    }
}

#[test]
fn test_shape_round_rect_honors_adj() {
    // roundRect corner radius comes from the adj value (fraction of the
    // short side, in 100k units); it was hardcoded to 0.1 (issue #361).
    let shape = r##"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 3333"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="EEF2FA"/></a:solidFill></p:spPr></p:sp>"##.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::RoundedRectangle { radius_fraction } => {
            assert!(
                (radius_fraction - 0.03333).abs() < 0.001,
                "adj 3333 must give a 3.3% radius, got {radius_fraction}"
            );
        }
        other => panic!("Expected RoundedRectangle, got {other:?}"),
    }
}

#[test]
fn test_shape_round_rect_default_adj() {
    // Without avLst, the OOXML default adj is 16667 (1/6 of the short side).
    let shape = make_shape(0, 0, 2_000_000, 1_000_000, "roundRect", None, None, None);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::RoundedRectangle { radius_fraction } => {
            assert!(
                (radius_fraction - 0.16667).abs() < 0.001,
                "default adj must give a 16.7% radius, got {radius_fraction}"
            );
        }
        other => panic!("Expected RoundedRectangle, got {other:?}"),
    }
}

#[test]
fn pie_preset_uses_adjusted_sector_angles_including_wraparound() {
    let cases = [
        (
            [9_000_000.0, 16_200_000.0],
            (0.172_673_165, 0.877_964_473),
            (0.5, 0.0),
            (0.0, 0.5),
        ),
        (
            [16_200_000.0, 1_800_000.0],
            (0.5, 0.0),
            (0.827_326_835, 0.877_964_473),
            (1.0, 0.5),
        ),
    ];

    for (adjustments, expected_start, expected_end, expected_cardinal) in cases {
        let kind = prst_to_shape_kind(
            "pie",
            200.0,
            100.0,
            false,
            false,
            ArrowHead::None,
            ArrowHead::None,
            &adjustments,
        );
        let ShapeKind::Path { subpaths } = kind else {
            panic!("pie should retain its adjusted elliptical sector outline");
        };
        assert_eq!(subpaths.len(), 1);
        assert!(subpaths[0].closed);

        let vertices = &subpaths[0].vertices;
        let has_vertex = |expected: (f64, f64), tolerance: f64| {
            vertices.iter().any(|actual| {
                (actual.0 - expected.0).abs() < tolerance
                    && (actual.1 - expected.1).abs() < tolerance
            })
        };
        assert!(
            has_vertex(expected_start, 1e-6),
            "missing adjusted arc start"
        );
        assert!(has_vertex(expected_end, 1e-6), "missing adjusted arc end");
        assert!(
            has_vertex(expected_cardinal, 0.01),
            "the arc should sweep through its intervening cardinal point"
        );
        assert!(
            has_vertex((0.5, 0.5), 1e-6),
            "the sector should close at centre"
        );
        assert!(
            vertices.len() > 20,
            "the curved edge should be sampled instead of drawn as a chord"
        );
    }
}

#[test]
fn pie_preset_uses_standard_defaults_and_equal_angles_mean_a_full_sweep() {
    let default_kind = prst_to_shape_kind(
        "pie",
        100.0,
        100.0,
        false,
        false,
        ArrowHead::None,
        ArrowHead::None,
        &[],
    );
    let ShapeKind::Path {
        subpaths: default_subpaths,
    } = default_kind
    else {
        panic!("the default pie should be a sector path");
    };
    let default_vertices = &default_subpaths[0].vertices;
    let default_has = |expected: (f64, f64)| {
        default_vertices.iter().any(|actual| {
            (actual.0 - expected.0).abs() < 1e-6 && (actual.1 - expected.1).abs() < 1e-6
        })
    };
    assert!(default_has((1.0, 0.5)));
    assert!(default_has((0.5, 0.0)));
    assert!(default_has((0.5, 0.5)));

    let full_kind = prst_to_shape_kind(
        "pie",
        100.0,
        100.0,
        false,
        false,
        ArrowHead::None,
        ArrowHead::None,
        &[0.0, 0.0],
    );
    let ShapeKind::Path {
        subpaths: full_subpaths,
    } = full_kind
    else {
        panic!("equal pie angles should retain the standard full sweep");
    };
    let full_vertices = &full_subpaths[0].vertices;
    for cardinal in [(1.0, 0.5), (0.5, 1.0), (0.0, 0.5), (0.5, 0.0)] {
        assert!(
            full_vertices.iter().any(|actual| {
                (actual.0 - cardinal.0).abs() < 1e-6 && (actual.1 - cardinal.1).abs() < 1e-6
            }),
            "full sweep should include cardinal {cardinal:?}"
        );
    }
    assert_eq!(full_vertices.last(), Some(&(0.5, 0.5)));
}

#[test]
fn wedge_round_rect_callout_matches_the_fixture_bottom_pointer() {
    let kind = prst_to_shape_kind(
        "wedgeRoundRectCallout",
        150.0,
        114.0,
        false,
        false,
        ArrowHead::None,
        ArrowHead::None,
        &[41_242.0, 92_245.0, 16_667.0],
    );
    let ShapeKind::Path { subpaths } = kind else {
        panic!("wedgeRoundRectCallout should retain its rounded wedge outline");
    };
    assert_eq!(subpaths.len(), 1);
    assert!(subpaths[0].closed);

    let vertices = &subpaths[0].vertices;
    let has_vertex = |expected: (f64, f64)| {
        vertices.iter().any(|actual| {
            (actual.0 - expected.0).abs() < 1e-6 && (actual.1 - expected.1).abs() < 1e-6
        })
    };

    let radius = 114.0 * 16_667.0 / 100_000.0;
    assert!(
        has_vertex((radius / 150.0, 0.0)),
        "the top-left corner should turn at the adjusted radius"
    );
    assert!(
        has_vertex((136.863 / 150.0, 162.159_3 / 114.0)),
        "adj1/adj2 should place the fixture pointer below and right of centre"
    );
    assert!(
        vertices.len() > 20,
        "the four rounded corners should be sampled instead of drawn as diagonals"
    );
}

#[test]
fn wedge_round_rect_callout_selects_the_top_edge_for_an_upward_pointer() {
    let kind = prst_to_shape_kind(
        "wedgeRoundRectCallout",
        200.0,
        100.0,
        false,
        false,
        ArrowHead::None,
        ArrowHead::None,
        &[70_000.0, -80_000.0, 10_000.0],
    );
    let ShapeKind::Path { subpaths } = kind else {
        panic!("wedgeRoundRectCallout should retain its rounded wedge outline");
    };

    let vertices = &subpaths[0].vertices;
    assert!(
        vertices
            .iter()
            .any(|actual| { (actual.0 - 1.2).abs() < 1e-6 && (actual.1 + 0.3).abs() < 1e-6 })
    );
    assert!(
        vertices
            .iter()
            .filter(|vertex| vertex.1 < -1e-6)
            .all(|vertex| (vertex.0 - 1.2).abs() < 1e-6),
        "only the upward wedge tip should leave the rounded rectangle"
    );
}

#[test]
fn brace_pair_preserves_open_curves_and_adjusted_short_side_radius() {
    for (width, height, adjustment) in [
        (240.0_f64, 300.0_f64, 8333.0_f64),
        (400.0, 120.0, 20000.0),
        (100.0, 200.0, 40000.0),
        (120.0, 80.0, 0.0),
    ] {
        let shape_xml = make_shape(
            0,
            0,
            (width * 12700.0) as i64,
            (height * 12700.0) as i64,
            "bracePair",
            None,
            Some(25400),
            Some("9C003B"),
        )
        .replace(
            "<a:avLst/>",
            &format!(r#"<a:avLst><a:gd name="adj" fmla="val {adjustment}"/></a:avLst>"#),
        );
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[make_slide_xml(&[shape_xml])]);
        let (doc, _) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let shape = get_shape(&first_fixed_page(&doc).elements[0]);
        let ShapeKind::Path { subpaths } = &shape.kind else {
            panic!("brace pair must have curved subpaths")
        };
        assert_eq!(subpaths.len(), 2);
        assert!(subpaths.iter().all(|path| !path.closed));
        let radius = width.min(height) * adjustment.clamp(0.0, 25000.0) / 100000.0;
        let left = &subpaths[0].vertices;
        let right = &subpaths[1].vertices;
        assert!((left[0].0 * width - 2.0 * radius).abs() < 1e-6);
        assert!((left[0].1 - 1.0).abs() < 1e-6);
        assert!((left.last().unwrap().1).abs() < 1e-6);
        assert!(
            left.iter()
                .any(|&(x, y)| x.abs() < 1e-6 && (y - 0.5).abs() < 1e-6)
        );
        assert!(
            right
                .iter()
                .any(|&(x, y)| (x - 1.0).abs() < 1e-6 && (y - 0.5).abs() < 1e-6)
        );
        assert!(
            subpaths
                .iter()
                .flat_map(|p| &p.vertices)
                .all(|&(x, y)| x.is_finite() && y.is_finite())
        );
    }
}

#[test]
fn brace_pair_separates_enclosed_fill_from_open_outline() {
    let shape = make_shape(
        0,
        0,
        3048000,
        3810000,
        "bracePair",
        Some("00FF00"),
        Some(25400),
        Some("9C003B"),
    );
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[make_slide_xml(&[shape])]);
    let (doc, _) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    assert_eq!(
        page.elements.len(),
        2,
        "fill perimeter and open strokes have different topology"
    );
    let fill = get_shape(&page.elements[0]);
    let outline = get_shape(&page.elements[1]);
    assert_eq!(fill.fill, Some(Color::new(0, 255, 0)));
    assert!(fill.stroke.is_none());
    let ShapeKind::Path { subpaths } = &fill.kind else {
        panic!("enclosed fill path")
    };
    assert_eq!(subpaths.len(), 1);
    assert!(subpaths[0].closed);
    assert!(subpaths[0].vertices.iter().any(|p| p.0 < 0.01));
    assert!(subpaths[0].vertices.iter().any(|p| p.0 > 0.99));
    assert_eq!(outline.fill, None);
    assert!(outline.stroke.is_some());
    let ShapeKind::Path { subpaths } = &outline.kind else {
        panic!("open outline path")
    };
    assert!(subpaths.iter().all(|p| !p.closed));
}
