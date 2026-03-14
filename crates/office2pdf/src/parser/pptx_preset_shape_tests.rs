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
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 7),
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
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 7),
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
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 10),
        other => panic!("Expected Polygon for star5, got {other:?}"),
    }
    assert_eq!(shape.fill, Some(Color::new(255, 215, 0)));
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
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 8),
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
        ShapeKind::Polygon { vertices } => assert_eq!(vertices.len(), 12),
        other => panic!("Expected Polygon for star6, got {other:?}"),
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
