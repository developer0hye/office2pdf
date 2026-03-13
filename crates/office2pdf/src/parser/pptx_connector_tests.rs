use super::*;

// ── Connector shape XML builders ────────────────────────────────────

/// Create a connector shape XML (mirrors real PPTX `<p:cxnSp>` structure).
#[allow(clippy::too_many_arguments)]
fn make_connector(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    border_hex: Option<&str>,
    border_width_emu: Option<i64>,
    dash: Option<&str>,
    flip_h: bool,
    flip_v: bool,
) -> String {
    make_connector_full(x, y, cx, cy, prst, border_hex, border_width_emu, dash, flip_h, flip_v, "", "")
}

/// Create a connector with arrowhead attributes.
#[allow(clippy::too_many_arguments)]
fn make_connector_with_arrows(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    border_hex: Option<&str>,
    border_width_emu: Option<i64>,
    dash: Option<&str>,
    flip_h: bool,
    flip_v: bool,
    tail_type: &str,
) -> String {
    let tail_xml = if tail_type.is_empty() {
        String::new()
    } else {
        format!(r#"<a:tailEnd type="{tail_type}"/>"#)
    };
    make_connector_full(x, y, cx, cy, prst, border_hex, border_width_emu, dash, flip_h, flip_v, "", &tail_xml)
}

#[allow(clippy::too_many_arguments)]
fn make_connector_full(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    border_hex: Option<&str>,
    border_width_emu: Option<i64>,
    dash: Option<&str>,
    flip_h: bool,
    flip_v: bool,
    adj_xml: &str,
    extra_ln_xml: &str,
) -> String {
    let flip_attrs = match (flip_h, flip_v) {
        (true, true) => r#" flipH="1" flipV="1""#,
        (true, false) => r#" flipH="1""#,
        (false, true) => r#" flipV="1""#,
        (false, false) => "",
    };

    let w_attr = border_width_emu
        .map(|w| format!(r#" w="{w}""#))
        .unwrap_or_default();

    let fill_xml = border_hex
        .map(|h| format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#))
        .unwrap_or_default();

    let dash_xml = dash
        .map(|d| format!(r#"<a:prstDash val="{d}"/>"#))
        .unwrap_or_default();

    let av_lst = if adj_xml.is_empty() {
        "<a:avLst/>".to_string()
    } else {
        format!("<a:avLst>{adj_xml}</a:avLst>")
    };

    format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="10" name="Connector"/><p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm{flip_attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}">{av_lst}</a:prstGeom><a:ln{w_attr}>{fill_xml}{dash_xml}{extra_ln_xml}</a:ln></p:spPr></p:cxnSp>"#
    )
}

/// Create a connector with a `<p:style>` section for theme-based line color.
fn make_connector_with_style(
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    prst: &str,
    scheme_color: &str,
    dash: Option<&str>,
    flip_h: bool,
    flip_v: bool,
) -> String {
    let flip_attrs = match (flip_h, flip_v) {
        (true, true) => r#" flipH="1" flipV="1""#,
        (true, false) => r#" flipH="1""#,
        (false, true) => r#" flipV="1""#,
        (false, false) => "",
    };

    let dash_xml = dash
        .map(|d| format!(r#"<a:prstDash val="{d}"/>"#))
        .unwrap_or_default();

    format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="10" name="Connector"/><p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm{flip_attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:ln>{dash_xml}</a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="{scheme_color}"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="{scheme_color}"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="{scheme_color}"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>"#
    )
}

// ── Tests ───────────────────────────────────────────────────────────

#[test]
fn test_straight_connector_parsed_as_line() {
    let connector = make_connector(
        500_000, 1_000_000, 3_000_000, 0,
        "straightConnector1", Some("0F6CFE"), Some(12700), Some("solid"),
        false, false,
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Connector should produce 1 element");

    let elem = &page.elements[0];
    assert!((elem.x - emu_to_pt(500_000)).abs() < 0.1);
    assert!((elem.y - emu_to_pt(1_000_000)).abs() < 0.1);

    let shape = get_shape(elem);
    match &shape.kind {
        ShapeKind::Line { x1, y1, x2, y2, .. } => {
            assert!((*x1).abs() < 0.1, "x1 should be 0");
            assert!((*y1).abs() < 0.1, "y1 should be 0");
            assert!((*x2 - emu_to_pt(3_000_000)).abs() < 0.1);
            assert!((*y2).abs() < 0.1);
        }
        _ => panic!("Expected Line shape, got {:?}", shape.kind),
    }
    let stroke = shape.stroke.as_ref().expect("Expected stroke on connector");
    assert!((stroke.width - 1.0).abs() < 0.1);
    assert_eq!(stroke.color, Color::new(0x0F, 0x6C, 0xFE));
}

#[test]
fn test_connector_with_line_preset() {
    let connector = make_connector(
        0, 0, 5_000_000, 2_000,
        "line", Some("FF0000"), Some(25400), Some("dash"),
        false, false,
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);

    let shape = get_shape(&page.elements[0]);
    assert!(matches!(shape.kind, ShapeKind::Line { .. }));
    let stroke = shape.stroke.as_ref().expect("Expected stroke");
    assert_eq!(stroke.color, Color::new(255, 0, 0));
    assert_eq!(stroke.style, BorderLineStyle::Dashed);
}

#[test]
fn test_connector_flip_h_reverses_line_direction() {
    let connector = make_connector(
        1_000_000, 2_000_000, 4_000_000, 2_000_000,
        "straightConnector1", Some("0000FF"), Some(12700), None,
        true, false,  // flipH only
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let width: f64 = emu_to_pt(4_000_000);
    let height: f64 = emu_to_pt(2_000_000);

    // flipH: start at (width, 0), end at (0, height)
    match &shape.kind {
        ShapeKind::Line { x1, y1, x2, y2, .. } => {
            assert!((*x1 - width).abs() < 0.1, "flipH: x1 should be {width}, got {x1}");
            assert!((*y1).abs() < 0.1, "flipH: y1 should be 0, got {y1}");
            assert!((*x2).abs() < 0.1, "flipH: x2 should be 0, got {x2}");
            assert!((*y2 - height).abs() < 0.1, "flipH: y2 should be {height}, got {y2}");
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_flip_v_reverses_line_direction() {
    let connector = make_connector(
        0, 0, 3_000_000, 2_000_000,
        "straightConnector1", Some("0000FF"), Some(12700), None,
        false, true,  // flipV only
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let width: f64 = emu_to_pt(3_000_000);
    let height: f64 = emu_to_pt(2_000_000);

    // flipV: start at (0, height), end at (width, 0)
    match &shape.kind {
        ShapeKind::Line { x1, y1, x2, y2, .. } => {
            assert!((*x1).abs() < 0.1, "flipV: x1 should be 0, got {x1}");
            assert!((*y1 - height).abs() < 0.1, "flipV: y1 should be {height}, got {y1}");
            assert!((*x2 - width).abs() < 0.1, "flipV: x2 should be {width}, got {x2}");
            assert!((*y2).abs() < 0.1, "flipV: y2 should be 0, got {y2}");
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_flip_h_and_v() {
    let connector = make_connector(
        0, 0, 3_000_000, 2_000_000,
        "straightConnector1", Some("0000FF"), Some(12700), None,
        true, true,  // both flips
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let width: f64 = emu_to_pt(3_000_000);
    let height: f64 = emu_to_pt(2_000_000);

    // flipH+V: start at (width, height), end at (0, 0)
    match &shape.kind {
        ShapeKind::Line { x1, y1, x2, y2, .. } => {
            assert!((*x1 - width).abs() < 0.1, "x1 should be {width}, got {x1}");
            assert!((*y1 - height).abs() < 0.1, "y1 should be {height}, got {y1}");
            assert!((*x2).abs() < 0.1, "x2 should be 0, got {x2}");
            assert!((*y2).abs() < 0.1, "y2 should be 0, got {y2}");
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_mixed_with_regular_shapes() {
    let rect = make_shape(0, 0, 1_000_000, 1_000_000, "rect", Some("FF0000"), None, None);
    let connector = make_connector(
        1_000_000, 500_000, 2_000_000, 0,
        "straightConnector1", Some("0000FF"), Some(12700), None,
        false, false,
    );
    let slide = make_slide_xml(&[rect, connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Should have rect + connector");
    assert!(matches!(get_shape(&page.elements[0]).kind, ShapeKind::Rectangle));
    assert!(matches!(get_shape(&page.elements[1]).kind, ShapeKind::Line { .. }));
}

#[test]
fn test_connector_with_style_based_line_color() {
    let connector = make_connector_with_style(
        0, 0, 3_000_000, 0,
        "straightConnector1", "accent1", Some("dash"),
        false, false,
    );
    let slide = make_slide_xml(&[connector]);
    let theme_xml = make_theme_xml(
        &[
            ("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "44546A"), ("lt2", "E7E6E6"),
            ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"),
            ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47"),
            ("hlink", "0563C1"), ("folHlink", "954F72"),
        ],
        "Calibri",
        "맑은 고딕",
    );
    let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1);
    let shape = get_shape(&page.elements[0]);
    let stroke = shape.stroke.as_ref().expect("Expected stroke from style");
    assert_eq!(stroke.color, Color::new(0x44, 0x72, 0xC4));
    assert_eq!(stroke.style, BorderLineStyle::Dashed);
}

#[test]
fn test_bent_connector3_parsed_as_polyline() {
    let connector = make_connector(
        1_000_000, 2_000_000, 500_000, 300_000,
        "bentConnector3", Some("FF0000"), Some(12700), None,
        false, false,
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "bentConnector3 should produce 1 element");

    let shape = get_shape(&page.elements[0]);
    // Bent connectors are rendered as polylines (Z-shaped paths)
    assert!(
        matches!(shape.kind, ShapeKind::Polyline { .. }),
        "bentConnector3 should be parsed as Polyline, got {:?}",
        shape.kind
    );
    let stroke = shape.stroke.as_ref().expect("Expected stroke");
    assert_eq!(stroke.color, Color::new(255, 0, 0));
}

#[test]
fn test_connector_tail_end_triangle() {
    let connector = make_connector_with_arrows(
        0, 0, 3_000_000, 0,
        "straightConnector1", Some("0000FF"), Some(12700), None,
        false, false,
        "triangle",
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    match &shape.kind {
        ShapeKind::Line { tail_end, head_end, .. } => {
            assert_eq!(*tail_end, ArrowHead::Triangle, "tail should be Triangle");
            assert_eq!(*head_end, ArrowHead::None, "head should be None");
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_bent_connector3_with_adj_value() {
    // bentConnector3 with adj1=74340 (74.34% of width for the bend point)
    let adj_xml = r#"<a:gd name="adj1" fmla="val 74340"/>"#;
    let connector = make_connector_full(
        0, 0, 1_000_000, 500_000,
        "bentConnector3", Some("FF0000"), Some(12700), None,
        false, false,
        adj_xml, "",
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);
    let width: f64 = emu_to_pt(1_000_000);
    let height: f64 = emu_to_pt(500_000);

    match &shape.kind {
        ShapeKind::Polyline { points, .. } => {
            // bentConnector3 with adj=0.7434: start → (mid_x, start_y) → (mid_x, end_y) → end
            assert_eq!(points.len(), 4, "bentConnector3 should have 4 points");
            let mid_x: f64 = width * 0.7434;
            assert!((points[0].0).abs() < 0.1, "start x should be 0");
            assert!((points[1].0 - mid_x).abs() < 0.5, "bend x should be at adj point");
            assert!((points[2].0 - mid_x).abs() < 0.5, "bend x should match");
            assert!((points[3].0 - width).abs() < 0.1, "end x should be width");
            assert!((points[3].1 - height).abs() < 0.1, "end y should be height");
        }
        _ => panic!("Expected Polyline shape, got {:?}", shape.kind),
    }
}
