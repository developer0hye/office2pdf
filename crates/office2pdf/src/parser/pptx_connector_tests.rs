use super::*;

// ── Connector shape XML builders ────────────────────────────────────

/// Create a straight connector shape XML (mirrors real PPTX `<p:cxnSp>` structure).
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

    format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="10" name="Connector"/><p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm{flip_attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:ln{w_attr}>{fill_xml}{dash_xml}</a:ln></p:spPr></p:cxnSp>"#
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
        500_000,
        1_000_000,
        3_000_000,
        0,
        "straightConnector1",
        Some("0F6CFE"),
        Some(12700),
        Some("solid"),
        false,
        false,
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
        ShapeKind::Line { x2, y2 } => {
            assert!((*x2 - emu_to_pt(3_000_000)).abs() < 0.1);
            assert!((*y2 - 0.0).abs() < 0.1);
        }
        _ => panic!("Expected Line shape, got {:?}", shape.kind),
    }
    let stroke = shape.stroke.as_ref().expect("Expected stroke on connector");
    assert!((stroke.width - 1.0).abs() < 0.1);
    assert_eq!(stroke.color, Color::new(0x0F, 0x6C, 0xFE));
}

#[test]
fn test_connector_with_line_preset() {
    // Some connectors use prst="line" instead of "straightConnector1"
    let connector = make_connector(
        0,
        0,
        5_000_000,
        2_000,
        "line",
        Some("FF0000"),
        Some(25400),
        Some("dash"),
        false,
        false,
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
    // flipH means the line goes from right-to-left within the bounding box
    let connector = make_connector(
        1_000_000,
        2_000_000,
        4_000_000,
        2_000_000,
        "straightConnector1",
        Some("0000FF"),
        Some(12700),
        None,
        true,  // flipH
        false,
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);

    // With flipH, line should go from (width, 0) to (0, height)
    match &shape.kind {
        ShapeKind::Line { x2, y2 } => {
            let width = emu_to_pt(4_000_000);
            let height = emu_to_pt(2_000_000);
            // flipH: start at (width, 0), end at (0, height)
            // which means x2 = -width (going left), y2 = height (going down)
            assert!(
                (*x2 - (-width)).abs() < 0.1,
                "flipH: x2 should be -{width}, got {x2}"
            );
            assert!(
                (*y2 - height).abs() < 0.1,
                "flipH: y2 should be {height}, got {y2}"
            );
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_flip_v_reverses_line_direction() {
    let connector = make_connector(
        0,
        0,
        3_000_000,
        2_000_000,
        "straightConnector1",
        Some("0000FF"),
        Some(12700),
        None,
        false,
        true, // flipV
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);

    match &shape.kind {
        ShapeKind::Line { x2, y2 } => {
            let width = emu_to_pt(3_000_000);
            let height = emu_to_pt(2_000_000);
            // flipV: start at (0, height), end at (width, 0)
            // which means x2 = width, y2 = -height
            assert!(
                (*x2 - width).abs() < 0.1,
                "flipV: x2 should be {width}, got {x2}"
            );
            assert!(
                (*y2 - (-height)).abs() < 0.1,
                "flipV: y2 should be -{height}, got {y2}"
            );
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_flip_h_and_v() {
    let connector = make_connector(
        0,
        0,
        3_000_000,
        2_000_000,
        "straightConnector1",
        Some("0000FF"),
        Some(12700),
        None,
        true, // flipH
        true, // flipV
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let shape = get_shape(&page.elements[0]);

    match &shape.kind {
        ShapeKind::Line { x2, y2 } => {
            let width = emu_to_pt(3_000_000);
            let height = emu_to_pt(2_000_000);
            // flipH+flipV: start at (width, height), end at (0, 0)
            // which means x2 = -width, y2 = -height
            assert!(
                (*x2 - (-width)).abs() < 0.1,
                "flipH+V: x2 should be -{width}, got {x2}"
            );
            assert!(
                (*y2 - (-height)).abs() < 0.1,
                "flipH+V: y2 should be -{height}, got {y2}"
            );
        }
        _ => panic!("Expected Line shape"),
    }
}

#[test]
fn test_connector_mixed_with_regular_shapes() {
    let rect = make_shape(
        0,
        0,
        1_000_000,
        1_000_000,
        "rect",
        Some("FF0000"),
        None,
        None,
    );
    let connector = make_connector(
        1_000_000,
        500_000,
        2_000_000,
        0,
        "straightConnector1",
        Some("0000FF"),
        Some(12700),
        None,
        false,
        false,
    );
    let slide = make_slide_xml(&[rect, connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Should have rect + connector");

    // First element: rectangle
    let s0 = get_shape(&page.elements[0]);
    assert!(matches!(s0.kind, ShapeKind::Rectangle));

    // Second element: connector line
    let s1 = get_shape(&page.elements[1]);
    assert!(matches!(s1.kind, ShapeKind::Line { .. }));
}

#[test]
fn test_connector_with_style_based_line_color() {
    // Connectors often inherit line color from <p:style><a:lnRef> when
    // <a:ln> has no explicit <a:solidFill>.
    let connector = make_connector_with_style(
        0,
        0,
        3_000_000,
        0,
        "straightConnector1",
        "accent1",
        Some("dash"),
        false,
        false,
    );
    let slide = make_slide_xml(&[connector]);
    // Use theme builder with known accent1 color
    let theme_xml = make_theme_xml(
        &[
            ("dk1", "000000"),
            ("lt1", "FFFFFF"),
            ("dk2", "44546A"),
            ("lt2", "E7E6E6"),
            ("accent1", "4472C4"),
            ("accent2", "ED7D31"),
            ("accent3", "A5A5A5"),
            ("accent4", "FFC000"),
            ("accent5", "5B9BD5"),
            ("accent6", "70AD47"),
            ("hlink", "0563C1"),
            ("folHlink", "954F72"),
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
    // Should get accent1 = #4472C4 as line color from style
    let stroke = shape.stroke.as_ref().expect("Expected stroke from style");
    assert_eq!(
        stroke.color,
        Color::new(0x44, 0x72, 0xC4),
        "Line color should come from accent1 theme color"
    );
    assert_eq!(stroke.style, BorderLineStyle::Dashed);
}

#[test]
fn test_bent_connector_parsed_as_line() {
    // bentConnector3 should render as a straight line approximation for now
    let connector = make_connector(
        1_000_000,
        2_000_000,
        500_000,
        300_000,
        "bentConnector3",
        Some("FF0000"),
        Some(12700),
        None,
        false,
        false,
    );
    let slide = make_slide_xml(&[connector]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "bentConnector3 should produce 1 element");

    let shape = get_shape(&page.elements[0]);
    // Bent connectors are rendered as straight lines (approximation)
    assert!(
        matches!(shape.kind, ShapeKind::Line { .. }),
        "bentConnector3 should be parsed as Line shape"
    );
    let stroke = shape.stroke.as_ref().expect("Expected stroke");
    assert_eq!(stroke.color, Color::new(255, 0, 0));
}
