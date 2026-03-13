use super::*;

#[test]
fn test_fixed_page_sets_page_size() {
    let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("width: 960pt"),
        "Expected slide width in: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 540pt"),
        "Expected slide height in: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_zero_margins() {
    let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("margin: 0pt"),
        "Expected zero margins for slide in: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_rectangle_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            150.0,
            ShapeKind::Rectangle,
            Some(Color::new(255, 0, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rect"));
    assert!(output.source.contains("200pt"));
    assert!(output.source.contains("rgb(255, 0, 0)"));
}

#[test]
fn test_fixed_page_ellipse_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            50.0,
            50.0,
            120.0,
            80.0,
            ShapeKind::Ellipse,
            Some(Color::new(0, 128, 255)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("ellipse"));
}

#[test]
fn test_fixed_page_line_shape() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            300.0,
            0.0,
            ShapeKind::Line { x1: 0.0, y1: 0.0, x2: 300.0, y2: 0.0, head_end: ArrowHead::None, tail_end: ArrowHead::None },
            None,
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("line"));
}

#[test]
fn test_fixed_page_shape_with_stroke() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Rectangle,
            None,
            Some(BorderSide {
                width: 1.5,
                color: Color::new(0, 0, 255),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("stroke"));
    assert!(output.source.contains("1.5pt"));
}

#[test]
fn test_shape_rotation_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(255, 0, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: Some(90.0),
                opacity: None,
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rotate"));
    assert!(output.source.contains("90deg"));
}

#[test]
fn test_shape_opacity_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(0, 255, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: Some(0.5),
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rgb(0, 255, 0, 128)"));
}

#[test]
fn test_shape_rotation_and_opacity_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 50.0,
            y: 50.0,
            width: 100.0,
            height: 100.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Ellipse,
                fill: Some(Color::new(0, 0, 255)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: Some(45.0),
                opacity: Some(0.75),
                shadow: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rotate"));
    assert!(output.source.contains("45deg"));
    assert!(output.source.contains("rgb(0, 0, 255, 191)"));
}

#[test]
fn test_fixed_page_image_element() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_fixed_image(
            100.0,
            150.0,
            400.0,
            300.0,
            ImageFormat::Png,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#image("));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_fixed_page_mixed_elements() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(50.0, 30.0, 800.0, 60.0, "Title"),
            make_shape_element(
                50.0,
                100.0,
                400.0,
                300.0,
                ShapeKind::Rectangle,
                Some(Color::new(200, 200, 200)),
                None,
            ),
            make_fixed_image(500.0, 100.0, 350.0, 300.0, ImageFormat::Jpeg),
            make_text_box(50.0, 420.0, 800.0, 40.0, "Footer text"),
        ],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Title"));
    assert!(output.source.contains("rect"));
    assert!(output.source.contains("#image("));
    assert!(output.source.contains("Footer text"));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_fixed_page_multiple_text_boxes() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(100.0, 50.0, 300.0, 40.0, "First"),
            make_text_box(100.0, 120.0, 300.0, 40.0, "Second"),
            make_text_box(100.0, 190.0, 300.0, 40.0, "Third"),
        ],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("First"));
    assert!(output.source.contains("Second"));
    assert!(output.source.contains("Third"));
}
