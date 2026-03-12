use super::*;

#[test]
fn fixed_element_positioned_returns_direct_fields() {
    let elem = FixedElement {
        x: 10.5,
        y: 20.0,
        width: 300.0,
        height: 150.5,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: None,
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    assert!((elem.x() - 10.5).abs() < f64::EPSILON);
    assert!((elem.y() - 20.0).abs() < f64::EPSILON);
    assert!((elem.width() - 300.0).abs() < f64::EPSILON);
    assert!((elem.height() - 150.5).abs() < f64::EPSILON);
}

#[test]
fn floating_image_positioned_maps_offsets_to_xy() {
    let fi = FloatingImage {
        image: ImageData {
            data: vec![],
            format: ImageFormat::Png,
            width: Some(200.0),
            height: Some(100.0),
            crop: None,
        },
        wrap_mode: WrapMode::Square,
        offset_x: 50.0,
        offset_y: 75.0,
    };
    assert!((fi.x() - 50.0).abs() < f64::EPSILON);
    assert!((fi.y() - 75.0).abs() < f64::EPSILON);
    assert!((fi.width() - 200.0).abs() < f64::EPSILON);
    assert!((fi.height() - 100.0).abs() < f64::EPSILON);
}

#[test]
fn floating_image_positioned_returns_zero_when_dimensions_absent() {
    let fi = FloatingImage {
        image: ImageData {
            data: vec![],
            format: ImageFormat::Jpeg,
            width: None,
            height: None,
            crop: None,
        },
        wrap_mode: WrapMode::None,
        offset_x: 10.0,
        offset_y: 20.0,
    };
    assert!((fi.x() - 10.0).abs() < f64::EPSILON);
    assert!((fi.y() - 20.0).abs() < f64::EPSILON);
    assert!(fi.width().abs() < f64::EPSILON);
    assert!(fi.height().abs() < f64::EPSILON);
}

#[test]
fn floating_text_box_positioned_maps_offsets_to_xy() {
    let ftb = FloatingTextBox {
        content: vec![],
        wrap_mode: WrapMode::TopAndBottom,
        width: 250.0,
        height: 180.0,
        offset_x: 30.0,
        offset_y: 45.0,
    };
    assert!((ftb.x() - 30.0).abs() < f64::EPSILON);
    assert!((ftb.y() - 45.0).abs() < f64::EPSILON);
    assert!((ftb.width() - 250.0).abs() < f64::EPSILON);
    assert!((ftb.height() - 180.0).abs() < f64::EPSILON);
}

#[test]
fn positioned_trait_works_through_dyn_dispatch() {
    let elem = FixedElement {
        x: 5.0,
        y: 10.0,
        width: 100.0,
        height: 50.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Ellipse,
            fill: None,
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    let ftb = FloatingTextBox {
        content: vec![],
        wrap_mode: WrapMode::Behind,
        width: 200.0,
        height: 80.0,
        offset_x: 15.0,
        offset_y: 25.0,
    };
    let items: Vec<&dyn Positioned> = vec![&elem, &ftb];
    assert!((items[0].x() - 5.0).abs() < f64::EPSILON);
    assert!((items[0].width() - 100.0).abs() < f64::EPSILON);
    assert!((items[1].x() - 15.0).abs() < f64::EPSILON);
    assert!((items[1].width() - 200.0).abs() < f64::EPSILON);
}

#[test]
fn fixed_element_positioned_with_zero_dimensions() {
    let elem = FixedElement {
        x: 0.0,
        y: 0.0,
        width: 0.0,
        height: 0.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: None,
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    assert!(elem.x().abs() < f64::EPSILON);
    assert!(elem.y().abs() < f64::EPSILON);
    assert!(elem.width().abs() < f64::EPSILON);
    assert!(elem.height().abs() < f64::EPSILON);
}

#[test]
fn floating_text_box_positioned_with_negative_offsets() {
    let ftb = FloatingTextBox {
        content: vec![],
        wrap_mode: WrapMode::InFront,
        width: 100.0,
        height: 60.0,
        offset_x: -10.0,
        offset_y: -5.0,
    };
    assert!((ftb.x() - (-10.0)).abs() < f64::EPSILON);
    assert!((ftb.y() - (-5.0)).abs() < f64::EPSILON);
}
