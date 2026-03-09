use super::*;

// ── Floating image codegen tests ──

#[test]
fn test_floating_image_square_wrap_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(200.0),
                    height: Some(100.0),
                    crop: None,
                },
                wrap_mode: WrapMode::Square,
                offset_x: 72.0,
                offset_y: 36.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#place("),
        "Expected #place() for floating image, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("float: true"),
        "Expected float: true for square wrap, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("dx: 72pt"),
        "Expected dx: 72pt, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_image_top_and_bottom_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(150.0),
                    height: Some(75.0),
                    crop: None,
                },
                wrap_mode: WrapMode::TopAndBottom,
                offset_x: 10.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#block("),
        "Expected #block() for topAndBottom wrap, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("#v(75pt)"),
        "Expected vertical space for image height, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_image_behind_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(100.0),
                    height: Some(50.0),
                    crop: None,
                },
                wrap_mode: WrapMode::Behind,
                offset_x: 0.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#place("),
        "Expected #place() for behind wrap, got:\n{}",
        output.source
    );
    assert!(
        !output.source.contains("float: true"),
        "Behind wrap should NOT use float, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_text_box_square_wrap_codegen() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Anchored box")],
            wrap_mode: WrapMode::Square,
            width: 200.0,
            height: 100.0,
            offset_x: 72.0,
            offset_y: 36.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#place("),
        "Expected #place() for floating text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("float: true"),
        "Expected float: true for square-wrapped text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("dx: 72pt"),
        "Expected dx: 72pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("width: 200pt"),
        "Expected width: 200pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("height: 100pt"),
        "Expected height: 100pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Anchored box"),
        "Expected text box content, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_text_box_top_and_bottom_codegen() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Top box")],
            wrap_mode: WrapMode::TopAndBottom,
            width: 150.0,
            height: 60.0,
            offset_x: 10.0,
            offset_y: 0.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#block(width: 100%)"),
        "Expected block wrapper for top-and-bottom text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("#v(60pt)"),
        "Expected reserved vertical space for text box height, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Top box"),
        "Expected text box content, got:\n{}",
        output.source
    );
}

// ── Math equation codegen tests ──

#[test]
fn test_codegen_display_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "frac(a, b)".to_string(),
            display: true,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$ frac(a, b) $"),
        "Expected display math '$ frac(a, b) $', got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_inline_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "x^2".to_string(),
            display: false,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$x^2$"),
        "Expected inline math '$x^2$', got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_complex_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "sum_(i=1)^n i".to_string(),
            display: true,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$ sum_(i=1)^n i $"),
        "Expected display math with sum, got:\n{}",
        output.source
    );
}

// ── Gradient codegen tests (US-050) ─────────────────────────────────

#[test]
fn test_gradient_background_codegen() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: Some(Color::new(255, 0, 0)),
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(255, 0, 0), 0%)"),
        "Should contain first stop. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(0, 0, 255), 100%)"),
        "Should contain second stop. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("angle: 90deg"),
        "Should contain angle. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_background_no_angle_codegen() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 255, 255),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 0),
                },
            ],
            angle: 0.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("angle:"),
        "Should not contain angle for 0 degrees. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_shape_fill_codegen() {
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(0, 128, 0),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 128),
                    },
                ],
                angle: 45.0,
            }),
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear for shape. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(0, 128, 0), 0%)"),
        "Should contain first stop. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("fill: rgb(255, 0, 0)"),
        "Should not contain fallback solid fill. Got: {}",
        output.source,
    );
}

// ── Shadow codegen tests ──────────────────────────────────────────

#[test]
fn test_shape_shadow_codegen() {
    use crate::ir::Shadow;

    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 4.0,
                distance: 3.0,
                direction: 45.0,
                color: Color::new(0, 0, 0),
                opacity: 0.5,
            }),
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("rgb(0, 0, 0, 128)"),
        "Shadow should use rgb with alpha. Got: {}",
        output.source,
    );
    let shadow_pos = output.source.find("rgb(0, 0, 0, 128)");
    let main_pos = output.source.find("rgb(255, 0, 0)");
    assert!(
        shadow_pos < main_pos,
        "Shadow should appear before main shape in output",
    );
}

#[test]
fn test_shape_no_shadow_no_extra_output() {
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("rgb(0, 0, 0,"),
        "No shadow should produce no rgb shadow. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_prefers_over_solid_fill() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: Some(Color::new(128, 128, 128)),
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 180.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Gradient should be preferred. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("fill: rgb(128, 128, 128)"),
        "Solid fallback should not appear. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_unsorted_stops_rendered_in_sorted_order() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
                GradientStop {
                    position: 0.5,
                    color: Color::new(0, 255, 0),
                },
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    let src = &output.source;
    let pos_red = src.find("(rgb(255, 0, 0), 0%)").expect("red stop missing");
    let pos_green = src
        .find("(rgb(0, 255, 0), 50%)")
        .expect("green stop missing");
    let pos_blue = src
        .find("(rgb(0, 0, 255), 100%)")
        .expect("blue stop missing");
    assert!(
        pos_red < pos_green && pos_green < pos_blue,
        "Stops should be in sorted order (0% < 50% < 100%). Got: {}",
        src,
    );
}
