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
fn test_fixed_page_omits_text_fully_occluded_by_later_opaque_jpeg() {
    let mut jpeg = Vec::new();
    image::DynamicImage::ImageRgb8(image::RgbImage::from_pixel(1, 1, image::Rgb([20, 40, 60])))
        .write_to(
            &mut std::io::Cursor::new(&mut jpeg),
            image::ImageFormat::Jpeg,
        )
        .unwrap();
    let mut cover = make_fixed_image(0.0, 0.0, 959.76, 540.0, ImageFormat::Jpeg);
    let FixedElementKind::Image(image) = &mut cover.kind else {
        unreachable!();
    };
    image.data = jpeg;
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(5.0, 525.4, 67.4, 9.6, "Sensitivity: Internal"),
            cover,
            make_text_box(120.0, 120.0, 720.0, 80.0, "Visible slide title"),
        ],
    )]);

    let output = generate_typst(&doc).unwrap();

    assert!(
        !output.source.contains("[Sensitivity:]"),
        "fully covered text must not remain in generated markup:\n{}",
        output.source
    );
    assert!(
        output.source.contains("[Visible]"),
        "text painted after the covering image must remain in generated markup:\n{}",
        output.source
    );

    let pdf =
        crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
            .unwrap();
    let extracted = pdf_extract::extract_text_from_mem(&pdf).unwrap();
    let searchable = extracted.split_whitespace().collect::<Vec<_>>().join(" ");
    assert!(
        !searchable.contains("Sensitivity"),
        "fully covered text must not remain searchable: {searchable:?}"
    );
    assert!(
        searchable.contains("Visible slide title"),
        "text painted after the covering image must remain searchable: {searchable:?}"
    );
}

#[test]
fn test_fixed_page_keeps_text_covered_by_a_non_page_sized_jpeg() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(10.0, 10.0, 60.0, 20.0, "Searchable caption"),
            make_fixed_image(0.0, 0.0, 100.0, 100.0, ImageFormat::Jpeg),
        ],
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("[Searchable]"),
        "partial-slide artwork is outside the conservative occlusion rule:\n{source}"
    );
}

#[test]
fn test_fixed_page_keeps_no_wrap_text_whose_frame_is_covered() {
    let mut overflow = make_text_box(950.0, 20.0, 5.0, 20.0, "Visible overflow");
    let FixedElementKind::TextBox(text_box) = &mut overflow.kind else {
        unreachable!();
    };
    text_box.no_wrap = true;
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            overflow,
            make_fixed_image(0.0, 0.0, 959.76, 540.0, ImageFormat::Jpeg),
        ],
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("[Visible]"),
        "no-wrap glyphs can escape a covered frame and must remain:\n{source}"
    );
}

#[test]
fn test_fixed_page_keeps_text_under_a_transparent_full_page_png() {
    let mut png = Vec::new();
    image::DynamicImage::ImageRgba8(image::RgbaImage::from_pixel(
        1,
        1,
        image::Rgba([0, 0, 0, 0]),
    ))
    .write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
    .unwrap();
    let mut cover = make_fixed_image(0.0, 0.0, 960.0, 540.0, ImageFormat::Png);
    let FixedElementKind::Image(image) = &mut cover.kind else {
        unreachable!();
    };
    image.data = png;
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(10.0, 10.0, 60.0, 20.0, "Visible below alpha"),
            cover,
        ],
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("[Visible]"),
        "transparent page artwork must not remove searchable text:\n{source}"
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
            ShapeKind::Line {
                x1: 0.0,
                y1: 0.0,
                x2: 300.0,
                y2: 0.0,
                head_end: ArrowHead::None,
                tail_end: ArrowHead::None,
            },
            None,
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
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
                join: LineJoin::Round,
                cap: LineCap::Flat,
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
                pattern_fill: None,
                stroke: None,
                rotation_deg: Some(90.0),
                opacity: None,
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("rotate"));
    assert!(output.source.contains("90deg"));
}

#[test]
fn shape_rotation_keeps_the_declared_box_center_as_its_pivot() {
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
                pattern_fill: None,
                stroke: None,
                rotation_deg: Some(90.0),
                opacity: None,
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains(
            "#move(dx: 175pt, dy: -25pt)[#rotate(90deg, origin: top + left, reflow: false)["
        ),
        "shape rotation must pivot around the unclamped 200x150pt box center:\n{}",
        output.source
    );
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
                pattern_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: Some(0.5),
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
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
                pattern_fill: None,
                stroke: None,
                rotation_deg: Some(45.0),
                opacity: Some(0.75),
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
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
    assert!(output.source.contains("#o2p-pptx-word([Footer]"));
    assert!(output.source.contains("#o2p-pptx-word([text]"));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_line_arrowhead_uses_place_overlay() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 0.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Line {
                    x1: 0.0,
                    y1: 0.0,
                    x2: 200.0,
                    y2: 0.0,
                    head_end: ArrowHead::None,
                    tail_end: ArrowHead::Triangle,
                },
                fill: None,
                gradient_fill: None,
                pattern_fill: None,
                stroke: Some(BorderSide {
                    width: 2.0,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                    cap: LineCap::Flat,
                }),
                rotation_deg: None,
                opacity: None,
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    // Arrowhead polygon must be inside #place(top + left) so it overlays
    // on the line rather than stacking below it in the layout.
    assert!(
        output.source.contains("#place(top + left)[#polygon("),
        "Arrowhead polygon must use #place overlay, got: {}",
        output.source,
    );
}

#[test]
fn test_polyline_segments_use_place_overlay() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 100.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Polyline {
                    points: vec![(0.0, 0.0), (100.0, 0.0), (100.0, 100.0), (200.0, 100.0)],
                    head_end: ArrowHead::None,
                    tail_end: ArrowHead::Triangle,
                },
                fill: None,
                gradient_fill: None,
                pattern_fill: None,
                stroke: Some(BorderSide {
                    width: 1.5,
                    color: Color::new(0, 0, 255),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                    cap: LineCap::Flat,
                }),
                rotation_deg: None,
                opacity: None,
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    // Each polyline segment must use #place overlay for correct positioning.
    let segment_count = output.source.matches("#place(top + left)[#line(").count();
    assert!(
        segment_count >= 3,
        "Expected 3 polyline segments with #place overlay, found {}: {}",
        segment_count,
        output.source,
    );
    // Arrowhead must also use #place overlay.
    assert!(
        output.source.contains("#place(top + left)[#polygon("),
        "Arrowhead polygon must use #place overlay, got: {}",
        output.source,
    );
}

#[test]
fn test_rotated_polyline_pre_rotates_points_without_typst_rotate_wrapper() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 120.0,
            height: 160.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Polyline {
                    points: vec![(120.0, 0.0), (20.0, 0.0), (20.0, 160.0), (0.0, 160.0)],
                    head_end: ArrowHead::None,
                    tail_end: ArrowHead::None,
                },
                fill: None,
                gradient_fill: None,
                pattern_fill: None,
                stroke: Some(BorderSide {
                    width: 1.0,
                    color: Color::new(67, 113, 187),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                    cap: LineCap::Flat,
                }),
                rotation_deg: Some(270.0),
                opacity: None,
                stroke_opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();

    assert!(
        !output.source.contains("#rotate(270deg)"),
        "Rotated polylines should emit transformed points directly: {}",
        output.source,
    );
    assert!(
        output.source.contains("start: (-20.000000000000014pt, 20.000000000000014pt), end: (-19.999999999999993pt, 120.00000000000001pt)")
            || output
                .source
                .contains("start: (-20pt, 20pt), end: (-20pt, 120pt)"),
        "Expected rotated first segment coordinates, got: {}",
        output.source,
    );
    assert!(
        output.source.contains("start: (-19.999999999999993pt, 120.00000000000001pt), end: (140pt, 119.99999999999999pt)")
            || output
                .source
                .contains("start: (-20pt, 120pt), end: (140pt, 120pt)"),
        "Expected rotated second segment coordinates, got: {}",
        output.source,
    );
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

/// A stated `a:ln/@cap` has to survive compilation, not merely appear in the
/// generated Typst: `cap: "round"` and `cap: "square"` must reach the drawn
/// stroke as PDF's round and projecting-square caps, which is what extends each
/// end by half the line width. A native macOS PowerPoint export of a 3.5pt
/// straight connector traces `linecap="1,1,1"` for `cap="rnd"` and
/// `linecap="2,2,2"` for `cap="sq"`, against `0,0,0` when the attribute is
/// absent (issue #1682).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_stated_line_cap_reaches_the_compiled_stroke() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};
    use typst::visualize::LineCap as CompiledLineCap;

    for (cap, expected) in [
        (LineCap::Flat, CompiledLineCap::Butt),
        (LineCap::Round, CompiledLineCap::Round),
        (LineCap::Square, CompiledLineCap::Square),
    ] {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                20.0,
                100.0,
                300.0,
                0.0,
                ShapeKind::Line {
                    x1: 0.0,
                    y1: 0.0,
                    x2: 300.0,
                    y2: 0.0,
                    head_end: ArrowHead::None,
                    tail_end: ArrowHead::None,
                },
                None,
                Some(BorderSide {
                    width: 3.5,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                    cap,
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        let painted =
            compiled_paint_sequence(&output.source, &output.images, 0).expect("the slide compiles");
        let stroke = painted
            .iter()
            .filter(|item| item.kind == PaintedKind::Shape)
            .find_map(|item| item.stroke.as_ref())
            .unwrap_or_else(|| panic!("the line is stroked: {painted:?}"));
        assert_eq!(stroke.cap, expected, "IR cap {cap:?}");
    }
}

#[test]
fn inset_page_sized_picture_keeps_text_visible_in_its_unfilled_margin() {
    let mut jpeg = Vec::new();
    image::DynamicImage::ImageRgb8(image::RgbImage::from_pixel(1, 1, image::Rgb([20, 40, 60])))
        .write_to(
            &mut std::io::Cursor::new(&mut jpeg),
            image::ImageFormat::Jpeg,
        )
        .unwrap();
    let mut cover = make_fixed_image(0.0, 0.0, 960.0, 540.0, ImageFormat::Jpeg);
    let FixedElementKind::Image(image) = &mut cover.kind else {
        unreachable!();
    };
    image.data = jpeg;
    image.fill_rect = Some(crate::ir::ImageFillRect {
        left: 0.25,
        ..crate::ir::ImageFillRect::default()
    });
    let document = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![
            make_text_box(10.0, 10.0, 200.0, 30.0, "Visible margin caption"),
            cover,
        ],
    )]);
    let generated = generate_typst(&document).unwrap();
    let pdf = crate::render::pdf::compile_to_pdf(
        &generated.source,
        &generated.images,
        None,
        &[],
        false,
        false,
    )
    .unwrap();
    let text = pdf_extract::extract_text_from_mem(&pdf).unwrap();
    assert!(
        text.split_whitespace()
            .collect::<Vec<_>>()
            .join(" ")
            .contains("Visible margin caption")
    );
}
