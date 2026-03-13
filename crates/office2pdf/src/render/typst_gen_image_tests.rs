use super::*;
use crate::ir::ImageCrop;

/// Minimal valid 1x1 red pixel PNG for testing.
const MINIMAL_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn make_quadrant_png() -> Vec<u8> {
    let mut image = image::RgbaImage::new(2, 2);
    image.put_pixel(0, 0, image::Rgba([255, 0, 0, 255]));
    image.put_pixel(1, 0, image::Rgba([0, 255, 0, 255]));
    image.put_pixel(0, 1, image::Rgba([0, 0, 255, 255]));
    image.put_pixel(1, 1, image::Rgba([255, 255, 0, 255]));

    let mut encoded = Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(image)
        .write_to(&mut encoded, RasterImageFormat::Png)
        .unwrap();
    encoded.into_inner()
}

fn make_image(format: ImageFormat, width: Option<f64>, height: Option<f64>) -> Block {
    Block::Image(ImageData {
        data: MINIMAL_PNG.to_vec(),
        format,
        width,
        height,
        crop: None,
        stroke: None,
    })
}

#[test]
fn test_image_basic_no_size() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#image(\"img-0.png\")"),
        "Expected #image(\"img-0.png\") in: {}",
        output.source
    );
}

#[test]
fn test_image_crop_preprocesses_raster_asset() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        data: make_quadrant_png(),
        format: ImageFormat::Png,
        width: Some(20.0),
        height: Some(20.0),
        crop: Some(ImageCrop {
            left: 0.5,
            top: 0.0,
            right: 0.0,
            bottom: 0.0,
        }),
        stroke: None,
    })])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 20pt, height: 20pt, fit: \"stretch\")"),
        "Expected original display size in: {}",
        output.source
    );

    let cropped =
        image::load_from_memory_with_format(&output.images[0].data, RasterImageFormat::Png)
            .unwrap()
            .to_rgba8();
    assert_eq!(cropped.dimensions(), (1, 2));
    assert_eq!(cropped.get_pixel(0, 0).0, [0, 255, 0, 255]);
    assert_eq!(cropped.get_pixel(0, 1).0, [255, 255, 0, 255]);
}

#[test]
fn test_image_with_width_only() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(100.0),
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 100pt)"),
        "Expected width param in: {}",
        output.source
    );
}

#[test]
fn test_image_with_height_only() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        Some(80.0),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", height: 80pt)"),
        "Expected height param in: {}",
        output.source
    );
}

#[test]
fn test_image_with_both_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(200.0),
        Some(150.0),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 200pt, height: 150pt, fit: \"stretch\")"),
        "Expected both dimensions with fit stretch in: {}",
        output.source
    );
}

#[test]
fn test_image_collects_asset() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images.len(), 1);
    assert_eq!(output.images[0].path, "img-0.png");
    assert_eq!(output.images[0].data, MINIMAL_PNG);
}

#[test]
fn test_multiple_images_numbered_sequentially() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_image(ImageFormat::Png, None, None),
        make_image(ImageFormat::Jpeg, Some(50.0), None),
    ])]);
    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images.len(), 2);
    assert_eq!(output.images[0].path, "img-0.png");
    assert_eq!(output.images[1].path, "img-1.jpeg");
    assert!(output.source.contains("img-0.png"));
    assert!(output.source.contains("img-1.jpeg"));
}

#[test]
fn test_image_format_extensions() {
    let formats = [
        (ImageFormat::Png, "png"),
        (ImageFormat::Jpeg, "jpeg"),
        (ImageFormat::Gif, "gif"),
        (ImageFormat::Bmp, "bmp"),
        (ImageFormat::Tiff, "tiff"),
        (ImageFormat::Svg, "svg"),
    ];
    for (i, (format, expected_ext)) in formats.iter().enumerate() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(*format, None, None)])]);
        let output = generate_typst(&doc).unwrap();
        let expected_path = format!("img-0.{expected_ext}");
        assert_eq!(
            output.images[0].path, expected_path,
            "Format {format:?} should produce .{expected_ext} extension (test #{i})"
        );
    }
}

#[test]
fn test_image_with_fractional_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(72.5),
        Some(96.25),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("width: 72.5pt"),
        "Expected fractional width in: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 96.25pt"),
        "Expected fractional height in: {}",
        output.source
    );
}

#[test]
fn test_image_mixed_with_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Before image"),
        make_image(ImageFormat::Png, Some(100.0), Some(80.0)),
        make_paragraph("After image"),
    ])]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Before image"));
    assert!(output.source.contains("#image(\"img-0.png\""));
    assert!(output.source.contains("After image"));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_no_images_produces_empty_assets() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Just text")])]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.images.is_empty());
}

#[test]
fn test_image_with_border_renders_box_stroke() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        data: MINIMAL_PNG.to_vec(),
        format: ImageFormat::Png,
        width: Some(127.0),
        height: Some(227.0),
        crop: None,
        stroke: Some(BorderSide {
            width: 6.0,
            color: Color { r: 152, g: 0, b: 0 },
            style: BorderLineStyle::Solid,
        }),
    })])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#box(stroke: "),
        "Expected #box(stroke: ...) wrapper in: {}",
        output.source
    );
    assert!(
        output.source.contains("#image(\"img-0.png\""),
        "Expected #image call in: {}",
        output.source
    );
}

#[test]
fn test_fixed_image_with_border_uses_rect_overlay() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 841.6,
            y: 257.1,
            width: 96.9,
            height: 226.2,
            kind: FixedElementKind::Image(ImageData {
                data: MINIMAL_PNG.to_vec(),
                format: ImageFormat::Png,
                width: Some(96.9),
                height: Some(226.2),
                crop: None,
                stroke: Some(BorderSide {
                    width: 5.87,
                    color: Color {
                        r: 0,
                        g: 176,
                        b: 80,
                    },
                    style: BorderLineStyle::Solid,
                }),
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    // The image should be placed without #box wrapper
    assert!(
        !output.source.contains("#box(stroke:"),
        "Fixed-page image should NOT use #box(stroke:) wrapper: {}",
        output.source
    );
    // Should have a separate #rect overlay for the border
    assert!(
        output.source.contains("#rect("),
        "Expected #rect() border overlay in: {}",
        output.source
    );
    // Image should have correct dimensions
    assert!(
        output.source.contains("width: 96.9pt"),
        "Expected width: 96.9pt in: {}",
        output.source
    );
}

#[test]
fn test_image_without_border_no_box() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(100.0),
        Some(80.0),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("#box(stroke:"),
        "Should NOT have #box wrapper when no stroke: {}",
        output.source
    );
}
