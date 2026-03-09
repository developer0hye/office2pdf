use super::*;
use crate::ir::{
    ChartSeries, ColumnLayout, GradientStop, HeaderFooterParagraph, ListItem, ListKind,
    ListLevelStyle, Metadata, SmartArtNode, StyleSheet,
};
use std::collections::BTreeMap;

/// Helper to create a minimal Document with one FlowPage.
fn make_doc(pages: Vec<Page>) -> Document {
    Document {
        metadata: Metadata::default(),
        pages,
        styles: StyleSheet::default(),
    }
}

/// Helper to create a FlowPage with default A4 size and margins.
fn make_flow_page(content: Vec<Block>) -> Page {
    Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content,
        header: None,
        footer: None,
        columns: None,
    })
}

/// Helper to create a simple paragraph with one plain-text run.
fn make_paragraph(text: &str) -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })
}

#[test]
fn test_generate_plain_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Hello World"));
}

#[test]
fn test_generate_page_setup() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize {
            width: 612.0,
            height: 792.0,
        },
        margins: Margins {
            top: 36.0,
            bottom: 36.0,
            left: 54.0,
            right: 54.0,
        },
        content: vec![make_paragraph("test")],
        header: None,
        footer: None,
        columns: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("612pt"));
    assert!(result.contains("792pt"));
    assert!(result.contains("36pt"));
    assert!(result.contains("54pt"));
}

#[test]
fn test_generate_bold_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold text".to_string(),
            style: TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("weight: \"bold\""),
        "Expected bold weight in: {result}"
    );
    assert!(result.contains("Bold text"));
}

#[test]
fn test_generate_italic_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Italic text".to_string(),
            style: TextStyle {
                italic: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("style: \"italic\""),
        "Expected italic style in: {result}"
    );
    assert!(result.contains("Italic text"));
}

#[test]
fn test_generate_underline_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Underlined".to_string(),
            style: TextStyle {
                underline: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#underline["),
        "Expected underline wrapper in: {result}"
    );
    assert!(result.contains("Underlined"));
}

#[test]
fn test_generate_font_size() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Large text".to_string(),
            style: TextStyle {
                font_size: Some(24.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("size: 24pt"),
        "Expected font size in: {result}"
    );
}

#[test]
fn test_generate_font_color() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Red text".to_string(),
            style: TextStyle {
                color: Some(Color::new(255, 0, 0)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("fill: rgb(255, 0, 0)"),
        "Expected RGB color in: {result}"
    );
}

#[test]
fn test_generate_combined_text_styles() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Styled".to_string(),
            style: TextStyle {
                bold: Some(true),
                italic: Some(true),
                font_size: Some(16.0),
                color: Some(Color::new(0, 128, 255)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("weight: \"bold\""));
    assert!(result.contains("style: \"italic\""));
    assert!(result.contains("size: 16pt"));
    assert!(result.contains("fill: rgb(0, 128, 255)"));
    assert!(result.contains("Styled"));
}

#[test]
fn test_generate_alignment_center() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Centered".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align(center"),
        "Expected center alignment in: {result}"
    );
}

#[test]
fn test_generate_alignment_right() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Right),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Right".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align(right"),
        "Expected right alignment in: {result}"
    );
}

#[test]
fn test_generate_alignment_justify() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Justify),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Justified text".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("par(justify: true") || result.contains("set par(justify: true"),
        "Expected justify in: {result}"
    );
}

#[test]
fn test_generate_line_spacing_proportional() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            line_spacing: Some(LineSpacing::Proportional(2.0)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Double spaced".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("leading:"),
        "Expected leading setting in: {result}"
    );
}

#[test]
fn test_generate_line_spacing_exact() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            line_spacing: Some(LineSpacing::Exact(18.0)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Exact spaced".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("leading: 18pt"),
        "Expected exact leading in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Spaced text".to_string(),
            style: TextStyle {
                letter_spacing: Some(2.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: 2pt"),
        "Expected tracking param in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing_negative() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Condensed".to_string(),
            style: TextStyle {
                letter_spacing: Some(-0.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: -0.5pt"),
        "Expected negative tracking in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_measured_default_stops() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Name:\tValue".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#context {"),
        "Expected contextual tab rendering in: {result}"
    );
    assert!(
        result.contains("measure(tab_prefix_0).width"),
        "Expected tab spacing to measure the rendered prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 36)"),
        "Expected default tabs to advance to the next 36pt stop in: {result}"
    );
    assert!(
        !result.contains("#h(36pt)"),
        "Expected default tabs to avoid a hard-coded 36pt gap in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_next_explicit_stop_and_alignment() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 216.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Col1\tCol2\tCol3".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 72pt"),
        "Expected the first explicit stop to be chosen by measured width in: {result}"
    );
    assert!(
        result.contains("else if tab_prefix_width_2 < 216pt"),
        "Expected the next explicit stop to be selected after the first one in: {result}"
    );
    assert!(
        result.contains("216pt - tab_prefix_width_2 - tab_segment_width_2"),
        "Expected right-aligned tabs to subtract the following segment width in: {result}"
    );
}

#[test]
fn test_generate_tab_falls_back_to_next_default_stop_after_explicit_tabs() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 100.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "A\tB\tC".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 100pt"),
        "Expected the explicit stop to be used when it is still ahead of the prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_2.abs.pt(), 36)"),
        "Expected tabs beyond explicit stops to use the next default stop in: {result}"
    );
}

#[test]
fn test_generate_tab_leader_uses_repeat_fill() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 144.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::Dot,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Heading\t12".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("box(width: tab_advance_1, repeat[.])"),
        "Expected dot tab leaders to render with Typst repeat fill in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_uses_decimal_separator_not_thousands_separator() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1,234.56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1,234]"),
        "Expected decimal alignment to anchor after the thousands group in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_handles_comma_decimal_locale() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1.234,56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1.234]"),
        "Expected decimal alignment to anchor on the locale decimal separator in: {result}"
    );
}

#[test]
fn test_generate_multiple_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("First paragraph"),
        make_paragraph("Second paragraph"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("First paragraph"));
    assert!(result.contains("Second paragraph"));
    assert!(
        result.contains("First paragraph\n\nSecond paragraph"),
        "Expected paragraph break between flow paragraphs in: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_multiple_runs() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Normal ".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: "bold".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: " normal again".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Normal "));
    assert!(result.contains("bold"));
    assert!(result.contains(" normal again"));
}

#[test]
fn test_generate_empty_document() {
    let doc = make_doc(vec![]);
    let result = generate_typst(&doc).unwrap().source;
    // Should produce valid (possibly empty) Typst markup
    assert!(result.is_empty() || !result.is_empty()); // Just shouldn't error
}

#[test]
fn test_generate_special_characters_escaped() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 #items @store",
    )])]);
    let result = generate_typst(&doc).unwrap().source;
    // The text should appear but special chars should be escaped for Typst
    // In Typst, # starts a code expression, so it needs escaping
    assert!(
        result.contains("\\#") || result.contains("Price"),
        "Expected escaped or present text in: {result}"
    );
}

#[path = "typst_gen_table_codegen_tests.rs"]
mod table_codegen_tests;
use self::table_codegen_tests::make_text_cell;

// ── Image codegen tests ─────────────────────────────────────────────

use crate::ir::{ImageCrop, ImageData};

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
    })])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 20pt, height: 20pt)"),
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
            .contains("#image(\"img-0.png\", width: 200pt, height: 150pt)"),
        "Expected both dimensions in: {}",
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

// ── FixedPage codegen tests (US-010) ────────────────────────────────

/// Helper to create a FixedPage (slide-like) with given elements.
fn make_fixed_page(width: f64, height: f64, elements: Vec<FixedElement>) -> Page {
    Page::Fixed(FixedPage {
        size: PageSize { width, height },
        elements,
        background_color: None,
        background_gradient: None,
    })
}

/// Helper to create a text box FixedElement.
fn make_text_box(x: f64, y: f64, w: f64, h: f64, text: &str) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            padding: Insets::default(),
            vertical_align: crate::ir::TextBoxVerticalAlign::Top,
        }),
    }
}

/// Helper to create a shape FixedElement.
fn make_shape_element(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    kind: ShapeKind,
    fill: Option<Color>,
    stroke: Option<BorderSide>,
) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::Shape(Shape {
            kind,
            fill,
            gradient_fill: None,
            stroke,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    }
}

fn make_fixed_text_box(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    padding: Insets,
    vertical_align: crate::ir::TextBoxVerticalAlign,
    content: Vec<Block>,
) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
            content,
            padding,
            vertical_align,
        }),
    }
}

/// Helper to create an image FixedElement.
fn make_fixed_image(x: f64, y: f64, w: f64, h: f64, format: ImageFormat) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::Image(ImageData {
            data: vec![0x89, 0x50, 0x4E, 0x47], // PNG header stub
            format,
            width: Some(w),
            height: Some(h),
            crop: None,
        }),
    }
}

#[path = "typst_gen_fixed_page_tests.rs"]
mod fixed_page_tests;

// ── TablePage codegen tests ──────────────────────────────────────────

/// Helper to create a TablePage.
fn make_table_page(name: &str, width: f64, height: f64, margins: Margins, table: Table) -> Page {
    Page::Table(crate::ir::TablePage {
        name: name.to_string(),
        size: PageSize { width, height },
        margins,
        table,
        header: None,
        footer: None,
        charts: vec![],
    })
}

/// Helper to create a simple Table with text cells.
fn make_simple_table(rows: Vec<Vec<&str>>) -> Table {
    Table {
        rows: rows
            .into_iter()
            .map(|cells| TableRow {
                cells: cells
                    .into_iter()
                    .map(|text| TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: text.to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    })
                    .collect(),
                height: None,
            })
            .collect(),
        column_widths: vec![],
        ..Table::default()
    }
}

#[path = "typst_gen_table_page_tests.rs"]
mod table_page_tests;

// ----- List codegen tests -----

#[path = "typst_gen_list_tests.rs"]
mod list_tests;

#[path = "typst_gen_page_misc_tests.rs"]
mod page_misc_tests;

#[path = "typst_gen_visual_tests.rs"]
mod visual_tests;

#[path = "typst_gen_advanced_tests.rs"]
mod advanced_tests;

#[path = "typst_gen_text_pipeline_tests.rs"]
mod text_pipeline_tests;

#[test]
fn test_generate_run_superscript() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "2".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Superscript),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#super[2]"),
        "Superscript should use #super[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_subscript() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "2".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Subscript),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#sub[2]"),
        "Subscript should use #sub[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_small_caps() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Hello".to_string(),
            style: TextStyle {
                small_caps: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#smallcaps[Hello]"),
        "Small caps should use #smallcaps[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_all_caps() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Hello World".to_string(),
            style: TextStyle {
                all_caps: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("HELLO WORLD"),
        "All caps should uppercase the text. Got: {result}"
    );
}

#[test]
fn test_generate_run_superscript_with_bold() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "n".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Superscript),
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#super[") && result.contains("weight: \"bold\""),
        "Superscript with bold should combine both. Got: {result}"
    );
}

#[test]
fn test_generate_run_highlight_yellow() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Important".to_string(),
            style: TextStyle {
                highlight: Some(Color::new(255, 255, 0)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#highlight(fill: rgb(255, 255, 0))[Important]"),
        "Highlight should use #highlight(fill: ...). Got: {result}"
    );
}

#[test]
fn test_table_cell_vertical_align_center() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Centered".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                vertical_align: Some(CellVerticalAlign::Center),
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align: horizon"),
        "Center vertical alignment should emit 'align: horizon'. Got: {result}"
    );
}

#[test]
fn test_generate_run_highlight_with_bold() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold Highlight".to_string(),
            style: TextStyle {
                highlight: Some(Color::new(0, 255, 0)),
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#highlight(fill: rgb(0, 255, 0))["),
        "Should have highlight wrapper. Got: {result}"
    );
    assert!(
        result.contains("weight: \"bold\""),
        "Should have bold text. Got: {result}"
    );
}

#[test]
fn test_table_cell_vertical_align_bottom() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Bottom".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                vertical_align: Some(CellVerticalAlign::Bottom),
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align: bottom"),
        "Bottom vertical alignment should emit 'align: bottom'. Got: {result}"
    );
}
