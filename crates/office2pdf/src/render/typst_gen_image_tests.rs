use super::*;
use crate::ir::{ImageCrop, ImageParagraphIndent, ImageParagraphSpacing};

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

#[cfg(not(target_arch = "wasm32"))]
fn floating_band_picture(
    wrap_mode: WrapMode,
    width: Option<f64>,
    height: Option<f64>,
    rotation_deg: Option<f64>,
    crop: Option<ImageCrop>,
) -> Document {
    let mut pixels: image::RgbaImage = image::RgbaImage::new(200, 100);
    let bands: [[u8; 4]; 4] = [
        [220, 40, 40, 255],
        [40, 180, 60, 255],
        [40, 80, 220, 255],
        [240, 190, 30, 255],
    ];
    for (x, _, pixel) in pixels.enumerate_pixels_mut() {
        *pixel = image::Rgba(bands[(x / 50) as usize]);
    }
    let mut encoded: Cursor<Vec<u8>> = Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(pixels)
        .write_to(&mut encoded, RasterImageFormat::Png)
        .unwrap();
    make_doc(vec![make_flow_page(vec![Block::FloatingImage(
        FloatingImage {
            image: ImageData {
                rotation_deg,
                flip_h: false,
                flip_v: false,
                data: encoded.into_inner(),
                format: ImageFormat::Png,
                width,
                height,
                crop,
                fill_rect: None,
                stroke: None,
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                shadow: None,
                paragraph_spacing: None,
                paragraph_indent: None,
            },
            wrap_mode,
            vertical_anchor: FloatingImageVerticalAnchor::Margin,
            offset_x: 24.0,
            offset_y: 36.0,
        },
    )])])
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_floating_picture_size(document: &Document, width: f64, height: f64) {
    let output = generate_typst(document).unwrap();
    let boxes = crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
        .expect("the floating picture compiles");
    let [picture] = boxes.as_slice() else {
        panic!("expected one painted picture, got {boxes:?}");
    };
    let edge_length = |first: (f64, f64), second: (f64, f64)| -> f64 {
        (second.0 - first.0).hypot(second.1 - first.1)
    };
    let actual_width: f64 = edge_length(picture.corners[0], picture.corners[1]);
    let actual_height: f64 = edge_length(picture.corners[1], picture.corners[2]);
    // Read the source image's painted extent, not its parent's clipping box:
    // cover can produce a correctly sized frame while cutting off its artwork.
    assert!(
        (actual_width - width).abs() < 0.01 && (actual_height - height).abs() < 0.01,
        "the full picture must occupy {width} x {height}pt, got {actual_width} x {actual_height}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_pictures_preserve_declared_dimensions_for_every_wrap_mode() {
    for wrap_mode in [
        WrapMode::None,
        WrapMode::Behind,
        WrapMode::InFront,
        WrapMode::Square,
        WrapMode::Tight,
        WrapMode::TopAndBottom,
    ] {
        for (width, height) in [(120.0, 120.0), (180.0, 45.0)] {
            let document = floating_band_picture(wrap_mode, Some(width), Some(height), None, None);
            assert_floating_picture_size(&document, width, height);
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_pictures_preserve_declared_dimensions_after_crop_and_rotation() {
    for rotation in [0.0, 27.0, 90.0] {
        let document = floating_band_picture(
            WrapMode::InFront,
            Some(135.0),
            Some(75.0),
            Some(rotation),
            Some(ImageCrop {
                left: 0.25,
                top: 0.0,
                right: 0.25,
                bottom: 0.0,
            }),
        );
        assert_floating_picture_size(&document, 135.0, 75.0);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_vector_pictures_preserve_declared_dimensions() {
    let mut document =
        floating_band_picture(WrapMode::InFront, Some(120.0), Some(120.0), None, None);
    let Page::Flow(page) = &mut document.pages[0] else {
        unreachable!();
    };
    let Block::FloatingImage(picture) = &mut page.content[0] else {
        unreachable!();
    };
    picture.image.format = ImageFormat::Svg;
    picture.image.data = br##"<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100"><rect width="200" height="100" fill="#dc2828"/><rect x="50" width="50" height="100" fill="#28b43c"/><rect x="100" width="50" height="100" fill="#2850dc"/><rect x="150" width="50" height="100" fill="#f0be1e"/></svg>"##.to_vec();
    assert_floating_picture_size(&document, 120.0, 120.0);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_pictures_with_one_dimension_keep_the_source_aspect_ratio() {
    for (width, height, expected_width, expected_height) in [
        (Some(120.0), None, 120.0, 60.0),
        (None, Some(40.0), 80.0, 40.0),
    ] {
        let document = floating_band_picture(WrapMode::InFront, width, height, None, None);
        assert_floating_picture_size(&document, expected_width, expected_height);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_pictures_in_an_authored_docx_preserve_their_frame_dimensions() {
    use crate::parser::{Parser, docx::DocxParser};

    let source: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
    let (document, _) = DocxParser
        .parse(source, &crate::config::ConvertOptions::default())
        .expect("the authored picture document parses");
    let output = generate_typst(&document).unwrap();
    let pictures = crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
        .expect("the authored picture document compiles");
    assert_eq!(pictures.len(), 4);
    for (picture, (width, height)) in
        pictures
            .iter()
            .zip([(120.0, 120.0), (180.0, 45.0), (135.0, 75.0), (135.0, 75.0)])
    {
        let edge_length = |first: (f64, f64), second: (f64, f64)| -> f64 {
            (second.0 - first.0).hypot(second.1 - first.1)
        };
        let actual_width = edge_length(picture.corners[0], picture.corners[1]);
        let actual_height = edge_length(picture.corners[1], picture.corners[2]);
        assert!(
            (actual_width - width).abs() < 0.01 && (actual_height - height).abs() < 0.01,
            "the parsed picture must paint {width} x {height}pt, got {actual_width} x {actual_height}pt"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_picture_source_crops_follow_each_occurrence_of_shared_media() {
    use crate::parser::{Parser, docx::DocxParser};

    let source: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
    let (document, _) = DocxParser
        .parse(source, &crate::config::ConvertOptions::default())
        .expect("the authored picture document parses");
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected a flowing Word page");
    };
    let pictures: Vec<&FloatingImage> = page
        .content
        .iter()
        .filter_map(|block| match block {
            Block::FloatingImage(picture) => Some(picture),
            _ => None,
        })
        .collect();
    assert_eq!(pictures.len(), 4);
    for (index, picture) in pictures.iter().enumerate() {
        let expected: Option<ImageCrop> = (index == 2).then_some(ImageCrop {
            left: 0.25,
            top: 0.0,
            right: 0.25,
            bottom: 0.0,
        });
        assert_eq!(
            picture.image.crop, expected,
            "source selection on picture {index}"
        );
    }

    let output = generate_typst(&document).unwrap();
    let decoded: Vec<image::RgbImage> = output
        .images
        .iter()
        .map(|asset| image::load_from_memory(&asset.data).unwrap().to_rgb8())
        .collect();
    assert!(
        decoded.iter().any(|pixels| {
            pixels.dimensions() == (100, 100)
                && pixels.get_pixel(0, 50).0 == [40, 180, 60]
                && pixels.get_pixel(99, 50).0 == [40, 80, 220]
        }),
        "the rendered cropped picture must contain only the selected green/blue half"
    );
    assert!(
        decoded.iter().any(|pixels| {
            pixels.dimensions() == (200, 100)
                && pixels.get_pixel(0, 50).0 == [220, 40, 40]
                && pixels.get_pixel(199, 50).0 == [240, 190, 30]
        }),
        "uncropped uses of the same media must retain all four bands"
    );
}

/// The left edges of the pictures the compiled document paints on its first
/// page, in layout order.
#[cfg(not(target_arch = "wasm32"))]
fn compiled_picture_left_edges_pt(document: &Document) -> Vec<f64> {
    let output = generate_typst(document).unwrap();
    crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
        .expect("the picture document compiles")
        .iter()
        .map(|picture| picture.corners[0].0)
        .collect()
}

/// docxide-pdf's `inline-picture-indent-*` cases, unmodified: one paragraph
/// holding only a 180x90pt inline picture, under eight combinations of `w:jc`
/// and `w:ind`, with 72pt side margins. Word places the picture where it
/// places a first line of text: that line runs from `left + firstLine` (a
/// hanging indent is a negative first line) to `right`, and the alignment
/// works inside it. Main drew every picture as if the paragraph had no
/// indent (issue #2201).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_inline_picture_paragraph_places_its_picture_by_the_paragraph_indent() {
    use crate::parser::{Parser, docx::DocxParser};

    let cases: [(&str, &[u8], f64); 8] = [
        (
            "left, firstLine 709",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_left_first.docx"
            ),
            107.45,
        ),
        (
            "both, firstLine 709",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_justify_first.docx"
            ),
            107.45,
        ),
        (
            "left, left 720 hanging 360",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_left_hanging.docx"
            ),
            90.0,
        ),
        (
            "center, firstLine 709",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_center_first.docx"
            ),
            225.38,
        ),
        (
            "center, left 720 hanging 360",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_center_hanging.docx"
            ),
            216.65,
        ),
        (
            "center, left 720 right 360",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_center_left_right.docx"
            ),
            216.65,
        ),
        (
            "center, no indent",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_center_zero.docx"
            ),
            207.65,
        ),
        (
            "right, firstLine 709",
            include_bytes!(
                "../../../../tests/fixtures/docx/docxide/inline_picture_indent_right_first.docx"
            ),
            343.3,
        ),
    ];
    for (case, source, word_left_pt) in cases {
        let (document, _) = DocxParser
            .parse(source, &crate::config::ConvertOptions::default())
            .expect("the picture document parses");
        let left_edges: Vec<f64> = compiled_picture_left_edges_pt(&document);
        assert!(
            left_edges.len() == 1 && (left_edges[0] - word_left_pt).abs() < 0.05,
            "{case}: Word puts the picture's left edge at {word_left_pt}pt, got {left_edges:?}"
        );
    }
}

/// The general rule behind
/// [`an_inline_picture_paragraph_places_its_picture_by_the_paragraph_indent`]:
/// pictures sit where the paragraph's lines of text would (issue #2201).
/// Several pictures in one paragraph wrap like words, so only the first line
/// takes the first-line indent and a wrapped picture starts at `left`. A
/// negative indent, or a hanging first line past a zero `left`, widens the
/// line into the margin, as Word does.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn inline_pictures_start_where_the_paragraphs_text_lines_would() {
    let picture = |width: f64, alignment: Option<Alignment>, indent: ImageParagraphIndent| {
        let Block::Image(mut image) = make_image(ImageFormat::Png, Some(width), Some(40.0)) else {
            unreachable!();
        };
        image.data = make_quadrant_png();
        image.alignment = alignment;
        image.paragraph_indent = Some(indent);
        image
    };
    let hanging = ImageParagraphIndent {
        left: 36.0,
        right: 0.0,
        first_line: -18.0,
    };
    // A4 with 72pt margins: the text column runs from 72pt to 523.28pt.
    let cases: [(&str, Block, &[f64]); 5] = [
        (
            "left 0, hanging 18",
            Block::Image(picture(
                100.0,
                None,
                ImageParagraphIndent {
                    left: 0.0,
                    right: 0.0,
                    first_line: -18.0,
                },
            )),
            &[54.0],
        ),
        (
            "right-aligned, right 18, firstLine 36",
            Block::Image(picture(
                100.0,
                Some(Alignment::Right),
                ImageParagraphIndent {
                    left: 0.0,
                    right: 18.0,
                    first_line: 36.0,
                },
            )),
            &[405.28],
        ),
        (
            "left -36",
            Block::Image(picture(
                100.0,
                None,
                ImageParagraphIndent {
                    left: -36.0,
                    right: 0.0,
                    first_line: 0.0,
                },
            )),
            &[36.0],
        ),
        (
            "right-aligned, right -18",
            Block::Image(picture(
                100.0,
                Some(Alignment::Right),
                ImageParagraphIndent {
                    left: 0.0,
                    right: -18.0,
                    first_line: 0.0,
                },
            )),
            &[441.28],
        ),
        (
            "two pictures too wide to share a line, left 36, hanging 18",
            Block::InlineImages(vec![
                picture(250.0, None, hanging),
                picture(250.0, None, hanging),
            ]),
            &[90.0, 108.0],
        ),
    ];
    for (case, block, expected_left_edges) in cases {
        let left_edges: Vec<f64> =
            compiled_picture_left_edges_pt(&make_doc(vec![make_flow_page(vec![block])]));
        assert!(
            left_edges.len() == expected_left_edges.len()
                && left_edges
                    .iter()
                    .zip(expected_left_edges)
                    .all(|(left, expected)| (left - expected).abs() < 0.05),
            "{case}: the pictures must start at {expected_left_edges:?}pt, got {left_edges:?}"
        );
    }
}

/// An inline picture between two text runs is a glyph-like object on their
/// line, as Word lays it out (issue #2227). Its bottom edge sits on the
/// baseline both runs share and the run after it starts at its right edge. A
/// picture taller than the text grows the line upward only: the line's descent
/// and leading stay the text's, so the next line follows as far below the
/// baseline as it does under a short picture, which in the body is one
/// ordinary line pitch. Main drew the picture as its own block above the
/// paragraph and set both runs on one line beneath it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_inline_picture_between_text_runs_stands_on_their_baseline() {
    use crate::parser::{Parser, docx::DocxParser};

    let line = |text: &str| docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(text));
    for place in ["body", "table cell"] {
        // The gap from the picture line's baseline down to the next one, for a
        // 72pt-wide picture 36pt tall (taller than the text) and 4pt tall.
        let mut gaps_below_baseline: Vec<f64> = Vec::new();
        for height_pt in [36.0, 4.0] {
            let case = format!("{height_pt}pt picture in the {place}");
            let picture_paragraph = docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Left "))
                .add_run(docx_rs::Run::new().add_image(
                    docx_rs::Pic::new(MINIMAL_PNG).size(914_400, (height_pt * 12_700.0) as u32),
                ))
                .add_run(docx_rs::Run::new().add_text(" End"));
            let docx = docx_rs::Docx::new().add_paragraph(line("First line"));
            let docx = match place {
                "body" => docx.add_paragraph(picture_paragraph),
                _ => docx.add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
                    docx_rs::TableCell::new().add_paragraph(picture_paragraph),
                ])])),
            }
            .add_paragraph(line("Next line"))
            .add_paragraph(line("Last line"));
            let mut cursor = Cursor::new(Vec::new());
            docx.build().pack(&mut cursor).unwrap();
            let (document, _) = DocxParser
                .parse(
                    &cursor.into_inner(),
                    &crate::config::ConvertOptions::default(),
                )
                .expect("the picture document parses");
            let output = generate_typst(&document).unwrap();
            let pictures =
                crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
                    .expect("the picture document compiles");
            let runs = crate::render::pdf::compiled_text_runs_with_images(
                &output.source,
                &output.images,
                0,
            )
            .expect("the picture document compiles");
            let run_with = |needle: &str| {
                runs.iter()
                    .find(|run| run.text.contains(needle))
                    .unwrap_or_else(|| panic!("{case}: no run holds {needle:?} in {runs:?}"))
            };
            let [picture] = pictures.as_slice() else {
                panic!("{case}: expected one picture, got {pictures:?}");
            };
            let (picture_left, picture_top) = picture.corners[0];
            let (picture_right, picture_bottom) = picture.corners[2];
            let (first, left, end, next) = (
                run_with("First"),
                run_with("Left"),
                run_with("End"),
                run_with("Next"),
            );
            assert!(
                (left.baseline_pt - picture_bottom).abs() < 0.05
                    && (end.baseline_pt - picture_bottom).abs() < 0.05,
                "{case}: the picture's bottom edge ({picture_bottom}pt) must sit on the baseline \
                 of the runs on either side, got {}pt and {}pt",
                left.baseline_pt,
                end.baseline_pt
            );
            assert!(
                left.left_pt < picture_left && (end.left_pt - picture_right).abs() < 0.05,
                "{case}: the picture ({picture_left}..{picture_right}pt) must follow `Left ` \
                 at {}pt and lead ` End`, which starts at {}pt",
                left.left_pt,
                end.left_pt
            );
            assert!(
                picture_top > first.baseline_pt,
                "{case}: the picture's top ({picture_top}pt) must stay below the line above, \
                 whose baseline is at {}pt",
                first.baseline_pt
            );
            gaps_below_baseline.push(next.baseline_pt - picture_bottom);
            if place == "body" {
                let pitch: f64 = run_with("Last").baseline_pt - next.baseline_pt;
                assert!(
                    (next.baseline_pt - picture_bottom - pitch).abs() < 0.05,
                    "{case}: the next line must follow one line pitch ({pitch}pt) below the \
                     picture line's baseline ({picture_bottom}pt), got {}pt",
                    next.baseline_pt
                );
            }
        }
        assert!(
            (gaps_below_baseline[0] - gaps_below_baseline[1]).abs() < 0.05,
            "{place}: a tall picture must leave the line's descent and leading alone, but the \
             next line follows {gaps_below_baseline:?}pt below the 36pt and 4pt picture lines"
        );
    }
}

/// A line under `w:lineRule="exact"` keeps its stated height whatever it
/// holds (issue #2228). An inline picture taller than the line's ascent still
/// stands on the baseline, but Word clips it to the line box instead of
/// growing the line: its export of docxide-pdf's `cell-inline-exact` shows
/// only the bottom 14.4pt of a 36pt picture, and the next paragraph starts one
/// 18pt line below the table's top. So the text and the next line sit where
/// they do without the picture, and the picture stays visible from the line's
/// top down to the baseline. A picture that fits is drawn whole, and an
/// at-least line, which states only a minimum, still grows to hold it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_inline_picture_taller_than_an_exact_line_is_clipped_to_the_line() {
    use crate::parser::{Parser, docx::DocxParser};

    struct Laid {
        left_baseline_pt: f64,
        end_baseline_pt: f64,
        next_baseline_pt: f64,
        picture: Option<crate::render::pdf::PlacedImageBox>,
    }
    let lay_out = |place: &str,
                   rule: docx_rs::LineSpacingType,
                   line_twips: i32,
                   picture_height_pt: Option<f64>| {
        let mut paragraph = docx_rs::Paragraph::new()
            .line_spacing(docx_rs::LineSpacing::new().line_rule(rule).line(line_twips))
            .add_run(docx_rs::Run::new().add_text("Left "));
        if let Some(height_pt) = picture_height_pt {
            paragraph = paragraph.add_run(docx_rs::Run::new().add_image(
                docx_rs::Pic::new(MINIMAL_PNG).size(914_400, (height_pt * 12_700.0) as u32),
            ));
        }
        let paragraph = paragraph.add_run(docx_rs::Run::new().add_text(" End"));
        // The paragraph, or the borderless table holding it, opens the page,
        // so its line's top is the top margin.
        let docx = docx_rs::Docx::new().page_margin(
            docx_rs::PageMargin::new()
                .top(1440)
                .bottom(1440)
                .left(1440)
                .right(1440),
        );
        let docx = match place {
            "body" => docx.add_paragraph(paragraph),
            _ => docx.add_table(docx_rs::Table::without_borders(vec![
                docx_rs::TableRow::new(vec![docx_rs::TableCell::new().add_paragraph(paragraph)]),
            ])),
        }
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Next line")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let (document, _) = DocxParser
            .parse(
                &cursor.into_inner(),
                &crate::config::ConvertOptions::default(),
            )
            .expect("the exact-line document parses");
        let output = generate_typst(&document).unwrap();
        let pictures = crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
            .expect("the exact-line document compiles");
        let runs =
            crate::render::pdf::compiled_text_runs_with_images(&output.source, &output.images, 0)
                .expect("the exact-line document compiles");
        let baseline_of = |needle: &str| {
            runs.iter()
                .find(|run| run.text.contains(needle))
                .unwrap_or_else(|| panic!("no run holds {needle:?} in {runs:?}"))
                .baseline_pt
        };
        assert!(
            pictures.len() <= 1,
            "expected at most one picture, got {pictures:?}"
        );
        Laid {
            left_baseline_pt: baseline_of("Left"),
            end_baseline_pt: baseline_of("End"),
            next_baseline_pt: baseline_of("Next"),
            picture: pictures.first().copied(),
        }
    };
    let line_top_pt: f64 = 72.0;
    let drawn_whole = |picture: &crate::render::pdf::PlacedImageBox| {
        let (left, top) = picture.corners[0];
        let (right, bottom) = picture.corners[2];
        picture
            .clip
            .is_none_or(|(clip_left, clip_top, clip_right, clip_bottom)| {
                clip_left <= left + 0.05
                    && clip_top <= top + 0.05
                    && clip_right >= right - 0.05
                    && clip_bottom >= bottom - 0.05
            })
    };
    for place in ["body", "table cell"] {
        // 18pt and 24pt exact lines.
        for line_twips in [360, 480] {
            let without_picture: Laid =
                lay_out(place, docx_rs::LineSpacingType::Exact, line_twips, None);
            let ascent_pt: f64 = without_picture.left_baseline_pt - line_top_pt;
            for height_pt in [36.0, 60.0, 4.0] {
                let case = format!(
                    "{height_pt}pt picture on a {line_twips}-twip exact line in the {place}"
                );
                let laid: Laid = lay_out(
                    place,
                    docx_rs::LineSpacingType::Exact,
                    line_twips,
                    Some(height_pt),
                );
                let picture = laid
                    .picture
                    .unwrap_or_else(|| panic!("{case}: the picture is not painted"));
                let picture_bottom_pt: f64 = picture.corners[2].1;
                assert!(
                    (laid.left_baseline_pt - without_picture.left_baseline_pt).abs() < 0.05
                        && (laid.end_baseline_pt - without_picture.left_baseline_pt).abs() < 0.05,
                    "{case}: the text must keep the exact line's baseline ({}pt), got {}pt and {}pt",
                    without_picture.left_baseline_pt,
                    laid.left_baseline_pt,
                    laid.end_baseline_pt
                );
                assert!(
                    (picture_bottom_pt - laid.left_baseline_pt).abs() < 0.05,
                    "{case}: the picture's bottom ({picture_bottom_pt}pt) must sit on the \
                     baseline ({}pt)",
                    laid.left_baseline_pt
                );
                assert!(
                    (laid.next_baseline_pt - without_picture.next_baseline_pt).abs() < 0.05,
                    "{case}: the next line must follow the exact line as it does with no \
                     picture ({}pt), got {}pt",
                    without_picture.next_baseline_pt,
                    laid.next_baseline_pt
                );
                if height_pt > ascent_pt {
                    let (clip_left, clip_top, clip_right, clip_bottom) = picture
                        .clip
                        .unwrap_or_else(|| panic!("{case}: the picture overflows unclipped"));
                    assert!(
                        (clip_top - line_top_pt).abs() < 0.05
                            && clip_bottom >= picture_bottom_pt - 0.05
                            && clip_left <= picture.corners[0].0 + 0.05
                            && clip_right >= picture.corners[2].0 - 0.05,
                        "{case}: the picture must stay visible across its width from the \
                         line's top ({line_top_pt}pt) to the baseline ({picture_bottom_pt}pt), \
                         got the clip {:?}",
                        picture.clip
                    );
                } else {
                    assert!(
                        drawn_whole(&picture),
                        "{case}: a picture inside the line must be drawn whole, got the clip {:?}",
                        picture.clip
                    );
                }
            }
        }
        let without_picture: Laid = lay_out(place, docx_rs::LineSpacingType::AtLeast, 360, None);
        let laid: Laid = lay_out(place, docx_rs::LineSpacingType::AtLeast, 360, Some(36.0));
        let picture = laid
            .picture
            .unwrap_or_else(|| panic!("at-least line in the {place}: the picture is not painted"));
        assert!(
            laid.next_baseline_pt > without_picture.next_baseline_pt + 10.0
                && drawn_whole(&picture),
            "an at-least line in the {place} must grow to hold the whole picture: the next line \
             moved from {}pt to {}pt, clip {:?}",
            without_picture.next_baseline_pt,
            laid.next_baseline_pt,
            picture.clip
        );
    }
}

fn make_image(format: ImageFormat, width: Option<f64>, height: Option<f64>) -> Block {
    Block::Image(ImageData {
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: MINIMAL_PNG.to_vec(),
        format,
        width,
        height,
        crop: None,
        stroke: None,
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: None,
        paragraph_indent: None,
    })
}

/// Word advances a picture paragraph by the picture plus the paragraph's own
/// `w:spacing`. Zeroing both gaps to keep Typst's 1.2em default away also
/// discarded the declared gap, so content below a figure sat one `w:after`
/// too high (issue #499).
#[test]
fn picture_paragraph_spacing_becomes_block_gaps() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: MINIMAL_PNG.to_vec(),
        format: ImageFormat::Png,
        width: None,
        height: None,
        crop: None,
        stroke: None,
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: Some(ImageParagraphSpacing {
            before: Some(6.0),
            after: Some(3.0),
        }),
        paragraph_indent: None,
    })])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("above: 6pt") && output.source.contains("below: 3pt"),
        "declared picture paragraph spacing should reach the block: {}",
        output.source
    );
}

/// Triangulation for [`picture_paragraph_spacing_becomes_block_gaps`]: a
/// picture with no declared spacing must still pin both gaps to zero, or
/// Typst's 1.2em default reopens ~24pt around the figure (issues #463, #491).
#[test]
fn picture_without_declared_spacing_keeps_zero_gaps() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("above: 0pt") && output.source.contains("below: 0pt"),
        "an unspaced picture should keep both gaps at zero: {}",
        output.source
    );
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
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
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
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: None,
        paragraph_indent: None,
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
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: MINIMAL_PNG.to_vec(),
        format: ImageFormat::Png,
        width: Some(127.0),
        height: Some(227.0),
        crop: None,
        stroke: Some(BorderSide {
            width: 6.0,
            color: Color { r: 152, g: 0, b: 0 },
            style: BorderLineStyle::Solid,
            join: LineJoin::Round,
            cap: LineCap::Flat,
        }),
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: None,
        paragraph_indent: None,
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

/// A bordered picture casts its shadow from the same stroked silhouette a
/// shape does — the frame grown by half the line width, then offset by `dist`
/// (issue #1057, measured on a native PowerPoint export; see
/// `shadow_outline_outset`). The picture path builds its crisp duplicate
/// separately from `write_shadow_shape`, so it needs its own guard.
#[test]
fn a_bordered_picture_casts_its_shadow_from_the_stroked_frame() {
    use crate::ir::Shadow;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 50.0,
            width: 200.0,
            height: 120.0,
            kind: FixedElementKind::Image(ImageData {
                rotation_deg: None,
                flip_h: false,
                flip_v: false,
                data: MINIMAL_PNG.to_vec(),
                format: ImageFormat::Png,
                width: Some(200.0),
                height: Some(120.0),
                crop: None,
                stroke: Some(BorderSide {
                    width: 3.0,
                    color: Color {
                        r: 255,
                        g: 255,
                        b: 255,
                    },
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                    cap: LineCap::Flat,
                }),
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                fill_rect: None,
                shadow: Some(Shadow {
                    blur_radius: 0.0,
                    distance: 4.0,
                    direction: 0.0,
                    color: Color { r: 0, g: 0, b: 0 },
                    opacity: 0.38,
                }),
                paragraph_spacing: None,
                paragraph_indent: None,
            }),
        }],
    )]);
    let source = generate_typst(&doc).unwrap().source;

    // 200x120 frame outset by half of the 3pt border, moved 4pt right.
    assert!(
        source.contains("dx: 2.5pt, dy: -1.5pt, rect(width: 203pt, height: 123pt"),
        "a 3pt border must grow the picture's shadow silhouette by 1.5pt a \
         side: {source}"
    );
}

/// The stroked silhouette turns its corners the way the border's join does, so
/// a picture's crisp duplicate carries the same arc a shape's does — an outline
/// half-width under DrawingML's round default (issue #1138). This path emits
/// the duplicate separately from `write_shadow_shape`, so it needs its own
/// guard.
#[test]
fn a_bordered_picture_rounds_its_shadow_corner_like_the_border_join() {
    use crate::ir::Shadow;

    let source_with_join = |join: LineJoin| {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![FixedElement {
                x: 100.0,
                y: 50.0,
                width: 200.0,
                height: 120.0,
                kind: FixedElementKind::Image(ImageData {
                    rotation_deg: None,
                    flip_h: false,
                    flip_v: false,
                    data: MINIMAL_PNG.to_vec(),
                    format: ImageFormat::Png,
                    width: Some(200.0),
                    height: Some(120.0),
                    crop: None,
                    stroke: Some(BorderSide {
                        width: 3.0,
                        color: Color {
                            r: 255,
                            g: 255,
                            b: 255,
                        },
                        style: BorderLineStyle::Solid,
                        join,
                        cap: LineCap::Flat,
                    }),
                    alignment: None,
                    stroke_outline: None,
                    clip_shape: None,
                    fill_rect: None,
                    // Crisp, so the border's own arc is the whole radius.
                    shadow: Some(Shadow {
                        blur_radius: 0.0,
                        distance: 4.0,
                        direction: 0.0,
                        color: Color { r: 0, g: 0, b: 0 },
                        opacity: 0.38,
                    }),
                    paragraph_spacing: None,
                    paragraph_indent: None,
                }),
            }],
        )]);
        generate_typst(&doc).unwrap().source
    };

    let rounded: String = source_with_join(LineJoin::Round);
    assert!(
        rounded.contains("rect(width: 203pt, height: 123pt, radius: 1.5pt"),
        "a 3pt round-joined border must arc the shadow corner by 1.5pt: \
         {rounded}"
    );
    // Triangulation: a mitre runs out to the corner of the outset frame, so
    // it leaves the silhouette square.
    let mitred: String = source_with_join(LineJoin::Miter);
    assert!(
        mitred.contains("rect(width: 203pt, height: 123pt, radius: 0pt"),
        "a mitred border leaves the shadow corner square: {mitred}"
    );
}

/// A blurred picture uses the same continuous Gaussian asset as a shape.
/// This path is emitted separately from `write_shadow_shape`, so it needs its
/// own guard (issues #1204, #1309).
#[test]
fn a_blurred_picture_uses_one_filtered_frame_asset() {
    use crate::ir::Shadow;

    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 100.0,
            y: 50.0,
            width: 200.0,
            height: 120.0,
            kind: FixedElementKind::Image(ImageData {
                rotation_deg: None,
                flip_h: false,
                flip_v: false,
                data: MINIMAL_PNG.to_vec(),
                format: ImageFormat::Png,
                width: Some(200.0),
                height: Some(120.0),
                crop: None,
                stroke: None,
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                fill_rect: None,
                shadow: Some(Shadow {
                    // The theme effect `customGeo.pptx` page 46 carries.
                    blur_radius: 40000.0 / 12700.0,
                    distance: 1.57,
                    direction: 90.0,
                    color: Color { r: 0, g: 0, b: 0 },
                    opacity: 0.38,
                }),
                paragraph_spacing: None,
                paragraph_indent: None,
            }),
        }],
    )]);
    let output = generate_typst(&doc).unwrap();
    let shadow = output
        .images
        .iter()
        .find(|image| image.path.ends_with(".svg"))
        .expect("one generated shadow SVG");
    let svg = std::str::from_utf8(&shadow.data).unwrap();
    assert!(
        svg.contains("<feGaussianBlur")
            && svg.contains("<rect")
            && svg.contains("width=\"200\"")
            && svg.contains("height=\"120\""),
        "the picture frame itself must feed one Gaussian filter: {svg}"
    );
    assert!(
        output.source.find(&shadow.path) < output.source.find("#image(\"img-1.png\""),
        "the filtered shadow must sit before the picture: {}",
        output.source,
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
                rotation_deg: None,
                flip_h: false,
                flip_v: false,
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
                    join: LineJoin::Round,
                    cap: LineCap::Flat,
                }),
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                fill_rect: None,
                shadow: None,
                paragraph_spacing: None,
                paragraph_indent: None,
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

/// OOXML mirrors the cropped/clipped picture frame before rotating it. Keep
/// the scale wrapper inside the rotate wrapper so non-square artwork follows
/// the same transform order (issue #1017).
#[test]
fn a_fixed_picture_flip_is_applied_inside_its_rotation() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 10.0,
            y: 20.0,
            width: 120.0,
            height: 60.0,
            kind: FixedElementKind::Image(ImageData {
                rotation_deg: Some(90.0),
                flip_h: true,
                flip_v: false,
                data: MINIMAL_PNG.to_vec(),
                format: ImageFormat::Png,
                width: Some(120.0),
                height: Some(60.0),
                crop: None,
                stroke: None,
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                fill_rect: None,
                shadow: None,
                paragraph_spacing: None,
                paragraph_indent: None,
            }),
        }],
    )]);

    let source = generate_typst(&doc).unwrap().source;
    let rotate = source
        .find("#rotate(90deg, origin: top + left)[")
        .expect("the outer picture rotation must be emitted");
    let flip = source
        .find("#scale(x: -100%, y: 100%, origin: top + left)[")
        .expect("flipH must become a horizontal mirror");
    let image = source
        .find("#image(\"")
        .expect("the picture must be emitted");
    assert!(
        rotate < flip && flip < image,
        "wrong transform order: {source}"
    );
}

/// A picture frame whose box reaches past the slide, mirrored and turned:
/// `Grafikk 310` on the CONTOSO deck's layout 15 (issue #1032). PowerPoint
/// mirrors it about its own centre and turns it about that same centre,
/// however far the box hangs off the slide.
///
/// Its two tests read the compiled frames, which needs a Typst compilation
/// the wasm test target does not build.
#[cfg(not(target_arch = "wasm32"))]
fn oversized_turned_picture() -> (Document, [(f64, f64); 4]) {
    const X: f64 = -89.37716535433071;
    const Y: f64 = -154.79425196850394;
    const WIDTH: f64 = 848.043937007874;
    const HEIGHT: f64 = 856.8;
    const ROTATION_DEG: f64 = 135.37011666666666;

    let doc: Document = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: X,
            y: Y,
            width: WIDTH,
            height: HEIGHT,
            kind: FixedElementKind::Image(ImageData {
                rotation_deg: Some(ROTATION_DEG),
                flip_h: false,
                flip_v: true,
                data: make_quadrant_png(),
                format: ImageFormat::Png,
                width: Some(WIDTH),
                height: Some(HEIGHT),
                crop: None,
                stroke: None,
                alignment: None,
                stroke_outline: None,
                clip_shape: None,
                fill_rect: None,
                shadow: None,
                paragraph_spacing: None,
                paragraph_indent: None,
            }),
        }],
    )]);

    // Mirror about the box's horizontal centre line, then turn about the
    // box centre — DrawingML's own order, evaluated here independently of
    // the markup under test.
    let (centre_x, centre_y): (f64, f64) = (X + WIDTH / 2.0, Y + HEIGHT / 2.0);
    let (sin, cos): (f64, f64) = ROTATION_DEG.to_radians().sin_cos();
    let seat = |u: f64, v: f64| -> (f64, f64) {
        let (dx, dy): (f64, f64) = (X + u - centre_x, Y + (HEIGHT - v) - centre_y);
        (
            centre_x + cos * dx - sin * dy,
            centre_y + sin * dx + cos * dy,
        )
    };
    let expected: [(f64, f64); 4] = [
        seat(0.0, 0.0),
        seat(WIDTH, 0.0),
        seat(WIDTH, HEIGHT),
        seat(0.0, HEIGHT),
    ];
    (doc, expected)
}

/// Typst resolves `origin: center` against the frame it lays the body out
/// in, and that frame is clamped to the region. A picture box taller than
/// the slide therefore turned about the slide's midpoint instead of its own,
/// landing 119pt off its seat and dragging the artwork with it (issue #1032).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_oversized_turned_picture_pivots_on_its_own_centre() {
    let (doc, expected) = oversized_turned_picture();
    let output = generate_typst(&doc).unwrap();
    let boxes = crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
        .expect("the slide compiles");
    let [placed] = boxes.as_slice() else {
        panic!("the slide carries exactly one picture, got {boxes:?}");
    };

    for (index, (want, got)) in expected.iter().zip(placed.corners.iter()).enumerate() {
        assert!(
            (want.0 - got.0).abs() < 0.05 && (want.1 - got.1).abs() < 0.05,
            "corner {index} sits at {got:?}, PowerPoint seats it at {want:?}"
        );
    }
}

/// The same pivot, read as the invariant it protects: a turn about the box
/// centre leaves that centre where the frame put it, whatever the angle.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_turned_pictures_centre_stays_on_its_frame_centre() {
    let (doc, _expected) = oversized_turned_picture();
    let output = generate_typst(&doc).unwrap();
    let boxes = crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
        .expect("the slide compiles");
    let [placed] = boxes.as_slice() else {
        panic!("the slide carries exactly one picture, got {boxes:?}");
    };
    let centre: (f64, f64) = (
        (placed.corners[0].0 + placed.corners[2].0) / 2.0,
        (placed.corners[0].1 + placed.corners[2].1) / 2.0,
    );
    // -89.37716535433071 + 848.043937007874 / 2, -154.79425196850394 + 856.8 / 2
    assert!(
        (centre.0 - 334.6448031496063).abs() < 0.05 && (centre.1 - 273.60574803149603).abs() < 0.05,
        "the turned picture's centre drifted to {centre:?}"
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

/// An `a:srcRect` crop reaches an SVG asset too, by narrowing its viewBox
/// (issue #892).
///
/// `preprocess_image_asset` cropped by decoding the bitmap, and
/// `raster_image_format` returns `None` for an SVG, so a cropped vector
/// graphic was drawn whole and stretched into a box sized for the retained
/// part. The deck in #841 does this to a 6-column dot grid with
/// `<a:srcRect r="49115"/>`: three columns should show, six did.
#[test]
fn an_svg_crop_narrows_the_view_box() {
    const SVG: &str = r#"<svg width="365" height="340" viewBox="0 0 365 340" xmlns="http://www.w3.org/2000/svg" overflow="hidden"><circle cx="10" cy="10" r="5"/></svg>"#;
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: SVG.as_bytes().to_vec(),
        format: ImageFormat::Svg,
        width: Some(20.0),
        height: Some(40.0),
        crop: Some(ImageCrop {
            left: 0.0,
            top: 0.4302,
            right: 0.49115,
            bottom: 0.0,
        }),
        stroke: None,
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: None,
        paragraph_indent: None,
    })])]);

    let output = generate_typst(&doc).unwrap();
    let svg = String::from_utf8(output.images[0].data.clone()).expect("the asset stays an SVG");

    // 49.115% off the right of 365 leaves 185.73; 43.02% off the top of 340
    // leaves y from 146.27 for 193.73.
    let view_box = svg
        .split("viewBox=\"")
        .nth(1)
        .and_then(|rest| rest.split('"').next())
        .expect("a viewBox survives");
    let values: Vec<f64> = view_box
        .split_whitespace()
        .map(|v| v.parse().expect("numeric viewBox"))
        .collect();
    let expected = [0.0, 146.27, 185.73, 193.73];
    for (index, (got, want)) in values.iter().zip(expected.iter()).enumerate() {
        assert!(
            (got - want).abs() < 0.05,
            "viewBox component {index}: got {got}, want {want} (viewBox {view_box})"
        );
    }
    // The viewport must shrink with the viewBox. Left at its old size,
    // `preserveAspectRatio` meets the smaller viewBox inside it — scaling by
    // one and centring the drawing instead of cropping it.
    for (name, want) in [("width", 185.73), ("height", 193.73)] {
        let got: f64 = svg
            .split(&format!("{name}=\""))
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .and_then(|value| value.parse().ok())
            .unwrap_or_else(|| panic!("a numeric {name} survives: {svg}"));
        assert!(
            (got - want).abs() < 0.05,
            "{name}: got {got}, want {want} ({svg})"
        );
    }
    assert!(
        svg.contains(r#"<circle cx="10" cy="10" r="5"/>"#),
        "the drawing itself is untouched: {svg}"
    );
}

/// An SVG with no crop is passed through byte for byte, so nothing is rewritten
/// on the far commoner path.
#[test]
fn an_uncropped_svg_is_left_alone() {
    const SVG: &str =
        r#"<svg width="10" height="10" viewBox="0 0 10 10"><rect width="10" height="10"/></svg>"#;
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: SVG.as_bytes().to_vec(),
        format: ImageFormat::Svg,
        width: Some(10.0),
        height: Some(10.0),
        crop: None,
        stroke: None,
        alignment: None,
        stroke_outline: None,
        clip_shape: None,
        fill_rect: None,
        shadow: None,
        paragraph_spacing: None,
        paragraph_indent: None,
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images[0].data, SVG.as_bytes());
}
