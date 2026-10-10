use super::*;

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn page_anchored_text_box_position_does_not_follow_flow_position() {
    fn anchored_text_baseline(preceding_paragraph_count: usize) -> f64 {
        let mut content: Vec<Block> = (0..preceding_paragraph_count)
            .map(|paragraph_index| {
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: format!("preceding flow paragraph {paragraph_index} ").repeat(12),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                        inline_box: None,
                    }],
                })
            })
            .collect();
        content.push(Block::FloatingTextBox(FloatingTextBox {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Page anchored text".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                    inline_box: None,
                }],
            })],
            wrap_mode: WrapMode::Behind,
            width: 120.0,
            height: 48.0,
            fill: None,
            stroke: None,
            shape_rotation_deg: None,
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Page,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
            offset_x: 0.0,
            offset_y: 240.0,
        }));
        let document: Document = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                first_header: None,
                first_footer: None,
                header: None,
                footer: None,
                columns: None,
                line_grid_pitch: None,
                line_grid_snaps_lines: false,
                page_numbering: None,
                size: PageSize {
                    width: 612.0,
                    height: 792.0,
                },
                margins: Margins {
                    top: 72.0,
                    right: 72.0,
                    bottom: 72.0,
                    left: 72.0,
                },
                content,
            })],
            styles: crate::ir::StyleSheet::default(),
        };
        let output: TypstOutput =
            generate_typst(&document).expect("the test document should generate");
        (0..4)
            .filter_map(|page_index| {
                crate::render::pdf::compiled_text_runs(&output.source, page_index).ok()
            })
            .flatten()
            .find(|run| run.text == "Page anchored text")
            .expect("the floating textbox text should remain searchable")
            .baseline_pt
    }

    let baseline_without_flow: f64 = anchored_text_baseline(0);
    let baseline_after_flow: f64 = anchored_text_baseline(8);
    assert!(
        (baseline_without_flow - baseline_after_flow).abs() < 0.25,
        "page anchoring should keep the baseline stable when preceding flow changes; before={baseline_without_flow}pt after={baseline_after_flow}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn page_anchored_text_box_paints_its_declared_frame_at_the_page_offset() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};

    let page_size: PageSize = PageSize::default();
    let document: Document = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Page frame")],
            wrap_mode: WrapMode::Behind,
            width: 120.0,
            height: 48.0,
            fill: Some(Color::white()),
            stroke: Some(BorderSide {
                width: 0.75,
                color: Color::new(0, 176, 240),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            shape_rotation_deg: None,
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Page,
            vertical_position_align: None,
            horizontal_align: Some(crate::ir::FrameAlign::End),
            horizontal_anchor: crate::ir::FrameAnchor::Page,
            offset_x: 0.0,
            offset_y: 240.0,
        },
    )])]);
    let output: TypstOutput = generate_typst(&document).expect("the page should generate");
    assert!(output.source.contains("fill: rgb(255, 255, 255)"));
    assert!(output.source.contains("paint: rgb(0, 176, 240)"));
    assert!(output.source.contains("thickness: 0.75pt"));

    let paints = compiled_paint_sequence(&output.source, &output.images, 0)
        .expect("the page should compile to paint geometry");
    let frame = paints
        .iter()
        .find(|paint| {
            paint.kind == PaintedKind::Shape
                && paint.rectangle_fill.is_some()
                && paint
                    .stroke
                    .as_ref()
                    .is_some_and(|stroke| (stroke.thickness_pt - 0.75).abs() < 0.01)
        })
        .expect("the text box should paint its filled and stroked frame");
    let expected_left: f64 = page_size.width - 120.0;

    assert!((frame.bounds.0 - expected_left).abs() < 0.25);
    assert!((frame.bounds.1 - 240.0).abs() < 0.25);
    assert!((frame.bounds.2 - page_size.width).abs() < 0.25);
    assert!((frame.bounds.3 - 288.0).abs() < 0.25);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_right_aligned_anchored_text_box_uses_the_page_margin() {
    use crate::parser::Parser;

    const FIXTURE: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/libreoffice/tdf105688.docx");
    let (document, _warnings) = crate::parser::docx::DocxParser
        .parse(FIXTURE, &crate::config::ConvertOptions::default())
        .expect("the reported DOCX should parse");
    let mut isolated_page = document
        .pages
        .iter()
        .find_map(|page| match page {
            Page::Flow(page) | Page::FlowContinuous(page)
                if page.content.iter().any(|block| {
                    matches!(
                        block,
                        Block::FloatingTextBox(text_box)
                            if (text_box.width - 156.25).abs() < 0.01
                    )
                }) =>
            {
                Some(page.clone())
            }
            _ => None,
        })
        .expect("the reported anchored text box should be in the flow");
    let text_box = isolated_page
        .content
        .iter()
        .find(|block| {
            matches!(
                block,
                Block::FloatingTextBox(text_box) if (text_box.width - 156.25).abs() < 0.01
            )
        })
        .expect("the reported anchored text box should be in the flow")
        .clone();
    isolated_page.content = vec![text_box];
    isolated_page.header = None;
    isolated_page.footer = None;
    isolated_page.first_header = None;
    isolated_page.first_footer = None;
    let page_width: f64 = isolated_page.size.width;
    let expected_right_margin: f64 = page_width - isolated_page.margins.right;
    let isolated_document = Document {
        metadata: document.metadata.clone(),
        pages: vec![Page::Flow(isolated_page)],
        styles: document.styles.clone(),
    };
    let output = generate_typst(&isolated_document)
        .expect("the isolated anchored text box should generate Typst");
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .expect("the isolated page should compile");
    let quote = runs
        .iter()
        .find(|run| run.text == "“")
        .expect("the anchored box should keep its opening quote");
    let paints = crate::render::pdf::compiled_paint_sequence(&output.source, &output.images, 0)
        .expect("the isolated page should produce paint geometry");
    let rightmost_shape_edge: f64 = paints
        .iter()
        .filter(|paint| paint.kind == crate::render::pdf::PaintedKind::Shape)
        .map(|paint| paint.bounds.2)
        .reduce(f64::max)
        .expect("the anchored table border should paint shape geometry");

    assert!(
        quote.left_pt > page_width / 2.0,
        "the quote should follow the right-aligned box on the right half of the page, got x={}pt",
        quote.left_pt
    );
    assert!(
        (rightmost_shape_edge - expected_right_margin).abs() < 0.5,
        "the box's right edge should align to the page's right margin at {expected_right_margin}pt, got {rightmost_shape_edge}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn floating_text_box_margin_alignment_tracks_a_different_page_geometry() {
    let page_size: PageSize = PageSize {
        width: 720.0,
        height: 900.0,
    };
    let margins: Margins = Margins {
        top: 36.0,
        right: 80.0,
        bottom: 36.0,
        left: 20.0,
    };
    let document = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: page_size,
            margins,
            content: vec![
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Anchor paragraph".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                        inline_box: None,
                    }],
                }),
                Block::FloatingTextBox(FloatingTextBox {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Margin aligned".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                            inline_box: None,
                        }],
                    })],
                    wrap_mode: WrapMode::Behind,
                    width: 100.0,
                    height: 32.0,
                    fill: None,
                    stroke: None,
                    shape_rotation_deg: None,
                    padding: Insets::default(),
                    vertical_align: TextBoxVerticalAlign::Top,
                    vertical_anchor: crate::ir::FrameAnchor::Text,
                    vertical_position_align: None,
                    horizontal_align: Some(crate::ir::FrameAlign::End),
                    horizontal_anchor: crate::ir::FrameAnchor::Margin,
                    offset_x: 0.0,
                    offset_y: 0.0,
                }),
            ],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: StyleSheet::default(),
    };
    let output = generate_typst(&document).expect("the synthetic page should generate Typst");
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .expect("the synthetic page should compile");
    let aligned = runs
        .iter()
        .find(|run| run.text == "Margin aligned")
        .expect("the floating box text should be present");
    let expected_x: f64 = page_size.width - margins.right - 100.0;

    assert!(
        (aligned.left_pt - expected_x).abs() < 0.25,
        "right alignment should use the page's own margin geometry at x={expected_x}pt, got {}pt",
        aligned.left_pt
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn explicit_pptx_no_fill_line_does_not_emit_a_default_rectangle_outline() {
    use crate::parser::Parser;
    use crate::parser::pptx::PptxParser;
    use crate::render::pdf::PaintedKind;

    const FIXTURE: &[u8] = include_bytes!("../../../../tests/fixtures/pptx/poi/with_japanese.pptx");

    let (document, _warnings) = PptxParser
        .parse(FIXTURE, &crate::config::ConvertOptions::default())
        .expect("the reported fixture should parse");
    let output = generate_typst(&document).expect("the reported fixture should render");
    let paints = crate::render::pdf::compiled_paint_sequence(&output.source, &output.images, 0)
        .expect("the first slide should compile");

    let rectangle_paints: Vec<_> = paints
        .iter()
        .filter(|paint| {
            paint.kind == PaintedKind::Shape
                && (paint.bounds.0 - 0.0).abs() < 0.01
                && (paint.bounds.1 - 0.0).abs() < 0.01
                && (paint.bounds.2 - 720.0).abs() < 0.01
                && (paint.bounds.3 - 36.0).abs() < 0.01
        })
        .collect();

    assert!(
        rectangle_paints.is_empty(),
        "Rectangle 2 has no fill or line and must paint nothing, got {rectangle_paints:?}"
    );
}

#[test]
fn explicit_no_line_does_not_emit_an_arrowhead() {
    let stroke: Option<BorderSide> = Some(BorderSide {
        width: 0.0,
        color: Color::black(),
        style: BorderLineStyle::None,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    });
    let mut output: String = String::new();

    super::super::shapes::write_arrowhead_at(&mut output, &stroke, (0.0, 0.0), (10.0, 0.0));

    assert!(
        output.is_empty(),
        "an explicitly disabled line must not leave an arrowhead, got {output}"
    );
}

// ── Floating image codegen tests ──

#[test]
fn test_floating_image_square_wrap_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    rotation_deg: None,
                    flip_h: false,
                    flip_v: false,
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(200.0),
                    height: Some(100.0),
                    crop: None,
                    stroke: None,
                    alignment: None,
                    clip_shape: None,
                    shadow: None,
                    paragraph_spacing: None,
                },
                wrap_mode: WrapMode::Square,
                vertical_anchor: FloatingImageVerticalAnchor::Page,
                offset_x: 72.0,
                offset_y: 36.0,
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
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
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    rotation_deg: None,
                    flip_h: false,
                    flip_v: false,
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(150.0),
                    height: Some(75.0),
                    crop: None,
                    stroke: None,
                    alignment: None,
                    clip_shape: None,
                    shadow: None,
                    paragraph_spacing: None,
                },
                wrap_mode: WrapMode::TopAndBottom,
                vertical_anchor: FloatingImageVerticalAnchor::Page,
                offset_x: 10.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
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
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    rotation_deg: None,
                    flip_h: false,
                    flip_v: false,
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(100.0),
                    height: Some(50.0),
                    crop: None,
                    stroke: None,
                    alignment: None,
                    clip_shape: None,
                    shadow: None,
                    paragraph_spacing: None,
                },
                wrap_mode: WrapMode::Behind,
                vertical_anchor: FloatingImageVerticalAnchor::Page,
                offset_x: 0.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
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
            fill: None,
            stroke: None,
            shape_rotation_deg: None,
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Text,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
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
fn a_rotated_floating_text_box_keeps_its_declared_center() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Turned label")],
            wrap_mode: WrapMode::None,
            width: 200.0,
            height: 150.0,
            fill: None,
            stroke: None,
            shape_rotation_deg: Some(90.0),
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Text,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
            offset_x: 72.0,
            offset_y: 36.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains(
            "#move(dx: 175pt, dy: -25pt)[#rotate(90deg, origin: top + left, reflow: false)["
        ),
        "floating WPG text must turn around the same unclamped center as its shape:\n{}",
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
            fill: None,
            stroke: None,
            shape_rotation_deg: None,
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Text,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
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
fn test_floating_text_box_content_is_top_left_aligned_inside_bounds() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Top aligned")],
            wrap_mode: WrapMode::None,
            width: 120.0,
            height: 40.0,
            fill: None,
            stroke: None,
            shape_rotation_deg: None,
            padding: Insets::default(),
            vertical_align: TextBoxVerticalAlign::Top,
            vertical_anchor: crate::ir::FrameAnchor::Text,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
            offset_x: 10.0,
            offset_y: 5.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();

    assert!(
        output
            .source
            .contains("#box(width: 120pt, height: 40pt, inset: 0pt)["),
        "Expected floating text box bounds to use a zero-inset box, got:\n{}",
        output.source
    );
    assert!(
        output
            .source
            .contains("#place(top + left, dy: -6pt)[\n#block(width: 120pt)["),
        "Expected floating text box content to be placed at the top-left of its bounds, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_text_box_applies_padding_and_center_alignment() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Centered")],
            wrap_mode: WrapMode::None,
            width: 120.0,
            height: 60.0,
            fill: None,
            stroke: None,
            shape_rotation_deg: None,
            padding: Insets {
                top: 3.0,
                right: 6.0,
                bottom: 3.0,
                left: 6.0,
            },
            vertical_align: TextBoxVerticalAlign::Center,
            vertical_anchor: crate::ir::FrameAnchor::Text,
            vertical_position_align: None,
            horizontal_align: None,
            horizontal_anchor: crate::ir::FrameAnchor::Text,
            offset_x: 10.0,
            offset_y: 5.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();

    assert!(output.source.contains(
        "#box(width: 120pt, height: 60pt, inset: (top: 3pt, right: 6pt, bottom: 3pt, left: 6pt))["
    ));
    assert!(output.source.contains("block(width: 108pt)["));
    assert!(
        output
            .source
            .contains("calc.max(54pt - measure(floating_text_box_content_0).height, 0pt)")
    );
    assert!(output.source.contains("#v(floating_text_box_slack_0 / 2)"));
}

#[test]
fn test_consecutive_floating_shapes_share_one_anchor_line() {
    let shape = Shape {
        kind: ShapeKind::Rectangle,
        fill: Some(Color::new(114, 159, 207)),
        gradient_fill: None,
        pattern_fill: None,
        stroke: None,
        rotation_deg: None,
        opacity: None,
        shadow: None,
        top_bevel: None,
    };
    let doc = make_doc(vec![make_flow_page(vec![
        Block::FloatingShape(FloatingShape {
            shape: shape.clone(),
            width: 100.0,
            height: 40.0,
            offset_x: 20.0,
            offset_y: 10.0,
            wrap_mode: WrapMode::None,
        }),
        Block::FloatingShape(FloatingShape {
            shape,
            width: 100.0,
            height: 40.0,
            offset_x: 160.0,
            offset_y: 10.0,
            wrap_mode: WrapMode::None,
        }),
    ])]);

    let output = generate_typst(&doc).unwrap();
    let anchor_count = output
        .source
        .matches("#box(width: 0pt, height: 0pt)")
        .count();

    assert_eq!(
        anchor_count, 1,
        "Consecutive floating shapes from one DOCX paragraph should share one anchor line. Got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("dx: 20pt, dy: 10pt")
            && output.source.contains("dx: 160pt, dy: 10pt"),
        "Expected both floating shapes to remain in the shared anchor group. Got:\n{}",
        output.source
    );
}

#[test]
fn a_translucent_floating_shape_fill_uses_shape_opacity() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingShape(
        FloatingShape {
            shape: Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(0, 0, 0)),
                gradient_fill: None,
                pattern_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: Some(0.16),
                shadow: None,
                top_bevel: None,
            },
            width: 100.0,
            height: 40.0,
            offset_x: 20.0,
            offset_y: 10.0,
            wrap_mode: WrapMode::None,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("fill: rgb(0, 0, 0, 41)"),
        "shape opacity must reach the floating shape fill:\n{}",
        output.source
    );
}

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
    assert_eq!(
        output.source.matches("gradient.linear(").count(),
        output.source.matches("space: rgb").count(),
        "Every gradient should interpolate in sRGB. Got: {}",
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
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
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

#[test]
fn test_light_upward_diagonal_pattern_fill_codegen() {
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: None,
            gradient_fill: None,
            pattern_fill: Some(PatternFill {
                preset: PatternPreset::LightUpwardDiagonal,
                foreground: Color::new(0, 0, 255),
                background: Color::new(255, 255, 255),
            }),
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();

    assert!(
        output.source.contains("tiling(size: (2.72pt, 2.72pt))"),
        "Should emit a repeating DrawingML pattern. Got: {}",
        output.source,
    );
    assert!(
        output
            .source
            .contains("fill: rgb(255, 255, 255), stroke: none"),
        "Should paint the pattern background without an outline. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("stroke: 0.24pt + rgb(0, 0, 255)"),
        "Should paint the light upward hatch in the foreground color. Got: {}",
        output.source,
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn test_all_pattern_presets_compile_to_pdf() {
    let elements: Vec<FixedElement> = PatternPreset::ALL
        .iter()
        .copied()
        .enumerate()
        .map(|(index, preset)| FixedElement {
            x: (index % 9) as f64 * 48.0,
            y: (index / 9) as f64 * 48.0,
            width: 40.0,
            height: 40.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: None,
                gradient_fill: None,
                pattern_fill: Some(PatternFill {
                    preset,
                    foreground: Color::new(0, 0, 255),
                    background: Color::new(255, 255, 255),
                }),
                stroke: None,
                rotation_deg: None,
                opacity: None,
                shadow: None,
                top_bevel: None,
            }),
        })
        .collect();
    let doc = make_doc(vec![make_fixed_page(440.0, 300.0, elements)]);
    let output = generate_typst(&doc).unwrap();
    let pdf =
        crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
            .expect("Every DrawingML preset pattern should compile as Typst");

    assert!(pdf.starts_with(b"%PDF"));
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
            pattern_fill: None,
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
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images.len(), 1, "the blur should be one asset");
    let shadow_pos = output.source.find(&output.images[0].path);
    let main_pos = output.source.find("rgb(255, 0, 0)");
    assert!(
        shadow_pos < main_pos,
        "Shadow should appear before main shape in output",
    );
}

/// A custom-geometry shape casts the shadow of its own outline. Before the
/// `<a:custGeom>` path was understood, these shapes fell back to a rectangle
/// and got a rectangle's shadow; once the path arrived the shadow vanished
/// altogether, because no silhouette knew how to follow it (issue #1205).
#[test]
fn test_path_shape_casts_its_outline_as_a_shadow() {
    use crate::ir::{Shadow, Subpath};

    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Path {
                subpaths: vec![Subpath::closed_outline(vec![
                    (0.5, 0.0),
                    (1.0, 1.0),
                    (0.0, 1.0),
                ])],
            },
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: None,
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 3.0,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.35,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source = generate_typst(&doc).unwrap().source;

    let shadow_pos = source
        .find("rgb(0, 0, 0, ")
        .unwrap_or_else(|| panic!("no shadow layer in {source}"));
    let main_pos = source
        .find("rgb(255, 0, 0)")
        .expect("the shape itself still draws");
    assert!(shadow_pos < main_pos, "the shadow sits under the shape");
    assert!(
        source[..shadow_pos].ends_with("even-odd\", fill: "),
        "the shadow is a curve following the path, not a rectangle: {source}"
    );
    assert_eq!(
        source.matches("curve.move(").count(),
        2,
        "one outline for the shadow and one for the shape: {source}"
    );
}

/// An unclosed subpath casts the shadow of its stroke rather than a filled
/// silhouette. Filling it would paint the area an elbow connector merely
/// brackets, while skipping it loses the offset grey band PowerPoint draws
/// under the whole connector (issues #1205, #1305).
#[test]
fn test_open_subpath_casts_an_offset_copy_of_its_stroke() {
    use crate::ir::{Shadow, Subpath};

    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Path {
                subpaths: vec![Subpath::open_outline(vec![
                    (0.0, 0.0),
                    (0.0, 1.0),
                    (1.0, 1.0),
                ])],
            },
            fill: None,
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 2.0,
                color: Color::new(138, 180, 226),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 3.0,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.35,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source = generate_typst(&doc).unwrap().source;

    let shadow = source
        .lines()
        .find(|line| line.contains("rgb(0, 0, 0, 89)"))
        .unwrap_or_else(|| panic!("the open connector casts no shadow stroke: {source}"));
    assert!(
        shadow.contains("dy: 3pt")
            && shadow.contains("thickness: 2pt")
            && shadow.contains("join: \"round\"")
            && shadow.contains("curve.move((0pt, 0pt))")
            && shadow.contains("curve.line((0pt, 150pt))")
            && shadow.contains("curve.line((200pt, 150pt))"),
        "the shadow must offset the source stroke without changing its path, width, or join: {shadow}"
    );
    assert!(
        !shadow.contains("fill:") && !shadow.contains("curve.close()"),
        "an open connector shadow must remain an open stroke: {shadow}"
    );
}

#[test]
fn test_blurred_open_subpath_filters_its_stroke_without_filling_it() {
    use crate::ir::{Shadow, Subpath};

    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Path {
                subpaths: vec![Subpath::open_outline(vec![
                    (0.0, 0.0),
                    (0.0, 1.0),
                    (1.0, 1.0),
                ])],
            },
            fill: None,
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 2.0,
                color: Color::new(138, 180, 226),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 6.0,
                distance: 3.0,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.35,
            }),
            top_bevel: None,
        }),
    };
    let output =
        generate_typst(&make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])])).unwrap();
    let asset = output.images.first().expect("one blurred shadow asset");
    let svg = std::str::from_utf8(&asset.data).expect("a generated shadow SVG");

    assert!(
        svg.contains("<feGaussianBlur stdDeviation=\"2\"")
            && svg.contains("stroke=\"rgb(0, 0, 0)\"")
            && svg.contains("stroke-width=\"2\"")
            && svg.contains("stroke-linejoin=\"round\"")
            && svg.contains("<path d=\"M ")
            && svg.contains(" L ")
            && svg.contains("fill=\"none\""),
        "the Gaussian source must be the connector's open stroke: {svg}"
    );
    assert!(
        !svg.contains(" Z "),
        "the filtered connector must not gain a closing segment: {svg}"
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
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
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
fn test_rectangle_top_bevel_draws_four_lit_rim_faces() {
    use crate::ir::TopBevel;

    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 100.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(79, 129, 189)),
            gradient_fill: None,
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: Some(TopBevel {
                width: 5.0,
                height: 2.0,
                light_rig_rotation_deg: 20.0,
            }),
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();

    assert!(
        output.source.matches("#polygon(").count() >= 4,
        "the bevel must overlay one face on each rectangular edge: {}",
        output.source,
    );
    assert!(
        output.source.contains("(5pt, 5pt)") && output.source.contains("(195pt, 95pt)"),
        "the four faces must meet on the requested 5pt inset: {}",
        output.source,
    );
    assert_eq!(
        output.source.matches("gradient.linear(").count(),
        output.source.matches("space: rgb").count(),
        "Every gradient should interpolate in sRGB. Got: {}",
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

#[test]
fn test_shape_shadow_blur_uses_one_continuous_gaussian_asset() {
    // A stack of flat-alpha copies quantises the falloff at every ring
    // boundary. The native PowerPoint export instead carries one continuous
    // Gaussian ramp, so a blurred shadow must travel as one filtered asset,
    // not as separately painted Typst shapes (issue #1309).
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
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 8.0,
                distance: 3.0,
                direction: 45.0,
                color: Color::new(0, 0, 0),
                opacity: 0.5,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();

    assert_eq!(
        output.images.len(),
        1,
        "one blurred shadow must produce one continuous asset"
    );
    let asset = &output.images[0];
    assert!(asset.path.ends_with(".svg"), "asset path: {}", asset.path);
    let svg = std::str::from_utf8(&asset.data).expect("shadow asset must be SVG");
    assert!(
        svg.contains("<feGaussianBlur") && svg.contains("stdDeviation=\"2.666666"),
        "shadow must carry sigma = blurRad / 3: {svg}"
    );
    assert!(
        svg.contains("opacity=\"0.5\""),
        "the continuous mask must apply the declared opacity once to the filtered silhouette: {svg}"
    );
    assert!(
        output.source.contains("#pdf.artifact(image(\"") && output.source.contains(&asset.path),
        "the decorative shadow asset must be emitted as a PDF artifact: {}",
        output.source
    );
    assert_eq!(
        output.source.matches("rgb(0, 0, 0, ").count(),
        0,
        "a blurred shadow must not retain any flat-alpha ring: {}",
        output.source
    );
}

/// One crisp duplicate of the shadow's silhouette, in points.
struct CrispShadowLayer {
    dx: f64,
    dy: f64,
    width: f64,
    height: f64,
}

/// The length `key` introduces in `fragment`, in points, or `None` when
/// `fragment` does not carry that key.
///
/// Read by key rather than by position: a direction that is a multiple of 90
/// degrees leaves a cosine residue of about 1e-17 in the offset, which no
/// exact literal can spell, and an outline writes a variable number of lengths.
fn optional_pt(fragment: &str, key: &str) -> Option<f64> {
    let start: usize = fragment.find(key)? + key.len();
    let rest: &str = &fragment[start..];
    let end: usize = rest.find("pt")?;
    rest[..end].trim().parse::<f64>().ok()
}

fn read_pt(fragment: &str, key: &str) -> f64 {
    optional_pt(fragment, key)
        .unwrap_or_else(|| panic!("no `{key}` length in crisp shadow layer: {fragment}"))
}

/// Every crisp shadow duplicate in `source`.
fn crisp_shadow_layers(source: &str) -> Vec<CrispShadowLayer> {
    let layers: Vec<CrispShadowLayer> = source
        .lines()
        .filter(|line| line.contains("rgb(0, 0, 0, "))
        .map(|line| CrispShadowLayer {
            dx: read_pt(line, "dx: "),
            dy: read_pt(line, "dy: "),
            width: read_pt(line, "width: "),
            height: read_pt(line, "height: "),
        })
        .collect();
    assert!(
        !layers.is_empty(),
        "no crisp shadow layer in output: {source}"
    );
    layers
}

/// The `(dx, dy, width, height)` of the first crisp shadow layer in `source`, in
/// points.
fn first_shadow_layer_geometry(source: &str) -> (f64, f64, f64, f64) {
    let layer: &CrispShadowLayer = &crisp_shadow_layers(source)[0];
    (layer.dx, layer.dy, layer.width, layer.height)
}

#[test]
fn test_shadow_silhouette_outsets_by_half_the_outline_width() {
    // PowerPoint casts `a:outerShdw` from the *stroked* silhouette — the fill
    // path grown by half the line width — and only then offsets it by `dist`.
    // Measured on `customGeo.pptx` page 46 against a native macOS PowerPoint
    // export: the banner's fill path is x [60, 672], y [36, 113.75] under a
    // 3pt outline, and the export's flattened shadow bitmap puts its
    // half-alpha silhouette at x [58.44, 673.56], y [36.00, 116.88] — the fill
    // path outset by 1.5pt and then moved 1.57pt down, agreeing on every edge
    // to within 0.09pt (issue #1057).
    use crate::ir::Shadow;

    let elem = FixedElement {
        x: 60.0,
        y: 36.0,
        width: 612.0,
        height: 77.75,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(79, 129, 189)),
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 3.0,
                color: Color::new(255, 255, 255),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            rotation_deg: None,
            opacity: None,
            // Blur left off so the silhouette itself is what the geometry
            // below measures; the continuous blur contract is covered by
            // `test_shape_shadow_blur_uses_one_continuous_gaussian_asset`.
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 1.57,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.38,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source = generate_typst(&doc).unwrap().source;

    let outset: f64 = 1.5;
    let (dx, dy, width, height) = first_shadow_layer_geometry(&source);
    assert!(
        (width - (612.0 + 2.0 * outset)).abs() < 0.01
            && (height - (77.75 + 2.0 * outset)).abs() < 0.01,
        "a 3pt outline must grow the shadow silhouette by {outset}pt a side, \
         got {width}x{height}pt: {source}"
    );
    // The silhouette grows about its own centre, so the outset shifts the
    // placement back by the same amount before `dist` moves it. Against the
    // shape's own x=60, y=36 that puts the silhouette at x [58.5, 673.5],
    // y [36.07, 116.82] — the export's, to within 0.09pt.
    assert!(
        (dx - -outset).abs() < 0.01 && (dy - (1.57 - outset)).abs() < 0.01,
        "the grown silhouette must stay centred on the shape, got dx {dx}pt \
         dy {dy}pt: {source}"
    );
}

#[test]
fn test_shadow_silhouette_ignores_outline_when_shape_has_none() {
    // The control for the outset above: a shape with no `a:ln` casts from its
    // fill path alone, so the crisp duplicate keeps the shape's own size and
    // sits exactly `dist` away (issue #1057).
    use crate::ir::Shadow;

    let elem = FixedElement {
        x: 60.0,
        y: 36.0,
        width: 612.0,
        height: 77.75,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(79, 129, 189)),
            gradient_fill: None,
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 1.57,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.38,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source = generate_typst(&doc).unwrap().source;

    let (dx, dy, width, height) = first_shadow_layer_geometry(&source);
    assert!(
        (width - 612.0).abs() < 0.01 && (height - 77.75).abs() < 0.01,
        "an outline-less shape casts its fill path unchanged, got \
         {width}x{height}pt: {source}"
    );
    assert!(
        dx.abs() < 0.01 && (dy - 1.57).abs() < 0.01,
        "an outline-less shadow sits exactly `dist` away, got dx {dx}pt \
         dy {dy}pt: {source}"
    );
}

/// `customGeo.pptx` page 46's title banner with its native theme shadow.
fn banner_shadow_output(stroke: Option<BorderSide>, rounded: bool) -> TypstOutput {
    use crate::ir::Shadow;

    let elem = FixedElement {
        x: 60.0,
        y: 36.0,
        width: 612.0,
        height: 77.75,
        kind: FixedElementKind::Shape(Shape {
            kind: if rounded {
                ShapeKind::RoundedRectangle {
                    radius_fraction: 50.0 / 77.75,
                }
            } else {
                ShapeKind::Rectangle
            },
            fill: Some(Color::new(79, 129, 189)),
            gradient_fill: None,
            pattern_fill: None,
            stroke,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 40000.0 / 12700.0,
                distance: 1.57,
                direction: 90.0,
                color: Color::new(0, 0, 0),
                opacity: 0.38,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    generate_typst(&doc).unwrap()
}

fn banner_outline(join: LineJoin) -> Option<BorderSide> {
    Some(BorderSide {
        width: 3.0,
        color: Color::new(255, 255, 255),
        style: BorderLineStyle::Solid,
        join,
        cap: LineCap::Flat,
    })
}

#[test]
fn test_blurred_shadow_filters_the_actual_stroked_silhouette() {
    // The filtered source is the stroked shape itself: its 3pt outline grows
    // the silhouette by 1.5pt, and the declared join still decides how the
    // corner turns before the Gaussian is applied (#1057, #1090, #1138,
    // #1204, #1309).
    for (join, expected) in [(LineJoin::Round, "round"), (LineJoin::Miter, "miter")] {
        let output = banner_shadow_output(banner_outline(join), false);
        assert_eq!(output.images.len(), 1);
        let svg = std::str::from_utf8(&output.images[0].data).unwrap();
        assert!(svg.contains("<feGaussianBlur"), "{svg}");
        assert!(svg.contains("stroke-width=\"3\""), "{svg}");
        assert!(
            svg.contains(&format!("stroke-linejoin=\"{expected}\"")),
            "{svg}"
        );
        assert!(svg.contains("opacity=\"0.38\""), "{svg}");
        assert!(
            !svg.contains("fill-opacity") && !svg.contains("stroke-opacity"),
            "group opacity must not compound where the fill and stroke overlap: {svg}"
        );
    }

    let bare = banner_shadow_output(None, false);
    let svg = std::str::from_utf8(&bare.images[0].data).unwrap();
    assert!(svg.contains("stroke=\"none\""), "{svg}");
}

#[test]
fn test_blurred_rounded_rectangle_keeps_its_native_arc_in_the_filter_source() {
    let output = banner_shadow_output(None, true);
    let svg = std::str::from_utf8(&output.images[0].data).unwrap();
    assert!(
        svg.contains("<rect") && svg.contains("rx=\"38.875\"") && svg.contains("ry=\"38.875\""),
        "the radius is clamped to half the 77.75pt height: {svg}"
    );
}

#[test]
fn test_shape_shadow_without_blur_keeps_single_duplicate() {
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
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 3.0,
                direction: 45.0,
                color: Color::new(0, 0, 0),
                opacity: 0.5,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source = generate_typst(&doc).unwrap().source;

    assert_eq!(
        source.matches("rgb(0, 0, 0, 128)").count(),
        1,
        "zero blur keeps the single offset duplicate: {source}"
    );
}

// ── Polygon shadow silhouette tests (issue #1206) ─────────────────

/// Every point the first crisp shadow outline in `source` draws, in page
/// coordinates: the layer writes its outline relative to its own `#place`, so
/// the placement has to be added back before any of it can be compared to
/// the shape it sits under.
fn first_crisp_shadow_points(source: &str) -> Vec<(f64, f64)> {
    let line: &str = source
        .lines()
        .find(|line| line.contains("rgb(0, 0, 0, "))
        .unwrap_or_else(|| panic!("no crisp shadow outline in output: {source}"));
    let (dx, dy): (f64, f64) = (read_pt(line, "dx: "), read_pt(line, "dy: "));
    let body: &str = &line[line
        .find("curve.move(")
        .unwrap_or_else(|| panic!("the shadow should follow an outline, not a box: {line}"))..];
    // Every length in the curve body is one coordinate of a point, written in
    // order, so reading the `pt` lengths off pairs them up.
    let lengths: Vec<f64> = body
        .split("pt")
        .filter_map(|fragment| {
            let start: usize = fragment
                .rfind(|character: char| {
                    !character.is_ascii_digit() && character != '.' && character != '-'
                })
                .map_or(0, |index| index + 1);
            fragment[start..].parse::<f64>().ok()
        })
        .collect();
    let points: Vec<(f64, f64)> = lengths
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| (dx + pair[0], dy + pair[1]))
        .collect();
    assert!(!points.is_empty(), "no shadow coordinates in: {line}");
    points
}

fn triangle_shadow_output(stroke: Option<BorderSide>, blur_radius: f64) -> TypstOutput {
    use crate::ir::Shadow;

    let elem = FixedElement {
        x: 0.0,
        y: 0.0,
        width: 120.0,
        height: 240.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
            },
            fill: Some(Color::new(192, 0, 0)),
            gradient_fill: None,
            pattern_fill: None,
            stroke,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius,
                distance: 0.0,
                direction: 0.0,
                color: Color::new(0, 0, 0),
                opacity: 1.0,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    generate_typst(&doc).unwrap()
}

fn triangle_shadow_source(stroke: Option<BorderSide>, blur_radius: f64) -> String {
    triangle_shadow_output(stroke, blur_radius).source
}

/// The defect #1206 reports. A ring is the silhouette pushed out by a fixed
/// distance, so a sharp apex has to rise by `outset / sin(half-angle)` — far
/// enough for both slanted edges to clear it. Scaling the vertices onto an
/// expanded bounding box lifts it by the outset alone, which is 4x short on
/// this triangle, and moves the slanted edges by less than the outset while
/// the horizontal base moves by exactly it.
#[test]
fn test_polygon_shadow_offsets_its_outline_instead_of_scaling_it() {
    let source: String = triangle_shadow_source(
        Some(BorderSide {
            width: 8.0,
            color: Color::new(0, 0, 0),
            style: BorderLineStyle::Solid,
            join: LineJoin::Miter,
            cap: LineCap::Flat,
        }),
        0.0,
    );
    let points: Vec<(f64, f64)> = first_crisp_shadow_points(&source);
    let outset: f64 = 4.0;
    let half_angle: f64 = (60.0_f64).atan2(240.0);

    let top: f64 = points.iter().fold(f64::MAX, |top, point| top.min(point.1));
    let expected_top: f64 = -outset / half_angle.sin();
    assert!(
        (top - expected_top).abs() < 0.05,
        "apex at {top:.3}pt, expected {expected_top:.3}pt; a scale would say {:.3}pt",
        -outset,
    );

    let bottom: f64 = points
        .iter()
        .fold(f64::MIN, |bottom, point| bottom.max(point.1));
    assert!(
        (bottom - (240.0 + outset)).abs() < 0.05,
        "the flat base moves by the outset itself, not {:.3}pt",
        bottom - 240.0,
    );
}

/// A blurred polygon keeps its own silhouette in one filtered SVG while the
/// visible shape remains a Typst polygon.
#[test]
fn test_polygon_shadow_uses_its_outline_as_the_gaussian_source() {
    let output = triangle_shadow_output(None, 12.0);
    let svg = std::str::from_utf8(&output.images[0].data).unwrap();
    assert!(
        svg.contains("<feGaussianBlur") && svg.contains("<polygon points=\""),
        "the filter must blur the polygon itself: {svg}",
    );
    assert!(
        output.source.contains("#polygon("),
        "the shape itself still draws as a polygon: {}",
        output.source,
    );
}

/// The SVG viewport follows the Gaussian to 2.6 sigma on every side without
/// scaling or translating the polygon inside that padding.
#[test]
fn test_blurred_polygon_asset_keeps_the_declared_gaussian_reach() {
    let output = triangle_shadow_output(None, 30.0);
    let svg = std::str::from_utf8(&output.images[0].data).unwrap();
    assert!(
        svg.contains("width=\"172\"") && svg.contains("height=\"292\""),
        "120x240pt plus 26pt of 2.6-sigma reach each side: {svg}",
    );
    assert!(
        svg.contains("points=\"86,26 146,266 26,266\""),
        "the polygon must be translated by padding, not scaled: {svg}",
    );
}

/// A `custGeom` shadow follows the same offset. Its hole is a hole in the
/// material too, so dilating the shape has to shrink it rather than grow it.
#[test]
fn test_path_shadow_shrinks_a_hole_while_it_grows_the_outline() {
    use crate::ir::{Shadow, Subpath};

    let elem = FixedElement {
        x: 0.0,
        y: 0.0,
        width: 200.0,
        height: 200.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Path {
                subpaths: vec![
                    Subpath::closed_outline(vec![(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)]),
                    Subpath::closed_outline(vec![
                        (0.25, 0.25),
                        (0.75, 0.25),
                        (0.75, 0.75),
                        (0.25, 0.75),
                    ]),
                ],
            },
            fill: Some(Color::new(192, 0, 0)),
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 10.0,
                color: Color::new(0, 0, 0),
                style: BorderLineStyle::Solid,
                join: LineJoin::Miter,
                cap: LineCap::Flat,
            }),
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 0.0,
                distance: 0.0,
                direction: 0.0,
                color: Color::new(0, 0, 0),
                opacity: 1.0,
            }),
            top_bevel: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let source: String = generate_typst(&doc).unwrap().source;
    let points: Vec<(f64, f64)> = first_crisp_shadow_points(&source);
    let outset: f64 = 5.0;

    let outer_left: f64 = points
        .iter()
        .fold(f64::MAX, |left, point| left.min(point.0));
    assert!(
        (outer_left + outset).abs() < 0.05,
        "the outline grows to {outer_left:.3}pt, expected {:.3}pt",
        -outset,
    );
    // The hole's own left edge sits at 50pt; a shrinking hole moves it in.
    let hole_left: f64 = points
        .iter()
        .filter(|point| point.0 > 0.0 && point.0 < 100.0)
        .fold(f64::MAX, |left, point| left.min(point.0));
    assert!(
        (hole_left - (50.0 + outset)).abs() < 0.05,
        "the hole's edge is at {hole_left:.3}pt, expected {:.3}pt",
        50.0 + outset,
    );
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn test_blurred_polygon_shadow_compiles_to_pdf() {
    let output = triangle_shadow_output(
        Some(BorderSide {
            width: 3.0,
            color: Color::new(0, 0, 0),
            style: BorderLineStyle::Solid,
            join: LineJoin::Round,
            cap: LineCap::Flat,
        }),
        24.0,
    );
    let pdf =
        crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
            .expect("a filtered polygon shadow should compile as Typst");
    assert!(pdf.starts_with(b"%PDF"));
}

// ── Inline text box codegen tests (issue #1690) ──

/// A `wp:inline` text box is one item on its paragraph's own line: a Typst
/// `box` written between the anchor paragraph's runs, so the line grows to hold
/// it and the text after it keeps the same baseline.
#[test]
fn test_inline_text_box_is_emitted_inside_its_anchor_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Anchor paragraph with a text box".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            },
            Run {
                text: String::new(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: Some(Box::new(InlineTextBox {
                    content: vec![
                        Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Box text".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                                inline_box: None,
                            }],
                        }),
                        Block::Table(Table {
                            rows: vec![TableRow {
                                cells: vec![
                                    TableCell {
                                        content: vec![Block::Paragraph(Paragraph {
                                            style: ParagraphStyle::default(),
                                            runs: vec![Run {
                                                text: "Table cell A".to_string(),
                                                style: TextStyle::default(),
                                                href: None,
                                                footnote: None,
                                                inline_box: None,
                                            }],
                                        })],
                                        ..TableCell::default()
                                    },
                                    TableCell {
                                        content: vec![Block::Paragraph(Paragraph {
                                            style: ParagraphStyle::default(),
                                            runs: vec![Run {
                                                text: "Table cell B".to_string(),
                                                style: TextStyle::default(),
                                                href: None,
                                                footnote: None,
                                                inline_box: None,
                                            }],
                                        })],
                                        ..TableCell::default()
                                    },
                                ],
                                height: None,
                                minimum_height: None,
                            }],
                            column_widths: vec![48.0, 48.0],
                            ..Table::default()
                        }),
                        Block::Image(ImageData {
                            rotation_deg: None,
                            flip_h: false,
                            flip_v: false,
                            data: vec![0x89, 0x50, 0x4E, 0x47],
                            format: ImageFormat::Png,
                            width: Some(24.0),
                            height: Some(24.0),
                            crop: None,
                            stroke: None,
                            alignment: None,
                            clip_shape: None,
                            shadow: None,
                            paragraph_spacing: None,
                        }),
                    ],
                    width: 144.0,
                    height: 36.0,
                    padding: Insets {
                        top: 3.6,
                        right: 7.2,
                        bottom: 3.6,
                        left: 7.2,
                    },
                    stroke: Some(BorderSide {
                        width: 0.75,
                        color: Color { r: 0, g: 0, b: 0 },
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                        cap: LineCap::Flat,
                    }),
                    fill: Some(Color {
                        r: 0xFF,
                        g: 0xFF,
                        b: 0xFF,
                    }),
                })),
            },
        ],
    })])]);

    let source = generate_typst(&doc).unwrap().source;

    // Nothing separates the anchor's text from the box. A paragraph break there
    // would put the box on a line of its own, which is the defect (#1690).
    assert!(
        source.contains("Anchor paragraph with a text box#box("),
        "the box must sit on the anchor's line, got:\n{source}"
    );

    assert!(
        source.contains("#box(width: 144pt, height: 36pt"),
        "the box takes the drawing's extent, got:\n{source}"
    );
    assert!(
        source.contains("inset: (top: 3.6pt, right: 7.2pt, bottom: 3.6pt, left: 7.2pt)"),
        "the box takes its `wps:bodyPr` insets, got:\n{source}"
    );
    assert!(
        source.contains("thickness: 0.75pt"),
        "the box strokes its `a:ln`, got:\n{source}"
    );
    assert!(
        source.contains("fill: rgb(255, 255, 255)"),
        "the box fills its `a:solidFill`, got:\n{source}"
    );
    // Typst aligns an inline box on the first baseline of its in-flow content,
    // which leaves the box hanging below the line rather than standing on it.
    // Placing the content out of flow makes the box's own bottom edge its
    // baseline, which is where Word rests it.
    assert!(
        source.contains("#place(top + left)[#block(width: 129.6"),
        "the box's content is placed out of flow at the inner width, got:\n{source}"
    );
    assert!(
        source.contains("Box text"),
        "the box's text is emitted inside it, got:\n{source}"
    );
    let inline_box_start: usize = source
        .find("#box(width: 144pt, height: 36pt")
        .expect("the inline text box is emitted");
    let table_start: usize = source
        .find("#table(")
        .expect("the inline text box's table is emitted");
    assert!(
        table_start > inline_box_start,
        "the table is emitted inside the inline text box, got:\n{source}"
    );
    assert!(
        source.contains("Table cell A") && source.contains("Table cell B"),
        "both table cells remain readable inside the box, got:\n{source}"
    );
    assert!(
        source.contains("#image(\"img-0.png\""),
        "the inline text box's picture is emitted with the rest of its flow, got:\n{source}"
    );
}

/// The box travels with the run, so it reaches body, table-cell, list and
/// tabbed-paragraph paths. Carrying its markup beside the runs left it out of
/// each path, where a run with no text of its own emits nothing (issues #1690,
/// #1889).
#[test]
fn test_inline_text_box_survives_a_tab_and_a_table_cell() {
    let nested_table = || {
        Block::Table(Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Nested table cell".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                            inline_box: None,
                        }],
                    })],
                    ..TableCell::default()
                }],
                height: None,
                minimum_height: None,
            }],
            column_widths: vec![72.0],
            ..Table::default()
        })
    };
    let anchor_runs = |leading: &str| {
        vec![
            Run {
                text: leading.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            },
            Run {
                text: String::new(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: Some(Box::new(InlineTextBox {
                    content: vec![
                        Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Box text".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                                inline_box: None,
                            }],
                        }),
                        nested_table(),
                    ],
                    width: 144.0,
                    height: 36.0,
                    padding: Insets::default(),
                    stroke: None,
                    fill: None,
                })),
            },
        ]
    };

    // A tab splits the paragraph's runs into fresh segments before they are
    // measured and written.
    let tabbed = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: anchor_runs("Label:\t"),
    })])]);
    let tabbed_source = generate_typst(&tabbed).unwrap().source;
    assert!(
        tabbed_source.contains("#box(width: 144pt, height: 36pt"),
        "a tabbed paragraph still draws its box, got:\n{tabbed_source}"
    );
    assert!(
        tabbed_source.contains("Box text"),
        "a tabbed paragraph still writes the box's text, got:\n{tabbed_source}"
    );
    let tabbed_box_start: usize = tabbed_source
        .find("#box(width: 144pt, height: 36pt")
        .expect("the tabbed box is emitted");
    let tabbed_box_end: usize = tabbed_source[tabbed_box_start..]
        .find("]]]")
        .map(|offset| tabbed_box_start + offset)
        .expect("the tabbed box closes");
    assert!(
        tabbed_source[tabbed_box_start..tabbed_box_end].contains("#table(")
            && tabbed_source[tabbed_box_start..tabbed_box_end].contains("Nested table cell"),
        "the tabbed box keeps its nested table inside its bounds, got:\n{tabbed_source}"
    );

    // A table cell's paragraphs go through the table generator, not the body's.
    let celled = make_doc(vec![make_flow_page(vec![Block::Table(Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: anchor_runs("In a cell"),
                })],
                ..TableCell::default()
            }],
            height: None,
            minimum_height: None,
        }],
        column_widths: vec![300.0],
        ..Table::default()
    })])]);
    let celled_source = generate_typst(&celled).unwrap().source;
    assert!(
        celled_source.contains("#box(width: 144pt, height: 36pt"),
        "a table cell still draws its box, got:\n{celled_source}"
    );
    assert!(
        celled_source.contains("Box text"),
        "a table cell still writes the box's text, got:\n{celled_source}"
    );
    let celled_box_start: usize = celled_source
        .find("#box(width: 144pt, height: 36pt")
        .expect("the cell's inline box is emitted");
    let celled_box_end: usize = celled_source[celled_box_start..]
        .find("]]]")
        .map(|offset| celled_box_start + offset)
        .expect("the cell's inline box closes");
    assert!(
        celled_source[celled_box_start..celled_box_end].contains("#table(")
            && celled_source[celled_box_start..celled_box_end].contains("Nested table cell"),
        "the cell's inline box keeps its nested table inside its bounds, got:\n{celled_source}"
    );

    let listed = make_doc(vec![make_flow_page(vec![Block::List(List {
        kind: ListKind::Unordered,
        items: vec![crate::ir::ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: anchor_runs("List item: "),
            }],
            level: 0,
            start_at: None,
        }],
        level_styles: Default::default(),
    })])]);
    let listed_source = generate_typst(&listed).unwrap().source;
    let listed_box_start: usize = listed_source
        .find("#box(width: 144pt, height: 36pt")
        .expect("the list item's inline box is emitted");
    let listed_box_end: usize = listed_source[listed_box_start..]
        .find("]]]")
        .map(|offset| listed_box_start + offset)
        .expect("the list item's inline box closes");
    assert!(
        listed_source[listed_box_start..listed_box_end].contains("#table(")
            && listed_source[listed_box_start..listed_box_end].contains("Nested table cell"),
        "the list item's inline box keeps its nested table inside its bounds, got:\n{listed_source}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn footnote_inline_text_box_preserves_nested_table_text() {
    for (anchor, label) in [("Source", "ReferenceCell"), ("Source:\t", "AppendixCell")] {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![make_paragraph(label)],
                    ..TableCell::default()
                }],
                height: None,
                minimum_height: None,
            }],
            column_widths: vec![120.0],
            ..Table::default()
        };
        let note = Run {
            text: String::new(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
            inline_box: Some(Box::new(InlineTextBox {
                content: vec![Block::Table(table)],
                width: 144.0,
                height: 36.0,
                padding: Insets::default(),
                stroke: None,
                fill: None,
            })),
        };
        let document = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: anchor.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                    inline_box: None,
                },
                Run {
                    text: String::new(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: Some(vec![note]),
                    inline_box: None,
                },
            ],
        })])]);
        let generated = generate_typst(&document).expect("footnote box generation succeeds");
        let runs = crate::render::pdf::compiled_text_runs_with_images(
            &generated.source,
            &generated.images,
            0,
        )
        .expect("footnote table compiles");
        let text: String = runs.iter().map(|run| run.text.as_str()).collect();
        assert!(
            text.contains(label),
            "footnote table text remains searchable: {text}"
        );
    }
}
