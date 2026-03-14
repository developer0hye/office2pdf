use super::*;
use crate::ir::{
    ChartSeries, ColumnLayout, GradientStop, HeaderFooterParagraph, ImageData, ListItem, ListKind,
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

#[path = "typst_gen_paragraph_tests.rs"]
mod paragraph_tests;

#[path = "typst_gen_table_codegen_tests.rs"]
mod table_codegen_tests;
use self::table_codegen_tests::make_text_cell;

#[path = "typst_gen_image_tests.rs"]
mod image_tests;

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
            fill: None,
            opacity: None,
            stroke: None,
            shape_kind: None,
                    no_wrap: false,
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
            fill: None,
            opacity: None,
            stroke: None,
            shape_kind: None,
                    no_wrap: false,
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
            stroke: None,
        }),
    }
}

#[path = "typst_gen_fixed_page_tests.rs"]
mod fixed_page_tests;

#[path = "typst_gen_fixed_page_textbox_tests.rs"]
mod fixed_page_textbox_tests;

// ── SheetPage codegen tests ──────────────────────────────────────────

/// Helper to create a SheetPage.
fn make_sheet_page(name: &str, width: f64, height: f64, margins: Margins, table: Table) -> Page {
    Page::Sheet(crate::ir::SheetPage {
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

#[path = "typst_gen_diagram_visual_tests.rs"]
mod diagram_visual_tests;

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

// ── generate_blocks helper tests ─────────────────────────────────────

#[test]
fn test_generate_blocks_empty_slice_produces_no_output() {
    let blocks: Vec<Block> = vec![];
    let mut out = String::new();
    let mut ctx = GenCtx::new();
    generate_blocks(&mut out, &blocks, &mut ctx).unwrap();
    assert!(
        out.is_empty(),
        "Empty block slice should produce no output. Got: {out:?}"
    );
}

#[test]
fn test_generate_blocks_single_block_no_leading_newline() {
    let blocks: Vec<Block> = vec![make_paragraph("Hello")];
    let mut out = String::new();
    let mut ctx = GenCtx::new();
    generate_blocks(&mut out, &blocks, &mut ctx).unwrap();
    assert!(
        !out.starts_with('\n'),
        "Single block should not start with newline. Got: {out:?}"
    );
    assert!(
        out.contains("Hello"),
        "Output should contain block text. Got: {out:?}"
    );
}

#[test]
fn test_generate_blocks_multiple_blocks_separated_by_newline() {
    let blocks: Vec<Block> = vec![make_paragraph("First"), make_paragraph("Second")];
    let mut out = String::new();
    let mut ctx = GenCtx::new();
    generate_blocks(&mut out, &blocks, &mut ctx).unwrap();
    // The output should contain both paragraphs separated by a newline
    let first_pos: usize = out.find("First").expect("Should contain 'First'");
    let second_pos: usize = out.find("Second").expect("Should contain 'Second'");
    assert!(
        first_pos < second_pos,
        "First should appear before Second. Got: {out:?}"
    );
    // There should be a newline between the two blocks
    let between: &str = &out[first_pos..second_pos];
    assert!(
        between.contains('\n'),
        "Blocks should be separated by newline. Got between: {between:?}"
    );
}

#[test]
fn test_generate_blocks_three_blocks_have_two_separators() {
    let blocks: Vec<Block> = vec![
        make_paragraph("A"),
        make_paragraph("B"),
        make_paragraph("C"),
    ];
    let mut out = String::new();
    let mut ctx = GenCtx::new();
    generate_blocks(&mut out, &blocks, &mut ctx).unwrap();
    assert!(out.contains("A"), "Should contain A. Got: {out:?}");
    assert!(out.contains("B"), "Should contain B. Got: {out:?}");
    assert!(out.contains("C"), "Should contain C. Got: {out:?}");
    // Verify ordering
    let pos_a: usize = out.find("A").expect("A");
    let pos_b: usize = out.find("B").expect("B");
    let pos_c: usize = out.find("C").expect("C");
    assert!(pos_a < pos_b && pos_b < pos_c, "Order should be A < B < C");
}

// ── Font weight inference with fallback tests ────────────────────────

#[test]
fn test_inferred_weight_not_emitted_when_font_unavailable() {
    use crate::render::font_context::FontSearchContext;
    // When "Pretendard ExtraBold" is not available (no font context has it),
    // `weight: "extrabold"` should NOT appear — it blocks fallback fonts.
    let context = FontSearchContext::for_test(Vec::new(), &["Arial"], &[], &[]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap()
    .source;
    assert!(
        !result.contains("weight: \"extrabold\""),
        "Should NOT emit extrabold weight when font is unavailable. Got: {result}"
    );
}

#[test]
fn test_inferred_weight_emitted_when_font_available_via_alias() {
    use crate::render::font_context::FontSearchContext;
    // When "Pretendard" family is available, "Pretendard ExtraBold" should
    // emit weight: "extrabold" so Typst picks the correct variant.
    let context = FontSearchContext::for_test(Vec::new(), &["Pretendard"], &[], &[]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap()
    .source;
    assert!(
        result.contains("weight: \"extrabold\""),
        "Should emit extrabold weight when font is available. Got: {result}"
    );
}

#[test]
fn test_explicit_bold_still_emitted_when_font_unavailable() {
    use crate::render::font_context::FontSearchContext;
    // Explicit bold from PPTX attributes should still be emitted even when
    // the font is unavailable — bold (weight 700) exists in most fonts.
    let context = FontSearchContext::for_test(Vec::new(), &["Arial"], &[], &[]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold text".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap()
    .source;
    assert!(
        result.contains("weight: \"bold\""),
        "Explicit bold should still be emitted. Got: {result}"
    );
    assert!(
        !result.contains("weight: \"extrabold\""),
        "Should use bold, not extrabold (from unavailable font name). Got: {result}"
    );
}
