use super::*;
use crate::ir::{
    ChartSeries, ColumnLayout, GradientStop, HeaderFooterParagraph, ImageData, ListItem, ListKind,
    ListLevelStyle, Metadata, SmartArtNode, StyleSheet,
};
use crate::render::typst_gen::shapes::{SHADOW_BLUR_EXTENT_SIGMA, shadow_blur_sigma};
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
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content,
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
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

/// The tracked Noto Sans CJK SC face's family, and its `hhea` ascender and
/// descender in em (1160 and -288 of 1000 units). It declares no line gap,
/// which [`noto_sans_cjk_context_with_line_gap`] rewrites.
const NOTO_SANS_CJK_SC: &str = "Noto Sans CJK SC";
const NOTO_SANS_CJK_ASCENDER_EM: f64 = 1.160;
const NOTO_SANS_CJK_DESCENDER_EM: f64 = 0.288;

/// A font context whose only caller-provided face is the tracked Noto Sans
/// CJK SC with its `hhea` line gap rewritten to `line_gap` units — the one
/// factor Batang, Gulim, Dotum and Gungsuh (152/1024) differ from Malgun
/// Gothic (0) by. Held in memory, it outranks any installed copy.
fn noto_sans_cjk_context_with_line_gap(
    line_gap: i16,
) -> crate::render::font_context::FontSearchContext {
    let base: &[u8] = include_bytes!("../../fonts/NotoSansCJKsc-GB2312.otf");
    let face: Vec<u8> = crate::test_support::make_face_with_hhea_line_gap(base.to_vec(), line_gap);
    let fonts = crate::render::pdf::load_fonts_from_bytes([face.as_slice()]);
    crate::render::font_context::resolve_font_search_context_from_fonts(&fonts)
}

/// The `(top_edge_em, bottom_edge_em)` of the first line box the generator
/// emits, or `None` when the source declares no fixed text edges.
fn emitted_line_box_em(source: &str) -> Option<(f64, f64)> {
    emitted_line_box(source, "em")
}

/// The same box for a slide's text, which states its edges in points so they
/// cannot drift with the size in force where the rule lands (issue #1115).
/// `font_size_pt` is the size the paragraph declares, which is what the box was
/// derived from.
fn emitted_slide_line_box_em(source: &str, font_size_pt: f64) -> Option<(f64, f64)> {
    let (top_pt, bottom_pt) = emitted_line_box(source, "pt")?;
    Some((top_pt / font_size_pt, bottom_pt / font_size_pt))
}

fn emitted_line_box(source: &str, unit: &str) -> Option<(f64, f64)> {
    let after_top: &str = source.split_once("top-edge: ")?.1;
    let (top, rest) = after_top.split_once(unit)?;
    let after_bottom: &str = rest.split_once("bottom-edge: -")?.1;
    let (bottom, _) = after_bottom.split_once(unit)?;
    Some((top.parse().ok()?, bottom.parse().ok()?))
}

/// Assert the generated line box spans `expected_pt` and seats the baseline
/// `hhea ascender + lineGap` below its top — a constant that does not scale
/// with the box, so every point the line gains over the font's own metric line
/// falls below the baseline (issues #508, #518). `east_asian_excess_em` is the
/// extra ascent Word adds for a line carrying East Asian text; pass 0 for a
/// Latin line. Compared numerically rather than as a formatted string so float
/// noise in the em split cannot break the assertion.
fn assert_line_advance(
    source: &str,
    family: &str,
    font_size: f64,
    expected_pt: f64,
    east_asian_excess_em: f64,
) {
    let (top, bottom) =
        emitted_line_box_em(source).unwrap_or_else(|| panic!("no line box emitted in: {source}"));
    let advance_pt: f64 = (top + bottom) * font_size;
    assert!(
        (advance_pt - expected_pt).abs() < 0.01,
        "line advance {advance_pt}pt should be {expected_pt}pt in: {source}"
    );
    let (ascender, _descender, _) =
        crate::render::pdf::font_line_metrics_em(family).expect("font metrics should resolve");
    let expected_top: f64 = ascender + east_asian_excess_em;
    assert!(
        (top - expected_top).abs() < 0.001,
        "baseline should sit {expected_top}em below the box top, not {top}em: {source}"
    );
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
            auto_fit: false,
            text_rotation_deg: None,
            shape_rotation_deg: None,
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
            pattern_fill: None,
            stroke,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
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
            auto_fit: false,
            text_rotation_deg: None,
            shape_rotation_deg: None,
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
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            data: vec![0x89, 0x50, 0x4E, 0x47], // PNG header stub
            format,
            width: Some(w),
            height: Some(h),
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
        }),
    }
}

#[path = "typst_gen_fixed_page_tests.rs"]
mod fixed_page_tests;

#[path = "typst_gen_fixed_page_textbox_tests.rs"]
mod fixed_page_textbox_tests;

#[cfg(not(target_arch = "wasm32"))]
#[path = "typst_gen_powerpoint_kerning_tests.rs"]
mod powerpoint_kerning_tests;

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
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })
}

/// Helper to create a simple Table with text cells.
fn make_simple_table(rows: Vec<Vec<&str>>) -> Table {
    Table {
        rows: rows
            .into_iter()
            .map(|cells| TableRow {
                minimum_height: None,
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

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_generate_run_baseline_shift_moves_text_by_its_run_size() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "A".to_string(),
                style: TextStyle {
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "1".to_string(),
                style: TextStyle {
                    font_size: Some(10.0),
                    baseline_shift: Some(BaselineShiftEm(0.3)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "B".to_string(),
                style: TextStyle {
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "2".to_string(),
                style: TextStyle {
                    font_size: Some(10.0),
                    baseline_shift: Some(BaselineShiftEm(-0.25)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let output = generate_typst(&doc).unwrap();
    let placed = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
    let baseline = |needle: &str| -> f64 {
        placed
            .iter()
            .find(|run| run.text == needle)
            .unwrap_or_else(|| panic!("missing {needle:?} in {placed:?}"))
            .baseline_pt
    };
    let body_baseline = baseline("A");

    let second_body_baseline: f64 = baseline("B");
    let superscript_baseline: f64 = baseline("1");
    let subscript_baseline: f64 = baseline("2");
    assert!(
        (second_body_baseline - body_baseline).abs() < 0.01,
        "body baselines differ: {body_baseline} and {second_body_baseline}; {placed:?}\n{}",
        output.source
    );
    assert!(
        (body_baseline - superscript_baseline - 3.0).abs() < 0.01,
        "superscript baseline {superscript_baseline} should be 3pt above {body_baseline}; {placed:?}"
    );
    assert!(
        (subscript_baseline - body_baseline - 2.5).abs() < 0.01,
        "subscript baseline {subscript_baseline} should be 2.5pt below {body_baseline}; {placed:?}"
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

/// A run whose fill declares an opacity reaches the page with its alpha
/// channel, so it composites against the backdrop instead of printing as solid
/// ink (issue #1121).
#[test]
fn test_generate_run_fill_alpha() {
    let source = generated_source_for_run_style(TextStyle {
        color: Some(Color::new(0, 0, 0)),
        color_alpha: Some(0.5),
        ..TextStyle::default()
    });
    assert!(
        source.contains("fill: rgb(0, 0, 0, 128)"),
        "A half-opacity run should emit a 4-argument rgb. Got: {source}"
    );
}

/// Triangulation for [`test_generate_run_fill_alpha`]: a different opacity on a
/// different colour scales the alpha channel with it.
#[test]
fn test_generate_run_fill_alpha_quarter() {
    let source = generated_source_for_run_style(TextStyle {
        color: Some(Color::new(255, 0, 0)),
        color_alpha: Some(0.25),
        ..TextStyle::default()
    });
    assert!(
        source.contains("fill: rgb(255, 0, 0, 64)"),
        "A quarter-opacity run should emit a 4-argument rgb. Got: {source}"
    );
}

/// A run that states no opacity keeps the 3-argument form, so no existing
/// output gains a redundant alpha channel.
#[test]
fn test_generate_run_without_fill_alpha_stays_opaque() {
    let source = generated_source_for_run_style(TextStyle {
        color: Some(Color::new(0, 0, 0)),
        color_alpha: None,
        ..TextStyle::default()
    });
    assert!(
        source.contains("fill: rgb(0, 0, 0)"),
        "An opaque run should emit a 3-argument rgb. Got: {source}"
    );
}

/// The Typst source for a one-run paragraph carrying `style`.
fn generated_source_for_run_style(style: TextStyle) -> String {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Sensitivity: Internal".to_string(),
            style,
            href: None,
            footnote: None,
        }],
    })])]);
    generate_typst(&doc).unwrap().source
}

#[test]
fn test_table_cell_vertical_align_center() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
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
            minimum_height: None,
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

// ── Continuous shadow blur (issues #390, #662, #784, #1309) ──────────
//
// PowerPoint renders `blurRad` as a Gaussian alpha ramp centred on the
// shadow silhouette with a std-dev of blurRad/3 (probe-fitted from native
// exports at blur 1-18.9pt, issue #784). The generated SVG follows that
// continuous filter to 2.6 sigma, where the two-sided tail is about 0.9%.

fn shadow_with(blur_radius: f64, opacity: f64) -> Shadow {
    Shadow {
        blur_radius,
        distance: 3.0,
        direction: 90.0,
        color: Color { r: 0, g: 0, b: 0 },
        opacity,
    }
}

#[test]
fn test_zero_blur_shadow_keeps_its_declared_alpha_byte() {
    assert_eq!(shadow_alpha(&shadow_with(0.0, 0.4)), 102);
}

#[test]
fn test_blur_sigma_is_a_third_of_the_declared_radius() {
    // Native PowerPoint rasterises an `outerShdw` as a Gaussian whose
    // std-dev is blurRad/3: one-factor probe exports of customGeo.pptx at
    // blurRad 1/3.15/6.3/12.6/18.9pt fit sigma/blurRad = 0.331-0.345 on
    // every silhouette edge of the flattened shadow bitmap (issue #784).
    // The 0.3 this replaced sat below that band at every radius, which cut
    // both the ramp's reach and its density about 10% short.
    for blur_radius in [40000.0 / 12700.0, 160000.0 / 12700.0] {
        let sigma = shadow_blur_sigma(&shadow_with(blur_radius, 0.38));
        let ratio = sigma / blur_radius;
        assert!(
            (0.32..=0.35).contains(&ratio),
            "sigma/blurRad {ratio} at blur {blur_radius}pt must sit inside \
             the probe-fitted band"
        );
    }
}

#[test]
fn test_blur_asset_reaches_the_declared_gaussian_extent() {
    let sigma = shadow_blur_sigma(&shadow_with(24.0, 0.6));
    let reach = SHADOW_BLUR_EXTENT_SIGMA * sigma;
    assert!(
        (reach - 20.8).abs() < 1e-9,
        "24pt blur should follow its 8pt sigma to 20.8pt, got {reach}"
    );
}

/// Word's East Asian line height follows the face a line is set in, not the
/// script of its characters.
///
/// `03_meeting_minutes_ko` has three Heading2 paragraphs sharing identical
/// `w:pPr` and `w:rPr` — Malgun Gothic in every `w:rFonts` slot — differing
/// only in whether their text is Hangul or Latin. Word gives them the same
/// height (gap from the preceding baseline 29.76 against 29.52); gating the
/// East Asian line box on the text left the Latin-only heading 2.37pt short
/// and dragged the table under it 5.47pt up (issue #643).
#[test]
fn a_latin_line_set_in_a_cjk_face_keeps_the_east_asian_line_box() {
    if crate::render::pdf::font_line_metrics_em("Malgun Gothic").is_none() {
        return; // no Korean face available (e.g. a runner with no CJK fonts)
    }
    let line_box = |text: &str, family: &str| {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some(family.to_string()),
                    east_asian_font_family: Some(family.to_string()),
                    font_size: Some(11.5),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        emitted_line_box_em(&generate_typst(&doc).unwrap().source)
    };

    let korean = line_box("결정 사항", "Malgun Gothic").expect("a Korean line has a line box");
    let latin =
        line_box("Action Items", "Malgun Gothic").expect("a Latin line in a CJK face has one too");
    assert_eq!(
        korean, latin,
        "two lines set in the same face at the same size must share a line box"
    );

    // The face still decides: Word does not give an Arial paragraph the East
    // Asian line even inside a Korean document, and snapping those inflated
    // every Western document by 30-50% (issue #354).
    assert_ne!(
        line_box("Action Items", "Arial"),
        Some(latin),
        "an Arial line must not take the CJK face's line box"
    );
}
