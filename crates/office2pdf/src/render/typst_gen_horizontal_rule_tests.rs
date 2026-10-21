//! Word's horizontal line, drawn on the line of the paragraph it is the only
//! content of (issue #2251).

#![cfg(not(target_arch = "wasm32"))]

use super::*;
use crate::ir::{HorizontalRule, HorizontalRulePaint, HorizontalRuleWidth};
use crate::render::pdf::{PaintedKind, PaintedPrimitive, compiled_paint_sequence};

const MARGIN_PT: f64 = 72.0;
const MEASURE_PT: f64 = 612.0 - 2.0 * MARGIN_PT;
const LINE_PT: f64 = 20.0;

fn letter_page(content: Vec<Block>) -> Page {
    Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize {
            width: 612.0,
            height: 792.0,
        },
        margins: Margins {
            top: MARGIN_PT,
            right: MARGIN_PT,
            bottom: MARGIN_PT,
            left: MARGIN_PT,
        },
        content,
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
        page_borders: None,
        line_numbering: None,
        vertical_alignment: None,
    })
}

/// A paragraph on an exact 20pt line with no gaps, so every line's top is a
/// whole multiple of the line above the first.
fn exact_line_style() -> ParagraphStyle {
    ParagraphStyle {
        line_spacing: Some(LineSpacing::Exact(LINE_PT)),
        space_before: Some(0.0),
        space_after: Some(0.0),
        ..ParagraphStyle::default()
    }
}

fn text_paragraph(text: &str) -> Block {
    Block::Paragraph(Paragraph {
        style: exact_line_style(),
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(12.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
            inline_box: None,
        }],
    })
}

fn rule_paragraph(rule: HorizontalRule, indent: (f64, f64)) -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_left: Some(indent.0),
            indent_right: Some(indent.1),
            horizontal_rule: Some(Box::new(rule)),
            paragraph_mark_text_style: Some(Box::new(TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(12.0),
                ..TextStyle::default()
            })),
            ..exact_line_style()
        },
        runs: Vec::new(),
    })
}

fn shaded(width: HorizontalRuleWidth, alignment: Alignment) -> HorizontalRule {
    HorizontalRule {
        width,
        height: 1.5,
        alignment,
        paint: HorizontalRulePaint::Shaded,
    }
}

/// The shapes the compiled first page paints.
fn painted_shapes(content: Vec<Block>) -> Vec<PaintedPrimitive> {
    let doc = make_doc(vec![letter_page(content)]);
    let output = generate_typst(&doc).expect("the document generates");
    compiled_paint_sequence(&output.source, &output.images, 0)
        .unwrap_or_else(|error| panic!("{error:?} in {}", output.source))
        .into_iter()
        .filter(|primitive| primitive.kind == PaintedKind::Shape)
        .collect()
}

/// The one shape painted on the page.
fn the_rule(content: Vec<Block>) -> PaintedPrimitive {
    let shapes = painted_shapes(content);
    assert_eq!(shapes.len(), 1, "one rule is painted: {shapes:#?}");
    shapes.into_iter().next().unwrap()
}

fn assert_close(actual: f64, expected: f64, what: &str) {
    assert!(
        (actual - expected).abs() < 0.01,
        "{what}: {actual} should be {expected}"
    );
}

fn assert_spans_x(rule: &PaintedPrimitive, left: f64, right: f64) {
    assert_close(rule.bounds.0, left, "rule left");
    assert_close(rule.bounds.2, right, "rule right");
}

/// Word for Mac paints its shaded rule as a 0.75pt black outline around the
/// rule's box, 0.2pt inside its top and bottom, whatever the `fillcolor`. A
/// 1.5pt rule under 540pt of Arial 12pt text in the header of docxide-pdf's
/// `isla_language_lesson_plan` is stroked at y 67.40–68.50 where its line ends
/// at 68.75 (measured on 1.5pt rules only). The box stands on the bottom of
/// the paragraph's line and spans the paragraph between its indents.
#[test]
fn a_shaded_rule_outlines_the_bottom_of_its_line_across_the_paragraph() {
    let rule = the_rule(vec![
        text_paragraph("Above"),
        rule_paragraph(
            shaded(HorizontalRuleWidth::Fraction(1.0), Alignment::Center),
            (0.0, 0.0),
        ),
        text_paragraph("Below"),
    ]);

    let line_bottom: f64 = MARGIN_PT + 2.0 * LINE_PT;
    assert_spans_x(&rule, MARGIN_PT, MARGIN_PT + MEASURE_PT);
    assert_close(rule.bounds.1, line_bottom - 1.3, "rule top");
    assert_close(rule.bounds.3, line_bottom - 0.2, "rule bottom");
    let stroke = rule.stroke.as_ref().expect("the shaded rule is outlined");
    assert_close(stroke.thickness_pt, 0.75, "outline thickness");
    assert_eq!(
        stroke.color,
        Some(typst::visualize::Color::from_u8(0, 0, 0, 255))
    );
    assert_eq!(rule.rectangle_fill, None, "the outline is hollow");
}

/// `o:hrpct` takes its share of the paragraph's width between its indents,
/// and `o:hralign` seats the rule in what is left over.
#[test]
fn a_partial_rule_takes_its_share_of_the_indented_measure_where_it_is_aligned() {
    let right_half = the_rule(vec![rule_paragraph(
        shaded(HorizontalRuleWidth::Fraction(0.5), Alignment::Right),
        (0.0, 0.0),
    )]);
    assert_spans_x(
        &right_half,
        MARGIN_PT + MEASURE_PT / 2.0,
        MARGIN_PT + MEASURE_PT,
    );

    let (indent_left, indent_right): (f64, f64) = (36.0, 18.0);
    let paragraph_pt: f64 = MEASURE_PT - indent_left - indent_right;
    let centred_quarter = the_rule(vec![rule_paragraph(
        shaded(HorizontalRuleWidth::Fraction(0.25), Alignment::Center),
        (indent_left, indent_right),
    )]);
    let left: f64 = MARGIN_PT + indent_left + paragraph_pt * 0.375;
    assert_spans_x(&centred_quarter, left, left + paragraph_pt * 0.25);

    let stated = the_rule(vec![rule_paragraph(
        shaded(HorizontalRuleWidth::Points(100.0), Alignment::Left),
        (indent_left, indent_right),
    )]);
    assert_spans_x(
        &stated,
        MARGIN_PT + indent_left,
        MARGIN_PT + indent_left + 100.0,
    );
}

/// TODO(unverified noshade paint): no Word export of an `o:hrnoshade` rule is
/// at hand. It is drawn as the VML says, a bar of its full height in its
/// `fillcolor`, standing on the bottom of its line.
#[test]
fn an_unshaded_rule_is_a_solid_bar_in_its_fill() {
    let rule = the_rule(vec![rule_paragraph(
        HorizontalRule {
            width: HorizontalRuleWidth::Fraction(1.0),
            height: 9.0,
            alignment: Alignment::Left,
            paint: HorizontalRulePaint::Solid(Color::new(0x00, 0x80, 0x80)),
        },
        (0.0, 0.0),
    )]);

    assert_spans_x(&rule, MARGIN_PT, MARGIN_PT + MEASURE_PT);
    assert_close(rule.bounds.1, MARGIN_PT + LINE_PT - 9.0, "bar top");
    assert_close(rule.bounds.3, MARGIN_PT + LINE_PT, "bar bottom");
    assert!(rule.stroke.is_none(), "the bar has no outline: {rule:#?}");
    assert_eq!(
        rule.rectangle_fill,
        Some(typst::visualize::Color::from_u8(0x00, 0x80, 0x80, 255))
    );
}

/// TODO(#2251 rules in table cells): Word drew every rule in the autofit
/// table cells of docxide-pdf's `maine_criminal_history_record` as a one-twip
/// sliver, not across the cell, and no export of a rule in another cell is at
/// hand. Until a native probe settles a cell's rule width, a rule in a cell
/// paints nothing and its paragraph keeps the blank line it had before.
#[test]
fn a_rule_in_a_table_cell_paints_nothing_yet() {
    let table = Table {
        rows: vec![TableRow {
            cant_split: false,
            minimum_height: None,
            cells: vec![TableCell {
                content: vec![
                    rule_paragraph(
                        shaded(HorizontalRuleWidth::Fraction(1.0), Alignment::Center),
                        (0.0, 0.0),
                    ),
                    text_paragraph("Cell"),
                ],
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };

    let shapes = painted_shapes(vec![Block::Table(table)]);
    assert!(shapes.is_empty(), "no rule is painted: {shapes:#?}");
}
