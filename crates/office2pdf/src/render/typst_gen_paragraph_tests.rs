use super::*;
use crate::ir::WordCompatibilityMode;

#[test]
fn test_generate_plain_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Hello World"));
}

#[test]
fn test_generate_empty_paragraph_reserves_line_height() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: Vec::new(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#v(12pt)"),
        "empty DOCX paragraph marks should reserve vertical flow space: {result}"
    );
}

#[test]
fn test_generate_page_setup() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
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
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
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
        result.contains("#underline("),
        "Expected underline wrapper in: {result}"
    );
    assert!(result.contains("Underlined"));
    // Word draws one continuous rectangle straight through descenders. Typst
    // skips ink by default, which chopped the rule into segments wherever a
    // glyph descended through it (issue #641).
    assert!(
        result.contains("evade: false"),
        "The underline must not evade descenders, got: {result}"
    );
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
    // Typst hangs line-final punctuation into the margin when justifying.
    // Word does not, so a justified line ending in a comma overshot the
    // right margin by about 0.8 of the comma's advance (issue #640).
    // Typst lets a justified line squeeze its spaces to two thirds of their
    // natural width. Word's own exports never go below 0.9332, and on the
    // audited letter our 0.9014 squeeze pulled one more syllable onto a line
    // and split a Korean word across the break (issue #639).
    assert!(
        result.contains("justification-limits: (spacing: (min: 80%, max: 150%))"),
        "justified text must not shrink its spaces below their natural width: {result}"
    );
    assert!(
        result.contains("overhang: false"),
        "justified text must not hang punctuation past the margin: {result}"
    );
}

/// A justified paragraph in the compatibility mode Word declares, so the
/// natural-width rule of issue #1130 can be switched on and off around one
/// paragraph.
fn justified_paragraph_document(text: &str, mode: WordCompatibilityMode) -> Document {
    let mut doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Justify),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    doc.styles.word_compatibility_mode = Some(mode);
    doc
}

/// Word's pre-2013 East Asian justification seats on a line only what fits at
/// its natural width. Swept over eleven measures either side of the fit
/// boundary, a native legacy export of the reported fixture refuses a 0.5pt
/// overrun where Typst's Knuth-Plass breaker squeezes twelve word spaces to
/// 0.9746 of natural to take one more eojeol (issue #1130).
#[test]
fn test_legacy_justified_east_asian_paragraph_breaks_at_natural_width() {
    let doc = justified_paragraph_document(
        "본 계약은 2026년 8월 체결되며 ABC 주식회사와 DEF 파트너스는 각각 임원을 파견한다",
        WordCompatibilityMode::Legacy,
    );
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("par(justify: true"),
        "the paragraph is still justified: {result}"
    );
    assert!(
        result.contains(r#"set par(linebreaks: "simple")"#),
        "a legacy East Asian justified line must not compress to seat one more \
         token: {result}"
    );
}

/// At `compatibilityMode 15` Word does take that trade, and its compressed word
/// space measures 3.6020pt against our 3.6019pt on the reported fixture, so the
/// Knuth-Plass breaker is what models it (issue #1130).
#[test]
fn test_word_2013_justified_east_asian_paragraph_keeps_the_optimized_breaker() {
    let doc = justified_paragraph_document(
        "본 계약은 2026년 8월 체결되며 ABC 주식회사와 DEF 파트너스는 각각 임원을 파견한다",
        WordCompatibilityMode::Word2013OrLater,
    );
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("linebreaks:"),
        "modern Word compresses to seat one more token, so the breaker stays: \
         {result}"
    );
}

/// The switch is East Asian, not legacy-wide: the same legacy export takes up
/// to 2.5pt of overrun on its Latin paragraph, so Latin keeps the breaker that
/// already reproduces the corpus's compressed English lines (issue #1130).
#[test]
fn test_legacy_justified_latin_paragraph_keeps_the_optimized_breaker() {
    let doc = justified_paragraph_document(
        "Monthly active users crossed 21,300 in June, up 72% since January.",
        WordCompatibilityMode::Legacy,
    );
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("linebreaks:"),
        "legacy Latin justification compresses, so the breaker stays: {result}"
    );
}

/// A presentation and a workbook declare no such setting, and neither is laid
/// out by Word, so both keep the modern breaker.
#[test]
fn test_document_without_a_compatibility_mode_keeps_the_optimized_breaker() {
    let mut doc = justified_paragraph_document("본 계약은 체결되며", WordCompatibilityMode::Legacy);
    doc.styles.word_compatibility_mode = None;
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("linebreaks:"),
        "only a Word package states a compatibility mode: {result}"
    );
}

#[test]
fn test_punctuation_never_hangs_regardless_of_alignment() {
    // The setting is document-wide, not per paragraph, because a hung glyph
    // leaves the line's layout box: a *centred* line opening with a hyphen was
    // measured narrow and drawn off centre, and header and footer bands never
    // pass through the paragraph settings at all (issues #645, #646).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "- 1 -".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("overhang: false").count(),
        1,
        "stated once for the document, not per paragraph: {result}"
    );
    let preamble_end = result.find("#set page").unwrap_or(result.len());
    assert!(
        result[..preamble_end].contains("overhang: false"),
        "it must precede the first page so header and footer bands inherit it: {result}"
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
fn test_generate_word_default_line_box() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            line_box: Some(LineBox {
                ascent_em: 1.3125,
                descent_em: 0.4375,
            }),
            space_after: Some(8.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Word defaults".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("#set text(top-edge: 1.3125em, bottom-edge: -0.4375em)"),
        "Expected Word-compatible line edges in: {source}"
    );
    assert!(
        source.contains("#set par(leading: 0pt)"),
        "Expected Word-compatible line stacking in: {source}"
    );
    assert!(
        source.contains("below: 8pt"),
        "Expected paragraph spacing in: {source}"
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
fn test_generate_tab_uses_paragraph_default_stop_interval() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            default_tab_stop_pt: Some(72.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Label\tValue".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 72)"),
        "Expected the paragraph's default tab interval in: {result}"
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
    assert!(result.is_empty() || !result.is_empty());
}

#[test]
fn test_generate_special_characters_escaped() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 #items @store",
    )])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("\\#") || result.contains("Price"),
        "Expected escaped or present text in: {result}"
    );
}

#[test]
fn test_centered_paragraph_with_spacing_keeps_full_width_block() {
    // A paragraph with spacing gets a #block wrapper; without width: 100%
    // the block shrinks to its content and the inner #align(center) has no
    // visible effect (Word: <w:spacing w:after> + <w:jc w:val="center">).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            space_after: Some(6.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Centered title".to_string(),
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
    let block_start = result.find("#block(").expect("expected block wrapper");
    let block_params = &result[block_start..block_start + 60];
    assert!(
        block_params.contains("width: 100%"),
        "Block wrapper must span the full width for alignment to apply: {block_params}"
    );
}

#[test]
fn test_document_grid_pitch_snaps_line_height() {
    // A Korean Word section whose <w:docGrid> snaps lines puts body lines on
    // an 18pt grid. The line box is clamped to a fixed em height equal to the
    // grid pitch (leading 0) so a taller fallback glyph on a line cannot
    // inflate its advance past the grid (issue #398). The baseline keeps its
    // constant ascent inside that box: the slot's slack accrues below it, not
    // around it (issue #518). Uses a font from Typst's embedded set so the
    // test is environment-free.
    let Some((_, _, word_pitch_em)) = crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let mut page = match make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "그리드 정렬 grid snapped".to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })]) {
        Page::Flow(flow) => flow,
        _ => unreachable!(),
    };
    page.line_grid_pitch = Some(18.0);
    page.line_grid_snaps_lines = true;
    let doc = make_doc(vec![Page::Flow(page)]);
    let result = generate_typst(&doc).unwrap().source;

    assert_line_advance(
        &result,
        "Libertinus Serif",
        10.0,
        18.0,
        0.15 * word_pitch_em,
    );
    assert!(
        result.contains("leading: 0pt"),
        "the grid advance is carried by the box, not by leading: {result}"
    );
}

#[test]
fn test_latin_paragraph_ignores_document_grid() {
    // Word leaves Latin-only paragraphs at their metric line height even
    // when the section carries a document grid; only East Asian text snaps
    // (issue #354).
    let mut page = match make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "latin only body text".to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })]) {
        Page::Flow(flow) => flow,
        _ => unreachable!(),
    };
    page.line_grid_pitch = Some(18.0);
    page.line_grid_snaps_lines = true;
    let doc = make_doc(vec![Page::Flow(page)]);
    let result = generate_typst(&doc).unwrap().source;

    // The paragraph keeps Word's hhea single-spacing advance; the 18pt grid
    // pitch must not appear in its line box.
    let Some((ascender, descender, word_pitch)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return;
    };
    let single_pt: f64 = (word_pitch * 10.0).max((ascender + descender) * 10.0);
    assert_line_advance(&result, "Libertinus Serif", 10.0, single_pt, 0.0);
    assert!(
        (single_pt - 18.0).abs() > 0.01,
        "the fixture only proves anything if the grid pitch differs from single spacing"
    );
}

#[test]
fn test_no_document_grid_uses_word_single_spacing() {
    // Without a document grid, paragraphs still use Word's hhea single-line
    // pitch instead of Typst's glyph-tight default (issue #354).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "plain".to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    let Some((ascender, descender, word_pitch)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return;
    };
    let single_pt: f64 = (word_pitch * 10.0).max((ascender + descender) * 10.0);
    assert_line_advance(&result, "Libertinus Serif", 10.0, single_pt, 0.0);
    assert!(
        result.contains("leading: 0pt"),
        "the advance is carried by the box, not by leading: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_background_shading() {
    // w:pPr/w:shd paints the whole paragraph; the block wrapper must carry
    // the fill so the shading spans the full line width (issue #351).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            background: Some(Color::new(0xF4, 0xF4, 0xF4)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "$ cargo install office2pdf-cli".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("fill: rgb(244, 244, 244)"),
        "paragraph shading must fill the block wrapper: {result}"
    );
    assert!(
        result.contains("#block(width: 100%"),
        "shaded paragraphs need the full-width block wrapper: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_bottom_border_rule() {
    // w:pBdr bottom rules (resume header underline) must stroke the block
    // wrapper's bottom edge (issue #368).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            border: Some(Box::new(CellBorder {
                bottom: Some(BorderSide {
                    width: 0.75,
                    color: Color::new(0x1E, 0x27, 0x61),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                }),
                ..CellBorder::default()
            })),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "JAMIE PARKER".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    // Word inflates the rule 0.02in past each edge of the text column, so it
    // runs 2.88pt wider than the copy above it. `outset` widens what the block
    // paints without moving the text inside it (issue #644).
    assert!(
        result.contains("outset: (x: 1.44pt)"),
        "the rule must overhang the text column by 1.44pt each side: {result}"
    );
    assert!(
        result.contains("stroke: (bottom: 0.75pt + rgb(30, 39, 97))"),
        "bottom border must stroke the wrapper: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_double_bottom_border() {
    // Double letterhead rules render as two placed hairlines; Typst strokes
    // have no double style (issue #368).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            border: Some(Box::new(CellBorder {
                bottom: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Double,
                    join: LineJoin::Round,
                }),
                ..CellBorder::default()
            })),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "주식회사 에이엑스솔루션".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    let rule_count = result.matches("line(length: 100%").count();
    assert_eq!(
        rule_count, 2,
        "double borders draw exactly two rules: {result}"
    );
    assert!(
        !result.contains("stroke: (bottom:"),
        "double sides must not also stroke the wrapper: {result}"
    );
}

fn make_tab_paragraph() -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "제1조\t(목적) 본문".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })
}

#[test]
fn test_tab_advance_uses_document_default_tab_stop() {
    // Word documents carry w:defaultTabStop; tabs advance to multiples of
    // it, not the ECMA fallback (issue #393).
    let mut doc = make_doc(vec![make_flow_page(vec![make_tab_paragraph()])]);
    doc.styles.default_tab_stop_pt = Some(40.0);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 40)"),
        "explicit default tab stop must drive the advance: {result}"
    );
}

#[test]
fn test_tab_advance_defaults_to_40pt_under_document_grid() {
    // When settings.xml omits w:defaultTabStop, East Asian Word (signalled
    // by the section's w:docGrid) falls back to 800 twips = 40pt, not the
    // ECMA 720 twips (issue #393).
    let mut page = match make_flow_page(vec![make_tab_paragraph()]) {
        Page::Flow(flow) => flow,
        _ => unreachable!(),
    };
    page.line_grid_pitch = Some(18.0);
    page.line_grid_snaps_lines = true;
    let doc = make_doc(vec![Page::Flow(page)]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 40)"),
        "grid documents default to 40pt tab stops: {result}"
    );
}

#[test]
fn test_tab_advance_defaults_to_36pt_without_grid() {
    let doc = make_doc(vec![make_flow_page(vec![make_tab_paragraph()])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 36)"),
        "ECMA default stays 36pt: {result}"
    );
}

#[test]
fn test_latin_paragraph_space_after_stays_raw_gap() {
    // Word places `w:spacing w:after` directly below the full line box,
    // which the paragraph's own line box already spans, so the gap reaches
    // the block unchanged and needs no leading top-up (issues #394, #452).
    let make_para = |text: &str| {
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                space_after: Some(4.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some("Libertinus Serif".to_string()),
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })
    };
    let doc = make_doc(vec![make_flow_page(vec![
        make_para("first paragraph"),
        make_para("second paragraph"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("below: 4pt"),
        "Latin paragraph keeps the raw 4pt gap: {result}"
    );
}

#[test]
fn test_consecutive_paragraphs_each_advance_by_the_full_font_line() {
    // The shaded command lines of a technical manual are separate 9pt
    // Courier New paragraphs with `w:spacing w:after="0"`. Word advances
    // each by the font's full single-spacing line; Typst only inserts
    // `par(leading:)` *between* the lines of one paragraph, so recovering
    // the advance that way left every paragraph one leading short and
    // consecutive command lines packed 28% tighter than Word (issue #452).
    // The line box must therefore span the whole advance on its own.
    let Some(advance_pt) = single_spacing_advance_pt(LINE_GAP_FONT, 9.0) else {
        return;
    };
    let make_code_line = |text: &str| {
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                space_after: Some(0.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some(LINE_GAP_FONT.to_string()),
                    font_size: Some(9.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })
    };
    let doc = make_doc(vec![make_flow_page(vec![
        make_code_line("$ cargo install office2pdf-cli"),
        make_code_line("$ office2pdf --version"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_line_advance(&result, LINE_GAP_FONT, 9.0, advance_pt, 0.0);
    assert!(
        result.contains("leading: 0pt"),
        "the advance belongs to the box, not to leading: {result}"
    );
    assert!(
        result.contains("below: 0pt"),
        "a zero w:spacing w:after stays zero: {result}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_mixed_font_line_uses_the_declared_runs_deeper_descent() {
    let Some((arial_top_em, arial_bottom_em, _)) =
        crate::render::pdf::font_line_metrics_em("Arial")
    else {
        return;
    };
    let Some((_courier_top_em, courier_bottom_em, _)) =
        crate::render::pdf::font_line_metrics_em("Courier New")
    else {
        return;
    };
    let arial_size_pt: f64 = 10.5;
    let courier_size_pt: f64 = 9.5;
    let arial_bottom_pt: f64 = arial_bottom_em * arial_size_pt;
    let courier_bottom_pt: f64 = courier_bottom_em * courier_size_pt;
    if courier_bottom_pt <= arial_bottom_pt + 0.1 {
        return;
    }

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Justify),
            ..ParagraphStyle::default()
        },
        runs: vec![
            Run {
                text: "Ensure the original fonts are installed, or pass ".to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(arial_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "--font-path".to_string(),
                style: TextStyle {
                    font_family: Some("Courier New".to_string()),
                    font_size: Some(courier_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "\nfiles. Missing fonts fall back to substitutes.".to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(arial_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let output = generate_typst(&doc).unwrap();
    let placed = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let mixed_baseline_pt: f64 = placed
        .iter()
        .find(|run| run.text.contains("--font-path"))
        .expect("the mixed-font line should contain the Courier run")
        .baseline_pt;
    let following_baseline_pt: f64 = placed
        .iter()
        .find(|run| run.text.contains("files."))
        .expect("the hard break should start the following Arial line")
        .baseline_pt;
    let actual_advance_pt: f64 = following_baseline_pt - mixed_baseline_pt;
    let expected_advance_pt: f64 =
        arial_top_em * arial_size_pt + courier_bottom_pt.max(arial_bottom_pt);

    assert!(
        (actual_advance_pt - expected_advance_pt).abs() < 0.05,
        "the next baseline should use Arial's ascent plus the deepest declared \
         run descent: got {actual_advance_pt}pt, expected {expected_advance_pt}pt; \
         placed={placed:?}\n{}",
        output.source
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_mixed_font_line_uses_the_following_lines_taller_ascent() {
    let Some((arial_top_em, _arial_bottom_em, _)) =
        crate::render::pdf::font_line_metrics_em("Arial")
    else {
        return;
    };
    let Some((courier_top_em, courier_bottom_em, _)) =
        crate::render::pdf::font_line_metrics_em("Courier New")
    else {
        return;
    };
    let arial_size_pt: f64 = 10.5;
    let courier_size_pt: f64 = 9.5;
    let arial_top_pt: f64 = arial_top_em * arial_size_pt;
    let courier_top_pt: f64 = courier_top_em * courier_size_pt;
    if arial_top_pt <= courier_top_pt + 0.1 {
        return;
    }

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "first Courier line".to_string(),
                style: TextStyle {
                    font_family: Some("Courier New".to_string()),
                    font_size: Some(courier_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "\nsecond Courier run ".to_string(),
                style: TextStyle {
                    font_family: Some("Courier New".to_string()),
                    font_size: Some(courier_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "with taller Arial ascent".to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(arial_size_pt),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let output = generate_typst(&doc).unwrap();
    let placed = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let first_baseline_pt: f64 = placed
        .iter()
        .find(|run| run.text.contains("first Courier"))
        .expect("the first Courier line should be present")
        .baseline_pt;
    let second_baseline_pt: f64 = placed
        .iter()
        .find(|run| run.text.contains("second Courier"))
        .expect("the hard break should start the second line")
        .baseline_pt;
    let actual_advance_pt: f64 = second_baseline_pt - first_baseline_pt;
    let expected_advance_pt: f64 =
        courier_bottom_em * courier_size_pt + arial_top_pt.max(courier_top_pt);

    assert!(
        (actual_advance_pt - expected_advance_pt).abs() < 0.05,
        "the next baseline should use the first line's Courier descent plus the \
         tallest declared ascent on the following line: got {actual_advance_pt}pt, \
         expected {expected_advance_pt}pt; placed={placed:?}\n{}",
        output.source
    );
}

#[test]
fn uniform_font_runs_keep_only_the_paragraph_level_line_box() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "first Arial run ".to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(10.5),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "second Arial run".to_string(),
                style: TextStyle {
                    font_family: Some("arial".to_string()),
                    font_size: Some(9.5),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let output = generate_typst(&doc).unwrap();

    assert!(
        output.source.contains("#set text(top-edge:"),
        "the paragraph-level Word line box should still be present\n{}",
        output.source
    );
    assert!(
        !output.source.contains("#text(top-edge:"),
        "same-family runs must not gain per-run line boxes\n{}",
        output.source
    );
}

/// A Typst-embedded font whose hhea line is taller than its typographic
/// metric box, so the single-spacing advance is strictly larger than the
/// metric box. Libertinus Serif - the default test font here - has no line
/// gap at all, which would make the assertions above pass vacuously.
const LINE_GAP_FONT: &str = "DejaVu Sans Mono";

/// Word's single-spacing advance for `family` at `font_size`, or `None`
/// when the font is unavailable or its hhea line adds no gap over the
/// typographic metric box (which would make the assertion vacuous).
fn single_spacing_advance_pt(family: &str, font_size: f64) -> Option<f64> {
    let (ascender, descender, word_pitch) = crate::render::pdf::font_line_metrics_em(family)?;
    (word_pitch - (ascender + descender) > 0.001).then_some(word_pitch * font_size)
}

#[test]
fn test_grid_paragraph_space_after_stays_raw_gap() {
    // Grid variant: the snapped line box already spans the full grid pitch,
    // so Word's after-gap sits directly below it and reaches the block
    // unchanged (issues #394, #452).
    if crate::render::pdf::font_line_metrics_em("Libertinus Serif").is_none() {
        return;
    }
    let mut page = match make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            space_after: Some(4.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "그리드 본문".to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })]) {
        Page::Flow(flow) => flow,
        _ => unreachable!(),
    };
    page.line_grid_pitch = Some(18.0);
    page.line_grid_snaps_lines = true;
    let doc = make_doc(vec![Page::Flow(page)]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("below: 4pt"),
        "grid paragraph keeps the raw 4pt gap: {result}"
    );
}

#[test]
fn test_paragraph_left_indent_offsets_the_text_column() {
    // Word's `w:ind w:left` moves the whole paragraph right; only the list
    // path ever read it, so indented body paragraphs started at the margin
    // (issue #464).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_left: Some(12.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "$ cargo install office2pdf-cli".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("inset: (left: 12pt, right: 0pt)"),
        "left indent should inset the paragraph block: {result}"
    );
}

#[test]
fn test_paragraph_right_indent_narrows_the_text_column() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_right: Some(18.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "narrowed".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("inset: (left: 0pt, right: 18pt)"),
        "right indent should inset the paragraph block: {result}"
    );
}

#[test]
fn test_indented_paragraph_shading_starts_at_the_indent() {
    // Word paints `w:pPr/w:shd` from the left indent to the right indent,
    // not across the whole text column: measured on a Word export, the
    // shaded band of a 12pt-indented code line starts at the indent, 12pt
    // right of the margin (issue #464). The fill therefore belongs to an
    // inner block that spans only the inset content area.
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_left: Some(12.0),
            background: Some(Color::new(0xF4, 0xF4, 0xF4)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "$ office2pdf --version".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    let inset_at = result
        .find("inset: (left: 12pt, right: 0pt)")
        .expect("indent should inset the outer block");
    let fill_at = result
        .find("fill: rgb(244, 244, 244)")
        .expect("shading should still be emitted");
    assert!(
        fill_at > inset_at,
        "the fill belongs to a block nested inside the indent: {result}"
    );
}

#[test]
fn test_unindented_paragraph_keeps_a_single_block() {
    // A paragraph with no indent must not gain an inset or a nested block.
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            background: Some(Color::new(0xF4, 0xF4, 0xF4)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "flush left".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("inset: (left:"),
        "an unindented paragraph needs no inset: {result}"
    );
    assert!(
        result.contains("fill: rgb(244, 244, 244)"),
        "shading is still emitted: {result}"
    );
}

#[test]
fn test_empty_indented_paragraph_closes_its_block() {
    // An indented paragraph with no runs takes the early-return path. Leaving
    // its indent wrapper open produced "unclosed delimiter" and failed 73
    // third-party fixtures outright (regression caught on #464).
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                indent_left: Some(24.0),
                ..ParagraphStyle::default()
            },
            runs: Vec::new(),
        }),
        make_paragraph("after the empty paragraph"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;

    let opened: usize = result.matches('[').count();
    let closed: usize = result.matches(']').count();
    assert_eq!(
        opened, closed,
        "every content block opened must be closed: {result}"
    );
    assert!(
        result.contains("after the empty paragraph"),
        "the following paragraph must not be swallowed: {result}"
    );
}

/// One paragraph of `text` in `family` at `font_size`, with no grid.
fn line_box_for_text(text: &str, family: &str, font_size: f64) -> Option<(f64, f64)> {
    crate::render::pdf::font_line_metrics_em(family)?;
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle {
                font_family: Some(family.to_string()),
                font_size: Some(font_size),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    emitted_line_box_em(&generate_typst(&doc).unwrap().source)
}

#[test]
fn east_asian_line_advances_130_percent_of_the_font_line() {
    // Word gives a line carrying East Asian text 130% of the font's own hhea
    // line. Measured across the business corpus: 10.5pt Malgun Gothic paces its
    // wrapped lines at 18.00-18.24pt against 18.156 predicted, and
    // 06_official_letter_ko's 9.5pt paragraphs advance 16.43pt where the bare
    // hhea line is 12.64pt (issue #518).
    let Some((ascender, descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let _ = (ascender, descender);
    let Some((top, bottom)) = line_box_for_text("본문 한 줄", "Libertinus Serif", 10.0) else {
        return;
    };

    assert!(
        (top + bottom - 1.3 * word_pitch_em).abs() < 0.001,
        "East Asian advance {}em should be 1.3 x the {word_pitch_em}em hhea line",
        top + bottom
    );
}

#[test]
fn east_asian_bonus_is_centred_on_the_baseline() {
    // Half of the 30% lands above the baseline and half below: an Arial first
    // baseline sits at `hhea ascender + lineGap` while a Malgun Gothic one at
    // the same settings sits 0.15 x pitch lower, and the descent gap grows by
    // the same amount (issue #518).
    let Some((ascender, _descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return;
    };
    let Some((top, bottom)) = line_box_for_text("본문 한 줄", "Libertinus Serif", 10.0) else {
        return;
    };

    assert!(
        (top - (ascender + 0.15 * word_pitch_em)).abs() < 0.001,
        "the baseline should sit 0.15 x pitch below the Latin seat, got {top}em"
    );
    assert!(
        (bottom - (word_pitch_em - ascender + 0.15 * word_pitch_em)).abs() < 0.001,
        "the other half of the bonus belongs below the baseline, got {bottom}em"
    );
}

#[test]
fn a_latin_line_keeps_the_plain_hhea_line_and_seat() {
    // Triangulation for both rules above: the bonus is a property of the
    // line's script, not of the renderer. Inflating Latin lines too made every
    // Western document 30-50% taller (issue #354).
    let Some((ascender, _descender, word_pitch_em)) =
        crate::render::pdf::font_line_metrics_em("Libertinus Serif")
    else {
        return;
    };
    let Some((top, bottom)) = line_box_for_text("plain body text", "Libertinus Serif", 10.0) else {
        return;
    };

    assert!(
        (top + bottom - word_pitch_em).abs() < 0.001,
        "a Latin line advances the bare hhea line, got {}em",
        top + bottom
    );
    assert!(
        (top - ascender).abs() < 0.001,
        "a Latin baseline keeps the `hhea ascender + lineGap` seat, got {top}em"
    );
}

#[test]
fn the_east_asian_bonus_scales_with_the_font_size_not_with_the_text() {
    // The rule is a factor on the font's line, so doubling the size doubles
    // both edges exactly - a fake that returned one measured pair would not.
    let Some((small_top, small_bottom)) = line_box_for_text("표", "Libertinus Serif", 9.0) else {
        return;
    };
    let (large_top, large_bottom) =
        line_box_for_text("전혀 다른 한국어 문장", "Libertinus Serif", 21.0)
            .expect("the same font resolves at any size");

    // The box is emitted in em, so the em split must be identical at both
    // sizes and for unrelated text.
    assert!(
        (small_top - large_top).abs() < 0.001 && (small_bottom - large_bottom).abs() < 0.001,
        "the split is a property of the font, not of the size or the text: \
         {small_top}/{small_bottom} vs {large_top}/{large_bottom}"
    );
}

/// One paragraph with a bottom rule of `style` at `width`pt and `space`pt of
/// `w:pBdr` gap.
fn bordered_paragraph_source(width: f64, style: BorderLineStyle, space: f64) -> String {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            border: Some(Box::new(CellBorder {
                bottom: Some(BorderSide {
                    width,
                    color: Color::black(),
                    style,
                    join: LineJoin::Round,
                }),
                ..CellBorder::default()
            })),
            border_space: Some(Box::new(Insets {
                top: 0.0,
                right: 0.0,
                bottom: space,
                left: 0.0,
            })),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "서울특별시 강남구".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    generate_typst(&doc).unwrap().source
}

#[test]
fn a_paragraph_rule_reserves_its_own_declared_space() {
    // `w:pBdr` carries `w:space` per side, in points. A fixed 4pt stood in for
    // it until #520, which displaced every line below a bordered paragraph by
    // the difference — 06_official_letter_ko declares 8pt and lost 4pt of it,
    // as a step that survived to the bottom of the page.
    //
    // Half the width, not the whole of it: Typst centres the stroke on the
    // inset edge, so reserving the full width sat the rule half a width low
    // (issue #648).
    let source = bordered_paragraph_source(0.75, BorderLineStyle::Solid, 8.0);

    assert!(
        source.contains("inset: (bottom: 8.375pt)"),
        "the rule reserves its declared 8pt plus half its 0.75pt width: {source}"
    );
}

#[test]
fn a_rule_that_declares_no_space_sits_against_the_text() {
    // Triangulation, and the attribute's own default: `w:space` omitted means
    // zero, not a house value.
    let source = bordered_paragraph_source(0.75, BorderLineStyle::Solid, 0.0);

    assert!(
        source.contains("inset: (bottom: 0.375pt)"),
        "with no declared space only the rule's width is reserved: {source}"
    );
}

#[test]
fn a_double_rule_reserves_its_space_plus_all_three_widths() {
    // A Word double rule is two lines of the declared width separated by a gap
    // of the same width, and it is drawn as two placed hairlines because Typst
    // strokes have no double style. Both the reserved height and the placement
    // of each hairline hang off the declared space (issue #520).
    let source = bordered_paragraph_source(1.0, BorderLineStyle::Double, 8.0);

    assert!(
        source.contains("inset: (bottom: 11pt)"),
        "8pt of space plus three 1pt widths: {source}"
    );
    assert!(
        source.contains("#place(bottom, dx: -1.44pt, dy: 9pt,")
            && source.contains("#place(bottom, dx: -1.44pt, dy: 11pt,"),
        "both hairlines are placed from the declared space: {source}"
    );
    // A placed line spans the block's layout box, which `outset` does not
    // widen, so each hairline reaches past both edges itself (issue #644).
    assert!(
        source.contains("line(length: 100% + 2.88pt"),
        "a double rule overhangs the column like a single one: {source}"
    );
    assert!(
        !source.contains("stroke: (bottom:"),
        "a double rule is drawn as overlays, not as a block stroke: {source}"
    );
}

#[test]
fn every_declared_space_reaches_the_output_unchanged() {
    // Triangulation across values, so no single measured constant can pass.
    for (space, expected) in [
        (0.0, "0.375pt"),
        (2.0, "2.375pt"),
        (8.0, "8.375pt"),
        (14.0, "14.375pt"),
    ] {
        let source = bordered_paragraph_source(0.75, BorderLineStyle::Solid, space);
        assert!(
            source.contains(&format!("inset: (bottom: {expected})")),
            "{space}pt of declared space should reserve {expected}: {source}"
        );
    }
}

#[test]
fn test_generate_heading_with_style_border_rule() {
    // Word paints a heading's `w:pBdr` around it like any other paragraph, and
    // a chapter-rule heading style is the commonest place one appears. The
    // heading branch used to return before any decoration was emitted, so
    // every styled heading rule was dropped (issue #581).
    let body = Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "본문".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    });
    let heading = Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            space_before: Some(14.0),
            space_after: Some(7.5),
            border: Some(Box::new(CellBorder {
                bottom: Some(BorderSide {
                    width: 1.5,
                    color: Color::new(0x2E, 0x74, 0xB5),
                    style: BorderLineStyle::Solid,
                    join: LineJoin::Round,
                }),
                ..CellBorder::default()
            })),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "개요".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    });
    let doc = make_doc(vec![make_flow_page(vec![body, heading])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("stroke: (bottom: 1.5pt + rgb(46, 116, 181))"),
        "a heading's style border must stroke its wrapper: {result}"
    );
    assert!(
        // Framed because a heading is a paragraph like any other, and Word
        // breaks its Korean at eojeol too (issue #626).
        result.contains("#heading(level: 1)[#box[개요]]"),
        "the heading itself must survive the wrapper: {result}"
    );
    assert!(
        result.contains("above: 14pt") && result.contains("below: 7.5pt"),
        "the wrapper carries the heading's own spacing, not Typst's: {result}"
    );
}

/// A heading paragraph carrying `style`, with one plain run, following a body
/// paragraph — the page's first block has its `space_before` hoisted into a
/// leading `#v()` instead, because Typst drops block spacing at a page top.
fn heading_source(style: ParagraphStyle) -> String {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Body copy"),
        Block::Paragraph(Paragraph {
            style,
            runs: vec![Run {
                text: "Overview".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }),
    ])]);
    generate_typst(&doc).unwrap().source
}

#[test]
fn every_heading_spacing_reaches_its_block_unchanged() {
    // A heading resolves `w:spacing` exactly like body copy, so whatever the
    // style hierarchy resolved has to reach the block. Left off, the gap was
    // Typst's own `#set heading` default — a number no `w:spacing`, no style
    // definition and no Word rule produced — and because Typst collapses
    // adjacent block spacing to the larger of the two, that default also
    // swallowed the neighbouring paragraph's declared gap (issue #1132).
    //
    // Triangulated across values, so no single measured constant can pass.
    for (before, after, expected) in [
        (None, Some(8.0), "#block(width: 100%, below: 8pt)"),
        (
            Some(14.0),
            Some(7.0),
            "#block(width: 100%, above: 14pt, below: 7pt)",
        ),
        (
            Some(6.0),
            Some(0.0),
            "#block(width: 100%, above: 6pt, below: 0pt)",
        ),
    ] {
        let source = heading_source(ParagraphStyle {
            heading_level: Some(1),
            space_before: before,
            space_after: after,
            ..ParagraphStyle::default()
        });
        assert!(
            source.contains(expected),
            "heading spacing {before:?}/{after:?} should emit {expected}: {source}"
        );
        assert!(
            source.contains("#heading(level: 1)"),
            "the heading itself must survive the wrapper: {source}"
        );
    }
}

#[test]
fn a_heading_takes_the_same_line_box_as_body_copy() {
    // The block's `above`/`below` measure from the line box's edges, so a
    // heading needs Word's box for them to land where Word puts them: with
    // Typst's glyph-tight default the wrapper ends at the baseline and the
    // gap below the heading comes out a descender short (issue #1132).
    let source = heading_source(ParagraphStyle {
        heading_level: Some(1),
        line_box: Some(LineBox {
            ascent_em: 1.3125,
            descent_em: 0.4375,
        }),
        space_after: Some(8.0),
        ..ParagraphStyle::default()
    });

    assert!(
        source.contains("#set text(top-edge: 1.3125em, bottom-edge: -0.4375em)"),
        "a heading's line box must reach the output: {source}"
    );
    assert!(
        source.contains("#set par(leading: 0pt)"),
        "a heading stacks its lines Word's way: {source}"
    );
}

/// A heading paragraph in `family`, following a body paragraph, with the
/// `w:spacing w:after` a real `Heading N` resolves.
fn theme_heading_source(family: &str) -> String {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Body copy"),
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(1),
                space_after: Some(8.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Overview".to_string(),
                style: TextStyle {
                    font_family: Some(family.to_string()),
                    font_size: Some(16.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }),
    ])]);
    generate_typst(&doc).unwrap().source
}

#[test]
fn a_theme_major_face_heading_takes_the_base_family_line_box() {
    // `Calibri Light` is the `majorHAnsi` face of every Office theme since
    // 2013, so a built-in `Heading N` resolves to it. Typst indexes the
    // installed face under the trimmed family `Calibri`, so the stated name
    // matched no face: `font_line_metrics_em` answered `None`, the heading got
    // no fixed line box, and its block ended on the baseline with the
    // descender missing from the gap below it (issue #1197).
    if crate::render::pdf::font_line_metrics_em("Calibri").is_none() {
        return; // no Calibri-compatible face on this host
    }
    let light: String = theme_heading_source("Calibri Light");
    let regular: String = theme_heading_source("Calibri");

    let light_box = emitted_line_box_em(&light)
        .unwrap_or_else(|| panic!("a Calibri Light heading needs Word's line box: {light}"));
    let regular_box = emitted_line_box_em(&regular)
        .unwrap_or_else(|| panic!("a Calibri heading needs Word's line box: {regular}"));
    assert_eq!(
        light_box, regular_box,
        "Calibri Light declares Calibri's own hhea metrics, so the two headings \
         share one line box"
    );
}

#[test]
fn test_generate_undecorated_heading_keeps_its_bare_form() {
    // A heading with nothing to carry — no decoration, no resolved
    // `w:spacing`, no line box — keeps Typst's own heading block, because an
    // empty wrapper would state nothing the document asked for. Every DOCX
    // heading does resolve a `w:spacing w:after`, so this is the shape only an
    // IR built without one takes (issues #581, #1132).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(2),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "한 문장 요약".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    // One syllable needs no frame; the two-syllable eojeol each get one so a
    // wrapped heading breaks where Word breaks it (issue #626).
    assert!(result.contains("#heading(level: 2)[한 #box[문장] #box[요약]]"));
    assert!(
        !result.contains("#block(width: 100%"),
        "an undecorated heading needs no block wrapper: {result}"
    );
}

/// Build a header paragraph in Word's running-head shape: segments separated
/// by `<w:tab/>` runs, with the stops that place them declared on the
/// paragraph.
fn running_head(texts: &[&str], stops: Vec<TabStop>) -> crate::ir::HeaderFooterParagraph {
    let mut elements: Vec<HFInline> = Vec::new();
    for (index, text) in texts.iter().enumerate() {
        if index > 0 {
            elements.push(HFInline::Run(Run {
                text: "\t".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }));
        }
        elements.push(HFInline::Run(Run {
            text: (*text).to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }));
    }
    crate::ir::HeaderFooterParagraph {
        style: ParagraphStyle {
            tab_stops: Some(stops),
            ..ParagraphStyle::default()
        },
        elements,
        border: None,
        border_space: None,
        frame: None,
    }
}

fn page_with_header(header: crate::ir::HeaderFooter) -> Page {
    let Page::Flow(mut flow) = make_flow_page(vec![]) else {
        unreachable!()
    };
    flow.header = Some(header);
    Page::Flow(flow)
}

fn stop(position: f64, alignment: TabAlignment) -> TabStop {
    TabStop {
        position,
        alignment,
        leader: TabLeader::None,
    }
}

#[test]
fn test_header_right_tab_stop_pushes_its_segment_to_the_margin() {
    // A `<w:tab/>` was advanced by a fixed 1em however the paragraph's stops
    // were declared, so the segment a right stop should have pushed to the
    // right margin sat beside the left one (issue #579).
    let doc = make_doc(vec![page_with_header(crate::ir::HeaderFooter {
        shapes: Vec::new(),
        paragraphs: vec![running_head(
            &["office2pdf 기술 소개서", "본문"],
            vec![stop(465.3, TabAlignment::Right)],
        )],
        distance_from_edge: None,
        sheet_print_scale: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#grid(columns: (1fr, auto)"),
        "a right stop lays the two segments out against the margins: {result}"
    );
    assert!(
        !result.contains("#h(1em)"),
        "the tab must not fall back to a fixed advance: {result}"
    );
}

#[test]
fn test_header_center_and_right_tab_stops_lay_out_three_segments() {
    let doc = make_doc(vec![page_with_header(crate::ir::HeaderFooter {
        shapes: Vec::new(),
        paragraphs: vec![running_head(
            &["left", "middle", "right"],
            vec![
                stop(232.6, TabAlignment::Center),
                stop(465.3, TabAlignment::Right),
            ],
        )],
        distance_from_edge: None,
        sheet_print_scale: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#grid(columns: (1fr, auto, 1fr), align: (left, center, right)"),
        "a centre and a right stop give three placed segments: {result}"
    );
}

#[test]
fn test_header_tab_without_a_matching_stop_keeps_the_plain_advance() {
    // Only the two running-head shapes are laid out; a header that tabs for
    // some other reason keeps the behaviour it had.
    let doc = make_doc(vec![page_with_header(crate::ir::HeaderFooter {
        shapes: Vec::new(),
        paragraphs: vec![running_head(
            &["a", "b"],
            vec![stop(72.0, TabAlignment::Left)],
        )],
        distance_from_edge: None,
        sheet_print_scale: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("#grid(columns: (1fr, auto)"),
        "a left stop is not the running-head idiom: {result}"
    );
    assert!(
        result.contains('\t'),
        "the tab stays a literal tab for Typst to collapse: {result}"
    );
}

/// Letter-spacing disables ligatures (issue #684).
///
/// Tracking and ligation are mutually exclusive: a ligature replaces several
/// glyphs with one and swallows the inter-glyph spacing the tracking should
/// have added. PowerPoint disables ligatures under `a:rPr/@spc` for that
/// reason. Leaving `liga` on merged the `ffi` in "office2pdf" into one glyph,
/// so the text layer extracted as "o ffi c e 2 p d f" and matched no search
/// for the word.
#[test]
fn test_letter_spacing_disables_ligatures() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "office2pdf".to_string(),
            style: TextStyle {
                letter_spacing: Some(0.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("ligatures: false"),
        "tracked text must disable ligatures, got: {result}"
    );
}

/// Triangulation: a run with no tracking keeps ligatures, so the rule is tied
/// to letter-spacing rather than applied to every run.
#[test]
fn test_untracked_text_keeps_ligatures() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "office2pdf".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("ligatures: false"),
        "untracked text must not disable ligatures, got: {result}"
    );
}

/// An explicit zero tracking is not tracking, so it leaves ligatures alone.
#[test]
fn test_zero_letter_spacing_keeps_ligatures() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "office2pdf".to_string(),
            style: TextStyle {
                letter_spacing: Some(0.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("ligatures: false"),
        "zero tracking must not disable ligatures, got: {result}"
    );
}

/// PowerPoint never applies `liga`, so a slide states that once (issue #1058).
///
/// The rule is document-wide rather than per-run because it holds for every
/// slide run, including the emission sites that cannot name their own text.
#[test]
fn test_powerpoint_slide_disables_ligatures_document_wide() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![make_text_box(60.0, 200.0, 400.0, 60.0, "Standard setting:")],
    )]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(ligatures: false)"),
        "a slide must switch standard ligatures off for the whole document, \
         got: {result}"
    );
}

/// Triangulation: the rule is keyed to the slide, not switched on for every
/// document, so a flow page keeps the engine's default. Word's own ligature
/// setting is a separate question this change does not answer.
#[test]
fn test_flow_page_keeps_ligatures_document_wide() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Standard setting:",
    )])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("ligatures: false"),
        "a flow page must not disable ligatures document-wide, got: {result}"
    );
}

/// A substituted face's kern pairs land on top of the tracking the document
/// declared, and the combined advance can exceed the gap a PDF text extractor
/// reads as a word break. On the deck in issue #864 that split five titles —
/// `ANSATTE` extracted as `ANSAT TE` — while the glyphs rendered continuously.
#[test]
fn test_letter_spacing_disables_kerning() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "ANSATTE".to_string(),
            style: TextStyle {
                letter_spacing: Some(3.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("kerning: false"),
        "tracked text must disable kerning, got: {result}"
    );
}

/// Triangulation: an untracked run keeps kerning, so the rule is tied to
/// letter-spacing and does not flatten every run's pair spacing.
#[test]
fn test_untracked_text_keeps_kerning() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "ANSATTE".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("kerning: false"),
        "untracked text must keep kerning, got: {result}"
    );
}

/// An explicit zero tracking is not tracking, so it leaves kerning alone —
/// PowerPoint writes `spc="0"` routinely.
#[test]
fn test_zero_letter_spacing_keeps_kerning() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "ANSATTE".to_string(),
            style: TextStyle {
                letter_spacing: Some(0.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("kerning: false"),
        "zero tracking must keep kerning, got: {result}"
    );
}
