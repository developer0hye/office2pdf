use super::*;
use crate::ir::PairKerning;

// ── Unicode NFC normalization tests ──────────────────────────────

#[test]
fn test_escape_typst_normalizes_korean_nfd_to_nfc() {
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let nfc_korean = "한글";
    let result = escape_typst(nfd_korean);
    assert_eq!(
        result, nfc_korean,
        "NFD Korean jamo should be normalized to composed hangul"
    );
}

#[test]
fn test_escape_typst_normalizes_combining_diacritics() {
    let nfd_cafe = "cafe\u{0301}";
    let nfc_cafe = "caf\u{00E9}";
    let result = escape_typst(nfd_cafe);
    assert_eq!(
        result, nfc_cafe,
        "Combining diacritics should be normalized to NFC"
    );
}

#[test]
fn test_escape_typst_nfc_with_special_chars() {
    let nfd_input = "cafe\u{0301} \\$5";
    let result = escape_typst(nfd_input);
    assert!(
        result.contains("caf\u{00E9}"),
        "Should contain NFC-normalized é: {result}"
    );
    assert!(
        result.contains("\\$"),
        "Should still escape $ sign: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_korean_in_paragraph() {
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_korean)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("한글"),
        "Generated Typst should contain NFC-composed Korean: {result}"
    );
    assert!(
        !result.contains('\u{1112}'),
        "Generated Typst should not contain decomposed jamo: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_diacritics_in_paragraph() {
    let nfd_resume = "re\u{0301}sume\u{0301}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_resume)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("r\u{00E9}sum\u{00E9}"),
        "Generated Typst should contain NFC-composed résumé: {result}"
    );
}

#[test]
fn test_escape_typst_already_nfc_unchanged() {
    let nfc_text = "Hello 한글 café";
    let result = escape_typst(nfc_text);
    assert_eq!(result, nfc_text, "Already-NFC text should be unchanged");
}

// --- US-103: Multi-column section layout codegen tests ---

#[test]
fn test_generate_flow_page_with_equal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Column text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(2, gutter: 36pt)"),
        "Should contain columns() call. Got: {result}"
    );
    assert!(
        result.contains("Column text"),
        "Should contain the text content. Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_three_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Three col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 3,
            spacing: 18.0,
            column_widths: None,
        }),
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(3, gutter: 18pt)"),
        "Should contain columns(3, ...). Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_unequal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Unequal col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: Some(vec![300.0, 150.0]),
        }),
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#grid(columns: (300pt, 150pt)"),
        "Unequal columns should use grid(). Got: {result}"
    );
}

#[test]
fn test_generate_column_break() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![
            make_paragraph("Before break"),
            Block::ColumnBreak,
            make_paragraph("After break"),
        ],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#colbreak()"),
        "Should contain colbreak(). Got: {result}"
    );
}

#[test]
fn test_generate_no_columns_no_wrapper() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Normal text")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#columns("),
        "Should not contain columns(). Got: {result}"
    );
    assert!(
        !result.contains("#grid(columns:"),
        "Should not contain grid(columns:). Got: {result}"
    );
}

// ── BiDi / RTL codegen tests ──────────────────────────────────────

#[test]
fn test_generate_rtl_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            direction: Some(TextDirection::Rtl),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "مرحبا بالعالم".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(dir: rtl)"),
        "RTL paragraph should emit #set text(dir: rtl). Got: {result}"
    );
}

#[test]
fn test_generate_ltr_paragraph_no_direction() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("dir: rtl"),
        "LTR paragraph should not emit dir: rtl. Got: {result}"
    );
}

#[test]
fn test_generate_mixed_rtl_ltr_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                direction: Some(TextDirection::Rtl),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "مرحبا 123".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }),
        make_paragraph("English text"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(dir: rtl)"),
        "Should contain RTL direction for Arabic paragraph. Got: {result}"
    );
    assert!(result.contains("مرحبا 123"), "Arabic text should appear");
    assert!(
        result.contains("English text"),
        "English text should appear"
    );
}

// --- US-204: Codegen/render robustness tests ---

#[test]
fn test_codegen_robustness_zero_pages() {
    let doc = make_doc(vec![]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.images.is_empty());
}

#[test]
fn test_codegen_robustness_flow_page_empty_content() {
    let doc = make_doc(vec![make_flow_page(vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_fixed_page_empty_elements() {
    let doc = make_doc(vec![Page::Fixed(FixedPage {
        size: PageSize::default(),
        elements: vec![],
        background_color: None,
        background_gradient: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_table_page_empty_rows() {
    let doc = make_doc(vec![Page::Sheet(SheetPage {
        name: String::new(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![],
            column_widths: vec![],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_paragraph_all_alignment_variants() {
    for alignment in [
        Some(Alignment::Left),
        Some(Alignment::Center),
        Some(Alignment::Right),
        Some(Alignment::Justify),
        None,
    ] {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment,
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: format!("Alignment: {alignment:?}"),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not fail for alignment {alignment:?}"
        );
    }
}

#[test]
fn test_generate_shape_shadow_all_kinds() {
    let shadow = Shadow {
        blur_radius: 4.0,
        color: Color { r: 0, g: 0, b: 0 },
        opacity: 0.5,
        direction: 45.0,
        distance: 3.0,
    };

    let shape_kinds = vec![
        ShapeKind::Rectangle,
        ShapeKind::Ellipse,
        ShapeKind::Line {
            x1: 0.0,
            y1: 0.0,
            x2: 100.0,
            y2: 0.0,
            head_end: ArrowHead::None,
            tail_end: ArrowHead::None,
        },
        ShapeKind::RoundedRectangle {
            radius_fraction: 0.1,
        },
        ShapeKind::Polygon {
            vertices: vec![(0.0, 0.0), (1.0, 0.0), (0.5, 1.0)],
        },
    ];

    for kind in shape_kinds {
        let doc = make_doc(vec![Page::Fixed(FixedPage {
            size: PageSize {
                width: 960.0,
                height: 540.0,
            },
            elements: vec![FixedElement {
                x: 100.0,
                y: 100.0,
                width: 200.0,
                height: 100.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: kind.clone(),
                    fill: Some(Color { r: 255, g: 0, b: 0 }),
                    gradient_fill: None,
                    pattern_fill: None,
                    stroke: None,
                    opacity: None,
                    shadow: Some(shadow.clone()),
                    top_bevel: None,
                    rotation_deg: None,
                }),
            }],
            background_color: None,
            background_gradient: None,
        })]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not panic for shape kind {kind:?} with shadow"
        );
    }
}

#[test]
fn test_column_break_with_empty_content() {
    let segments = split_at_column_breaks(&[]);
    assert_eq!(segments.len(), 1);
    assert!(segments[0].is_empty());
}

#[test]
fn test_column_break_only_breaks() {
    let blocks = vec![Block::ColumnBreak, Block::ColumnBreak];
    let segments = split_at_column_breaks(&blocks);
    assert_eq!(segments.len(), 3);
    assert!(segments.iter().all(|segment| segment.is_empty()));
}

// --- US-315: text escaping for Typst-significant characters ---

#[test]
fn test_escape_typst_backslash() {
    assert_eq!(escape_typst("path\\to\\file"), "path\\\\to\\\\file");
}

#[test]
fn test_escape_typst_hash() {
    assert_eq!(escape_typst("#hashtag"), "\\#hashtag");
}

#[test]
fn test_escape_typst_dollar() {
    assert_eq!(escape_typst("$100"), "\\$100");
}

#[test]
fn test_escape_typst_brackets() {
    assert_eq!(escape_typst("[content]"), "\\[content\\]");
}

#[test]
fn test_escape_typst_braces() {
    assert_eq!(escape_typst("{code}"), "\\{code\\}");
}

#[test]
fn test_escape_typst_all_special_chars() {
    let input = r"#*_`<>@\~/$[]{}";
    let result = escape_typst(input);
    assert_eq!(result, "\\#\\*\\_\\`\\<\\>\\@\\\\\\~\\/\\$\\[\\]\\{\\}");
}

#[test]
fn test_escape_typst_in_paragraph_output() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 path\\to",
    )])]);
    let output = generate_typst(&doc).unwrap().source;
    assert!(
        output.contains("\\$100"),
        "Dollar sign should be escaped in output: {output}"
    );
    assert!(
        output.contains("path\\\\to"),
        "Backslash should be escaped in output: {output}"
    );
}

// --- US-316: single-stop gradient fallback ---

#[test]
fn test_gradient_single_stop_fallback_to_solid() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![GradientStop {
                position: 0.5,
                color: Color::new(255, 128, 0),
            }],
            angle: 0.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("gradient.linear"),
        "Single-stop gradient should fall back to solid fill: {}",
        output.source,
    );
    assert!(
        output.source.contains("rgb(255, 128, 0)"),
        "Single-stop gradient should use the stop color as solid fill: {}",
        output.source,
    );
}

#[test]
fn test_gradient_two_stops_still_works() {
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
        output.source.contains("gradient.linear"),
        "Two-stop gradient should still produce gradient.linear: {}",
        output.source,
    );
}

// --- US-382/383: unstyled run after styled run must not create `](` pattern ---

#[test]
fn test_unstyled_run_with_parens_after_styled_run() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "bold text".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "(parenthetical note)".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("](\\(") || !result.contains("]("),
        "Unstyled text with parens after styled run must be wrapped safely. Got: {result}"
    );
    assert!(
        result.contains("#[") || result.contains("\\("),
        "Unstyled text should be wrapped in #[...] to prevent syntax issues. Got: {result}"
    );
}

#[test]
fn test_escape_typst_escapes_leading_numeric_enum_marker() {
    // "2026. 07. 17." at line start would otherwise be re-typeset as a
    // Typst numbered list, dropping the zero padding ("2026. 7. 17.").
    let result = escape_typst("2026. 07. 17.");
    assert!(
        result.starts_with("2026\\."),
        "leading digits-period must be escaped: {result}"
    );
}

#[test]
fn test_escape_typst_keeps_mid_text_numbers_untouched() {
    let result = escape_typst("가격은 2026. 07 기준");
    assert!(result.contains("07"), "digits must survive: {result}");
}

#[test]
fn test_escape_typst_numeric_without_following_space_untouched() {
    // "3.14" is not an enum marker.
    assert_eq!(escape_typst("3.14"), "3.14");
}

// ── Preserved-space tests (issue #352) ───────────────────────────

#[test]
fn test_escape_typst_preserves_consecutive_spaces() {
    // Word keeps literal space runs (xml:space="preserve") that documents
    // use for manual alignment; Typst markup collapses them to one space.
    let result = escape_typst("Invoice #: INV-0342    Date: July 10");
    assert!(
        result.contains("#\"    \";"),
        "runs of spaces must survive markup collapsing: {result}"
    );
}

#[test]
fn test_escape_typst_preserves_leading_space_runs() {
    // Leading indentation ("      2. 계정 현행화 양식 1부.", code lines)
    // is stripped by markup whitespace handling.
    let result = escape_typst("      2. indented");
    assert!(
        result.starts_with("#\"      \";"),
        "leading space runs must survive: {result}"
    );
    assert!(
        result.ends_with("2. indented"),
        "text must follow: {result}"
    );
}

#[test]
fn test_escape_typst_preserves_spaces_after_hard_linebreak() {
    // Code blocks carry hard breaks followed by indentation.
    let result = escape_typst("match x {\n  b\"w:p\" => 1,\n}");
    assert!(
        result.contains("#linebreak();#\"  \";"),
        "post-break indentation must survive: {result}"
    );
}

#[test]
fn test_escape_typst_terminates_hard_linebreak_before_parenthesis() {
    // A fee-schedule cell reads "New York" over "(07) Western". Typst parses
    // an unterminated `#linebreak()(07)` as a call on the break's content
    // and aborts the whole conversion with "expected function, found
    // content".
    let result = escape_typst("New York\n(07) Western");
    assert!(
        result.contains("#linebreak();(07) Western"),
        "the break must be terminated before literal text: {result}"
    );
}

#[test]
fn test_escape_typst_terminates_hard_linebreak_before_field_access() {
    // A leading `.name` would otherwise read as a field access on the break
    // ("linebreak does not have field"). A leading `.5` is a number and
    // never chained, so the case below is the one that failed documents.
    let result = escape_typst("Runtime\n.NET pricing");
    assert!(
        result.contains("#linebreak();.NET pricing"),
        "the break must be terminated before a leading field access: {result}"
    );
}

#[test]
fn test_escape_typst_single_interior_space_untouched() {
    assert_eq!(escape_typst("a b"), "a b");
}

// ── Smart-typography escape tests (issue #353) ───────────────────

#[test]
fn test_escape_typst_keeps_straight_double_quotes() {
    // Typst smart quotes turned literal "quoted" into curly “quoted”.
    let result = escape_typst("run \"quoted\" text");
    assert!(
        result.contains("\\\"quoted\\\""),
        "straight double quotes must be escaped so smartquote cannot rewrite them: {result}"
    );
}

#[test]
fn test_escape_typst_keeps_straight_single_quotes() {
    let result = escape_typst("it's 'fine'");
    assert!(
        result.contains("it\\'s \\'fine\\'"),
        "straight apostrophes must be escaped: {result}"
    );
}

#[test]
fn test_escape_typst_keeps_double_hyphens() {
    // `--` ligates to an en dash, corrupting CLI flags like --font-path.
    let result = escape_typst("office2pdf --font-path dir --version");
    assert!(
        result.contains("\\-\\-font\\-path") || result.contains("\\-\\-font-path"),
        "double hyphens must not ligate to an en dash: {result}"
    );
    assert!(
        !result.contains("--"),
        "no raw double hyphen may remain: {result}"
    );
}

#[test]
fn test_escape_typst_keeps_hyphen_before_digits() {
    // A hyphen before digits becomes a Unicode minus (−18%) in markup.
    let result = escape_typst("blended CAC, -18%");
    assert!(
        result.contains("\\-18"),
        "hyphen before digits must stay a hyphen-minus: {result}"
    );
}

/// Word's East Asian/Latin auto space becomes a quarter of the *run's* size.
///
/// Sized in points rather than `em` because the spacing is emitted between the
/// run's `#text(size:)` calls: an `em` there resolves against the paragraph's
/// default size, which put 2.75pt at every boundary of a 10.5pt run and made a
/// line wide enough to re-wrap (issue #521).
#[test]
fn the_auto_space_marker_becomes_a_quarter_of_the_runs_size() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "2026\u{E001}년".to_string(),
            style: TextStyle {
                font_family: Some("Malgun Gothic".to_string()),
                font_size: Some(10.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("spacing: 2.625pt)[\\u{00A0}]"),
        "0.25 x 10.5pt should reach the output as points: {result}"
    );
    assert!(
        !result.contains('\u{E001}'),
        "the marker must never be emitted literally: {result}"
    );
}

#[test]
fn the_auto_space_scales_with_the_run_not_the_document() {
    // Triangulation: a different run size must produce a different gap, so a
    // single measured constant cannot pass.
    for (size, expected) in [
        (9.5, "spacing: 2.375pt)[\\u{00A0}]"),
        (16.0, "spacing: 4pt)[\\u{00A0}]"),
    ] {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "3\u{E001}자".to_string(),
                style: TextStyle {
                    font_family: Some("Malgun Gothic".to_string()),
                    font_size: Some(size),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(expected),
            "a {size}pt run should emit {expected}: {result}"
        );
    }
}

// ── Hangul eojeol line breaking (issue #626) ─────────────────────────

/// The Korean sentence the issue measured, cut to the part that fits one
/// assertion. Word breaks it only at the spaces.
const EOJEOL_SENTENCE: &str = "본 계약은 갑과 을이";

/// A Korean paragraph, optionally justified, in the given font.
fn korean_paragraph(text: &str, alignment: Option<Alignment>, family: Option<&str>) -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment,
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle {
                font_family: family.map(str::to_string),
                east_asian_font_family: family.map(str::to_string),
                font_size: family.map(|_| 10.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })
}

#[test]
fn a_docx_paragraph_keeps_each_hangul_eojeol_whole() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        EOJEOL_SENTENCE,
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    // A one-syllable eojeol needs no frame — nothing can break inside it.
    assert!(
        result.contains("본 #box[계약은] #box[갑과] #box[을이]"),
        "each multi-syllable eojeol should be an unbreakable inline box: {result}"
    );
}

/// `w:wordWrap w:val="0"` asks Word for character-level breaking of Hangul,
/// and it overrides the style chain. A paragraph that says so must not get the
/// eojeol frames #626 gives every other flow-page paragraph (issue #730).
#[test]
fn a_docx_paragraph_asking_for_character_breaking_gets_no_eojeol_frames() {
    let mut paragraph = korean_paragraph(EOJEOL_SENTENCE, None, None);
    if let Block::Paragraph(ref mut p) = paragraph {
        p.style.word_wrap = Some(false);
    }
    let doc = make_doc(vec![make_flow_page(vec![paragraph])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(EOJEOL_SENTENCE),
        "the text stays one run so Typst may break inside an eojeol: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no eojeol frame may be emitted when wordWrap is off: {result}"
    );
}

/// Triangulation: `w:val="1"` is the word-level setting, so it must keep the
/// frames rather than being treated as "the property is present, back off".
#[test]
fn a_docx_paragraph_asking_for_word_breaking_keeps_its_eojeol_frames() {
    let mut paragraph = korean_paragraph(EOJEOL_SENTENCE, None, None);
    if let Block::Paragraph(ref mut p) = paragraph {
        p.style.word_wrap = Some(true);
    }
    let doc = make_doc(vec![make_flow_page(vec![paragraph])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("본 #box[계약은] #box[갑과] #box[을이]"),
        "wordWrap=1 keeps each eojeol whole: {result}"
    );
}

/// Justification does not decide how a Korean line breaks — the style rule
/// does. A one-factor `w:jc` probe in a package that defines a default
/// paragraph style measured Word breaking `… 체결되며 ABC | 주식회사와 …` at
/// exactly the same two eojeol boundaries for `left`, `both`, `center` and
/// `right`, stretching the justified first line 55.12pt to the measure rather
/// than pulling `주` up onto it (issue #1084).
#[test]
fn a_justified_docx_paragraph_keeps_its_eojeol_frames() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        EOJEOL_SENTENCE,
        Some(Alignment::Justify),
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("본 #box[계약은] #box[갑과] #box[을이]"),
        "a justified line keeps each eojeol whole, as Word does: {result}"
    );
}

/// Triangulation for the arm removed by #1084: the justified paragraphs that
/// *were* measured breaking mid-eojeol — `02_contract_ko`'s, in a package
/// defining no paragraph style — reach codegen carrying the effective
/// `w:wordWrap w:val="0"` of the #732 style rule, and that is what has to
/// suppress the frames.
#[test]
fn a_justified_docx_paragraph_in_a_style_less_package_gets_no_eojeol_frames() {
    let mut paragraph = korean_paragraph(EOJEOL_SENTENCE, Some(Alignment::Justify), None);
    if let Block::Paragraph(ref mut p) = paragraph {
        p.style.word_wrap = Some(false);
    }
    let doc = make_doc(vec![make_flow_page(vec![paragraph])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(EOJEOL_SENTENCE),
        "the text stays one run so Typst may break inside an eojeol: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no eojeol frame may be emitted when wordWrap is off: {result}"
    );
}

#[test]
fn a_slide_paragraph_keeps_syllable_breaking() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![make_text_box(10.0, 10.0, 300.0, 100.0, EOJEOL_SENTENCE)],
    )]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(EOJEOL_SENTENCE),
        "PowerPoint breaks Korean mid-word, and our slide output matches it: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no eojeol frame may reach a slide: {result}"
    );
}

#[test]
fn a_sheet_cell_keeps_syllable_breaking() {
    let doc = make_doc(vec![make_sheet_page(
        "Sheet1",
        595.0,
        842.0,
        Margins::default(),
        make_simple_table(vec![vec![EOJEOL_SENTENCE]]),
    )]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(EOJEOL_SENTENCE),
        "a spreadsheet cell keeps Excel's own breaking: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no eojeol frame may reach a sheet cell: {result}"
    );
}

#[test]
fn a_docx_table_cell_keeps_each_hangul_eojeol_whole() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell(EOJEOL_SENTENCE)],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("본 #box[계약은] #box[갑과] #box[을이]"),
        "Word breaks a table cell's Korean at eojeol too: {result}"
    );
}

#[test]
fn latin_text_is_untouched() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "The parties agree to cooperate",
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("The parties agree to cooperate"),
        "Latin already breaks at spaces and needs no frame: {result}"
    );
    assert!(!result.contains("#box["), "no frame for Latin: {result}");
}

#[test]
fn only_the_tokens_carrying_hangul_are_framed() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "2026년 VAT 별도 API 연동",
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#box[2026년] VAT #box[별도] API #box[연동]"),
        "a Latin/digit token keeps its own break opportunities: {result}"
    );
}

/// The auto space of issue #521 marks a boundary *inside* one eojeol, and
/// nothing may break there. It used to hold that by sitting inside the eojeol
/// frame, because Typst maps every `#h()` to a space in the paragraph text;
/// a frame is laid out at its natural width, though, so the gap could take no
/// part in a justified line's stretch (issue #1193). It is emitted outside
/// the frame now, and what forbids the break is the character itself: U+00A0
/// is Unicode line-break class GL, which allows no break on either side.
#[test]
fn the_east_asian_auto_space_cannot_host_a_break() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "2026\u{E001}년 계약",
        None,
        Some("Malgun Gothic"),
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("#h(2.625pt)"),
        "a spacer is a breakable space in the paragraph text: {result}"
    );
    assert!(
        result.contains("spacing: 2.625pt)[\\u{00A0}]"),
        "the boundary takes a quarter-em no-break space: {result}"
    );
}

/// Word spreads a justified line's stretch demand over the East Asian/Latin
/// auto spaces as well as the word spaces, and past half an em per gap it
/// sets every expandable gap on the line to one common width (issue #1053).
/// A rigid `#h()` could take none of it, so the paragraph states the ceiling
/// that lets Typst's justifier reproduce that (issue #1193).
#[test]
fn a_justified_paragraph_caps_its_gaps_where_word_does() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "2026\u{E001}년 계약 조건",
        Some(Alignment::Justify),
        Some("Malgun Gothic"),
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#set par(justification-limits: (spacing: (min: 80%, max: 0.0001% + 0.5em)))"
        ),
        "every gap on the line stops at its own half em: {result}"
    );
}

/// Word's phase 3, measured on `korean_alignment_autospace.docx`: the first
/// line of the wrapping justified paragraph stretches 55.06pt to a 453.61pt
/// measure, and its native Word export puts all sixteen gaps — eleven word
/// spaces at 6.8065pt and five East Asian/Latin auto spaces at 6.8028pt — at
/// one common width. Before issue #1193 the whole demand landed in the word
/// spaces: 8.70pt each, with the auto spaces left at their 2.62pt quarter em,
/// which displaced the `자` of `제3자` by 9.50pt.
#[cfg(not(target_arch = "wasm32"))]
fn host_can_shape_korean_family(family: &str) -> bool {
    let context = crate::render::font_context::resolve_font_search_context(&[]);
    let candidates: Vec<String> =
        crate::render::font_subst::with_font_search_context(Some(&context), || {
            crate::render::font_subst::family_candidates(family)
        });
    candidates.iter().any(|candidate| {
        context.covers_script(candidate, crate::render::font_subst::TextScript::Korean)
    })
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_heavily_stretched_justified_line_gives_every_gap_one_width() {
    // The fixture is set in Malgun Gothic throughout. If neither that family
    // nor one of its metric candidates carries Hangul, the ASCII spaces can
    // still resolve through the Latin painting tail while the eojeols cannot
    // be shaped. That split line is not the Word line this test measures.
    if !host_can_shape_korean_family("Malgun Gothic") {
        return;
    }

    let data = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/korean_alignment_autospace.docx"
    ))
    .expect("fixture");
    let (doc, _warnings) = crate::parser::Parser::parse(
        &crate::parser::docx::DocxParser,
        &data,
        &crate::config::ConvertOptions::default(),
    )
    .expect("parse");
    let source = generate_typst(&doc).unwrap().source;
    let runs = crate::render::pdf::compiled_text_runs(&source, 0).expect("compile");

    // The stretched line is the wrapping paragraph's first, which opens on
    // `본` — the one token of the file that appears nowhere else. A run's
    // baseline names the line it sits on.
    let baseline: f64 = runs
        .iter()
        .find(|run| run.text.contains('본'))
        .expect("the wrapping paragraph's first line")
        .baseline_pt;
    let mut line: Vec<&crate::render::pdf::PlacedTextRun> = runs
        .iter()
        .filter(|run| (run.baseline_pt - baseline).abs() < 0.01)
        .collect();
    line.sort_by(|left, right| left.left_pt.total_cmp(&right.left_pt));

    // A gap is a run holding nothing but a space, and the width the line gave
    // it is the distance to whatever it sits before.
    let mut word_gaps: Vec<f64> = Vec::new();
    let mut auto_gaps: Vec<f64> = Vec::new();
    for (index, run) in line.iter().enumerate() {
        let Some(next) = line.get(index + 1) else {
            break;
        };
        let width: f64 = next.left_pt - run.left_pt;
        match run.text.as_str() {
            " " => word_gaps.push(width),
            "\u{00A0}" => auto_gaps.push(width),
            _ => {}
        }
    }

    // A word space is a run of its own only where the eojeols around it are
    // framed. The coverage guard proves the chain can shape the Hangul; this
    // second guard keeps the measurement honest if the compiled line still
    // contains no separately placed word-space runs.
    if word_gaps.is_empty() {
        return;
    }
    assert!(
        !auto_gaps.is_empty(),
        "the line's auto spaces are gaps the justifier can see"
    );
    let widest: f64 = word_gaps
        .iter()
        .chain(&auto_gaps)
        .copied()
        .fold(f64::MIN, f64::max);
    let narrowest: f64 = word_gaps
        .iter()
        .chain(&auto_gaps)
        .copied()
        .fold(f64::MAX, f64::min);
    assert!(
        widest - narrowest < 0.05,
        "every gap on the line takes one common width, \
         but they run {narrowest:.4}pt to {widest:.4}pt"
    );

    // 6.80pt is Word's width for *this* face at 10.5pt. A runner that
    // substitutes another Korean face stretches the same line over its own
    // advances, and only the shared width above carries over.
    if line.iter().any(|run| run.family == "Malgun Gothic") {
        assert!(
            (narrowest - 6.80).abs() < 0.05,
            "Word's common width here is 6.80pt, not {narrowest:.4}pt"
        );
    }
}

/// The ceiling is Word's answer for a line carrying the auto space. Typst's
/// line breaker prices a line against the allowance it is given, so stating
/// it where no auto space exists would move breaks in ordinary justified
/// paragraphs for nothing.
#[test]
fn a_justified_paragraph_without_the_auto_space_keeps_the_document_ceiling() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "계약 조건 확인",
        Some(Alignment::Justify),
        Some("Malgun Gothic"),
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("justification-limits").count(),
        1,
        "only the document-wide rule states a ceiling: {result}"
    );
}

/// A frame seats its baseline on its own bottom edge, so under the fixed text
/// edges Word's line model needs (issues #354, #508) the framed text would sink
/// by the descent. The frame restores both edges and shifts its baseline back.
#[test]
fn a_framed_eojeol_keeps_the_paragraphs_baseline() {
    // The correction exists only under a fixed line box, and the paragraph
    // derives that box from its Korean face's own metrics. Without one — every
    // CI runner here — there is no box to restore and nothing to assert.
    if crate::render::pdf::font_line_metrics_em("Malgun Gothic").is_none() {
        return; // no Korean face available (e.g. a runner with no CJK fonts)
    }
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        EOJEOL_SENTENCE,
        None,
        Some("Malgun Gothic"),
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    let (top_em, bottom_em) =
        emitted_line_box_em(&result).expect("a Korean paragraph declares a fixed line box");
    let top_pt: f64 = top_em * 10.5;
    let bottom_pt: f64 = bottom_em * 10.5;
    let expected: String = format!(
        "#box(baseline: {}pt)[#text(top-edge: {}pt, bottom-edge: -{}pt)[",
        format_f64(bottom_pt),
        format_f64(top_pt),
        format_f64(bottom_pt)
    );
    assert!(
        result.contains(&expected),
        "the frame should restore the line box and shift the baseline back by the descent\n\
         expected: {expected}\nin: {result}"
    );
    assert_eq!(
        result.matches(&expected).count(),
        3,
        "every multi-syllable eojeol should carry the correction: {result}"
    );
}

/// Triangulation: the shift is the paragraph's own descent, not a constant.
#[test]
fn the_frames_baseline_shift_scales_with_the_font_size() {
    // Same premise as the test above: the shift is the fixed line box's own
    // descent, and that box needs the Korean face's measured metrics.
    if crate::render::pdf::font_line_metrics_em("Malgun Gothic").is_none() {
        return; // no Korean face available (e.g. a runner with no CJK fonts)
    }
    let mut shifts: Vec<String> = Vec::new();
    for size in [10.5_f64, 20.0_f64] {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: EOJEOL_SENTENCE.to_string(),
                style: TextStyle {
                    font_family: Some("Malgun Gothic".to_string()),
                    east_asian_font_family: Some("Malgun Gothic".to_string()),
                    font_size: Some(size),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        let (_top_em, bottom_em) = emitted_line_box_em(&result).expect("fixed line box");
        let expected: String = format!("#box(baseline: {}pt)", format_f64(bottom_em * size));
        assert!(
            result.contains(&expected),
            "a {size}pt paragraph should shift by {expected}: {result}"
        );
        shifts.push(expected);
    }
    assert_ne!(
        shifts[0], shifts[1],
        "the shift must not be a single measured constant"
    );
}

/// Letter spacing crosses a frame boundary by a rule that is not one step per
/// item — measured on typst 0.14, framing a 13pt tracked heading's words made
/// it narrower and a 9pt one's wider — so a tracked run keeps today's
/// emission rather than a guessed correction.
#[test]
fn a_letter_spaced_run_is_not_framed() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "활용 설치부터".to_string(),
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
        result.contains("#text(tracking: 0.5pt, ligatures: false, kerning: false)[활용 설치부터]"),
        "a tracked run stays one text item: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no frame for tracked text: {result}"
    );
}

/// Issue #1023: the PowerPoint advance-grid path splits words and spaces into
/// separate Typst items, and Typst trims tracking at every item boundary, so
/// the #841 deck's tracked footer lost ~4pt per word gap while every
/// intra-word advance matched the native export. A tracked slide run keeps
/// one shaped item instead, the same trade the framed-eojeol exemption makes.
#[test]
fn a_letter_spaced_slide_run_declines_the_advance_grid() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![FixedElement {
            x: 69.84,
            y: 500.0,
            width: 400.0,
            height: 30.0,
            kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "CONTOSO ALLE ANSATTE".to_string(),
                        style: TextStyle {
                            letter_spacing: Some(2.0),
                            ..TextStyle::default()
                        },
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
        }],
    )]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("#o2p-pptx-word(") && !result.contains("#o2p-pptx-space()"),
        "a tracked slide run must not be split into grid items: {result}"
    );
    assert!(
        result.contains("CONTOSO ALLE ANSATTE"),
        "the run stays one shaped item: {result}"
    );
}

#[test]
fn an_untracked_slide_run_still_takes_the_advance_grid() {
    // Triangulation: the exclusion keys on the spacing, not on the page kind.
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_text_box(
            69.84,
            500.0,
            400.0,
            30.0,
            "CONTOSO ALLE ANSATTE",
        )],
    )]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#o2p-pptx-word([CONTOSO]"),
        "an untracked slide run keeps the 1/8pt grid: {result}"
    );
}

#[test]
fn an_unspaced_eojeol_is_still_framed() {
    // Triangulation: the exclusion must key on the spacing, not on the words.
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "활용 설치부터",
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#box[활용] #box[설치부터]"),
        "an unspaced paragraph is framed as usual: {result}"
    );
}

/// Cutting a run at an eojeol boundary can leave a date at the start of the
/// next escaping unit, where Typst reads `2026. ` as an enumeration marker and
/// puts the date on a line of its own — which is what happened to the official
/// letter's `시행일자: 2026. 7. 17.`.
#[test]
fn a_date_after_an_eojeol_is_not_retypeset_as_a_list_item() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "시행일자: 2026. 7. 17.",
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#box[시행일자:] 2026\\. 7. 17."),
        "the date's first dot must be escaped once the frame precedes it: {result}"
    );
}

/// The same hazard without any frame: whichever way a run is cut, the one
/// leading space this function emits literally must not hide the marker.
#[test]
fn a_leading_space_does_not_hide_an_enum_marker() {
    assert_eq!(escape_typst(" 2026. 7. 17."), " 2026\\. 7. 17.");
    assert_eq!(escape_typst("2026. 7. 17."), "2026\\. 7. 17.");
    assert_eq!(
        escape_typst(" 2026 7 17"),
        " 2026 7 17",
        "a bare number is not a marker and must not gain an escape"
    );
    // An indentation run leaves as a code-mode string, which cannot open an
    // enumeration, so it must not gain an escape it does not need.
    assert!(
        escape_typst("      2. indented").ends_with("2. indented"),
        "an indented number keeps its plain dot"
    );
}

/// A token no line could hold would take a frame of its own and start a new
/// line, costing a line Word does not spend. Such a token is not an eojeol.
#[test]
fn a_pathologically_long_token_is_not_framed() {
    let long_token: String = "가".repeat(40);
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        &format!("계약 {long_token} 종료"),
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(&format!("#box[계약] {long_token} #box[종료]")),
        "an over-long token keeps today's syllable breaking: {result}"
    );
}

/// Word keeps a Latin token whole while it fits a table-cell line, then
/// exposes character boundaries only when the token is wider than the cell.
/// The break must not add an invisible character to the PDF text layer, and a
/// split that crosses run boundaries must retain every run's link and style.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_docx_table_cell_breaks_only_an_overlong_latin_token() {
    const FAMILY: &str = "Office2pdf Missing Hand Fixture";
    const SIZE_PT: f64 = 66.0;
    let family_chain = vec![FAMILY.to_string()];
    let welc_width =
        crate::render::pdf::glyph_advances_em_with_typst_fallback(&family_chain, false, "WELC")
            .expect("Typst's fallback face should measure")
            .iter()
            .sum::<f64>()
            * SIZE_PT;
    let welco_width =
        crate::render::pdf::glyph_advances_em_with_typst_fallback(&family_chain, false, "WELCO")
            .expect("Typst's fallback face should measure")
            .iter()
            .sum::<f64>()
            * SIZE_PT;
    let available_width = (welc_width + welco_width) / 2.0;
    let linked_style = |color| TextStyle {
        font_family: Some(FAMILY.to_string()),
        font_size: Some(SIZE_PT),
        color: Some(color),
        ..TextStyle::default()
    };
    let href = Some("https://example.com/event".to_string());
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                line_spacing: Some(LineSpacing::Proportional(0.8)),
                ..ParagraphStyle::default()
            },
            runs: vec![
                Run {
                    text: "AGES ".to_string(),
                    style: linked_style(Color::black()),
                    href: None,
                    footnote: None,
                },
                Run {
                    text: "WEL".to_string(),
                    style: linked_style(Color::new(255, 0, 0)),
                    href: href.clone(),
                    footnote: None,
                },
                Run {
                    text: "CO".to_string(),
                    style: linked_style(Color::new(0, 0, 255)),
                    href: href.clone(),
                    footnote: None,
                },
                Run {
                    text: "ME".to_string(),
                    style: linked_style(Color::new(0, 128, 0)),
                    href,
                    footnote: None,
                },
            ],
        })],
        vertical_align: Some(CellVerticalAlign::Center),
        padding: Some(Insets::default()),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: Some(333.808),
            minimum_height: None,
        }],
        column_widths: vec![available_width],
        default_cell_padding: Some(Insets::default()),
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let output = generate_typst(&doc).unwrap();

    assert_eq!(
        output.source.matches("#linebreak()").count(),
        1,
        "the overlong token should gain one measured break: {}",
        output.source
    );
    assert!(
        !output.source.contains("AGE#linebreak()"),
        "the fitting AGES token must stay whole: {}",
        output.source
    );
    assert_eq!(
        output.source.matches("https://example.com/event").count(),
        4,
        "both halves of the split run must retain their hyperlink: {}",
        output.source
    );
    assert!(
        !output.source.contains(['\u{200B}', '\u{2060}', '\u{00A0}']),
        "no hidden break character may reach the generated source: {}",
        output.source
    );

    let mut placed = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
    placed.sort_by(|left, right| {
        left.baseline_pt
            .total_cmp(&right.baseline_pt)
            .then_with(|| left.left_pt.total_cmp(&right.left_pt))
    });
    let extracted: String = placed.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(
        extracted, "AGES WELCOME",
        "the break object must not change the searchable characters"
    );
    let mut lines: Vec<(f64, String)> = Vec::new();
    for run in placed {
        match lines.last_mut() {
            Some((baseline, text)) if (run.baseline_pt - *baseline).abs() < 0.1 => {
                text.push_str(&run.text);
            }
            _ => lines.push((run.baseline_pt, run.text)),
        }
    }
    assert_eq!(
        lines
            .iter()
            .map(|(_, text)| text.trim().to_string())
            .collect::<Vec<_>>(),
        vec!["AGES", "WELC", "OME"],
        "the overlong token should wrap at the measured character boundary: {lines:#?}"
    );
    let pitches: Vec<f64> = lines.windows(2).map(|pair| pair[1].0 - pair[0].0).collect();
    assert!(
        pitches.iter().all(|pitch| *pitch > SIZE_PT),
        "splitting styled inline items must retain the original line metrics: {pitches:?}"
    );
    assert!(
        pitches
            .windows(2)
            .all(|pair| (pair[1] - pair[0]).abs() < 0.1),
        "every split line must keep the same pitch: {pitches:?}"
    );
}

/// A table cell whose single paragraph names a family and a size, so the
/// eojeol width guard has metrics to measure against (issue #626).
fn make_text_cell_styled(text: &str, family: &str, font_size: f64) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle {
                    font_family: Some(family.to_string()),
                    east_asian_font_family: Some(family.to_string()),
                    font_size: Some(font_size),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })],
        ..TableCell::default()
    }
}

/// A slide text box whose content is a one-item bullet list.
fn make_text_box_with_list(x: f64, y: f64, w: f64, h: f64, text: &str) -> FixedElement {
    let mut element: FixedElement = make_text_box(x, y, w, h, text);
    if let FixedElementKind::TextBox(ref mut data) = element.kind {
        data.content = vec![Block::List(List {
            kind: ListKind::Unordered,
            items: vec![ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: text.to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            }],
            level_styles: BTreeMap::new(),
        })];
    }
    element
}

// Typst line-leading markup at an eojeol boundary (issue #626)

/// Cutting a run at an eojeol boundary makes the inter-word text its own
/// escaping unit, so a bare ` + ` reaches `escape_typst` at the start of a
/// content block — where Typst reads it as a numbered-list marker, deletes the
/// `+` from the page and puts a `1.` in the text layer instead.
#[test]
fn a_plus_between_two_eojeol_is_not_retypeset_as_a_list_item() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "런타임 초기화 + 프로필 생성",
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#box[초기화] \\+ #box[프로필]"),
        "the `+` must be escaped once a frame precedes it: {result}"
    );
}

/// The same hazard for `=`, which Typst reads as a heading marker and which
/// was not in the escape set at all.
#[test]
fn an_equals_between_two_eojeol_is_not_retypeset_as_a_heading() {
    let doc = make_doc(vec![make_flow_page(vec![korean_paragraph(
        "부하 시험 = 결과 보고",
        None,
        None,
    )])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#box[시험] \\= #box[결과]"),
        "the `=` must be escaped once a frame precedes it: {result}"
    );
}

/// The whole set of Typst markup that is only meaningful at a line start,
/// exercised directly: through one leading space and at index 0, and only when
/// the marker is really one.
#[test]
fn every_line_leading_marker_is_neutralised_through_one_space() {
    // Bullet, numbered list, heading — the three that need the positional rule.
    assert_eq!(escape_typst(" + "), " \\+ ");
    assert_eq!(escape_typst("+ x"), "\\+ x");
    assert_eq!(escape_typst(" = "), " \\= ");
    assert_eq!(escape_typst("= x"), "\\= x");
    assert_eq!(
        escape_typst(" == "),
        " \\== ",
        "escaping the first equals is enough to break a level-2 heading"
    );
    assert_eq!(escape_typst(" - "), " \\- ");
    assert_eq!(escape_typst(" 2. x"), " 2\\. x");
    // A term list opens with `/`, which is escaped wherever it appears.
    assert_eq!(escape_typst(" / term: x"), " \\/ term: x");

    // Not markers: no trailing whitespace, or a leading run of two spaces that
    // leaves as a code-mode string.
    assert_eq!(escape_typst(" =x"), " =x");
    assert_eq!(escape_typst("+"), "+");
    assert_eq!(escape_typst("a = b"), "a = b");
    assert!(
        escape_typst("  = x").ends_with("= x"),
        "an indented equals keeps its plain form: {}",
        escape_typst("  = x")
    );
}

/// A DOCX list item is a Word paragraph like any other, so its Korean breaks
/// at eojeol too. It used to bypass the whole path by calling `generate_run`
/// directly.
#[test]
fn a_docx_list_item_keeps_each_hangul_eojeol_whole() {
    let doc = make_doc(vec![make_flow_page(vec![Block::List(List {
        kind: ListKind::Unordered,
        items: vec![ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: EOJEOL_SENTENCE.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }],
            level: 0,
            start_at: None,
        }],
        level_styles: BTreeMap::new(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("본 #box[계약은] #box[갑과] #box[을이]"),
        "a list item's Korean should be framed like any other paragraph: {result}"
    );
}

/// A slide's list keeps PowerPoint's own mid-word breaking.
#[test]
fn a_slide_list_item_keeps_syllable_breaking() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![make_text_box_with_list(
            10.0,
            10.0,
            300.0,
            100.0,
            EOJEOL_SENTENCE,
        )],
    )]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(EOJEOL_SENTENCE),
        "PowerPoint breaks a bullet's Korean mid-word: {result}"
    );
    assert!(
        !result.contains("#box["),
        "no eojeol frame may reach a slide list: {result}"
    );
}

/// A token wider than the column cannot break inside its frame, so the frame
/// would take a line of its own and still overflow it. Word breaks such a
/// token at character level, and so must we.
#[test]
fn a_token_wider_than_its_column_is_not_framed() {
    // 20 syllables at 10.5pt Malgun Gothic is 210pt — far wider than the
    // 150pt column, and short enough that a character ceiling would let it
    // through.
    let long_token: String = "가나다라마바사아자차카타파하가나다라마바".to_string();
    // The premise is that the token *measures* over the column, so the guard
    // needs the same advance the generator does. Without a Korean face — every
    // CI runner here — nothing measures and the generator falls back to its
    // character ceiling, which this deliberately 20-character token clears.
    if crate::render::pdf::text_advance_em("Malgun Gothic", false, &long_token).is_none() {
        return; // no Korean face available (e.g. a runner with no CJK fonts)
    }
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell_styled(
                &format!("주 {long_token}"),
                "Malgun Gothic",
                10.5,
            )],
            height: None,
        }],
        column_widths: vec![150.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains(&format!("[{long_token}]]")),
        "an over-wide token must keep the engine's syllable breaking: {result}"
    );
}

/// Triangulation for the guard above: the same token in a column wide enough
/// to hold it *is* framed, so the rule keys on width and not on the token.
///
/// Deliberately unguarded, unlike its partner: without a Korean face the token
/// is framed by the character ceiling instead of by the width rule, so the
/// assertion still holds and keeps guarding "a frame is emitted at all" on a
/// runner with no CJK fonts. Only the *width* half of the triangulation needs
/// the face, and that half lives in the test above.
#[test]
fn the_same_token_is_framed_when_the_column_can_hold_it() {
    let long_token: String = "가나다라마바사아자차카타파하가나다라마바".to_string();
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![make_text_cell_styled(
                &format!("주 {long_token}"),
                "Malgun Gothic",
                10.5,
            )],
            height: None,
        }],
        column_widths: vec![400.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(&format!("[{long_token}]]")),
        "a token the column can hold keeps its frame: {result}"
    );
}

// ----- Pair kerning (issue #628) -----

/// A DOCX-shaped document: `w:docDefaults` resolved to a run style, which
/// carries the kerning decision every run inherits.
fn make_doc_with_default_text(pages: Vec<Page>, default_text: TextStyle) -> Document {
    Document {
        metadata: Metadata::default(),
        pages,
        styles: StyleSheet {
            default_text: Some(default_text),
            ..StyleSheet::default()
        },
    }
}

fn styled_paragraph(text: &str, style: TextStyle) -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: text.to_string(),
            style,
            href: None,
            footnote: None,
        }],
    })
}

#[test]
fn test_unkerned_run_states_the_decision_on_itself_not_document_wide() {
    // Word writes no `w:kern` in the business mocks, so it sets every glyph at
    // its nominal advance. The decision travels on the run rather than as a
    // document-wide `#set text(kerning: false)`: that rule would also reach
    // the list markers and header fields whose text the emitter cannot name,
    // and RTL text under it loses glyphs to typst 0.14.2's shaping defect
    // (issue #628 review, defect 1).
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![styled_paragraph(
            "Body text",
            TextStyle {
                pair_kerning: Some(PairKerning::Never),
                ..TextStyle::default()
            },
        )])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        !source.contains("#set text(kerning: false)"),
        "no document-wide kerning rule may be emitted: {source}"
    );
    assert!(
        source.contains("#text(kerning: false)[Body text]"),
        "a run with no other property still states its own kerning: {source}"
    );
}

#[test]
fn test_run_without_kern_element_emits_kerning_false() {
    // The masthead of 08_newsletter_en: 22pt Arial Bold, centred, which Word
    // does not kern. Leaving the OpenType feature on pulled `LY` in by 2.02pt.
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "THE MONTHLY RENDER",
        TextStyle {
            font_family: Some("Arial".to_string()),
            font_size: Some(22.0),
            bold: Some(true),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: false"),
        "an unkerned run must state it on its own #text(): {source}"
    );
}

#[test]
fn test_run_at_or_above_kern_threshold_keeps_kerning() {
    // `w:kern w:val="32"` is 16pt. A 20pt title is at or above it, so Word
    // kerns it and the generator must not switch the feature off; the 11pt
    // body under the same threshold is below it and must be left alone.
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![
            styled_paragraph(
                "JAMIE PARKER",
                TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(20.0),
                    bold: Some(true),
                    pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
                    ..TextStyle::default()
                },
            ),
            styled_paragraph(
                "Product designer",
                TextStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(11.0),
                    pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
                    ..TextStyle::default()
                },
            ),
        ])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("size: 20pt, weight: \"bold\", kerning: true"),
        "a run at or above the threshold keeps kerning: {source}"
    );
    assert!(
        source.contains("size: 11pt, kerning: false"),
        "body text below the threshold stays unkerned: {source}"
    );
}

#[test]
fn test_run_below_kern_threshold_disables_kerning() {
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "Body copy",
        TextStyle {
            font_family: Some("Arial".to_string()),
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: false"),
        "a run below the threshold must not be kerned: {source}"
    );
}

#[test]
fn test_format_without_kerning_model_emits_no_kerning_parameter() {
    // XLSX states no threshold at all, and a PPTX run inherits `None` until a
    // `kern` turns up on its own `a:rPr` or a list style above it — an
    // unstated rule must leave the engine's own default standing.
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "Slide title",
        TextStyle {
            font_family: Some("Arial".to_string()),
            font_size: Some(28.0),
            bold: Some(true),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        !source.contains("kerning"),
        "a format that states no kerning rule must emit none: {source}"
    );
}

#[test]
fn test_rtl_run_keeps_kerning_despite_the_word_rule() {
    // typst 0.14.2 mis-orders RTL glyph ranges when the `kern` feature is
    // off, so Word's rule is not applied to a document that shapes
    // right-to-left — see `with_rtl_shaping_exemption`.
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![styled_paragraph(
            "مرحبا بالعالم",
            TextStyle {
                font_family: Some("Arial".to_string()),
                pair_kerning: Some(PairKerning::Never),
                ..TextStyle::default()
            },
        )])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: true"),
        "an RTL run must override the document-wide kerning: false: {source}"
    );
}

#[test]
fn test_bare_rtl_run_is_never_left_under_a_kerning_false() {
    // A run with no other text property is emitted bare, so nothing may switch
    // kerning off around it: in a document Word does not kern, the Hebrew run
    // below must still reach the engine with the feature on.
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![
            styled_paragraph(
                "שלום עולם",
                TextStyle {
                    pair_kerning: Some(PairKerning::Never),
                    ..TextStyle::default()
                },
            ),
            styled_paragraph(
                "Latin body copy",
                TextStyle {
                    pair_kerning: Some(PairKerning::Never),
                    ..TextStyle::default()
                },
            ),
        ])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    let hebrew_line: &str = source
        .lines()
        .find(|line| line.contains("שלום"))
        .expect("the Hebrew run is emitted");
    assert!(
        !hebrew_line.contains("kerning: false"),
        "the RTL run must not be wrapped in kerning: false: {hebrew_line}"
    );
    assert!(
        !source.contains("#set text(kerning: false)"),
        "and nothing document-wide may switch it off either: {source}"
    );
    // The Latin run travels with it: bidi reordering is decided over a whole
    // shaped paragraph, and a run's own codepoints do not bound the scope the
    // exemption has to cover. See
    // `test_neutral_run_beside_an_rtl_run_keeps_kerning`.
    assert!(
        !source.contains("kerning: false"),
        "no run in a document carrying RTL may state kerning: false: {source}"
    );
}

/// A paragraph whose runs are given verbatim, so a test can mix scripts inside
/// one shaped bidi paragraph.
fn paragraph_of_runs(style: ParagraphStyle, texts: &[(&str, TextStyle)]) -> Block {
    Block::Paragraph(Paragraph {
        style,
        runs: texts
            .iter()
            .map(|(text, run_style)| Run {
                text: (*text).to_string(),
                style: run_style.clone(),
                href: None,
                footnote: None,
            })
            .collect(),
    })
}

#[test]
fn test_neutral_run_beside_an_rtl_run_keeps_kerning() {
    // typst shapes a whole bidi *paragraph*, not a run: the runs around a
    // right-to-left one are reordered with it, so a run whose own codepoints
    // are all Latin or neutral still reaches the shaper as a right-to-left
    // segment. Disabling `kern` there is what inverts the glyph ranges, and
    // krilla then panics building the text group ("byte range starts at 3 but
    // ends at 0"). The exemption therefore cannot be decided from one run's
    // own text (issue #628 follow-up).
    let unkerned = TextStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(14.0),
        pair_kerning: Some(PairKerning::Never),
        ..TextStyle::default()
    };
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![paragraph_of_runs(
            ParagraphStyle::default(),
            &[
                ("We met a girl ", unkerned.clone()),
                ("مرحبا", unkerned.clone()),
                (" whose father was a diver.", unkerned.clone()),
            ],
        )])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        !source.contains("kerning: false"),
        "no run of a paragraph carrying RTL may state kerning: false: {source}"
    );
}

#[test]
fn test_rtl_paragraph_direction_keeps_kerning_on_neutral_text() {
    // FDO76312.docx's cells: `w:bidi` makes the paragraph's base direction
    // right-to-left, so its neutral characters — an ellipsis run, a row of
    // full stops — take an RTL bidi level even though no strong RTL codepoint
    // appears anywhere in the file. Two such characters in a row are enough to
    // trip the shaping defect, so the direction has to be read as well as the
    // text (issue #628 follow-up).
    let unkerned = TextStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(14.0),
        pair_kerning: Some(PairKerning::Never),
        ..TextStyle::default()
    };
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![paragraph_of_runs(
            ParagraphStyle {
                direction: Some(TextDirection::Rtl),
                ..ParagraphStyle::default()
            },
            &[("She taught .......... how to use a computer.", unkerned)],
        )])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("dir: rtl"),
        "the paragraph is emitted right-to-left: {source}"
    );
    assert!(
        !source.contains("kerning: false"),
        "an RTL-directed paragraph keeps kerning on its neutral text: {source}"
    );
}

#[test]
fn test_latin_run_is_not_exempted_from_the_word_rule() {
    // The RTL exemption is a shaping workaround, not a licence to keep
    // kerning: every other script must still follow Word.
    for text in ["JAMIE PARKER", "안녕하세요 세계", "你好世界", "สวัสดีชาวโลก"]
    {
        let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
            text,
            TextStyle {
                font_family: Some("Arial".to_string()),
                font_size: Some(20.0),
                pair_kerning: Some(PairKerning::Never),
                ..TextStyle::default()
            },
        )])]);

        let source = generate_typst(&doc).unwrap().source;

        assert!(
            source.contains("kerning: false"),
            "{text} must follow Word's rule: {source}"
        );
    }
}

/// An Arabic list under a document Word does not kern.
fn arabic_list_source(kind: crate::ir::ListKind, marker_text: Option<&str>) -> String {
    use crate::ir::{ListItem, ListLevelStyle};

    let item_style = TextStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(11.0),
        pair_kerning: Some(PairKerning::Never),
        ..TextStyle::default()
    };
    let mut level_styles = std::collections::BTreeMap::new();
    if let Some(marker_text) = marker_text {
        level_styles.insert(
            0,
            ListLevelStyle {
                kind,
                numbering_pattern: None,
                full_numbering: false,
                marker_text: Some(marker_text.to_string()),
                marker_style: Some(item_style.clone()),
            },
        );
    }
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![Block::List(List {
            kind,
            items: vec![ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "بند أول".to_string(),
                        style: item_style,
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            }],
            level_styles,
        })])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    generate_typst(&doc).unwrap().source
}

#[test]
fn test_ordered_list_marker_takes_the_safe_kerning_answer() {
    // The marker's text is `#numbering`'s result — the emitter cannot know
    // which script the pattern produces, so it must not switch kerning off
    // around it (issue #628 review, defect 1).
    let source: String = arabic_list_source(crate::ir::ListKind::Ordered, None);

    let marker: &str = source
        .split_once("numbering: (..nums) => [")
        .expect("an ordered list emits a numbering function")
        .1;
    let marker_params: &str = marker
        .split_once("#numbering(")
        .expect("the marker wraps a numbering call")
        .0;
    assert!(
        marker_params.contains("kerning: true"),
        "the numbering marker keeps kerning: {marker_params}"
    );
}

#[test]
fn test_rtl_list_marker_keeps_kerning() {
    // A marker whose own text is right-to-left is the case that dropped
    // glyphs: it has to be read as RTL and left kerned.
    let source: String = arabic_list_source(crate::ir::ListKind::Unordered, Some("أولاً"));

    assert!(
        !source.contains("kerning: false"),
        "nothing in an RTL-marked list may switch kerning off: {source}"
    );
    assert!(
        source.contains("أولاً"),
        "the marker survives into the source: {source}"
    );
}

#[test]
fn test_list_item_runs_keep_the_rtl_exemption() {
    let source: String = arabic_list_source(crate::ir::ListKind::Unordered, None);

    let item_line: &str = source
        .lines()
        .find(|line| line.contains("بند"))
        .expect("the item text is emitted");
    assert!(
        item_line.contains("kerning: true"),
        "an RTL list item keeps kerning: {item_line}"
    );
}

#[test]
fn test_header_field_never_states_kerning_false() {
    // A page-number field's text is the engine's, in whatever numbering format
    // the section states, so the emitter cannot name it and must not switch
    // kerning off around it.
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let field_style = TextStyle {
        font_size: Some(8.0),
        pair_kerning: Some(PairKerning::Never),
        ..TextStyle::default()
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::PageNumber(field_style)],
                border: None,
                border_space: None,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: true"),
        "the field takes the script-safe answer: {source}"
    );
    assert!(
        !source.contains("kerning: false"),
        "and never the one that can drop glyphs: {source}"
    );
}

#[test]
fn test_unembedded_aptos_footer_uses_noto_sans_on_an_aptos_host() {
    // Reduced from `Place your event title here.docx`: the footer fixes its
    // face and size, but the package carries no font data. A host Office
    // installation must not make this fixed footer narrower than the
    // LibreOffice/Noto Sans ground truth (issue #1463).
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(24.0),
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Sensitivity: Internal".to_string(),
                    style: TextStyle {
                        font_family: Some("Aptos".to_string()),
                        font_size: Some(8.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let context = FontSearchContext::for_test(Vec::new(), &["Aptos", "Noto Sans"], &["Aptos"], &[]);

    let source = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap()
    .source;

    assert!(
        source.contains("font: (\"Noto Sans\""),
        "the fixed footer must lead with the reference-compatible family: {source}"
    );
    assert!(
        !source.contains("font: (\"Aptos\""),
        "the incidental host face must not remain first: {source}"
    );
}

// ── Synthetic oblique for faces with no italic (issue #686) ──────

use crate::render::font_context::FontSearchContext;
use crate::render::font_subst::{TextScript, with_font_search_context};

/// A context where Calibri ships an italic face and Malgun Gothic does not —
/// the split the issue measured against a native Word and PowerPoint export.
fn calibri_italic_malgun_upright_context() -> FontSearchContext {
    FontSearchContext::for_test(Vec::new(), &["Calibri", "Malgun Gothic"], &[], &[])
        .with_italic_and_scripts(
            &["Calibri"],
            &[
                ("Calibri", &[TextScript::Latin]),
                ("Malgun Gothic", &[TextScript::Latin, TextScript::Korean]),
            ],
        )
}

fn italic_run(text: &str, family: &str) -> Run {
    Run {
        text: text.to_string(),
        style: TextStyle {
            font_family: Some(family.to_string()),
            font_size: Some(14.0),
            italic: Some(true),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    }
}

fn generated_source(runs: &[Run], context: &FontSearchContext) -> String {
    with_font_search_context(Some(context), || {
        let mut out = String::new();
        generate_runs(&mut out, runs, EojeolWrap::Syllable);
        out
    })
}

#[test]
fn test_italic_on_face_without_italic_emits_a_synthetic_oblique() {
    // Word slants a Malgun Gothic `<w:i/>` run itself — the measured text
    // matrix is `trm="38 0 12.91406 38"`, a 0.340 slope. Typst has no such
    // fallback, so without this the emphasis vanishes silently (issue #686).
    let source = generated_source(
        &[italic_run("한국어", "Malgun Gothic")],
        &calibri_italic_malgun_upright_context(),
    );

    assert!(
        source.contains("skew(ax: -18.778deg, origin: bottom + left)"),
        "Hangul on a face with no italic must be slanted by hand, got:\n{source}"
    );
    assert!(
        !source.contains("style: \"italic\""),
        "the upright face must not also be asked for an italic it does not have, got:\n{source}"
    );
}

#[test]
fn test_italic_on_face_with_italic_is_left_to_the_engine() {
    // Triangulation: the same run on a family that *does* ship an italic must
    // take the unchanged path, or every italic in the corpus gains a box.
    let source = generated_source(
        &[italic_run("Report", "Calibri")],
        &calibri_italic_malgun_upright_context(),
    );

    assert!(
        source.contains("style: \"italic\""),
        "a real italic face must still be selected by style, got:\n{source}"
    );
    assert!(
        !source.contains("skew("),
        "a face with a real italic must not be slanted by hand, got:\n{source}"
    );
}

#[test]
fn test_mixed_script_italic_slants_only_the_face_that_needs_it() {
    // The audited deck's captions run Hangul and Latin through one `i="1"`
    // run: GT keeps Calibri-Italic for the Latin part and synthesises the
    // slant only for the Hangul.
    let source = generated_source(
        &[italic_run("PDF 변환", "Calibri")],
        &calibri_italic_malgun_upright_context(),
    );

    assert!(
        source.contains("skew(ax:"),
        "the Hangul half must be slanted by hand, got:\n{source}"
    );
    assert!(
        source.contains("style: \"italic\""),
        "the Latin half must keep the real italic face, got:\n{source}"
    );
    assert!(
        !source.contains("skew(ax: -18.778deg, origin: bottom + left)[PDF"),
        "the Latin half must not be inside a slant box, got:\n{source}"
    );
}

#[test]
fn test_synthetic_oblique_keeps_a_break_opportunity_per_syllable() {
    // A slant box is atomic for line breaking, so one box around the whole
    // run would turn a wrapping Korean caption into a single overflowing
    // line. One box per syllable keeps the break opportunities Korean text
    // already has.
    let source = generated_source(
        &[italic_run("가나다", "Malgun Gothic")],
        &calibri_italic_malgun_upright_context(),
    );

    assert_eq!(
        source.matches("#box(skew(").count(),
        3,
        "each Hangul syllable needs its own box, got:\n{source}"
    );
}

#[test]
fn test_synthetic_oblique_keeps_latin_words_whole() {
    // Latin wraps at spaces, so a word is already atomic: one box per word
    // keeps its kerning and ligatures instead of splitting every letter.
    let context = FontSearchContext::for_test(Vec::new(), &["Bodoni Ornaments"], &[], &[])
        .with_italic_and_scripts(&[], &[("Bodoni Ornaments", &[TextScript::Latin])]);
    let source = generated_source(&[italic_run("two words", "Bodoni Ornaments")], &context);

    assert_eq!(
        source.matches("#box(skew(").count(),
        2,
        "one box per word, with the space left outside them, got:\n{source}"
    );
}

#[test]
fn test_no_font_context_leaves_italic_alone() {
    // On WASM there is no font context at all. Guessing "no italic face" there
    // would slant every italic run in the document.
    let mut out = String::new();
    with_font_search_context(None, || {
        generate_runs(
            &mut out,
            &[italic_run("한국어", "Malgun Gothic")],
            EojeolWrap::Syllable,
        );
    });

    assert!(
        !out.contains("skew("),
        "an unknown font context must not synthesise anything, got:\n{out}"
    );
}

#[test]
fn test_synthetic_oblique_inside_an_eojeol_frame_keeps_the_frame_s_descent() {
    // An eojeol frame shifts its own baseline up by the descent it expects its
    // content to carry (issue #626). A slant box that ended at the baseline
    // carried none, so the frame's shift over-corrected and every framed
    // Korean italic in `10_research_report_ko` dropped 3.97pt below the
    // unframed text on its own line.
    let context = calibri_italic_malgun_upright_context();
    let source = with_font_search_context(Some(&context), || {
        let mut out = String::new();
        generate_runs(
            &mut out,
            &[italic_run("가나 다라", "Malgun Gothic")],
            EojeolWrap::Eojeol {
                line_box_em: Some((1.28789, 0.44121)),
                measure_pt: Some(400.0),
            },
        );
        out
    });

    // 0.44121em at the run's 14pt.
    assert!(
        source.contains("inset: (bottom: 6.17694pt), baseline: 6.17694pt"),
        "a framed slant box must claim the frame's own descent, got:\n{source}"
    );
}

#[test]
fn test_synthetic_oblique_outside_a_frame_states_no_seat() {
    // Triangulation for the seat: with nothing depending on the box's descent
    // the box ends at the baseline, which is what keeps the unframed paths
    // emitting the same geometry they did before.
    let source = generated_source(
        &[italic_run("가나", "Malgun Gothic")],
        &calibri_italic_malgun_upright_context(),
    );

    assert!(
        !source.contains("inset:"),
        "an unframed slant box must not pad itself, got:\n{source}"
    );
}

/// The RTL exemption outranks the tracking rule of issue #864: typst 0.14.2
/// mis-orders RTL glyph ranges when the `kern` feature is off, and losing
/// glyphs is worse than the spurious word break the rule exists to prevent.
#[test]
fn test_rtl_run_keeps_kerning_despite_tracking() {
    let doc = make_doc_with_default_text(
        vec![make_flow_page(vec![styled_paragraph(
            "مرحبا بالعالم",
            TextStyle {
                font_family: Some("Arial".to_string()),
                letter_spacing: Some(3.0),
                pair_kerning: Some(PairKerning::Never),
                ..TextStyle::default()
            },
        )])],
        TextStyle {
            font_size: Some(11.0),
            pair_kerning: Some(PairKerning::Never),
            ..TextStyle::default()
        },
    );

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: true"),
        "a tracked RTL run must keep kerning on: {source}"
    );
    assert!(
        !source.contains("kerning: false"),
        "the tracking rule must not reach RTL text: {source}"
    );
}

/// Issue #1073: slide 13 of the #841 deck sets `RAPPORTSTATUS` at 38pt with
/// `spc="300"` under a master `titleStyle` declaring `kern="1200"`. PowerPoint
/// applies both, tightening `TA` and `AT` by ~1.9pt each; switching the `kern`
/// feature off for the whole run because it is tracked set the line 3.94pt
/// wide and, the paragraph being centred, moved its origin 0.52pt left.
#[test]
fn test_tracked_run_at_its_stated_kern_threshold_keeps_kerning() {
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "RAPPORTSTATUS",
        TextStyle {
            font_family: Some("Posterama".to_string()),
            font_size: Some(38.0),
            bold: Some(true),
            letter_spacing: Some(3.0),
            pair_kerning: Some(PairKerning::AtOrAbovePt(12.0)),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: true"),
        "a tracked run at its stated threshold keeps kerning: {source}"
    );
    assert!(
        !source.contains("kerning: false"),
        "the tracking rule must not override a stated threshold: {source}"
    );
}

#[test]
fn test_tracked_run_below_its_stated_kern_threshold_stays_unkerned() {
    // Triangulation: the threshold is what decides, not the mere presence of a
    // rule. The same deck's 10pt tracked footer sits under `kern="1200"` and
    // PowerPoint does not kern it.
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "CONTOSO ALLE ANSATTE",
        TextStyle {
            font_family: Some("Posterama".to_string()),
            font_size: Some(10.0),
            letter_spacing: Some(2.0),
            pair_kerning: Some(PairKerning::AtOrAbovePt(12.0)),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;

    assert!(
        source.contains("kerning: false"),
        "a tracked run below its stated threshold stays unkerned: {source}"
    );
}

#[test]
fn test_tracked_run_in_a_substituted_face_stays_unkerned() {
    // The protection of issue #864 is specifically against a *substitute's*
    // kern pairs riding on top of tracking the document sized for another
    // face: the stated threshold says what PowerPoint does with the real font
    // and nothing about the stand-in. With Posterama absent, the tracked title
    // must still reach the engine with the feature off.
    use crate::render::font_context::FontSearchContext;
    let context = FontSearchContext::for_test(Vec::new(), &["Arial"], &[], &[]);
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "RAPPORTSTATUS",
        TextStyle {
            font_family: Some("Posterama".to_string()),
            font_size: Some(38.0),
            bold: Some(true),
            letter_spacing: Some(3.0),
            pair_kerning: Some(PairKerning::AtOrAbovePt(12.0)),
            ..TextStyle::default()
        },
    )])]);

    let source = crate::render::typst_gen::generate_typst_with_options_and_font_context(
        &doc,
        &crate::config::ConvertOptions::default(),
        Some(&context),
    )
    .unwrap()
    .source;

    assert!(
        source.contains("kerning: false"),
        "a tracked run whose face is substituted stays unkerned: {source}"
    );
}

// ── Excel's whole-point advance grid (issue #1088) ───────────────

/// A spreadsheet page whose single cell holds `text` in `style`.
fn sheet_page_with_cell(text: &str, style: TextStyle) -> Page {
    sheet_page_with_aligned_cell(text, style, None)
}

fn sheet_page_with_aligned_cell(
    text: &str,
    style: TextStyle,
    alignment: Option<Alignment>,
) -> Page {
    let mut content: Block = styled_paragraph(text, style);
    let Block::Paragraph(paragraph) = &mut content else {
        unreachable!("styled_paragraph always returns a paragraph")
    };
    paragraph.style.alignment = alignment;

    Page::Sheet(SheetPage {
        name: String::new(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![content],
                    ..TableCell::default()
                }],
                height: None,
                minimum_height: None,
            }],
            column_widths: vec![400.0],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
    })
}

/// The first `tracking:` the generator emitted, in points.
fn emitted_tracking_pt(source: &str) -> Option<f64> {
    let after_tracking: &str = source.split_once("tracking: ")?.1;
    let (value, _) = after_tracking.split_once("pt")?;
    value.parse().ok()
}

/// The last explicit horizontal space the generator emitted, in points.
fn emitted_last_horizontal_space_pt(source: &str) -> Option<f64> {
    let (_, after_space) = source.rsplit_once("#h(")?;
    let (value, _) = after_space.split_once("pt)")?;
    value.parse().ok()
}

/// The embedded Libertinus Serif faces make this deterministic on every
/// target, the same pin the digit-advance (#621) and token-advance (#624)
/// measurements take.
///
/// Ground truth is the face's own `hmtx`, read independently of this crate:
/// "Total" advances 0.597, 0.504, 0.316, 0.457 and 0.264em. The four that
/// carry a gap occupy 18.74pt at 10pt against Excel's 6 + 5 + 3 + 5 = 19pt, so
/// each gap widens by 0.065pt.
#[test]
fn test_sheet_cell_line_takes_excels_whole_point_advance_grid() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    let tracking: f64 =
        emitted_tracking_pt(&source).unwrap_or_else(|| panic!("no tracking emitted: {source}"));
    assert!(
        (tracking - 0.065).abs() < 1e-9,
        "'Total' at 10pt should be tracked 0.065pt onto the whole-point grid, got {tracking}"
    );
}

/// Excel includes the final glyph's rounded advance when it places a
/// right-aligned line from the cell's trailing edge. Typst drops tracking after
/// that glyph, so the difference needs its own trailing spacer (issue #1233).
///
/// Libertinus Serif's final `l` in "Total" advances 0.264em: 2.64pt at 10pt,
/// which Excel rounds to 3pt. The line therefore reserves the missing 0.36pt.
#[test]
fn test_right_aligned_sheet_cell_reserves_the_rounded_trailing_advance() {
    let doc = make_doc(vec![sheet_page_with_aligned_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
        Some(Alignment::Right),
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("#h(0.36pt)"),
        "a right-aligned sheet line must reserve its rounded final advance: {source}"
    );
}

/// Excel rounds glyph advances at the cell's declared size and scales the
/// resulting whole-point grid afterwards (issue #1238). At half scale,
/// Libertinus Serif 10pt "Total" is emitted at 5pt: its four gap advances
/// occupy 9.37 printed points and Excel's declared 19pt grid prints at 9.5pt,
/// so each gap takes 0.0325pt. The final 2.64pt declared advance rounds to 3pt
/// and contributes a separate 0.18 printed-point reserve.
#[test]
fn test_scaled_sheet_cell_uses_the_declared_size_advance_grid() {
    let mut page = sheet_page_with_aligned_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(5.0),
            ..TextStyle::default()
        },
        Some(Alignment::Right),
    );
    let Page::Sheet(sheet) = &mut page else {
        unreachable!("sheet_page_with_aligned_cell returns a sheet")
    };
    sheet.table.print_scale = Some(0.5);

    let source = generate_typst(&make_doc(vec![page])).unwrap().source;
    let tracking: f64 =
        emitted_tracking_pt(&source).unwrap_or_else(|| panic!("no tracking emitted: {source}"));
    assert!(
        (tracking - 0.0325).abs() < 1e-9,
        "the declared-size grid needs 0.0325pt tracking after scaling, got {tracking}"
    );
    assert!(
        source.contains("#h(0.18pt)"),
        "the rounded final advance must reserve 0.18 printed points: {source}"
    );
}

/// The trailing correction follows the advance's rounding direction rather
/// than always widening the line. A single glyph also needs it even though it
/// has no inter-glyph gap that could take `tracking`.
#[test]
fn test_right_aligned_sheet_cell_trailing_advance_can_narrow_the_line() {
    let doc = make_doc(vec![sheet_page_with_aligned_cell(
        "O",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
        Some(Alignment::Right),
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert_eq!(
        emitted_tracking_pt(&source),
        None,
        "a one-glyph line still has no gap to track: {source}"
    );
    assert!(
        source.contains("#h(-0.02pt)"),
        "Libertinus Serif O advances 7.02pt and rounds down to 7pt: {source}"
    );
}

/// Width cannot move a left-aligned origin, and Excel separately snaps a
/// centred origin to a whole point. Keep the measured issue #1233 scope on
/// right alignment instead of moving either unaffected class.
#[test]
fn test_non_right_aligned_sheet_cells_do_not_reserve_the_trailing_advance() {
    for alignment in [None, Some(Alignment::Left), Some(Alignment::Center)] {
        let doc = make_doc(vec![sheet_page_with_aligned_cell(
            "Total",
            TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            alignment,
        )]);

        let source = generate_typst(&doc).unwrap().source;
        assert!(
            !source.contains("#h(0.36pt)"),
            "only a right-aligned sheet line takes the trailing reserve: {source}"
        );
    }
}

/// A paragraph-wide trailing reserve cannot represent independently aligned
/// explicit lines or tab segments. Decline all such paragraphs rather than
/// correcting only their last segment (issue #1233).
#[test]
fn test_segmented_right_aligned_sheet_cells_do_not_reserve_a_trailing_advance() {
    for text in [
        "Total\nNext".to_string(),
        "Total\tNext".to_string(),
        "Total\u{000B}Next".to_string(),
    ] {
        let doc = make_doc(vec![sheet_page_with_aligned_cell(
            &text,
            TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            Some(Alignment::Right),
        )]);

        let source = generate_typst(&doc).unwrap().source;
        assert_eq!(
            emitted_last_horizontal_space_pt(&source),
            None,
            "a segmented paragraph must not receive a partial trailing reserve: {source}"
        );
    }
}

/// Small caps are shaped as separate transformed glyph runs downstream, so a
/// reserve measured from the source run would not describe their final glyph.
#[test]
fn test_small_caps_right_aligned_sheet_cell_declines_the_trailing_reserve() {
    let doc = make_doc(vec![sheet_page_with_aligned_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            small_caps: Some(true),
            ..TextStyle::default()
        },
        Some(Alignment::Right),
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert_eq!(
        emitted_last_horizontal_space_pt(&source),
        None,
        "small caps must decline a source-run trailing reserve: {source}"
    );
}

/// All-caps text is shaped from its uppercase glyphs, whose final advance can
/// differ from the source case. Its reserve must therefore match literal
/// uppercase text, not the original mixed-case string.
#[test]
fn test_all_caps_right_aligned_sheet_cell_measures_the_uppercase_terminal_glyph() {
    let source_for = |text: &str, all_caps: bool| {
        let doc = make_doc(vec![sheet_page_with_aligned_cell(
            text,
            TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                all_caps: Some(all_caps),
                ..TextStyle::default()
            },
            Some(Alignment::Right),
        )]);
        generate_typst(&doc).unwrap().source
    };

    let transformed = source_for("Total", true);
    let literal_uppercase = source_for("TOTAL", false);
    let mixed_case = source_for("Total", false);
    let transformed_space = emitted_last_horizontal_space_pt(&transformed);
    let uppercase_space = emitted_last_horizontal_space_pt(&literal_uppercase);
    let mixed_case_space = emitted_last_horizontal_space_pt(&mixed_case);

    assert_eq!(
        transformed_space, uppercase_space,
        "all caps must measure the same terminal glyph as literal uppercase text"
    );
    assert_ne!(
        transformed_space, mixed_case_space,
        "this fixture must distinguish uppercase L from mixed-case l"
    );
}

/// The grid is not a widening: an advance just over a half point rounds *down*
/// and the line comes out narrower than the face sets it.
///
/// "Achievement" at 9pt: its ten gap-carrying advances sum to 45.918pt of
/// `hmtx` and quantize to 45pt, so each gap loses 0.0918pt. A model that only
/// ever added width — the direction issue #1088 measures on 9pt Arial — would
/// get the sign wrong here.
#[test]
fn test_sheet_cell_grid_narrows_a_line_whose_advances_round_down() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "Achievement",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(9.0),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    let tracking: f64 =
        emitted_tracking_pt(&source).unwrap_or_else(|| panic!("no tracking emitted: {source}"));
    assert!(
        (tracking - -0.0918).abs() < 1e-9,
        "'Achievement' at 9pt should be tracked -0.0918pt, got {tracking}"
    );
}

/// A bold cell is quantized on the bold face's advances, which are not the
/// regular face's: bold "Total"'s gap-carrying advances sum to 20.67pt at 10pt
/// and quantize to 22pt, giving 0.3325pt against the regular face's 0.065pt.
#[test]
fn test_sheet_cell_grid_measures_the_runs_own_weight() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            bold: Some(true),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    let tracking: f64 =
        emitted_tracking_pt(&source).unwrap_or_else(|| panic!("no tracking emitted: {source}"));
    assert!(
        (tracking - 0.3325).abs() < 1e-9,
        "bold 'Total' at 10pt should be tracked 0.3325pt, got {tracking}"
    );
}

/// Excel accumulates the bare `hmtx` advances, so a pair kern landing on top of
/// the correction would move every glyph after it back off the grid. The run
/// states the decision on itself rather than leaning on an enclosing rule.
#[test]
fn test_sheet_cell_grid_run_is_unkerned_and_unligated() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("kerning: false"),
        "a grid-corrected cell run must reach the engine unkerned: {source}"
    );
    assert!(
        source.contains("ligatures: false"),
        "a ligature would swallow the gaps the correction rides in: {source}"
    );
}

/// The grid is Excel's alone. A Word page sets the same run on the face's own
/// advances, and no native `docx` export puts its glyph origins on whole
/// points.
#[test]
fn test_flow_page_keeps_the_faces_own_advances() {
    let doc = make_doc(vec![make_flow_page(vec![styled_paragraph(
        "Total",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
    )])]);

    let source = generate_typst(&doc).unwrap().source;
    assert_eq!(
        emitted_tracking_pt(&source),
        None,
        "a Word paragraph must not be tracked onto Excel's grid: {source}"
    );
}

/// Typst drops the tracking after a shaped item's last glyph, so a one-glyph
/// cell has no gap to carry a correction; emitting one would state a spacing
/// the engine then ignores while costing the run its kerning and ligatures.
#[test]
fn test_single_glyph_sheet_cell_is_left_alone() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "7",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    assert_eq!(
        emitted_tracking_pt(&source),
        None,
        "a one-glyph cell has no gap to track: {source}"
    );
}

/// The correction is anchored on the run's last origin, not on its total width.
///
/// A two-glyph cell is where the two anchors separate most: "OK" advances 7.02
/// and 6.37pt at Libertinus Serif 10pt, and Excel prints the `K` exactly 7pt
/// after the `O`. Spreading both roundings over the one gap there is — the
/// anchor that would make the run's *total* advance exact — pulls that gap to
/// 6.63pt instead, which is how the corresponding Arial cell in the inventory
/// mock came out 4% wide.
#[test]
fn test_sheet_cell_grid_puts_the_last_origin_on_the_grid() {
    let doc = make_doc(vec![sheet_page_with_cell(
        "OK",
        TextStyle {
            font_family: Some("Libertinus Serif".to_string()),
            font_size: Some(10.0),
            ..TextStyle::default()
        },
    )]);

    let source = generate_typst(&doc).unwrap().source;
    let tracking: f64 =
        emitted_tracking_pt(&source).unwrap_or_else(|| panic!("no tracking emitted: {source}"));
    assert!(
        (tracking - -0.02).abs() < 1e-9,
        "'OK' should be tracked -0.02pt so its second glyph lands 7pt along, got {tracking}"
    );
}
