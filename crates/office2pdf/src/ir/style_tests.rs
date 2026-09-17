use super::*;

#[test]
fn test_color_constructors() {
    let black = Color::black();
    assert_eq!(black, Color { r: 0, g: 0, b: 0 });
    let white = Color::white();
    assert_eq!(
        white,
        Color {
            r: 255,
            g: 255,
            b: 255,
        }
    );
}

#[test]
fn test_stylesheet_default_is_empty() {
    let ss = StyleSheet::default();
    assert!(ss.styles.is_empty());
}

#[test]
fn test_text_style_default_has_none_text_effects() {
    let ts = TextStyle::default();
    assert!(ts.vertical_align.is_none());
    assert!(ts.baseline_shift.is_none());
    assert!(ts.all_caps.is_none());
    assert!(ts.small_caps.is_none());
}

#[test]
fn test_text_style_superscript() {
    let ts = TextStyle {
        vertical_align: Some(VerticalTextAlign::Superscript),
        ..TextStyle::default()
    };
    assert_eq!(ts.vertical_align, Some(VerticalTextAlign::Superscript));
}

#[test]
fn test_text_style_subscript() {
    let ts = TextStyle {
        vertical_align: Some(VerticalTextAlign::Subscript),
        ..TextStyle::default()
    };
    assert_eq!(ts.vertical_align, Some(VerticalTextAlign::Subscript));
}

#[test]
fn test_text_style_caps() {
    let ts = TextStyle {
        all_caps: Some(true),
        small_caps: Some(true),
        ..TextStyle::default()
    };
    assert_eq!(ts.all_caps, Some(true));
    assert_eq!(ts.small_caps, Some(true));
}

// ── TextStyle::merge_from tests ──────────────────────────────────────

#[test]
fn text_style_merge_from_all_none_source_preserves_target() {
    let mut target = TextStyle {
        font_family: Some("Arial".to_string()),
        east_asian_font_family: None,
        font_size: Some(12.0),
        bold: Some(true),
        italic: Some(false),
        underline: Some(true),
        strikethrough: Some(false),
        color: Some(Color::new(255, 0, 0)),
        color_alpha: None,
        highlight: Some(Color::new(0, 255, 0)),
        vertical_align: Some(VerticalTextAlign::Superscript),
        baseline_shift: Some(BaselineShiftEm(0.2)),
        all_caps: Some(true),
        small_caps: Some(false),
        letter_spacing: Some(1.5),
        pair_kerning: Some(PairKerning::Never),
    };
    let original: TextStyle = target.clone();
    let source = TextStyle::default();

    target.merge_from(&source);

    assert_eq!(target, original);
}

#[test]
fn text_style_merge_from_all_some_source_overwrites_target() {
    let mut target = TextStyle {
        font_family: Some("Arial".to_string()),
        east_asian_font_family: None,
        font_size: Some(12.0),
        bold: Some(true),
        italic: Some(true),
        underline: Some(true),
        strikethrough: Some(true),
        color: Some(Color::new(255, 0, 0)),
        color_alpha: None,
        highlight: Some(Color::new(0, 255, 0)),
        vertical_align: Some(VerticalTextAlign::Superscript),
        baseline_shift: Some(BaselineShiftEm(0.2)),
        all_caps: Some(true),
        small_caps: Some(true),
        letter_spacing: Some(1.5),
        pair_kerning: Some(PairKerning::Never),
    };
    let source = TextStyle {
        font_family: Some("Times".to_string()),
        east_asian_font_family: None,
        font_size: Some(24.0),
        bold: Some(false),
        italic: Some(false),
        underline: Some(false),
        strikethrough: Some(false),
        color: Some(Color::new(0, 0, 255)),
        color_alpha: None,
        highlight: Some(Color::new(128, 128, 128)),
        vertical_align: Some(VerticalTextAlign::Subscript),
        baseline_shift: Some(BaselineShiftEm(-0.25)),
        all_caps: Some(false),
        small_caps: Some(false),
        letter_spacing: Some(3.0),
        pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
    };

    target.merge_from(&source);

    assert_eq!(target, source);
}

#[test]
fn text_style_merge_from_partial_overlap() {
    let mut target = TextStyle {
        font_family: Some("Arial".to_string()),
        east_asian_font_family: None,
        font_size: Some(12.0),
        bold: None,
        italic: Some(true),
        ..TextStyle::default()
    };
    let source = TextStyle {
        font_family: Some("Helvetica".to_string()),
        east_asian_font_family: None,
        bold: Some(true),
        italic: None,
        color: Some(Color::new(100, 100, 100)),
        color_alpha: None,
        ..TextStyle::default()
    };

    target.merge_from(&source);

    assert_eq!(target.font_family, Some("Helvetica".to_string()));
    assert_eq!(target.font_size, Some(12.0));
    assert_eq!(target.bold, Some(true));
    assert_eq!(target.italic, Some(true));
    assert_eq!(target.color, Some(Color::new(100, 100, 100)));
    assert!(target.underline.is_none());
}

#[test]
fn text_style_merge_from_into_default_target() {
    let mut target = TextStyle::default();
    let source = TextStyle {
        font_size: Some(18.0),
        bold: Some(false),
        letter_spacing: Some(2.0),
        ..TextStyle::default()
    };

    target.merge_from(&source);

    assert_eq!(target.font_size, Some(18.0));
    assert_eq!(target.bold, Some(false));
    assert_eq!(target.letter_spacing, Some(2.0));
    assert!(target.font_family.is_none());
}

// ── ParagraphStyle::merge_from tests ─────────────────────────────────

#[test]
fn paragraph_style_merge_from_all_none_source_preserves_target() {
    let mut target = ParagraphStyle {
        alignment: Some(Alignment::Center),
        word_wrap: None,
        indent_left: Some(10.0),
        indent_right: Some(5.0),
        indent_first_line: Some(20.0),
        line_spacing: Some(LineSpacing::Proportional(1.5)),
        line_box: Some(LineBox {
            ascent_em: 1.0,
            descent_em: 0.25,
        }),
        space_before: Some(6.0),
        space_after: Some(12.0),
        space_before_percent: None,
        space_after_percent: None,
        heading_level: Some(2),
        direction: Some(TextDirection::Rtl),
        tab_stops: Some(vec![TabStop {
            position: 72.0,
            alignment: TabAlignment::Left,
            leader: TabLeader::None,
        }]),
        default_tab_stop_pt: Some(40.0),
        background: Some(Color::new(0xEE, 0xEE, 0xEE)),
        border: None,
        border_space: None,
        paragraph_mark_font_family: None,
        sheet_number_format_reserved_glyphs: None,
    };
    let original: ParagraphStyle = target.clone();
    let source = ParagraphStyle::default();

    target.merge_from(&source);

    assert_eq!(target.alignment, original.alignment);
    assert_eq!(target.indent_left, original.indent_left);
    assert_eq!(target.indent_right, original.indent_right);
    assert_eq!(target.indent_first_line, original.indent_first_line);
    assert_eq!(target.line_box, original.line_box);
    assert_eq!(target.space_before, original.space_before);
    assert_eq!(target.space_after, original.space_after);
    assert_eq!(target.heading_level, original.heading_level);
    assert_eq!(target.direction, original.direction);
    assert_eq!(target.tab_stops, original.tab_stops);
    assert_eq!(target.default_tab_stop_pt, original.default_tab_stop_pt);
}

#[test]
fn paragraph_style_merge_from_all_some_source_overwrites_target() {
    let mut target = ParagraphStyle {
        alignment: Some(Alignment::Left),
        indent_left: Some(10.0),
        space_before: Some(6.0),
        ..ParagraphStyle::default()
    };
    let source = ParagraphStyle {
        alignment: Some(Alignment::Right),
        word_wrap: Some(false),
        indent_left: Some(20.0),
        indent_right: Some(15.0),
        indent_first_line: Some(30.0),
        line_spacing: Some(LineSpacing::Exact(14.0)),
        line_box: Some(LineBox {
            ascent_em: 1.3125,
            descent_em: 0.4375,
        }),
        space_before: Some(8.0),
        space_after: Some(16.0),
        space_before_percent: None,
        space_after_percent: None,
        heading_level: Some(1),
        direction: Some(TextDirection::Rtl),
        background: Some(Color::new(0xF4, 0xF4, 0xF4)),
        border: None,
        border_space: None,
        paragraph_mark_font_family: None,
        sheet_number_format_reserved_glyphs: None,
        tab_stops: Some(vec![TabStop {
            position: 144.0,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }]),
        default_tab_stop_pt: Some(72.0),
    };

    target.merge_from(&source);

    assert_eq!(target.alignment, Some(Alignment::Right));
    assert_eq!(target.indent_left, Some(20.0));
    assert_eq!(target.indent_right, Some(15.0));
    assert_eq!(target.indent_first_line, Some(30.0));
    assert_eq!(target.line_box, source.line_box);
    assert_eq!(target.space_before, Some(8.0));
    assert_eq!(target.space_after, Some(16.0));
    assert_eq!(target.heading_level, Some(1));
    assert_eq!(target.direction, Some(TextDirection::Rtl));
    assert_eq!(target.default_tab_stop_pt, Some(72.0));
    assert_eq!(
        target.tab_stops,
        Some(vec![TabStop {
            position: 144.0,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }])
    );
}

#[test]
fn paragraph_style_merge_from_partial_overlap() {
    let mut target = ParagraphStyle {
        alignment: Some(Alignment::Left),
        indent_left: Some(10.0),
        space_before: Some(6.0),
        ..ParagraphStyle::default()
    };
    let source = ParagraphStyle {
        alignment: Some(Alignment::Center),
        space_after: Some(12.0),
        ..ParagraphStyle::default()
    };

    target.merge_from(&source);

    assert_eq!(target.alignment, Some(Alignment::Center));
    assert_eq!(target.indent_left, Some(10.0));
    assert_eq!(target.space_before, Some(6.0));
    assert_eq!(target.space_after, Some(12.0));
    assert!(target.heading_level.is_none());
}

#[test]
fn paragraph_style_merge_from_keeps_spacing_units_mutually_exclusive() {
    let mut target = ParagraphStyle {
        space_before_percent: Some(0.2),
        space_after_percent: Some(0.25),
        ..ParagraphStyle::default()
    };

    target.merge_from(&ParagraphStyle {
        space_before: Some(6.0),
        space_after: Some(12.0),
        ..ParagraphStyle::default()
    });

    assert_eq!(target.space_before, Some(6.0));
    assert_eq!(target.space_after, Some(12.0));
    assert!(target.space_before_percent.is_none());
    assert!(target.space_after_percent.is_none());

    target.merge_from(&ParagraphStyle {
        space_before_percent: Some(0.3),
        space_after_percent: Some(0.4),
        ..ParagraphStyle::default()
    });

    assert!(target.space_before.is_none());
    assert!(target.space_after.is_none());
    assert_eq!(target.space_before_percent, Some(0.3));
    assert_eq!(target.space_after_percent, Some(0.4));
}

// ----- Pair kerning threshold model (issue #628) -----

#[test]
fn pair_kerning_never_is_off_at_every_size() {
    // Word writes no `w:kern` in any of the business mocks, and writes
    // `w:val="0"` when a user turns kerning off explicitly. Both mean the
    // same thing: every glyph advances by its nominal width.
    for size_pt in [8.0_f64, 10.0, 11.0, 20.0, 22.0, 72.0] {
        assert!(
            !PairKerning::Never.applies_at(Some(size_pt)),
            "Never must stay off at {size_pt}pt"
        );
    }
    assert!(!PairKerning::Never.applies_at(None));
}

#[test]
fn pair_kerning_threshold_applies_from_its_own_size_up() {
    // `w:kern w:val="28"` is 28 half-points = 14pt: Word kerns 14pt and
    // larger, and leaves 13.5pt and smaller alone.
    let threshold: PairKerning = PairKerning::AtOrAbovePt(14.0);

    assert!(!threshold.applies_at(Some(11.0)));
    assert!(!threshold.applies_at(Some(13.5)));
    assert!(
        threshold.applies_at(Some(14.0)),
        "inclusive at the threshold"
    );
    assert!(threshold.applies_at(Some(20.0)));
    assert!(threshold.applies_at(Some(22.0)));
}

#[test]
fn pair_kerning_threshold_reads_an_unstated_size_as_body_text() {
    // A run that states no size takes the document default, which Word ships
    // at 10pt — below every threshold a real document states.
    assert!(!PairKerning::AtOrAbovePt(14.0).applies_at(None));
    // A threshold at or below that default still kerns such a run.
    assert!(PairKerning::AtOrAbovePt(8.0).applies_at(None));
}

#[test]
fn text_style_merge_from_lets_only_a_stated_kerning_decision_win() {
    // Only a level that actually states `w:kern` may change the answer. The
    // stated cases are a threshold and Word's explicit `w:val="0"`; silence
    // is inheritance, which is what an absent element means at every level
    // below `w:docDefaults` (issue #628 review, defect 3).
    let mut target = TextStyle {
        pair_kerning: Some(PairKerning::AtOrAbovePt(16.0)),
        ..TextStyle::default()
    };

    target.merge_from(&TextStyle::default());
    assert_eq!(
        target.pair_kerning,
        Some(PairKerning::AtOrAbovePt(16.0)),
        "an unstated run property leaves the inherited threshold in place"
    );

    target.merge_from(&TextStyle {
        pair_kerning: Some(PairKerning::Never),
        ..TextStyle::default()
    });
    assert_eq!(
        target.pair_kerning,
        Some(PairKerning::Never),
        "an explicit w:val=\"0\" does turn the inherited threshold off"
    );
}

#[test]
fn text_style_silent_run_keeps_its_styles_kerning_threshold() {
    // Word's Title style states a 14pt threshold and a 28pt size; a run that
    // restates only the size must still be kerned.
    let mut resolved = TextStyle {
        font_size: Some(28.0),
        pair_kerning: Some(PairKerning::AtOrAbovePt(14.0)),
        ..TextStyle::default()
    };
    resolved.merge_from(&TextStyle {
        font_size: Some(28.0),
        bold: Some(true),
        ..TextStyle::default()
    });

    assert_eq!(resolved.pair_kerning, Some(PairKerning::AtOrAbovePt(14.0)));
    assert!(
        resolved
            .pair_kerning
            .expect("stated")
            .applies_at(resolved.font_size),
        "a 28pt run above a 14pt threshold keeps kerning"
    );
}

#[test]
fn text_style_default_states_no_pair_kerning() {
    assert!(TextStyle::default().pair_kerning.is_none());
}
