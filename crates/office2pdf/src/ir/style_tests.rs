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
        font_size: Some(12.0),
        bold: Some(true),
        italic: Some(false),
        underline: Some(true),
        strikethrough: Some(false),
        color: Some(Color::new(255, 0, 0)),
        highlight: Some(Color::new(0, 255, 0)),
        vertical_align: Some(VerticalTextAlign::Superscript),
        all_caps: Some(true),
        small_caps: Some(false),
        letter_spacing: Some(1.5),
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
        font_size: Some(12.0),
        bold: Some(true),
        italic: Some(true),
        underline: Some(true),
        strikethrough: Some(true),
        color: Some(Color::new(255, 0, 0)),
        highlight: Some(Color::new(0, 255, 0)),
        vertical_align: Some(VerticalTextAlign::Superscript),
        all_caps: Some(true),
        small_caps: Some(true),
        letter_spacing: Some(1.5),
    };
    let source = TextStyle {
        font_family: Some("Times".to_string()),
        font_size: Some(24.0),
        bold: Some(false),
        italic: Some(false),
        underline: Some(false),
        strikethrough: Some(false),
        color: Some(Color::new(0, 0, 255)),
        highlight: Some(Color::new(128, 128, 128)),
        vertical_align: Some(VerticalTextAlign::Subscript),
        all_caps: Some(false),
        small_caps: Some(false),
        letter_spacing: Some(3.0),
    };

    target.merge_from(&source);

    assert_eq!(target, source);
}

#[test]
fn text_style_merge_from_partial_overlap() {
    let mut target = TextStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(12.0),
        bold: None,
        italic: Some(true),
        ..TextStyle::default()
    };
    let source = TextStyle {
        font_family: Some("Helvetica".to_string()),
        bold: Some(true),
        italic: None,
        color: Some(Color::new(100, 100, 100)),
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
        indent_left: Some(10.0),
        indent_right: Some(5.0),
        indent_first_line: Some(20.0),
        line_spacing: Some(LineSpacing::Proportional(1.5)),
        space_before: Some(6.0),
        space_after: Some(12.0),
        heading_level: Some(2),
        direction: Some(TextDirection::Rtl),
        tab_stops: Some(vec![TabStop {
            position: 72.0,
            alignment: TabAlignment::Left,
            leader: TabLeader::None,
        }]),
    };
    let original: ParagraphStyle = target.clone();
    let source = ParagraphStyle::default();

    target.merge_from(&source);

    assert_eq!(target.alignment, original.alignment);
    assert_eq!(target.indent_left, original.indent_left);
    assert_eq!(target.indent_right, original.indent_right);
    assert_eq!(target.indent_first_line, original.indent_first_line);
    assert_eq!(target.space_before, original.space_before);
    assert_eq!(target.space_after, original.space_after);
    assert_eq!(target.heading_level, original.heading_level);
    assert_eq!(target.direction, original.direction);
    assert_eq!(target.tab_stops, original.tab_stops);
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
        indent_left: Some(20.0),
        indent_right: Some(15.0),
        indent_first_line: Some(30.0),
        line_spacing: Some(LineSpacing::Exact(14.0)),
        space_before: Some(8.0),
        space_after: Some(16.0),
        heading_level: Some(1),
        direction: Some(TextDirection::Rtl),
        tab_stops: Some(vec![TabStop {
            position: 144.0,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }]),
    };

    target.merge_from(&source);

    assert_eq!(target.alignment, Some(Alignment::Right));
    assert_eq!(target.indent_left, Some(20.0));
    assert_eq!(target.indent_right, Some(15.0));
    assert_eq!(target.indent_first_line, Some(30.0));
    assert_eq!(target.space_before, Some(8.0));
    assert_eq!(target.space_after, Some(16.0));
    assert_eq!(target.heading_level, Some(1));
    assert_eq!(target.direction, Some(TextDirection::Rtl));
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
