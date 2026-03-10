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
