use super::*;

#[test]
fn test_default_page_size_is_a4() {
    let size = PageSize::default();
    assert!((size.width - 595.28).abs() < 0.01);
    assert!((size.height - 841.89).abs() < 0.01);
}

#[test]
fn test_default_margins_are_one_inch() {
    let margins = Margins::default();
    assert!((margins.top - 72.0).abs() < 0.01);
    assert!((margins.left - 72.0).abs() < 0.01);
}

#[test]
fn test_fixed_page_background_color() {
    use crate::ir::Color;
    let page = FixedPage {
        size: PageSize::default(),
        elements: vec![],
        background_color: Some(Color::new(255, 0, 0)),
        background_gradient: None,
    };
    assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
}

#[test]
fn test_fixed_page_no_background_color() {
    let page = FixedPage {
        size: PageSize::default(),
        elements: vec![],
        background_color: None,
        background_gradient: None,
    };
    assert!(page.background_color.is_none());
}
