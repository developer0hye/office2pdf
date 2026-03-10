use super::*;

#[test]
fn test_parse_sqref_single_range() {
    let ranges = parse_sqref("A1:C3");
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].start_col, 1);
    assert_eq!(ranges[0].start_row, 1);
    assert_eq!(ranges[0].end_col, 3);
    assert_eq!(ranges[0].end_row, 3);
}

#[test]
fn test_parse_sqref_single_cell() {
    let ranges = parse_sqref("B5");
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].start_col, 2);
    assert_eq!(ranges[0].start_row, 5);
    assert_eq!(ranges[0].end_col, 2);
    assert_eq!(ranges[0].end_row, 5);
}

#[test]
fn test_parse_sqref_multiple_ranges() {
    let ranges = parse_sqref("A1:B2 D4:E5");
    assert_eq!(ranges.len(), 2);
}

#[test]
fn test_interpolate_color_extremes() {
    let white = Color::new(255, 255, 255);
    let red = Color::new(255, 0, 0);

    let at_min = interpolate_color(white, red, 0.0);
    assert_eq!(at_min, white);

    let at_max = interpolate_color(white, red, 1.0);
    assert_eq!(at_max, red);
}

#[test]
fn test_interpolate_color_midpoint() {
    let white = Color::new(255, 255, 255);
    let red = Color::new(255, 0, 0);

    let mid = interpolate_color(white, red, 0.5);
    assert_eq!(mid.r, 255);
    assert_eq!(mid.g, 128);
    assert_eq!(mid.b, 128);
}
