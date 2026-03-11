use super::units::{emu_to_pt, half_points_to_pt, twips_to_pt};

// ── emu_to_pt ───────────────────────────────────────────────────────

#[test]
fn emu_to_pt_zero() {
    assert_eq!(emu_to_pt(0i64), 0.0);
}

#[test]
fn emu_to_pt_one_inch() {
    // 1 inch = 914400 EMU = 72 points
    let result: f64 = emu_to_pt(914_400i64);
    assert!((result - 72.0).abs() < 1e-10);
}

#[test]
fn emu_to_pt_one_point() {
    // 1 pt = 12700 EMU
    let result: f64 = emu_to_pt(12_700i64);
    assert!((result - 1.0).abs() < 1e-10);
}

#[test]
fn emu_to_pt_negative_value() {
    let result: f64 = emu_to_pt(-12_700i64);
    assert!((result - (-1.0)).abs() < 1e-10);
}

#[test]
fn emu_to_pt_accepts_u32() {
    // u32 should work via Into<i64>
    let result: f64 = emu_to_pt(12_700u32);
    assert!((result - 1.0).abs() < 1e-10);
}

#[test]
fn emu_to_pt_accepts_i32() {
    let result: f64 = emu_to_pt(-12_700i32);
    assert!((result - (-1.0)).abs() < 1e-10);
}

#[test]
fn emu_to_pt_fractional_result() {
    // 6350 EMU = 0.5 pt
    let result: f64 = emu_to_pt(6_350i64);
    assert!((result - 0.5).abs() < 1e-10);
}

#[test]
fn emu_to_pt_large_value() {
    // Typical PPTX slide width: 12_192_000 EMU = 960 pt
    let result: f64 = emu_to_pt(12_192_000i64);
    assert!((result - 960.0).abs() < 1e-10);
}

// ── twips_to_pt ─────────────────────────────────────────────────────

#[test]
fn twips_to_pt_zero() {
    assert_eq!(twips_to_pt(0.0), 0.0);
}

#[test]
fn twips_to_pt_twenty_twips_is_one_point() {
    assert!((twips_to_pt(20.0) - 1.0).abs() < 1e-10);
}

#[test]
fn twips_to_pt_negative() {
    assert!((twips_to_pt(-20.0) - (-1.0)).abs() < 1e-10);
}

#[test]
fn twips_to_pt_typical_page_width() {
    // 12240 twips = 612 pt = 8.5 inches
    let result: f64 = twips_to_pt(12240.0);
    assert!((result - 612.0).abs() < 1e-10);
}

#[test]
fn twips_to_pt_accepts_integer_via_into() {
    // i32 -> f64 via Into
    let result: f64 = twips_to_pt(720i32);
    assert!((result - 36.0).abs() < 1e-10);
}

#[test]
fn twips_to_pt_accepts_u16_via_into() {
    let result: f64 = twips_to_pt(240u16);
    assert!((result - 12.0).abs() < 1e-10);
}

// ── half_points_to_pt ───────────────────────────────────────────────

#[test]
fn half_points_to_pt_zero() {
    assert_eq!(half_points_to_pt(0.0), 0.0);
}

#[test]
fn half_points_to_pt_two_half_points_is_one_point() {
    assert!((half_points_to_pt(2.0) - 1.0).abs() < 1e-10);
}

#[test]
fn half_points_to_pt_typical_font_size() {
    // 24 half-points = 12 pt font
    let result: f64 = half_points_to_pt(24.0);
    assert!((result - 12.0).abs() < 1e-10);
}

#[test]
fn half_points_to_pt_odd_value() {
    // 25 half-points = 12.5 pt
    let result: f64 = half_points_to_pt(25.0);
    assert!((result - 12.5).abs() < 1e-10);
}
