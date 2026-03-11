use super::*;

// --- Paper size constants ---

#[test]
fn a4_width_matches_iso_216_spec() {
    // A4 = 210mm × 297mm. 210mm = 595.276pt (at 72pt/inch, 25.4mm/inch).
    // The commonly used rounded value is 595.28pt.
    assert!((A4_WIDTH_PT - 595.28).abs() < 0.01);
}

#[test]
fn a4_height_matches_iso_216_spec() {
    assert!((A4_HEIGHT_PT - 841.89).abs() < 0.01);
}

#[test]
fn letter_width_matches_us_standard() {
    // US Letter = 8.5in × 11in. 8.5 × 72 = 612.0pt.
    assert!((LETTER_WIDTH_PT - 612.0).abs() < 0.01);
}

#[test]
fn letter_height_matches_us_standard() {
    assert!((LETTER_HEIGHT_PT - 792.0).abs() < 0.01);
}

#[test]
fn legal_width_matches_us_standard() {
    // US Legal = 8.5in × 14in. 8.5 × 72 = 612.0pt, 14 × 72 = 1008.0pt.
    assert!((LEGAL_WIDTH_PT - 612.0).abs() < 0.01);
}

#[test]
fn legal_height_matches_us_standard() {
    assert!((LEGAL_HEIGHT_PT - 1008.0).abs() < 0.01);
}

// --- Margin constants ---

#[test]
fn default_margin_is_one_inch() {
    // 1 inch = 72 points.
    assert!((DEFAULT_MARGIN_PT - 72.0).abs() < 0.01);
}

// --- Heading font sizes ---

#[test]
fn heading_font_sizes_has_six_levels() {
    assert_eq!(HEADING_FONT_SIZES.len(), 6);
}

#[test]
fn heading_font_sizes_decrease_monotonically() {
    for i in 0..HEADING_FONT_SIZES.len() - 1 {
        assert!(
            HEADING_FONT_SIZES[i] > HEADING_FONT_SIZES[i + 1],
            "Heading level {} size ({}) should be larger than level {} size ({})",
            i + 1,
            HEADING_FONT_SIZES[i],
            i + 2,
            HEADING_FONT_SIZES[i + 1]
        );
    }
}

#[test]
fn heading_1_is_24pt() {
    assert!((HEADING_FONT_SIZES[0] - 24.0).abs() < 0.01);
}

#[test]
fn heading_6_is_11pt() {
    assert!((HEADING_FONT_SIZES[5] - 11.0).abs() < 0.01);
}

// --- Streaming chunk size ---

#[test]
fn default_streaming_chunk_size_is_1000() {
    assert_eq!(DEFAULT_STREAMING_CHUNK_SIZE, 1000);
}

// --- Unit conversion constants ---

#[test]
fn points_per_inch_is_72() {
    assert!((POINTS_PER_INCH - 72.0).abs() < 0.01);
}

#[test]
fn pixels_per_inch_is_96() {
    assert!((PIXELS_PER_INCH - 96.0).abs() < 0.01);
}

#[test]
fn px_to_pt_conversion_uses_correct_ratio() {
    // 96px = 72pt, so 1px = 0.75pt
    let one_px_in_pt: f64 = POINTS_PER_INCH / PIXELS_PER_INCH;
    assert!((one_px_in_pt - 0.75).abs() < 0.001);
}
