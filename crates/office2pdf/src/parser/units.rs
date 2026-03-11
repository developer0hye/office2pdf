//! Centralized unit conversion functions for OOXML document formats.
//!
//! EMU (English Metric Units), twips, and half-points are the three primary
//! measurement units used across DOCX, PPTX, and XLSX formats. This module
//! provides a single source of truth for converting each to PDF points.

/// Convert English Metric Units (EMU) to points.
/// 1 inch = 914 400 EMU, 1 inch = 72 pt, therefore 1 pt = 12 700 EMU.
///
/// Accepts any integer type that implements `Into<i64>` (i32, u32, i64, etc.).
pub fn emu_to_pt(emu: impl Into<i64>) -> f64 {
    emu.into() as f64 / 12700.0
}

/// Convert twips to points. 1 pt = 20 twips.
///
/// Accepts any numeric type that implements `Into<f64>` (f64, i32, i64, etc.).
pub fn twips_to_pt(twips: impl Into<f64>) -> f64 {
    twips.into() / 20.0
}

/// Convert half-points to points. 1 pt = 2 half-points.
///
/// Used for font sizes in DOCX run properties (`w:sz` attribute).
pub fn half_points_to_pt(hp: impl Into<f64>) -> f64 {
    hp.into() / 2.0
}
