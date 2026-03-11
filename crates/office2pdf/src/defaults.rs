//! Centralized document default constants.
//!
//! All hardcoded numeric values for paper sizes, margins, heading styles,
//! streaming parameters, and unit conversions live here so they are defined
//! once and referenced everywhere.

// ---------------------------------------------------------------------------
// Paper sizes in points (1 pt = 1/72 inch)
// ---------------------------------------------------------------------------

/// A4 width: 210 mm = 595.28 pt (ISO 216).
pub const A4_WIDTH_PT: f64 = 595.28;

/// A4 height: 297 mm = 841.89 pt (ISO 216).
pub const A4_HEIGHT_PT: f64 = 841.89;

/// US Letter width: 8.5 in = 612.0 pt.
pub const LETTER_WIDTH_PT: f64 = 612.0;

/// US Letter height: 11 in = 792.0 pt.
pub const LETTER_HEIGHT_PT: f64 = 792.0;

/// US Legal width: 8.5 in = 612.0 pt.
pub const LEGAL_WIDTH_PT: f64 = 612.0;

/// US Legal height: 14 in = 1008.0 pt.
pub const LEGAL_HEIGHT_PT: f64 = 1008.0;

// ---------------------------------------------------------------------------
// Margins
// ---------------------------------------------------------------------------

/// Default page margin: 1 inch = 72 points.
pub const DEFAULT_MARGIN_PT: f64 = 72.0;

// ---------------------------------------------------------------------------
// Heading font sizes
// ---------------------------------------------------------------------------

/// Default font sizes for heading levels 1-6.
/// Index 0 = Heading 1, index 5 = Heading 6.
pub const HEADING_FONT_SIZES: [f64; 6] = [24.0, 20.0, 16.0, 14.0, 12.0, 11.0];

// ---------------------------------------------------------------------------
// Streaming
// ---------------------------------------------------------------------------

/// Default chunk size (in rows) for XLSX streaming mode.
pub const DEFAULT_STREAMING_CHUNK_SIZE: usize = 1000;

// ---------------------------------------------------------------------------
// Unit conversion
// ---------------------------------------------------------------------------

/// Points per inch (PostScript definition).
pub const POINTS_PER_INCH: f64 = 72.0;

/// CSS reference pixels per inch.
pub const PIXELS_PER_INCH: f64 = 96.0;

#[cfg(test)]
#[path = "defaults_tests.rs"]
mod tests;
