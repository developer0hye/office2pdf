/// Supported input document formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub enum Format {
    Docx,
    Pptx,
    Xlsx,
}

impl Format {
    /// Detect format from file extension.
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "docx" => Some(Self::Docx),
            "pptx" => Some(Self::Pptx),
            "xlsx" => Some(Self::Xlsx),
            _ => None,
        }
    }
}

/// A range of slide numbers (1-indexed) for PPTX conversion.
#[derive(Debug, Clone, PartialEq, Eq)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub struct SlideRange {
    /// Start slide number (1-indexed, inclusive).
    pub start: u32,
    /// End slide number (1-indexed, inclusive).
    pub end: u32,
}

impl SlideRange {
    /// Create a new slide range (1-indexed, inclusive on both ends).
    pub fn new(start: u32, end: u32) -> Self {
        Self { start, end }
    }

    /// Check if a 1-indexed slide number is within this range.
    pub fn contains(&self, slide_number: u32) -> bool {
        slide_number >= self.start && slide_number <= self.end
    }

    /// Parse a slide range string like "1-5" or "3".
    pub fn parse(s: &str) -> Result<Self, String> {
        if let Some((start_str, end_str)) = s.split_once('-') {
            let start: u32 = start_str
                .trim()
                .parse()
                .map_err(|_| format!("invalid start number: {start_str}"))?;
            let end: u32 = end_str
                .trim()
                .parse()
                .map_err(|_| format!("invalid end number: {end_str}"))?;
            if start == 0 || end == 0 {
                return Err("slide numbers must be >= 1".to_string());
            }
            if start > end {
                return Err(format!("start ({start}) must be <= end ({end})"));
            }
            Ok(Self::new(start, end))
        } else {
            let n: u32 = s
                .trim()
                .parse()
                .map_err(|_| format!("invalid slide number: {s}"))?;
            if n == 0 {
                return Err("slide number must be >= 1".to_string());
            }
            Ok(Self::new(n, n))
        }
    }
}

/// PDF standard to enforce compliance with.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub enum PdfStandard {
    /// PDF/A-2b for archival purposes.
    PdfA2b,
}

/// Paper size for output PDF.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub enum PaperSize {
    /// A4: 595.28pt × 841.89pt (210mm × 297mm).
    A4,
    /// US Letter: 612pt × 792pt (8.5in × 11in).
    Letter,
    /// US Legal: 612pt × 1008pt (8.5in × 14in).
    Legal,
    /// Custom dimensions in points.
    Custom { width: f64, height: f64 },
}

impl PaperSize {
    /// Returns (width, height) in points.
    pub fn dimensions(&self) -> (f64, f64) {
        match self {
            Self::A4 => (595.28, 841.89),
            Self::Letter => (612.0, 792.0),
            Self::Legal => (612.0, 1008.0),
            Self::Custom { width, height } => (*width, *height),
        }
    }

    /// Parse a paper size string (case-insensitive): "a4", "letter", "legal".
    pub fn parse(s: &str) -> Result<Self, String> {
        match s.to_ascii_lowercase().as_str() {
            "a4" => Ok(Self::A4),
            "letter" => Ok(Self::Letter),
            "legal" => Ok(Self::Legal),
            _ => Err(format!(
                "unknown paper size: {s}; expected one of: a4, letter, legal"
            )),
        }
    }
}

/// Options controlling the conversion process.
#[derive(Debug, Clone, Default)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub struct ConvertOptions {
    /// Filter XLSX sheets by name. Only sheets whose names are in this list
    /// will be included. If `None`, all sheets are included.
    pub sheet_names: Option<Vec<String>>,
    /// Filter PPTX slides by range (1-indexed). If `None`, all slides are included.
    pub slide_range: Option<SlideRange>,
    /// PDF standard to enforce. If `None`, produces a standard PDF 1.7.
    pub pdf_standard: Option<PdfStandard>,
    /// Override paper size for the output PDF. If `None`, uses the source document's size.
    pub paper_size: Option<PaperSize>,
    /// Additional font directories to search for fonts.
    #[cfg_attr(feature = "typescript", ts(type = "Array<string>"))]
    pub font_paths: Vec<std::path::PathBuf>,
    /// Force landscape orientation. If `Some(true)`, swaps width/height so width > height.
    /// If `Some(false)`, forces portrait. If `None`, uses source document orientation.
    pub landscape: Option<bool>,
    /// Enable tagged PDF output with document structure tags (H1-H6, P, Table, Figure).
    /// When `true`, the output PDF includes accessibility tags that map document
    /// structure for screen readers and assistive technologies.
    pub tagged: bool,
    /// Enable PDF/UA (Universal Accessibility) compliance. Implies `tagged: true`.
    /// Combines tagged PDF with the PDF/UA-1 standard for full accessibility compliance.
    pub pdf_ua: bool,
    /// Enable streaming mode for large file processing.
    /// In streaming mode, XLSX files are processed in chunks of rows to bound memory usage.
    /// Each chunk is compiled independently and the resulting PDFs are merged.
    /// Requires the `pdf-ops` feature for PDF merging.
    pub streaming: bool,
    /// Chunk size (in rows) for streaming mode. Defaults to 1000 if `None`.
    /// Only used when `streaming` is `true`.
    pub streaming_chunk_size: Option<usize>,
}

#[cfg(test)]
#[path = "config_tests.rs"]
mod tests;

#[cfg(all(test, feature = "typescript"))]
#[path = "config_ts_tests.rs"]
mod ts_tests;
