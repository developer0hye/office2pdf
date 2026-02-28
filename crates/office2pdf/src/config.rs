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
mod tests {
    use super::*;

    #[test]
    fn test_format_from_extension() {
        assert_eq!(Format::from_extension("docx"), Some(Format::Docx));
        assert_eq!(Format::from_extension("DOCX"), Some(Format::Docx));
        assert_eq!(Format::from_extension("pptx"), Some(Format::Pptx));
        assert_eq!(Format::from_extension("xlsx"), Some(Format::Xlsx));
        assert_eq!(Format::from_extension("pdf"), None);
        assert_eq!(Format::from_extension("txt"), None);
    }

    #[test]
    fn test_slide_range_single() {
        let r = SlideRange::parse("3").unwrap();
        assert_eq!(r.start, 3);
        assert_eq!(r.end, 3);
        assert!(!r.contains(2));
        assert!(r.contains(3));
        assert!(!r.contains(4));
    }

    #[test]
    fn test_slide_range_range() {
        let r = SlideRange::parse("2-5").unwrap();
        assert_eq!(r.start, 2);
        assert_eq!(r.end, 5);
        assert!(!r.contains(1));
        assert!(r.contains(2));
        assert!(r.contains(3));
        assert!(r.contains(5));
        assert!(!r.contains(6));
    }

    #[test]
    fn test_slide_range_parse_errors() {
        assert!(SlideRange::parse("abc").is_err());
        assert!(SlideRange::parse("0").is_err());
        assert!(SlideRange::parse("5-2").is_err());
        assert!(SlideRange::parse("0-3").is_err());
        assert!(SlideRange::parse("a-b").is_err());
    }

    #[test]
    fn test_convert_options_default() {
        let opts = ConvertOptions::default();
        assert!(opts.sheet_names.is_none());
        assert!(opts.slide_range.is_none());
    }

    #[test]
    fn test_convert_options_with_sheets() {
        let opts = ConvertOptions {
            sheet_names: Some(vec!["Sheet1".to_string(), "Data".to_string()]),
            ..Default::default()
        };
        assert_eq!(opts.sheet_names.as_ref().unwrap().len(), 2);
    }

    #[test]
    fn test_convert_options_with_slide_range() {
        let opts = ConvertOptions {
            slide_range: Some(SlideRange::new(1, 3)),
            ..Default::default()
        };
        assert!(opts.slide_range.as_ref().unwrap().contains(2));
    }

    #[test]
    fn test_pdf_standard_enum_exists() {
        let std = PdfStandard::PdfA2b;
        assert_eq!(format!("{std:?}"), "PdfA2b");
    }

    #[test]
    fn test_convert_options_pdf_standard_default_none() {
        let opts = ConvertOptions::default();
        assert!(opts.pdf_standard.is_none());
    }

    #[test]
    fn test_convert_options_with_pdf_standard() {
        let opts = ConvertOptions {
            pdf_standard: Some(PdfStandard::PdfA2b),
            ..Default::default()
        };
        assert_eq!(opts.pdf_standard, Some(PdfStandard::PdfA2b));
    }

    // --- PaperSize tests ---

    #[test]
    fn test_paper_size_a4_dimensions() {
        let (w, h) = PaperSize::A4.dimensions();
        assert!((w - 595.28).abs() < 0.01);
        assert!((h - 841.89).abs() < 0.01);
    }

    #[test]
    fn test_paper_size_letter_dimensions() {
        let (w, h) = PaperSize::Letter.dimensions();
        assert!((w - 612.0).abs() < 0.01);
        assert!((h - 792.0).abs() < 0.01);
    }

    #[test]
    fn test_paper_size_legal_dimensions() {
        let (w, h) = PaperSize::Legal.dimensions();
        assert!((w - 612.0).abs() < 0.01);
        assert!((h - 1008.0).abs() < 0.01);
    }

    #[test]
    fn test_paper_size_custom_dimensions() {
        let ps = PaperSize::Custom {
            width: 400.0,
            height: 600.0,
        };
        assert_eq!(ps.dimensions(), (400.0, 600.0));
    }

    #[test]
    fn test_paper_size_parse() {
        assert_eq!(PaperSize::parse("a4").unwrap(), PaperSize::A4);
        assert_eq!(PaperSize::parse("A4").unwrap(), PaperSize::A4);
        assert_eq!(PaperSize::parse("letter").unwrap(), PaperSize::Letter);
        assert_eq!(PaperSize::parse("LETTER").unwrap(), PaperSize::Letter);
        assert_eq!(PaperSize::parse("legal").unwrap(), PaperSize::Legal);
        assert!(PaperSize::parse("tabloid").is_err());
    }

    #[test]
    fn test_convert_options_paper_size_default_none() {
        let opts = ConvertOptions::default();
        assert!(opts.paper_size.is_none());
    }

    #[test]
    fn test_convert_options_font_paths_default_empty() {
        let opts = ConvertOptions::default();
        assert!(opts.font_paths.is_empty());
    }

    #[test]
    fn test_convert_options_landscape_default_none() {
        let opts = ConvertOptions::default();
        assert!(opts.landscape.is_none());
    }

    #[test]
    fn test_convert_options_with_paper_size() {
        let opts = ConvertOptions {
            paper_size: Some(PaperSize::Letter),
            ..Default::default()
        };
        assert_eq!(opts.paper_size, Some(PaperSize::Letter));
    }

    #[test]
    fn test_convert_options_with_font_paths() {
        let opts = ConvertOptions {
            font_paths: vec![
                std::path::PathBuf::from("/usr/share/fonts"),
                std::path::PathBuf::from("/home/user/.fonts"),
            ],
            ..Default::default()
        };
        assert_eq!(opts.font_paths.len(), 2);
    }

    #[test]
    fn test_convert_options_with_landscape() {
        let opts = ConvertOptions {
            landscape: Some(true),
            ..Default::default()
        };
        assert_eq!(opts.landscape, Some(true));
    }

    #[test]
    fn test_convert_options_tagged_default_false() {
        let opts = ConvertOptions::default();
        assert!(!opts.tagged);
    }

    #[test]
    fn test_convert_options_pdf_ua_default_false() {
        let opts = ConvertOptions::default();
        assert!(!opts.pdf_ua);
    }

    #[test]
    fn test_convert_options_with_tagged() {
        let opts = ConvertOptions {
            tagged: true,
            ..Default::default()
        };
        assert!(opts.tagged);
    }

    #[test]
    fn test_convert_options_with_pdf_ua() {
        let opts = ConvertOptions {
            pdf_ua: true,
            ..Default::default()
        };
        assert!(opts.pdf_ua);
    }

    #[test]
    fn test_convert_options_streaming_default_false() {
        let opts = ConvertOptions::default();
        assert!(!opts.streaming);
    }

    #[test]
    fn test_convert_options_streaming_chunk_size_default_none() {
        let opts = ConvertOptions::default();
        assert!(opts.streaming_chunk_size.is_none());
    }

    #[test]
    fn test_convert_options_with_streaming() {
        let opts = ConvertOptions {
            streaming: true,
            ..Default::default()
        };
        assert!(opts.streaming);
    }

    #[test]
    fn test_convert_options_with_streaming_chunk_size() {
        let opts = ConvertOptions {
            streaming: true,
            streaming_chunk_size: Some(500),
            ..Default::default()
        };
        assert!(opts.streaming);
        assert_eq!(opts.streaming_chunk_size, Some(500));
    }
}

#[cfg(all(test, feature = "typescript"))]
mod ts_tests {
    use super::*;
    use ts_rs::TS;

    fn cfg() -> ts_rs::Config {
        ts_rs::Config::new()
    }

    #[test]
    fn test_format_ts_declaration() {
        let decl = Format::decl(&cfg());
        assert!(decl.contains("Format"), "Format TS decl: {decl}");
        assert!(decl.contains("Docx"), "should contain Docx variant");
        assert!(decl.contains("Pptx"), "should contain Pptx variant");
        assert!(decl.contains("Xlsx"), "should contain Xlsx variant");
    }

    #[test]
    fn test_paper_size_ts_declaration() {
        let decl = PaperSize::decl(&cfg());
        assert!(decl.contains("PaperSize"), "PaperSize TS decl: {decl}");
        assert!(decl.contains("A4"), "should contain A4 variant");
        assert!(decl.contains("Letter"), "should contain Letter variant");
        assert!(decl.contains("Legal"), "should contain Legal variant");
        assert!(decl.contains("Custom"), "should contain Custom variant");
    }

    #[test]
    fn test_pdf_standard_ts_declaration() {
        let decl = PdfStandard::decl(&cfg());
        assert!(decl.contains("PdfStandard"), "PdfStandard TS decl: {decl}");
        assert!(decl.contains("PdfA2b"), "should contain PdfA2b variant");
    }

    #[test]
    fn test_slide_range_ts_declaration() {
        let decl = SlideRange::decl(&cfg());
        assert!(decl.contains("SlideRange"), "SlideRange TS decl: {decl}");
        assert!(decl.contains("start"), "should contain start field");
        assert!(decl.contains("end"), "should contain end field");
        assert!(decl.contains("number"), "fields should be number type");
    }

    #[test]
    fn test_convert_options_ts_declaration() {
        let decl = ConvertOptions::decl(&cfg());
        assert!(
            decl.contains("ConvertOptions"),
            "ConvertOptions TS decl: {decl}"
        );
        assert!(
            decl.contains("tagged"),
            "should contain tagged field: {decl}"
        );
        assert!(
            decl.contains("pdf_ua"),
            "should contain pdf_ua field: {decl}"
        );
    }

    #[test]
    fn test_format_ts_export() {
        let ts = Format::export_to_string(&cfg()).unwrap();
        assert!(ts.contains("Format"));
    }

    #[test]
    fn test_convert_options_ts_export() {
        let ts = ConvertOptions::export_to_string(&cfg()).unwrap();
        assert!(ts.contains("ConvertOptions"));
    }
}
