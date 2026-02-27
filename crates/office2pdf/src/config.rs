/// Supported input document formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
pub enum PdfStandard {
    /// PDF/A-2b for archival purposes.
    PdfA2b,
}

/// Options controlling the conversion process.
#[derive(Debug, Clone, Default)]
pub struct ConvertOptions {
    /// Filter XLSX sheets by name. Only sheets whose names are in this list
    /// will be included. If `None`, all sheets are included.
    pub sheet_names: Option<Vec<String>>,
    /// Filter PPTX slides by range (1-indexed). If `None`, all slides are included.
    pub slide_range: Option<SlideRange>,
    /// PDF standard to enforce. If `None`, produces a standard PDF 1.7.
    pub pdf_standard: Option<PdfStandard>,
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
}
