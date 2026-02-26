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

/// Options controlling the conversion process.
#[derive(Debug, Clone, Default)]
pub struct ConvertOptions {
    // Placeholder for future options (paper size, font paths, etc.)
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
}
