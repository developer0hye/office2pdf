use thiserror::Error;

/// Errors that can occur during document conversion.
#[derive(Debug, Error)]
pub enum ConvertError {
    #[error("unsupported file format: {0}")]
    UnsupportedFormat(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("parse error: {0}")]
    Parse(String),

    #[error("render error: {0}")]
    Render(String),
}

/// A non-fatal warning emitted when an element cannot be fully processed.
#[derive(Debug, Clone)]
pub struct ConvertWarning {
    /// Description of the element that caused the warning.
    pub element: String,
    /// Reason the element could not be processed.
    pub reason: String,
}

impl std::fmt::Display for ConvertWarning {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}: {}", self.element, self.reason)
    }
}

/// Result of a successful conversion, containing PDF bytes and any warnings.
#[derive(Debug)]
pub struct ConvertResult {
    /// The generated PDF bytes.
    pub pdf: Vec<u8>,
    /// Warnings collected during conversion (non-fatal issues).
    pub warnings: Vec<ConvertWarning>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_convert_warning_display() {
        let w = ConvertWarning {
            element: "Table at index 2".to_string(),
            reason: "malformed cell structure".to_string(),
        };
        assert_eq!(w.to_string(), "Table at index 2: malformed cell structure");
    }

    #[test]
    fn test_convert_warning_clone() {
        let w = ConvertWarning {
            element: "Paragraph".to_string(),
            reason: "unsupported style".to_string(),
        };
        let w2 = w.clone();
        assert_eq!(w2.element, "Paragraph");
        assert_eq!(w2.reason, "unsupported style");
    }

    #[test]
    fn test_convert_result_fields() {
        let result = ConvertResult {
            pdf: vec![0x25, 0x50, 0x44, 0x46],
            warnings: vec![ConvertWarning {
                element: "Image".to_string(),
                reason: "missing data".to_string(),
            }],
        };
        assert_eq!(result.pdf, vec![0x25, 0x50, 0x44, 0x46]);
        assert_eq!(result.warnings.len(), 1);
        assert_eq!(result.warnings[0].element, "Image");
    }

    #[test]
    fn test_convert_result_empty_warnings() {
        let result = ConvertResult {
            pdf: vec![1, 2, 3],
            warnings: vec![],
        };
        assert!(result.warnings.is_empty());
    }
}
