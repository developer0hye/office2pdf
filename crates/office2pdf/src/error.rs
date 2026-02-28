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
///
/// Warnings are structured so that callers can programmatically inspect
/// what was degraded during conversion.
#[derive(Debug, Clone, PartialEq)]
pub enum ConvertWarning {
    /// An element type is not supported and was completely omitted.
    UnsupportedElement {
        /// Document format (e.g. "DOCX", "PPTX", "XLSX").
        format: String,
        /// Name or description of the unsupported element.
        element: String,
    },
    /// An element was partially rendered (some features degraded).
    PartialElement {
        /// Document format (e.g. "DOCX", "PPTX", "XLSX").
        format: String,
        /// Name or description of the element.
        element: String,
        /// Detail about what was degraded.
        detail: String,
    },
    /// A fallback representation was used instead of full rendering.
    FallbackUsed {
        /// Document format (e.g. "DOCX", "PPTX", "XLSX").
        format: String,
        /// Original element type.
        from: String,
        /// Fallback representation used.
        to: String,
    },
    /// An element was skipped during parsing.
    ParseSkipped {
        /// Document format (e.g. "DOCX", "PPTX", "XLSX").
        format: String,
        /// Reason the element was skipped.
        reason: String,
    },
}

impl ConvertWarning {
    /// Returns the document format associated with this warning.
    pub fn format(&self) -> &str {
        match self {
            Self::UnsupportedElement { format, .. }
            | Self::PartialElement { format, .. }
            | Self::FallbackUsed { format, .. }
            | Self::ParseSkipped { format, .. } => format,
        }
    }
}

impl std::fmt::Display for ConvertWarning {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UnsupportedElement { format, element } => {
                write!(f, "[{format}] unsupported element: {element}")
            }
            Self::PartialElement {
                format,
                element,
                detail,
            } => {
                write!(f, "[{format}] partial rendering of {element}: {detail}")
            }
            Self::FallbackUsed { format, from, to } => {
                write!(f, "[{format}] fallback: {from} rendered as {to}")
            }
            Self::ParseSkipped { format, reason } => {
                write!(f, "[{format}] skipped: {reason}")
            }
        }
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
    fn test_unsupported_element_display() {
        let w = ConvertWarning::UnsupportedElement {
            format: "DOCX".to_string(),
            element: "OLE object".to_string(),
        };
        assert_eq!(w.to_string(), "[DOCX] unsupported element: OLE object");
    }

    #[test]
    fn test_partial_element_display() {
        let w = ConvertWarning::PartialElement {
            format: "PPTX".to_string(),
            element: "scheme color".to_string(),
            detail: "tint modifier ignored".to_string(),
        };
        assert_eq!(
            w.to_string(),
            "[PPTX] partial rendering of scheme color: tint modifier ignored"
        );
    }

    #[test]
    fn test_fallback_used_display() {
        let w = ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "chart".to_string(),
            to: "data table".to_string(),
        };
        assert_eq!(
            w.to_string(),
            "[DOCX] fallback: chart rendered as data table"
        );
    }

    #[test]
    fn test_parse_skipped_display() {
        let w = ConvertWarning::ParseSkipped {
            format: "PPTX".to_string(),
            reason: "slide 3 failed to parse: missing XML".to_string(),
        };
        assert_eq!(
            w.to_string(),
            "[PPTX] skipped: slide 3 failed to parse: missing XML"
        );
    }

    #[test]
    fn test_warning_format_accessor() {
        let w = ConvertWarning::FallbackUsed {
            format: "XLSX".to_string(),
            from: "chart".to_string(),
            to: "data table".to_string(),
        };
        assert_eq!(w.format(), "XLSX");
    }

    #[test]
    fn test_warning_clone_and_eq() {
        let w = ConvertWarning::ParseSkipped {
            format: "DOCX".to_string(),
            reason: "element panicked".to_string(),
        };
        let w2 = w.clone();
        assert_eq!(w, w2);
    }

    #[test]
    fn test_convert_result_fields() {
        let result = ConvertResult {
            pdf: vec![0x25, 0x50, 0x44, 0x46],
            warnings: vec![ConvertWarning::UnsupportedElement {
                format: "DOCX".to_string(),
                element: "Image".to_string(),
            }],
        };
        assert_eq!(result.pdf, vec![0x25, 0x50, 0x44, 0x46]);
        assert_eq!(result.warnings.len(), 1);
        assert_eq!(result.warnings[0].format(), "DOCX");
    }

    #[test]
    fn test_convert_result_empty_warnings() {
        let result = ConvertResult {
            pdf: vec![1, 2, 3],
            warnings: vec![],
        };
        assert!(result.warnings.is_empty());
    }

    #[test]
    fn test_all_variants_carry_format() {
        let variants = [
            ConvertWarning::UnsupportedElement {
                format: "DOCX".to_string(),
                element: "x".to_string(),
            },
            ConvertWarning::PartialElement {
                format: "PPTX".to_string(),
                element: "x".to_string(),
                detail: "y".to_string(),
            },
            ConvertWarning::FallbackUsed {
                format: "XLSX".to_string(),
                from: "x".to_string(),
                to: "y".to_string(),
            },
            ConvertWarning::ParseSkipped {
                format: "DOCX".to_string(),
                reason: "x".to_string(),
            },
        ];
        let expected_formats = ["DOCX", "PPTX", "XLSX", "DOCX"];
        for (w, expected) in variants.iter().zip(expected_formats.iter()) {
            assert_eq!(w.format(), *expected);
        }
    }
}
