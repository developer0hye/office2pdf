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
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
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

/// Per-stage timing and size metrics from a conversion.
#[derive(Debug, Clone)]
#[cfg_attr(feature = "typescript", derive(ts_rs::TS))]
pub struct ConvertMetrics {
    /// Time spent parsing the input document (DOCX/PPTX/XLSX → IR).
    #[cfg_attr(feature = "typescript", ts(type = "number"))]
    pub parse_duration: std::time::Duration,
    /// Time spent generating Typst source code (IR → Typst).
    #[cfg_attr(feature = "typescript", ts(type = "number"))]
    pub codegen_duration: std::time::Duration,
    /// Time spent compiling Typst to PDF (Typst → PDF).
    #[cfg_attr(feature = "typescript", ts(type = "number"))]
    pub compile_duration: std::time::Duration,
    /// Total end-to-end conversion time.
    #[cfg_attr(feature = "typescript", ts(type = "number"))]
    pub total_duration: std::time::Duration,
    /// Size of the input file in bytes.
    pub input_size_bytes: u64,
    /// Size of the output PDF in bytes.
    pub output_size_bytes: u64,
    /// Number of pages in the output PDF.
    pub page_count: u32,
}

/// Result of a successful conversion, containing PDF bytes and any warnings.
#[derive(Debug)]
pub struct ConvertResult {
    /// The generated PDF bytes.
    pub pdf: Vec<u8>,
    /// Warnings collected during conversion (non-fatal issues).
    pub warnings: Vec<ConvertWarning>,
    /// Per-stage timing metrics, populated when instrumentation is enabled.
    pub metrics: Option<ConvertMetrics>,
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
            metrics: None,
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
            metrics: None,
        };
        assert!(result.warnings.is_empty());
    }

    #[test]
    fn test_convert_metrics_fields() {
        use std::time::Duration;
        let metrics = ConvertMetrics {
            parse_duration: Duration::from_millis(100),
            codegen_duration: Duration::from_millis(50),
            compile_duration: Duration::from_millis(200),
            total_duration: Duration::from_millis(360),
            input_size_bytes: 1024,
            output_size_bytes: 2048,
            page_count: 5,
        };
        assert_eq!(metrics.parse_duration, Duration::from_millis(100));
        assert_eq!(metrics.codegen_duration, Duration::from_millis(50));
        assert_eq!(metrics.compile_duration, Duration::from_millis(200));
        assert_eq!(metrics.total_duration, Duration::from_millis(360));
        assert_eq!(metrics.input_size_bytes, 1024);
        assert_eq!(metrics.output_size_bytes, 2048);
        assert_eq!(metrics.page_count, 5);
    }

    #[test]
    fn test_convert_metrics_clone() {
        use std::time::Duration;
        let metrics = ConvertMetrics {
            parse_duration: Duration::from_millis(10),
            codegen_duration: Duration::from_millis(20),
            compile_duration: Duration::from_millis(30),
            total_duration: Duration::from_millis(65),
            input_size_bytes: 512,
            output_size_bytes: 1024,
            page_count: 1,
        };
        let cloned = metrics.clone();
        assert_eq!(cloned.parse_duration, metrics.parse_duration);
        assert_eq!(cloned.total_duration, metrics.total_duration);
    }

    #[test]
    fn test_convert_result_with_metrics() {
        use std::time::Duration;
        let result = ConvertResult {
            pdf: vec![0x25, 0x50, 0x44, 0x46],
            warnings: vec![],
            metrics: Some(ConvertMetrics {
                parse_duration: Duration::from_millis(10),
                codegen_duration: Duration::from_millis(20),
                compile_duration: Duration::from_millis(30),
                total_duration: Duration::from_millis(65),
                input_size_bytes: 100,
                output_size_bytes: 200,
                page_count: 1,
            }),
        };
        assert!(result.metrics.is_some());
        let m = result.metrics.unwrap();
        assert_eq!(m.page_count, 1);
    }

    #[test]
    fn test_convert_error_debug_format() {
        let e = ConvertError::UnsupportedFormat("txt".to_string());
        let dbg = format!("{e:?}");
        assert!(dbg.contains("UnsupportedFormat"));
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

#[cfg(all(test, feature = "typescript"))]
mod ts_tests {
    use super::*;
    use ts_rs::TS;

    fn cfg() -> ts_rs::Config {
        ts_rs::Config::new()
    }

    #[test]
    fn test_convert_warning_ts_declaration() {
        let decl = ConvertWarning::decl(&cfg());
        assert!(
            decl.contains("ConvertWarning"),
            "ConvertWarning TS decl: {decl}"
        );
        assert!(
            decl.contains("UnsupportedElement"),
            "should contain UnsupportedElement variant: {decl}"
        );
        assert!(
            decl.contains("PartialElement"),
            "should contain PartialElement variant: {decl}"
        );
    }

    #[test]
    fn test_convert_metrics_ts_declaration() {
        let decl = ConvertMetrics::decl(&cfg());
        assert!(
            decl.contains("ConvertMetrics"),
            "ConvertMetrics TS decl: {decl}"
        );
        assert!(
            decl.contains("page_count"),
            "should contain page_count field: {decl}"
        );
        assert!(
            decl.contains("number"),
            "numeric fields should be number type: {decl}"
        );
    }

    #[test]
    fn test_convert_warning_ts_export() {
        let ts = ConvertWarning::export_to_string(&cfg()).unwrap();
        assert!(ts.contains("ConvertWarning"));
    }

    #[test]
    fn test_convert_metrics_ts_export() {
        let ts = ConvertMetrics::export_to_string(&cfg()).unwrap();
        assert!(ts.contains("ConvertMetrics"));
    }
}
