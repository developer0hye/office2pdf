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

    #[error("file is encrypted/password-protected and cannot be converted")]
    UnsupportedEncryption,
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
#[path = "error_tests.rs"]
mod tests;

#[cfg(all(test, feature = "typescript"))]
#[path = "error_ts_tests.rs"]
mod ts_tests;
