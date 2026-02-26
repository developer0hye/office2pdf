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
