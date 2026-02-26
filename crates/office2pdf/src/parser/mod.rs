pub mod docx;
pub mod pptx;
pub mod xlsx;

use crate::error::{ConvertError, ConvertWarning};
use crate::ir::Document;

/// Trait for parsing an input file format into the IR.
pub trait Parser {
    /// Parse raw file bytes into a Document IR and any non-fatal warnings.
    fn parse(&self, data: &[u8]) -> Result<(Document, Vec<ConvertWarning>), ConvertError>;
}
