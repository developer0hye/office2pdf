pub(crate) mod chart;
pub(crate) mod cond_fmt;
pub mod docx;
pub(crate) mod metadata;
pub(crate) mod omml;
pub mod pptx;
pub(crate) mod smartart;
pub mod xlsx;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::Document;

/// Trait for parsing an input file format into the IR.
pub trait Parser {
    /// Parse raw file bytes into a Document IR and any non-fatal warnings.
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError>;
}
