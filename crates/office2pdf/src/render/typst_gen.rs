use crate::error::ConvertError;
use crate::ir::Document;

/// Generate Typst markup from a Document IR.
pub fn generate_typst(_doc: &Document) -> Result<String, ConvertError> {
    Err(ConvertError::Render(
        "Typst codegen not yet implemented".to_string(),
    ))
}
