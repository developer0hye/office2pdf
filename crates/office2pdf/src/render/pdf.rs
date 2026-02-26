use crate::error::ConvertError;

/// Compile Typst markup to PDF bytes.
pub fn compile_to_pdf(_typst_source: &str) -> Result<Vec<u8>, ConvertError> {
    Err(ConvertError::Render(
        "PDF compilation not yet implemented".to_string(),
    ))
}
