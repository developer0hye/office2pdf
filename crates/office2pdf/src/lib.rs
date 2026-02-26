pub mod config;
pub mod error;
pub mod ir;
pub mod parser;
pub mod render;

use std::path::Path;

use config::{ConvertOptions, Format};
use error::ConvertError;
use parser::Parser;

/// Convert a file at the given path to PDF bytes.
pub fn convert(path: impl AsRef<Path>) -> Result<Vec<u8>, ConvertError> {
    convert_with_options(path, &ConvertOptions::default())
}

/// Convert a file at the given path to PDF bytes with options.
pub fn convert_with_options(
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<Vec<u8>, ConvertError> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .ok_or_else(|| ConvertError::UnsupportedFormat("no file extension".to_string()))?;

    let format = Format::from_extension(ext)
        .ok_or_else(|| ConvertError::UnsupportedFormat(ext.to_string()))?;

    let data = std::fs::read(path)?;
    convert_bytes(&data, format, options)
}

/// Convert raw bytes of a known format to PDF bytes.
pub fn convert_bytes(
    data: &[u8],
    format: Format,
    _options: &ConvertOptions,
) -> Result<Vec<u8>, ConvertError> {
    let parser: Box<dyn Parser> = match format {
        Format::Docx => Box::new(parser::docx::DocxParser),
        Format::Pptx => Box::new(parser::pptx::PptxParser),
        Format::Xlsx => Box::new(parser::xlsx::XlsxParser),
    };

    let doc = parser.parse(data)?;
    let typst_source = render::typst_gen::generate_typst(&doc)?;
    render::pdf::compile_to_pdf(&typst_source)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_convert_unsupported_format() {
        let result = convert("test.txt");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(matches!(err, ConvertError::UnsupportedFormat(_)));
    }

    #[test]
    fn test_convert_no_extension() {
        let result = convert("test");
        assert!(result.is_err());
    }

    #[test]
    fn test_convert_bytes_docx_not_yet_implemented() {
        let result = convert_bytes(b"fake", Format::Docx, &ConvertOptions::default());
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), ConvertError::Parse(_)));
    }
}
