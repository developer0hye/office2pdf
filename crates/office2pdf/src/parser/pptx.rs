use crate::error::ConvertError;
use crate::ir::Document;
use crate::parser::Parser;

pub struct PptxParser;

impl Parser for PptxParser {
    fn parse(&self, _data: &[u8]) -> Result<Document, ConvertError> {
        Err(ConvertError::Parse(
            "PPTX parser not yet implemented".to_string(),
        ))
    }
}
