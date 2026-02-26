use crate::error::ConvertError;
use crate::ir::Document;
use crate::parser::Parser;

pub struct DocxParser;

impl Parser for DocxParser {
    fn parse(&self, _data: &[u8]) -> Result<Document, ConvertError> {
        Err(ConvertError::Parse(
            "DOCX parser not yet implemented".to_string(),
        ))
    }
}
