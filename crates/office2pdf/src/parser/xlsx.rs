use crate::error::ConvertError;
use crate::ir::Document;
use crate::parser::Parser;

pub struct XlsxParser;

impl Parser for XlsxParser {
    fn parse(&self, _data: &[u8]) -> Result<Document, ConvertError> {
        Err(ConvertError::Parse(
            "XLSX parser not yet implemented".to_string(),
        ))
    }
}
