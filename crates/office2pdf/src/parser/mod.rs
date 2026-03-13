pub(crate) mod chart;
pub(crate) mod cond_fmt;
pub mod docx;
pub(crate) mod embedded_fonts;
pub(crate) mod metadata;
pub(crate) mod omml;
pub mod pptx;
pub(crate) mod smartart;
pub(crate) mod units;
pub mod xlsx;
pub(crate) mod xml_util;

use std::io::Cursor;

use zip::ZipArchive;

#[cfg(test)]
#[path = "units_tests.rs"]
mod units_tests;

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

/// Open a byte slice as a ZIP archive, returning a `ConvertError::Parse` on failure.
pub(crate) fn open_zip(data: &[u8]) -> Result<ZipArchive<Cursor<&[u8]>>, ConvertError> {
    let cursor: Cursor<&[u8]> = Cursor::new(data);
    ZipArchive::new(cursor)
        .map_err(|error| parse_err(format!("Failed to open ZIP archive: {error}")))
}

/// Convenience constructor for `ConvertError::Parse`.
pub(crate) fn parse_err(msg: impl std::fmt::Display) -> ConvertError {
    ConvertError::Parse(msg.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_zip_returns_archive_for_valid_zip() {
        // Build a minimal valid ZIP in memory
        let buf: Vec<u8> = Vec::new();
        let cursor = Cursor::new(buf);
        let mut writer = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::default();
        writer.start_file("hello.txt", options).unwrap();
        std::io::Write::write_all(&mut writer, b"world").unwrap();
        let cursor = writer.finish().unwrap();
        let zip_bytes: Vec<u8> = cursor.into_inner();

        let mut archive = open_zip(&zip_bytes).expect("should open valid ZIP");
        assert_eq!(archive.len(), 1);
        let file = archive.by_name("hello.txt");
        assert!(file.is_ok());
    }

    #[test]
    fn open_zip_returns_parse_error_for_invalid_data() {
        let result = open_zip(b"this is not a zip file");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ConvertError::Parse(ref msg) if msg.contains("Failed to open ZIP archive")),
            "Expected Parse error with ZIP context, got: {err:?}"
        );
    }

    #[test]
    fn open_zip_returns_parse_error_for_empty_data() {
        let result = open_zip(b"");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ConvertError::Parse(_)),
            "Expected Parse error, got: {err:?}"
        );
    }

    #[test]
    fn parse_err_creates_parse_variant_with_string_message() {
        let err = parse_err("something went wrong");
        match err {
            ConvertError::Parse(msg) => assert_eq!(msg, "something went wrong"),
            other => panic!("Expected Parse variant, got: {other:?}"),
        }
    }

    #[test]
    fn parse_err_works_with_format_display_types() {
        let io_error = std::io::Error::new(std::io::ErrorKind::NotFound, "file missing");
        let err = parse_err(format!("I/O failure: {io_error}"));
        match err {
            ConvertError::Parse(msg) => assert!(
                msg.contains("I/O failure") && msg.contains("file missing"),
                "Unexpected message: {msg}"
            ),
            other => panic!("Expected Parse variant, got: {other:?}"),
        }
    }

    #[test]
    fn parse_err_accepts_display_impl_directly() {
        // Verify it works with any Display implementor, not just String/&str
        let number: i32 = 42;
        let err = parse_err(number);
        match err {
            ConvertError::Parse(msg) => assert_eq!(msg, "42"),
            other => panic!("Expected Parse variant, got: {other:?}"),
        }
    }
}
