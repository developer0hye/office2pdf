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
    render_document(&doc)
}

/// Render an IR Document to PDF bytes.
///
/// Takes a fully constructed [`ir::Document`] and runs it through
/// the Typst codegen → PDF compilation pipeline.
pub fn render_document(doc: &ir::Document) -> Result<Vec<u8>, ConvertError> {
    let output = render::typst_gen::generate_typst(doc)?;
    render::pdf::compile_to_pdf(&output.source, &output.images)
}

#[cfg(test)]
mod tests {
    use super::*;
    use ir::*;

    /// Helper: create a minimal IR Document with a single FlowPage containing one paragraph.
    fn make_simple_document(text: &str) -> Document {
        Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: text.to_string(),
                        style: TextStyle::default(),
                    }],
                })],
            })],
            styles: StyleSheet::default(),
        }
    }

    // --- Format detection tests ---

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
        assert!(matches!(
            result.unwrap_err(),
            ConvertError::UnsupportedFormat(_)
        ));
    }

    #[test]
    fn test_format_detection_all_supported_extensions() {
        // Tested via config.rs, but verify pipeline dispatches correctly
        assert!(convert_bytes(b"fake", Format::Docx, &ConvertOptions::default()).is_err());
        assert!(convert_bytes(b"fake", Format::Pptx, &ConvertOptions::default()).is_err());
        assert!(convert_bytes(b"fake", Format::Xlsx, &ConvertOptions::default()).is_err());
    }

    #[test]
    fn test_convert_bytes_propagates_parse_error() {
        // All stub parsers should return Parse errors
        for format in [Format::Docx, Format::Pptx, Format::Xlsx] {
            let result = convert_bytes(b"fake", format, &ConvertOptions::default());
            assert!(result.is_err());
            assert!(
                matches!(result.unwrap_err(), ConvertError::Parse(_)),
                "Expected Parse error for {format:?}"
            );
        }
    }

    #[test]
    fn test_convert_nonexistent_file_returns_io_error() {
        let result = convert("nonexistent_file.docx");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
    }

    // --- Pipeline integration tests with mock IR documents ---

    #[test]
    fn test_render_document_empty_document() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty(), "PDF bytes should not be empty");
        assert!(pdf.starts_with(b"%PDF"), "Should be valid PDF");
    }

    #[test]
    fn test_render_document_single_paragraph() {
        let doc = make_simple_document("Hello, World!");
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_render_document_styled_text() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Center),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![
                        Run {
                            text: "Bold text ".to_string(),
                            style: TextStyle {
                                bold: Some(true),
                                font_size: Some(16.0),
                                ..TextStyle::default()
                            },
                        },
                        Run {
                            text: "and italic".to_string(),
                            style: TextStyle {
                                italic: Some(true),
                                color: Some(Color::new(255, 0, 0)),
                                ..TextStyle::default()
                            },
                        },
                    ],
                })],
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_render_document_multiple_flow_pages() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![
                Page::Flow(FlowPage {
                    size: PageSize::default(),
                    margins: Margins::default(),
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Page 1".to_string(),
                            style: TextStyle::default(),
                        }],
                    })],
                }),
                Page::Flow(FlowPage {
                    size: PageSize::default(),
                    margins: Margins::default(),
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Page 2".to_string(),
                            style: TextStyle::default(),
                        }],
                    })],
                }),
            ],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_render_document_page_break() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Before break".to_string(),
                            style: TextStyle::default(),
                        }],
                    }),
                    Block::PageBreak,
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "After break".to_string(),
                            style: TextStyle::default(),
                        }],
                    }),
                ],
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_with_options_delegates_to_convert_bytes() {
        // convert_with_options on a nonexistent file should produce an IO error,
        // confirming it reads the file before calling convert_bytes
        let result = convert_with_options("nonexistent.docx", &ConvertOptions::default());
        assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
    }

    #[test]
    fn test_convert_delegates_to_convert_with_options() {
        // convert("nonexistent.docx") should behave same as convert_with_options
        let result = convert("nonexistent.docx");
        assert!(matches!(result.unwrap_err(), ConvertError::Io(_)));
    }
}
