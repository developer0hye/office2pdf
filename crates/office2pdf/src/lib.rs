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

    // --- Image pipeline integration tests ---

    /// Build a minimal valid 1×1 red PNG with correct CRC checksums.
    fn make_test_png() -> Vec<u8> {
        /// Compute CRC32 over PNG chunk type + data.
        fn png_crc32(chunk_type: &[u8], data: &[u8]) -> u32 {
            let mut crc: u32 = 0xFFFF_FFFF;
            for &byte in chunk_type.iter().chain(data.iter()) {
                crc ^= byte as u32;
                for _ in 0..8 {
                    if crc & 1 != 0 {
                        crc = (crc >> 1) ^ 0xEDB8_8320;
                    } else {
                        crc >>= 1;
                    }
                }
            }
            crc ^ 0xFFFF_FFFF
        }

        let mut png = Vec::new();
        png.extend_from_slice(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        let ihdr_data: [u8; 13] = [
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00,
        ];
        let ihdr_type = b"IHDR";
        png.extend_from_slice(&(ihdr_data.len() as u32).to_be_bytes());
        png.extend_from_slice(ihdr_type);
        png.extend_from_slice(&ihdr_data);
        png.extend_from_slice(&png_crc32(ihdr_type, &ihdr_data).to_be_bytes());
        let idat_data: [u8; 15] = [
            0x78, 0x01, 0x01, 0x04, 0x00, 0xFB, 0xFF, 0x00, 0xFF, 0x00, 0x00, 0x03, 0x01, 0x01,
            0x00,
        ];
        let idat_type = b"IDAT";
        png.extend_from_slice(&(idat_data.len() as u32).to_be_bytes());
        png.extend_from_slice(idat_type);
        png.extend_from_slice(&idat_data);
        png.extend_from_slice(&png_crc32(idat_type, &idat_data).to_be_bytes());
        let iend_type = b"IEND";
        png.extend_from_slice(&0u32.to_be_bytes());
        png.extend_from_slice(iend_type);
        png.extend_from_slice(&png_crc32(iend_type, &[]).to_be_bytes());
        png
    }

    #[test]
    fn test_render_document_with_image() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Image(ImageData {
                    data: make_test_png(),
                    format: ImageFormat::Png,
                    width: Some(100.0),
                    height: Some(80.0),
                })],
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty(), "PDF should not be empty");
        assert!(pdf.starts_with(b"%PDF"), "Should be valid PDF");
    }

    #[test]
    fn test_render_document_image_mixed_with_text() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Image below:".to_string(),
                            style: TextStyle::default(),
                        }],
                    }),
                    Block::Image(ImageData {
                        data: make_test_png(),
                        format: ImageFormat::Png,
                        width: Some(200.0),
                        height: None,
                    }),
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Image above.".to_string(),
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
}
