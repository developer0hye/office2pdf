pub mod config;
pub mod error;
pub mod ir;
pub mod parser;
pub mod render;

use std::path::Path;

use config::{ConvertOptions, Format};
use error::{ConvertError, ConvertResult};
use parser::Parser;

/// Convert a file at the given path to PDF bytes with warnings.
pub fn convert(path: impl AsRef<Path>) -> Result<ConvertResult, ConvertError> {
    convert_with_options(path, &ConvertOptions::default())
}

/// Convert a file at the given path to PDF bytes with options.
pub fn convert_with_options(
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
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

/// Convert raw bytes of a known format to PDF bytes with warnings.
pub fn convert_bytes(
    data: &[u8],
    format: Format,
    _options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    let parser: Box<dyn Parser> = match format {
        Format::Docx => Box::new(parser::docx::DocxParser),
        Format::Pptx => Box::new(parser::pptx::PptxParser),
        Format::Xlsx => Box::new(parser::xlsx::XlsxParser),
    };

    let (doc, warnings) = parser.parse(data)?;
    let pdf = render_document(&doc)?;
    Ok(ConvertResult { pdf, warnings })
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
                header: None,
                footer: None,
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
                header: None,
                footer: None,
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
                    header: None,
                    footer: None,
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
                    header: None,
                    footer: None,
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
                header: None,
                footer: None,
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
                header: None,
                footer: None,
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
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    // --- End-to-end integration tests: raw document bytes → PDF ---

    /// Build a minimal DOCX as bytes using docx-rs builder.
    fn build_test_docx() -> Vec<u8> {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello from DOCX")),
            )
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Second paragraph").bold()),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Build a minimal XLSX as bytes using umya-spreadsheet.
    fn build_test_xlsx() -> Vec<u8> {
        use std::io::Cursor;
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            sheet.get_cell_mut("A1").set_value("Name");
            sheet.get_cell_mut("B1").set_value("Value");
            sheet.get_cell_mut("A2").set_value("Item 1");
            sheet.get_cell_mut("B2").set_value("100");
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Build a minimal PPTX as bytes using zip + raw XML.
    fn build_test_pptx() -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        ).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/presentation.xml
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
        ).unwrap();

        // ppt/_rels/presentation.xml.rels
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/slides/slide1.xml — one text box with "Hello from PPTX"
        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello from PPTX</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
        ).unwrap();

        // ppt/slides/_rels/slide1.xml.rels (empty rels)
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
        ).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_e2e_docx_to_pdf() {
        let docx_bytes = build_test_docx();
        let result = convert_bytes(&docx_bytes, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            !result.pdf.is_empty(),
            "DOCX→PDF should produce non-empty output"
        );
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Output should be valid PDF"
        );
        assert!(
            result.warnings.is_empty(),
            "Normal DOCX should produce no warnings"
        );
    }

    #[test]
    fn test_e2e_xlsx_to_pdf() {
        let xlsx_bytes = build_test_xlsx();
        let result = convert_bytes(&xlsx_bytes, Format::Xlsx, &ConvertOptions::default()).unwrap();
        assert!(
            !result.pdf.is_empty(),
            "XLSX→PDF should produce non-empty output"
        );
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Output should be valid PDF"
        );
    }

    #[test]
    fn test_e2e_pptx_to_pdf() {
        let pptx_bytes = build_test_pptx();
        let result = convert_bytes(&pptx_bytes, Format::Pptx, &ConvertOptions::default()).unwrap();
        assert!(
            !result.pdf.is_empty(),
            "PPTX→PDF should produce non-empty output"
        );
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Output should be valid PDF"
        );
    }

    #[test]
    fn test_e2e_docx_with_table_to_pdf() {
        use std::io::Cursor;
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell A")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell B")),
            ),
        ])]);
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Table below:")),
            )
            .add_table(table);
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(!result.pdf.is_empty());
        assert!(result.pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_e2e_convert_with_options_from_temp_file() {
        let docx_bytes = build_test_docx();
        let dir = std::env::temp_dir();
        let input = dir.join("office2pdf_test_input.docx");
        let output = dir.join("office2pdf_test_output.pdf");
        std::fs::write(&input, &docx_bytes).unwrap();

        let result = convert(&input).unwrap();
        assert!(!result.pdf.is_empty());
        assert!(result.pdf.starts_with(b"%PDF"));

        // Also test convert_with_options with the file path
        let result2 = convert_with_options(&input, &ConvertOptions::default()).unwrap();
        assert!(!result2.pdf.is_empty());
        assert!(result2.pdf.starts_with(b"%PDF"));

        // Write PDF to output and verify file exists
        std::fs::write(&output, &result.pdf).unwrap();
        assert!(output.exists());
        let written = std::fs::read(&output).unwrap();
        assert!(written.starts_with(b"%PDF"));

        // Cleanup
        let _ = std::fs::remove_file(&input);
        let _ = std::fs::remove_file(&output);
    }

    #[test]
    fn test_e2e_unsupported_format_error_message() {
        let result = convert("document.odt");
        let err = result.unwrap_err();
        match err {
            ConvertError::UnsupportedFormat(ref ext) => {
                assert_eq!(ext, "odt", "Error should mention the unsupported extension");
            }
            _ => panic!("Expected UnsupportedFormat error, got {err:?}"),
        }
    }

    #[test]
    fn test_e2e_missing_file_error() {
        let result = convert("nonexistent_document.docx");
        assert!(
            matches!(result.unwrap_err(), ConvertError::Io(_)),
            "Missing file should produce IO error"
        );
    }

    // --- US-018: Font fallback - system font discovery tests ---

    #[test]
    fn test_render_document_with_system_font_in_ir() {
        // A Document with a system font name (e.g., "Arial") in the IR should
        // compile to valid PDF. With system font discovery enabled, the font
        // is used if available; otherwise Typst falls back to embedded fonts.
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Hello with system font".to_string(),
                        style: TextStyle {
                            font_family: Some("Arial".to_string()),
                            ..TextStyle::default()
                        },
                    }],
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_render_document_with_multiple_font_families() {
        // Different runs can specify different system fonts — all should compile
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![
                        Run {
                            text: "Calibri text ".to_string(),
                            style: TextStyle {
                                font_family: Some("Calibri".to_string()),
                                ..TextStyle::default()
                            },
                        },
                        Run {
                            text: "and Times New Roman text".to_string(),
                            style: TextStyle {
                                font_family: Some("Times New Roman".to_string()),
                                ..TextStyle::default()
                            },
                        },
                    ],
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    // --- US-017: Enhanced error handling tests ---

    /// Build a PPTX with two slides: one valid, one with broken XML.
    /// The parser should skip the broken slide with a warning and still produce a PDF.
    fn build_pptx_with_broken_slide() -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::FileOptions::default();

        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        ).unwrap();

        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

        // Two slides referenced
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst></p:presentation>"#,
        ).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
        ).unwrap();

        // Slide 1: valid
        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Valid slide content</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
        ).unwrap();

        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
        ).unwrap();

        // Slide 2: intentionally missing from the ZIP archive.
        // The presentation.xml references it via rId3, but no slide2.xml exists.
        // This should trigger a warning (missing slide file), not a fatal error.

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_convert_result_has_pdf_and_warnings() {
        let docx_bytes = build_test_docx();
        let result = convert_bytes(&docx_bytes, Format::Docx, &ConvertOptions::default()).unwrap();
        // ConvertResult has both pdf and warnings fields
        assert!(result.pdf.starts_with(b"%PDF"));
        let _warnings: &Vec<error::ConvertWarning> = &result.warnings;
    }

    #[test]
    fn test_pptx_broken_slide_emits_warning_and_produces_pdf() {
        let pptx_bytes = build_pptx_with_broken_slide();
        let result = convert_bytes(&pptx_bytes, Format::Pptx, &ConvertOptions::default()).unwrap();

        // Should still produce a valid PDF (from the good slide)
        assert!(
            !result.pdf.is_empty(),
            "Should produce PDF despite broken slide"
        );
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Output should be valid PDF"
        );

        // Should have at least one warning about the broken slide
        assert!(
            !result.warnings.is_empty(),
            "Should emit warning for broken slide"
        );
        // Verify the warning mentions the broken element
        let warning_text = result.warnings[0].to_string();
        assert!(
            warning_text.contains("slide") || warning_text.contains("Slide"),
            "Warning should mention the problematic slide: {warning_text}"
        );
    }

    #[test]
    fn test_render_document_with_list() {
        use crate::ir::{
            Document, FlowPage, List, ListItem, ListKind, Margins, Metadata, Page, PageSize,
            Paragraph, ParagraphStyle, Run, StyleSheet, TextStyle,
        };
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![ir::Block::List(List {
                    kind: ListKind::Unordered,
                    items: vec![
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Hello".to_string(),
                                    style: TextStyle::default(),
                                }],
                            }],
                            level: 0,
                        },
                        ListItem {
                            content: vec![Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "World".to_string(),
                                    style: TextStyle::default(),
                                }],
                            }],
                            level: 0,
                        },
                    ],
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(
            pdf.starts_with(b"%PDF"),
            "Should produce valid PDF with list"
        );
    }

    #[test]
    fn test_e2e_docx_with_list_produces_pdf() {
        use std::io::Cursor;
        // Build a DOCX with a bulleted list and verify it converts to PDF
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ));
        let numbering = docx_rs::Numbering::new(1, 0);
        let nums = docx_rs::Numberings::new()
            .add_abstract_numbering(abstract_num)
            .add_numbering(numbering);

        let docx = docx_rs::Docx::new()
            .numberings(nums)
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Bullet 1"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            )
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Bullet 2"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            )
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Regular text")),
            );

        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Should produce valid PDF with list content"
        );
    }

    #[test]
    fn test_normal_docx_has_no_warnings() {
        let docx_bytes = build_test_docx();
        let result = convert_bytes(&docx_bytes, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.warnings.is_empty(),
            "Normal DOCX should produce no warnings"
        );
    }

    // --- US-020: Header/footer integration tests ---

    #[test]
    fn test_render_document_with_header() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Body content".to_string(),
                        style: TextStyle::default(),
                    }],
                })],
                header: Some(ir::HeaderFooter {
                    paragraphs: vec![ir::HeaderFooterParagraph {
                        style: ParagraphStyle::default(),
                        elements: vec![ir::HFInline::Run(Run {
                            text: "My Header".to_string(),
                            style: TextStyle::default(),
                        })],
                    }],
                }),
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_render_document_with_page_number_footer() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Body content".to_string(),
                        style: TextStyle::default(),
                    }],
                })],
                header: None,
                footer: Some(ir::HeaderFooter {
                    paragraphs: vec![ir::HeaderFooterParagraph {
                        style: ParagraphStyle::default(),
                        elements: vec![
                            ir::HFInline::Run(Run {
                                text: "Page ".to_string(),
                                style: TextStyle::default(),
                            }),
                            ir::HFInline::PageNumber,
                        ],
                    }],
                }),
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(!pdf.is_empty());
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_e2e_docx_with_header_footer_to_pdf() {
        use std::io::Cursor;
        let header = docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Document Title")),
        );
        let footer = docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Page ")
                    .add_field_char(docx_rs::FieldCharType::Begin, false)
                    .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                    .add_field_char(docx_rs::FieldCharType::Separate, false)
                    .add_text("1")
                    .add_field_char(docx_rs::FieldCharType::End, false),
            ),
        );
        let docx = docx_rs::Docx::new()
            .header(header)
            .footer(footer)
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body paragraph")),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "DOCX with header/footer should produce valid PDF"
        );
    }

    // --- US-021: Page orientation (landscape/portrait) tests ---

    #[test]
    fn test_render_document_with_landscape_page() {
        // A landscape FlowPage should render to valid PDF
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize {
                    width: 841.9, // A4 landscape
                    height: 595.3,
                },
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    runs: vec![Run {
                        text: "Landscape page".to_string(),
                        style: TextStyle::default(),
                    }],
                    style: ParagraphStyle::default(),
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(
            !pdf.is_empty(),
            "Landscape FlowPage should produce non-empty PDF"
        );
        assert!(pdf.starts_with(b"%PDF"), "Should produce valid PDF");
    }

    #[test]
    fn test_e2e_landscape_docx_to_pdf() {
        use std::io::Cursor;
        // Build a landscape DOCX with swapped dimensions
        let docx = docx_rs::Docx::new()
            .page_size(16838, 11906)
            .page_orient(docx_rs::PageOrientationType::Landscape)
            .page_margin(
                docx_rs::PageMargin::new()
                    .top(1440)
                    .bottom(1440)
                    .left(1440)
                    .right(1440),
            )
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Landscape document")),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Landscape DOCX should produce valid PDF"
        );
    }
}
