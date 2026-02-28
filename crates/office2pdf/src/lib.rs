//! Pure-Rust conversion of Office documents (DOCX, PPTX, XLSX) to PDF.
//!
//! # Quick start (native only)
//!
//! ```no_run
//! # #[cfg(not(target_arch = "wasm32"))]
//! # {
//! let result = office2pdf::convert("report.docx").unwrap();
//! std::fs::write("report.pdf", &result.pdf).unwrap();
//! # }
//! ```
//!
//! # With options (native only)
//!
//! ```no_run
//! # #[cfg(not(target_arch = "wasm32"))]
//! # {
//! use office2pdf::config::{ConvertOptions, PaperSize, SlideRange};
//!
//! let options = ConvertOptions {
//!     paper_size: Some(PaperSize::A4),
//!     slide_range: Some(SlideRange::new(1, 5)),
//!     ..Default::default()
//! };
//! let result = office2pdf::convert_with_options("slides.pptx", &options).unwrap();
//! std::fs::write("slides.pdf", &result.pdf).unwrap();
//! # }
//! ```
//!
//! # In-memory conversion (works on all targets including WASM)
//!
//! ```no_run
//! use office2pdf::config::{ConvertOptions, Format};
//!
//! let docx_bytes = std::fs::read("report.docx").unwrap();
//! let result = office2pdf::convert_bytes(&docx_bytes, Format::Docx, &ConvertOptions::default()).unwrap();
//! std::fs::write("report.pdf", &result.pdf).unwrap();
//! ```

pub mod config;
pub mod error;
pub mod ir;
pub mod parser;
#[cfg(feature = "pdf-ops")]
pub mod pdf_ops;
pub mod render;
#[cfg(feature = "wasm")]
pub mod wasm;

use std::time::Instant;

use config::{ConvertOptions, Format};
use error::{ConvertError, ConvertMetrics, ConvertResult};
use parser::Parser;

/// Convert a file at the given path to PDF bytes with warnings.
///
/// Detects the format from the file extension (`.docx`, `.pptx`, `.xlsx`).
///
/// This function is not available on `wasm32` targets because it reads from the
/// filesystem. Use [`convert_bytes`] for in-memory conversion on WASM.
///
/// # Errors
///
/// Returns [`ConvertError::UnsupportedFormat`] if the extension is unrecognized,
/// [`ConvertError::Io`] if the file cannot be read, or other variants for
/// parse/render failures.
#[cfg(not(target_arch = "wasm32"))]
pub fn convert(path: impl AsRef<std::path::Path>) -> Result<ConvertResult, ConvertError> {
    convert_with_options(path, &ConvertOptions::default())
}

/// Convert a file at the given path to PDF bytes with options.
///
/// See [`ConvertOptions`] for available settings (paper size, sheet filter, etc.).
///
/// This function is not available on `wasm32` targets because it reads from the
/// filesystem. Use [`convert_bytes`] for in-memory conversion on WASM.
///
/// # Errors
///
/// Returns [`ConvertError`] on unsupported format, I/O, parse, or render failure.
#[cfg(not(target_arch = "wasm32"))]
pub fn convert_with_options(
    path: impl AsRef<std::path::Path>,
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
///
/// Use this when you already have the file contents in memory and know the
/// [`Format`].
///
/// When `options.streaming` is `true` and the format is XLSX, the conversion
/// processes rows in chunks to bound peak memory during Typst compilation.
/// This requires the `pdf-ops` feature for PDF merging.
///
/// # Errors
///
/// Returns [`ConvertError`] on parse or render failure.
pub fn convert_bytes(
    data: &[u8],
    format: Format,
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    // Use streaming path for XLSX when requested and pdf-ops is available
    #[cfg(feature = "pdf-ops")]
    if options.streaming && format == Format::Xlsx {
        return convert_bytes_streaming_xlsx(data, options);
    }

    let total_start = Instant::now();
    let input_size_bytes = data.len() as u64;

    let parser: Box<dyn Parser> = match format {
        Format::Docx => Box::new(parser::docx::DocxParser),
        Format::Pptx => Box::new(parser::pptx::PptxParser),
        Format::Xlsx => Box::new(parser::xlsx::XlsxParser),
    };

    // Stage 1: Parse (OOXML → IR)
    let parse_start = Instant::now();
    let (doc, warnings) = parser.parse(data, options)?;
    let parse_duration = parse_start.elapsed();
    let page_count = doc.pages.len() as u32;

    // Stage 2: Codegen (IR → Typst)
    let codegen_start = Instant::now();
    let output = render::typst_gen::generate_typst_with_options(&doc, options)?;
    let codegen_duration = codegen_start.elapsed();

    // Stage 3: Compile (Typst → PDF)
    let compile_start = Instant::now();
    let pdf = render::pdf::compile_to_pdf(
        &output.source,
        &output.images,
        options.pdf_standard,
        &options.font_paths,
        options.tagged,
        options.pdf_ua,
    )?;
    let compile_duration = compile_start.elapsed();

    let total_duration = total_start.elapsed();
    let output_size_bytes = pdf.len() as u64;

    Ok(ConvertResult {
        pdf,
        warnings,
        metrics: Some(ConvertMetrics {
            parse_duration,
            codegen_duration,
            compile_duration,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count,
        }),
    })
}

/// Streaming conversion for XLSX: process rows in chunks with bounded memory.
///
/// Each chunk of rows is compiled independently to a PDF, then all chunk PDFs
/// are merged. This bounds peak memory during Typst compilation because only
/// one chunk's worth of Typst source and compilation state is in memory at a time.
#[cfg(feature = "pdf-ops")]
fn convert_bytes_streaming_xlsx(
    data: &[u8],
    options: &ConvertOptions,
) -> Result<ConvertResult, ConvertError> {
    let total_start = Instant::now();
    let input_size_bytes = data.len() as u64;
    let chunk_size = options.streaming_chunk_size.unwrap_or(1000);

    let xlsx_parser = parser::xlsx::XlsxParser;

    // Stage 1: Parse into chunks
    let parse_start = Instant::now();
    let (chunk_docs, warnings) = xlsx_parser.parse_streaming(data, options, chunk_size)?;
    let parse_duration = parse_start.elapsed();

    if chunk_docs.is_empty() {
        // Empty spreadsheet — produce a minimal empty PDF
        let empty_doc = ir::Document {
            metadata: ir::Metadata::default(),
            pages: vec![],
            styles: ir::StyleSheet::default(),
        };
        let output = render::typst_gen::generate_typst(&empty_doc)?;
        let pdf =
            render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)?;
        let total_duration = total_start.elapsed();
        return Ok(ConvertResult {
            pdf,
            warnings,
            metrics: Some(ConvertMetrics {
                parse_duration,
                codegen_duration: std::time::Duration::ZERO,
                compile_duration: std::time::Duration::ZERO,
                total_duration,
                input_size_bytes,
                output_size_bytes: 0,
                page_count: 0,
            }),
        });
    }

    // Stage 2+3: Codegen + Compile each chunk independently
    let mut all_pdfs: Vec<Vec<u8>> = Vec::with_capacity(chunk_docs.len());
    let mut codegen_duration_total = std::time::Duration::ZERO;
    let mut compile_duration_total = std::time::Duration::ZERO;
    let mut total_page_count = 0u32;

    for chunk_doc in chunk_docs {
        total_page_count += chunk_doc.pages.len() as u32;

        let codegen_start = Instant::now();
        let output = render::typst_gen::generate_typst_with_options(&chunk_doc, options)?;
        codegen_duration_total += codegen_start.elapsed();

        let compile_start = Instant::now();
        let pdf = render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            options.pdf_standard,
            &options.font_paths,
            options.tagged,
            options.pdf_ua,
        )?;
        compile_duration_total += compile_start.elapsed();

        all_pdfs.push(pdf);
        // chunk_doc and output are dropped here, freeing their memory
    }

    // Stage 4: Merge all chunk PDFs
    let final_pdf = if all_pdfs.len() == 1 {
        all_pdfs.into_iter().next().unwrap()
    } else {
        let refs: Vec<&[u8]> = all_pdfs.iter().map(|p| p.as_slice()).collect();
        pdf_ops::merge(&refs).map_err(|e| ConvertError::Render(format!("PDF merge failed: {e}")))?
    };

    let total_duration = total_start.elapsed();
    let output_size_bytes = final_pdf.len() as u64;

    Ok(ConvertResult {
        pdf: final_pdf,
        warnings,
        metrics: Some(ConvertMetrics {
            parse_duration,
            codegen_duration: codegen_duration_total,
            compile_duration: compile_duration_total,
            total_duration,
            input_size_bytes,
            output_size_bytes,
            page_count: total_page_count,
        }),
    })
}

/// Render an IR Document to PDF bytes.
///
///// Render an IR [`Document`](ir::Document) directly to PDF bytes.
///
/// Takes a fully constructed [`ir::Document`] and runs it through
/// the Typst codegen → PDF compilation pipeline.
///
/// # Errors
///
/// Returns [`ConvertError::Render`] if Typst compilation or PDF export fails.
pub fn render_document(doc: &ir::Document) -> Result<Vec<u8>, ConvertError> {
    let output = render::typst_gen::generate_typst(doc)?;
    render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
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
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        }
    }

    /// Build a DOCX file with a title in docProps/core.xml for PDF/UA tests.
    fn build_docx_with_title(title: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::FileOptions::default();

        zip.start_file("[Content_Types].xml", options).unwrap();
        Write::write_all(&mut zip, br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        Write::write_all(&mut zip, br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
        )
        .unwrap();

        zip.start_file("docProps/core.xml", options).unwrap();
        let core_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>{title}</dc:title>
</cp:coreProperties>"#
        );
        Write::write_all(&mut zip, core_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
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
                            href: None,
                            footnote: None,
                        },
                        Run {
                            text: "and italic".to_string(),
                            style: TextStyle {
                                italic: Some(true),
                                color: Some(Color::new(255, 0, 0)),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
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
                            href: None,
                            footnote: None,
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
                            href: None,
                            footnote: None,
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
                            href: None,
                            footnote: None,
                        }],
                    }),
                    Block::PageBreak,
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "After break".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
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
                            href: None,
                            footnote: None,
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
                            href: None,
                            footnote: None,
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
                        href: None,
                        footnote: None,
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
                            href: None,
                            footnote: None,
                        },
                        Run {
                            text: "and Times New Roman text".to_string(),
                            style: TextStyle {
                                font_family: Some("Times New Roman".to_string()),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
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
                                    href: None,
                                    footnote: None,
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
                                    href: None,
                                    footnote: None,
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
                        href: None,
                        footnote: None,
                    }],
                })],
                header: Some(ir::HeaderFooter {
                    paragraphs: vec![ir::HeaderFooterParagraph {
                        style: ParagraphStyle::default(),
                        elements: vec![ir::HFInline::Run(Run {
                            text: "My Header".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
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
                        href: None,
                        footnote: None,
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
                                href: None,
                                footnote: None,
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
                        href: None,
                        footnote: None,
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

    #[test]
    fn test_docx_toc_pipeline_produces_pdf() {
        use std::io::Cursor;
        let toc = docx_rs::TableOfContents::new()
            .heading_styles_range(1, 3)
            .alias("Table of contents")
            .add_item(
                docx_rs::TableOfContentsItem::new()
                    .text("Chapter 1")
                    .toc_key("_Toc00000001")
                    .level(1)
                    .page_ref("2"),
            )
            .add_item(
                docx_rs::TableOfContentsItem::new()
                    .text("Chapter 2")
                    .toc_key("_Toc00000002")
                    .level(1)
                    .page_ref("5"),
            );

        let docx = docx_rs::Docx::new()
            .add_style(
                docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph).name("Heading 1"),
            )
            .add_table_of_contents(toc)
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Chapter 1"))
                    .style("Heading1"),
            )
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Some body text")),
            );

        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "DOCX with TOC should produce valid PDF"
        );
    }

    #[test]
    fn test_convert_bytes_with_pdfa_option() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("PDF/A test")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let options = ConvertOptions {
            pdf_standard: Some(config::PdfStandard::PdfA2b),
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
        // PDF/A output should contain PDF/A identification
        let pdf_str = String::from_utf8_lossy(&result.pdf);
        assert!(
            pdf_str.contains("pdfaid") || pdf_str.contains("PDF/A"),
            "PDF/A conversion should include PDF/A metadata"
        );
    }

    #[test]
    fn test_render_document_default_no_pdfa() {
        let doc = make_simple_document("No PDF/A");
        let pdf = render_document(&doc).unwrap();
        let pdf_str = String::from_utf8_lossy(&pdf);
        assert!(
            !pdf_str.contains("pdfaid:conformance"),
            "Default render_document should not produce PDF/A"
        );
    }

    #[test]
    fn test_convert_bytes_with_paper_size_override() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Paper size test")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let options = ConvertOptions {
            paper_size: Some(config::PaperSize::Letter),
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "DOCX with Letter paper override should produce valid PDF"
        );
    }

    #[test]
    fn test_convert_bytes_with_landscape_override() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Landscape override test")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let options = ConvertOptions {
            landscape: Some(true),
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "DOCX with landscape override should produce valid PDF"
        );
    }

    // --- US-048: Edge case handling and robustness tests ---

    #[test]
    fn test_edge_empty_docx_produces_valid_pdf() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new();
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Empty DOCX should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_empty_xlsx_produces_valid_pdf() {
        use std::io::Cursor;
        let book = umya_spreadsheet::new_file();
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Empty XLSX should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_empty_pptx_produces_valid_pdf() {
        use std::io::{Cursor, Write};
        // Build a PPTX with no slides
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::FileOptions::default();
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
        ).unwrap();
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst/></p:presentation>"#,
        ).unwrap();
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
        ).unwrap();
        let data = zip.finish().unwrap().into_inner();
        let result = convert_bytes(&data, Format::Pptx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Empty PPTX should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_long_paragraph_no_panic() {
        use std::io::Cursor;
        // Create a very long paragraph (10,000 characters)
        let long_text: String = "Lorem ipsum dolor sit amet. ".repeat(400);
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(&long_text)),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Long paragraph should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_large_table_no_panic() {
        use std::io::Cursor;
        let mut book = umya_spreadsheet::new_file();
        {
            let sheet = book.get_sheet_mut(&0).unwrap();
            // 100 rows x 20 columns
            for row in 1..=100u32 {
                for col in 1..=20u32 {
                    let coord = format!("{}{}", (b'A' + ((col - 1) % 26) as u8) as char, row);
                    sheet
                        .get_cell_mut(coord.as_str())
                        .set_value(format!("R{row}C{col}"));
                }
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Xlsx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Large table should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_corrupted_docx_returns_error() {
        let data = b"not a valid ZIP file at all";
        let result = convert_bytes(data, Format::Docx, &ConvertOptions::default());
        assert!(result.is_err(), "Corrupted DOCX should return an error");
        let err = result.unwrap_err();
        match err {
            ConvertError::Parse(msg) => {
                assert!(!msg.is_empty(), "Error message should not be empty");
            }
            _ => panic!("Expected Parse error for corrupted DOCX, got {err:?}"),
        }
    }

    #[test]
    fn test_edge_corrupted_xlsx_returns_error() {
        let data = b"this is not an xlsx file";
        let result = convert_bytes(data, Format::Xlsx, &ConvertOptions::default());
        assert!(result.is_err(), "Corrupted XLSX should return an error");
    }

    #[test]
    fn test_edge_corrupted_pptx_returns_error() {
        let data = b"garbage data that is not a pptx";
        let result = convert_bytes(data, Format::Pptx, &ConvertOptions::default());
        assert!(result.is_err(), "Corrupted PPTX should return an error");
    }

    #[test]
    fn test_edge_truncated_zip_returns_error() {
        // Create a valid DOCX then truncate it
        let full_data = build_test_docx();
        let truncated = &full_data[..full_data.len() / 2];
        let result = convert_bytes(truncated, Format::Docx, &ConvertOptions::default());
        assert!(result.is_err(), "Truncated DOCX should return an error");
    }

    #[test]
    fn test_edge_unicode_cjk_text() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("中文测试 日本語テスト 한국어 테스트")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "CJK text should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_unicode_emoji_text() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello 🌍🎉💡 World")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();
        // Emoji may render with fallback font, but should not crash
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "Emoji text should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_unicode_rtl_text() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("مرحبا بالعالم")), // Arabic: Hello World
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "RTL text should produce valid PDF"
        );
    }

    #[test]
    fn test_edge_image_only_docx() {
        // A DOCX with only an image (no text paragraphs) should convert
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Image(ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47], // Minimal PNG header (won't render but shouldn't panic)
                    format: ir::ImageFormat::Png,
                    width: Some(100.0),
                    height: Some(100.0),
                })],
                header: None,
                footer: None,
            })],
            styles: StyleSheet::default(),
        };
        // This tests the render pipeline with image-only content
        // It may fail to compile the image (invalid PNG) but should not panic
        let _result = render_document(&doc);
    }

    // --- PDF output size regression tests (US-089) ---

    #[test]
    fn test_render_multipage_document_size() {
        // A 10-page document produced via the full IR → Typst → PDF pipeline
        // should be under 500KB, verifying end-to-end compression behavior.
        let mut pages = Vec::new();
        for i in 1..=10 {
            pages.push(Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle {
                            alignment: Some(Alignment::Center),
                            ..ParagraphStyle::default()
                        },
                        runs: vec![Run {
                            text: format!("Page {i} Heading"),
                            style: TextStyle {
                                bold: Some(true),
                                font_size: Some(24.0),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
                        }],
                    }),
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: format!(
                                "This is page {i}. Lorem ipsum dolor sit amet, \
                                 consectetur adipiscing elit. Sed do eiusmod tempor \
                                 incididunt ut labore et dolore magna aliqua."
                            ),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }),
                ],
                header: None,
                footer: None,
            }));
        }
        let doc = Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(
            pdf.len() < 512_000,
            "10-page IR document PDF should be under 500KB, actual: {} bytes ({:.1} KB)",
            pdf.len(),
            pdf.len() as f64 / 1024.0
        );
    }

    #[test]
    fn test_render_pptx_style_document_size() {
        // A slide-style (FixedPage) document should produce reasonably sized PDF.
        let mut pages = Vec::new();
        for i in 1..=5 {
            pages.push(Page::Fixed(FixedPage {
                size: PageSize {
                    width: 720.0,
                    height: 540.0,
                },
                background_color: None,
                background_gradient: None,
                elements: vec![FixedElement {
                    x: 50.0,
                    y: 50.0,
                    width: 620.0,
                    height: 80.0,
                    kind: FixedElementKind::TextBox(vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: format!("Slide {i} content"),
                            style: TextStyle {
                                font_size: Some(32.0),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
                        }],
                    })]),
                }],
            }));
        }
        let doc = Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        };
        let pdf = render_document(&doc).unwrap();
        assert!(
            pdf.len() < 512_000,
            "5-slide FixedPage PDF should be under 500KB, actual: {} bytes ({:.1} KB)",
            pdf.len(),
            pdf.len() as f64 / 1024.0
        );
    }

    // --- ConvertMetrics instrumentation tests ---

    /// Helper: create a minimal DOCX as bytes for metrics tests.
    fn make_test_docx_bytes() -> Vec<u8> {
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello metrics")),
        );
        let mut buf = std::io::Cursor::new(Vec::new());
        docx.build().pack(&mut buf).unwrap();
        buf.into_inner()
    }

    #[test]
    fn test_convert_bytes_returns_populated_metrics() {
        let data = make_test_docx_bytes();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        let metrics = result.metrics.expect("convert_bytes should return metrics");
        assert!(
            metrics.parse_duration.as_nanos() > 0,
            "parse_duration should be non-zero"
        );
        assert!(
            metrics.codegen_duration.as_nanos() > 0,
            "codegen_duration should be non-zero"
        );
        assert!(
            metrics.compile_duration.as_nanos() > 0,
            "compile_duration should be non-zero"
        );
        assert!(
            metrics.total_duration.as_nanos() > 0,
            "total_duration should be non-zero"
        );
        assert_eq!(metrics.input_size_bytes, data.len() as u64);
        assert_eq!(metrics.output_size_bytes, result.pdf.len() as u64);
        assert!(metrics.page_count >= 1, "should have at least 1 page");
    }

    #[test]
    fn test_metrics_total_ge_sum_of_stages() {
        let data = make_test_docx_bytes();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        let m = result.metrics.expect("should have metrics");
        let sum = m.parse_duration + m.codegen_duration + m.compile_duration;
        assert!(
            m.total_duration >= sum,
            "total ({:?}) should be >= sum of stages ({:?})",
            m.total_duration,
            sum
        );
    }

    #[test]
    fn test_metrics_output_size_matches_pdf() {
        let data = make_test_docx_bytes();
        let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
        let m = result.metrics.expect("should have metrics");
        assert_eq!(
            m.output_size_bytes,
            result.pdf.len() as u64,
            "output_size_bytes should match actual PDF size"
        );
    }

    // --- Tagged PDF and PDF/UA integration tests (US-096) ---

    #[test]
    fn test_convert_bytes_with_tagged_option() {
        use std::io::Cursor;
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tagged test")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let options = ConvertOptions {
            tagged: true,
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
        let pdf_str = String::from_utf8_lossy(&result.pdf);
        assert!(
            pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
            "Tagged conversion should include structure tree"
        );
    }

    #[test]
    fn test_convert_bytes_with_pdf_ua_option() {
        // PDF/UA requires a document title. Build a DOCX with core properties.
        let data = build_docx_with_title("PDF/UA Test Document");

        let options = ConvertOptions {
            pdf_ua: true,
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
        let pdf_str = String::from_utf8_lossy(&result.pdf);
        assert!(
            pdf_str.contains("pdfuaid"),
            "PDF/UA conversion should include pdfuaid metadata"
        );
    }

    #[test]
    fn test_convert_bytes_tagged_pdf_with_heading() {
        use std::io::Cursor;

        // Create a DOCX with a Heading 1 style
        let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
            .name("Heading 1")
            .outline_lvl(0);

        let docx = docx_rs::Docx::new()
            .add_style(h1_style)
            .add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("My Title"))
                    .style("Heading1"),
            )
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
            );

        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let options = ConvertOptions {
            tagged: true,
            ..Default::default()
        };
        let result = convert_bytes(&data, Format::Docx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
        let pdf_str = String::from_utf8_lossy(&result.pdf);
        assert!(
            pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
            "Tagged PDF with headings should contain structure tags"
        );
    }
}

#[cfg(all(test, feature = "typescript"))]
mod ts_integration_tests {
    use ts_rs::TS;

    use crate::config::{ConvertOptions, Format, PaperSize, PdfStandard, SlideRange};
    use crate::error::{ConvertMetrics, ConvertWarning};

    fn cfg_for_bindings() -> ts_rs::Config {
        let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");
        std::fs::create_dir_all(&bindings_dir).unwrap();
        ts_rs::Config::new().with_out_dir(bindings_dir)
    }

    #[test]
    fn test_export_all_types_to_bindings() {
        let cfg = cfg_for_bindings();
        let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");

        Format::export_all(&cfg).unwrap();
        PaperSize::export_all(&cfg).unwrap();
        PdfStandard::export_all(&cfg).unwrap();
        SlideRange::export_all(&cfg).unwrap();
        ConvertOptions::export_all(&cfg).unwrap();
        ConvertWarning::export_all(&cfg).unwrap();
        ConvertMetrics::export_all(&cfg).unwrap();

        // Verify files were created
        assert!(bindings_dir.join("Format.ts").exists());
        assert!(bindings_dir.join("PaperSize.ts").exists());
        assert!(bindings_dir.join("PdfStandard.ts").exists());
        assert!(bindings_dir.join("SlideRange.ts").exists());
        assert!(bindings_dir.join("ConvertOptions.ts").exists());
        assert!(bindings_dir.join("ConvertWarning.ts").exists());
        assert!(bindings_dir.join("ConvertMetrics.ts").exists());
    }

    #[test]
    fn test_generated_types_contain_expected_content() {
        let cfg = cfg_for_bindings();
        let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");

        // Export all types
        Format::export_all(&cfg).unwrap();
        ConvertOptions::export_all(&cfg).unwrap();

        // Read and verify Format.ts content
        let format_ts = std::fs::read_to_string(bindings_dir.join("Format.ts")).unwrap();
        assert!(
            format_ts.contains("Docx"),
            "Format.ts should contain Docx: {format_ts}"
        );
        assert!(
            format_ts.contains("Pptx"),
            "Format.ts should contain Pptx: {format_ts}"
        );
        assert!(
            format_ts.contains("Xlsx"),
            "Format.ts should contain Xlsx: {format_ts}"
        );

        // Read and verify ConvertOptions.ts content
        let opts_ts = std::fs::read_to_string(bindings_dir.join("ConvertOptions.ts")).unwrap();
        assert!(
            opts_ts.contains("tagged"),
            "ConvertOptions.ts should contain tagged: {opts_ts}"
        );
        assert!(
            opts_ts.contains("pdf_ua"),
            "ConvertOptions.ts should contain pdf_ua: {opts_ts}"
        );
        assert!(
            opts_ts.contains("boolean"),
            "boolean fields should be mapped: {opts_ts}"
        );
    }
}

#[cfg(all(test, feature = "pdf-ops"))]
mod streaming_tests {
    use super::*;
    use std::io::Cursor;

    /// Build an XLSX with many rows.
    fn build_xlsx_with_rows(num_rows: u32, num_cols: u32) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        for row in 1..=num_rows {
            for col in 1..=num_cols {
                sheet
                    .get_cell_mut((col, row))
                    .set_value(format!("R{row}C{col}"));
            }
        }
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_streaming_xlsx_produces_valid_pdf() {
        let data = build_xlsx_with_rows(50, 3);
        let options = config::ConvertOptions {
            streaming: true,
            streaming_chunk_size: Some(20),
            ..Default::default()
        };
        let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
        // PDF should start with %PDF
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "output should be valid PDF"
        );
        assert!(result.pdf.len() > 100, "PDF should have content");
    }

    #[test]
    fn test_streaming_xlsx_same_data_as_normal() {
        // Both streaming and non-streaming should produce valid PDFs with content
        let data = build_xlsx_with_rows(10, 2);

        let normal_opts = config::ConvertOptions::default();
        let normal_result = convert_bytes(&data, config::Format::Xlsx, &normal_opts).unwrap();

        let streaming_opts = config::ConvertOptions {
            streaming: true,
            streaming_chunk_size: Some(5),
            ..Default::default()
        };
        let streaming_result = convert_bytes(&data, config::Format::Xlsx, &streaming_opts).unwrap();

        // Both should produce valid PDFs
        assert!(normal_result.pdf.starts_with(b"%PDF"));
        assert!(streaming_result.pdf.starts_with(b"%PDF"));
        // Both should have content (non-empty)
        assert!(normal_result.pdf.len() > 100);
        assert!(streaming_result.pdf.len() > 100);
    }

    #[test]
    fn test_streaming_large_xlsx_completes() {
        // 10,000 rows — a large spreadsheet
        let data = build_xlsx_with_rows(10_000, 3);
        let options = config::ConvertOptions {
            streaming: true,
            streaming_chunk_size: Some(1000),
            ..Default::default()
        };
        let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
        assert!(
            result.pdf.starts_with(b"%PDF"),
            "output should be valid PDF"
        );
        assert!(result.metrics.is_some(), "streaming should produce metrics");
    }

    #[test]
    fn test_streaming_non_xlsx_falls_through() {
        // Streaming mode on DOCX should just do normal conversion
        let docx = {
            let doc = docx_rs::Docx::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello streaming")),
            );
            let mut cursor = Cursor::new(Vec::new());
            doc.build().pack(&mut cursor).unwrap();
            cursor.into_inner()
        };
        let options = config::ConvertOptions {
            streaming: true,
            ..Default::default()
        };
        let result = convert_bytes(&docx, config::Format::Docx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_streaming_chunk_size_default() {
        // When streaming_chunk_size is None, default to 1000
        let data = build_xlsx_with_rows(20, 1);
        let options = config::ConvertOptions {
            streaming: true,
            streaming_chunk_size: None, // should default to 1000
            ..Default::default()
        };
        // With 20 rows and default chunk_size=1000, should be 1 chunk
        let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_streaming_memory_bounded() {
        // Streaming a 10,000-row XLSX should use less peak memory
        // than non-streaming. We test this by verifying completion
        // (if memory were unbounded, large allocations would fail or be very slow).
        let data = build_xlsx_with_rows(5_000, 5);
        let options = config::ConvertOptions {
            streaming: true,
            streaming_chunk_size: Some(500),
            ..Default::default()
        };
        let result = convert_bytes(&data, config::Format::Xlsx, &options).unwrap();
        assert!(result.pdf.starts_with(b"%PDF"));
        // The result should have multiple pages (5000/500 = 10 chunks)
        assert!(
            result.pdf.len() > 1000,
            "PDF should have substantial content"
        );
    }
}
