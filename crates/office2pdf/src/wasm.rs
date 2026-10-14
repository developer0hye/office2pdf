//! WebAssembly bindings for office2pdf via `wasm-bindgen`.
//!
//! This module is only available when the `wasm` feature is enabled.
//! It exports JavaScript-callable functions for converting Office documents
//! to PDF in browser or Node.js environments.
//!
//! # Running WASM integration tests
//!
//! WASM integration tests use `wasm-bindgen-test` and require `wasm-pack`:
//!
//! ```bash
//! # Install wasm-pack (one-time setup)
//! cargo install wasm-pack
//!
//! # Run WASM tests in Node.js
//! cd crates/office2pdf
//! wasm-pack test --node --features wasm
//!
//! # Or run in a headless browser
//! wasm-pack test --headless --chrome --features wasm
//! ```
//!
//! These tests verify end-to-end WASM conversion by building the library as
//! a WASM module, loading it, and calling the exported functions.

use wasm_bindgen::prelude::*;

use crate::config::{ConvertOptions, Format};
use crate::convert_bytes;
use crate::error::{ConvertResult as CoreConvertResult, ConvertWarning};

/// Structured warning returned by the result-bearing WASM API.
#[wasm_bindgen]
#[derive(Clone)]
pub struct ConversionWarning {
    kind: String,
    format: String,
    message: String,
    from: Option<String>,
    to: Option<String>,
    element: Option<String>,
    detail: Option<String>,
    reason: Option<String>,
}

#[wasm_bindgen]
impl ConversionWarning {
    #[wasm_bindgen(getter)]
    pub fn kind(&self) -> String {
        self.kind.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn format(&self) -> String {
        self.format.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn message(&self) -> String {
        self.message.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn from(&self) -> Option<String> {
        self.from.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn to(&self) -> Option<String> {
        self.to.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn element(&self) -> Option<String> {
        self.element.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn detail(&self) -> Option<String> {
        self.detail.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn reason(&self) -> Option<String> {
        self.reason.clone()
    }
}

impl From<ConvertWarning> for ConversionWarning {
    fn from(warning: ConvertWarning) -> Self {
        let message = warning.to_string();
        match warning {
            ConvertWarning::UnsupportedElement { format, element } => Self {
                kind: "unsupported-element".to_string(),
                format,
                message,
                from: None,
                to: None,
                element: Some(element),
                detail: None,
                reason: None,
            },
            ConvertWarning::PartialElement {
                format,
                element,
                detail,
            } => Self {
                kind: "partial-element".to_string(),
                format,
                message,
                from: None,
                to: None,
                element: Some(element),
                detail: Some(detail),
                reason: None,
            },
            ConvertWarning::FallbackUsed { format, from, to } => Self {
                kind: "fallback-used".to_string(),
                format,
                message,
                from: Some(from),
                to: Some(to),
                element: None,
                detail: None,
                reason: None,
            },
            ConvertWarning::ParseSkipped { format, reason } => Self {
                kind: "parse-skipped".to_string(),
                format,
                message,
                from: None,
                to: None,
                element: None,
                detail: None,
                reason: Some(reason),
            },
        }
    }
}

/// PDF bytes and non-fatal warnings from a WASM conversion.
#[wasm_bindgen]
pub struct ConversionResult {
    pdf: Vec<u8>,
    warnings: Vec<ConversionWarning>,
}

#[wasm_bindgen]
impl ConversionResult {
    #[wasm_bindgen(getter)]
    pub fn pdf(&self) -> Vec<u8> {
        self.pdf.clone()
    }

    #[wasm_bindgen(getter, js_name = warningCount)]
    pub fn warning_count(&self) -> usize {
        self.warnings.len()
    }

    #[wasm_bindgen(js_name = warningAt)]
    pub fn warning_at(&self, index: usize) -> Option<ConversionWarning> {
        self.warnings.get(index).cloned()
    }
}

impl From<CoreConvertResult> for ConversionResult {
    fn from(result: CoreConvertResult) -> Self {
        Self {
            pdf: result.pdf,
            warnings: result.warnings.into_iter().map(Into::into).collect(),
        }
    }
}

/// Per-instance WASM converter with caller-provided in-memory fonts.
#[wasm_bindgen]
#[derive(Default)]
pub struct Office2PdfConverter {
    options: ConvertOptions,
}

#[wasm_bindgen]
impl Office2PdfConverter {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self::default()
    }

    /// Register a standalone font or font collection for later conversions.
    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(&mut self, data: &[u8]) -> Result<(), JsValue> {
        if crate::render::pdf::load_fonts_from_bytes([data]).is_empty() {
            return Err(JsValue::from_str(
                "registered font contains no usable font faces",
            ));
        }
        self.options.font_bytes.push(data.to_vec());
        Ok(())
    }

    /// Remove all caller-registered font buffers from this converter.
    #[wasm_bindgen(js_name = clearFonts)]
    pub fn clear_fonts(&mut self) {
        self.options.font_bytes.clear();
    }

    /// Set the family appended to every generated font chain.
    #[wasm_bindgen(js_name = setLastResortFontFamily)]
    pub fn set_last_resort_font_family(&mut self, family: &str) -> Result<(), JsValue> {
        let family = family.trim();
        if family.is_empty() {
            return Err(JsValue::from_str(
                "last-resort font family must not be empty",
            ));
        }
        self.options.last_resort_font_family = Some(family.to_string());
        Ok(())
    }

    /// Remove the configured last-resort family.
    #[wasm_bindgen(js_name = clearLastResortFontFamily)]
    pub fn clear_last_resort_font_family(&mut self) {
        self.options.last_resort_font_family = None;
    }

    #[wasm_bindgen(js_name = convertToPdf)]
    pub fn convert_to_pdf(&self, data: &[u8], format: &str) -> Result<ConversionResult, JsValue> {
        let format = Format::from_extension(format)
            .ok_or_else(|| JsValue::from_str(&format!("unsupported format: {format}")))?;
        self.convert_format(data, format)
    }

    #[wasm_bindgen(js_name = convertDocxToPdf)]
    pub fn convert_docx_to_pdf(&self, data: &[u8]) -> Result<ConversionResult, JsValue> {
        self.convert_format(data, Format::Docx)
    }

    #[wasm_bindgen(js_name = convertPptxToPdf)]
    pub fn convert_pptx_to_pdf(&self, data: &[u8]) -> Result<ConversionResult, JsValue> {
        self.convert_format(data, Format::Pptx)
    }

    #[wasm_bindgen(js_name = convertXlsxToPdf)]
    pub fn convert_xlsx_to_pdf(&self, data: &[u8]) -> Result<ConversionResult, JsValue> {
        self.convert_format(data, Format::Xlsx)
    }
}

impl Office2PdfConverter {
    fn convert_format(&self, data: &[u8], format: Format) -> Result<ConversionResult, JsValue> {
        convert_bytes(data, format, &self.options)
            .map(Into::into)
            .map_err(|error| JsValue::from_str(&error.to_string()))
    }
}

/// Internal: convert with format string, returning a `String` error (testable on native).
fn convert_to_pdf_inner(data: &[u8], format: &str) -> Result<Vec<u8>, String> {
    let fmt =
        Format::from_extension(format).ok_or_else(|| format!("unsupported format: {format}"))?;
    let result = convert_bytes(data, fmt, &ConvertOptions::default()).map_err(|e| e.to_string())?;
    Ok(result.pdf)
}

/// Internal: convert with a known `Format`, returning a `String` error (testable on native).
fn convert_format_inner(data: &[u8], format: Format) -> Result<Vec<u8>, String> {
    let result =
        convert_bytes(data, format, &ConvertOptions::default()).map_err(|e| e.to_string())?;
    Ok(result.pdf)
}

/// Convert with default options while preserving structured warnings.
fn convert_format_with_result_inner(
    data: &[u8],
    format: Format,
) -> Result<ConversionResult, String> {
    convert_bytes(data, format, &ConvertOptions::default())
        .map(Into::into)
        .map_err(|error| error.to_string())
}

/// Convert an Office document to PDF.
///
/// `data` is the raw bytes of the input document (DOCX, PPTX, or XLSX).
/// `format` is one of `"docx"`, `"pptx"`, or `"xlsx"` (case-insensitive).
///
/// Returns the PDF bytes on success, or throws a JS error string on failure.
#[wasm_bindgen(js_name = "convertToPdf")]
pub fn convert_to_pdf(data: &[u8], format: &str) -> Result<Vec<u8>, JsValue> {
    convert_to_pdf_inner(data, format).map_err(|e| JsValue::from_str(&e))
}

/// Convert a DOCX document to PDF.
///
/// `data` is the raw bytes of a `.docx` file.
///
/// Returns the PDF bytes on success, or throws a JS error string on failure.
#[wasm_bindgen(js_name = "convertDocxToPdf")]
pub fn convert_docx_to_pdf(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    convert_format_inner(data, Format::Docx).map_err(|e| JsValue::from_str(&e))
}

/// Convert a PPTX document to PDF.
///
/// `data` is the raw bytes of a `.pptx` file.
///
/// Returns the PDF bytes on success, or throws a JS error string on failure.
#[wasm_bindgen(js_name = "convertPptxToPdf")]
pub fn convert_pptx_to_pdf(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    convert_format_inner(data, Format::Pptx).map_err(|e| JsValue::from_str(&e))
}

/// Convert an XLSX document to PDF.
///
/// `data` is the raw bytes of a `.xlsx` file.
///
/// Returns the PDF bytes on success, or throws a JS error string on failure.
#[wasm_bindgen(js_name = "convertXlsxToPdf")]
pub fn convert_xlsx_to_pdf(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    convert_format_inner(data, Format::Xlsx).map_err(|e| JsValue::from_str(&e))
}

/// Convert a document with default options and return PDF bytes plus warnings.
#[wasm_bindgen(js_name = "convertToPdfWithResult")]
pub fn convert_to_pdf_with_result(data: &[u8], format: &str) -> Result<ConversionResult, JsValue> {
    let format = Format::from_extension(format)
        .ok_or_else(|| JsValue::from_str(&format!("unsupported format: {format}")))?;
    convert_format_with_result_inner(data, format).map_err(|error| JsValue::from_str(&error))
}

/// Convert a DOCX with default options and return PDF bytes plus warnings.
#[wasm_bindgen(js_name = "convertDocxToPdfWithResult")]
pub fn convert_docx_to_pdf_with_result(data: &[u8]) -> Result<ConversionResult, JsValue> {
    convert_format_with_result_inner(data, Format::Docx).map_err(|error| JsValue::from_str(&error))
}

/// Convert a PPTX with default options and return PDF bytes plus warnings.
#[wasm_bindgen(js_name = "convertPptxToPdfWithResult")]
pub fn convert_pptx_to_pdf_with_result(data: &[u8]) -> Result<ConversionResult, JsValue> {
    convert_format_with_result_inner(data, Format::Pptx).map_err(|error| JsValue::from_str(&error))
}

/// Convert an XLSX with default options and return PDF bytes plus warnings.
#[wasm_bindgen(js_name = "convertXlsxToPdfWithResult")]
pub fn convert_xlsx_to_pdf_with_result(data: &[u8]) -> Result<ConversionResult, JsValue> {
    convert_format_with_result_inner(data, Format::Xlsx).map_err(|error| JsValue::from_str(&error))
}

#[cfg(test)]
#[path = "wasm_tests.rs"]
mod tests;

// ---------------------------------------------------------------------------
// WASM integration tests (run via `wasm-pack test --node --features wasm`)
//
// These tests compile ONLY when targeting wasm32 and are executed inside a
// real WASM runtime (Node.js or headless browser). They call the actual
// `#[wasm_bindgen]`-exported functions and verify end-to-end conversion.
// ---------------------------------------------------------------------------
#[cfg(all(target_arch = "wasm32", test))]
mod wasm_tests {
    use super::*;
    use wasm_bindgen_test::*;

    /// Helper: create a minimal valid DOCX via docx-rs builder.
    fn make_minimal_docx() -> Vec<u8> {
        use std::io::Cursor;
        let doc = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello WASM")),
        );
        let mut buf = Cursor::new(Vec::new());
        doc.build().pack(&mut buf).unwrap();
        buf.into_inner()
    }

    /// DOCX that declares SimSun but deliberately embeds no font.
    fn make_cjk_docx_without_embedded_font() -> Vec<u8> {
        include_bytes!("../../../tests/fixtures/docx/wasm_registered_cjk.docx").to_vec()
    }

    fn noto_sans_sc_subset_bytes() -> Vec<u8> {
        let carrier = include_bytes!("../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
        let embedded =
            crate::parser::embedded_fonts::extract_embedded_font_data(carrier, Format::Docx)
                .expect("the #943 fixture should carry one embedded font");
        embedded
            .font_bytes()
            .next()
            .expect("the #943 fixture should expose one deobfuscated font")
            .to_vec()
    }

    /// Helper: create a minimal valid PPTX.
    fn make_minimal_pptx() -> Vec<u8> {
        use std::io::{Cursor, Write};
        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options =
            zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#)
        .unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#)
        .unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>"#,
        )
        .unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#)
        .unwrap();

        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1000000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Hello WASM</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#,
        )
        .unwrap();

        zip.finish().unwrap().into_inner()
    }

    /// Helper: create a minimal valid XLSX.
    fn make_minimal_xlsx() -> Vec<u8> {
        use std::io::Cursor;
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Hello WASM");
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    #[wasm_bindgen_test]
    fn wasm_convert_docx_to_pdf_produces_valid_pdf() {
        let docx = make_minimal_docx();
        let result = convert_docx_to_pdf(&docx);
        assert!(result.is_ok(), "DOCX to PDF conversion failed in WASM");
        let pdf = result.unwrap();
        assert!(
            pdf.starts_with(b"%PDF"),
            "Output should start with %PDF magic bytes"
        );
        assert!(pdf.len() > 100, "PDF output should have meaningful size");
    }

    #[wasm_bindgen_test]
    fn wasm_docx_table_style_line_spacing_reaches_pdf() {
        use crate::parser::Parser;

        let data =
            include_bytes!("../../../tests/fixtures/docx/docxide/table_style_line_spacing.docx");
        let (document, _) = crate::parser::docx::DocxParser
            .parse(data, &ConvertOptions::default())
            .expect("styled tables parse in WASM");
        let crate::ir::Page::Flow(page) = &document.pages[0] else {
            panic!("expected flow page")
        };
        let table = page
            .content
            .iter()
            .find_map(|block| match block {
                crate::ir::Block::Table(table) => Some(table),
                _ => None,
            })
            .expect("fixture contains a table");
        let paragraph = table.rows[0].cells[0]
            .content
            .iter()
            .find_map(|block| match block {
                crate::ir::Block::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .expect("first cell contains text");
        assert!(matches!(paragraph.style.line_spacing,
            Some(crate::ir::LineSpacing::Proportional(value)) if (value - 1.0).abs() < 1e-9));
        let pdf = convert_docx_to_pdf(data).expect("styled tables convert through the WASM API");
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[wasm_bindgen_test]
    fn wasm_convert_docx_uses_document_embedded_cjk_font() {
        let docx = include_bytes!("../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
        let pdf = convert_docx_to_pdf(docx)
            .expect("DOCX with an embedded CJK face should convert in WASM");

        assert!(
            pdf.windows(b"NotoSansSC".len())
                .any(|window| window == b"NotoSansSC"),
            "the PDF should embed the document-provided Noto Sans SC subset"
        );
    }

    #[wasm_bindgen_test]
    fn wasm_converter_registers_last_resort_font_and_surfaces_warning() {
        let docx = make_cjk_docx_without_embedded_font();
        let mut converter = Office2PdfConverter::new();
        converter
            .register_font(&noto_sans_sc_subset_bytes())
            .expect("the valid Noto subset should register");
        converter
            .set_last_resort_font_family("Noto Sans SC")
            .expect("a non-empty family should be accepted");

        let result = converter
            .convert_docx_to_pdf(&docx)
            .expect("registered CJK fallback should convert");

        assert!(
            result
                .pdf
                .windows(b"NotoSansSC".len())
                .any(|window| window == b"NotoSansSC"),
            "the PDF should embed the caller-registered Noto Sans SC subset"
        );
        assert!(result.warnings.iter().any(|warning| {
            warning.kind == "fallback-used"
                && warning.from.as_deref() == Some("SimSun")
                && warning.to.as_deref() == Some("Noto Sans SC")
        }));
    }

    #[cfg(not(feature = "wasm-cjk-font"))]
    #[wasm_bindgen_test]
    fn wasm_result_reports_notdef_when_no_cjk_face_is_available() {
        let result = convert_docx_to_pdf_with_result(&make_cjk_docx_without_embedded_font())
            .expect("conversion should finish even when glyph coverage is missing");

        assert!(result.warnings.iter().any(|warning| {
            warning.kind == "fallback-used"
                && warning.from.as_deref() == Some("SimSun")
                && warning.to.as_deref() == Some(".notdef")
        }));
    }

    #[cfg(feature = "wasm-cjk-font")]
    #[wasm_bindgen_test]
    fn wasm_bundled_cjk_feature_renders_chinese_without_caller_font() {
        let result = convert_docx_to_pdf_with_result(&make_cjk_docx_without_embedded_font())
            .expect("the bundled Chinese face should make default conversion readable");

        assert_eq!(result.warnings.len(), 1);
        let warning = &result.warnings[0];
        assert_eq!(warning.kind, "fallback-used");
        assert_eq!(warning.from.as_deref(), Some("SimSun"));
        assert_eq!(warning.to.as_deref(), Some("Noto Sans CJK SC"));
        assert!(
            result
                .pdf
                .windows(b"NotoSansCJKsc".len())
                .any(|window| window == b"NotoSansCJKsc"),
            "the PDF should embed the feature-gated bundled face"
        );
    }

    #[cfg(feature = "wasm-cjk-font")]
    #[wasm_bindgen_test]
    fn wasm_bundled_cjk_feature_applies_to_render_document() {
        use crate::parser::Parser;

        let (document, _) = crate::parser::docx::DocxParser
            .parse(
                &make_cjk_docx_without_embedded_font(),
                &ConvertOptions::default(),
            )
            .expect("fixture should parse to IR");
        let pdf = crate::render_document(&document)
            .expect("direct IR rendering should use the feature-gated face");

        assert!(
            pdf.windows(b"NotoSansCJKsc".len())
                .any(|window| window == b"NotoSansCJKsc"),
            "direct IR output should embed the feature-gated bundled face"
        );
    }

    #[wasm_bindgen_test]
    fn wasm_converter_rejects_invalid_font_and_empty_family() {
        let mut converter = Office2PdfConverter::new();
        assert!(converter.register_font(b"not a font").is_err());
        assert!(converter.set_last_resort_font_family("   ").is_err());
        assert!(converter.options.font_bytes.is_empty());
        assert!(converter.options.last_resort_font_family.is_none());
    }

    #[wasm_bindgen_test]
    fn wasm_convert_to_pdf_with_docx_format_string() {
        let docx = make_minimal_docx();
        let result = convert_to_pdf(&docx, "docx");
        assert!(
            result.is_ok(),
            "convert_to_pdf with 'docx' format failed in WASM"
        );
        let pdf = result.unwrap();
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[wasm_bindgen_test]
    fn wasm_docx_source_crops_follow_shared_picture_occurrences() {
        use crate::parser::Parser;

        let data: &[u8] =
            include_bytes!("../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
        let (document, _) = crate::parser::docx::DocxParser
            .parse(data, &ConvertOptions::default())
            .expect("shared-picture fixture should parse in WASM");
        let crate::ir::Page::Flow(page) = &document.pages[0] else {
            panic!("flow page");
        };
        let images: Vec<&crate::ir::ImageData> = page
            .content
            .iter()
            .filter_map(|block| match block {
                crate::ir::Block::FloatingImage(image) => Some(&image.image),
                _ => None,
            })
            .collect();
        assert_eq!(images.len(), 4);
        assert!(images[0].crop.is_none());
        assert!(images[1].crop.is_none());
        assert!(images[3].crop.is_none());
        let crop = images[2].crop.expect("third occurrence has its own crop");
        assert_eq!(
            (crop.left, crop.top, crop.right, crop.bottom),
            (0.25, 0.0, 0.25, 0.0)
        );
        let pdf = convert_docx_to_pdf(data).expect("cropped pictures should convert in WASM");
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[wasm_bindgen_test]
    fn wasm_docx_cached_smartart_reaches_pdf_conversion() {
        use crate::parser::Parser;

        let data: &[u8] =
            include_bytes!("../../../tests/fixtures/docx/docxide/smartart_cached_drawing.docx");
        let (document, _) = crate::parser::docx::DocxParser
            .parse(data, &ConvertOptions::default())
            .unwrap();
        let crate::ir::Page::Flow(page) = &document.pages[0] else {
            panic!("flow page")
        };
        let diagrams: usize = page
            .content
            .iter()
            .filter_map(|block| {
                if let crate::ir::Block::Paragraph(paragraph) = block {
                    Some(paragraph)
                } else {
                    None
                }
            })
            .flat_map(|paragraph| &paragraph.runs)
            .filter_map(|run| run.inline_box.as_ref())
            .flat_map(|frame| &frame.content)
            .filter(|block| matches!(block, crate::ir::Block::DrawingGroup(_)))
            .count();
        assert_eq!(diagrams, 2, "both cached diagrams reach the renderer");
        let pdf = convert_docx_to_pdf(data)
            .expect("cached diagrams convert through the exported WASM API");
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[wasm_bindgen_test]
    fn wasm_convert_pptx_to_pdf_produces_valid_pdf() {
        let pptx = make_minimal_pptx();
        let result = convert_pptx_to_pdf(&pptx);
        assert!(result.is_ok(), "PPTX to PDF conversion failed in WASM");
        let pdf = result.unwrap();
        assert!(
            pdf.starts_with(b"%PDF"),
            "Output should start with %PDF magic bytes"
        );
    }

    #[wasm_bindgen_test]
    fn wasm_convert_xlsx_to_pdf_produces_valid_pdf() {
        let xlsx = make_minimal_xlsx();
        let result = convert_xlsx_to_pdf(&xlsx);
        assert!(result.is_ok(), "XLSX to PDF conversion failed in WASM");
        let pdf = result.unwrap();
        assert!(
            pdf.starts_with(b"%PDF"),
            "Output should start with %PDF magic bytes"
        );
    }

    #[wasm_bindgen_test]
    fn wasm_convert_to_pdf_invalid_data_returns_error() {
        let result = convert_docx_to_pdf(b"not a valid docx");
        assert!(result.is_err(), "Should fail on invalid input data");
    }

    #[wasm_bindgen_test]
    fn wasm_convert_to_pdf_unsupported_format_returns_error() {
        let result = convert_to_pdf(b"dummy", "txt");
        assert!(result.is_err(), "Should fail on unsupported format string");
    }
}
