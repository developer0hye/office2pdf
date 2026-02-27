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

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    /// Helper: create a minimal valid DOCX via docx-rs builder.
    fn make_minimal_docx() -> Vec<u8> {
        let doc = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello WASM")),
        );
        let mut buf = Cursor::new(Vec::new());
        doc.build().pack(&mut buf).unwrap();
        buf.into_inner()
    }

    /// Helper: create a minimal valid PPTX.
    fn make_minimal_pptx() -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
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
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value("Hello WASM");
        let mut cursor = Cursor::new(Vec::new());
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
        cursor.into_inner()
    }

    // --- Tests for convert_to_pdf_inner (generic format string API) ---

    #[test]
    fn test_convert_to_pdf_inner_docx() {
        let docx = make_minimal_docx();
        let result = convert_to_pdf_inner(&docx, "docx");
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_to_pdf_inner_pptx() {
        let pptx = make_minimal_pptx();
        let result = convert_to_pdf_inner(&pptx, "pptx");
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_to_pdf_inner_xlsx() {
        let xlsx = make_minimal_xlsx();
        let result = convert_to_pdf_inner(&xlsx, "xlsx");
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_to_pdf_inner_case_insensitive() {
        let docx = make_minimal_docx();
        assert!(convert_to_pdf_inner(&docx, "DOCX").is_ok());
        assert!(convert_to_pdf_inner(&docx, "Docx").is_ok());
    }

    #[test]
    fn test_convert_to_pdf_inner_unsupported_format() {
        let result = convert_to_pdf_inner(b"dummy", "txt");
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("unsupported format"));
    }

    #[test]
    fn test_convert_to_pdf_inner_invalid_data() {
        let result = convert_to_pdf_inner(b"not a docx", "docx");
        assert!(result.is_err());
    }

    // --- Tests for convert_format_inner (typed format API) ---

    #[test]
    fn test_convert_format_inner_docx() {
        let docx = make_minimal_docx();
        let result = convert_format_inner(&docx, Format::Docx);
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_format_inner_pptx() {
        let pptx = make_minimal_pptx();
        let result = convert_format_inner(&pptx, Format::Pptx);
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_format_inner_xlsx() {
        let xlsx = make_minimal_xlsx();
        let result = convert_format_inner(&xlsx, Format::Xlsx);
        assert!(result.is_ok(), "failed: {:?}", result.err());
        assert!(result.unwrap().starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_format_inner_docx_invalid() {
        assert!(convert_format_inner(b"bad", Format::Docx).is_err());
    }

    #[test]
    fn test_convert_format_inner_pptx_invalid() {
        assert!(convert_format_inner(b"bad", Format::Pptx).is_err());
    }

    #[test]
    fn test_convert_format_inner_xlsx_invalid() {
        assert!(convert_format_inner(b"bad", Format::Xlsx).is_err());
    }
}

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
