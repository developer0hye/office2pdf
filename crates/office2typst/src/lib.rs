pub use office2pdf::config;
pub use office2pdf::error;
pub use office2pdf::ir;

use std::collections::HashSet;

use office2pdf::config::{ConvertOptions, Format};
use office2pdf::error::{ConvertError, ConvertWarning};
use office2pdf::ir::Document;
use office2pdf::parser::{self, Parser};
use office2pdf::render::typst_gen::{self, TypstOutput};

#[derive(Debug)]
pub struct TypstResult {
    pub document: Document,
    pub typst: TypstOutput,
    pub warnings: Vec<ConvertWarning>,
}

const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn is_ole2(data: &[u8]) -> bool {
    data.len() >= OLE2_MAGIC.len() && data[..OLE2_MAGIC.len()] == OLE2_MAGIC
}

fn extract_panic_message(payload: &Box<dyn std::any::Any + Send>) -> String {
    if let Some(s) = payload.downcast_ref::<String>() {
        s.clone()
    } else if let Some(s) = payload.downcast_ref::<&str>() {
        (*s).to_string()
    } else {
        "unknown panic".to_string()
    }
}

fn dedup_warnings(warnings: &mut Vec<ConvertWarning>) {
    let mut seen: HashSet<String> = HashSet::new();
    warnings.retain(|warning| seen.insert(warning.to_string()));
}

fn parser_for_format(format: Format) -> Box<dyn Parser> {
    match format {
        Format::Docx => Box::new(parser::docx::DocxParser),
        Format::Pptx => Box::new(parser::pptx::PptxParser),
        Format::Xlsx => Box::new(parser::xlsx::XlsxParser),
    }
}

pub fn parse_bytes(
    data: &[u8],
    format: Format,
    options: &ConvertOptions,
) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
    if is_ole2(data) {
        return Err(ConvertError::UnsupportedEncryption);
    }

    let parser: Box<dyn Parser> = parser_for_format(format);
    let parse_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        parser.parse(data, options)
    }));

    let (document, mut warnings) = match parse_result {
        Ok(result) => result?,
        Err(panic_info) => {
            return Err(ConvertError::Parse(format!(
                "upstream parser panicked: {}",
                extract_panic_message(&panic_info)
            )))
        }
    };

    dedup_warnings(&mut warnings);
    Ok((document, warnings))
}

#[cfg(not(target_arch = "wasm32"))]
pub fn parse_path(
    path: impl AsRef<std::path::Path>,
    options: &ConvertOptions,
) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .ok_or_else(|| ConvertError::UnsupportedFormat("no file extension".to_string()))?;
    let format =
        Format::from_extension(ext).ok_or_else(|| ConvertError::UnsupportedFormat(ext.to_string()))?;
    let data: Vec<u8> = std::fs::read(path)?;
    parse_bytes(&data, format, options)
}

pub fn generate_typst(doc: &Document) -> Result<TypstOutput, ConvertError> {
    typst_gen::generate_typst(doc)
}

pub fn generate_typst_with_options(
    doc: &Document,
    options: &ConvertOptions,
) -> Result<TypstOutput, ConvertError> {
    typst_gen::generate_typst_with_options(doc, options)
}

pub fn convert_bytes(
    data: &[u8],
    format: Format,
    options: &ConvertOptions,
) -> Result<TypstResult, ConvertError> {
    let (document, warnings) = parse_bytes(data, format, options)?;
    let typst = generate_typst_with_options(&document, options)?;
    Ok(TypstResult {
        document,
        typst,
        warnings,
    })
}

#[cfg(not(target_arch = "wasm32"))]
pub fn convert_path(
    path: impl AsRef<std::path::Path>,
    options: &ConvertOptions,
) -> Result<TypstResult, ConvertError> {
    let (document, warnings) = parse_path(path, options)?;
    let typst = generate_typst_with_options(&document, options)?;
    Ok(TypstResult {
        document,
        typst,
        warnings,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_minimal_docx() -> Vec<u8> {
        let doc = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello Typst")),
        );
        let mut cursor = std::io::Cursor::new(Vec::new());
        doc.build()
            .pack(&mut cursor)
            .expect("DOCX pack should succeed");
        cursor.into_inner()
    }

    #[test]
    fn convert_bytes_docx_produces_typst_source() {
        let docx = make_minimal_docx();
        let result = convert_bytes(&docx, Format::Docx, &ConvertOptions::default())
            .expect("DOCX should convert to Typst");
        assert!(
            result.typst.source.contains("Hello Typst"),
            "Typst source should contain original text, got: {}",
            result.typst.source
        );
        assert!(
            !result.document.pages.is_empty(),
            "Parsed document should contain at least one page"
        );
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn parse_path_rejects_unsupported_extension() {
        let err = parse_path(
            "/tmp/not-supported.txt",
            &ConvertOptions {
                ..Default::default()
            },
        )
        .expect_err("Unsupported extension should return an error");
        match err {
            ConvertError::UnsupportedFormat(ext) => assert_eq!(ext, "txt"),
            other => panic!("Expected UnsupportedFormat, got {other:?}"),
        }
    }
}
