use typst::diag::FileResult;
use typst::foundations::{Bytes, Datetime};
use typst::syntax::{FileId, Source, VirtualPath};
use typst::text::Font;
use typst::utils::LazyHash;
use typst::{Library, LibraryExt, World};
use typst_kit::fonts::{FontSearcher, Fonts};

use crate::error::ConvertError;

/// Compile Typst markup to PDF bytes.
pub fn compile_to_pdf(typst_source: &str) -> Result<Vec<u8>, ConvertError> {
    let world = MinimalWorld::new(typst_source);

    let warned = typst::compile::<typst::layout::PagedDocument>(&world);
    let document = warned.output.map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("Typst compilation failed: {}", messages.join("; ")))
    })?;

    let options = typst_pdf::PdfOptions::default();
    typst_pdf::pdf(&document, &options).map_err(|errors| {
        let messages: Vec<String> = errors.iter().map(|e| e.message.to_string()).collect();
        ConvertError::Render(format!("PDF export failed: {}", messages.join("; ")))
    })
}

/// Minimal World implementation providing Typst compiler with source and fonts.
struct MinimalWorld {
    library: LazyHash<Library>,
    book: LazyHash<typst::text::FontBook>,
    fonts: Vec<typst_kit::fonts::FontSlot>,
    source: Source,
}

impl MinimalWorld {
    fn new(source_text: &str) -> Self {
        let mut searcher = FontSearcher::new();
        searcher.include_system_fonts(false);
        let font_data: Fonts = searcher.search();

        let main_id = FileId::new(None, VirtualPath::new("main.typ"));
        let source = Source::new(main_id, source_text.to_string());

        Self {
            library: LazyHash::new(Library::default()),
            book: LazyHash::new(font_data.book),
            fonts: font_data.fonts,
            source,
        }
    }
}

impl World for MinimalWorld {
    fn library(&self) -> &LazyHash<Library> {
        &self.library
    }

    fn book(&self) -> &LazyHash<typst::text::FontBook> {
        &self.book
    }

    fn main(&self) -> FileId {
        self.source.id()
    }

    fn source(&self, id: FileId) -> FileResult<Source> {
        if id == self.source.id() {
            Ok(self.source.clone())
        } else {
            Err(typst::diag::FileError::NotFound(
                id.vpath().as_rootless_path().into(),
            ))
        }
    }

    fn file(&self, id: FileId) -> FileResult<Bytes> {
        if id == self.source.id() {
            Ok(Bytes::new(self.source.text().as_bytes().to_vec()))
        } else {
            Err(typst::diag::FileError::NotFound(
                id.vpath().as_rootless_path().into(),
            ))
        }
    }

    fn font(&self, index: usize) -> Option<Font> {
        self.fonts.get(index).and_then(|slot| slot.get())
    }

    fn today(&self, _offset: Option<i64>) -> Option<Datetime> {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compile_simple_text() {
        let result = compile_to_pdf("Hello, World!").unwrap();
        assert!(!result.is_empty(), "PDF bytes should not be empty");
        assert!(
            result.starts_with(b"%PDF"),
            "PDF should start with %PDF magic bytes"
        );
    }

    #[test]
    fn test_compile_with_page_setup() {
        let source = r#"#set page(width: 612pt, height: 792pt)
Hello from a US Letter page."#;
        let result = compile_to_pdf(source).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_styled_text() {
        let source = r#"#text(weight: "bold", size: 16pt)[Bold Title]

#text(style: "italic")[Italic body text]

#underline[Underlined text]"#;
        let result = compile_to_pdf(source).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_colored_text() {
        let source = r#"#text(fill: rgb(255, 0, 0))[Red text]
#text(fill: rgb(0, 128, 255))[Blue text]"#;
        let result = compile_to_pdf(source).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_alignment() {
        let source = r#"#align(center)[Centered text]

#align(right)[Right-aligned text]"#;
        let result = compile_to_pdf(source).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_invalid_source_returns_error() {
        // Invalid Typst source should produce a compilation error
        let result = compile_to_pdf("#invalid-func-that-does-not-exist()");
        assert!(result.is_err(), "Invalid source should produce an error");
    }

    #[test]
    fn test_compile_empty_source() {
        // Empty source should still produce valid PDF (empty page)
        let result = compile_to_pdf("").unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }

    #[test]
    fn test_compile_multiple_paragraphs() {
        let source = "First paragraph.\n\nSecond paragraph.\n\nThird paragraph.";
        let result = compile_to_pdf(source).unwrap();
        assert!(!result.is_empty());
        assert!(result.starts_with(b"%PDF"));
    }
}
