#![cfg(not(target_arch = "wasm32"))] // native-only unit tests (filesystem, system fonts)
use super::test_support::{build_test_docx, build_test_pptx, build_test_xlsx};
use super::*;

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

    let result2 = convert_with_options(&input, &ConvertOptions::default()).unwrap();
    assert!(!result2.pdf.is_empty());
    assert!(result2.pdf.starts_with(b"%PDF"));

    std::fs::write(&output, &result.pdf).unwrap();
    assert!(output.exists());
    let written = std::fs::read(&output).unwrap();
    assert!(written.starts_with(b"%PDF"));

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

#[test]
fn test_e2e_docx_with_list_produces_pdf() {
    use std::io::Cursor;

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

#[test]
fn test_e2e_landscape_docx_to_pdf() {
    use std::io::Cursor;

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
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape document")),
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
        .add_style(docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph).name("Heading 1"))
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

/// The PDF declares the language the document was written in, not English
/// whatever the source said.
#[test]
fn test_docx_language_reaches_pdf_catalog() {
    use std::io::{Cursor, Read, Write};

    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Aftale om levering")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    // docx-rs cannot state a language, so write Danish Word's into the
    // document defaults.
    let mut archive = zip::ZipArchive::new(Cursor::new(cursor.into_inner())).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name: String = entry.name().to_string();
        let mut content: Vec<u8> = Vec::new();
        entry.read_to_end(&mut content).unwrap();
        if name == "word/styles.xml" {
            let xml: String = String::from_utf8(content).unwrap();
            let start: usize = xml.find("<w:rPrDefault>").unwrap() + "<w:rPrDefault>".len();
            let end: usize = xml.find("</w:rPrDefault>").unwrap();
            content = format!(
                r#"{}<w:rPr><w:lang w:val="da-DK"/></w:rPr>{}"#,
                &xml[..start],
                &xml[end..]
            )
            .into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .unwrap();
        out.write_all(&content).unwrap();
    }
    let data: Vec<u8> = out.finish().unwrap().into_inner();

    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    let pdf: String = String::from_utf8_lossy(&result.pdf).into_owned();
    assert!(
        pdf.contains("/Lang (da-DK)") || pdf.contains("/Lang(da-DK)"),
        "the PDF must declare Danish: {:?}",
        pdf.match_indices("/Lang")
            .map(|(at, _)| &pdf[at..at + 16])
            .collect::<Vec<_>>()
    );
}

#[test]
fn test_current_style_language_reaches_pdf_catalog_after_tracked_change() {
    let data: &[u8] =
        include_bytes!("../../../tests/fixtures/docx/issue-2054-default-language.docx");
    let result = convert_bytes(data, Format::Docx, &ConvertOptions::default()).unwrap();
    let pdf: String = String::from_utf8_lossy(&result.pdf).into_owned();
    assert!(
        pdf.contains("/Lang (de-DE)") || pdf.contains("/Lang(de-DE)"),
        "current German formatting must override historical English formatting"
    );
}

/// A narrow Word column of long words set justified, with `settings` spliced
/// into `word/settings.xml` ahead of its first child.
fn build_narrow_justified_docx(settings: &str) -> Vec<u8> {
    use std::io::{Cursor, Read, Write};

    const TEXT: &str = "Internationalization considerations notwithstanding, the \
        telecommunications infrastructure modernization programme demonstrates \
        extraordinary interoperability characteristics, notwithstanding \
        unquestionably counterproductive misunderstandings regarding \
        responsibilities, accountability, and organizational transformation.";
    // A 250pt page with 36pt margins leaves a 178pt column.
    let docx = docx_rs::Docx::new()
        .page_size(5000, 12000)
        .page_margin(
            docx_rs::PageMargin::new()
                .top(720)
                .bottom(720)
                .left(720)
                .right(720),
        )
        .add_paragraph(
            docx_rs::Paragraph::new()
                .align(docx_rs::AlignmentType::Both)
                .add_run(docx_rs::Run::new().add_text(TEXT)),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();

    let mut archive = zip::ZipArchive::new(Cursor::new(cursor.into_inner())).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name: String = entry.name().to_string();
        let mut content: Vec<u8> = Vec::new();
        entry.read_to_end(&mut content).unwrap();
        if name == "word/settings.xml" {
            let xml: String = String::from_utf8(content).unwrap();
            let first_child: usize = xml.find("<w:defaultTabStop").unwrap();
            content =
                format!("{}{settings}{}", &xml[..first_child], &xml[first_child..]).into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default())
            .unwrap();
        out.write_all(&content).unwrap();
    }
    out.finish().unwrap().into_inner()
}

/// The lines of a converted PDF's text layer that end in a hyphenation point.
/// Typst marks one with a soft hyphen in the text layer.
fn hyphenated_line_ends(pdf: &[u8]) -> Vec<String> {
    let text: String = pdf_extract::extract_text_from_mem(pdf).unwrap();
    text.lines()
        .map(str::trim_end)
        .filter(|line| line.ends_with('\u{ad}') || line.ends_with('-'))
        .map(str::to_string)
        .collect()
}

/// Word does not hyphenate a document unless it declares
/// `w:autoHyphenation`, so a justified paragraph has to wrap whole words.
#[test]
fn test_justified_docx_paragraph_is_not_hyphenated() {
    let generated = build_narrow_justified_docx("");
    let native_audit_fixture: &[u8] =
        include_bytes!("../../../tests/fixtures/docx/issue-2052-default-hyphenation.docx");
    for data in [generated.as_slice(), native_audit_fixture] {
        let result = convert_bytes(data, Format::Docx, &ConvertOptions::default()).unwrap();
        let hyphenated: Vec<String> = hyphenated_line_ends(&result.pdf);
        assert!(
            hyphenated.is_empty(),
            "no line may end in a hyphenation point: {hyphenated:?}"
        );
    }
}

/// Triangulation: the same paragraph in a document that turns automatic
/// hyphenation on does break words, so the rule follows the setting.
#[test]
fn test_docx_auto_hyphenation_hyphenates_justified_paragraph() {
    let data = build_narrow_justified_docx("<w:autoHyphenation/>");
    let result = convert_bytes(&data, Format::Docx, &ConvertOptions::default()).unwrap();
    assert!(
        !hyphenated_line_ends(&result.pdf).is_empty(),
        "a narrow justified column of long words must hyphenate when the document asks"
    );
}
