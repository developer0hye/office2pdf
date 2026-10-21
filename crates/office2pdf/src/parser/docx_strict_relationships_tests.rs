use super::*;
use std::borrow::Cow;
use std::io::Write;

const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// The Strict-conformance twin of a Transitional package: WordprocessingML
/// and the relationships namespace move to their Strict names, and every
/// officeDocument relationship type moves with the latter.
fn strict_twin(data: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().expect("entry name should decode").into_owned();
        let mut content = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut content).unwrap();
        if name.ends_with(".xml") || name.ends_with(".rels") {
            content = String::from_utf8(content)
                .unwrap()
                .replace(TRANSITIONAL_WORD, STRICT_WORD)
                .replace(TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS)
                .into_bytes();
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output.write_all(&content).unwrap();
    }
    output.finish().unwrap().into_inner()
}

fn pack(docx: docx_rs::Docx) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn strict_package_takes_its_styles_part_document_defaults() {
    for (face, half_points) in [("Raleway", 18), ("Georgia", 30)] {
        let styles = docx_rs::Styles::new()
            .default_fonts(docx_rs::RunFonts::new().ascii(face).hi_ansi(face))
            .default_size(half_points);
        let data = build_docx_bytes_with_stylesheet(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Quarterly"))],
            styles,
        );
        let (document, _) = DocxParser
            .parse(&strict_twin(&data), &ConvertOptions::default())
            .unwrap();
        let run = first_run(&document);
        assert_eq!(run.style.font_family.as_deref(), Some(face));
        assert_eq!(run.style.font_size, Some(half_points as f64 / 2.0));
    }
}

#[test]
fn strict_package_renders_like_its_transitional_original() {
    let numberings = docx_rs::Numberings::new()
        .add_abstract_numbering(
            docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
                0,
                docx_rs::Start::new(3),
                docx_rs::NumberFormat::new("upperRoman"),
                docx_rs::LevelText::new("%1)"),
                docx_rs::LevelJc::new("left"),
            )),
        )
        .add_numbering(docx_rs::Numbering::new(1, 0));
    let item = |text: &str| {
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text(text))
            .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0))
    };
    // An external target that happens to spell a Strict relationship type
    // must survive: only `Type` names a relationship type.
    let strict_spelled_target = format!("{STRICT_RELATIONSHIPS}/styles");
    let link = docx_rs::Hyperlink::new(&strict_spelled_target, docx_rs::HyperlinkType::External)
        .add_run(docx_rs::Run::new().add_text("Schema"));
    let picture = docx_rs::Pic::new(&super::image_tests::make_test_bmp()).size(95_250, 95_250);
    let docx = docx_rs::Docx::new()
        .styles(
            docx_rs::Styles::new()
                .default_fonts(docx_rs::RunFonts::new().ascii("Georgia").hi_ansi("Georgia"))
                .default_size(26)
                .add_style(
                    docx_rs::Style::new("Caption", docx_rs::StyleType::Paragraph)
                        .name("Caption")
                        .bold()
                        .size(18),
                ),
        )
        .numberings(numberings)
        .header(docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Report header")),
        ))
        .footer(docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Report footer")),
        ))
        .add_paragraph(item("First item"))
        .add_paragraph(item("Second item"))
        .add_paragraph(
            docx_rs::Paragraph::new()
                .style("Caption")
                .add_run(docx_rs::Run::new().add_text("Figure ").add_image(picture))
                .add_hyperlink(link),
        );
    let transitional = pack(docx);
    let typst = |data: &[u8]| {
        let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
        crate::render::typst_gen::generate_typst(&document)
            .unwrap()
            .source
    };
    let expected = typst(&transitional);
    for needle in [
        "start: 3",
        "Report header",
        "Report footer",
        &strict_spelled_target,
    ] {
        assert!(
            expected.contains(needle),
            "{needle} missing from {expected}"
        );
    }
    assert_eq!(typst(&strict_twin(&transitional)), expected);
}

#[test]
fn strict_fixture_takes_its_theme_face_and_stated_size() {
    // LibreOffice's tdf79272 package: Strict relationship types throughout,
    // `w:docDefaults` naming the minor theme face (Calibri) at `w:sz` 22, and
    // no face or size on any run.
    const FIXTURE: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/libreoffice/tdf79272_strictDxa.docx");
    let (document, _) = DocxParser
        .parse(FIXTURE, &ConvertOptions::default())
        .unwrap();
    let mut runs: Vec<&Run> = Vec::new();
    fn collect<'a>(blocks: &'a [Block], runs: &mut Vec<&'a Run>) {
        for block in blocks {
            match block {
                Block::Paragraph(paragraph) => runs.extend(&paragraph.runs),
                Block::Table(table) => {
                    for row in &table.rows {
                        for cell in &row.cells {
                            collect(&cell.content, runs);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    collect(all_blocks(&document), &mut runs);
    let text_runs: Vec<&Run> = runs
        .into_iter()
        .filter(|run| !run.text.trim().is_empty())
        .collect();
    assert!(!text_runs.is_empty());
    for run in text_runs {
        assert_eq!(run.style.font_family.as_deref(), Some("Calibri"), "{run:?}");
        assert_eq!(run.style.font_size, Some(11.0), "{run:?}");
    }
}

#[test]
fn transitional_package_passes_through_unchanged() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Unchanged")),
    ]);
    let normalized = crate::parser::docx_strict_relationships::normalize(&data);
    assert!(matches!(normalized, Cow::Borrowed(_)));
}
