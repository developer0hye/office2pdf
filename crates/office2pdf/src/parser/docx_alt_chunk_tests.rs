use super::*;
use crate::ir::{Block, Document, Page};
use std::io::{Cursor, Write};

fn package(chunk: &[u8], content_type: &str, target: &str, mode: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let entries = [
        ("[Content_Types].xml", format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/import.html" ContentType="{content_type}"/></Types>"#)),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.into()),
        ("word/_rels/document.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="import" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="{target}" TargetMode="{mode}"/></Relationships>"#)),
        ("word/document.xml", r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><d:altChunk xmlns:d="http://schemas.openxmlformats.org/wordprocessingml/2006/main" rel:id="import"/><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.into()),
    ];
    for (name, data) in entries {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.start_file("word/import.html", zip::write::SimpleFileOptions::default())
        .unwrap();
    zip.write_all(chunk).unwrap();
    zip.finish().unwrap().into_inner()
}

fn paragraphs(document: &Document) -> Vec<&crate::ir::Paragraph> {
    fn collect<'a>(blocks: &'a [Block], out: &mut Vec<&'a crate::ir::Paragraph>) {
        for block in blocks {
            match block {
                Block::Paragraph(p) => out.push(p),
                Block::Table(t) => {
                    for row in &t.rows {
                        for cell in &row.cells {
                            collect(&cell.content, out);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    let mut result = Vec::new();
    for page in &document.pages {
        if let Page::Flow(flow) | Page::FlowContinuous(flow) = page {
            collect(&flow.content, &mut result);
        }
    }
    result
}

#[test]
fn alt_chunk_html_keeps_reading_order_entities_and_inherited_styles() {
    let bytes = package(br#"<!doctype html><html><head><style>p {font-family:Arial;font-size:14pt;color:#123456} .report {font-weight:bold} p.report {font-size:16pt}</style></head><body><p class="report extra">Imported <em>styled</em> &amp; valid<br>line<p>Next &nbsp; item<table><tr><td>Cell one<td>Cell two</table><script>never printed</script></body></html>"#, "text/html", "./import.html", "Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    let ps = paragraphs(&document);
    let text: Vec<String> = ps
        .iter()
        .map(|p| p.runs.iter().map(|r| r.text.as_str()).collect())
        .collect();
    assert_eq!(
        text,
        [
            "Before",
            "Imported styled & valid\nline",
            "Next \u{a0} item",
            "Cell one",
            "Cell two",
            "After"
        ]
    );
    let styled = ps[1]
        .runs
        .iter()
        .find(|r| r.text.contains("styled"))
        .unwrap();
    assert_eq!(styled.style.font_family.as_deref(), Some("Arial"));
    assert_eq!(styled.style.font_size, Some(16.0));
    assert_eq!(styled.style.bold, Some(true));
    assert_eq!(styled.style.italic, Some(true));
}

#[test]
fn alt_chunk_mhtml_decodes_folded_headers_transfer_encoding_and_charset() {
    let mime = b"MIME-Version: 1.0\r\nContent-Type: multipart/related;\r\n boundary=\"report\"\r\n\r\n--report\r\nContent-Type: text/html; charset=windows-1252\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<html><body><p>R=E9sum=E9 =\r\ncomplete</p></body></html>\r\n--report--\r\n";
    let bytes = package(mime, "message/rfc822", "/word/import.html", "Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(
        paragraphs(&document).iter().any(|p| p
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>()
            == "Résumé complete")
    );
}

#[test]
fn alt_chunk_external_or_unsupported_import_is_reported_without_fetching() {
    for (content_type, target, mode) in [
        (
            "text/html",
            "https://example.invalid/report.html",
            "External",
        ),
        ("application/rtf", "import.html", "Internal"),
        ("text/html", "../../../import.html", "Internal"),
    ] {
        let bytes = package(b"<p>Not imported</p>", content_type, target, mode);
        let (document, warnings) = DocxParser
            .parse(&bytes, &ConvertOptions::default())
            .unwrap();
        assert!(
            !warnings.is_empty(),
            "lost content must be reported for {target}"
        );
        assert_eq!(paragraphs(&document).len(), 2);
    }
}

#[test]
fn alt_chunk_mixed_blocks_do_not_drop_unwrapped_text_or_rest_of_import() {
    let bytes=package(b"<html><body>Intro <strong>bold</strong><p>Main paragraph</p>Tail<div>Nested text<p>Inside</p>End</div><p>Image caption<img src='not-fetched.png'></p></body></html>","text/html","import.html","Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    let text: Vec<String> = paragraphs(&document)
        .iter()
        .map(|p| p.runs.iter().map(|r| r.text.as_str()).collect())
        .collect();
    assert_eq!(
        text,
        [
            "Before",
            "Intro bold",
            "Main paragraph",
            "Tail",
            "Nested text",
            "Inside",
            "End",
            "Image caption",
            "After"
        ]
    );
    assert!(
        warnings
            .iter()
            .any(|warning| format!("{warning:?}").contains("images"))
    );
}

#[test]
fn alt_chunk_base64_mime_import_preserves_unicode() {
    let mime=b"MIME-Version: 1.0\nContent-Type: text/html; charset=utf-8\nContent-Transfer-Encoding: base64\n\nPHA+7ZWc6riAIOyInOyEnDwvcD4=";
    let bytes = package(mime, "message/rfc822", "import.html", "Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(
        paragraphs(&document).iter().any(|p| p
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>()
            == "한글 순서")
    );
}

#[test]
fn alt_chunk_tracked_fixture_paints_searchable_imported_content() {
    let bytes = include_bytes!("../../../../tests/fixtures/docx/altchunk_html.docx");
    let (document, warnings) = DocxParser.parse(bytes, &ConvertOptions::default()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    assert_eq!(
        paragraphs(&document)[0]
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>(),
        "Imported HTML keeps important text in the document."
    );
    #[cfg(not(target_arch = "wasm32"))]
    {
        let options = ConvertOptions::default();
        let pdf = crate::convert_bytes(bytes, crate::config::Format::Docx, &options).unwrap();
        assert!(
            pdf_extract::extract_text_from_mem(&pdf.pdf)
                .unwrap()
                .contains("Imported HTML keeps")
        );
    }
}

#[test]
fn alt_chunk_span_spacing_follows_word_import_without_changing_auto_spacing() {
    let bytes=package(b"<html><body><p style='margin-bottom:13px'><span style='margin:0in'>Zero span</span></p><p style='margin-bottom:13px'><span>Plain span</span></p><p><span style='margin:0in'>Auto paragraph</span></p></body></html>","text/html","import.html","Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    let ps = paragraphs(&document);
    assert_eq!(ps[1].style.space_after, Some(0.0));
    assert_eq!(ps[2].style.space_after, Some(9.75));
    assert_eq!(ps[3].style.space_after, Some(14.0));
}

#[test]
fn alt_chunk_mhtml_preserves_literal_utf8_in_quoted_printable_parts() {
    let mime = "MIME-Version: 1.0\r\nContent-Type: text/html; charset=utf-8\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<p>Članak ovlaštenih inženjera =3D complete</p>";
    let bytes = package(mime.as_bytes(), "message/rfc822", "import.html", "Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    assert_eq!(
        paragraphs(&document)[1]
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>(),
        "Članak ovlaštenih inženjera = complete"
    );
}

#[test]
fn alt_chunk_auto_spacing_does_not_shift_the_first_cell_paragraph() {
    let bytes=package(b"<html><body><table><tr><td><p>First in cell</p><p>Second in cell</p></td></tr></table></body></html>","text/html","import.html","Internal");
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    let ps = paragraphs(&document);
    assert_eq!(ps[1].style.space_before, Some(0.0));
    assert_eq!(ps[2].style.space_before, Some(14.0));
}

#[test]
fn alt_chunk_strict_word_and_relationship_namespaces_import_text() {
    let original = package(
        b"<p>Strict import</p>",
        "text/html",
        "import.html",
        "Internal",
    );
    let mut archive = zip::ZipArchive::new(Cursor::new(original)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let mut data = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut data).unwrap();
        let xml = String::from_utf8(data)
            .unwrap()
            .replace(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
            )
            .replace(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
            );
        writer
            .start_file(
                entry.name().expect("entry name should decode").into_owned(),
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
        writer.write_all(xml.as_bytes()).unwrap();
    }
    let bytes = writer.finish().unwrap().into_inner();
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(
        paragraphs(&document).iter().any(|p| p
            .runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>()
            == "Strict import")
    );
}
