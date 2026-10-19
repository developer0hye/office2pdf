//! Round-trip and surgical-edit tests for the lossless PPTX model
//! (issue #746).
//!
//! Saved containers are re-read with the `zip` crate and rendered through the
//! existing `office2pdf` pipeline, so the model under test never verifies its
//! own output.

use std::io::{Cursor, Read};
use std::path::PathBuf;

use pptx_model::{PptxDocument, PptxError};

fn fixture_bytes(relative: &str) -> Vec<u8> {
    let path: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures")
        .join(relative);
    std::fs::read(path).expect("fixture should be readable")
}

/// `(name, raw compressed payload)` for every non-directory entry, in order.
fn raw_zip_entries(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("container should open");
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index_raw(index).expect("raw entry should open");
        if file.is_dir() {
            continue;
        }
        let mut payload: Vec<u8> = Vec::new();
        file.read_to_end(&mut payload)
            .expect("raw entry should read");
        entries.push((
            file.name().expect("entry name should decode").into_owned(),
            payload,
        ));
    }
    entries
}

fn zip_entry(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("container should open");
    let mut file = archive.by_name(name).expect("entry should exist");
    let mut payload: Vec<u8> = Vec::new();
    file.read_to_end(&mut payload)
        .expect("entry should decompress");
    payload
}

fn convert_to_pdf(container: &[u8]) -> Vec<u8> {
    office2pdf::convert_bytes(
        container,
        office2pdf::config::Format::Pptx,
        &office2pdf::config::ConvertOptions::default(),
    )
    .expect("conversion should succeed")
    .pdf
}

#[test]
fn slides_enumerate_in_presentation_order() {
    let two_slides =
        PptxDocument::from_bytes(fixture_bytes("pptx/WithMaster.pptx")).expect("should load");
    assert_eq!(two_slides.slide_count(), 2);
    assert_eq!(
        two_slides.slide_part_names(),
        ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]
    );

    // The rels part of this fixture lists relationships out of order, so the
    // expected sequence can only come from the sldIdLst -> rels mapping.
    let ten_slides =
        PptxDocument::from_bytes(fixture_bytes("pptx/10-slides.pptx")).expect("should load");
    let expected: Vec<String> = (1..=10)
        .map(|number| format!("ppt/slides/slide{number}.xml"))
        .collect();
    assert_eq!(ten_slides.slide_part_names(), expected.as_slice());

    let empty =
        PptxDocument::from_bytes(fixture_bytes("pptx/no-slides.pptx")).expect("should load");
    assert_eq!(empty.slide_count(), 0);
}

#[test]
fn text_runs_are_listed_in_document_order() {
    let document =
        PptxDocument::from_bytes(fixture_bytes("pptx/WithMaster.pptx")).expect("should load");
    assert_eq!(
        document.slide_text_runs(0).expect("slide 1 should parse"),
        ["First page title", "First page subtitle"]
    );
    assert_eq!(
        document.slide_text_runs(1).expect("slide 2 should parse"),
        ["2", "nd", " page subtitle", "Footer from the master slide"]
    );
}

#[test]
fn editing_one_text_run_touches_only_that_slide_part() {
    let original: Vec<u8> = fixture_bytes("pptx/WithMaster.pptx");
    let mut document = PptxDocument::from_bytes(original.clone()).expect("should load");
    document
        .set_slide_text_run(0, 0, "Edited page title")
        .expect("edit should succeed");
    assert_eq!(
        document.package().dirty_part_names(),
        ["ppt/slides/slide1.xml"],
        "exactly the edited slide part should be dirty"
    );

    let saved: Vec<u8> = document.save_to_bytes().expect("save should succeed");

    // The edited slide differs only in the spliced character data.
    let expected_slide: Vec<u8> = String::from_utf8(zip_entry(&original, "ppt/slides/slide1.xml"))
        .expect("slide XML should be UTF-8")
        .replace("First page title", "Edited page title")
        .into_bytes();
    assert_eq!(
        zip_entry(&saved, "ppt/slides/slide1.xml"),
        expected_slide,
        "the edit should splice character data and leave the rest of the slide byte-identical"
    );

    // Every other package part keeps its exact compressed payload.
    for ((name, original_payload), (saved_name, saved_payload)) in raw_zip_entries(&original)
        .iter()
        .zip(raw_zip_entries(&saved).iter())
    {
        assert_eq!(name, saved_name, "entry order should be preserved");
        if name != "ppt/slides/slide1.xml" {
            assert_eq!(
                original_payload, saved_payload,
                "untouched entry {name} should keep its exact compressed payload"
            );
        }
    }

    // The saved file reopens in this model with the edit applied.
    let reloaded = PptxDocument::from_bytes(saved).expect("saved file should reload");
    assert_eq!(
        reloaded.slide_text_runs(0).expect("slide 1 should parse"),
        ["Edited page title", "First page subtitle"]
    );
}

#[test]
fn replacement_text_is_escaped_and_survives_reload() {
    let mut document =
        PptxDocument::from_bytes(fixture_bytes("pptx/WithMaster.pptx")).expect("should load");
    let replacement: &str = "Q3 <Sales & \"Profit\">";
    document
        .set_slide_text_run(0, 1, replacement)
        .expect("edit should succeed");

    let saved: Vec<u8> = document.save_to_bytes().expect("save should succeed");
    let slide_xml: String = String::from_utf8(zip_entry(&saved, "ppt/slides/slide1.xml"))
        .expect("slide XML should be UTF-8");
    assert!(
        !slide_xml.contains("<Sales"),
        "markup characters must be escaped in the stored XML"
    );

    let reloaded = PptxDocument::from_bytes(saved).expect("saved file should reload");
    assert_eq!(
        reloaded.slide_text_runs(0).expect("slide 1 should parse"),
        ["First page title", replacement]
    );
}

#[test]
fn empty_text_runs_can_be_read_and_edited() {
    // No committed fixture carries a self-closing <a:t/>, so craft one from
    // WithMaster.pptx through the package layer.
    let mut package = ooxml_package::OpcPackage::from_bytes(fixture_bytes("pptx/WithMaster.pptx"))
        .expect("package should load");
    let slide_xml: String = String::from_utf8(
        package
            .part_bytes("ppt/slides/slide1.xml")
            .expect("slide should read")
            .into_owned(),
    )
    .expect("slide XML should be UTF-8");
    package
        .set_part_bytes(
            "ppt/slides/slide1.xml",
            slide_xml
                .replace("<a:t>First page title</a:t>", "<a:t/>")
                .into_bytes(),
        )
        .expect("slide should accept the crafted XML");
    let crafted: Vec<u8> = package.save_to_bytes().expect("package should save");

    let mut document = PptxDocument::from_bytes(crafted).expect("crafted file should load");
    assert_eq!(
        document.slide_text_runs(0).expect("slide 1 should parse"),
        ["", "First page subtitle"],
        "a self-closing a:t is an empty run"
    );

    document
        .set_slide_text_run(0, 0, "Restored title")
        .expect("empty run should accept text");
    let reloaded = PptxDocument::from_bytes(document.save_to_bytes().expect("save should succeed"))
        .expect("saved file should reload");
    assert_eq!(
        reloaded.slide_text_runs(0).expect("slide 1 should parse"),
        ["Restored title", "First page subtitle"]
    );
}

#[test]
fn no_op_round_trip_projects_to_an_identical_pdf() {
    let original: Vec<u8> = fixture_bytes("pptx/WithMaster.pptx");
    let document = PptxDocument::from_bytes(original.clone()).expect("should load");
    let saved: Vec<u8> = document.save_to_bytes().expect("save should succeed");

    let original_pdf: Vec<u8> = convert_to_pdf(&original);
    let round_trip_pdf: Vec<u8> = convert_to_pdf(&saved);
    assert_eq!(
        original_pdf, round_trip_pdf,
        "a no-op round trip must not change the rendered PDF"
    );
}

#[test]
fn edited_presentation_projects_the_new_text_into_the_pdf() {
    let mut document =
        PptxDocument::from_bytes(fixture_bytes("pptx/WithMaster.pptx")).expect("should load");
    document
        .set_slide_text_run(0, 0, "Quarterly Review 2026")
        .expect("edit should succeed");
    let saved: Vec<u8> = document.save_to_bytes().expect("save should succeed");

    let pdf_text: String = pdf_extract::extract_text_from_mem(&convert_to_pdf(&saved))
        .expect("PDF text should extract");
    // pdf-extract pads inter-word gaps with variable spacing; compare on
    // whitespace-normalized text.
    let normalized: String = pdf_text.split_whitespace().collect::<Vec<&str>>().join(" ");
    assert!(
        normalized.contains("Quarterly Review 2026"),
        "edited text should reach the rendered PDF, got: {normalized:?}"
    );
    assert!(
        !normalized.contains("First page title"),
        "replaced text should no longer render"
    );
}

#[test]
fn non_presentation_input_is_rejected() {
    let docx_error: PptxError =
        PptxDocument::from_bytes(fixture_bytes("docx/unit_test_headers.docx")).unwrap_err();
    assert!(
        matches!(docx_error, PptxError::NotAPresentation(_)),
        "a DOCX container must be rejected as not-a-presentation, got: {docx_error}"
    );

    let garbage_error: PptxError =
        PptxDocument::from_bytes(b"not an OPC container".to_vec()).unwrap_err();
    assert!(
        matches!(garbage_error, PptxError::Package(_)),
        "garbage bytes must surface the package error, got: {garbage_error}"
    );
}

#[test]
fn out_of_range_indices_are_reported() {
    let mut document =
        PptxDocument::from_bytes(fixture_bytes("pptx/WithMaster.pptx")).expect("should load");

    let slide_error: PptxError = document.slide_text_runs(5).unwrap_err();
    assert!(
        matches!(
            slide_error,
            PptxError::SlideIndexOutOfRange {
                index: 5,
                slide_count: 2,
            }
        ),
        "expected SlideIndexOutOfRange, got: {slide_error}"
    );

    let run_error: PptxError = document.set_slide_text_run(0, 99, "text").unwrap_err();
    assert!(
        matches!(
            &run_error,
            PptxError::TextRunIndexOutOfRange {
                index: 99,
                run_count: 2,
                ..
            }
        ),
        "expected TextRunIndexOutOfRange, got: {run_error}"
    );
}

/// Writes round-trip artifacts for the manual acceptance check of issue #746:
/// open both files in Microsoft PowerPoint and confirm neither raises a
/// repair prompt. Set `PPTX_MODEL_ARTIFACT_DIR` to choose the output
/// directory (defaults to the system temp directory).
#[test]
#[ignore = "writes local artifacts for the manual PowerPoint acceptance check"]
fn write_powerpoint_verification_artifacts() {
    let output_dir: PathBuf = std::env::var_os("PPTX_MODEL_ARTIFACT_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir);
    let original: Vec<u8> = fixture_bytes("pptx/WithMaster.pptx");

    let no_op = PptxDocument::from_bytes(original.clone()).expect("should load");
    let no_op_path: PathBuf = output_dir.join("withmaster-no-op-roundtrip.pptx");
    std::fs::write(
        &no_op_path,
        no_op.save_to_bytes().expect("save should succeed"),
    )
    .expect("artifact should write");

    let mut edited = PptxDocument::from_bytes(original).expect("should load");
    edited
        .set_slide_text_run(0, 0, "Edited page title")
        .expect("edit should succeed");
    let edited_path: PathBuf = output_dir.join("withmaster-edited-roundtrip.pptx");
    std::fs::write(
        &edited_path,
        edited.save_to_bytes().expect("save should succeed"),
    )
    .expect("artifact should write");

    println!("no-op artifact:  {}", no_op_path.display());
    println!("edited artifact: {}", edited_path.display());
}
