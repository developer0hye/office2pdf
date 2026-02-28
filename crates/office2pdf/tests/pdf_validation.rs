//! PDF validation tests using qpdf.
//!
//! These tests convert real fixture files to PDF and validate the output
//! using `qpdf --check`. Validation only runs when:
//! - `OFFICE2PDF_VALIDATE_PDF=1` environment variable is set
//! - `qpdf` is installed on the system
//!
//! In CI, both conditions are met. Locally, tests pass without qpdf
//! (validation is simply skipped).

mod common;

use std::path::PathBuf;

fn fixture_path(format: &str, name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures")
        .join(format)
        .join(name)
}

// ---------------------------------------------------------------------------
// Env-var gating
// ---------------------------------------------------------------------------

#[test]
fn validation_skipped_without_env_var() {
    // This test verifies the gating logic: when the env var is not set,
    // validation should be skipped (return false). We can only test this
    // path when the env var is actually unset.
    if std::env::var("OFFICE2PDF_VALIDATE_PDF").unwrap_or_default() == "1" {
        // In CI the env var is set — skip this test since we cannot safely
        // unset it in a parallel test environment.
        return;
    }

    let dummy_pdf = b"%PDF-1.4 dummy content";
    let result = common::validate_pdf_with_qpdf(dummy_pdf);
    assert!(
        !result,
        "validation should be skipped when env var is not set"
    );
}

// ---------------------------------------------------------------------------
// DOCX PDF validation
// ---------------------------------------------------------------------------

#[test]
fn validate_docx_table_pdf() {
    let path = fixture_path("docx", "table.docx");
    let result = office2pdf::convert(&path).expect("DOCX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}

#[test]
fn validate_docx_styles_pdf() {
    let path = fixture_path("docx", "styles_en.docx");
    let result = office2pdf::convert(&path).expect("DOCX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}

// ---------------------------------------------------------------------------
// PPTX PDF validation
// ---------------------------------------------------------------------------

#[test]
fn validate_pptx_sample_pdf() {
    let path = fixture_path("pptx", "powerpoint_sample.pptx");
    let result = office2pdf::convert(&path).expect("PPTX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}

#[test]
fn validate_pptx_test_slides_pdf() {
    let path = fixture_path("pptx", "test_slides.pptx");
    let result = office2pdf::convert(&path).expect("PPTX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}

// ---------------------------------------------------------------------------
// XLSX PDF validation
// ---------------------------------------------------------------------------

#[test]
fn validate_xlsx_temperature_pdf() {
    let path = fixture_path("xlsx", "temperature.xlsx");
    let result = office2pdf::convert(&path).expect("XLSX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}

#[test]
fn validate_xlsx_formatted_pdf() {
    let path = fixture_path("xlsx", "SH106-Formatted.xlsx");
    let result = office2pdf::convert(&path).expect("XLSX conversion should succeed");
    assert!(
        result.pdf.starts_with(b"%PDF"),
        "output should be a valid PDF"
    );
    common::validate_pdf_with_qpdf(&result.pdf);
}
