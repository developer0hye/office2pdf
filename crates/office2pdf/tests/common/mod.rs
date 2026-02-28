//! Shared test utilities for integration tests.

/// Extract all visible text content from PDF bytes.
///
/// Returns the concatenated text from all pages. Useful for verifying
/// that key content markers from source documents appear in the final PDF.
///
/// Panics if the PDF cannot be parsed.
pub fn extract_pdf_text(pdf_bytes: &[u8]) -> String {
    pdf_extract::extract_text_from_mem(pdf_bytes).expect("should extract text from PDF")
}

/// Validate PDF bytes using `qpdf --check`.
///
/// Returns `true` if validation was performed and passed, `false` if skipped.
///
/// Validation is skipped when:
/// - `OFFICE2PDF_VALIDATE_PDF` env var is not set to `"1"`
/// - `qpdf` is not installed on the system
///
/// Panics if `qpdf --check` reports the PDF is invalid.
pub fn validate_pdf_with_qpdf(pdf_bytes: &[u8]) -> bool {
    // Gate on environment variable
    if std::env::var("OFFICE2PDF_VALIDATE_PDF").unwrap_or_default() != "1" {
        return false;
    }

    // Check if qpdf is available
    match std::process::Command::new("qpdf").arg("--version").output() {
        Ok(output) if output.status.success() => {}
        _ => {
            eprintln!("[WARN] qpdf not installed, skipping PDF validation");
            return false;
        }
    }

    // Write PDF bytes to a temp file
    let temp_path = std::env::temp_dir().join(format!(
        "office2pdf_test_{}_{}.pdf",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));

    std::fs::write(&temp_path, pdf_bytes).expect("should write temp PDF file");

    let output = std::process::Command::new("qpdf")
        .arg("--check")
        .arg(&temp_path)
        .output()
        .expect("should run qpdf");

    // Clean up temp file before asserting
    let _ = std::fs::remove_file(&temp_path);

    assert!(
        output.status.success(),
        "qpdf --check failed:\nstdout: {}\nstderr: {}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );

    true
}
