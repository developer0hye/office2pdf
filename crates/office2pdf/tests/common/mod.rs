//! Shared test utilities for integration tests.
//!
//! Not all test binaries use every function — suppress dead code warnings.
#![allow(dead_code)]

use std::path::{Path, PathBuf};

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

/// Check if `pdftoppm` (from poppler-utils) is available on the system.
pub fn is_pdftoppm_available() -> bool {
    std::process::Command::new("pdftoppm")
        .arg("-v")
        .output()
        .is_ok_and(|o| o.status.success() || !o.stderr.is_empty())
}

/// Check if `pdftotext` (from poppler-utils) is available on the system.
pub fn is_pdftotext_available() -> bool {
    std::process::Command::new("pdftotext")
        .arg("-v")
        .output()
        .is_ok_and(|o| o.status.success() || !o.stderr.is_empty())
}

/// Render all pages of a PDF file to PNG images using `pdftoppm`.
///
/// Returns the list of generated PNG file paths, sorted by page number.
/// The files are named `{prefix}-{page}.png` in the output directory.
pub fn render_pdf_to_pngs(
    pdf_path: &Path,
    output_dir: &Path,
    prefix: &str,
    dpi: u32,
) -> Vec<PathBuf> {
    std::fs::create_dir_all(output_dir).expect("create output dir");

    let output_prefix = output_dir.join(prefix);
    let status = std::process::Command::new("pdftoppm")
        .args(["-png", "-r", &dpi.to_string()])
        .arg(pdf_path)
        .arg(&output_prefix)
        .status()
        .expect("run pdftoppm");

    assert!(
        status.success(),
        "pdftoppm failed for {}",
        pdf_path.display()
    );

    collect_pngs(output_dir, prefix)
}

/// Render all pages of in-memory PDF bytes to PNG images using `pdftoppm`.
///
/// Writes the bytes to a temp file, then delegates to `render_pdf_to_pngs`.
pub fn render_pdf_bytes_to_pngs(
    pdf_bytes: &[u8],
    output_dir: &Path,
    prefix: &str,
    dpi: u32,
) -> Vec<PathBuf> {
    std::fs::create_dir_all(output_dir).expect("create output dir");

    let temp_pdf = output_dir.join(format!("{prefix}.pdf"));
    std::fs::write(&temp_pdf, pdf_bytes).expect("write temp PDF");

    let result = render_pdf_to_pngs(&temp_pdf, output_dir, prefix, dpi);
    let _ = std::fs::remove_file(&temp_pdf);
    result
}

/// Extract text from a PDF file using `pdftotext`.
///
/// Returns the extracted text. Useful for structural comparison against ground truth.
pub fn extract_text_from_pdf_file(pdf_path: &Path) -> String {
    let output = std::process::Command::new("pdftotext")
        .args(["-enc", "UTF-8"])
        .arg(pdf_path)
        .arg("-") // output to stdout
        .output()
        .expect("run pdftotext");

    assert!(
        output.status.success(),
        "pdftotext failed for {}",
        pdf_path.display()
    );
    String::from_utf8_lossy(&output.stdout).into_owned()
}

/// Extract text from in-memory PDF bytes using `pdftotext`.
pub fn extract_text_from_pdf_bytes(pdf_bytes: &[u8], work_dir: &Path) -> String {
    std::fs::create_dir_all(work_dir).expect("create work dir");
    let temp_pdf = work_dir.join(format!("_pdftotext_tmp_{}.pdf", std::process::id()));
    std::fs::write(&temp_pdf, pdf_bytes).expect("write temp PDF");
    let result = extract_text_from_pdf_file(&temp_pdf);
    let _ = std::fs::remove_file(&temp_pdf);
    result
}

/// Get page count from a PDF file using `pdfinfo`.
pub fn pdf_page_count(pdf_path: &Path) -> Option<u32> {
    let output = std::process::Command::new("pdfinfo")
        .arg(pdf_path)
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }
    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("Pages:") {
            return rest.trim().parse().ok();
        }
    }
    None
}

/// Collect PNG files matching `{prefix}-*.png` in a directory, sorted by page number.
fn collect_pngs(dir: &Path, prefix: &str) -> Vec<PathBuf> {
    let mut pngs: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read output dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| {
            p.extension().is_some_and(|ext| ext == "png")
                && p.file_name().unwrap().to_string_lossy().starts_with(prefix)
        })
        .collect();
    pngs.sort();
    pngs
}
