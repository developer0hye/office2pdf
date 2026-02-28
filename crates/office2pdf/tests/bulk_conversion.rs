//! Bulk conversion smoke tests for all fixture files.
//!
//! These tests iterate over ALL fixture files in `tests/fixtures/` and attempt
//! to convert each one to PDF. The goal is to detect panics — conversion errors
//! are acceptable, but panics are not.
//!
//! Run with: `cargo test -p office2pdf --test bulk_conversion -- --nocapture --ignored`

use std::fmt::Write as FmtWrite;
use std::io::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use office2pdf::config::{ConvertOptions, Format};

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Outcome {
    Success,
    Error,
    Panic,
}

struct FileResult {
    path: PathBuf,
    outcome: Outcome,
    detail: String,
}

struct Summary {
    format: &'static str,
    total: usize,
    success: usize,
    error: usize,
    panic: usize,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixtures_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures")
}

/// Recursively discover all files with the given extension under `dir`.
fn discover_files(dir: &Path, extension: &str) -> Vec<PathBuf> {
    let mut files = Vec::new();
    collect_files_recursive(dir, extension, &mut files);
    files.sort();
    files
}

fn collect_files_recursive(dir: &Path, extension: &str, out: &mut Vec<PathBuf>) {
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_files_recursive(&path, extension, out);
        } else if path
            .extension()
            .is_some_and(|ext| ext.eq_ignore_ascii_case(extension))
        {
            out.push(path);
        }
    }
}

/// Attempt to convert a single file, catching panics.
fn convert_file(path: &Path, format: Format) -> FileResult {
    let data = match std::fs::read(path) {
        Ok(d) => d,
        Err(e) => {
            return FileResult {
                path: path.to_path_buf(),
                outcome: Outcome::Error,
                detail: format!("IO error: {e}"),
            };
        }
    };

    let result = catch_unwind(AssertUnwindSafe(|| {
        office2pdf::convert_bytes(&data, format, &ConvertOptions::default())
    }));

    match result {
        Ok(Ok(convert_result)) => {
            let pdf_size = convert_result.pdf.len();
            FileResult {
                path: path.to_path_buf(),
                outcome: Outcome::Success,
                detail: format!("OK ({pdf_size} bytes)"),
            }
        }
        Ok(Err(e)) => FileResult {
            path: path.to_path_buf(),
            outcome: Outcome::Error,
            detail: format!("{e}"),
        },
        Err(panic_info) => {
            let msg = if let Some(s) = panic_info.downcast_ref::<String>() {
                s.clone()
            } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                (*s).to_string()
            } else {
                "unknown panic".to_string()
            };
            FileResult {
                path: path.to_path_buf(),
                outcome: Outcome::Panic,
                detail: format!("PANIC: {msg}"),
            }
        }
    }
}

/// Run bulk conversion for a single format, returning results and summary.
fn run_bulk_test(
    format_name: &'static str,
    extension: &str,
    format: Format,
) -> (Vec<FileResult>, Summary) {
    let dir = fixtures_dir().join(extension);
    let files = discover_files(&dir, extension);

    println!("\n{}", "=".repeat(60));
    println!("  Bulk {format_name} conversion: {} files", files.len());
    println!("{}\n", "=".repeat(60));

    let mut results = Vec::with_capacity(files.len());

    for (i, path) in files.iter().enumerate() {
        let filename = path
            .file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_default();
        print!("[{}/{}] {filename} ... ", i + 1, files.len());
        std::io::stdout().flush().ok();

        let result = convert_file(path, format);
        match result.outcome {
            Outcome::Success => println!("OK"),
            Outcome::Error => println!("ERROR: {}", result.detail),
            Outcome::Panic => println!("PANIC: {}", result.detail),
        }
        results.push(result);
    }

    let success = results
        .iter()
        .filter(|r| r.outcome == Outcome::Success)
        .count();
    let error = results
        .iter()
        .filter(|r| r.outcome == Outcome::Error)
        .count();
    let panic = results
        .iter()
        .filter(|r| r.outcome == Outcome::Panic)
        .count();

    let summary = Summary {
        format: format_name,
        total: files.len(),
        success,
        error,
        panic,
    };

    (results, summary)
}

/// Format results as a report string.
fn format_report(results: &[FileResult], summary: &Summary) -> String {
    let mut report = String::new();

    writeln!(report, "# Bulk Conversion Report: {}", summary.format).unwrap();
    writeln!(
        report,
        "Total: {} | Success: {} | Error: {} | Panic: {}",
        summary.total, summary.success, summary.error, summary.panic
    )
    .unwrap();
    let rate = if summary.total > 0 {
        (summary.success as f64 / summary.total as f64) * 100.0
    } else {
        0.0
    };
    writeln!(report, "Success rate: {rate:.1}%").unwrap();
    writeln!(report).unwrap();

    // List panics first (most critical)
    let panics: Vec<_> = results
        .iter()
        .filter(|r| r.outcome == Outcome::Panic)
        .collect();
    if !panics.is_empty() {
        writeln!(report, "## PANICS ({} files)", panics.len()).unwrap();
        for r in &panics {
            writeln!(report, "  - {} :: {}", r.path.display(), r.detail).unwrap();
        }
        writeln!(report).unwrap();
    }

    // List errors
    let errors: Vec<_> = results
        .iter()
        .filter(|r| r.outcome == Outcome::Error)
        .collect();
    if !errors.is_empty() {
        writeln!(report, "## ERRORS ({} files)", errors.len()).unwrap();
        for r in &errors {
            writeln!(report, "  - {} :: {}", r.path.display(), r.detail).unwrap();
        }
        writeln!(report).unwrap();
    }

    // List successes
    let successes: Vec<_> = results
        .iter()
        .filter(|r| r.outcome == Outcome::Success)
        .collect();
    if !successes.is_empty() {
        writeln!(report, "## SUCCESSES ({} files)", successes.len()).unwrap();
        for r in &successes {
            writeln!(report, "  - {} :: {}", r.path.display(), r.detail).unwrap();
        }
        writeln!(report).unwrap();
    }

    report
}

/// Print summary table to stdout.
fn print_summary_table(summaries: &[&Summary]) {
    println!("\n{}", "=".repeat(60));
    println!("  BULK CONVERSION SUMMARY");
    println!("{}", "=".repeat(60));
    println!(
        "{:<8} {:>6} {:>8} {:>6} {:>6} {:>8}",
        "Format", "Total", "Success", "Error", "Panic", "Rate"
    );
    println!("{:-<50}", "");

    let mut total_all = 0;
    let mut success_all = 0;
    let mut error_all = 0;
    let mut panic_all = 0;

    for s in summaries {
        let rate = if s.total > 0 {
            (s.success as f64 / s.total as f64) * 100.0
        } else {
            0.0
        };
        println!(
            "{:<8} {:>6} {:>8} {:>6} {:>6} {:>7.1}%",
            s.format, s.total, s.success, s.error, s.panic, rate
        );
        total_all += s.total;
        success_all += s.success;
        error_all += s.error;
        panic_all += s.panic;
    }

    let rate_all = if total_all > 0 {
        (success_all as f64 / total_all as f64) * 100.0
    } else {
        0.0
    };
    println!("{:-<50}", "");
    println!(
        "{:<8} {:>6} {:>8} {:>6} {:>6} {:>7.1}%",
        "TOTAL", total_all, success_all, error_all, panic_all, rate_all
    );
    println!();
}

/// Write results to a file.
fn write_results_file(all_reports: &str) {
    let output_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/bulk_conversion_results.txt");
    if let Err(e) = std::fs::write(&output_path, all_reports) {
        eprintln!("Warning: could not write results file: {e}");
    } else {
        println!("Results written to: {}", output_path.display());
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[test]
#[ignore]
fn test_bulk_docx() {
    let (results, summary) = run_bulk_test("DOCX", "docx", Format::Docx);
    let report = format_report(&results, &summary);
    println!("{report}");
    write_results_file(&report);

    print_summary_table(&[&summary]);

    assert_eq!(
        summary.panic, 0,
        "{} DOCX file(s) caused a panic! See output above for details.",
        summary.panic
    );
}

#[test]
#[ignore]
fn test_bulk_pptx() {
    let (results, summary) = run_bulk_test("PPTX", "pptx", Format::Pptx);
    let report = format_report(&results, &summary);
    println!("{report}");
    write_results_file(&report);

    print_summary_table(&[&summary]);

    assert_eq!(
        summary.panic, 0,
        "{} PPTX file(s) caused a panic! See output above for details.",
        summary.panic
    );
}

#[test]
#[ignore]
fn test_bulk_xlsx() {
    let (results, summary) = run_bulk_test("XLSX", "xlsx", Format::Xlsx);
    let report = format_report(&results, &summary);
    println!("{report}");
    write_results_file(&report);

    print_summary_table(&[&summary]);

    assert_eq!(
        summary.panic, 0,
        "{} XLSX file(s) caused a panic! See output above for details.",
        summary.panic
    );
}

#[test]
#[ignore]
fn test_bulk_all_formats() {
    let (docx_results, docx_summary) = run_bulk_test("DOCX", "docx", Format::Docx);
    let (pptx_results, pptx_summary) = run_bulk_test("PPTX", "pptx", Format::Pptx);
    let (xlsx_results, xlsx_summary) = run_bulk_test("XLSX", "xlsx", Format::Xlsx);

    // Combine all reports
    let mut all_reports = String::new();
    writeln!(
        all_reports,
        "{}",
        format_report(&docx_results, &docx_summary)
    )
    .unwrap();
    writeln!(
        all_reports,
        "{}",
        format_report(&pptx_results, &pptx_summary)
    )
    .unwrap();
    writeln!(
        all_reports,
        "{}",
        format_report(&xlsx_results, &xlsx_summary)
    )
    .unwrap();

    write_results_file(&all_reports);

    print_summary_table(&[&docx_summary, &pptx_summary, &xlsx_summary]);

    let total_panics = docx_summary.panic + pptx_summary.panic + xlsx_summary.panic;
    assert_eq!(
        total_panics, 0,
        "{total_panics} file(s) caused panics across all formats! See output above for details."
    );
}

/// Asserts that the overall conversion success rate meets the 70% target (US-205).
///
/// This test runs all formats and verifies the combined success rate is at or
/// above 70%. Password-protected or intentionally broken files that return
/// `ConvertError` are acceptable — only the success rate matters.
#[test]
#[ignore]
fn test_bulk_success_rate_target() {
    const TARGET_RATE: f64 = 70.0;

    let (_docx_results, docx_summary) = run_bulk_test("DOCX", "docx", Format::Docx);
    let (_pptx_results, pptx_summary) = run_bulk_test("PPTX", "pptx", Format::Pptx);
    let (_xlsx_results, xlsx_summary) = run_bulk_test("XLSX", "xlsx", Format::Xlsx);

    let summaries = [&docx_summary, &pptx_summary, &xlsx_summary];
    print_summary_table(&summaries);

    let total: usize = summaries.iter().map(|s| s.total).sum();
    let success: usize = summaries.iter().map(|s| s.success).sum();
    let rate = if total > 0 {
        (success as f64 / total as f64) * 100.0
    } else {
        0.0
    };

    // Per-format rates
    for s in &summaries {
        let fmt_rate = if s.total > 0 {
            (s.success as f64 / s.total as f64) * 100.0
        } else {
            0.0
        };
        println!("{}: {}/{} ({:.1}%)", s.format, s.success, s.total, fmt_rate);
    }
    println!("Overall: {success}/{total} ({rate:.1}%)");

    assert!(
        rate >= TARGET_RATE,
        "Overall success rate {rate:.1}% is below the {TARGET_RATE}% target. \
         {success}/{total} files converted successfully."
    );
}
