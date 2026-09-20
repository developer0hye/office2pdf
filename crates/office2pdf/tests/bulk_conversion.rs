#![cfg(not(target_arch = "wasm32"))] // native-only integration tests (fs, qpdf, criterion)
//! Bulk conversion regression tests for all fixture files.
//!
//! These tests iterate over ALL fixture files in `tests/fixtures/` and attempt
//! to convert each one to PDF. The CI gate rejects panics and any fixture that
//! converted successfully in the committed baseline but no longer succeeds.
//! New ordinary conversion errors are reported without blocking corpus growth.
//!
//! Run the gate with:
//! `cargo test --release -p office2pdf --test bulk_conversion test_bulk_regression_gate -- --ignored --nocapture --exact`
//!
//! Set `UPDATE_BULK_BASELINE=1` to intentionally replace the baseline with the
//! current outcomes after reviewing all changes.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as FmtWrite;
use std::io::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use office2pdf::config::{ConvertOptions, Format};

// ---------------------------------------------------------------------------
// Denylist — adversarial, XML-bomb, or OOM-inducing fixtures.
// Excluded from bulk testing so they do not skew quality metrics.
// See: https://github.com/developer0hye/office2pdf/issues/77
// ---------------------------------------------------------------------------

const DENYLIST: &[&str] = &[
    // ── DOCX — fuzzer-generated / corrupted zip structures ───────────
    "clusterfuzz-testcase-minimized-POIFuzzer-6709287337197568.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-4791943399604224.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-4959857092198400.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-4961551840247808.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-5166796835258368.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-5313273089884160.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-5564805011079168.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-5569740188549120.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-6061520554164224.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-6120975439364096.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-6442791109263360.docx",
    "clusterfuzz-testcase-minimized-POIXWPFFuzzer-6733884933668864.docx",
    // Crash reporter — corrupted zip
    "crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx",
    // Truncated archive — incomplete zip
    "truncated62886.docx",
    // ── PPTX — fuzzer-generated / corrupted zip structures ───────────
    "clusterfuzz-testcase-minimized-POIFuzzer-5205835528404992.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-4838644450394112.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-4986044400861184.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-5463285576892416.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-5471515212382208.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-5611274456596480.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-6071540680032256.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-6254434927378432.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-6372932378820608.pptx",
    "clusterfuzz-testcase-minimized-POIXSLFFuzzer-6435650376957952.pptx",
    // Corrupted archive (OOM / hang)
    "Divino_Revelado.pptx",
    // ── XLSX — fuzzer-generated / corrupted zip structures ───────────
    "clusterfuzz-testcase-minimized-POIFuzzer-5040805309710336.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-4828727001088000.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-5089447305609216.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-5185049589579776.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-5265527465181184.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-5937385319563264.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-6123461607817216.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-6419366255919104.xlsx",
    "clusterfuzz-testcase-minimized-POIXSSFFuzzer-6448258963341312.xlsx",
    "clusterfuzz-testcase-minimized-XLSX2CSVFuzzer-5025401116950528.xlsx",
    "clusterfuzz-testcase-minimized-XLSX2CSVFuzzer-5542865479270400.xlsx",
    "clusterfuzz-testcase-minimized-XLSX2CSVFuzzer-5636439151607808.xlsx",
    "clusterfuzz-testcase-minimized-XLSX2CSVFuzzer-6504225896792064.xlsx",
    "clusterfuzz-testcase-minimized-XLSX2CSVFuzzer-6594557414080512.xlsx",
    // Crash reporters — corrupted zip
    "crash-274d6342e4842d61be0fb48eaadad6208ae767ae.xlsx",
    "crash-9bf3cd4bd6f50a8a9339d363c2c7af14b536865c.xlsx",
    // Corrupted / truncated archive
    "58616.xlsx",
    // ── XLSX — adversarial / OOM-inducing ────────────────────────────
    // XML billion-laughs attack PoCs
    "poc-xmlbomb.xlsx",
    "poc-xmlbomb-empty.xlsx",
    // XML bomb variants (lol9 entity expansion)
    "54764.xlsx",
    "54764-2.xlsx",
    // Shared string table bomb (OOM)
    "poc-shared-strings.xlsx",
    // Hangs during conversion (CI timeout)
    "bug62181.xlsx",
];

// ---------------------------------------------------------------------------
// Expected errors — files that produce errors by design (e.g. encrypted).
// These exercise a valid rejection path and must
// not count against the conversion success rate.
// See: https://github.com/developer0hye/office2pdf/issues/82
// ---------------------------------------------------------------------------

const EXPECTED_ERRORS: &[&str] = &[
    // Conflicting ASCII and Unicode ZIP part names select different content.
    "unicode-path.docx",
    // Encrypted DOCX (OLE2 containers, password-protected)
    "Encrypted_LO_Standard_abc.docx",
    "Encrypted_MSO2007_abc.docx",
    "Encrypted_MSO2010_abc.docx",
    "Encrypted_MSO2013_abc.docx",
    "bug53475-password-is-pass.docx",
    "bug53475-password-is-solrcell.docx",
    // Encrypted XLSX (OLE2 container, password-protected)
    "protected_passtika.xlsx",
    // Invalid cell coordinates past Excel's grid are rejected before layout.
    "too-many-cols-rows.xlsx",
];

/// Returns `true` if the file should be skipped due to being on the denylist.
fn is_denylisted(path: &Path) -> bool {
    path.file_name()
        .and_then(|f| f.to_str())
        .is_some_and(|name| DENYLIST.contains(&name))
}

/// Returns `true` if the file is expected to produce a conversion error.
fn is_expected_error(path: &Path) -> bool {
    path.file_name()
        .and_then(|f| f.to_str())
        .is_some_and(|name| EXPECTED_ERRORS.contains(&name))
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Deserialize, serde::Serialize)]
#[serde(rename_all = "snake_case")]
enum Outcome {
    Success,
    Error,
    /// Error that was expected (e.g. encrypted file → UnsupportedEncryption).
    ExpectedError,
    Panic,
}

struct FileResult {
    path: PathBuf,
    outcome: Outcome,
    detail: String,
}

#[derive(Debug, serde::Deserialize, serde::Serialize)]
struct BulkBaseline {
    version: u32,
    fixtures: BTreeMap<String, Outcome>,
}

impl BulkBaseline {
    fn from_entries(entries: impl IntoIterator<Item = (String, Outcome)>) -> Self {
        Self {
            version: 1,
            fixtures: entries.into_iter().collect(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize)]
#[serde(rename_all = "snake_case")]
enum RegressionKind {
    Panic,
    PreviousSuccessFailed,
    PreviousSuccessMissing,
    SuccessRateDecreased,
}

#[derive(Debug, serde::Serialize)]
struct Regression {
    path: String,
    kind: RegressionKind,
    detail: String,
}

#[derive(Debug, serde::Serialize)]
struct BaselineComparison {
    baseline_success_rate: f64,
    current_success_rate: f64,
    regressions: Vec<Regression>,
    improvements: Vec<String>,
    new_files: Vec<String>,
}

impl BaselineComparison {
    fn passed(&self) -> bool {
        self.regressions.is_empty()
    }
}

#[derive(serde::Serialize)]
struct MachineReport<'a> {
    summaries: Vec<MachineSummary<'a>>,
    files: Vec<MachineFileResult>,
    comparison: &'a BaselineComparison,
}

#[derive(serde::Serialize)]
struct MachineSummary<'a> {
    format: &'a str,
    total: usize,
    skipped: usize,
    success: usize,
    error: usize,
    expected_error: usize,
    panic: usize,
    success_rate: f64,
}

#[derive(serde::Serialize)]
struct MachineFileResult {
    path: String,
    outcome: Outcome,
    detail: String,
}

struct Summary {
    format: &'static str,
    total: usize,
    skipped: usize,
    success: usize,
    error: usize,
    expected_error: usize,
    panic: usize,
}

impl Summary {
    /// Effective total excludes expected errors from the success rate denominator.
    fn effective_total(&self) -> usize {
        self.total - self.expected_error
    }

    fn success_rate(&self) -> f64 {
        let eff = self.effective_total();
        if eff > 0 {
            (self.success as f64 / eff as f64) * 100.0
        } else {
            0.0
        }
    }
}

fn compare_against_baseline(
    baseline: &BulkBaseline,
    current_results: &[FileResult],
) -> BaselineComparison {
    let current: BTreeMap<String, &FileResult> = current_results
        .iter()
        .map(|result| (fixture_key(&result.path), result))
        .collect();
    let mut regressions: Vec<Regression> = Vec::new();
    let mut improvements: Vec<String> = Vec::new();

    for (path, result) in &current {
        if result.outcome == Outcome::Panic {
            regressions.push(Regression {
                path: path.clone(),
                kind: RegressionKind::Panic,
                detail: result.detail.clone(),
            });
        }
    }

    for (path, baseline_outcome) in &baseline.fixtures {
        let current_outcome = current.get(path).map(|result| result.outcome);
        if *baseline_outcome == Outcome::Success {
            match current_outcome {
                Some(Outcome::Success) => {}
                Some(outcome) => regressions.push(Regression {
                    path: path.clone(),
                    kind: RegressionKind::PreviousSuccessFailed,
                    detail: format!("baseline success became {outcome:?}"),
                }),
                None => regressions.push(Regression {
                    path: path.clone(),
                    kind: RegressionKind::PreviousSuccessMissing,
                    detail: "baseline success is missing from the fixture corpus".to_string(),
                }),
            }
        } else if current_outcome == Some(Outcome::Success) {
            improvements.push(path.clone());
        }
    }

    let baseline_effective: BTreeSet<&str> = baseline
        .fixtures
        .iter()
        .filter(|(_, outcome)| **outcome != Outcome::ExpectedError)
        .map(|(path, _)| path.as_str())
        .collect();
    let baseline_success = baseline
        .fixtures
        .iter()
        .filter(|(path, outcome)| {
            baseline_effective.contains(path.as_str()) && **outcome == Outcome::Success
        })
        .count();
    let current_success = baseline_effective
        .iter()
        .filter(|path| {
            current
                .get(**path)
                .is_some_and(|result| result.outcome == Outcome::Success)
        })
        .count();
    let denominator = baseline_effective.len();
    let baseline_success_rate = percentage(baseline_success, denominator);
    let current_success_rate = percentage(current_success, denominator);
    if current_success_rate + f64::EPSILON < baseline_success_rate {
        regressions.push(Regression {
            path: "(aggregate)".to_string(),
            kind: RegressionKind::SuccessRateDecreased,
            detail: format!(
                "baseline fixture success rate decreased from {baseline_success_rate:.2}% to {current_success_rate:.2}%"
            ),
        });
    }

    let new_files = current
        .keys()
        .filter(|path| !baseline.fixtures.contains_key(*path))
        .cloned()
        .collect();

    BaselineComparison {
        baseline_success_rate,
        current_success_rate,
        regressions,
        improvements,
        new_files,
    }
}

fn percentage(numerator: usize, denominator: usize) -> f64 {
    if denominator == 0 {
        0.0
    } else {
        (numerator as f64 / denominator as f64) * 100.0
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixtures_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures")
}

fn baseline_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/bulk_conversion_baseline.json")
}

fn report_dir() -> PathBuf {
    std::env::var_os("BULK_REPORT_DIR").map_or_else(
        || PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../target/bulk-conversion"),
        PathBuf::from,
    )
}

fn fixture_key(path: &Path) -> String {
    let relative = path.strip_prefix(fixtures_dir()).unwrap_or(path);
    relative.to_string_lossy().replace('\\', "/")
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
    let expected = is_expected_error(path);

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
            outcome: if expected {
                Outcome::ExpectedError
            } else {
                Outcome::Error
            },
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
    let all_files = discover_files(&dir, extension);
    let (denied, files): (Vec<_>, Vec<_>) = all_files.into_iter().partition(|p| is_denylisted(p));
    let skipped = denied.len();

    println!("\n{}", "=".repeat(60));
    println!(
        "  Bulk {format_name} conversion: {} files ({skipped} denylisted, skipped)",
        files.len() + skipped
    );
    println!("{}\n", "=".repeat(60));

    if skipped > 0 {
        for p in &denied {
            println!(
                "  SKIP: {}",
                p.file_name()
                    .map(|f| f.to_string_lossy().to_string())
                    .unwrap_or_default()
            );
        }
        println!();
    }

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
            Outcome::ExpectedError => println!("EXPECTED ERROR: {}", result.detail),
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
    let expected_error = results
        .iter()
        .filter(|r| r.outcome == Outcome::ExpectedError)
        .count();
    let panic = results
        .iter()
        .filter(|r| r.outcome == Outcome::Panic)
        .count();

    let summary = Summary {
        format: format_name,
        total: files.len(),
        skipped,
        success,
        error,
        expected_error,
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
        "Total: {} | Skipped: {} | Success: {} | Error: {} | Expected Error: {} | Panic: {}",
        summary.total,
        summary.skipped,
        summary.success,
        summary.error,
        summary.expected_error,
        summary.panic
    )
    .unwrap();
    writeln!(report, "Success rate: {:.1}%", summary.success_rate()).unwrap();
    writeln!(report).unwrap();

    // List panics first (most critical)
    let panics: Vec<_> = results
        .iter()
        .filter(|r| r.outcome == Outcome::Panic)
        .collect();
    if !panics.is_empty() {
        writeln!(report, "## PANICS ({} files)", panics.len()).unwrap();
        for r in &panics {
            writeln!(report, "  - {} :: {}", fixture_key(&r.path), r.detail).unwrap();
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
            writeln!(report, "  - {} :: {}", fixture_key(&r.path), r.detail).unwrap();
        }
        writeln!(report).unwrap();
    }

    // List expected errors
    let expected: Vec<_> = results
        .iter()
        .filter(|r| r.outcome == Outcome::ExpectedError)
        .collect();
    if !expected.is_empty() {
        writeln!(
            report,
            "## EXPECTED ERRORS ({} files, excluded from success rate)",
            expected.len()
        )
        .unwrap();
        for r in &expected {
            writeln!(report, "  - {} :: {}", fixture_key(&r.path), r.detail).unwrap();
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
            writeln!(report, "  - {} :: {}", fixture_key(&r.path), r.detail).unwrap();
        }
        writeln!(report).unwrap();
    }

    report
}

fn format_baseline_comparison(comparison: &BaselineComparison) -> String {
    let mut report = String::new();
    writeln!(report, "# Baseline Comparison").unwrap();
    writeln!(
        report,
        "Baseline success rate: {:.2}% | Current success rate: {:.2}%",
        comparison.baseline_success_rate, comparison.current_success_rate
    )
    .unwrap();
    writeln!(
        report,
        "Gate: {}",
        if comparison.passed() { "PASS" } else { "FAIL" }
    )
    .unwrap();
    writeln!(report).unwrap();

    if !comparison.regressions.is_empty() {
        writeln!(report, "## Regressions ({})", comparison.regressions.len()).unwrap();
        for regression in &comparison.regressions {
            writeln!(
                report,
                "- `{}` ({:?}): {}",
                regression.path, regression.kind, regression.detail
            )
            .unwrap();
        }
        writeln!(report).unwrap();
    }

    if !comparison.improvements.is_empty() {
        writeln!(
            report,
            "## Improvements ({})",
            comparison.improvements.len()
        )
        .unwrap();
        for path in &comparison.improvements {
            writeln!(report, "- `{path}`").unwrap();
        }
        writeln!(report).unwrap();
    }

    if !comparison.new_files.is_empty() {
        writeln!(report, "## New Files ({})", comparison.new_files.len()).unwrap();
        for path in &comparison.new_files {
            writeln!(report, "- `{path}`").unwrap();
        }
        writeln!(report).unwrap();
    }

    report
}

/// Print summary table to stdout.
fn print_summary_table(summaries: &[&Summary]) {
    println!("\n{}", "=".repeat(72));
    println!("  BULK CONVERSION SUMMARY");
    println!("{}", "=".repeat(72));
    println!(
        "{:<8} {:>6} {:>8} {:>8} {:>6} {:>9} {:>6} {:>8}",
        "Format", "Total", "Skipped", "Success", "Error", "Expected", "Panic", "Rate"
    );
    println!("{:-<72}", "");

    let mut total_all = 0;
    let mut skipped_all = 0;
    let mut success_all = 0;
    let mut error_all = 0;
    let mut expected_all = 0;
    let mut panic_all = 0;

    for s in summaries {
        println!(
            "{:<8} {:>6} {:>8} {:>8} {:>6} {:>9} {:>6} {:>7.1}%",
            s.format,
            s.total,
            s.skipped,
            s.success,
            s.error,
            s.expected_error,
            s.panic,
            s.success_rate()
        );
        total_all += s.total;
        skipped_all += s.skipped;
        success_all += s.success;
        error_all += s.error;
        expected_all += s.expected_error;
        panic_all += s.panic;
    }

    let eff_total = total_all - expected_all;
    let rate_all = if eff_total > 0 {
        (success_all as f64 / eff_total as f64) * 100.0
    } else {
        0.0
    };
    println!("{:-<72}", "");
    println!(
        "{:<8} {:>6} {:>8} {:>8} {:>6} {:>9} {:>6} {:>7.1}%",
        "TOTAL", total_all, skipped_all, success_all, error_all, expected_all, panic_all, rate_all
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

fn load_baseline() -> BulkBaseline {
    let path = baseline_path();
    let data = std::fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("read bulk baseline {}: {error}", path.display()));
    let baseline: BulkBaseline = serde_json::from_str(&data)
        .unwrap_or_else(|error| panic!("parse bulk baseline {}: {error}", path.display()));
    assert_eq!(
        baseline.version, 1,
        "unsupported bulk baseline version {}",
        baseline.version
    );
    baseline
}

fn baseline_from_results(results: &[FileResult]) -> BulkBaseline {
    BulkBaseline::from_entries(
        results
            .iter()
            .map(|result| (fixture_key(&result.path), result.outcome)),
    )
}

fn update_baseline(results: &[FileResult]) {
    let baseline = baseline_from_results(results);
    let json = serde_json::to_string_pretty(&baseline).expect("serialize bulk baseline");
    let path = baseline_path();
    std::fs::write(&path, format!("{json}\n"))
        .unwrap_or_else(|error| panic!("write bulk baseline {}: {error}", path.display()));
    println!("Baseline updated: {}", path.display());
}

fn write_gate_reports(
    results: &[FileResult],
    summaries: &[&Summary],
    comparison: &BaselineComparison,
) {
    let output_dir = report_dir();
    std::fs::create_dir_all(&output_dir).unwrap_or_else(|error| {
        panic!("create report directory {}: {error}", output_dir.display())
    });

    let mut markdown = String::new();
    for (format_results, summary) in split_results_by_format(results, summaries) {
        writeln!(markdown, "{}", format_report(format_results, summary)).unwrap();
    }
    writeln!(markdown, "{}", format_baseline_comparison(comparison)).unwrap();
    std::fs::write(output_dir.join("report.md"), markdown).expect("write bulk report.md");

    let machine_report = MachineReport {
        summaries: summaries
            .iter()
            .map(|summary| MachineSummary {
                format: summary.format,
                total: summary.total,
                skipped: summary.skipped,
                success: summary.success,
                error: summary.error,
                expected_error: summary.expected_error,
                panic: summary.panic,
                success_rate: summary.success_rate(),
            })
            .collect(),
        files: results
            .iter()
            .map(|result| MachineFileResult {
                path: fixture_key(&result.path),
                outcome: result.outcome,
                detail: result.detail.clone(),
            })
            .collect(),
        comparison,
    };
    let json = serde_json::to_string_pretty(&machine_report).expect("serialize bulk report");
    std::fs::write(output_dir.join("report.json"), format!("{json}\n"))
        .expect("write bulk report.json");
    println!("Gate reports written to: {}", output_dir.display());
}

fn split_results_by_format<'a>(
    results: &'a [FileResult],
    summaries: &'a [&Summary],
) -> Vec<(&'a [FileResult], &'a Summary)> {
    let mut offset: usize = 0;
    summaries
        .iter()
        .map(|summary| {
            let end = offset + summary.total;
            let slice = &results[offset..end];
            offset = end;
            (slice, *summary)
        })
        .collect()
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

#[test]
#[ignore]
fn test_bulk_regression_gate() {
    let (docx_results, docx_summary) = run_bulk_test("DOCX", "docx", Format::Docx);
    let (pptx_results, pptx_summary) = run_bulk_test("PPTX", "pptx", Format::Pptx);
    let (xlsx_results, xlsx_summary) = run_bulk_test("XLSX", "xlsx", Format::Xlsx);
    let all_results: Vec<FileResult> = docx_results
        .into_iter()
        .chain(pptx_results)
        .chain(xlsx_results)
        .collect();
    let summaries = [&docx_summary, &pptx_summary, &xlsx_summary];

    if std::env::var("UPDATE_BULK_BASELINE").as_deref() == Ok("1") {
        update_baseline(&all_results);
    }

    let baseline = load_baseline();
    let comparison = compare_against_baseline(&baseline, &all_results);
    print_summary_table(&summaries);
    println!("{}", format_baseline_comparison(&comparison));
    write_gate_reports(&all_results, &summaries, &comparison);

    assert!(
        comparison.passed(),
        "{} bulk fixture regression(s) detected; inspect {}/report.md",
        comparison.regressions.len(),
        report_dir().display()
    );
}

/// Verifies that `is_denylisted` correctly identifies files on the denylist
/// and does not reject normal files.
#[test]
fn test_denylist_filtering() {
    // Every entry in DENYLIST should be recognized regardless of parent directory
    for name in DENYLIST {
        let path = PathBuf::from(format!("tests/fixtures/any/dir/{name}"));
        assert!(
            is_denylisted(&path),
            "Expected {name} to be denylisted, but it was not"
        );
    }

    // Denylist should cover all three formats
    let docx_count = DENYLIST.iter().filter(|n| n.ends_with(".docx")).count();
    let pptx_count = DENYLIST.iter().filter(|n| n.ends_with(".pptx")).count();
    let xlsx_count = DENYLIST.iter().filter(|n| n.ends_with(".xlsx")).count();
    assert!(
        docx_count >= 14,
        "Expected ≥14 DOCX entries, got {docx_count}"
    );
    assert!(
        pptx_count >= 10,
        "Expected ≥10 PPTX entries, got {pptx_count}"
    );
    assert!(
        xlsx_count >= 15,
        "Expected ≥15 XLSX entries, got {xlsx_count}"
    );

    // Normal files must not be denylisted
    let normal = PathBuf::from("tests/fixtures/xlsx/poi/sample.xlsx");
    assert!(
        !is_denylisted(&normal),
        "Normal file should not be denylisted"
    );

    // A file whose name contains a denylisted name as substring must not match
    let substring = PathBuf::from("tests/fixtures/xlsx/poi/not-poc-xmlbomb.xlsx.bak");
    assert!(
        !is_denylisted(&substring),
        "Substring match should not trigger denylist"
    );
}

/// Verifies that `is_expected_error` identifies fixtures rejected by design.
#[test]
fn test_expected_error_filtering() {
    // Every entry in EXPECTED_ERRORS should be recognized
    for name in EXPECTED_ERRORS {
        let path = PathBuf::from(format!("tests/fixtures/docx/poi/{name}"));
        assert!(
            is_expected_error(&path),
            "Expected {name} to be in expected-error list, but it was not"
        );
    }

    // Normal files must not be expected-error
    let normal = PathBuf::from("tests/fixtures/docx/poi/sample.docx");
    assert!(
        !is_expected_error(&normal),
        "Normal file should not be in expected-error list"
    );
}

/// Verifies that expected errors do not count against the success rate.
#[test]
fn test_summary_success_rate_excludes_expected_errors() {
    let summary = Summary {
        format: "TEST",
        total: 10,
        skipped: 0,
        success: 7,
        error: 1,
        expected_error: 2,
        panic: 0,
    };
    // effective_total = 10 - 2 = 8, success_rate = 7/8 = 87.5%
    assert_eq!(summary.effective_total(), 8);
    assert!((summary.success_rate() - 87.5).abs() < 0.01);
}

#[test]
fn baseline_gate_rejects_previously_successful_fixture_errors() {
    let baseline = baseline(&[
        ("docx/stable.docx", Outcome::Success),
        ("docx/known-error.docx", Outcome::Error),
    ]);
    let current = results(&[
        ("docx/stable.docx", Outcome::Error),
        ("docx/known-error.docx", Outcome::Error),
    ]);

    let comparison = compare_against_baseline(&baseline, &current);

    assert!(!comparison.passed());
    assert!(comparison.regressions.iter().any(|regression| {
        regression.path == "docx/stable.docx"
            && regression.kind == RegressionKind::PreviousSuccessFailed
    }));
}

#[test]
fn baseline_gate_rejects_missing_successes_and_any_panic() {
    let baseline = baseline(&[("pptx/stable.pptx", Outcome::Success)]);
    let current = results(&[("xlsx/new.xlsx", Outcome::Panic)]);

    let comparison = compare_against_baseline(&baseline, &current);

    assert!(!comparison.passed());
    assert!(comparison.regressions.iter().any(|regression| {
        regression.path == "pptx/stable.pptx"
            && regression.kind == RegressionKind::PreviousSuccessMissing
    }));
    assert!(comparison.regressions.iter().any(|regression| {
        regression.path == "xlsx/new.xlsx" && regression.kind == RegressionKind::Panic
    }));
}

#[test]
fn baseline_gate_allows_new_conversion_errors_and_records_improvements() {
    let baseline = baseline(&[
        ("xlsx/stable.xlsx", Outcome::Success),
        ("xlsx/known-error.xlsx", Outcome::Error),
    ]);
    let current = results(&[
        ("xlsx/stable.xlsx", Outcome::Success),
        ("xlsx/known-error.xlsx", Outcome::Success),
        ("xlsx/new-error.xlsx", Outcome::Error),
    ]);

    let comparison = compare_against_baseline(&baseline, &current);

    assert!(comparison.passed());
    assert_eq!(comparison.new_files, vec!["xlsx/new-error.xlsx"]);
    assert_eq!(comparison.improvements, vec!["xlsx/known-error.xlsx"]);
    assert_eq!(comparison.baseline_success_rate, 50.0);
    assert_eq!(comparison.current_success_rate, 100.0);
}

#[test]
fn bulk_report_uses_fixture_relative_paths() {
    let fixture = fixtures_dir().join("pptx/libreoffice/example.pptx");
    let results = vec![FileResult {
        path: fixture,
        outcome: Outcome::Error,
        detail: "example error".to_string(),
    }];
    let summary = Summary {
        format: "PPTX",
        total: 1,
        skipped: 0,
        success: 0,
        error: 1,
        expected_error: 0,
        panic: 0,
    };

    let report = format_report(&results, &summary);

    assert!(report.contains("pptx/libreoffice/example.pptx"));
    assert!(!report.contains(env!("CARGO_MANIFEST_DIR")));
}

fn baseline(entries: &[(&str, Outcome)]) -> BulkBaseline {
    BulkBaseline::from_entries(
        entries
            .iter()
            .map(|(path, outcome)| ((*path).to_string(), *outcome)),
    )
}

fn results(entries: &[(&str, Outcome)]) -> Vec<FileResult> {
    entries
        .iter()
        .map(|(path, outcome)| FileResult {
            path: PathBuf::from(path),
            outcome: *outcome,
            detail: String::new(),
        })
        .collect()
}

/// Asserts that the overall conversion success rate meets the 70% target (US-205).
///
/// This test runs all formats and verifies the combined success rate is at or
/// above 70%. Expected errors (encrypted files) are excluded from the
/// denominator.
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
    let expected: usize = summaries.iter().map(|s| s.expected_error).sum();
    let eff_total = total - expected;
    let rate = if eff_total > 0 {
        (success as f64 / eff_total as f64) * 100.0
    } else {
        0.0
    };

    // Per-format rates
    for s in &summaries {
        println!(
            "{}: {}/{} ({:.1}%)",
            s.format,
            s.success,
            s.effective_total(),
            s.success_rate()
        );
    }
    println!("Overall: {success}/{eff_total} ({rate:.1}%)");

    assert!(
        rate >= TARGET_RATE,
        "Overall success rate {rate:.1}% is below the {TARGET_RATE}% target. \
         {success}/{eff_total} files converted successfully."
    );
}
