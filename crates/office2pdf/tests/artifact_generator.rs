//! Artifact generator for visual comparison of classified fixtures.
//!
//! Converts each input file, renders both output and ground truth PDFs to PNGs,
//! extracts text, and saves all artifacts for analysis. Does NOT compute quality
//! scores — the actual analysis is done externally (by Ralph / Claude) by reading
//! the generated images and text files.
//!
//! Requires `pdftoppm` and `pdftotext` from poppler-utils.
//!
//! Run with:
//!   cargo test -p office2pdf --test artifact_generator -- --ignored --nocapture
//!
//! Environment variables:
//!   VISUAL_DPI — rendering DPI (default: 150)
//!
//! Output structure (per file):
//!   tests/classified_fixtures/_work/<safe_name>/
//!     output-01.png, output-02.png, ...   — rendered output PDF pages
//!     gt-01.png, gt-02.png, ...           — rendered ground truth PDF pages
//!     output.txt                          — extracted text from output PDF
//!     gt.txt                              — extracted text from ground truth PDF
//!
//!   tests/classified_fixtures/_work/report.json — factual summary (page counts, text lengths)

mod common;

use std::path::PathBuf;

use office2pdf::config::{ConvertOptions, Format};

// ── Report structures ────────────────────────────────────────────────────────

#[derive(serde::Serialize)]
struct Report {
    dpi: u32,
    total_files: usize,
    converted: usize,
    failed_conversion: usize,
    files: Vec<FileResult>,
}

#[derive(serde::Serialize, Clone)]
struct FileResult {
    input: String,
    format: String,
    status: String,
    page_count_output: Option<u32>,
    page_count_gt: Option<u32>,
    text_len_output: Option<usize>,
    text_len_gt: Option<usize>,
    work_dir: String,
    output_pngs: Vec<String>,
    gt_pngs: Vec<String>,
}

// ── Main test ────────────────────────────────────────────────────────────────

#[test]
#[ignore]
fn test_visual_comparison_all() {
    let dpi: u32 = std::env::var("VISUAL_DPI")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(150);

    assert!(
        common::is_pdftoppm_available(),
        "pdftoppm (poppler-utils) is required but not found"
    );
    assert!(
        common::is_pdftotext_available(),
        "pdftotext (poppler-utils) is required but not found"
    );

    let fixtures_dir =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/classified_fixtures");
    let pairs_json =
        std::fs::read_to_string(fixtures_dir.join("pairs.json")).expect("read pairs.json");

    #[derive(serde::Deserialize)]
    struct Pair {
        input: String,
        ground_truth: String,
    }

    let pairs: Vec<Pair> = serde_json::from_str(&pairs_json).expect("parse pairs.json");
    let work_dir = fixtures_dir.join("_work");
    std::fs::create_dir_all(&work_dir).expect("create _work dir");

    println!("\n===== Artifact Generation =====");
    println!("DPI: {dpi} | Files: {}\n", pairs.len());

    let mut results: Vec<FileResult> = Vec::new();

    for (idx, pair) in pairs.iter().enumerate() {
        let input_path = fixtures_dir.join("input").join(&pair.input);
        let gt_path = fixtures_dir.join("gt").join(&pair.ground_truth);

        let short_name = pair.input.rsplit('/').next().unwrap_or(&pair.input);
        let ext = input_path
            .extension()
            .map(|e| e.to_string_lossy().to_string())
            .unwrap_or_default();
        print!("[{}/{}] {short_name} ... ", idx + 1, pairs.len());

        if !input_path.exists() || !gt_path.exists() {
            let missing = if !input_path.exists() {
                "input"
            } else {
                "ground truth"
            };
            println!("SKIP ({missing} not found)");
            results.push(FileResult {
                input: pair.input.clone(),
                format: ext,
                status: format!("skip_{missing}_missing"),
                page_count_output: None,
                page_count_gt: None,
                text_len_output: None,
                text_len_gt: None,
                work_dir: String::new(),
                output_pngs: vec![],
                gt_pngs: vec![],
            });
            continue;
        }

        // ── Convert ──────────────────────────────────────────────────────
        let input_bytes = std::fs::read(&input_path).expect("read input");
        let format = Format::from_extension(&ext).expect("known format");

        let convert_result =
            office2pdf::convert_bytes(&input_bytes, format, &ConvertOptions::default());

        let pdf_bytes = match convert_result {
            Ok(result) => result.pdf,
            Err(e) => {
                println!("FAIL ({e})");
                results.push(FileResult {
                    input: pair.input.clone(),
                    format: ext,
                    status: format!("conversion_error: {e}"),
                    page_count_output: None,
                    page_count_gt: None,
                    text_len_output: None,
                    text_len_gt: None,
                    work_dir: String::new(),
                    output_pngs: vec![],
                    gt_pngs: vec![],
                });
                continue;
            }
        };

        // ── Setup per-file work directory ────────────────────────────────
        let safe_name = short_name
            .replace(' ', "_")
            .replace(['(', ')', '[', ']', ','], "_");
        let file_work_dir = work_dir.join(&safe_name);
        // Clean previous run artifacts
        if file_work_dir.exists() {
            let _ = std::fs::remove_dir_all(&file_work_dir);
        }
        std::fs::create_dir_all(&file_work_dir).expect("create file work dir");

        // ── Render to PNGs (keep them for analysis) ──────────────────────
        let output_pngs =
            common::render_pdf_bytes_to_pngs(&pdf_bytes, &file_work_dir, "output", dpi);
        let gt_pngs = common::render_pdf_to_pngs(&gt_path, &file_work_dir, "gt", dpi);

        // ── Extract text ─────────────────────────────────────────────────
        let text_out = common::extract_text_from_pdf_bytes(&pdf_bytes, &file_work_dir);
        let text_gt = common::extract_text_from_pdf_file(&gt_path);

        std::fs::write(file_work_dir.join("output.txt"), &text_out).expect("write output.txt");
        std::fs::write(file_work_dir.join("gt.txt"), &text_gt).expect("write gt.txt");

        let pc_output = output_pngs.len() as u32;
        let pc_gt = gt_pngs.len() as u32;

        let png_names = |pngs: &[PathBuf]| -> Vec<String> {
            pngs.iter()
                .filter_map(|p| p.file_name().map(|f| f.to_string_lossy().into_owned()))
                .collect()
        };

        println!(
            "pages={pc_output}/{pc_gt}  text={}/{}chars",
            text_out.len(),
            text_gt.len()
        );

        results.push(FileResult {
            input: pair.input.clone(),
            format: ext,
            status: "ok".into(),
            page_count_output: Some(pc_output),
            page_count_gt: Some(pc_gt),
            text_len_output: Some(text_out.len()),
            text_len_gt: Some(text_gt.len()),
            work_dir: safe_name,
            output_pngs: png_names(&output_pngs),
            gt_pngs: png_names(&gt_pngs),
        });
    }

    // ── Write report.json ────────────────────────────────────────────────
    let converted = results.iter().filter(|r| r.status == "ok").count();
    let report = Report {
        dpi,
        total_files: results.len(),
        converted,
        failed_conversion: results.len() - converted,
        files: results,
    };

    let report_path = work_dir.join("report.json");
    let report_json = serde_json::to_string_pretty(&report).expect("serialize report");
    std::fs::write(&report_path, &report_json).expect("write report.json");

    println!("\nReport: {}", report_path.display());
    println!(
        "Converted: {}/{} | Artifacts in: {}",
        report.converted,
        report.total_files,
        work_dir.display()
    );
}
