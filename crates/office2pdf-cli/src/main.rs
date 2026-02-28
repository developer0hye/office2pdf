use std::path::{Path, PathBuf};
use std::process;

use anyhow::{Context, Result};
use clap::Parser;
use office2pdf::config::{ConvertOptions, PaperSize, PdfStandard, SlideRange};

#[cfg(feature = "server")]
mod server;

#[cfg(feature = "server")]
#[derive(clap::Subcommand)]
enum Commands {
    /// Start an HTTP server for document conversion
    Serve {
        /// Host address to bind to
        #[arg(long, default_value = "127.0.0.1")]
        host: String,
        /// Port to listen on
        #[arg(long, default_value_t = 3000)]
        port: u16,
    },
}

#[derive(Parser)]
#[command(
    name = "office2pdf",
    version,
    about = "Convert DOCX, XLSX, PPTX to PDF"
)]
#[cfg_attr(
    feature = "server",
    command(subcommand_negates_reqs = true, args_conflicts_with_subcommands = true)
)]
struct Cli {
    #[cfg(feature = "server")]
    #[command(subcommand)]
    command: Option<Commands>,

    /// Input file paths (.docx, .xlsx, .pptx)
    #[arg(required = true)]
    inputs: Vec<PathBuf>,

    /// Output PDF file path (only valid with a single input file)
    #[arg(short, long, conflicts_with = "outdir")]
    output: Option<PathBuf>,

    /// Output directory for converted files
    #[arg(long)]
    outdir: Option<PathBuf>,

    /// XLSX sheet names to include (comma-separated, e.g. "Sheet1,Data")
    #[arg(long, value_delimiter = ',')]
    sheets: Option<Vec<String>>,

    /// PPTX slide range to include (e.g. "1-5" or "3")
    #[arg(long)]
    slides: Option<String>,

    /// Produce PDF/A-2b compliant output for archival purposes
    #[arg(long = "pdf-a")]
    pdf_a: bool,

    /// Paper size for output (a4, letter, legal)
    #[arg(long)]
    paper: Option<String>,

    /// Additional font directory to search (can be repeated)
    #[arg(long = "font-path")]
    font_path: Vec<PathBuf>,

    /// Force landscape orientation
    #[arg(long)]
    landscape: bool,

    /// Print per-stage timing metrics to stderr
    #[arg(long)]
    metrics: bool,

    /// Number of parallel conversion jobs (default: number of CPU cores)
    #[arg(short = 'j', long, default_value_t = 0)]
    jobs: usize,
}

/// Result of a batch conversion.
struct BatchResult {
    /// Successfully converted files: (input, output) pairs.
    succeeded: Vec<(PathBuf, PathBuf)>,
    /// Failed files: (input, error message) pairs.
    failed: Vec<(PathBuf, String)>,
}

fn main() {
    if let Err(err) = run() {
        eprintln!("Error: {err:#}");
        process::exit(1);
    }
}

/// Determine the output path for a given input file.
fn determine_output_path(input: &Path, output: Option<&Path>, outdir: Option<&Path>) -> PathBuf {
    if let Some(out) = output {
        out.to_path_buf()
    } else if let Some(dir) = outdir {
        let filename = input.file_name().unwrap_or_default();
        dir.join(filename).with_extension("pdf")
    } else {
        input.with_extension("pdf")
    }
}

/// Convert a single file and write the PDF output.
fn convert_single(
    input: &Path,
    output: &Path,
    options: &ConvertOptions,
    show_metrics: bool,
) -> Result<()> {
    let result = office2pdf::convert_with_options(input, options)
        .with_context(|| format!("converting {:?}", input))?;

    for warning in &result.warnings {
        eprintln!("Warning: {warning}");
    }

    if show_metrics && let Some(ref m) = result.metrics {
        eprintln!("--- Metrics: {:?} ---", input);
        eprintln!("  Parse:   {:?}", m.parse_duration);
        eprintln!("  Codegen: {:?}", m.codegen_duration);
        eprintln!("  Compile: {:?}", m.compile_duration);
        eprintln!("  Total:   {:?}", m.total_duration);
        eprintln!("  Input:   {} bytes", m.input_size_bytes);
        eprintln!("  Output:  {} bytes", m.output_size_bytes);
        eprintln!("  Pages:   {}", m.page_count);
    }

    std::fs::write(output, result.pdf)
        .with_context(|| format!("writing output to {:?}", output))?;

    Ok(())
}

/// Convert multiple files independently, collecting results.
///
/// When `jobs > 1` and there are multiple inputs, files are converted in
/// parallel using a rayon thread pool. `jobs == 0` means "use all available
/// CPU cores" (rayon's default).
fn convert_batch(
    inputs: &[PathBuf],
    outdir: Option<&Path>,
    options: &ConvertOptions,
    show_metrics: bool,
    jobs: usize,
) -> BatchResult {
    let convert_one = |input: &PathBuf| -> Result<(PathBuf, PathBuf), (PathBuf, String)> {
        let output_path = determine_output_path(input, None, outdir);
        match convert_single(input, &output_path, options, show_metrics) {
            Ok(()) => {
                println!("Converted: {:?} -> {:?}", input, output_path);
                Ok((input.clone(), output_path))
            }
            Err(err) => {
                eprintln!("Failed: {:?}: {err:#}", input);
                Err((input.clone(), format!("{err:#}")))
            }
        }
    };

    let effective_jobs = if jobs == 0 {
        std::thread::available_parallelism()
            .map(|n| n.get())
            .unwrap_or(1)
    } else {
        jobs
    };

    let results: Vec<_> = if effective_jobs > 1 && inputs.len() > 1 {
        use rayon::iter::{IntoParallelRefIterator, ParallelIterator};
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(effective_jobs)
            .build()
            .expect("failed to create rayon thread pool");
        pool.install(|| inputs.par_iter().map(convert_one).collect())
    } else {
        inputs.iter().map(convert_one).collect()
    };

    let mut batch = BatchResult {
        succeeded: Vec::new(),
        failed: Vec::new(),
    };
    for r in results {
        match r {
            Ok(pair) => batch.succeeded.push(pair),
            Err(pair) => batch.failed.push(pair),
        }
    }
    batch
}

fn run() -> Result<()> {
    let cli = Cli::parse();

    // Handle subcommands
    #[cfg(feature = "server")]
    if let Some(Commands::Serve { host, port }) = cli.command {
        return server::start_server(&host, port);
    }

    // --output is only valid with a single input file
    if cli.inputs.len() > 1 && cli.output.is_some() {
        anyhow::bail!("--output cannot be used with multiple input files; use --outdir instead");
    }

    let slide_range = cli
        .slides
        .map(|s| SlideRange::parse(&s))
        .transpose()
        .map_err(|e| anyhow::anyhow!("invalid --slides value: {e}"))?;

    let pdf_standard = if cli.pdf_a {
        Some(PdfStandard::PdfA2b)
    } else {
        None
    };

    let paper_size = cli
        .paper
        .map(|s| PaperSize::parse(&s))
        .transpose()
        .map_err(|e| anyhow::anyhow!("invalid --paper value: {e}"))?;

    let landscape = if cli.landscape { Some(true) } else { None };

    let options = ConvertOptions {
        sheet_names: cli.sheets,
        slide_range,
        pdf_standard,
        paper_size,
        font_paths: cli.font_path,
        landscape,
    };

    // Create outdir if specified and doesn't exist
    if let Some(ref outdir) = cli.outdir {
        std::fs::create_dir_all(outdir)
            .with_context(|| format!("creating output directory {:?}", outdir))?;
    }

    let show_metrics = cli.metrics;

    // Single file with explicit --output
    if let Some(output) = cli.output {
        let input = &cli.inputs[0];
        convert_single(input, &output, &options, show_metrics)?;
        println!("Converted: {:?} -> {:?}", input, output);
        return Ok(());
    }

    // Batch conversion (works for 1 or many files)
    let result = convert_batch(
        &cli.inputs,
        cli.outdir.as_deref(),
        &options,
        show_metrics,
        cli.jobs,
    );

    // Print summary when there are multiple files
    let total = result.succeeded.len() + result.failed.len();
    if total > 1 {
        println!(
            "\nSummary: {} succeeded, {} failed (out of {} files)",
            result.succeeded.len(),
            result.failed.len(),
            total
        );
        if !result.failed.is_empty() {
            println!("Failed files:");
            for (path, err) in &result.failed {
                println!("  {:?}: {err}", path);
            }
        }
    }

    if !result.failed.is_empty() {
        process::exit(1);
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn make_test_docx() -> Vec<u8> {
        let docx = docx_rs::Docx::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello batch")),
        );
        let mut buf = Cursor::new(Vec::new());
        docx.build().pack(&mut buf).unwrap();
        buf.into_inner()
    }

    // --- Unit tests for determine_output_path ---

    #[test]
    fn test_determine_output_path_default() {
        let input = PathBuf::from("/tmp/report.docx");
        let result = determine_output_path(&input, None, None);
        assert_eq!(result, PathBuf::from("/tmp/report.pdf"));
    }

    #[test]
    fn test_determine_output_path_with_output() {
        let input = PathBuf::from("/tmp/report.docx");
        let output = PathBuf::from("/custom/output.pdf");
        let result = determine_output_path(&input, Some(&output), None);
        assert_eq!(result, PathBuf::from("/custom/output.pdf"));
    }

    #[test]
    fn test_determine_output_path_with_outdir() {
        let input = PathBuf::from("/tmp/report.docx");
        let outdir = PathBuf::from("/output");
        let result = determine_output_path(&input, None, Some(&outdir));
        assert_eq!(result, PathBuf::from("/output/report.pdf"));
    }

    #[test]
    fn test_determine_output_path_outdir_replaces_extension() {
        let input = PathBuf::from("/data/slides.pptx");
        let outdir = PathBuf::from("/pdfs");
        let result = determine_output_path(&input, None, Some(&outdir));
        assert_eq!(result, PathBuf::from("/pdfs/slides.pdf"));
    }

    // --- Integration tests for batch conversion ---

    #[test]
    fn test_batch_convert_multiple_files() {
        let dir = std::env::temp_dir().join("office2pdf_batch_test_multi");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let file1 = dir.join("doc1.docx");
        let file2 = dir.join("doc2.docx");
        std::fs::write(&file1, &docx_data).unwrap();
        std::fs::write(&file2, &docx_data).unwrap();

        let inputs = vec![file1, file2];
        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 1);

        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 0);
        assert!(dir.join("doc1.pdf").exists());
        assert!(dir.join("doc2.pdf").exists());

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_partial_failure() {
        let dir = std::env::temp_dir().join("office2pdf_batch_test_fail");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let file1 = dir.join("good.docx");
        let file2 = dir.join("bad.txt");
        std::fs::write(&file1, &docx_data).unwrap();
        std::fs::write(&file2, b"not a valid document").unwrap();

        let inputs = vec![file1, file2.clone()];
        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 1);

        assert_eq!(result.succeeded.len(), 1);
        assert_eq!(result.failed.len(), 1);
        assert!(dir.join("good.pdf").exists());
        assert_eq!(result.failed[0].0, file2);

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_with_outdir() {
        let dir = std::env::temp_dir().join("office2pdf_batch_test_outdir");
        let outdir = dir.join("output");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        std::fs::create_dir_all(&outdir).unwrap();

        let docx_data = make_test_docx();
        let file1 = dir.join("report.docx");
        let file2 = dir.join("memo.docx");
        std::fs::write(&file1, &docx_data).unwrap();
        std::fs::write(&file2, &docx_data).unwrap();

        let inputs = vec![file1, file2];
        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, Some(&outdir), &options, false, 1);

        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 0);
        assert!(outdir.join("report.pdf").exists());
        assert!(outdir.join("memo.pdf").exists());
        // Original directory should NOT have PDFs
        assert!(!dir.join("report.pdf").exists());
        assert!(!dir.join("memo.pdf").exists());

        let _ = std::fs::remove_dir_all(&dir);
    }

    // --- Parallel batch conversion tests ---

    #[test]
    fn test_batch_convert_parallel_jobs_2() {
        let dir = std::env::temp_dir().join("office2pdf_parallel_test_j2");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let inputs: Vec<PathBuf> = (0..4)
            .map(|i| {
                let path = dir.join(format!("doc{i}.docx"));
                std::fs::write(&path, &docx_data).unwrap();
                path
            })
            .collect();

        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 2);

        assert_eq!(result.succeeded.len(), 4);
        assert_eq!(result.failed.len(), 0);
        for i in 0..4 {
            let pdf_path = dir.join(format!("doc{i}.pdf"));
            assert!(pdf_path.exists(), "doc{i}.pdf should exist");
            let pdf_bytes = std::fs::read(&pdf_path).unwrap();
            assert!(pdf_bytes.len() > 100, "PDF should have real content");
        }

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_parallel_partial_failure() {
        let dir = std::env::temp_dir().join("office2pdf_parallel_fail_test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let good = dir.join("good.docx");
        let bad = dir.join("bad.txt");
        std::fs::write(&good, &docx_data).unwrap();
        std::fs::write(&bad, b"not a valid document").unwrap();

        let inputs = vec![good, bad.clone()];
        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 2);

        assert_eq!(result.succeeded.len(), 1);
        assert_eq!(result.failed.len(), 1);
        assert!(dir.join("good.pdf").exists());
        assert_eq!(result.failed[0].0, bad);

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_parallel_with_outdir() {
        let dir = std::env::temp_dir().join("office2pdf_parallel_outdir_test");
        let outdir = dir.join("output");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        std::fs::create_dir_all(&outdir).unwrap();

        let docx_data = make_test_docx();
        let inputs: Vec<PathBuf> = (0..3)
            .map(|i| {
                let path = dir.join(format!("file{i}.docx"));
                std::fs::write(&path, &docx_data).unwrap();
                path
            })
            .collect();

        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, Some(&outdir), &options, false, 2);

        assert_eq!(result.succeeded.len(), 3);
        assert_eq!(result.failed.len(), 0);
        for i in 0..3 {
            assert!(outdir.join(format!("file{i}.pdf")).exists());
            // Original directory should NOT have PDFs
            assert!(!dir.join(format!("file{i}.pdf")).exists());
        }

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_single_file_with_jobs() {
        // Single file should work fine even with jobs > 1
        let dir = std::env::temp_dir().join("office2pdf_parallel_single_test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let input = dir.join("single.docx");
        std::fs::write(&input, &docx_data).unwrap();

        let inputs = vec![input];
        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 4);

        assert_eq!(result.succeeded.len(), 1);
        assert_eq!(result.failed.len(), 0);
        assert!(dir.join("single.pdf").exists());

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_batch_convert_sequential_jobs_1() {
        // jobs=1 should use sequential path
        let dir = std::env::temp_dir().join("office2pdf_sequential_test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let inputs: Vec<PathBuf> = (0..3)
            .map(|i| {
                let path = dir.join(format!("seq{i}.docx"));
                std::fs::write(&path, &docx_data).unwrap();
                path
            })
            .collect();

        let options = ConvertOptions::default();
        let result = convert_batch(&inputs, None, &options, false, 1);

        assert_eq!(result.succeeded.len(), 3);
        assert_eq!(result.failed.len(), 0);

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_convert_single_with_metrics() {
        let dir = std::env::temp_dir().join("office2pdf_metrics_test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();

        let docx_data = make_test_docx();
        let input = dir.join("report.docx");
        let output = dir.join("report.pdf");
        std::fs::write(&input, &docx_data).unwrap();

        let options = ConvertOptions::default();
        // Should succeed with metrics=true (metrics printed to stderr)
        convert_single(&input, &output, &options, true).unwrap();
        assert!(output.exists());

        let _ = std::fs::remove_dir_all(&dir);
    }
}
