use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::process;

use anyhow::{Context, Result};
use clap::Parser;
use office2pdf::config::{ConvertOptions, PaperSize, PdfStandard, SlideRange};
use office2pdf::pdf_ops;

#[cfg(feature = "server")]
mod metrics;
#[cfg(feature = "server")]
mod server;

#[derive(clap::Subcommand)]
enum Commands {
    /// Merge multiple PDF files into one
    Merge {
        /// Input PDF files to merge
        #[arg(required = true)]
        files: Vec<PathBuf>,
        /// Output file path
        #[arg(short, long, default_value = "merged.pdf")]
        output: PathBuf,
    },
    /// Split a PDF into parts by page ranges
    Split {
        /// Input PDF file
        input: PathBuf,
        /// Page ranges (e.g. "1-5,10-15")
        #[arg(long, required = true, value_delimiter = ',')]
        pages: Vec<String>,
        /// Output directory for split files
        #[arg(long, default_value = ".")]
        outdir: PathBuf,
    },
    #[cfg(feature = "server")]
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
    about = "Convert DOCX, XLSX, PPTX to PDF",
    subcommand_negates_reqs = true,
    args_conflicts_with_subcommands = true
)]
struct Cli {
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

    /// Produce tagged PDF with document structure tags for accessibility
    #[arg(long)]
    tagged: bool,

    /// Produce PDF/UA-1 compliant output for universal accessibility (implies --tagged)
    #[arg(long = "pdf-ua")]
    pdf_ua: bool,

    /// Enable streaming mode for large XLSX files (processes rows in chunks)
    #[arg(long)]
    streaming: bool,

    /// Chunk size (rows) for streaming mode (default: 1000)
    #[arg(long, default_value = None)]
    streaming_chunk_size: Option<usize>,

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

    let mut seen_warnings = HashSet::new();
    for warning in &result.warnings {
        let rendered = warning.to_string();
        if seen_warnings.insert(rendered.clone()) {
            eprintln!("Warning: {rendered}");
        }
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

/// Handle a CLI subcommand.
fn handle_command(cmd: Commands) -> Result<()> {
    match cmd {
        Commands::Merge { files, output } => {
            let inputs: Vec<Vec<u8>> = files
                .iter()
                .map(|f| std::fs::read(f).with_context(|| format!("reading {:?}", f)))
                .collect::<Result<_>>()?;

            let refs: Vec<&[u8]> = inputs.iter().map(|v| v.as_slice()).collect();
            let merged = pdf_ops::merge(&refs).map_err(|e| anyhow::anyhow!("{e}"))?;

            std::fs::write(&output, merged)
                .with_context(|| format!("writing output to {:?}", output))?;

            println!("Merged {} files -> {:?}", files.len(), output);
            Ok(())
        }
        Commands::Split {
            input,
            pages,
            outdir,
        } => {
            let data = std::fs::read(&input).with_context(|| format!("reading {:?}", input))?;

            let ranges: Vec<pdf_ops::PageRange> = pages
                .iter()
                .map(|s| {
                    pdf_ops::PageRange::parse(s)
                        .map_err(|e| anyhow::anyhow!("invalid page range '{s}': {e}"))
                })
                .collect::<Result<_>>()?;

            let parts = pdf_ops::split(&data, &ranges).map_err(|e| anyhow::anyhow!("{e}"))?;

            std::fs::create_dir_all(&outdir)
                .with_context(|| format!("creating output directory {:?}", outdir))?;

            let stem = input.file_stem().unwrap_or_default().to_string_lossy();

            for (i, (part, range)) in parts.iter().zip(ranges.iter()).enumerate() {
                let filename = format!("{}_pages_{}-{}.pdf", stem, range.start, range.end);
                let out_path = outdir.join(&filename);
                std::fs::write(&out_path, part)
                    .with_context(|| format!("writing {:?}", out_path))?;
                println!(
                    "Split part {} (pages {}-{}) -> {:?}",
                    i + 1,
                    range.start,
                    range.end,
                    out_path
                );
            }
            Ok(())
        }
        #[cfg(feature = "server")]
        Commands::Serve { host, port } => server::start_server(&host, port),
    }
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
    if let Some(cmd) = cli.command {
        return handle_command(cmd);
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
        tagged: cli.tagged,
        pdf_ua: cli.pdf_ua,
        streaming: cli.streaming,
        streaming_chunk_size: cli.streaming_chunk_size,
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
#[path = "main_tests.rs"]
mod tests;
