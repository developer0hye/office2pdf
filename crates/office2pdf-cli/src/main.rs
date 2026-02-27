use std::path::{Path, PathBuf};
use std::process;

use anyhow::{Context, Result};
use clap::Parser;
use office2pdf::config::{ConvertOptions, PaperSize, PdfStandard, SlideRange};

#[derive(Parser)]
#[command(
    name = "office2pdf",
    version,
    about = "Convert DOCX, XLSX, PPTX to PDF"
)]
struct Cli {
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
fn convert_single(input: &Path, output: &Path, options: &ConvertOptions) -> Result<()> {
    let result = office2pdf::convert_with_options(input, options)
        .with_context(|| format!("converting {:?}", input))?;

    for warning in &result.warnings {
        eprintln!("Warning: {warning}");
    }

    std::fs::write(output, result.pdf)
        .with_context(|| format!("writing output to {:?}", output))?;

    Ok(())
}

/// Convert multiple files independently, collecting results.
fn convert_batch(
    inputs: &[PathBuf],
    outdir: Option<&Path>,
    options: &ConvertOptions,
) -> BatchResult {
    let mut result = BatchResult {
        succeeded: Vec::new(),
        failed: Vec::new(),
    };

    for input in inputs {
        let output_path = determine_output_path(input, None, outdir);
        match convert_single(input, &output_path, options) {
            Ok(()) => {
                println!("Converted: {:?} -> {:?}", input, output_path);
                result.succeeded.push((input.clone(), output_path));
            }
            Err(err) => {
                eprintln!("Failed: {:?}: {err:#}", input);
                result.failed.push((input.clone(), format!("{err:#}")));
            }
        }
    }

    result
}

fn run() -> Result<()> {
    let cli = Cli::parse();

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

    // Single file with explicit --output
    if let Some(output) = cli.output {
        let input = &cli.inputs[0];
        convert_single(input, &output, &options)?;
        println!("Converted: {:?} -> {:?}", input, output);
        return Ok(());
    }

    // Batch conversion (works for 1 or many files)
    let result = convert_batch(&cli.inputs, cli.outdir.as_deref(), &options);

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
        let result = convert_batch(&inputs, None, &options);

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
        let result = convert_batch(&inputs, None, &options);

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
        let result = convert_batch(&inputs, Some(&outdir), &options);

        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 0);
        assert!(outdir.join("report.pdf").exists());
        assert!(outdir.join("memo.pdf").exists());
        // Original directory should NOT have PDFs
        assert!(!dir.join("report.pdf").exists());
        assert!(!dir.join("memo.pdf").exists());

        let _ = std::fs::remove_dir_all(&dir);
    }
}
