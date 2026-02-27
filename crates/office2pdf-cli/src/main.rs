use std::path::PathBuf;
use std::process;

use anyhow::{Context, Result};
use clap::Parser;
use office2pdf::config::{ConvertOptions, PdfStandard, SlideRange};

#[derive(Parser)]
#[command(
    name = "office2pdf",
    version,
    about = "Convert DOCX, XLSX, PPTX to PDF"
)]
struct Cli {
    /// Input file path (.docx, .xlsx, .pptx)
    input: PathBuf,

    /// Output PDF file path (default: input with .pdf extension)
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// XLSX sheet names to include (comma-separated, e.g. "Sheet1,Data")
    #[arg(long, value_delimiter = ',')]
    sheets: Option<Vec<String>>,

    /// PPTX slide range to include (e.g. "1-5" or "3")
    #[arg(long)]
    slides: Option<String>,

    /// Produce PDF/A-2b compliant output for archival purposes
    #[arg(long = "pdf-a")]
    pdf_a: bool,
}

fn main() {
    if let Err(err) = run() {
        eprintln!("Error: {err:#}");
        process::exit(1);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();

    let output = cli
        .output
        .unwrap_or_else(|| cli.input.with_extension("pdf"));

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

    let options = ConvertOptions {
        sheet_names: cli.sheets,
        slide_range,
        pdf_standard,
    };

    let result = office2pdf::convert_with_options(&cli.input, &options)
        .with_context(|| format!("converting {:?}", cli.input))?;

    // Print any warnings to stderr
    for warning in &result.warnings {
        eprintln!("Warning: {warning}");
    }

    std::fs::write(&output, result.pdf)
        .with_context(|| format!("writing output to {:?}", output))?;

    println!("Converted: {:?} -> {:?}", cli.input, output);
    Ok(())
}
