use std::path::PathBuf;
use std::process;

use anyhow::{Context, Result};
use clap::Parser;

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

    let result =
        office2pdf::convert(&cli.input).with_context(|| format!("converting {:?}", cli.input))?;

    // Print any warnings to stderr
    for warning in &result.warnings {
        eprintln!("Warning: {warning}");
    }

    std::fs::write(&output, result.pdf)
        .with_context(|| format!("writing output to {:?}", output))?;

    println!("Converted: {:?} -> {:?}", cli.input, output);
    Ok(())
}
