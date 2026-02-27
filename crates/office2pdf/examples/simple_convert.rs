//! Convert an Office document to PDF.
//!
//! Usage:
//!   cargo run --example simple_convert -- input.docx output.pdf

use std::env;
use std::fs;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: {} <input> <output.pdf>", args[0]);
        eprintln!("  Supported formats: .docx, .pptx, .xlsx");
        process::exit(1);
    }

    let input = &args[1];
    let output = &args[2];

    match office2pdf::convert(input) {
        Ok(result) => {
            if !result.warnings.is_empty() {
                eprintln!("{} warning(s):", result.warnings.len());
                for w in &result.warnings {
                    eprintln!("  - {w}");
                }
            }
            fs::write(output, &result.pdf).expect("failed to write PDF");
            println!("Wrote {} bytes to {output}", result.pdf.len());
        }
        Err(e) => {
            eprintln!("Conversion failed: {e}");
            process::exit(1);
        }
    }
}
