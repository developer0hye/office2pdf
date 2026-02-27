//! Convert with custom options (paper size, slide range, sheet filter).
//!
//! Usage:
//!   cargo run --example custom_options -- input.pptx output.pdf

use std::env;
use std::fs;
use std::process;

use office2pdf::config::{ConvertOptions, PaperSize, SlideRange};

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: {} <input> <output.pdf>", args[0]);
        process::exit(1);
    }

    let input = &args[1];
    let output = &args[2];

    let options = ConvertOptions {
        // Use A4 paper
        paper_size: Some(PaperSize::A4),
        // Only include slides 1 through 3 (for PPTX)
        slide_range: Some(SlideRange::new(1, 3)),
        // Only include these sheets (for XLSX)
        sheet_names: Some(vec!["Summary".to_string(), "Data".to_string()]),
        // Force landscape orientation
        landscape: Some(true),
        ..Default::default()
    };

    match office2pdf::convert_with_options(input, &options) {
        Ok(result) => {
            fs::write(output, &result.pdf).expect("failed to write PDF");
            println!("Wrote {} bytes to {output}", result.pdf.len());
        }
        Err(e) => {
            eprintln!("Conversion failed: {e}");
            process::exit(1);
        }
    }
}
