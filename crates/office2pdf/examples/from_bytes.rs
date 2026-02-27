//! Convert from in-memory bytes instead of a file path.
//!
//! Usage:
//!   cargo run --example from_bytes -- input.xlsx output.pdf

use std::env;
use std::fs;
use std::process;

use office2pdf::config::{ConvertOptions, Format};

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: {} <input> <output.pdf>", args[0]);
        process::exit(1);
    }

    let input = &args[1];
    let output = &args[2];

    // Read file into memory
    let data = fs::read(input).expect("failed to read input file");

    // Detect format from extension
    let ext = input.rsplit('.').next().unwrap_or("");
    let format = Format::from_extension(ext).unwrap_or_else(|| {
        eprintln!("Unsupported extension: .{ext}");
        process::exit(1);
    });

    // Convert from bytes
    match office2pdf::convert_bytes(&data, format, &ConvertOptions::default()) {
        Ok(result) => {
            fs::write(output, &result.pdf).expect("failed to write PDF");
            println!(
                "Converted {} bytes of {:?} → {} bytes of PDF",
                data.len(),
                format,
                result.pdf.len()
            );
        }
        Err(e) => {
            eprintln!("Conversion failed: {e}");
            process::exit(1);
        }
    }
}
