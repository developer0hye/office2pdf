use std::env;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use office2pdf::config::{ConvertOptions, Format};
use office2pdf::parser::Parser;
use office2pdf::parser::docx::DocxParser;
use office2pdf::parser::pptx::PptxParser;
use office2pdf::parser::xlsx::XlsxParser;
use office2pdf::render::typst_gen::generate_typst_with_options;

fn usage() -> &'static str {
    "usage: dump_typst_bundle <input.(docx|pptx|xlsx)> <output_dir>"
}

fn parse_args() -> Result<(PathBuf, PathBuf), String> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        return Err(usage().to_string());
    }
    Ok((PathBuf::from(&args[1]), PathBuf::from(&args[2])))
}

fn detect_format(input: &Path) -> Result<Format, String> {
    let ext = input
        .extension()
        .and_then(|ext| ext.to_str())
        .ok_or_else(|| "input file has no extension".to_string())?;
    Format::from_extension(ext).ok_or_else(|| format!("unsupported extension: {ext}"))
}

fn main() -> Result<(), Box<dyn Error>> {
    let (input, output_dir): (PathBuf, PathBuf) = parse_args().map_err(|msg| {
        eprintln!("{msg}");
        std::io::Error::new(std::io::ErrorKind::InvalidInput, msg)
    })?;

    let data: Vec<u8> = fs::read(&input)?;
    let options = ConvertOptions::default();
    let format = detect_format(&input).map_err(|msg| {
        eprintln!("{msg}");
        std::io::Error::new(std::io::ErrorKind::InvalidInput, msg)
    })?;

    let (doc, warnings) = match format {
        Format::Docx => DocxParser.parse(&data, &options)?,
        Format::Pptx => PptxParser.parse(&data, &options)?,
        Format::Xlsx => XlsxParser.parse(&data, &options)?,
    };

    for warning in &warnings {
        eprintln!("Warning: {warning}");
    }

    let typst_output = generate_typst_with_options(&doc, &options)?;

    fs::create_dir_all(&output_dir)?;
    let main_typ = output_dir.join("main.typ");
    fs::write(&main_typ, typst_output.source.as_bytes())?;

    for image in typst_output.images {
        let image_path = output_dir.join(image.path);
        fs::write(image_path, image.data)?;
    }

    println!("Wrote Typst bundle: {}", output_dir.display());
    println!("Main Typst source: {}", main_typ.display());
    Ok(())
}
