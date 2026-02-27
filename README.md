# office2pdf

[![CI](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml/badge.svg)](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml)
[![crates.io](https://img.shields.io/crates/v/office2pdf.svg)](https://crates.io/crates/office2pdf)
[![docs.rs](https://docs.rs/office2pdf/badge.svg)](https://docs.rs/office2pdf)
[![License](https://img.shields.io/crates/l/office2pdf.svg)](LICENSE)

Pure-Rust library and CLI for converting DOCX, XLSX, and PPTX files to PDF.

No LibreOffice, no Chromium, no Docker — just a single binary powered by [Typst](https://github.com/typst/typst).

## Features

- **DOCX** — paragraphs, inline formatting (bold/italic/underline/color), tables, images, lists, headers/footers, page setup
- **PPTX** — slides, text boxes, shapes, images, slide masters, speaker notes, gradient backgrounds, shadow/reflection effects
- **XLSX** — sheets, cell formatting, merged cells, column widths, row heights, conditional formatting (DataBar, IconSet)
- **PDF/A-2b** — archival-compliant output via `--pdf-a`
- **WASM** — runs in browsers and Node.js via WebAssembly (optional `wasm` feature)
- **Zero external dependencies** — runs as a standalone executable

## Installation

### Library

```toml
[dependencies]
office2pdf = "0.3"
```

### CLI

```sh
cargo install office2pdf-cli
```

## Quick Start

### As a library

```rust
// Simple one-liner
let result = office2pdf::convert("report.docx").unwrap();
std::fs::write("report.pdf", &result.pdf).unwrap();

// With options
use office2pdf::config::{ConvertOptions, PaperSize};

let options = ConvertOptions {
    paper_size: Some(PaperSize::A4),
    ..Default::default()
};
let result = office2pdf::convert_with_options("slides.pptx", &options).unwrap();
std::fs::write("slides.pdf", &result.pdf).unwrap();

// In-memory conversion
use office2pdf::config::Format;

let docx_bytes = std::fs::read("report.docx").unwrap();
let result = office2pdf::convert_bytes(
    &docx_bytes,
    Format::Docx,
    &ConvertOptions::default(),
).unwrap();
std::fs::write("report.pdf", &result.pdf).unwrap();
```

### CLI

```sh
# Single file
office2pdf report.docx

# Explicit output path
office2pdf report.docx -o output.pdf

# Batch conversion
office2pdf *.docx --outdir pdfs/

# With options
office2pdf slides.pptx --paper a4 --landscape
office2pdf spreadsheet.xlsx --sheets "Sheet1,Summary"
office2pdf document.docx --pdf-a
office2pdf report.docx --font-path /usr/share/fonts/custom
```

### WASM (Browser / Node.js)

Build with `wasm-pack`:

```sh
wasm-pack build crates/office2pdf --target web --features wasm
```

Use from JavaScript:

```js
import init, { convertDocxToPdf, convertToPdf } from './pkg/office2pdf.js';

await init();

const docxBytes = new Uint8Array(await file.arrayBuffer());
const pdfBytes = convertDocxToPdf(docxBytes);

// Or use the generic API with a format string
const pdfBytes2 = convertToPdf(xlsxBytes, "xlsx");
```

Available functions: `convertToPdf(data, format)`, `convertDocxToPdf(data)`, `convertPptxToPdf(data)`, `convertXlsxToPdf(data)`.

## CLI Options

| Flag | Description |
|------|-------------|
| `-o, --output <PATH>` | Output file path (single input only) |
| `--outdir <DIR>` | Output directory for batch conversion |
| `--paper <SIZE>` | Paper size: `a4`, `letter`, `legal` |
| `--landscape` | Force landscape orientation |
| `--pdf-a` | Produce PDF/A-2b compliant output |
| `--sheets <NAMES>` | XLSX sheet filter (comma-separated) |
| `--slides <RANGE>` | PPTX slide range (e.g. `1-5` or `3`) |
| `--font-path <DIR>` | Additional font directory (repeatable) |

## Supported Formats

| Format | Status | Key Features |
|--------|--------|-------------|
| DOCX | Supported | Text, tables, images, lists, headers/footers, page setup |
| PPTX | Supported | Slides, text boxes, shapes, images, masters, gradients, effects |
| XLSX | Supported | Sheets, formatting, merged cells, column/row sizing, conditional formatting |

## License

Licensed under [Apache License, Version 2.0](LICENSE).
