# office2pdf

[![CI](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml/badge.svg)](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml)
[![crates.io](https://img.shields.io/crates/v/office2pdf.svg)](https://crates.io/crates/office2pdf)
[![docs.rs](https://docs.rs/office2pdf/badge.svg)](https://docs.rs/office2pdf)
[![License](https://img.shields.io/crates/l/office2pdf.svg)](LICENSE)

Pure-Rust library and CLI for converting DOCX, XLSX, and PPTX files to PDF.

No LibreOffice, no Chromium, no Docker — just a single binary powered by [Typst](https://github.com/typst/typst).

## Features

- **DOCX** — paragraphs, inline formatting (bold/italic/underline/color), tables, images, drawing shapes, ordered/nested lists, syntax-highlighted code, headers/footers, page setup
- **PPTX** — slides, text boxes, shapes, tables (with theme-based table styles), images, slide masters, speaker notes, gradient backgrounds, shadow/reflection effects; hidden slides are omitted by default and can be included explicitly
- **XLSX** — sheets (hidden ones skipped, as Excel does), chartsheets (one page-sized chart each), cell formatting, merged cells, column widths, row heights, Excel tables (built-in style banding, header/foot rules, bold header), conditional formatting (DataBar, IconSet, and formula rules)
- **PDF/A-2b** — archival-compliant output via `--pdf-a`
- **Embedded font extraction** — fonts embedded in PPTX/DOCX are automatically extracted, deobfuscated, and used during conversion
- **macOS Office font auto-discovery** — PowerPoint/Word/Excel bundled fonts are searched automatically; mutable per-user Office cloud caches require an explicit font path
- **WASM** — runs in browsers and Node.js via WebAssembly, with optional caller-provided or feature-gated Simplified Chinese fonts
- **Zero external dependencies** — runs as a standalone executable

## Installation

### Library

```toml
[dependencies]
office2pdf = "0.7.0"
```

### CLI

```sh
cargo install office2pdf-cli
```

#### Prebuilt binaries

Every [GitHub release](https://github.com/developer0hye/office2pdf/releases) ships standalone CLI binaries — no Rust toolchain needed:

| Platform | Asset |
|----------|-------|
| Linux x86_64 (glibc) | `office2pdf-<version>-x86_64-unknown-linux-gnu.tar.gz` |
| Linux x86_64 (static musl) | `office2pdf-<version>-x86_64-unknown-linux-musl.tar.gz` |
| Linux ARM64 | `office2pdf-<version>-aarch64-unknown-linux-gnu.tar.gz` |
| macOS Apple Silicon | `office2pdf-<version>-aarch64-apple-darwin.tar.gz` |
| macOS Intel | `office2pdf-<version>-x86_64-apple-darwin.tar.gz` |
| Windows x86_64 | `office2pdf-<version>-x86_64-pc-windows-msvc.zip` |

On Linux and macOS, download, extract, and place the binary on your `PATH`:

```sh
VERSION=v0.7.0
TARGET=x86_64-unknown-linux-gnu  # pick your platform's target from the table above
curl -L "https://github.com/developer0hye/office2pdf/releases/download/${VERSION}/office2pdf-${VERSION}-${TARGET}.tar.gz" | tar xz
sudo install "office2pdf-${VERSION}-${TARGET}/office2pdf" /usr/local/bin/
```

On Windows, unzip the archive and add `office2pdf.exe` to your `PATH`.

The macOS binaries are not notarized. Binaries downloaded with a browser are quarantined by Gatekeeper; clear the flag with `xattr -d com.apple.quarantine office2pdf` (downloads via `curl` are unaffected).

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
    include_hidden_slides: true,
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
office2pdf slides.pptx --include-hidden-slides
office2pdf spreadsheet.xlsx --sheets "Sheet1,Summary"
office2pdf document.docx --pdf-a
office2pdf report.docx --font-path /usr/share/fonts/custom
```

On macOS, `office2pdf` automatically searches fonts bundled in Microsoft Office
applications before falling back to regular system fonts. Where a face under
`/System/Library/Fonts` shares the PostScript name of the bundled face, line
boxes are measured from the system copy, as Word does. It does not
automatically read the mutable per-user `CloudFonts` or `PreviewFont` caches,
whose contents depend on previously opened documents. Pass such a cache (or any
other custom font directory) explicitly with `--font-path` or
`ConvertOptions::font_paths` when that host-specific behavior is intentional.

### WASM (Browser / Node.js)

Build with `wasm-pack`:

```sh
wasm-pack build crates/office2pdf --target web --features wasm --locked
```

The default build does not add a CJK font. For zero-configuration Simplified
Chinese fallback, opt in to the 3.3 MiB GB2312 subset (this feature implies
`wasm`):

```sh
wasm-pack build crates/office2pdf --target web --features wasm-cjk-font --locked
```

Use from JavaScript:

```js
import init, {
  Office2PdfConverter,
  convertDocxToPdf,
  convertToPdf,
} from './pkg/office2pdf.js';

await init();

const docxBytes = new Uint8Array(await file.arrayBuffer());
const pdfBytes = convertDocxToPdf(docxBytes);

// Or use the generic API with a format string
const pdfBytes2 = convertToPdf(xlsxBytes, "xlsx");

// Register a font for this converter and make it the final fallback for every
// emitted font chain. The result-bearing API preserves structured warnings.
const fontBytes = new Uint8Array(await fontFile.arrayBuffer());
const converter = new Office2PdfConverter();
converter.registerFont(fontBytes);
converter.setLastResortFontFamily("Noto Sans SC");

const result = converter.convertDocxToPdf(docxBytes);
for (let index = 0; index < result.warningCount; index += 1) {
  const warning = result.warningAt(index);
  console.warn(warning.kind, warning.from, warning.to, warning.message);
}
const pdfBytes3 = result.pdf;
```

The compatibility functions return PDF bytes directly:
`convertToPdf(data, format)`, `convertDocxToPdf(data)`,
`convertPptxToPdf(data)`, and `convertXlsxToPdf(data)`. Their
`*WithResult` counterparts return `ConversionResult`, which exposes `pdf`,
`warningCount`, and `warningAt(index)`. `Office2PdfConverter` provides the same
result-bearing methods together with `registerFont`, `clearFonts`,
`setLastResortFontFamily`, and `clearLastResortFontFamily`.

Browser and Node.js builds use the bundled Typst fallback fonts and also honor
font faces embedded inside DOCX and PPTX files. Filesystem font paths remain
unavailable in WASM, but callers can supply standalone TTF, OTF, or TTC bytes
through `Office2PdfConverter`. When no registered or embedded face covers a
CJK run, the result-bearing API emits a `fallback-used` warning whose `to`
value is `.notdef` in the default build.

With `wasm-cjk-font`, `Noto Sans CJK SC` is automatically appended when the
caller has not configured another last-resort family. The subset covers the
complete GB2312 repertoire (7,445 characters, including 6,763 Han characters)
plus printable ASCII. It is not full Traditional Chinese, Japanese, or Korean
coverage. The warning `to` value identifies `Noto Sans CJK SC` when the bundle
is used. An explicit `setLastResortFontFamily` call takes precedence.

Native Rust callers can use the same per-conversion path through
`ConvertOptions::font_bytes` and `ConvertOptions::last_resort_font_family`.

## CLI Options

| Flag | Description |
|------|-------------|
| `-o, --output <PATH>` | Output file path (single input only) |
| `--outdir <DIR>` | Output directory for batch conversion |
| `--paper <SIZE>` | Paper size: `a4`, `letter`, `legal` |
| `--landscape` | Force landscape orientation |
| `--pdf-a` | Produce PDF/A-2b compliant output |
| `--sheets <NAMES>` | XLSX sheet filter (comma-separated); the only way to print a hidden sheet |
| `--slides <RANGE>` | PPTX slide range by source ordinal (e.g. `1-5` or `3`); hidden slides are omitted unless opted in |
| `--include-hidden-slides` | Include PPTX slides marked `show="0"` or `show="false"` (default: omitted) |
| `--font-path <DIR>` | Additional font directory override (repeatable) |

## Supported Formats

| Format | Status | Key Features |
|--------|--------|-------------|
| DOCX | Supported | Text, tables, images, drawing shapes, lists, code highlighting, headers/footers, page setup |
| PPTX | Supported | Slides, text boxes, shapes, tables, images, masters, gradients, effects |
| XLSX | Supported | Sheets, formatting, merged cells, column/row sizing, conditional formatting |

## Workspace Crates

| Crate | Published | Purpose |
|-------|-----------|---------|
| `office2pdf` | [crates.io](https://crates.io/crates/office2pdf) | Conversion library (this README) |
| `office2pdf-cli` | [crates.io](https://crates.io/crates/office2pdf-cli) | Command-line interface |
| `ooxml-package` | not yet | Lossless OPC (OOXML container) model: byte-for-byte round trip, dirty-part tracking, surgical save |
| `pptx-model` | not yet | Lossless, editable PPTX model over `ooxml-package`: slide enumeration, surgical text-run edits |

`ooxml-package` and `pptx-model` are the foundation for round-trip OOXML
editing: a load/save cycle preserves every package part — including parts and
XML they do not model — and saving rewrites only edited entries. The
`office2pdf` rendering pipeline stays a derived projection for PDF
preview/export; it is not the editable source of truth.

## Contributing files

Before attaching files to help improve office2pdf, complete a security review
and obtain your organization's internal approval for public sharing where
applicable. Submitters are responsible for authorized, lawful disclosure;
office2pdf and its maintainers disclaim related liability to the extent
permitted by applicable law. Read the [submission policy](CONTRIBUTING.md)
before submitting documents, screenshots, logs, or test fixtures.

## License

Licensed under [Apache License, Version 2.0](LICENSE).
