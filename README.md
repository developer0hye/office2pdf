# office2pdf

[![CI](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml/badge.svg)](https://github.com/developer0hye/office2pdf/actions/workflows/ci.yml)
[![crates.io](https://img.shields.io/crates/v/office2pdf.svg)](https://crates.io/crates/office2pdf)
[![docs.rs](https://docs.rs/office2pdf/badge.svg)](https://docs.rs/office2pdf)
[![License](https://img.shields.io/crates/l/office2pdf.svg)](LICENSE)

Pure-Rust library and CLI for converting DOCX, XLSX, and PPTX files to PDF.

No LibreOffice, no Chromium, no Docker — just a single binary powered by [Typst](https://github.com/typst/typst).

## Features

- **DOCX** — paragraphs, inline formatting (bold/italic/underline/color), tables (including floating tables), Word body-flow inline text boxes with paragraphs/tables/pictures, images, drawing shapes, ordered/nested lists, syntax-highlighted code, headers/footers, page setup
- **PPTX** — slides, text boxes, shapes, tables (with theme-based table styles), images, slide masters, speaker notes, gradient backgrounds, shadow/reflection effects; hidden slides are omitted by default and can be included explicitly
- **XLSX** — sheets (hidden ones skipped, as Excel does), chartsheets (one page-sized chart each), cell formatting including accounting currency alignment, merged cells, column widths, row heights, print scaling, Excel tables (built-in style banding, header/foot rules, bold header), conditional formatting (DataBar, IconSet, and formula rules)
- **PDF/A-2b** — archival-compliant output via `--pdf-a`
- **Embedded font extraction** — fonts embedded in PPTX/DOCX are automatically extracted, deobfuscated, and used during conversion
- **macOS Office font auto-discovery** — PowerPoint/Word/Excel bundled fonts are searched automatically; mutable per-user Office cloud caches require an explicit font path
- **WASM** — runs in browsers and Node.js via WebAssembly, with optional caller-provided or feature-gated Simplified Chinese fonts
- **Zero external dependencies** — runs as a standalone executable

DOCX packages with conflicting ASCII ZIP part names are rejected.

## Installation

### Library

```toml
[dependencies]
office2pdf = "0.10.0"
```

### Native builds without Typst's bundled fonts

The default `embedded-fonts` Cargo feature includes Typst's bundled font set.
To omit that set (including `NewCM10-Regular.otf`) from a native build, disable
default features (available since 0.8.0):

```toml
[dependencies]
office2pdf = { version = "0.10.0", default-features = false }
```

For the CLI:

```sh
cargo install office2pdf-cli --locked --no-default-features
```

Native font discovery, document-embedded fonts, caller-provided fonts, and
Office-specific bundled Noto/Selawik fallbacks remain available. Documents that
use Typst's fonts may render differently; supply suitable fonts with
`ConvertOptions::font_paths` / `font_bytes` (CLI: `--font-path`). Equations need
a suitable math font installed when Typst's math fonts are omitted.

Cargo features are additive: another dependency enabling
`office2pdf/embedded-fonts` or `typst-kit/embedded-fonts` can include the assets
again. WASM builds always retain Typst's bundled fonts because they cannot
search system font directories.

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
VERSION=v0.10.0
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
office2pdf spreadsheet.xlsx --xlsx-ui-script latin
office2pdf document.docx --pdf-a
office2pdf report.docx --font-path /usr/share/fonts/custom
```

Font names such as `Arial Narrow` select the family's width member when available.
Aptos and its width members require an explicit font input.

On macOS, `office2pdf` automatically searches fonts bundled in Microsoft Office
applications before falling back to regular system fonts. Where a face under
`/System/Library/Fonts` shares the PostScript name of the bundled face, line
boxes are measured from the system copy, as Word does. It does not
automatically read the mutable per-user `CloudFonts` or `PreviewFont` caches,
whose contents depend on previously opened documents. Pass such a cache (or any
other custom font directory) explicitly with `--font-path` or
`ConvertOptions::font_paths` when that host-specific behavior is intentional.

XLSX theme fonts (`<scheme val="minor"/>`, as Excel writes the default font)
resolve through the theme face for the script of Excel's user interface. The
default reproduces an Excel with a Korean UI, used for the converter's historical
reference exports; pass `--xlsx-ui-script latin`
(`ConvertOptions::xlsx_ui_script = XlsxUiScript::Latin`) to reproduce an
English, Danish, German or other Latin-script Excel.

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
| `--xlsx-ui-script <SCRIPT>` | UI script of the Excel to reproduce: `hangul` (default) or `latin`; picks the theme face for theme (scheme) fonts |
| `--slides <RANGE>` | PPTX slide range by source ordinal (e.g. `1-5` or `3`); hidden slides are omitted unless opted in |
| `--include-hidden-slides` | Include PPTX slides marked `show="0"` or `show="false"` (default: omitted) |
| `--font-path <DIR>` | Additional font directory override (repeatable) |

## Supported Formats

| Format | Status | Key Features |
|--------|--------|-------------|
| DOCX | Supported | Text, tables (including floating tables), Word body-flow inline text boxes with paragraphs/tables/pictures, images, drawing shapes, lists, code highlighting, headers/footers, page setup |
| PPTX | Supported | Slides, text boxes, shapes, tables, images, masters, gradients, effects |
| XLSX | Supported | Sheets, formatting, merged cells, column/row sizing, conditional formatting |

An inline DOCX picture that shares its paragraph with text stays on that text's
line, in run order, with its bottom edge on the baseline; in the IR it rides
`Run::inline_box`. Pictures with no text beside them remain `Image` or
`InlineImages` blocks.
DOCX picture crops apply per picture, including body, header and footer images.
Crops extending beyond the original image are not yet rendered.
Published library builds can omit DOCX picture rotation.

Damaged DOCX PNG pictures recover available pixels with warnings and keep their
drawing extent; undecodable headers use a black placeholder. Partial previews
may differ from Word's damaged-image rendering.

DOCX line-based section page borders support page/text offsets, per-edge paint,
first-page selection, and front/back placement. Custom `FlowPage` values use
`page_borders: None` to omit decorations. Decorative art borders are not supported.

DOCX legal numbering renders a list level's placeholders as decimal, preserving
the parent's own format. On/off flags and complete level overrides are honored.
Numbering-level run properties format marker fonts, sizes, weight and color
without changing body text. Inline prefixes retain ancestor counters, decimal
and zero-padded decimal, alphabetic and Roman formats, and literal label text.
Inline markers retain their size without stretching body line spacing.
Custom `TextStyle` values use `line_box: None` for automatic run edges.

DOCX child numbering supports never-restart and earlier-level restart rules;
omitted rules restart after any higher-level item. Word ignores restart rules
inside complete level overrides. Skipped child levels retain their actual style
and indentation. Skipped resettable ancestors advance later visible counters;
child levels with restart values of 0 or above 7 count physical items only.
Child restart values 1–7
suppress markers and numbering indents when they name the same/later level;
root flags are ignored. Wide numbered markers with tab suffixes advance the
first body line to default or custom left tab stops. Space/nothing suffixes
place list text after the marker, including inline numbers; continuations keep
their indent. Tabs in a list item's first paragraph, when it has a hanging
indent, advance from where its text starts to the paragraph's next stop,
counted from the margin.

DOCX body paragraph tabs also count stops and the default grid from the margin,
past the paragraph's left and first-line indents. When a tab's left or default
stop leaves no room for the next word before the line's end, the tab and that
word start the next line. A line left holding only tabs keeps their runs'
height. Centred, right and decimal stops keep their field on the line.
A negative left or right indent widens a body paragraph's lines into the
margin, as Word does; table cells still clamp it to zero until they clip their
content at the cell edge (#2261). A body list item's marker likewise starts at
its signed left indent minus its hanging span, and its lines at the signed left
indent, even left of the margin. While a paragraph style's `w:ind` still
stands in for a direct `w:numPr` level's (#2383), such an item keeps its marker
at the margin.

DOCX section body text supports top, centered and bottom page alignment,
including paragraph after-spacing and page breaks. Continuous sections use their
common alignment; mixed alignments fall back to top on their shared explicit
page. Automatic pagination within mixed continuous sections is approximate and
warns. Vertical justification (`both`) warns and retains top-aligned content.
Custom `FlowPage` values use `vertical_alignment: None` for top alignment.

DOCX section line numbering supports intervals, starts, page/section restarts,
continuous numbering and margin clearance. The start is a zero-based offset
(default 0 displays 1). A missing or zero interval disables numbering.
Numbers use the document default font;
custom line-number styles and paragraph suppression are not supported. Multi-column
number placement follows Typst's outside-margin behavior. Custom `FlowPage` values
use `line_numbering: None` to disable numbering.

DOCX and PPTX default language tags set the PDF language and text-layout language.
Per-run languages and DOCX East Asian/complex-script language tags are not yet used.

Automatic hyphenation is off unless enabled in the DOCX document settings.
Paragraph-level suppression and advanced hyphenation controls are not yet supported.

XLSX workbooks with cells beyond Excel's 16,384-column or 1,048,576-row grid
are rejected before layout expansion.

A workbook whose shared string table holds an item longer than Excel's
32,767-character cell limit is rejected as well, before the reader copies that
item into every cell referencing it.

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

`ooxml-package` is unpublished. Its public `PackageError::Container` variant
exposes `zip::ZipError` from ZIP 9, so maintainers must choose the crate's
semver/API compatibility strategy before publication. ZIP 9 cannot preserve an
invalid DOS timestamp while raw-copying an entry; `ooxml-package` rejects
entries when ZIP 9 reports an invalid DOS timestamp during load, before a
rewrite can normalize it. Round-trip tests cover valid, matching timestamps
in the local and central ZIP headers of untouched entries.

DOCX internal HTML/MHTML `altChunk` imports retain text, inline formatting,
headings, basic tables, and simple lists. Tag/class/id and inline CSS support
font, color, alignment, spacing, indentation, and cell widths. Images and row
spans in imports emit warnings while retaining surrounding text; RTF and
external imports are skipped with warnings. This is not a browser CSS engine.

## DOCX compatibility triage

The [docxide-pdf reference](references/docxide-pdf.md) documents a pinned external
Word-reference corpus and its current baseline. With that checkout and a current
CLI build, run `scripts/audit_docxide_corpus.py --help` for repeatable conversion,
page-count, and searchable-text triage. Geometry, pixels, fonts, and direct visual
inspection remain required before calling a document correct.

The render differ labels the raw 5% pixel-difference mask so thin horizontal
and vertical rules survive the cluster census. Its 20pt² reporting floor does
not establish materiality: inspect smaller marks and disconnected hairline
fragments even when no cluster is listed.

Picture outlines in PPTX and cached DOCX drawings retain ellipse, rounded-rectangle
and custom-path geometry, line opacity, and picture flips/rotation.
PPTX shapes also paint `<a:alpha>` on their outlines and on solid or gradient
fills, including colours supplied by `<p:style>` fill and line references.

DOCX body SmartArt imports saved drawing caches as vector shapes and searchable
text, inline or anchored in Word flow. The IR represents them as `DrawingGroup`
blocks; inline frames reserve their effect extents. Picture fills apply source
crops and signed stretch rectangles before shape clipping. Unavailable linked caches
and unsupported per-path painting produce warnings.
Header/footer and hyperlinked diagrams are not yet imported.

## Unattended issue runner

With Git, Python 3, authenticated `gh` with repository write access, and the selected
agent CLI logged in, run a single attempt from a trusted repository:

```sh
scripts/issue_loop.sh --agent codex --expect-user developer0hye --max-issues 1
scripts/issue_loop_watch.py --agent codex --dump 20
```

Claude remains the default. Both agents start fresh per issue and use the same
queue, prior-attempt history, main-check preflight, and optional `--reflect-every N`
prompt-improvement proposer; reflection opens a PR for review and must not merge it.
The Codex adapter uses CLI authentication, structured reports and JSONL events,
with `--max-runtime 7200` seconds by default instead of Claude's `--max-turns`.
`--model NAME` selects either CLI's model. The default Codex sandbox is
`danger-full-access` with approvals disabled; `--codex-sandbox` can restrict it.

Each attempt keeps its prompt, report and outcome, plus the transcript and exact
session ID when emitted by the CLI,
under `target/issue-loop-logs/runs/issue-N/`; Codex also writes live `events.jsonl`
and `events.stderr.log`. Inspect past runs with `scripts/issue_loop_history.py list`.
API failures abort rather than indicate an empty queue; Codex setup failures stop
without blocking an issue, and timeouts stop its process group. A final report is
not completion proof: the runner checks GitHub issue/PR state for progress, while
the agent must complete and verify CI, merge and cleanup under `AGENTS.md`.
Progress counts use every page of the repository's merged and open PR listings,
matching issue references in PR bodies without relying on the search index.
Inspect an issue's merged count and newest open citing PR with
`scripts/issue_loop.sh --print-progress N`; an API failure returns an error
without reporting a partial count.
See `scripts/issue_loop.sh --help` for limits, paths and stop behavior.

## Contributing files

Before attaching files to help improve office2pdf, complete a security review
and obtain your organization's internal approval for public sharing where
applicable. Submitters are responsible for authorized, lawful disclosure;
office2pdf and its maintainers disclaim related liability to the extent
permitted by applicable law. Read the [submission policy](CONTRIBUTING.md)
before submitting documents, screenshots, logs, or test fixtures.

Visual PRs follow the [evidence contract](assets/bugfixes/README.md) and finish
the [memory and speed pass](AGENTS.md#final-memory-and-speed-pass) after visual
acceptance. CI requires its completion record or a non-placeholder exemption
explanation; review verifies no runtime impact for exemptions, measurements,
and any revalidation.

## License

Licensed under [Apache License, Version 2.0](LICENSE).
