# office2pdf - Product Requirements Document (PRD)

## 1. Overview

**office2pdf** is a library and CLI tool that converts DOCX, XLSX, and PPTX files to PDF using pure Rust.
It operates standalone without external runtimes (LibreOffice, Chromium, Docker), using the Typst engine as the layout/PDF backend.

### Core Values
- **Zero dependency**: No external binaries/services required, runs as a single executable
- **High-quality output**: Reproduces the original document's layout/styles at 95% fidelity
- **Library-first**: Embeddable in other Rust projects, CLI is a thin wrapper

---

## 2. Target Users

| User | Use Case |
|---|---|
| **Backend developers** | Server-side document → PDF conversion (reports, invoices, printouts) |
| **DevOps/Infrastructure** | Lightweight conversion pipeline without LibreOffice/Docker |
| **CLI users** | Quick batch conversion from terminal |
| **Rust developers** | Embedding as a crate in projects |

---

## 3. Functional Requirements

### 3.1 Input Formats

#### DOCX (Word)

| Priority | Feature | Description |
|---|---|---|
| P0 | Text | Paragraphs, line breaks, page breaks |
| P0 | Inline formatting | Bold, italic, underline, strikethrough, font, size, color |
| P0 | Paragraph formatting | Alignment (left/right/center/justify), indentation, line spacing |
| P0 | Tables | Basic tables, cell merging, borders, background color |
| P0 | Images | Inline images, basic size adjustment |
| P1 | Lists | Numbered, bulleted (multi-level) |
| P1 | Headers/Footers | Text, page numbers |
| P1 | Page setup | Paper size, margins, orientation (portrait/landscape) |
| P1 | Styles | Document stylesheet (Heading 1~6, etc.) application |
| P2 | Table of Contents | TOC rendering |
| P2 | Hyperlinks | Clickable links in PDF |
| P2 | Footnotes/Endnotes | Footnote, Endnote rendering |
| P3 | Text wrapping | Text flow around images |
| P3 | Equations | OMML math equations |
| P3 | Charts | Embedded chart rendering |

#### PPTX (PowerPoint)

| Priority | Feature | Description |
|---|---|---|
| P0 | Slide → Page | 1 slide = 1 page mapping |
| P0 | Text boxes | Position, size, text content, formatting |
| P0 | Basic shapes | Rectangles, circles, lines, etc. |
| P0 | Images | Image placement within slides |
| P1 | Backgrounds | Solid color, gradient, image backgrounds |
| P1 | Master/Layout | Slide master → Layout → Slide inheritance chain |
| P1 | Themes | Theme colors, theme fonts interpretation |
| P1 | Tables | Table rendering within slides |
| P2 | Group shapes | Grouped shape handling |
| P2 | Shape styles | Shadow, reflection, rotation, transparency |
| P3 | SmartArt | SmartArt diagrams |
| P3 | Charts | Embedded charts |

#### XLSX (Excel)

| Priority | Feature | Description |
|---|---|---|
| P0 | Cell data | Text, number, date value output |
| P0 | Basic table layout | Rows/columns → PDF table conversion |
| P0 | Cell merging | Merged cells handling |
| P1 | Cell formatting | Font, color, background color, borders |
| P1 | Column width/Row height | Original size reflection |
| P1 | Number formats | Currency, percentage, date format strings |
| P1 | Sheet selection | Specific sheet or all sheets conversion |
| P2 | Print area | Configured Print Area reflection |
| P2 | Headers/Footers | Sheet headers/footers |
| P2 | Page breaks | Manual page break handling |
| P3 | Conditional formatting | Conditional formatting rendering |
| P3 | Charts | Embedded chart rendering |

### 3.2 Output

| Feature | Description |
|---|---|
| PDF output | Valid PDF file generation |
| PDF version | PDF 1.7 default, PDF/A option |
| Font embedding | Embed used fonts in PDF |
| Metadata | Title, author, creation date, etc. |

---

## 4. Non-Functional Requirements

| Category | Requirement |
|---|---|
| **Performance** | 10-page document < 1 second, 100-page < 5 seconds |
| **Memory** | < 500MB for 100-page documents |
| **Binary size** | CLI < 50MB (including fonts) |
| **Platforms** | Windows, macOS, Linux |
| **Error handling** | Skip unparseable elements with warnings, continue overall conversion |
| **Font fallback** | System font discovery + built-in default fonts |

---

## 5. Architecture

### 5.1 Conversion Pipeline

```
Input file (.docx/.pptx/.xlsx)
    │
    ▼
[1. Parser] ─── docx-rs / ppt-rs / umya-spreadsheet
    │
    ▼
[2. IR (Intermediate Representation)]  ← Format-independent document model
    │
    ▼
[3. Typst Codegen] ─── IR → Typst markup generation
    │
    ▼
[4. Typst Compile] ─── typst-as-lib / typst crate
    │
    ▼
[5. PDF Export] ─── typst-pdf
    │
    ▼
output.pdf
```

### 5.2 IR (Intermediate Representation) Design

All input formats are converted to a common IR before being output to Typst:

```rust
pub struct Document {
    pub metadata: Metadata,
    pub pages: Vec<Page>,
    pub styles: StyleSheet,
}

pub enum Page {
    Flow(FlowPage),     // DOCX: flowing text pages
    Fixed(FixedPage),   // PPTX: fixed coordinate pages
    Table(TablePage),   // XLSX: table-based pages
}

pub struct FlowPage {
    pub size: PageSize,
    pub margins: Margins,
    pub header: Option<HeaderFooter>,
    pub footer: Option<HeaderFooter>,
    pub content: Vec<Block>,
}

pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(Image),
    PageBreak,
    List(List),
}
```

### 5.3 Project Structure

```
office2pdf/
├── Cargo.toml                    # workspace root
├── crates/
│   ├── office2pdf/                  # library crate
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs            # public API
│   │       ├── error.rs          # thiserror error types
│   │       ├── ir/               # Intermediate Representation
│   │       │   ├── mod.rs
│   │       │   ├── document.rs   # Document, Page, Block, etc.
│   │       │   ├── style.rs      # Style/formatting model
│   │       │   └── elements.rs   # Paragraph, Table, Image, etc.
│   │       ├── parser/           # Input format parsers
│   │       │   ├── mod.rs        # Parser trait
│   │       │   ├── docx.rs       # DOCX → IR
│   │       │   ├── pptx.rs       # PPTX → IR
│   │       │   └── xlsx.rs       # XLSX → IR
│   │       ├── render/           # IR → PDF rendering
│   │       │   ├── mod.rs
│   │       │   ├── typst_gen.rs  # IR → Typst markup generation
│   │       │   └── pdf.rs        # Typst compile + PDF output
│   │       └── config.rs         # Conversion options
│   └── office2pdf-cli/              # CLI crate
│       ├── Cargo.toml
│       └── src/
│           └── main.rs           # clap-based CLI
├── tests/                        # Integration tests
│   ├── fixtures/                 # Test document files
│   └── integration_tests.rs
└── fonts/                        # Built-in default fonts
```

### 5.4 Key Dependencies

| Crate | Purpose | Notes |
|---|---|---|
| `typst` | Layout engine | or `typst-as-lib` |
| `typst-pdf` | PDF output | |
| `typst-kit` | Font discovery | |
| `docx-rs` | DOCX parsing | v0.4.19, actively maintained |
| `umya-spreadsheet` | XLSX parsing (with formatting) | Style/formatting extraction capable |
| `ppt-rs` | PPTX parsing | Rust port of python-pptx |
| `clap` | CLI argument parsing | v4 derive |
| `thiserror` | Library errors | |
| `anyhow` | CLI errors | |

---

## 6. Public API (Library)

```rust
use office2pdf::{Document, ConvertOptions, Format};

// Simple usage
let pdf_bytes = office2pdf::convert("input.docx")?;
std::fs::write("output.pdf", pdf_bytes)?;

// With options
let options = ConvertOptions::builder()
    .paper_size(PaperSize::A4)
    .font_paths(vec!["./fonts"])
    .pdf_standard(PdfStandard::PdfA2b)
    .build();

let pdf_bytes = office2pdf::convert_with_options("input.xlsx", &options)?;

// Convert from bytes
let docx_bytes = std::fs::read("input.docx")?;
let pdf_bytes = office2pdf::convert_bytes(&docx_bytes, Format::Docx, &options)?;

// Convert specific sheets/slides only
let options = ConvertOptions::builder()
    .sheet_names(vec!["Sheet1"])  // XLSX: specific sheets only
    .slide_range(1..=5)          // PPTX: slides 1-5 only
    .build();
```

---

## 7. CLI Interface

```bash
# Basic conversion
office2pdf input.docx                     # → input.pdf
office2pdf input.pptx -o output.pdf       # specify output path

# Options
office2pdf input.xlsx --sheets "Sheet1,Sheet2"   # specific sheets only
office2pdf input.pptx --slides 1-5               # specific slides only
office2pdf input.docx --paper a4 --landscape     # paper settings
office2pdf input.docx --font-path ./fonts        # font path

# Batch conversion
office2pdf *.docx --outdir ./pdfs/               # batch convert multiple files

# Info
office2pdf --version
office2pdf --help
```

---

## 8. Implementation Phases

### Phase 1: MVP (Basic Text + Images)
- Project structure setup (workspace, CI)
- IR definition
- DOCX P0 features → IR → Typst → PDF
- PPTX P0 features → IR → Typst → PDF
- XLSX P0 features → IR → Typst → PDF
- Basic CLI functionality

### Phase 2: Formatting + Styles
- All P1 feature implementation
- Font embedding / fallback
- Enhanced error handling

### Phase 3: Advanced Features
- P2 feature implementation
- PDF/A support
- Performance optimization
- Batch conversion

### Phase 4: Polish
- P3 features (where feasible)
- Edge case handling
- Documentation, examples, crates.io publishing

---

## 9. Validation Methods

| Method | Description |
|---|---|
| **Golden tests** | Test document → PDF → screenshot comparison |
| **Round-trip** | Convert known documents and verify via text extraction |
| **Manual comparison** | MS Office PDF output vs office2pdf output comparison |
| **CI** | Automatic conversion of test document set on each commit |

---

## 10. Risks

| Risk | Impact | Mitigation |
|---|---|---|
| OOXML spec complexity | Parsing crates may not support all elements | Skip unsupported elements + warning logs |
| Typst IR conversion limits | Some layouts may not be expressible in Typst | Best approximation within Typst capabilities |
| Font compatibility | Original fonts may not be available | System font discovery + fallback font mapping |
| Parsing crate maintenance | Dependency crates may become abandoned | Prepare forks, consider self-implementation for core parsers |
