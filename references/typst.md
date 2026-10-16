# typst

- **URL:** https://github.com/typst/typst
- **Stars:** ~51.7k
- **Language:** Rust
- **License:** Apache 2.0

## What it does

Markup-based typesetting system. Compiles `.typ` source to PDF/PNG/SVG/HTML. Embeddable as a Rust library.

## Why relevant

- **Our backend** — office2pdf generates Typst markup and compiles it.
- `crates/typst-library/` defines all available Typst elements and their parameters — the definitive reference for what markup we should generate.
- Understanding the layout engine, `World` trait, and compilation pipeline is essential for performance tuning.

## Key areas to study

- `crates/typst-library/src/layout/` — Page, table, grid, columns
- `crates/typst-library/src/text/` — Text, paragraph, heading
- `crates/typst-library/src/math/` — Math equation support
- `crates/typst/src/compile.rs` — Compilation pipeline

## When to consult

- Improving Typst codegen quality or discovering new Typst features
- Debugging layout issues in generated PDFs
- Optimizing compilation performance

## DOCX heading weight (#2176)

- Typst 0.15.1's [`HeadingElem::show_set`](https://github.com/typst/typst/blob/v0.15.1/crates/typst-library/src/model/heading.rs) supplies bold by default.
- A regular text scope inside the semantic heading lets resolved Word runs select
  their own explicit or family-derived weight while keeping outline metadata.
- Word's native reproduction also inherits regular typography from Normal;
  that parent treatment must suppress synthesized parser bold.
