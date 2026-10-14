# typst

- **URL:** https://github.com/typst/typst
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

## Automatic Word cell borders

- Typst 0.15.1 [grid line layout](https://github.com/typst/typst/blob/v0.15.1/crates/typst-layout/src/grid/layouter.rs) retains zero-width vertical line geometry over resolved rows, merges and page fragments. Replace private lines after layout to avoid guessed paragraph-height border extents.
- [PDF table tagging](https://github.com/typst/typst/blob/v0.15.1/crates/typst-pdf/src/tags/context/table.rs) reads `CellGrid` fills for background attributes. The line approach preserves those fills and needs no resolved-model copies or dependency fork.
- A rejected transparent-fill prototype needed a copy-on-write API absent from registry Typst ([upstream request](https://github.com/typst/typst/issues/8955)). Workspace-only patches would not make that API available to published library consumers.
