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

- Diagonal cell paint reads Typst 0.15.1's completed fill rectangles instead of
  measuring cell content: `typst-layout/src/grid/layouter.rs` resolves automatic
  rows and merged tracks before painting fills. A private fill marker is restored
  to the source background before diagonal strokes and cell text are painted.
