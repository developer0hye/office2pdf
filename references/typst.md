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

## Flow page-break rules (`typst-layout` 0.15.1, `flow/distribute.rs`)

- `#v` is relative spacing: `rel()` only subtracts it from the region and never
  checks fit, so spacing cut off at a region's end carries nothing over.
- A breakable block moves whole to the next region when the region is already
  full, or when its first frame is empty and its spill holds a non-empty frame;
  spacing alone makes neither.
- An unbreakable block (`single()`) that does not fit finishes the region and
  moves whole; where `may_progress()` is false, as in a fresh page or a fixed
  box, it overflows in place. An empty paragraph's blocks are unbreakable for
  that reason (#2253).
- A tag, such as `#metadata(none) <label>`, waits in `work.tags` until the next
  frame or placed item lands, so a section label written before a block that
  moves to the next region lands there with it, at the body top (#2254).
