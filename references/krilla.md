# krilla

- **URL:** https://github.com/LaurenzV/krilla
- **Language:** Rust
- **License:** MIT / Apache 2.0

## What it does

High-level PDF creation library built on `pdf-writer`. Now the PDF backend for Typst (since PR #5420). Supports PDF/A-1 through PDF/A-4 and PDF/UA-1.

## Why relevant

- **The engine producing our PDF output** — Typst uses krilla for PDF generation.
- Supports tagged PDF for accessibility (PDF/UA), which maps to office2pdf's accessibility features.
- Understanding krilla's capabilities and limitations directly affects what PDF features we can offer.
- Font subsetting, gradient fills, and PDF standard compliance are handled here.

## Key areas to study

- PDF/A compliance requirements
- PDF/UA tagging (structure tags: H1-H6, P, Table, Figure)
- Font subsetting behavior
- SVG text: [#455](https://github.com/LaurenzV/krilla/pull/455) batches solid-fill
  glyphs; [#456](https://github.com/LaurenzV/krilla/pull/456) preserves visible
  logical text through bounded ActualText spans without changing paint.
  Both crates share a published-version backport to preserve type identity.
  Extraction tests use [pdf-extract#165](https://github.com/jrmuizel/pdf-extract/pull/165).

## When to consult

- Debugging PDF output issues (fonts, compliance, accessibility)
- Understanding PDF/A or PDF/UA requirements
- Investigating PDF file size optimization
