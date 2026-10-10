# Issue 2060 visual evidence

Authored synthetic fixture: `tests/fixtures/docx/issue-2060-aptos-narrow.docx`. Native reference exported on
2026-10-10 with Microsoft Word 16.113.4 on macOS 26.6.2 (25G83).
The one-page reference uses Quartz PDFContext; current output uses Typst 0.15.1.
Word has a 400.08 x 600pt media box; current output is 400 x 600pt.

SHA-256:

- Source: `c07221487d3ef3bd80d8267a98ab13aeb143e70ab00f03439273d2910de3ac06`
- Native GT PDF: `a3285294189128f4016c4634a03f69c644cc9b7e6bf176fd73ecab4865e937e3`
- Current output PDF: `e81112b0b061dbaa405e9bc58b4b9e8e654ba2f1b0af25ec5bfbb0df6388edda`
- Explicit Aptos-Narrow.ttf input: `ded7515aa064578485a09c11a798a05d0303bdacdd7f2de92d0521f3bc7793ba`

The before converter is `6a296de889c0119cb59875a96b2292204436d957`;
current output uses this PR. Both conversions receive an isolated font directory
containing only the same Aptos-Narrow.ttf from the local Office installation.
The licensed font is not committed. Build products, PDFs and font-input evidence
remain outside the worktree. Native re-export timestamps can change PDF hashes.

Reproduction (supply a licensed Aptos Narrow font in `$FONT_DIR`):

```sh
stage="$HOME/Library/Containers/com.microsoft.Word/Data/pr-review-2061"
mkdir -p "$stage"
cp tests/fixtures/docx/issue-2060-aptos-narrow.docx "$stage/"
osascript scripts/macos/export_word_pdfs.applescript "$stage/gt" aptos-narrow "$stage/issue-2060-aptos-narrow.docx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-2060-aptos-narrow.docx --font-path "$FONT_DIR" -o after.pdf
python3 scripts/compare_layout.py "$stage/gt/aptos-narrow.pdf" after.pdf --json --audit --fine-shift 0.5
python3 scripts/compare_text_layer.py "$stage/gt/aptos-narrow.pdf" after.pdf
python3 scripts/compare_render.py "$stage/gt/aptos-narrow.pdf" after.pdf --fine-shift 0.5 --artifacts-dir audit --cluster-report clusters.json --strict-clusters
```

Before substitutes Calibri for Aptos Narrow despite the explicit input, changing
the letter shapes and widths. After embeds Aptos Narrow, agrees with all three
native text baselines within 0.01242pt, and preserves identical searchable text.
Full native/before/after pages, the 5% diff and matched crops of all three lines
were inspected at 150 DPI. Black regular 12pt text, spacing, left alignment and
line ends agree; no shapes, hairlines, borders, emphasis, rotation, clipping or
overlap are present. Sparse glyph-edge specks near the ends of the first and
third lines are below the material-cluster floor; the strict report has zero
clusters. JPEGs are progressive quality 86, stripped then reset to 150 DPI.

Portable tests separately verify the absent-width case: a regular-only family
reports its actual fallback and measures with that regular face's advances.
The Aptos source-policy regression keeps incidental host fonts out of both
paint and metric chains while allowing explicit font inputs.
