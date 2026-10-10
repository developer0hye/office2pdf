# Issue 2050 visual evidence

Authored synthetic fixture: `tests/fixtures/xlsx/issue-2050-latin-ui-theme.xlsx`. Native reference exported on
2026-10-10 with Microsoft Excel 16.113.4 on macOS 26.6.2 (25G83), running
with an English user interface. Excel was launched with the process-only
`-AppleLanguages '(en)'` argument; its normal Korean UI was restored afterward.
The reference is a one-page A4 Quartz PDFContext export; current output uses
Typst 0.15.1. Both embed Calibri.

SHA-256:

- Source: `3e6d0b601a77c3a3ffdc0fc3f90c3c25b25c8614d94bf2f6735399ed0f157e11`
- Native GT PDF: `0e08cc212e683ed5caf45aeca8b02727811f177d43cf4b7b5f446ee34eea835c`
- Current output PDF: `a297f9aa54e3f7ea8f750162e871306a1f9c2a2b7e4ef432f7ae76841cb0f63c`

The source uses Office's minor Latin Calibri / Hang Malgun Gothic theme,
Calibri 11pt scheme fonts, three declared 36pt rows and three automatic rows.
English Excel prints their pitches at 33pt and 14pt respectively. The before
converter is `6a296de889c0119cb59875a96b2292204436d957`, using the default
Korean UI calibration. Current output uses this PR with `--xlsx-ui-script latin`.
The same current converter without that option produces a pixel-identical page
to before (exact decoded AE = 0), preserving the existing default.

Reproduction (start an otherwise idle Excel session in English first):

```sh
open -a 'Microsoft Excel' --args -AppleLanguages '(en)'
stage="$HOME/Library/Containers/com.microsoft.Excel/Data/pr-review-2051"
mkdir -p "$stage"
cp tests/fixtures/xlsx/issue-2050-latin-ui-theme.xlsx "$stage/"
osascript scripts/macos/export_excel_pdfs.applescript "$stage/gt" latin-ui-theme "$stage/issue-2050-latin-ui-theme.xlsx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/xlsx/issue-2050-latin-ui-theme.xlsx --xlsx-ui-script latin -o after.pdf
python3 scripts/compare_layout.py "$stage/gt/latin-ui-theme-sheet-0001.pdf" after.pdf --json --audit --fine-shift 0.5
python3 scripts/compare_text_layer.py "$stage/gt/latin-ui-theme-sheet-0001.pdf" after.pdf
python3 scripts/compare_render.py "$stage/gt/latin-ui-theme-sheet-0001.pdf" after.pdf --fine-shift 0.5 --artifacts-dir audit --cluster-report clusters.json --strict-clusters
```

Before paints wider Malgun Gothic glyphs and accumulates 17pt of row drift.
After matches all six native text baselines exactly and preserves identical
searchable text, with zero material clusters under the strict audit.
Full GT/before/after pages, the 5% diff and matched GT/output/diff crops of all
six lines were inspected at 150 DPI. The regular black Calibri 11pt labels
match in shape, alignment, size and row pitch. There are no rules, borders,
shapes, emphasis, rotation, clipping or overlap. The residual pixels are sparse
glyph-edge rasterization specks, not displacement. Progressive quality-86
JPEGs have metadata stripped and density reset to 150 DPI. PDFs, native export
staging and inspection crops are preserved outside the worktree; native export
timestamps can change PDF hashes.
