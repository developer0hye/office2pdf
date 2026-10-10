# Issue 2058 visual evidence

Authored synthetic fixture: `tests/fixtures/docx/issue-2058-empty-paragraph-marks.docx`. Native reference exported on
2026-10-10 with Microsoft Word 16.113.4 on macOS 26.6.2 (25G83).
`pdfinfo` reports one 349.92 x 600pt page, Quartz PDFContext, PDF 1.3 for
Word and one 350 x 600pt page, Typst 0.15.1, PDF 1.7 for the current converter.

SHA-256 from `shasum -a 256`:

- Source: `338ad7d989d64c788bca67435c900e50518f53dfe6980436ccc910aef2464ad3`
- Native GT PDF: `ce983a9339c45fd32fc0258aa56ae716d18529b19e5c28b8987d9a6457bdc4e9`
- Current output PDF: `cd480cfcc5a4ecedae2b87155107c58154ac3afc3aa67dc278e2740cbc30289c`

The before converter was built from `6a296de889c0119cb59875a96b2292204436d957` with
`cargo build --locked --profile ci -p office2pdf-cli`. Current output uses
this PR. Build products and original PDFs are retained outside the worktree.
Native PDF export timestamps can change the hash on a subsequent export.

Reproduction from the repository root (stage inside Word's own container):

```sh
stage="$HOME/Library/Containers/com.microsoft.Word/Data/pr-review-2059"
mkdir -p "$stage"
cp tests/fixtures/docx/issue-2058-empty-paragraph-marks.docx "$stage/"
osascript scripts/macos/export_word_pdfs.applescript "$stage/gt" empty-paragraph-marks "$stage/issue-2058-empty-paragraph-marks.docx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-2058-empty-paragraph-marks.docx -o after.pdf
pdfinfo "$stage/gt/empty-paragraph-marks.pdf"
pdfinfo after.pdf
shasum -a 256 "$stage/gt/empty-paragraph-marks.pdf" after.pdf
/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' '/Applications/Microsoft Word.app/Contents/Info.plist'
sw_vers
```

The exported reference was copied to the audit's `gt.pdf`; the current
converter produced `after.pdf`. Both were passed to `compare_layout.py
--json --audit --fine-shift 0.5` and `compare_render.py --fine-shift 0.5
--strict-clusters` at 150 DPI. JPEGs preserve the rendered pixel dimensions,
use progressive quality 86 and retain only a reset 150 DPI density.

The fixture has empty paragraphs with Arial 12pt and 20pt paragraph marks,
without a line-spacing value in either paragraph properties or defaults.
Before reserves a flat 12pt and puts the final text baseline 12.86766pt above
Word. After follows the mark's line metrics: all four text lines match within
0.09234pt vertically, with no missing/extra/changed-wrap lines or fine shifts.

The text layer is identical. Full pages, the 5% diff and matched crops of all
four text regions were inspected at 150 DPI. Remaining differences are sparse
glyph-edge rasterization; no material cluster reaches the 20pt2 report floor.
The strict report passes with zero clusters. Black Arial regular 12pt text
stays legible, and there are no drawings, rules, emphasis, clipping or overlaps.
