# Issue 2054 visual evidence

Authored synthetic fixture: `tests/fixtures/docx/issue-2054-default-language.docx`. Native reference exported on
2026-10-10 with Microsoft Word 16.113.4 on macOS 26.6.2 (25G83).
`pdfinfo` reports one 250.08 x 600pt page, Quartz PDFContext, PDF 1.3 for
Word and one 250 x 600pt page, Typst 0.15.1, PDF 1.7 for the current converter.

SHA-256 from `shasum -a 256`:

- Source: `8e27ff891383ec5acc4ff52492e9136aa0cf67c923914342d3ad965c6ccbdb7d`
- Native GT PDF: `7342a33b10d5f0be91d036f19a7f39c3face2a290de1577daaf0bd49c31ebda8`
- Current output PDF: `7771bcf207df12f217880321dd3bc5a68383d915bb95f509d7e28603dcb818ee`

The before converter was built from `6a296de889c0119cb59875a96b2292204436d957` with
`cargo build --locked --profile ci -p office2pdf-cli`. Current output uses
this PR. Build products and original PDFs are retained outside the worktree.
Native PDF export timestamps can change the hash on a subsequent export.

Reproduction from the repository root (stage inside Word's own container):

```sh
stage="$HOME/Library/Containers/com.microsoft.Word/Data/pr-review-2055"
mkdir -p "$stage"
cp tests/fixtures/docx/issue-2054-default-language.docx "$stage/"
osascript scripts/macos/export_word_pdfs.applescript "$stage/gt" default-language "$stage/issue-2054-default-language.docx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-2054-default-language.docx -o after.pdf
pdfinfo "$stage/gt/default-language.pdf"
pdfinfo after.pdf
shasum -a 256 "$stage/gt/default-language.pdf" after.pdf
/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' '/Applications/Microsoft Word.app/Contents/Info.plist'
sw_vers
```

The exported reference was copied to the audit's `gt.pdf`; the current
converter produced `after.pdf`. Both were passed to `compare_layout.py
--json --audit --fine-shift 0.5` and `compare_render.py --fine-shift 0.5
--strict-clusters` at 150 DPI. JPEGs preserve the rendered pixel dimensions,
use progressive quality 86 and retain only a reset 150 DPI density.

Native Word and the current converter have nine lines. The current German
language reaches the PDF catalog and selects German text layout; tracked
historical English language and paragraph-mark language do not replace it.
German hyphenation and line-break choices still differ from line 2 (#2064).
All 60 current material cluster IDs were inspected and assigned to that issue.
There are no drawings, rules or emphasis styles in the fixture.

The text-layer report differs because Word encodes hard line-end hyphens and
Typst encodes soft hyphens. Removing either line-end hyphen plus its following
newline and normalizing whitespace reproduces the exact source text on both
sides. No source word is missing. The layout report cannot pair the resulting
lines and reports nine missing/nine extra lines; these are tracked reflow, not
proof of nine missing source lines. With no matched lines, its zero shift counts
are not a geometric parity claim.
