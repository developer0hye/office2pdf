# Issue 2052 visual evidence

Authored synthetic fixture: `tests/fixtures/docx/issue-2052-default-hyphenation.docx`. Native reference exported on
2026-10-10 with Microsoft Word 16.113.4 on macOS 26.6.2 (25G83).
`pdfinfo` reports one 250.08 x 600pt page, Quartz PDFContext, PDF 1.3 for
Word and one 250 x 600pt page, Typst 0.15.1, PDF 1.7 for the current converter.

SHA-256 from `shasum -a 256`:

- Source: `4004b41cff977b75d5381706ba7e721ee53444c78b237af46a8fb95063ad5867`
- Native GT PDF: `99d9b06bab1a266e2274b5f3fe8a22efbbd8c95f1797d444e11000979e13926a`
- Current output PDF: `2260d9e6ad5f1f16aa7c6b174282569ea1e7c316833d907a2988c64500f4c9fd`

The before converter was built from `6a296de889c0119cb59875a96b2292204436d957` with
`cargo build --locked --profile ci -p office2pdf-cli`. Current output uses
this PR. Build products and original PDFs are retained outside the worktree.
Native PDF export timestamps can change the hash on a subsequent export.

Reproduction from the repository root (stage inside Word's own container):

```sh
stage="$HOME/Library/Containers/com.microsoft.Word/Data/pr-review-2053"
mkdir -p "$stage"
cp tests/fixtures/docx/issue-2052-default-hyphenation.docx "$stage/"
osascript scripts/macos/export_word_pdfs.applescript "$stage/gt" default-hyphenation "$stage/issue-2052-default-hyphenation.docx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-2052-default-hyphenation.docx -o after.pdf
pdfinfo "$stage/gt/default-hyphenation.pdf"
pdfinfo after.pdf
shasum -a 256 "$stage/gt/default-hyphenation.pdf" after.pdf
/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' '/Applications/Microsoft Word.app/Contents/Info.plist'
sw_vers
```

The exported reference was copied to the audit's `gt.pdf`; the current
converter produced `after.pdf`. Both were passed to `compare_layout.py
--json --audit --fine-shift 0.5` and `compare_render.py --fine-shift 0.5
--strict-clusters` at 150 DPI. JPEGs preserve the rendered pixel dimensions,
use progressive quality 86 and retain only a reset 150 DPI density.

Native Word has 12 lines; after has 13. Unwanted automatic hyphens are gone,
and normalized searchable content is identical. The remaining justified-wrap
difference is tracked in #2063. All 35 current cluster IDs were inspected in
matched GT/output/diff crops and dispositioned there; no cluster was waived
as antialiasing. There are no drawings, rules or emphasis styles in the fixture.
