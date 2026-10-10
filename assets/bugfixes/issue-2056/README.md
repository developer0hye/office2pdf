# Issue 2056 visual evidence

Authored synthetic fixture: `tests/fixtures/docx/issue-2056-wrapped-header.docx`. Native reference exported on
2026-10-10 with Microsoft Word 16.113.4 on macOS 26.6.2 (25G83).
`pdfinfo` reports one 250.08 x 600pt page, Quartz PDFContext, PDF 1.3 for
Word and one 250 x 600pt page, Typst 0.15.1, PDF 1.7 for the current converter.

SHA-256 from `shasum -a 256`:

- Source: `3604269c9bf69df15a4b43f159828c8dec9b4c8203f42a4bf2dfd8bc9347b4fe`
- Native GT PDF: `388ad3d11da99ac805a4afc5aa2946617bc9fa0a9c704873048a2ca84a882c5a`
- Current output PDF: `219c968b882957e87175e241a30a256f8d03e6c3d01281c1878de5f6cc13b88e`

The before converter was built from `6a296de889c0119cb59875a96b2292204436d957` with
`cargo build --locked --profile ci -p office2pdf-cli`. Current output uses
this PR. Build products and original PDFs are retained outside the worktree.
Native PDF export timestamps can change the hash on a subsequent export.

Reproduction from the repository root (stage inside Word's own container):

```sh
stage="$HOME/Library/Containers/com.microsoft.Word/Data/pr-review-2057"
mkdir -p "$stage"
cp tests/fixtures/docx/issue-2056-wrapped-header.docx "$stage/"
osascript scripts/macos/export_word_pdfs.applescript "$stage/gt" wrapped-header "$stage/issue-2056-wrapped-header.docx"
cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-2056-wrapped-header.docx -o after.pdf
pdfinfo "$stage/gt/wrapped-header.pdf"
pdfinfo after.pdf
shasum -a 256 "$stage/gt/wrapped-header.pdf" after.pdf
/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' '/Applications/Microsoft Word.app/Contents/Info.plist'
sw_vers
```

The exported reference was copied to the audit's `gt.pdf`; the current
converter produced `after.pdf`. Both were passed to `compare_layout.py
--json --audit --fine-shift 0.5` and `compare_render.py --fine-shift 0.5
--strict-clusters` at 150 DPI. JPEGs preserve the rendered pixel dimensions,
use progressive quality 86 and retain only a reset 150 DPI density.

Native Word and current output have seven header lines and one body line.
Before puts the body between header lines 2 and 3; after keeps it below the
complete header. The text layer is identical. All eight lines match, with no
missing, extra or changed-wrap lines, visibility or rectangle failures.

Remaining: the body baseline is 128.50788pt against Word's 125.76pt, a
2.74788pt excess gap (#2065). All three current material cluster IDs belong
to that displacement and were inspected in matched crops. An independent
three-paragraph unwrapped control on pre-fix main also has a 2.75256pt body
gap, so this residual predates the wrapped-line repair. Header baselines stay
within the 0.5pt audit threshold. No drawings, rules or emphasis styles occur.

The PR is limited to wrapped-line reservation and preserving words across
OOXML run boundaries. Empty header paragraphs and paragraph spacing retain
their previous behavior and remain unresolved under #2056. A native probe
confirmed that the first header paragraph retains space-before and adjacent
after/before gaps collapse; those changes cannot be accepted independently
of the header's drawn paragraph layout.
