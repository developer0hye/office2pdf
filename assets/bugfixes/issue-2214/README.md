# Issue 2214 visual evidence

A header line with `w:lineRule="exact"` is laid out at its stated height, with
the baseline 0.8 of that height below the line's top, whatever the face.

Source: `tests/fixtures/docx/docxide/first_header_exact_line.docx`, the
unmodified docxide-pdf
[`first-header-no-minimum-control`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/first-header-no-minimum-control)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (synthetic, Apache-2.0). Its
`reference.pdf` reports `Creator: Microsoft Word`, one 500 × 500pt page, and
`check_gt_integrity.py --source` passes. The first-page header is Aptos 12pt in
a 160pt exact line, 6pt below the page top.

SHA-256:

- Source: `7bd704d343e6a18e28af36c0289354c9287ae8657fd96e1e33f27e5d4ae722da`
- Word reference: `516625f29371787dd86519c797c639e828af11cd58ca11ba31660d8f40d500ba`
- Before output (main `2796155e`): `fabd71a2a1c597cf7c151a14c63ef460cc1011df1b3ee64ea86d0a1cb169da1b`
- After output: `0a59075d2e37f491671961e11ba21fdcba623cb639cbebe75f31c539650a6c15`

Both CLIs were built with `cargo build --locked --profile ci -p
office2pdf-cli` in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and
`render-clusters-page-1.json` from `compare_render.py --dpi 150 --fine-shift
0.5 --strict-clusters`. The JPEGs are page 1 rendered by pdftoppm at 150 DPI
and saved as progressive quality 86 without metadata. `compare.jpg` is the
defect evidence from #2219.

Baselines from `mutool draw -F stext`:

| line | Word | before | after |
|---|---:|---:|---:|
| `Header` (160pt exact line) | 134.08pt | 18.83pt | 134.00pt |
| `Before` | 177.35pt | 177.26pt | 177.26pt |
| `Second` | 209.35pt | 209.22pt | 209.22pt |

The same source's `first-header-small-control` sibling
(`first_header_small_control.docx`, a 12pt exact line shorter than the face)
moves from 18.83pt to 15.60pt against Word's 15.50pt; its layout audit passes
too.

Remaining: `Header` is drawn in the Noto Sans stand-in for unembedded Aptos,
which is taller and about 8.5% wider than Word's face (#2237).

Corpus scan: the CI-profile CLIs from main and this branch, run over 199
tracked DOCX fixtures and 552 docxide-pdf cases. Only 12 outputs change, all
of them in header lines alone:

- The two tracked fixtures above, and the five docxide `first-header-*` cases.
- `cont-first-*` (three cases, plus the tracked
  `continuous_framed_section_mark.docx`): Times New Roman 14pt in a 16pt exact
  header line moves from 41.42pt to 41.15pt. Word puts it at 41.03pt.
- Scraped
  [`bulgarian_farmland_allocation_order`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/scraped/bulgarian_farmland_allocation_order):
  three 18pt exact header lines move from 45.22, 87.37 and 129.52pt to 49.80,
  67.80 and 85.80pt. Word puts them at 50.03, 68.03 and 86.03pt.

Conversions that fail on main fail the same way here.
