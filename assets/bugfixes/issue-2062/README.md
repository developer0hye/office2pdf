# Issue 2062 visual evidence

The drawn DOCX header and footer story keeps Word's empty paragraphs,
paragraph gaps and `w:lineRule="auto"` multiples.

Fixture: `tests/fixtures/docx/issue-2056-header-multiples.docx`, the base of
`scripts/probes/issue-2056-header-multiples.json`, unmodified: A4, `w:top`
72pt, `w:header` 36pt, and an Arial 12pt header of one paragraph wrapping to
three lines, two empty paragraphs and `Last`. This is the issue's own
reproduction shape. The GT is that package's native export from
`probe_harness.py --backend office` (Microsoft Word for Mac 16.114.3, macOS
27.0.1 build 26A434, `Producer: Quartz PDFContext`), staged in Word's own
container; the probe's no-patch re-zip control was layout-identical.
`check_gt_integrity.py --source` passes.

SHA-256:

- Fixture: `904e8196fc7a0bae1162395389e703a0c0487e517de965c1af86e5a8c4b6fe0a`
- Word export: `229a5ef3bc26722e59293295b182987a98aad5a3534b8b3e784f4b90125860c9`
- Before output (main `b7b29dcd`): `4bf1e2159702ae956f7df08456882ee11539610846186a7ced450db19ba90aed`
- After output: `a7a1d9e530507bd0327056a8bdef4e90030d627d1fc8eccf670f7a812258f51d`

Both CLIs were built with `cargo build --locked -p office2pdf-cli --profile
ci`. `layout-audit.json` comes from `compare_layout.py --json --audit
--fine-shift 0.5`, and `render-clusters-page-1.json` from `compare_render.py
--dpi 150 --fine-shift 0.5 --strict-clusters`. The JPEGs are pdftoppm 150 DPI
renders, saved as progressive quality 86 without metadata.

Result, as first-character baselines from `mutool draw -F stext`:

| line | Word | before | after |
|---|---:|---:|---:|
| wrapped `Header` lines 1-3 | 47.28 / 60.96 / 74.88 | 47.26 / 61.06 / 74.86 | 47.26 / 61.06 / 74.86 |
| `Last` | 116.16 | 88.66 | 116.25 |
| `Body marker` | 130.08 | 130.05 | 130.05 |

Main dropped the two empty paragraphs from the drawn header, so `Last` sat
27.50pt high, directly under the wrapped lines, while #2225 already started
the body below Word's. After the fix the layout audit has no findings. The
strict render audit's 34 clusters are glyph edges of the second wrapped line
and `Last`, whose baselines sit 0.09-0.10pt below Word's, under its 0.24pt
export grid.

The other probes (`issue-2056-header-spacing`, `issue-2062-footer-spacing`,
`issue-2062-header-multiple-seat`) are pinned by tests; their tables are in
the pull request.
