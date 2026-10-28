# Issue 2056 visual evidence

This directory holds the evidence for the fourth #2056 fix: Word's header
measure keeps paragraph spacing and `w:lineRule="auto"` multiples above the
body. Earlier evidence stays at the merge commits of its pull requests:
wrapped lines (#2057) at `d65372fb7f0172edc1c8940e947c5cc068743bf6`, exact
lines (#2219) at `6a4b7992b1a592644e713b1443253bba6f551713`, and empty
paragraphs (#2225) at `68c55ebe6b99c2a09e7d5591c12d371f1eb72ce4`.

Fixture: `tests/fixtures/docx/issue-2056-header-last-after.docx`, the
`last-after-480` variant of `scripts/probes/issue-2056-header-spacing.json`,
unmodified. A4, `w:top` 72pt, `w:header` 36pt, and three single-spaced Arial
12pt header paragraphs (`One`, `Two`, `Three`); the last states
`w:after="480"`. The GT is that variant's native export from
`probe_harness.py --backend office` (Microsoft Word for Mac 16.114.3, macOS
27.0.1 build 26A434, `Producer: Quartz PDFContext`), staged in Word's own
container. The probe's no-patch re-zip control was layout-identical to the
base. `check_gt_integrity.py --source` passes.

SHA-256:

- Fixture: `5411fa7097fffe2de476d707937be188432a46d11c29af1839593d1737ab392e`
- Word export: `b663817ea6e88f37788d36d9d8087b8fd02fd45fe3deb9979428d84634e011de`
- Before output (main `e8d7378f`): `55dd20f3d0c58aaa645b4a525ae120bbe1db66973fd76c7f2ed552980eebbfe4`
- After output: `d64a0e399ad7b68e5aa65abbabcca433add97263c94e0e995f8879a813875393`

Both CLIs were built with `cargo build --locked -p office2pdf-cli --profile
ci`. `layout-audit.json` comes from `compare_layout.py --json --audit
--fine-shift 0.5`, and `render-clusters-page-1.json` from `compare_render.py
--dpi 150 --fine-shift 0.5 --strict-clusters`. The JPEGs are pdftoppm 150 DPI
renders, saved as progressive quality 86 without metadata.

Result, as first-character baselines from `mutool draw -F stext`:

| line | Word | before | after |
|---|---:|---:|---:|
| `One` | 47.28 | 47.25 | 47.25 |
| `Two` | 60.96 | 61.05 | 61.05 |
| `Three` | 74.88 | 74.85 | 74.85 |
| `Body marker` | 112.56 | 91.31 | 112.65 |

Main ignored the 24pt `w:after`, so the body sat 21.25pt high, directly under
`Three`. After the fix the layout audit has no findings and the strict render
audit has no material cluster.

The probe tables for every spacing and multiple variant are on #2056 and in
the pull request.
