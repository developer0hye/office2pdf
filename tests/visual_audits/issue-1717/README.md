# Issue #1717 visual audit

The fixture is a reduced derivative of `tests/fixtures/xlsx/1000-customers.xlsx`,
from `my-sample-files/files-hub` (CC0-1.0). In `xl/worksheets/sheet1.xml`,
`<col min="5" max="5" width="21.7" customWidth="1"/>` sets column E, and
`<c r="E2" t="str"><v>Customer Quality Executive</v></c>` supplies the target
value. `xl/styles.xml` defines Normal as 12pt Calibri with `scheme="minor"`;
the XLSX parser resolves that themed font to Malgun Gothic on this host.
Excel's PDF embeds a Malgun Gothic subset from its bundled DFonts.

The native reference is a one-page A4 PDF exported by Microsoft Excel for Mac
16.113.3 on macOS 26.6.2 (build 25G83), using Quartz. Both converter PDFs were
created by Typst 0.15.1 from the same fixture. The pre-fix PDF was made from
base commit `f0f7d6e`; the post-fix PDF contains this branch's implementation.
The issue-local `provenance.json` records the source, exporter, PDF and image
hashes.

At 21.7 characters, the native line is 149.3262pt wide by Malgun Gothic's
fractional `hmtx` advances; Excel's whole-point sheet grid totals 148pt. The
available line width is 147pt. The old ratio estimate was 146.0652pt and did
not trigger continuation, so Typst wrapped the value. The corrected fractional
width gives the spill clip box enough room for the unwrapped line. The page
column continuation decision remains based on Excel's per-glyph whole-point
grid.

The pre-fix layout audit records 2 native lines versus 3 output lines and one
wrap. The fixed output has 2/2 lines, no wrap, missing or extra text, visibility
mismatch, or position shift. The strict render audit uses a 0.5pt fine-detail
threshold at 150 DPI and finds no material clusters. Visual review confirms
all five headers and the E2 value are present, in the same positions and style;
the value is one line as in Excel. The only remaining difference is dispersed
glyph-edge rasterization and antialiasing below the cluster floor.

Reproduction commands for the committed machine reports:

```sh
python3 scripts/compare_layout.py --json --audit --fine-shift 0.5 --noise-floor 0.5 \
  <native-excel.pdf> <fixed-output.pdf> > assets/bugfixes/issue-1717/layout-audit.json
python3 scripts/compare_render.py <native-excel.pdf> <fixed-output.pdf> \
  --page 1 --dpi 150 --fine-shift 0.5 --audit \
  --artifacts-dir target/issue-1717/render \
  --cluster-report assets/bugfixes/issue-1717/render-clusters-page-1.json \
  --cluster-dispositions tests/visual_audits/issue-1717/cluster-dispositions-page-1.json \
  --strict-clusters
```
