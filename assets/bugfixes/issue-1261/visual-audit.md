# Issue #1261 visual audit

Source: `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
(`2b4a2d8dceda58758593c88409875efbda05780559154c02bd13fef4f7a1c65b`).
The baseline and candidate use the identical tracked `Cargo.lock`
(`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`).
The clean current-main baseline PDF is
`c1e72f694108c6a80967ef2f233f2dbf41a0aaeee46d99b2e6d34709bcf75e7a`;
the candidate PDF is
`9050ee63df62a7dcf93fdf1694698fa555b075dbe33ea68b627179748ba71588`.
The native reference is a fresh Excel for Mac 16.112.3 export
(`923f0a5925957599e256100d185655649625c32ff6b306bd62a921e0590bfcc8`).
Page 2 is the `Monthly college budget` worksheet, fitted to A3 portrait at
0.78 scale.

## Before-change enumeration

- Page count/order: native and baseline contain the same two worksheets in the
  same order; page 2 is the A3 budget sheet.
- Element presence: the heading, three chart cards, cash-flow chart, selector
  band, budget table, all fills/rules, and the sensitivity footer are present.
  All ten x14 sparkline destinations (`Q28`, `Q29`, `Q37`, `Q40`, `Q45`,
  `Q49`, `Q53`, `Q59`, `Q67`, and `Q72`) are empty in the baseline.
- Position and size: because the ten sparkline images are absent, none has a
  measurable baseline frame. Existing sheet/table geometry is unchanged by
  this issue; its independently measured residuals are listed below.
- Rotation/flip: the native sparklines are unrotated and unflipped; the patch
  does not alter any existing element transform.
- Fill: the native sparkline bitmaps have opaque white backgrounds and a
  theme-resolved `#29744F` series. The baseline has no paint in their cells.
- Stroke/border: each native sparkline is a one-pixel solid line; no marker,
  axis, dash, or border belongs to the reported x14 group. Existing table and
  chart rules are unchanged.
- Text content: the baseline has 2,611 whitespace-normalized extracted
  characters, versus 2,606 in native. The difference is pre-existing numeric
  overflow/text extraction, not sparkline text; sparklines add no text.
- Font family/weight/style and text colour: no text style is changed. Native
  and output both embed Trebuchet MS/Trebuchet MS Bold for the table.
- Alignment and line/paragraph spacing: no text alignment or spacing is
  changed. The baseline's fitted-row seat residual is tracked in #1545.
- Clipping/overflow: no new clipping occurs. The two totals where Excel emits
  `#####` remain independently tracked in #1263.
- Hairline inventory (at most 1pt): all ten native sparkline strokes are
  absent before the fix. Existing chart axes/gridlines and table rules remain
  present; chart-gridline presence/alpha defects are #1271 and #1274.
- Weight/emphasis inventory: the page heading, section headings, total rows,
  and category labels keep their existing bold/regular states. No italic or
  underlined run is introduced or removed.

## After-change audit

- The package reader extracts the worksheet's x14 extension before the main
  spreadsheet parser can discard it. It resolves all ten source ranges and
  destination cells, `displayEmptyCellsAs="gap"`, and the theme-6/tint colour
  to `#29744F`.
- The renderer reconstructs Excel's declared-space 42 x 18 opaque bitmap and
  then applies the 0.78 print scale. `mutool draw -F trace` gives the same
  `31.98 x 14.04pt` frame at `x=729.30pt` for native and candidate. Their ten
  y origins agree to trace precision: `443.82`, `458.64`, `577.20`, `621.66`,
  `695.76`, `755.04`, `814.32`, `903.24`, `1021.80`, and `1095.90pt`.
- The constant `Q40` series is pixel-identical to the native 42 x 18 image.
  Per-image AE for the other nine source rasters is bounded at 64-125 of 756
  pixels and hugs the one-pixel line edge; values, colour, extrema, frame, and
  line presence agree. Poppler and MuPDF 300-DPI Q-column crops were inspected
  side by side at full scale and show no missing, displaced, clipped, or
  recoloured sparkline.
- A clean baseline-to-candidate 300-DPI census matches all 62 text instances
  exactly and finds exactly ten material clusters, one in each sparkline cell.
  No table, chart, text, fill, border, footer, or page-layout pixel changes.
  The normalized selectable-text count remains exactly 2,611 characters.
- `compare_text_layer.py --page 2 --json` likewise reports identical baseline
  and candidate census/content (2,784 normalized codepoints). Native versus
  candidate retains the pre-existing four-space/eight-codepoint content delta
  caused by tracked chart-label and numeric-overflow output, not by the
  non-text sparkline images.
- The native-to-candidate `compare_layout.py --audit --noise-floor 0.5
  --fine-shift 0.5` report retains the expected pre-existing findings: the
  fitted cell-seat cadence (#1545), numeric overflow (#1263), chart manual
  layout/tick/category/no-fill/grid/frame/alpha defects (#1265, #1266, #1267,
  #1270, #1271, #1272, #1274), and marker size/fill (#1185). No finding is
  attributable to the sparkline implementation.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5 --strict-clusters`
  passes with all 501 native-to-candidate clusters explicitly dispositioned in
  `render-clusters-page-2.json`. The twelve compact clusters belonging to the
  nine non-identical sparkline rasters are accepted only from the exact frame,
  colour, value, per-image, and dual-renderer crop evidence above. The flat
  sparkline produces no native-to-candidate cluster.
- Re-running the full checklist finds no remaining deviation attributable to
  #1261. Every other visible deviation on the same comparison has an open issue
  reference above.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg` (300 DPI, progressive JPEG,
  quality 86, metadata stripped)
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Matched Poppler and MuPDF full-resolution Q-column crops, native image
  extractions, per-image diffs, and baseline/candidate diffs were inspected
  locally; the canonical repository evidence retains the full-page images and
  machine-readable reports.
