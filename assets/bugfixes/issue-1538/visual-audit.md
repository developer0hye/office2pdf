# Issue #1538 visual audit

Source: `Gift Budget and Tracker1.xlsx` from #982
(`25f5dc75dab19ea12042979a61842314ddc226e3e45d447e36b2a2a104112613`).
The baseline and candidate use the identical tracked `Cargo.lock`
(`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`).
The candidate was rebuilt from a clean `office2pdf`/`office2pdf-cli` target and
is `ae4c550be64f229f5ce151498c410aeb048c0776d7ef60a05648b967105dea1f`.
The compared native page is worksheet page 2, an A3 landscape page exported by
Excel for Mac 16.112.

## Before-change enumeration

- Page count/order: native and baseline each contain the same two worksheets in
  the same order; page 2 is the fitted A3 sheet.
- Element presence: sidebar, explanatory copy, three gift thumbnails, chart,
  legend, both rose title bands, tracker table, and sensitivity footer are all
  present.
- Position: the eight canonical worksheet-grid fill regions share the issue
  defect. Their baseline transform is `dx +0.185pt`, `dy +0.700pt`; the sidebar
  union starts at `(63.325, 61.380)` instead of `(63.140, 60.680)`.
- Size: each affected baseline fill is `+0.18pt` wider and taller because its
  positive-axis 1pt background bleed was not scaled with the 0.82 fitted sheet.
  The worst baseline edge delta is `+0.88pt`.
- Rotation/flip: no page-2 element is rotated or flipped by this change.
- Fill: all native colours are present, but the rose title bands, pale sidebar,
  and five dark occasion cells have the shared position/extent defect above.
- Stroke/border: chart axes and gridlines plus tracker boundaries are present.
  No dashed or dotted border changes in this issue.
- Text content: all 34 native trace lines match the baseline; no line is missing
  or extra.
- Font family/weight/style: existing Segoe UI/Aptos substitutions are unchanged.
  The page headings and occasion labels remain emphasized; no italic or
  underlined run occurs in the affected regions.
- Text colour: heading, body, chart, tracker, and footer colours are unchanged.
- Alignment: body text already sits within the 0.5pt Excel gate and must not
  follow the fill-only translation.
- Line/paragraph spacing: no matched line pitch exceeds the 0.5pt gate.
- Clipping/overflow: no element is missing, newly clipped, or overflowing.
- Hairlines (at most 1pt): all eleven horizontal chart gridlines, the chart's
  zero-axis rule, the teal series line, and the tracker's horizontal fill
  boundaries are present. The gridlines and table boundaries are solid; no
  dash-pattern mismatch is visible at 300 DPI.
- Weight/emphasis inventory: `GIFT BUDGET AND TRACKER`, `MONTHLY OVERVIEW`, the
  lower tracker heading, and the five white occasion labels retain their native
  visible weight. Body, axis, and legend labels remain regular.

## After-change audit

- Excel constructs the fitted paper box in declared sheet space and then scales
  it. The candidate moves only table paint by `(-0.185, -0.700)pt`, counter-moves
  cell content to preserve its established seat, and scales the 1pt positive-axis
  background bleed and 0.25pt seam overlap by 0.82.
- The sidebar union, both rose title bands, and all five dark occasion fills now
  match their native trace bounds. Six are exact; the two remaining numerical
  maxima are `0.000006pt` and `0.000036pt`, far below the unchanged 0.5pt gate.
- `compare_layout.py --noise-floor 0.5 --audit --fine-shift 0.5` confirms two
  pages in the same order, 34/34 page-2 text lines, zero fine/large text shifts,
  zero missing/extra lines, zero wrap/reflow changes, zero painted-text
  visibility mismatches, and zero visible-fill occlusions. Page 1 has no
  findings. The five remaining page-2 rectangle findings are chart-only:
  fitted drawing-foreground position is #1542 and the independent 0.5275pt
  column width is #1543.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5 --audit
  --strict-clusters` passes with all 46 current clusters explicitly
  dispositioned in `render-clusters-page-2.json`. Forty IDs retain the prior
  reviewed glyph/photo/shape-edge classifications from #1510. The six new thin
  clusters hug corrected fill edges and are one-pixel native-clipping versus
  Typst-antialiasing rims.
- The before/after selectable-text censuses are identical: 1,054 normalized
  output characters and the same two-space delta against native. The native PDF
  alone encodes the visibly correct `Birthday Budget` legend label as
  `Birt hday Bu dget`; this is the unchanged native extraction artifact already
  recorded in #1510, not converter text loss.
- Full-page and matched title/sidebar/tracker crops were rendered at 300 DPI
  with both Poppler and MuPDF and inspected at full resolution. They show the
  corrected fill bounds, unchanged text seats, unchanged chart and image
  content, and no new clipping.
- Re-running the full checklist finds no remaining visible deviation attributable
  to #1538. The trace-level chart residuals are explicitly tracked in #1542 and
  #1543; every material 5% pixel cluster is dispositioned.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg` (300 DPI, progressive JPEG,
  quality 86, metadata stripped)
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Matched Poppler and MuPDF full-resolution crops and 5% diffs were inspected
  locally; the PR contract retains the canonical full-page JPEG evidence only.
