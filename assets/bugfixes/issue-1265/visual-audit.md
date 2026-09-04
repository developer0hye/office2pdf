# Issue #1265 visual audit

Source: `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
(`2b4a2d8dceda58758593c88409875efbda05780559154c02bd13fef4f7a1c65b`).
The clean current-main baseline and candidate use the identical tracked
`Cargo.lock`
(`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`).
The baseline PDF is
`ea28d82a406433f0da43cad236e62c1deb46e5299ffe66b2e080215409b3efe3`;
the candidate PDF is
`3511ee9e9bbeaf8becaebc15b100edd24bffa4dee1133f919dfcc36d87b7bf46`.
The native reference is a fresh Excel for Mac 16.112.3 export
(`923f0a5925957599e256100d185655649625c32ff6b306bd62a921e0590bfcc8`).
`check_gt_integrity.py` reports two printable worksheets, one hidden
worksheet, two PDF pages, and no invalid reference condition. Page 2 is the
`Monthly college budget` worksheet, fitted to A3 portrait at 0.78 scale.

## Before-change enumeration

- Page count/order: native and baseline contain the same two printable
  worksheets in the same order; page 2 is the A3 budget sheet.
- Element presence: the title, three summary cards, four charts, selector band,
  budget table, ten sparklines, rules, fills, and sensitivity footer are all
  present. No element is added or removed by this fix.
- Position and size: the cash-flow line plot starts 64.48pt left of Excel's
  plot. Its category-label line starts at 121.974pt rather than Excel's
  186.457pt. The chart XML states a plot rectangle, relative to its 786.71pt by
  75.55pt frame, of x=0.142768482, y=0.074307877, w=0.857231518, and
  h=0.559675204; the renderer instead uses automatic plot gutters.
- Rotation/flip: no element is rotated or flipped by the patch.
- Fill: all page, chart, table, total-row, and sparkline fills are unchanged.
- Stroke/border: the line plot and its automatic gridlines use the wrong plot
  rectangle. The stroke widths and styles themselves are unchanged.
- Text content: native, baseline, and candidate contain the same text tokens.
  The patch only changes chart-label coordinates.
- Font family/weight/style and text colour: chart and worksheet text keep the
  same fonts, emphasis, and colours. The integrity probe records the existing
  Calibri substitution in the native reference.
- Alignment and line/paragraph spacing: worksheet-cell alignment and spacing
  are unchanged. The category labels use their existing automatic vertical
  band, but their horizontal centres do not follow the stated plot rectangle.
- Clipping/overflow: no new clipping or overflow occurs. Existing worksheet
  and chart discrepancies remain independently tracked below.
- Hairline inventory (at most 1pt): the top separator; two summary-card
  dividers; chart axes and gridlines; cash-flow series stroke; budget section
  and row rules; two blue total-row rules; and ten sparkline strokes are
  present. This patch changes only the geometry of the 12 cash-flow plot
  strokes; their widths and dash styles are preserved.
- Weight/emphasis inventory: the page title, `CASH FLOW`, selector label,
  table section and category headings, and total rows remain bold. Body rows
  and month labels remain regular; no italic or underlined run changes.

## Layout rule probe

- The source's `c:manualLayout` uses factor modes and an outer-target plot
  rectangle. Applied to the unscaled chart frame, its expected inner rectangle
  is x=112.317pt, y=5.614pt, w=674.393pt, h=42.283pt.
- A focused visual regression test uses that real chart-frame size and layout.
  Before the fix it fails with x=30pt and w=756.71pt, proving the line-family
  renderer took its private automatic-gutter path.
- After the fix all four plot edges resolve to the stated rectangle within
  0.02pt. Category-label horizontal centres follow it, while their vertical
  baseline remains on the automatic label band. Value labels move with the
  plot and retain their 6pt gutter.

## After-change audit

- The line-family renderer now uses the shared stated-plot-rectangle resolver,
  with the prior automatic rectangle as fallback when no layout is declared.
  Value-label and category-label horizontal positions share the resolved plot
  displacement.
- The month-label line moves from x=121.974pt to x=186.533pt, leaving only
  +0.076pt relative to native Excel. Its baseline has exactly 0pt
  baseline-to-candidate drift and remains only +0.267pt from native.
- The clean baseline-to-candidate layout census matches all 62 text lines with
  no missing or extra line. Only the month-label line moves, by +64.5585pt x
  and 0pt y. All 847 canonical rectangles match; only 12 line-chart strokes
  change geometry. Visibility, fill visibility, wraps, reflow, and text-token
  content are identical.
- The clean baseline-to-candidate 300-DPI 5%-fuzz pixel comparison contains
  only the corrected chart region and measures 149,972 changed pixels
  (0.86175%). Full-resolution chart, upper-table, lower-table, full-page, and
  difference crops were inspected.
- The native-to-candidate layout audit has no text shift at or above 5pt. The
  former 64.48pt month-label displacement is gone; remaining sub-5pt chart and
  fitted-row differences are mapped to their existing issues.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5
  --strict-clusters` passes with all 497 native-to-candidate material clusters
  explicitly dispositioned. Of these, 415 map to open issues, 74 are inspected
  shape-edge antialiasing, and eight are sub-0.5pt matching glyph-edge
  rasterization. There are no duplicate, unknown, or undispositioned clusters.
- The prior #1263 audit had 499 clusters. Nineteen former #1265 displacement
  clusters and two former #1271 gridline clusters disappear; the corrected
  plot location introduces eleven #1271 gridline clusters and eight harmless
  glyph-edge clusters.
- Re-running the full checklist finds no remaining deviation attributable to
  #1265. Other visible differences remain tracked in #1185, #1266, #1267,
  #1270, #1271, #1272, #1274, and #1545. The first #1271 cluster also touches
  the independently tracked #1185 series marker.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg` (300 DPI, progressive JPEG,
  quality 86, metadata stripped)
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Matched full-resolution chart, upper-table, lower-table, full-page, and
  5%-fuzz difference images were inspected locally; the repository retains the
  canonical full-page images and machine-readable reports.
