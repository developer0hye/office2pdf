# Issue #1266 visual audit

Source: `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
(`2b4a2d8dceda58758593c88409875efbda05780559154c02bd13fef4f7a1c65b`).
The clean current-main baseline and candidate use the identical tracked
`Cargo.lock`
(`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`).
The baseline PDF is
`3511ee9e9bbeaf8becaebc15b100edd24bffa4dee1133f919dfcc36d87b7bf46`;
the candidate PDF is
`d7ea150b671bb850ddd5a904ff3f207360cf755bd47f2bf6ffcd09cbdcac8999`.
The native reference is a fresh Excel for Mac 16.112.3 export
(`923f0a5925957599e256100d185655649625c32ff6b306bd62a921e0590bfcc8`).
`check_gt_integrity.py` reports two printable worksheets, one hidden
worksheet, two PDF pages, and no invalid reference condition. Page 2 is the
`Monthly college budget` worksheet, fitted to A3 portrait at 0.78 scale.

## Before-change enumeration

- Page count/order: native and baseline contain the same two printable
  worksheets in the same order; page 2 is the A3 budget sheet.
- Element presence: the title, three summary cards, four charts, selector
  band, budget table, ten sparklines, rules, fills, and sensitivity footer are
  present. This fix adds or removes no element.
- Position and size: the `january income:` and `january expenses:` horizontal
  value-label lines sit 3.452pt and 3.465pt above their native printed
  baselines. The corresponding plot bottoms differ from native by only
  0.600pt, isolating the remaining 2.85pt printed error to the label band.
  Undoing the exact 0.78 print scale gives a 3.65 chart-point label-gap error.
- Rotation/flip: no element is rotated or flipped by the patch.
- Fill: page, chart, table, total-row, and sparkline fills are unchanged.
- Stroke/border: all separator, chart, table, and sparkline stroke widths,
  colours, positions, and dash styles are unchanged by the patch.
- Text content: native, baseline, and candidate retain the same value-tick
  strings. The baseline-to-candidate text-layer census is identical, including
  all 2,782 normalized content characters.
- Font family/weight/style and text colour: every worksheet and chart run
  keeps its existing face resolution, emphasis, size, and colour. The native
  reference's recorded Calibri substitution remains an integrity annotation,
  not a changed candidate condition.
- Alignment and line/paragraph spacing: worksheet-cell alignment and spacing
  are unchanged. Only the two horizontal value-label lines move down.
- Clipping/overflow: the lower labels remain inside the summary-card area and
  create no new clipping or overflow.
- Hairline inventory (at most 1pt): the top separator; two summary-card
  dividers; chart axes, tick marks, and gridlines; cash-flow series stroke;
  budget section and row rules; two blue total-row rules; and ten sparkline
  strokes are present. Full-resolution 300-DPI crops confirm that this text-only
  change alters none of their geometry or dash patterns. Existing gridline
  presence and opacity deviations remain tracked in #1271 and #1274.
- Weight/emphasis inventory: the page title, `CASH FLOW`, selector label,
  table section and category headings, and total rows remain bold. Body rows
  and chart labels remain regular; no italic or underlined run changes.

## Layout rule probe

- The reported workbook's value-axis labels resolve to 10pt. Native Excel
  seats the zero-label baseline 15.03 chart points below the inner plot bottom;
  translating that baseline through the existing Typst text box requires a
  7.65pt box-top gap.
- A focused regression using the real chart frame and stated inner plot
  rectangle fails before the fix with a 4.00pt gap, then passes at 7.65pt.
- PowerPoint retains its separately measured size-dependent formula. Excel
  chartsheets and Word charts retain their prior 4pt fallback, so the measured
  worksheet rule does not leak into other hosts.

## After-change audit

- The two printed value-label lines move down by 2.84693pt and 2.84699pt,
  exactly the 3.65 chart-point correction at the sheet's 0.78 scale. Their
  native baseline errors fall from -3.452pt/-3.465pt to -0.605pt/-0.618pt.
  That remainder follows the already-tracked fitted drawing origin in #1542;
  the chart-relative label gap itself now matches the measured rule.
- The clean baseline-to-candidate layout census matches all 62 text lines with
  no missing or extra line. Only the two horizontal value-label lines move;
  both keep exactly the same x position and width. All 847 canonical
  rectangles match with zero geometry delta, and visibility, fill visibility,
  wraps, reflow, and text-token content are identical.
- The clean baseline-to-candidate 300-DPI 5%-fuzz pixel comparison contains
  only the two corrected label bands and measures 11,003 changed pixels
  (0.063224%). Ink coverage and colour distribution are identical, and no
  baseline-to-candidate component reaches the material 20pt2 cluster floor.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5
  --strict-clusters` passes with all 497 native-to-candidate material clusters
  explicitly dispositioned. Of these, 415 map to open issues, 74 are inspected
  shape-edge antialiasing, and eight are sub-0.5pt matching glyph-edge
  rasterization. There are no duplicate, unknown, or undispositioned clusters.
- The material cluster IDs are unchanged from the prior clean-main audit: the
  label correction affects only smaller glyph components. Full-resolution
  crops show that the 14 clusters formerly assigned to #1266 sit on category
  glyphs and bar edges; with the chart-relative label gap fixed, their
  remaining fitted-drawing offset is assigned to #1542.
- Re-running the full checklist finds no remaining deviation attributable to
  #1266. Other visible differences remain tracked in #1185, #1267, #1270,
  #1271, #1272, #1274, #1542, and #1545.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg` (300 DPI, progressive JPEG,
  quality 86, metadata stripped)
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Matched full-resolution page, top-chart, value-label-band, baseline shift,
  difference, and all 54 native fine-shift crops were inspected locally; the
  repository retains the canonical full-page images and machine-readable
  reports.
