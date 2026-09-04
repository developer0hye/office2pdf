# Issue #1263 visual audit

Source: `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
(`2b4a2d8dceda58758593c88409875efbda05780559154c02bd13fef4f7a1c65b`).
The baseline and candidate use the identical tracked `Cargo.lock`
(`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`).
The clean current-main baseline PDF is
`9050ee63df62a7dcf93fdf1694698fa555b075dbe33ea68b627179748ba71588`;
the candidate PDF is
`ea28d82a406433f0da43cad236e62c1deb46e5299ffe66b2e080215409b3efe3`.
The native reference is a fresh Excel for Mac 16.112.3 export
(`923f0a5925957599e256100d185655649625c32ff6b306bd62a921e0590bfcc8`).
Page 2 is the `Monthly college budget` worksheet, fitted to A3 portrait at
0.78 scale.

## Before-change enumeration

- Page count/order: native and baseline contain the same two worksheets in the
  same order; page 2 is the A3 budget sheet.
- Element presence: the title, three summary cards, four charts, selector band,
  budget table, ten sparklines, rules, fills, and sensitivity footer are all
  present. No element is added or removed by this fix.
- Position and size: baseline and candidate have the same 675 canonical
  rectangles and the same positions and sizes for all 60 unaffected text
  lines. The two changed cells keep their existing O-column boxes.
- Rotation/flip: no element is rotated or flipped by the patch.
- Fill: all page, chart, table, total-row, and sparkline fills are unchanged.
- Stroke/border: chart axes, chart gridlines, section rules, cell borders,
  total-row rules, and sparkline strokes are unchanged.
- Text content: O37 prints `19,150` and O72 prints `19,449`, whereas native
  Excel prints `#####` in both cells. The four-digit O32 value `7,500` fits and
  remains visible.
- Font family/weight/style and text colour: the three cells use Trebuchet MS
  Bold 10pt with the same total-row colour. The patch does not alter their
  style or any other text run.
- Alignment and line/paragraph spacing: the cells remain right-aligned. No
  line or paragraph spacing changes.
- Clipping/overflow: the two five-digit annual totals visibly overrun Excel's
  usable fixed-format width instead of taking Excel's hash replacement.
- Hairline inventory (at most 1pt): the top separator; two summary-card
  dividers; chart axes and gridlines; line-chart stroke; budget section and
  row rules; two blue total-row rules; and ten sparkline strokes are present.
  Their baseline-to-candidate pixels and geometry are unchanged.
- Weight/emphasis inventory: the page title, selector label, table section and
  category headings, and total rows remain bold. Body rows remain regular;
  no italic or underlined run is introduced or removed.

## Native rule probe

- One-factor Excel exports held style, format, and content constant while
  varying O37's value or only column O's width.
- At the original 40pt cell width, `999`, `1,000`, and `9,999` remain visible;
  `10,000`, `19,150`, `99,999`, and `100,000` become five hashes.
- At an effective 42pt cell width, `19,150` becomes six hashes. At 43pt it is
  visible. The exact-width equality therefore fits; only a strictly wider
  measured value overflows.
- Built-in format 38, `#,##0_);[Red](#,##0)`, reserves the hidden advance of
  `)` through its `_x` control. Trebuchet MS Bold 10pt advances each digit or
  hash by 6pt and the comma or closing parenthesis by 4pt after Excel's
  per-glyph point snapping. Thus `19,150` measures 38pt including the hidden
  parenthesis; the original 35pt usable box holds five whole 6pt hashes.

## After-change audit

- Fixed-format numeric cells now measure their formatted text plus `_x`
  skip-width glyphs against the usable cell width. `General`, text, and
  conditional sections stay on their existing paths. Resolved font metrics
  are weight-aware; the issue face has deterministic reference advances for
  wasm or hosts without Trebuchet MS.
- O37 and O72 each print exactly five hashes, matching native Excel. O32 keeps
  `7,500`; adjacent values, the `% INC` column, cell backgrounds, total rules,
  and the two sparklines remain unchanged in full-resolution 300-DPI crops.
- The clean baseline-to-candidate layout census matches all 60 unaffected text
  lines with 0pt x/y drift and all 675 canonical rectangles with 0pt geometry
  drift. The only unmatched lines are the two intended value-to-hash
  substitutions.
- Baseline-to-candidate 300-DPI pixel comparison finds exactly two material
  clusters, both confined to the O37/O72 glyph boxes (48pt2 each). The 5%-fuzz
  AE is 3,180 pixels; no other page pixels change materially.
- Baseline-to-candidate text-layer comparison changes only those two displayed
  strings: normalized content length goes from 2,784 to 2,782 codepoints. No
  control-character class changes.
- Native-to-candidate full-page, section, total-cell, and pixel-difference
  images were inspected at 300 DPI. Both corrected cells now show the same
  five hashes as native with the same bold emphasis and colour.
- The native-to-candidate `compare_layout.py --audit --noise-floor 0.5
  --fine-shift 0.5` report retains the pre-existing fitted-row cadence and
  chart differences. They remain tracked in #1185, #1265, #1266, #1267,
  #1270, #1271, #1272, #1274, and #1545; none is attributable to this numeric
  overflow patch.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5 --strict-clusters`
  passes with all 499 native-to-candidate material clusters explicitly
  dispositioned. The prior #1261 audit had 501 clusters; the only two removed
  clusters are the former #1263 numeric values. The remaining census maps 425
  clusters to the open issues above and 74 thin edge clusters to inspected
  shape-edge rasterization.
- Re-running the full checklist finds no remaining deviation attributable to
  #1263. Every other visible deviation on the same comparison has an open
  issue reference above.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg` (300 DPI, progressive JPEG,
  quality 86, metadata stripped)
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Matched full-resolution top-chart, upper-table, lower-table, and total-cell
  crops plus the complete 5%-fuzz diff were inspected locally; the canonical
  repository evidence retains the full-page images and machine-readable
  reports.
