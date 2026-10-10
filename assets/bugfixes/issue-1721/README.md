# Issue 1721 visual evidence

The five compared pages are, in order: both pages of `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`, the single page of `tests/golden_mocks/business/sources/xlsx/09_expense_report_en.xlsx`, and both pages of `tests/golden_mocks/business/sources/xlsx/04_payroll_ko.xlsx`. The budget reference is a retained native Excel export from 2026-09-21 with an exact source SHA-256 match (`2b4a2d8dceda58758593c88409875efbda05780559154c02bd13fef4f7a1c65b`); the two business source/reference pairs match the tracked golden manifest. References were freshly compared, not re-exported for this review.

Before is main `71e103115ee19b1c72cd765ff04a27b5b8dede6f`; after is `7912c897` (including the merged-total regression correction). Both use the same explicit Excel DFonts directory, including the Aptos face used by native footers. Reproduce with `cargo run --locked --profile ci -p office2pdf-cli -- FIXTURE -o output.pdf --font-path /Applications/Microsoft\ Excel.app/Contents/Resources/DFonts`. Join each PDF set with `pdfunite` in the order above.

The JPEGs stack all five full 150-DPI pages without resizing, with white padding where page widths differ; progressive quality 86, stripped metadata, then reset 150-DPI density. `layout-audit.json` is the combined `compare_layout.py --json --audit --fine-shift 1` report. Strict render reports enumerate every current cluster (21/53/34/7/0) and its open issue; their passing status means complete dispositions, not pixel equality. Original PDFs, full pages/diffs, matched crops, binaries, source hashes and logs are preserved outside the worktree.

## Observations and remaining differences

Full pages, pixel diffs, all content-region crops and all 39 fine-shift crops were inspected. All five pages retain their text, shapes, fills, emphasis and page order; no missing/extra visual text, changed wrap, clipping, rotation or visibility mismatch was found. The four target budget labels now match native baselines; tight expense-report totals remain at 227pt and 241pt. Month-row vertical error improves from -2.126pt to +0.780pt. Default `pdftotext` sequences differ in table/column ordering; whitespace-stripped `pdftotext -layout` content matches exactly across the five pages, with no new invisible characters. This extraction-order residual is retained under #1874 rather than claimed as an exact text-layer pass.

| Current finding | Disposition / source |
| --- | --- |
| Extra cyan cash-flow marker | #1804; `xl/charts/chart2.xml` selected-period scatter series. |
| Thin reconstructed trend strokes | #1784; sheet2 `x14:sparklineGroup` source/destination ranges. All ten trend regions were inspected. |
| 36 right-aligned month-header fine shifts, +1.706 to +3.1305pt horizontally | #1774; sheet2 `C27:N27`, `C31:N31`, `C39:N39`, `numFmtId=164`, Cambria Bold and `horizontal=right`. The 0.780pt vertical remainder is #1874. |
| Two `#####` runs +1.5597pt right; expense title -2pt high | #1874; source total cells retain overflow hashes, and expense `A1` is a bottom-aligned horizontal merge. Both residuals predate this change. |
| Grey/white rule intersections, two visible-coverage findings | #1830; sheet2 border records including `O32`'s white left/thick grey top sides. Raw bounding boxes match within 0.00005pt; visible coverage does not. |
| Bar edges remain about 0.78pt high | #1778; chart frames in `xl/drawings/drawing1.xml`. |
| Small glyph pacing, footer origin, title fill top edge, and table-border edge differences | #1874; page1 text anchors match, title band top is +1pt; payroll centered header anchors differ by up to 0.758pt, with matching baselines. Source fonts/styles and thin borders were inspected; these are not blanket antialiasing waivers. |

Hairlines: page1 has no thin rule; page2 has double grey separators, chart gridlines/card dividers, grey row rules, teal total rules and brown section rules, plus sparklines. Page2 drawing paths use 0.25/0.75/1/2.25pt widths before the sheet transform; table rules use the 0.78pt sheet grid. The business tables use 1pt grid borders; small edge-coverage differences remain in #1874. All are solid; there is no dashed or dotted pattern to preserve. Bold headings/totals match; no italic or underlined runs occur in the inspected content. No new material root cause was found beyond the listed open issues.

Combined PDF SHA-256: gt `2d089074e5cbc728109aca2ea7d28bce3765d5f32d3821bd9c042921dc4487de`; before `ed98460c76a3967fc3d5f620d5ebf3fdb4cd991fcea62651294fa2a3a2e94927`; after `2ea5d6ad67212ad16b3244856f81922e400d81ede04fe1fa78189e7402ed9fc3`.
