# Issue #1510 visual audit

Source: `Gift Budget and Tracker1.xlsx` (`25f5dc75dab19ea12042979a61842314ddc226e3e45d447e36b2a2a104112613`). The baseline and candidate were built with the identical tracked `Cargo.lock` (`750f50c507b281f84b945b09573333ffa02dc45cddd566481af8535ae101e612`) and rendered fresh at 300 DPI. The compared native page is worksheet page 2, an A3 landscape page.

## Before-change enumeration

- Page count/order: two pages in the native export and both office2pdf PDFs; worksheet order is unchanged.
- Element presence: title panel, explanatory copy, chart, legend, tracker table, three gift thumbnails, and footer are all present.
- Position: the issue #1510 defect is the left footer origin. The native trace places `_x000D_` at `60 * 0.82 = 49.2pt`; the baseline seats it at `50pt` (`dx +0.8pt`). The independent uniform A3 body-fill transform found by the audit is tracked in #1538.
- Size: page content, shapes, images, and footer type retain their existing sizes. The footer prefix is already 8.2pt after issue #1210.
- Rotation/flip: no page-2 element changes rotation or flip.
- Fill: the title/sidebar, chart bars, and tracker fills are present with their existing colours.
- Stroke/border: chart axes/gridlines, legend marks, and tracker boundaries are present. The remaining edge pixels are covered by the strict cluster dispositions.
- Text content: all 34 native text lines match; no line is missing or extra.
- Font family/weight/style: existing Segoe UI/Aptos substitutions and emphasis are unchanged. No new bold, italic, or underline mismatch appears.
- Text colour: footer prefix colour and body colour already match the issue #1210 correction.
- Alignment: body, chart, and tracker alignment are unchanged; the footer alone is 0.8pt right.
- Line/paragraph spacing: all matched line pitches remain below the 0.5pt fine gate.
- Clipping/overflow: no content is newly clipped or overflows; the footer stays inside the native print box.
- Hairlines (at most 1pt): chart gridlines/axes and tracker rules are present at matching positions and dash patterns in the 300 DPI Poppler and MuPDF pages.
- Weight/emphasis inventory: the page title, overview/tracker headings, legend labels, and dark occasion labels keep the same visible emphasis. No italic or underlined run is present in the affected footer.

## After-change audit

- The candidate trace emits the footer prefix at `transform="1 0 0 1 49.2 817.1643"`; its 8.2pt text matrix is unchanged, proving the size is applied exactly once.
- `compare_layout.py --audit --fine-shift 0.5` matches 34/34 page-2 lines with no fine or large text shift, no missing/extra line, no wrap/reflow change, and no visibility mismatch. The full report is `layout-audit.json`; its overall exit remains nonzero only for the independent A3 body-fill geometry tracked in #1538.
- `compare_render.py --page 2 --dpi 300 --fine-shift 0.5 --audit --strict-clusters` passes with 50/50 clusters dispositioned in `render-clusters-page-2.json`.
- The before/after text-layer census is identical (1,054 normalized characters). The native export has two extra extracted spaces only because its chart legend encodes `Birt hday Bu dget`; the visible legend reads `Birthday Budget` on both sides, so this is a native-PDF extraction artifact rather than converter text loss.
- Full-page and matched footer crops from both Poppler and MuPDF show the same element census and confirm that the old displaced-footer cluster is gone. Forty-nine cluster IDs are unchanged from the already reviewed issue #1210 census; the one new compact tail-glyph cluster is classified as glyph-edge rasterization.
- Re-running the complete checklist finds no remaining visible deviation attributable to issue #1510. Residual raster clusters are explicitly accepted as glyph-edge rasterization, photo resampling, or shape-edge antialiasing; the separate trace-level body geometry is tracked in #1538.

## Evidence

- Full page: `gt.jpg`, `before.jpg`, `after.jpg`
- Matched Poppler and MuPDF footer crops were rendered and inspected locally at full resolution; the PR contract retains the canonical full-page JPEG evidence only.
- Vector/layout audit: `layout-audit.json`
- Strict pixel census: `render-clusters-page-2.json`
- Cluster dispositions are embedded in `render-clusters-page-2.json`.
- Text-layer censuses were generated and inspected locally; their results are recorded above.
