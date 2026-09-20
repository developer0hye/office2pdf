# Small Malgun Gothic descent evidence

Synthetic source for issue #1815. Each font/size occupies its own 25pt ruled row,
so another cell cannot supply a shared row baseline. The first cell style uses
Arial 4; the visible Arial 4 control uses a separate font record. Native re-zip
and default-font size 8/11 controls preserve every baseline and row boundary.

Native Malgun Gothic seats are 3pt at sizes 8–10 and 4pt at sizes 11–14.
Main at `bd443885` seats each of these seven labels 1pt lower. Before the fix,
the renderer left those entries unmeasured because an older 4pt workbook floor
masked them (#1208).
Gulim/Batang and Arial control floor differences belong to #1814; per-glyph
placement remains tracked in #1659. These measurements do not establish the
unfloored Gulim/Batang descent. The current fix resolves all seven Malgun baseline shifts. Other deviations
remain explicitly tracked; this is not a whole-page equivalence claim.

`compare.jpg` shows native Excel left and exact-main output right at 150 DPI.
All 22 labels and 23 horizontal rules remain present; extraction content matches.
The full pages, pixel diff and all 14 shifted-region crops were inspected.

The after audit retains seven 1pt fine shifts (#1814), with no missing/extra text,
wrap, visibility, visible-fill, rectangle or large-shift finding. All seven
Malgun baselines now match. Full after pages, diff, seven emitted shift crops
and three supplementary paired crops covering every row were inspected.
The three material clusters map by exact ID to #1814 and #1659.

README usage/API instructions do not change for this rendering correction.
