# Horizontal bar-chart category-label vertical baselines (#1621)

## What this confirms

The issue's own six expense-chart baselines, divided by the sheet's own
0.78 fit-to-height print scale, are **exact integers**: 243, 263, 283, 302,
322, 341 sheet points. That generalizes past the issue's own "alternating
15.60/14.82pt cadence" framing (an artifact of `20 x 0.78` and `19 x 0.78`
aliasing at print resolution) to the real underlying rule: native quantizes
each category label's vertical baseline to a **whole sheet point**, in sheet
space, before the print-scale wrapper — the same regime `excel-worksheet-column-edges-whole-sheet-points`
and `excel-centred-cell-whole-point-seat` already document elsewhere in this
codebase. `measurements.json` extends this to a second chart (income, 5
categories, a different row pitch) and three font sizes (8/10/14pt); every
one of the 18 measured native baselines (income: 4 rows at 8pt + 5 at 10pt +
3 at 14pt; expense: 6 rows at 10pt) is a whole sheet point. Run
`python3 assets/validation/issue-1621/verify_measurements.py` to check this
programmatically (`PASS:` on success).

office2pdf's own *current* (unrounded) baseline is an excellent model of the
underlying continuous geometry: for a fixed chart, `ours_sheet - row_top`
(row_top read straight off `generate_chart_axis`'s own local coordinates) is
constant to 5-6 decimal places across every category, confirming translation
invariance and ruling out any bug in the row-pitch or plot-rectangle math
itself. The only missing piece is the quantization step.

## What this falsifies

A single shared rounding rule applied to office2pdf's own continuous
baseline does **not** reproduce native for both charts:

- `floor(ours_sheet)` reproduces all 6 expense-chart rows exactly, but only
  1 of 5 income-chart rows.
- `round(ours_sheet)` (nearest integer; no measured row lands on an exact
  .5 boundary, so the tie-breaking convention doesn't matter here) reproduces
  all 5 income-chart rows exactly, but only 3 of 6 expense-chart rows.

Neither survives being applied to the other chart, so "the fix is a plain
floor (or round) of the current output" is wrong. `verify_measurements.py`
asserts this split holds against the raw data, so a future change to either
rule's behaviour is caught rather than silently re-introducing the same
guess.

## The more promising, still-incomplete model

Decomposing the continuous baseline as
`ours_sheet = sheet_frame_top_pt + row_top + row/2 + K`, `K` (`ours_sheet -
sheet_frame_top_pt - row_top - row/2`) is constant to 4 decimal places
*within* a chart — expected, since nothing in office2pdf's current code
rounds anything on this path. Rounding `sheet_frame_top_pt + row_top +
row/2` (the row's geometric centre, independent of font size, since `row`
and `row_top` don't depend on it) and a separate, size-dependent integer `K`
and summing them reproduces:

- **Income, all 3 sizes (11/11 rows):** `K = 4` at both 8pt and 10pt (native
  did not move at all between those two sizes), `K = 5` at 14pt. Consistent
  with a coarsely quantized ascent-like term, in the same shape as
  `line_category_baseline_pt`'s already-shipped `above = (ascent *
  size).round()`.
- **Expense, 10pt (1/6 rows):** the same reasoning requires `K = 4` (matching
  income at the same size, same font), but 5 of 6 rows actually need `K = 3`
  — a full sheet point less. Only `transportation` (fractional part `.364`)
  matches `K = 4`; its fractional distance from the nearest integer boundary
  is not the largest of the six (`tuition & fees`, at `.441`, is closer to
  the tie line yet still needs `K = 3`), so proximity to the rounding
  boundary alone does not explain which row is the exception.

That 5/6 vs 1/6 split, with the mismatching rows all off by exactly one
sheet point in the same direction, points at the *reference point* — not
`row_top + row/2` — being wrong for the expense chart's own row height
(19.60pt vs income's 23.12pt), not at the font term. Candidates not yet
tested: the plot rectangle's own top/bottom edges may be independently
snapped to whole sheet points before dividing into `N` row bands (mirroring
`column_value_chrome_y`'s existing snap for column-chart gridlines), which
would make the effective row centre depend on category count in a way this
fixture's fixed 5-vs-6-category comparison can't isolate from a fixed row-
height difference.

## What's needed before this can ship

A category-**count** probe (`c:cat`/`c:val` cache patch, not a one-line regex
like the size probe) on one chart, holding font size fixed, to determine
whether the `K` discrepancy tracks row height/category count (implicating
the plot-rectangle vertical sizing) or is chart-specific for some other
reason. That probe is more invasive to construct than the size probe here
(it needs to rewrite both the category cache and the value cache to stay
internally consistent) and was not completed in this session. Shipping a fix
before that probe would mean picking between `K=3` and `K=4` by guessing
which chart is the "normal" one — exactly the kind of unverified constant
this project's methodology asks not to ship.

## Reproduction

```sh
python3 scripts/probe_harness.py assets/validation/issue-1621/category-size-probe.json --backend office
python3 assets/validation/issue-1621/verify_measurements.py
```

The probe targets the **income** chart (`xl/charts/chart3.xml`) at 8pt and
14pt only. The expense chart (`chart4.xml`) cannot be size-probed past its
own declared 10pt without hitting #1756 (a fixed `c:manualLayout` plot
rectangle that doesn't grow with a larger category-label gutter) — an
unrelated, independently filed defect that would otherwise contaminate this
measurement. An 18pt income-chart probe was tried and discarded for a
different reason: at 18pt native rotates or drops labels for a crowded axis
(`chart_category_labels_rotated`'s own regime), which is a different code
path than the flat baseline placement this issue is about, so that data
point isn't in `measurements.json`.
`measurements.json` records the raw native and office2pdf baselines this
README's numbers are computed from, including the internal geometry
(`row`, `row_top`, `plot_y`, `plot_h`, `sheet_frame_top_pt`) read directly
off `generate_chart_axis` via temporary instrumentation (not shipped; removed
before this commit).
