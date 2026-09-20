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
3 at 14pt; expense: 6 rows at 10pt) is a whole sheet point. A later session's
category-count probe (see below) adds 6 more native baselines (income
widened to 6 categories); all 24 measured native baselines, across every
chart/size/category-count combination recorded, are whole sheet points. Run
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

- **Income, all 3 sizes (12/12 rows):** `K = 4` at both 8pt and 10pt (native
  did not move at all between those two sizes), `K = 5` at 14pt. Consistent
  with a coarsely quantized ascent-like term, in the same shape as
  `line_category_baseline_pt`'s already-shipped `above = (ascent *
  size).round()`.
- **Expense, 10pt (1/6 rows):** the same reasoning requires `K = 4` (matching
  income at the same size, same font), but 5 of 6 rows actually need `K = 3`
  — a full sheet point less.

That split was measured with `round()`. The category-count probe below (run
in a later session) shows it was the wrong rounding function, not evidence
that `K` depends on the chart.

## What this falsifies (new): a chart-cache-only category-count patch

A probe that patches only `xl/charts/chart3.xml`'s own `c:cat`/`c:val` cache
(`c:ptCount` plus adding or removing `c:pt` entries) is a **complete no-op**
against native Excel: exporting a 5→4-category patch built this way produced
a PDF **layout-identical** to the unpatched control (0 missing/extra/deviant
lines across 59 matched text lines — native Office's own export timestamps
mean the PDF bytes themselves are never identical run to run). Native Excel
re-derives the category axis from the workbook-scoped `CategoriesIncome`
defined-name array literal in `xl/workbook.xml` and the series values from
the live range named in `c:val`'s own `c:f` (`chart_calcs!$D$19:$D$23`, cells
carrying `ca="1"` volatile formulas) — both outside the chart part — and
overwrites the chart's cache with them on open, regardless of what the cache
itself says. office2pdf, which reads the chart-part cache directly with no
formula evaluation, *does* honor a cache-only patch — so this mismatch would
have silently produced a probe row reading "this factor does nothing" against
native while our own code changed correctly, exactly the wrong-answer shape
`scripts/probe_harness.py` is designed to guard against (its guard doesn't
reach here because the factor is expressible in a way that's individually
well-formed, just ineffective).

`category-count-cache-only-noop-probe.json` reproduces this directly: run it
with `--backend office2pdf` and the "4" variant shows a real 15.78pt shift (1
missing line, 4 deviant); run it with `--backend office` against native Excel
and the same variant is layout-identical to the control (0 missing, 0
deviant) -- the same patch, two different verdicts.

A **working** category-count probe therefore has to patch two parts
together: the workbook-level array literal (categories) and the chart's
`c:f` value range (values), plus the chart's own cache so office2pdf tracks
the same change. `scripts/probe_harness.py`'s spec model patches exactly one
part per variant, so this session hand-built the package with
`build_category_count_probe.py` instead of a harness JSON spec — a documented
gap in the harness, not a new defect in shipped conversion code.

## New measurement: income chart widened to 6 categories

`build_category_count_probe.py 6 OUT.xlsx` appends a synthetic "probe extra"
category (value 0 — magnitude doesn't matter for a baseline-placement probe)
to the income chart, holding font size fixed at the chart's declared 10pt.
This drops `row_sheet_pt` from 23.12pt (5 categories) to **19.27pt** —
closely bracketing the expense chart's 19.60pt row height from the other
side, the exact comparison the previous session flagged as needed to tell
whether the `K` discrepancy tracks row height/category count or is
chart-specific. Native Excel accepted the two-part patch with no repair
prompt and rendered all 6 categories; `measurements.json`'s
`category_count_probes.6` records the 6 resulting baselines, every one again
a whole sheet point (245, 264, 283, 302, 322, 341).

Four of those six sheet values — 283, 302, 322, 341 — are **identical** to
four of the six-category *expense* chart's own values, despite coming from a
different chart, a different row height (19.27pt vs 19.60pt), and a
different plot rectangle. That is strong direct evidence that native's exact
rounding outcome is driven by a row's geometric position, not by which chart
it belongs to.

## Refined model: `floor`, not `round`, plus a size-only `K`

Recomputing the same decomposition with `floor` instead of `round` for the
row-centre term — `floor(sheet_frame_top_pt + row_top + row/2) + K` — and
testing it against every measured row across five (chart, size)
configurations (income @ 5 categories x 3 sizes, expense @ 6 categories @
10pt, income @ 6 categories @ 10pt via the category-count probe) with `K`
depending **only on font size** (`K=4` at 8/10pt, `K=5` at 14pt, exactly the
values already established for the income chart) reproduces **18 of the 24
measured rows exactly**, including every interior row of every configuration
and the bottom-edge row of both income variants. The previous session's
"expense needs `K=3` for 5 of 6 rows" reading was an artifact of `round`;
under `floor` the *same* `K=4` used for income at 10pt reproduces the same
rows.

The six remaining exceptions are still real and are always the plot
rectangle's **edge rows** (adjacent to the plot's own top or bottom
boundary), always off by exactly one whole sheet point — the income chart's
topmost row at each of its 3 measured sizes, plus both edges of the expense
chart:

| Chart | Edge row | Direction |
| --- | --- | --- |
| Income (5-cat, all sizes) | top only | `K+1` |
| Income (6-cat probe) | top only | `K+1` |
| Expense (6-cat) | top **and** bottom | `K-1` |

Income's bottom edge matches the interior-row formula exactly in both
category counts; expense's does not. **Falsified:** naively snapping the
plot rectangle's own continuous top/bottom edges to the nearest whole sheet
point (`round()`), recomputing `row` from the snapped span, and reapplying
`floor(centre) + K` does not explain this — for the income chart it leaves
the top-row deviation exactly as large and *introduces a new one-point miss*
on a previously-exact interior row (`from savings`, 5-category), so a plain
independent-edge-snap is ruled out as the mechanism, though "the plot
rectangle's edges get special treatment" remains the live, narrower lead.

## What's needed before this can ship

The remaining question is now narrowly scoped to *why* the edge rows deviate
and in which direction — not whether `K` is chart-specific (it isn't) or
which rounding function applies to interior rows (`floor`, settled). Shipping
a fix before that edge-row rule is isolated would mean guessing at it —
still the kind of unverified constant this project's methodology asks not to
ship.

## Reproduction

```sh
python3 scripts/probe_harness.py assets/validation/issue-1621/category-size-probe.json --backend office
python3 scripts/probe_harness.py assets/validation/issue-1621/category-count-cache-only-noop-probe.json --backend office2pdf  # shows a real change
python3 scripts/probe_harness.py assets/validation/issue-1621/category-count-cache-only-noop-probe.json --backend office     # shows a no-op
python3 assets/validation/issue-1621/build_category_count_probe.py 6 /tmp/income-6cat.xlsx  # then export via Excel/office2pdf and extract baselines by hand
python3 assets/validation/issue-1621/verify_measurements.py
```

The size probe targets the **income** chart (`xl/charts/chart3.xml`) at 8pt
and 14pt only. The expense chart (`chart4.xml`) cannot be size-probed past
its own declared 10pt without hitting #1756 (a fixed `c:manualLayout` plot
rectangle that doesn't grow with a larger category-label gutter) — an
unrelated, independently filed defect that would otherwise contaminate this
measurement. An 18pt income-chart probe was tried and discarded for a
different reason: at 18pt native rotates or drops labels for a crowded axis
(`chart_category_labels_rotated`'s own regime), which is a different code
path than the flat baseline placement this issue is about, so that data
point isn't in `measurements.json`.

The category-count probe isn't wired into `probe_harness.py` (see "What this
falsifies (new)" above) — `build_category_count_probe.py` writes the package,
which then needs a manual native Excel export (staged inside
`~/Library/Containers/com.microsoft.Excel/Data/probes/`, per
`scripts/probe_harness.py`'s own container-sandbox notes) and a manual
`mutool draw -F trace` baseline extraction, because `compare_layout.py`'s own
report only surfaces deviations past its noise floor, not every matched
line's raw position.

`measurements.json` records the raw native and office2pdf baselines this
README's numbers are computed from, including the internal geometry
(`row`, `row_top`, `plot_y`, `plot_h`, `sheet_frame_top_pt`) read directly
off `generate_chart_axis` via temporary instrumentation (not shipped; removed
before that commit). The category-count probe's `row_sheet_pt` for N=6 is
derived analytically from the already-recorded `plot_h_sheet_pt` (unchanged
by category count) rather than re-instrumenting the Rust code.
