# Horizontal bar-chart value-axis label gap (#1622)

## What this confirms

`EXCEL_WORKSHEET_HORIZONTAL_VALUE_LABEL_GAP_PT` (the box-top gap below a
horizontal bar chart's plot bottom, for the `Spreadsheet` + `Bar` branch of
`horizontal_value_label_gap`) was a flat `7.65pt`, derived in #1549 from a
single 10pt measurement on the `issue_1181_fit_to_height.xlsx` income chart.
A fresh native Excel for Mac 16.112 export of the same chart (page 2, 0.78
fit scale) reproduces the issue's own baselines exactly (283.14pt income,
283.92pt expenses, printed) but shows the *current* main (post typst 0.15,
#1694) output landing 0.785-0.798pt too high — worse than the issue's
0.605-0.618pt, because the flat constant never absorbed a size term and the
typst upgrade shifted the box model slightly.

Isolating the value-axis font size with `office`-backend one-factor probes
(`value-axis-label-gap-size-probe.json`: 8pt and 14pt; `-1200.json`: 12pt),
patching only `c:valAx/c:txPr/a:defRPr@sz` on the income chart (`chart3.xml`,
which #1621 already established avoids the #1756 manualLayout-growth
contamination `chart4` hits), gives:

| size | native "0%" baseline (printed) | clean? |
| --- | ---: | --- |
| 8pt | 281.58pt | yes — no large/rect-geometry shifts vs the 1000 control |
| 10pt | 283.14pt | yes — the issue's own control point |
| 12pt | (unusable) | no — 1 re-wrapped line, native re-lays the label row |
| 14pt | (unusable) | no — 3 large shifts + 4 rectangle-geometry deviations |

Translating each clean native baseline through Typst's own box-top-to-
baseline offset for the same real Trebuchet MS font stack (measured with
`typst compile` against a minimal `box(width: 24pt)[align(center)[text(...)]]`
matching the emitted source exactly: 0.7373047pt per declared point,
confirmed linear at 8/10/14pt) leaves a *required* gap of 8.131547pt at 8pt
and 8.656938pt at 10pt — not flat. Fitting the two clean points gives
`GAP(size) = 6.029984 + 0.262695 * size_pt`, the same PT+EM shape already
shipped for the PowerPoint sibling constant just above it in the source.

Applying the fix and re-running `compare_layout.py --fine-shift 0.01` against
the same fresh native GT: the income chart's six labels match exactly (no
longer listed at a 0.01pt gate) and the expenses chart's six labels land
0.01215pt off (down from 0.798pt) — both charts' remaining residual is at
least two orders of magnitude below the repository's 0.5pt gate, and is
consistent with ordinary native-export quantization between the two charts'
independently-derived plot rectangles.

## What this does not confirm

Two points fit a line exactly, so `0.262695` is a two-point fit, not an
independently verified linear model — no third clean native point was
available in this workbook to check it (12pt and 14pt both trigger native's
own plot re-layout). A future probe on a *different* fixture (no manualLayout
growth trap) at three or more clean sizes would strengthen this.

The fitted constants fold in Trebuchet MS's own Typst box metric. A
value-axis face whose box-top-to-baseline offset differs from Trebuchet MS's
(including the built-in fallback font the `office2pdf` lib test harness's
`default_font_search_paths()` silently substitutes when it cannot find a real
Trebuchet MS, since it only searches Microsoft Office's bundled `DFonts` and
Office bundles no regular-weight Trebuchet MS) carries a proportional
residual against this constant. That is a distinct, unmeasured defect class,
not something this fix addresses.

## Reproduce

```sh
python3 scripts/probe_harness.py assets/validation/issue-1622/value-axis-label-gap-size-probe.json --backend office
python3 scripts/probe_harness.py assets/validation/issue-1622/value-axis-label-gap-size-probe-1200.json --backend office
```
