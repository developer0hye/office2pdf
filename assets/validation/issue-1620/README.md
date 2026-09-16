# Excel horizontal bar-chart category-label clearance (#1620)

The public fixture `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx` is
unmodified. Its page-2 income chart (`xl/charts/chart3.xml`, a horizontal bar
chart with a `<c:plotArea><c:layout><c:manualLayout>` inner rectangle) draws
its category axis at a declared 10pt Trebuchet MS, and its worksheet prints
at a 0.78 fit-to-height scale. `category-size-probe.json` beside this file
patches only that size, one factor at a time, to 8, 14 and 18pt;
`native-category-label-measurements.json` records the resulting native label
offsets.

## Rule

`chart_category_label_box_w` right-aligns a bar chart's category label inside
a box whose clearance from the plot is `CHART_LABEL_EDGE_PAD_EM` (0.927em),
calibrated from PowerPoint's own export of `bar-chart.pptx` (#998) and shared
by every host. Comparing a fresh native Excel export against that shared box
at 8, 10, 14 and 18pt category-axis size shows Excel keeping a narrower
clearance — the label's own origin sits a further amount to the right of the
shared box, not a flat point offset.

That amount must be measured, and applied, in the same coordinate space this
fix's own correction runs in. `excel_bar_category_label_x_shift_pt`'s result
is added to `plot.dx`, and `write_placed_sheet_anchor` (`typst_gen.rs`) emits
the whole chart inside the fitted sheet's own `#scale(print_scale)` wrapper —
every point this fix adds is a **sheet** point, pre-scale, not a **printed**
point. A first version of this fix measured the em ratio directly against
`mutool`-traced *printed* points and got 0.0262, which under-applies by
exactly this sheet's own 0.78 print scale once rendered (confirmed by
converting with that constant and remeasuring: the applied page effect was
0.2044pt at 10pt, not the printed 0.2610pt gap it was meant to close — `0.262
x 0.78 = 0.2044` to four figures). Dividing each printed measurement by 0.78
before comparing to the declared (unscaled) axis size gives the sheet-space
ratio this addition point actually needs:

| category size | native printed dx0 | / 0.78 = sheet dx0 | sheet dx0 / size |
| ---: | ---: | ---: | ---: |
| 8pt | 0.2083pt | 0.26705pt | 0.03338em |
| 10pt | 0.2610pt | 0.33462pt | 0.03346em |
| 14pt | 0.3677pt | 0.47141pt | 0.03367em |
| 18pt | 0.4733pt | 0.60680pt | 0.03371em |

    EXCEL_BAR_CATEGORY_LABEL_X_SHIFT_EM * size_pt ≈ 0.0336 * size_pt   (sheet points)

The bar rectangles themselves (the plot's value-zero edge) already match
native exactly at every size probed, so only the label's own box moves — the
fix translates that box by the measured amount (`plot.dx +
excel_bar_category_label_x_shift_pt(chart)`) rather than touching the gutter
or plot-rectangle formulas that size the plot.

## Reproduction

Run from the repository root with Microsoft Excel available on macOS:

```sh
python3 scripts/probe_harness.py assets/validation/issue-1620/category-size-probe.json --backend office
```

Convert with `--font-path '/Applications/Microsoft Excel.app/Contents/Resources/DFonts'`
and compare page 2 with `compare_layout.py --audit --fine-shift 0.5`. The
control (unpatched re-zip) must reproduce the base export's plot rectangle and
label positions exactly before any variant row is trusted. The probe's own
`compare_layout.py` runs report printed-point deltas (its output is a real
rendered PDF, already print-scaled); divide by the sheet's fit scale before
comparing to a sheet-space quantity, as above.

## Scope note

Raising the *expense* chart's category-axis size (`xl/charts/chart4.xml`,
patched identically in the fixture's own two-chart layout) past 10pt uncovers
a much larger, unrelated defect: office2pdf keeps a `c:manualLayout` plot
rectangle fixed regardless of the category-label gutter a larger font
demands, while native Excel grows the rectangle inward. That produced tens of
points of divergence at 14–18pt and is filed separately as #1756 — it does
not affect this fixture at its own declared 10pt size, where #1620's residual
reproduces cleanly.

## Reconciling #1620's reported dx with this fix's measured gap

#1620's own body measured `other`/`from savings`/`financial aid`/`other
expenses`/`discretionary` at -0.729 to -0.741 printed points against `main
295bc540e950d8bdbace56840c3861a55f667778`. Re-running
`compare_layout.py --audit --fine-shift 0.5 --noise-floor 0.5` on a fresh
native export against current `main` (`c3bfafd7`) gives each label's *total*
printed dx **before this fix** as only -0.254 to -0.266pt, not the filed
-0.73pt. That is not a measurement error: commit `082fa901` (#1272, merged
2026-09-07, one day after #1620 was filed) pulled an overrunning
`c:manualLayout` plot rectangle back inside its chart area for this exact
chart family, moving `plot.dx` for these two charts on its own, independent
of the category-label-clearance defect #1620 describes.

That -0.254 to -0.266pt is the gap this fix's sheet-space-corrected constant
actually closes: after conversion with `EXCEL_BAR_CATEGORY_LABEL_X_SHIFT_EM =
0.0336`, the same five labels measure **+0.0003 to +0.0078 printed points**
from native — at the noise floor of the measurement, not merely under the
issue's 0.5pt gate. (An earlier, page-space-calibrated draft of this fix used
0.0262 and left a -0.050 to -0.062pt residual; still inside the gate, but a
real, avoidable bias from adding a printed-point ratio in a sheet-space
context — see the "Rule" section above.)
