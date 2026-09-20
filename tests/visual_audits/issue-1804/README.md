# Extra selected-period scatter marker

Public fixture: `tests/fixtures/xlsx/two_line_sheet_footer.xlsx`, page 2. Native Excel omits the selected-period scatter marker; office2pdf paints a cyan circle over the first cash-flow category. The source `xl/charts/chart2.xml` has a markerless line series and two scatter series. Scatter series idx=1 declares a 14pt circle and cached Y=69 at point zero; both scatter series share the category/value axis IDs used by the line chart.

A one-factor native probe adding explicit X=1..13 to both scatter series does not restore the marker. Excel reads the original cached/formula Y value as 69. These observations rule out a simple missing-X explanation and a missing Y value; the precise native suppression rule remains unverified.

Model vision findings: inspected the full native/current page pair and the chart crop. The converter has a cyan circle near the first category where native has a plain line. The reported marker cluster is p2-28a2a48b6db9, bbox x189.12 y305.28 w11.52 h12pt. Other chart, cell-position and sparkline differences remain separate existing issues; this evidence does not claim full-page parity.

The filing image uses the converter at 7a89a402 with Excel's DFonts supplied through --font-path. Its page-2 pixels are identical to the page-2 candidate image used to assemble the comparison; the footer-spacing change only affects page 1. Both sides are 150 DPI. No renderer fix is claimed here.
