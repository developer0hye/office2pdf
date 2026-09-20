# Repeated chart-axis labels reported as a wrap

On page 2 of public `tests/fixtures/xlsx/two_line_sheet_footer.xlsx`, `compare_layout.py` reports a wrap for `0%10%20%30%40%50%0%10%20%30%40%50%`. The source has two independent horizontal bar charts, `xl/charts/chart3.xml` and `chart4.xml`, each with its own `c:valAx` and `c:numFmt formatCode="0%"`.

Native glyph baselines are 283.2336pt and 283.7951pt; the current converter uses 283.1400pt and 283.90785pt. Each independent axis is within 0.12pt of native. Native line grouping combines both sequences (mean baseline 283.51435pt), while output grouping keeps two equal text keys. The audit calls this a wrap despite preserved text and axis placement. The exact matcher branch responsible still needs investigation.

Model vision findings: inspected the full GT/output page, pixel difference and full-scale chart-region crops. Both charts retain six percentage ticks with no text moving to a new chart row. Other differences, including the extra marker and thinner sparklines, are separate issues. The 150-DPI comparison is reused from #1804; its converter page is pixel-identical to the current footer candidate's second page. This is harness evidence, not a renderer fix.
