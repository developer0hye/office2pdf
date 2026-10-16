# Word section columns and Typst grid children

- [Typst grid](https://typst.app/docs/reference/layout/grid/) takes each cell's
  content as a positional argument inside the grid call.
- [Typst 0.15.1 grid implementation](https://github.com/typst/typst/blob/v0.15.1/crates/typst-library/src/layout/grid/mod.rs)
- [docxide-pdf corpus](https://github.com/sverrejb/docxide-pdf/tree/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/tests/fixtures/cases/case80)
  supplies an unequal three-column section with a table and list.

Our unequal-column emitter closed the grid before its content arguments.
The following markup brackets therefore became painted text and ordinary
page flow; the intended cells never received their content. Emit each column
segment as a separate grid argument, then close the call after the last cell.

Compiled controls vary the track widths, gutter and column count. They check
actual text origins, common baselines and exact text, including authored
brackets that must survive. Source-string tests alone missed this defect.

This emission repair does not establish complete Word column fidelity.
Separator rules, nested content sizing and balancing still require native
full-file comparison; case80's before audit records independent deviations.
