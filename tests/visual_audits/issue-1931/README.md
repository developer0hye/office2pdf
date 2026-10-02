# XLSX footer margin floor

`base.xlsx` is a synthetic copy of the public alignment workbook from issue
#1812 with an odd header and footer marker added. The one-factor probes sweep
`pageMargins/@footer` and `@header` independently. Excel for Mac 16.113.3 on
macOS 26.6.2 exported the base and every variant; both re-zip controls were
layout-identical.

The footer sweep found that values from 0 through 0.26in share the same printed
position, 3pt closer to the page edge than 0.3in. At 0.28in the difference is
1pt; at 0.5in it is 15pt in the opposite direction. This matches a 0.25in
minimum followed by the renderer's whole-point floor. The header sweep is
different: 0 through 0.5in match 0.3in, then the header moves at larger values.
Header seating remains tracked by issue #1731.

`explicit-zero-footer.xlsx` is the base with only `@footer` changed to zero.
`visual-source.xlsx` uses that zero plus a 0.3in left margin and no header text,
so the footer audit is not obscured by the independent left-origin and header
differences. It retains the three labelled alignment samples and contains no
private data. `visual-export-probe.json` exports the visual source through
Excel with a layout-identical re-zip control.

The pre-fix visual audit showed one Letter page with all three cell labels and
the footer present, no wraps, clipping, fills, rules or shapes, and matching
text appearance. The footer baseline was 2pt above Excel (and 1pt left); the
fix moves it 3pt toward the native position. After the fix, bottom-aligned text
is +1pt vertically, top-aligned text -1pt, and the footer marker -1pt
horizontally/+1pt vertically. These sub-material residuals are recorded
against the standing tracker #1874.
