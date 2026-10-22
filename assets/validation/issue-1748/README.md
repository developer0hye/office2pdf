# Excel bottom-legend row seat (#1748)

All numbers are page 2 of native Microsoft Excel for Mac 16.112 exports of the
public #982 workbook (`tests/fixtures/xlsx/issue_1603_gift_budget.xlsx`, SHA-256
`25f5dc75…`) with only the legend's `a:defRPr/@sz` changed (Verdana rows also
change its typeface). They are in chart points: page points / the 0.82 fit.
No new export was made: the data are the committed #1617 batch, the #1748
filing's 6/9/18pt baselines, and the #1240 probes' 7/11/14pt baselines.

## Rule

With `L` the legend face's `hhea` ascent + descent per em and `s` the size:

    key centre    = frame bottom - 9 - 0.495 x L x s
    label top     = round(key centre - L x s / 2 - 1)   (whole sheet point)
    label baseline = label top + round(ascent x s)

The filled keys and the line sample both centre on the key centre. Nine Segoe UI
(6-18pt) and six Verdana (6-14pt) sizes give the same intercept within 0.003pt.
The six Segoe UI baselines need a rise between 0.958 and 1.222pt. The tightest
row is 14pt, 0.042pt inside that window. The rounded ascent drops from 7.55pt
to 6.47pt between 7pt and 6pt, which is why the 6pt row shares the 9pt baseline.
The old three-point cadence put that row one chart point lower.

`verify_seat_rule.py` checks every row; run it from the repository root.

## Limits

- Every row shares one frame and one sheet-grid phase. The unscaled control of
  #1607 has the same fractional frame top, so it is not an independent phase.
- No native PDF of these packages is committed, so the evidence contract's
  `layout-audit.json` must be generated on a Mac (see #1748).
- Label baselines were measured only in Segoe UI. Other faces follow the
  rule's own metrics, and no native export has checked them.
