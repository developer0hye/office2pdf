# Explicit worksheet print-scale evidence

Synthetic source for issue #1817. `source.xlsx` declares
`pageSetup scale="85"`; `base.xlsx` omits scale (native 100%). No fit-to-page
flag is set, so the percentage is the sheet's selected scaling mode.

Re-run the one-factor probe with

```sh
python3 scripts/probe_harness.py tests/visual_audits/issue-1817/probe.json --backend office
```

It reproduces the native 85% and 78% exports and passes its unchanged re-zip
control gate. Native row 6 adjacent label origins are 91.8pt apart at 85%,
84.24pt at 78%, and 108pt at 100%; `pageMargins left="0.7"` puts the first
column origin 50.4pt from the page edge.

## What each file records

- `base.xlsx`, `source.xlsx`, `probe.json` — the probe inputs; still current.
- `baselines.json`, `layout.json`, `native-control-baselines.json` — the
  per-label trace capture taken when the defect was filed at `bd443885`. The
  `gt.pdf` blocks still reproduce exactly from a fresh native export, so the
  native figures above remain re-derivable. The `before.pdf` blocks are a
  historical capture of that commit: two of the 40 comparable labels
  (`Bottom row 2`, `Bottom row 3`) have since moved 1pt through the bottom-seat
  fixes that landed between `bd443885` and `e5040555`, so read them as the
  filing record rather than as current output.
- `cluster-dispositions-page-1.json` — the strict diff-cluster dispositions for
  the fix, keyed to the render report in `assets/bugfixes/issue-1817/`.

Current before/after evidence for the fix lives in
`assets/bugfixes/issue-1817/` (`gt.jpg`, `before.jpg`, `after.jpg`,
`layout-audit.json`, `render-clusters-page-1.json`).

## Remaining deviations after the fix

The explicit percentage now reaches the columns, rows, type and table-inherited
cell inset (#1932): the 39 large text shifts and the 17.7% width error are gone,
and the rectangle census matches with zero geometry deviation. What is left on
the 85% comparison is the bottom/top cell seat, which differs by whole points
independently of print scale (#1721, #1814, #1815, and #1874 for the
sub-material remainder).
