# Line-legend marker floor rule (#1618)

## Rule (confirmed)

Native Excel for Mac 16.112 seats a line-legend key's marker centre by
flooring the sample stroke's own continuous target to a whole chart point,
the same way `WorksheetMarkerPlacement` already floors/rounds a plot marker
(#1105) — but here it is an **unconditional floor**, not the plot's
odd/even-outline branch (every marker in both datasets below has no explicit
outline, `<c:marker><c:spPr><a:solidFill/></c:spPr></c:marker>`, so the
outline-dependent branch is unprobed, not ruled out):

```
phase  = (marker_size mod 2) / 2
centre = floor(target) - phase
```

`target` is the sample line's own continuous position on that axis (its y for
the marker's y, its x-extent midpoint for the marker's x), and `centre` is
the marker's own centre, in unscaled chart-design points (page pt / the
sheet's print-fit scale — the JSON files below and #1617's are both already
converted). `verify_floor_phase_rule.py` checks this against every row of
both datasets — 9 legend sizes (6-18pt), two independent export batches, both
axes — and every row matches within 0.02pt (the measurement tooling's own
rounding). Run it with no arguments.

This reproduces the issue's own reported numbers to the digit: at 9pt,
`target = 987.8459` chart pt, `floor(987.8459) - 0.5 = 986.5`, `986.5 x 0.82 =
808.93` pt — the native circle centre `x` the issue quotes.

**Floor, not round, is pinned, not assumed.** Every row in the fresh batch
has `frac(target) < 0.5`, where floor and round agree — they do not
discriminate the rule on their own. The #1617 batch supplies six rows with
`frac(target) >= 0.5`, where the two disagree, and native matches floor at
every one: Verdana 9pt (`target=436.571`, floor gives centre 435.5, round
would give 436.5; observed 435.5), Verdana 10pt (`435.969` -> 435.5, observed
435.5), Verdana 12pt (`434.766` -> 433.5, observed 433.5), and the marker24
batch's 8pt (`436.716` -> 436.0), 11pt (`434.744` -> 434.0) and 14pt
(`432.769` -> 432.0) rows. `verify_floor_phase_rule.py` checks all of these.

## What blocks implementing it generally

**X is algebraically reachable but boundary-sensitive on this machine.** The
legend entry's left edge (`entry_x`, from `LegendBox::entry_origin`) is a
pure Rust computation — cumulative label/key widths, no font-ascent
ambiguity. It is chart-frame-local, not worksheet-absolute, so a correct
implementation floors `sheet_frame_origin_pt.0 + entry_x +
LEGEND_KEY_LEN_PT / 2` — the same absolute space `WorksheetMarkerPlacement`
already floors/rounds the plot markers in — not `entry_x` alone
(`floor(a + b) != floor(a) + b` for fractional `b`). `sheet_frame_origin_pt`
is in scope at the two call sites read while investigating this (around
`generate_chart_axis`'s legend loop and `generate_chart_line_plot`'s); the
other two `line_legend_key` call sites (radar and the standalone line-plot
legend, roughly lines 5556 and 5841 of `typst_gen_diagrams.rs` as of this
writing) were not individually re-checked.

`native-legend-line-key-measurements.json` holds four **fresh** native
exports (6/9/12/18pt) gathered for this issue, independent of #1617's batch,
specifically to get the line's own x-extent (#1617 recorded the marker but
not the sample line's x0/x1). office2pdf's own (Selawik-substituted, see
below) continuous x target tracks these within 0.04pt at all four sizes, and
`floor(ours) - phase` reproduces native's marker left at all four — but the
floor margin (distance from the continuous target to the next integer) is as
small as 0.053pt (12pt row). This fixture renders its Segoe UI legend as
Selawik (Microsoft's metric-compatible substitute, since Segoe UI itself
is not available on this machine) — see the `office2pdf-line-y-drift.json`
note below for the same substitution on the `y` axis. A font-metric
difference well under a point is enough to flip which integer the floor
lands on, so implementing `x` alone still needs validating against a real
Segoe UI export (or another face this repository can measure directly) before
it can be trusted unattended, not just algebraic soundness.

**Y has two separate blockers**, and confusing them is the trap for whoever
picks this up next:

1. *office2pdf's own sample-line `y` doesn't match native away from 9pt* —
   a pre-existing defect, not introduced by this investigation.
   `office2pdf-line-y-drift.json` compares office2pdf's own (currently
   unfixed) sample-line `y` against native across the same nine sizes: the
   two agree to 0.06pt at 9pt — the exact size #1618 was filed against — but
   diverge to **-2.87pt at 18pt**, growing roughly monotonically past 9pt.
   The cause: `line_legend_key`'s `box_height` is
   `SERIES_MARKER_SIZE_PT.min(marker_cap)`, which saturates at 5.0pt once the
   #1617 marker cap (`floor(0.6 x legend_pt)`) reaches 5 — every legend size
   from 9pt up computes the *same* 5.0pt box height, so the key-internal
   offset (`box_height / 2`) the sample line is seated at freezes while
   native's row height keeps scaling continuously (the drift table above
   still shows office2pdf's absolute `y` moving, ~3pt from 9pt to 18pt — just
   not by enough to track native's own, larger movement). The quantity
   native's row height actually tracks is already implemented and already
   validated to within 0.002pt at every size in #1617's dataset (the only one
   of the two with a `key_height_chart_pt` column;
   `native-legend-line-key-measurements.json`, gathered fresh for #1618,
   does not carry that field):
   `legend_key_line_box_pt` (`0.45 x (ascent_em + descent_em) x
   legend_pt`, the same share the bar-family key's flat rectangle uses) —
   `line_legend_key` simply is not the one using it. Filed separately as
   [#1753](https://github.com/developer0hye/office2pdf/issues/1753), since it
   is an independent root cause (the sample line's own seat, not the
   marker's offset from it) and this project's convention is one issue per
   root cause.

2. *Rust does not have the sample line's resolved absolute `y` to floor,
   even once #1753 makes it accurate.* The marker and the line share the
   same local `key_mid` value inside `line_legend_key`'s `#box(...,
   baseline: LEGEND_KEY_BASELINE_PT)`, but where that box ends up on the
   page depends on Typst's own inline layout — the ambient label text's
   ascent and the box's baseline shift interact (see
   `LEGEND_KEY_BASELINE_PT`'s own doc comment: "raising the box also grows
   the line's ascent and carries the baseline with it, so the offset is not
   a plain translation"). Fixing #1753 corrects what `y` Typst *renders*;
   it does not hand Rust that resolved number to floor at generation time.
   That needs either replicating Typst's ascent/baseline computation in
   Rust (a font-metrics derivation, not yet attempted) or placing the key at
   an explicit, Rust-computed absolute position instead of relying on
   inline box flow — both bigger than this issue's own scope.

Because of (2), the marker's `y` floor cannot be implemented generally today
regardless of #1753. Implementing only the `x` half would still fail the
issue's own joint 0.5pt gate on `y` at every size but the one this was
originally measured at, and — being a raster change — would need the full
visual evidence contract (`gt.jpg`/`before.jpg`/`after.jpg`) for a defect
that would not actually close. Per the project's TDD rule ("the test must
pin the general rule, not the one input from the issue"), landing a fix that
only holds at 9pt is exactly the kind of overfit this repository's own
history warns against — see `LEGEND_KEY_BASELINE_PT`'s own doc comment,
calibrated rather than derived for the same reason.

## Reproduction

The fresh export batch used the unchanged public Gift Budget and Tracker1.xlsx
attachment from #982 (SHA-256
`25f5dc75dab19ea12042979a61842314ddc226e3e45d447e36b2a2a104112613`), patching
only the bottom legend's `a:defRPr/@sz` in `xl/charts/chart1.xml` (the
series' own declared `c:marker/c:size` of 5 was left unchanged):

```sh
python3 scripts/probe_harness.py assets/validation/issue-1618/legend-size-x-probe.json --backend office
```

then, on each export's page 2 (`target/probes/issue-1618-legend-size-x/pdfs/`):

```sh
python3 assets/validation/issue-1617/measure_legend_row.py 0.82 target/probes/issue-1618-legend-size-x/pdfs/*.pdf   # key height, sample-line y, marker box
python3 assets/validation/issue-1618/measure_legend_line_x.py 0.82 target/probes/issue-1618-legend-size-x/pdfs/*.pdf  # sample-line x0/x1/mid
```

Both read the same `mutool draw -F trace -o - <pdf> 2` output and divide
every coordinate by the fixture's 0.82 print-fit scale to get chart-design
points. Both this batch's and #1617's row data are pre-extracted into the
JSON files here, so re-running the verification script does not require
Microsoft Excel — only regenerating the exports does.

`office2pdf-line-y-drift.json` was gathered by converting the same
legend-size variants with office2pdf's own CLI, built with
`cargo build --locked -p office2pdf-cli --release` (**not** `-p office2pdf`,
which only builds the library crate and silently leaves a stale binary in the
shared `CARGO_TARGET_DIR`), and running the same trace extraction on its
output.
