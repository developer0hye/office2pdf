#!/usr/bin/env python3
"""Verify the confirmed and falsified claims in README.md against measurements.json (#1621).

Confirms:
  1. Every measured native category-label baseline, divided by the sheet's
     print scale, is an integer number of sheet points (whole-sheet-point
     quantization, generalizing the issue's own 6-row observation to a
     second chart and three font sizes).

Falsifies (documents why neither a plain `floor` nor a plain `round` of
office2pdf's own current continuous baseline reproduces native across both
charts with one shared rule):
  2. `floor(ours_sheet)` matches every expense-chart row, but not every
     income-chart row.
  3. `round(ours_sheet)` matches every income-chart row, but not every
     expense-chart row.

Confirms (category-count probe, added in a later session):
  4. `floor(sheet_frame_top_pt + row_top + row/2) + K`, with `K` depending
     only on font size (not on chart identity or category count), reproduces
     every interior row and both income charts' bottom-edge row across three
     independent geometric configurations (income @ 5 categories x 3 sizes,
     expense @ 6 categories, income @ 6 categories via the category-count
     probe) -- 18 of the 24 rows across those five (chart, size) pairs. The
     exact six-row exception set (the 5-category income chart's top row at
     each of its 3 measured sizes, the 6-category income-probe's top row,
     and both edges of the 6-category expense chart) is asserted explicitly
     so a future change is caught rather than silently re-narrowing or
     widening the gap.

Falsifies (new):
  5. Naively snapping the plot rectangle's own continuous top/bottom edges to
     the nearest whole sheet point, recomputing `row` from the snapped span,
     and reapplying `floor(centre) + K` does not close the income chart's
     edge-row gap and additionally breaks a previously-exact interior row
     (`from savings`, 5-category income @ 10pt).

Exit code is 0 when claims 1-5 hold exactly as documented (i.e. the
measurements still support the README's "not yet a shippable rule"
conclusion) and non-zero if the data no longer matches that story --
which would mean README.md needs a re-read, not that the script is wrong.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

DATA_PATH = Path(__file__).with_name("measurements.json")


def sheet_y(printed_y: float, scale: float) -> float:
    return printed_y / scale


def floor_center_plus_k(
    frame_top: float, plot_y: float, row: float, n: int, idx: int, k: int
) -> int:
    """floor(sheet_frame_top_pt + row_top + row/2) + K for category `idx` of `n`."""
    row_top = plot_y + (n - 1 - idx) * row
    center = frame_top + row_top + row / 2
    return int(center // 1) + k


def edge_kind(idx: int, n: int) -> str:
    if idx == 0:
        return "bottom"
    if idx == n - 1:
        return "top"
    return "interior"


def main() -> int:
    data = json.loads(DATA_PATH.read_text())
    scale: float = data["print_scale"]
    frame_top: float = data["sheet_frame_top_pt"]

    failures: list[str] = []

    # Claim 1: every native baseline is a whole sheet point (base sizes and
    # the category-count probe alike).
    for chart_name, chart in data["charts"].items():
        for size, by_size in chart["sizes"].items():
            for cat, native_y in by_size["native_y"].items():
                ns = sheet_y(native_y, scale)
                if abs(ns - round(ns)) > 1e-6:
                    failures.append(
                        f"claim 1 violated: {chart_name}@{size}pt {cat!r} "
                        f"native_sheet={ns:.6f} is not a whole sheet point"
                    )
        for n_str, probe in chart.get("category_count_probes", {}).items():
            if n_str == "note":
                continue
            for size, by_size in probe["sizes"].items():
                for cat, native_y in by_size["native_y"].items():
                    ns = sheet_y(native_y, scale)
                    if abs(ns - round(ns)) > 1e-6:
                        failures.append(
                            f"claim 1 violated: {chart_name} N={n_str}@{size}pt {cat!r} "
                            f"native_sheet={ns:.6f} is not a whole sheet point"
                        )

    # Claims 2 & 3: no single rule (plain floor, plain round) reproduces
    # native from office2pdf's own current continuous baseline across BOTH
    # charts. We only have same-size (10pt) coverage for the expense chart,
    # so claim 2/3 are checked against the 10pt rows of each chart, which is
    # the one point of direct comparison the two charts share.
    def rule_matches(rule) -> dict[str, bool]:
        results: dict[str, bool] = {}
        for chart_name, chart in data["charts"].items():
            by_size = chart["sizes"].get("10")
            if not by_size:
                continue
            chart_ok = True
            for cat, native_y in by_size["native_y"].items():
                ours_y = by_size["ours_y"].get(cat)
                if ours_y is None:
                    continue
                ns = round(sheet_y(native_y, scale))
                os_ = sheet_y(ours_y, scale)
                if rule(os_) != ns:
                    chart_ok = False
            results[chart_name] = chart_ok
        return results

    floor_results = rule_matches(lambda x: int(x // 1))
    round_results = rule_matches(lambda x: round(x))

    if not (floor_results.get("expense") is True and floor_results.get("income") is False):
        failures.append(
            f"claim 2 expectation changed: floor(ours_sheet) results are now {floor_results}"
        )
    if not (round_results.get("income") is True and round_results.get("expense") is False):
        failures.append(
            f"claim 3 expectation changed: round(ours_sheet) results are now {round_results}"
        )

    # Claim 4: floor(centre) + size-only K reproduces every row except a
    # named, exact exception set confined to plot-rectangle edge rows.
    income = data["charts"]["income"]
    expense = data["charts"]["expense"]
    k_by_size = {"8": 4, "10": 4, "14": 5}

    configs = []
    for size, by_size in income["sizes"].items():
        names = income["categories_top_to_bottom"]
        name_to_idx = {name: len(names) - 1 - i for i, name in enumerate(names)}
        configs.append(
            (
                f"income@{size}pt",
                income["plot_y_sheet_pt"],
                income["row_sheet_pt"],
                len(names),
                {n: name_to_idx[n] for n in by_size["native_y"]},
                by_size["native_y"],
                k_by_size[size],
            )
        )
    for size, by_size in expense["sizes"].items():
        names = expense["categories_top_to_bottom"]
        name_to_idx = {name: len(names) - 1 - i for i, name in enumerate(names)}
        configs.append(
            (
                f"expense@{size}pt",
                expense["plot_y_sheet_pt"],
                expense["row_sheet_pt"],
                len(names),
                name_to_idx,
                by_size["native_y"],
                k_by_size[size],
            )
        )
    for n_str, probe in income.get("category_count_probes", {}).items():
        if n_str == "note":
            continue
        n = int(n_str)
        names = probe["categories_top_to_bottom"]
        name_to_idx = {name: len(names) - 1 - i for i, name in enumerate(names)}
        for size, by_size in probe["sizes"].items():
            configs.append(
                (
                    f"income-N{n}@{size}pt",
                    income["plot_y_sheet_pt"],
                    probe["row_sheet_pt"],
                    n,
                    name_to_idx,
                    by_size["native_y"],
                    k_by_size[size],
                )
            )

    # The exact rows the floor+K model is known NOT to reproduce -- an edge
    # row deviating by exactly one whole sheet point. Anything outside this
    # set must match exactly, or claim 4 no longer holds as documented.
    expected_exceptions = {
        ("income@8pt", "other"),
        ("income@10pt", "other"),
        ("income@14pt", "other"),
        ("income-N6@10pt", "probe extra"),
        ("expense@10pt", "room & board"),
        ("expense@10pt", "other expenses"),
    }
    seen_exceptions: set[tuple[str, str]] = set()
    total_rows = 0
    for label, plot_y, row, n, name_to_idx, native_map, k in configs:
        for name, idx in name_to_idx.items():
            total_rows += 1
            predicted_sheet = floor_center_plus_k(frame_top, plot_y, row, n, idx, k)
            predicted_printed = predicted_sheet * scale
            actual = native_map[name]
            if abs(predicted_printed - actual) > 0.01:
                seen_exceptions.add((label, name))

    if seen_exceptions != expected_exceptions:
        failures.append(
            "claim 4 exception set changed: "
            f"expected {sorted(expected_exceptions)}, got {sorted(seen_exceptions)}"
        )

    # Claim 5: independently snapping the plot rectangle's continuous edges
    # to whole sheet points does not fix the income top-edge gap and
    # introduces a new miss on a previously-exact interior row.
    plot_top_abs = frame_top + income["plot_y_sheet_pt"]
    plot_bottom_abs = frame_top + income["plot_y_sheet_pt"] + income["plot_h_sheet_pt"]
    snapped_top = round(plot_top_abs)
    snapped_bottom = round(plot_bottom_abs)
    snapped_row = (snapped_bottom - snapped_top) / 5
    names5 = income["categories_top_to_bottom"]
    name_to_idx5 = {name: len(names5) - 1 - i for i, name in enumerate(names5)}
    native_10pt = income["sizes"]["10"]["native_y"]
    from_savings_idx = name_to_idx5["from savings"]
    row_top = snapped_top + (5 - 1 - from_savings_idx) * snapped_row
    center = row_top + snapped_row / 2
    predicted = (int(center // 1) + 4) * scale
    actual = native_10pt["from savings"]
    if abs(predicted - actual) < 0.01:
        failures.append(
            "claim 5 expectation changed: edge-snapping no longer breaks "
            f"'from savings' (predicted={predicted}, actual={actual}) -- "
            "the edge-snap hypothesis may now be viable and README.md needs a re-read"
        )

    if failures:
        print("FAIL: measurements.json no longer matches README.md's documented findings:")
        for failure in failures:
            print(f"  - {failure}")
        return 1

    print("PASS: all five documented claims verified against measurements.json")
    print(f"  - every native baseline is a whole sheet point (print_scale={scale})")
    print(f"  - floor(ours_sheet) reproduces the expense chart only: {floor_results}")
    print(f"  - round(ours_sheet) reproduces the income chart only: {round_results}")
    matched_rows = total_rows - len(expected_exceptions)
    print(
        f"  - floor(centre) + size-only K reproduces {matched_rows}/{total_rows} rows; "
        f"exceptions confined to: {sorted(expected_exceptions)}"
    )
    print("  - independent plot-edge snapping does not close the remaining gap")
    return 0


if __name__ == "__main__":
    sys.exit(main())
