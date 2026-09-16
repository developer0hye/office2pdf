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

Exit code is 0 when claims 1-3 hold exactly as documented (i.e. the
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


def main() -> int:
    data = json.loads(DATA_PATH.read_text())
    scale: float = data["print_scale"]

    failures: list[str] = []

    # Claim 1: every native baseline is a whole sheet point.
    for chart_name, chart in data["charts"].items():
        for size, by_size in chart["sizes"].items():
            for cat, native_y in by_size["native_y"].items():
                ns = sheet_y(native_y, scale)
                if abs(ns - round(ns)) > 1e-6:
                    failures.append(
                        f"claim 1 violated: {chart_name}@{size}pt {cat!r} "
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

    if failures:
        print("FAIL: measurements.json no longer matches README.md's documented findings:")
        for failure in failures:
            print(f"  - {failure}")
        return 1

    print("PASS: all three documented claims verified against measurements.json")
    print(f"  - every native baseline is a whole sheet point (print_scale={scale})")
    print(f"  - floor(ours_sheet) reproduces the expense chart only: {floor_results}")
    print(f"  - round(ours_sheet) reproduces the income chart only: {round_results}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
