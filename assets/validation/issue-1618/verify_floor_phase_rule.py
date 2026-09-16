#!/usr/bin/env python3
"""Verify the legend-marker floor rule against native Excel measurements (#1618).

Native Excel seats a line-legend key's marker on a whole chart point, on the
same axis the plot marker's own WorksheetMarkerPlacement already floors to
(#1105): given the continuous sample-line target and the marker's own size,

    phase = (size mod 2) / 2
    center = floor(target) - phase

This reproduces every row of both native datasets below to the reported
precision (their own export tooling rounds to 3-4 decimal digits), on both
axes, across nine legend sizes (6-18pt) and two independent export batches
(the #1617 investigation's and a fresh one gathered for #1618). Run with no
arguments; it exits non-zero if any row's prediction misses by more than
0.02pt (the reported measurement resolution).
"""
from __future__ import annotations

import json
import math
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOLERANCE_PT = 0.02


def phase(size: float) -> float:
    return (size % 2.0) / 2.0


def predict_center(target: float, size: float) -> float:
    return math.floor(target) - phase(size)


def check(label: str, target: float, size: float, observed_left: float) -> bool:
    observed_center = observed_left + size / 2.0
    predicted_center = predict_center(target, size)
    ok = abs(predicted_center - observed_center) <= TOLERANCE_PT
    status = "OK  " if ok else "FAIL"
    print(
        f"[{status}] {label}: target={target:.4f} size={size:.1f} "
        f"predicted_center={predicted_center:.4f} observed_center={observed_center:.4f}"
    )
    return ok


def main() -> int:
    all_ok = True

    fresh = json.loads((HERE / "native-legend-line-key-measurements.json").read_text())
    for row in fresh["rows"]:
        all_ok &= check(
            f"issue-1618 fresh export / y / {row['legend_size_pt']}pt",
            row["line_y_chart_pt"],
            row["marker_h_chart_pt"],
            row["marker_top_chart_pt"],
        )
        all_ok &= check(
            f"issue-1618 fresh export / x / {row['legend_size_pt']}pt",
            row["line_mid_x_chart_pt"],
            row["marker_w_chart_pt"],
            row["marker_left_chart_pt"],
        )

    batch_1617 = json.loads(
        (HERE.parent / "issue-1617" / "native-legend-marker-measurements.json").read_text()
    )
    for row in batch_1617["rows"]:
        all_ok &= check(
            f"issue-1617 batch / y / {row['batch']}/{row['export']}",
            row["sample_line_y_chart_pt"],
            row["marker_h_chart_pt"],
            row["marker_top_chart_pt"],
        )

    if all_ok:
        print("\nAll rows match the floor(target) - phase rule within "
              f"{TOLERANCE_PT}pt.")
        return 0
    print("\nSome rows did not match — see FAIL lines above.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
