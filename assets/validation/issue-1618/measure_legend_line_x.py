#!/usr/bin/env python3
"""Line-legend sample stroke x-extent from a native Excel export, in chart points (#1618).

Companion to assets/validation/issue-1617/measure_legend_row.py, which reads
the same trace for the key height, sample-line y, line width and marker box
but does not report the sample stroke's own x0/x1 — needed here to get a
native, per-size continuous x target for the marker floor rule, independent
of any office2pdf-side font substitution.
"""
import subprocess
import sys
import xml.etree.ElementTree as ET


def mat(s: str) -> list[float]:
    v = [float(x) for x in s.split()]
    return v if len(v) == 6 else v + [0.0, 0.0]


def measure(pdf: str, page: int, scale: float) -> None:
    out = subprocess.run(
        ["mutool", "draw", "-F", "trace", "-o", "-", pdf, str(page)],
        capture_output=True,
        text=True,
    ).stdout
    root = ET.fromstring(out)
    elements = []
    keys = []
    for el in root.iter():
        if el.tag not in ("fill_path", "stroke_path"):
            continue
        m = mat(el.get("transform")) if el.get("transform") else [1, 0, 0, 1, 0, 0]
        xs: list[float] = []
        ys: list[float] = []
        for p in el.iter():
            if p.tag in ("moveto", "lineto", "curveto") and p.get("x") is not None:
                xs.append(m[0] * float(p.get("x")) + m[2] * float(p.get("y")) + m[4])
                ys.append(m[1] * float(p.get("x")) + m[3] * float(p.get("y")) + m[5])
            if p.tag == "curveto":
                for a, b in (("x1", "y1"), ("x2", "y2"), ("x3", "y3")):
                    if p.get(a):
                        xs.append(m[0] * float(p.get(a)) + m[2] * float(p.get(b)) + m[4])
                        ys.append(m[1] * float(p.get(a)) + m[3] * float(p.get(b)) + m[5])
        if not xs:
            continue
        bbox = (min(xs) / scale, min(ys) / scale, max(xs) / scale, max(ys) / scale)
        elements.append((el.tag, bbox, el.get("linewidth")))
        # Bar-family keys are flat LEGEND_KEY_LEN_PT (19.2pt)-wide rectangles;
        # collecting them locates the legend row's own y-band.
        if el.tag == "fill_path" and abs(bbox[2] - bbox[0] - 19.2) < 0.05 and bbox[3] - bbox[1] < 40:
            keys.append(bbox)
    if not keys:
        print(pdf, "no keys found")
        return
    y0 = min(k[1] for k in keys) - 25
    y1 = max(k[3] for k in keys) + 25
    # The line key's own sample stroke: same 19.2pt width, within the row's
    # y-band, to the right of every bar-family key (issue #801's line key
    # always trails the bar keys in this fixture's series order).
    line = [
        bbox
        for tag, bbox, _ in elements
        if tag == "stroke_path" and abs(bbox[2] - bbox[0] - 19.2) < 0.05 and y0 <= bbox[1] <= y1 and bbox[0] > keys[-1][2]
    ]
    for bbox in line:
        x0, y, x1, _ = bbox
        print(f"{pdf}: line x0={x0:.4f} x1={x1:.4f} mid={(x0 + x1) / 2:.4f} y={y:.4f}")


if __name__ == "__main__":
    scale = float(sys.argv[1])
    for path in sys.argv[2:]:
        measure(path, 2, scale)
