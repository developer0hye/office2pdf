"""Check the #1748 bottom-legend seat rule against the committed native data.

Reads the #1617 native batch (sample-line centres for nine Segoe UI and six
Verdana legend sizes) and the six native Segoe UI label baselines, all in
chart points (page points / the 0.82 fit). Run from the repository root with
no arguments; exits non-zero on any mismatch.
"""

import json
import math
import sys
from pathlib import Path

# hhea ascent, descent per em; Segoe UI as the crate holds it, Verdana 2059/430.
FACES = {"Segoe UI": (2210 / 2048, 514 / 2048), "Verdana": (2059 / 2048, 430 / 2048)}
FRAME_BOTTOM = 450.9859  # chart frame's bottom edge on page 2, chart points
KEY_PAD, KEY_SHARE, LABEL_RISE = 9.0, 0.495, 1.0
# 6pt and 9pt/18pt: the #1748 filing; 7, 11, 14pt: the #1240 probes.
BASELINES = {6: 439, 7: 440, 9: 439, 11: 438, 14: 437, 18: 436}


def key_centre(face, size):
    ascent, descent = FACES[face]
    return FRAME_BOTTOM - KEY_PAD - KEY_SHARE * (ascent + descent) * size


def label_baseline(size):
    ascent, descent = FACES["Segoe UI"]
    top = key_centre("Segoe UI", size) - (ascent + descent) * size / 2 - LABEL_RISE
    return math.floor(top + 0.5) + math.floor(ascent * size + 0.5)


rows = json.loads(
    Path("assets/validation/issue-1617/native-legend-marker-measurements.json").read_text()
)["rows"]
failures = 0
for row in rows:
    face, size = row["legend_face"], row["legend_size_pt"]
    delta = key_centre(face, size) - row["sample_line_y_chart_pt"]
    failures += abs(delta) > 0.01
    print(f"{face} {size}pt key centre delta {delta:+.4f}")
for size, native in BASELINES.items():
    got = label_baseline(size)
    failures += got != native
    print(f"Segoe UI {size}pt baseline {got} native {native}")
sys.exit(1 if failures else 0)
