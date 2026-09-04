#!/usr/bin/env python3
"""Trace-based layout differ: a per-page deviation vector instead of a pixel scalar.

Parses ``mutool draw -F trace`` for a ground-truth PDF and an office2pdf
output, reconstructs text and vector-path bounds in device points, and
conservatively recovers ``ignore_text`` geometry when a nearby path supplies
visible ink. It then matches lines by their text and reports typed deviations:

- matched / missing / extra lines, safe split/join topology differences for
  distant text objects, and wrap-point differences (text that is present but
  breaks at a different word) counted separately from real loss;
- spatial-anchor dy statistics, per-line dx0, line-width drift, inter-line
  pitch deltas between consecutive matched lines. Horizontal text uses its
  true baseline; a rotated or skewed `fill_text` stays one visual run and uses
  the minimum fully transformed glyph x/y as its comparable anchor;
- painted visibility from the page media box, active rectangular clips, and
  trace order: text outside the visible bounds, covered by a later opaque
  image, single closed axis-aligned rectangular fill, or fully extended
  shading under a single closed axis-aligned rectangular clip, or painted with
  the same colour as a flat background, is distinguished from text that
  remains visibly painted. Unmatched hidden trace lines are reported but do
  not become visual missing/extra findings;
- visible-fill occlusions: a later opaque, differently coloured rectangle
  cutting into a thin earlier rule is compared by final overlap area, colour,
  and paint order rather than by the raw rectangle count;
- a fill/stroke path census plus geometry-aware x/y/size/edge deltas for
  axis-aligned rectangles and straight lines. Touching same-paint operation
  splits are merged before matching while raw counts stay visible.

A noise floor (default 0.12pt — native Word exports quantise coordinates to a
0.24pt grid; use 0.5 for Excel GT, whose Quartz export rounds every advance to
a whole point) separates measurement noise from deviations: raw statistics are
always reported, but a line only counts as *deviant* past the floor.

Pixel metrics (`compare_render.py`) remain the tripwire for whatever this tool
does not model — colour, images, effects. This tool is the primary signal for
position, size, pitch, wrap, and element presence, which is what actually
changes when a layout defect is fixed.

Usage:
    compare_layout.py GT.pdf OUTPUT.pdf [--page N] [--noise-floor PT]
        [--fine-shift PT] [--audit]
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import statistics
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from spatial_match import minimum_cost_pairs

# mutool 1.23.x opens a page as `<page mediabox="...">` with no `number`
# attribute; later builds add one. Requiring it parses zero pages and reports
# that as "no differences found" instead of as a failure. Keep the attributes
# so the media box can bound the visual text census.
PAGE_RE = re.compile(r"<page\b([^>]*)>(.*?)</page>", re.S)
TEXT_RE = re.compile(r"<(fill_text|ignore_text)\b([^>]*)>(.*?)</\1>", re.S)
SPAN_RE = re.compile(r"<span\b([^>]*)>(.*?)</span>", re.S)
GLYPH_RE = re.compile(
    r'<g unicode="([^"]*)" glyph="[^"]*" x="([-0-9.e]+)" y="([-0-9.e]+)" adv="([-0-9.e]+)"'
)
PATH_RE = re.compile(r"<(fill_path|stroke_path)\b([^>]*)>(.*?)</\1>", re.S)
FILL_IMAGE_RE = re.compile(r"<fill_image\b([^>]*)/>", re.S)
CLIP_PATH_RE = re.compile(r"<clip_path\b([^>]*)>(.*?)</clip_path>", re.S)
FILL_SHADE_RE = re.compile(r"<fill_shade\b([^>]*)/>", re.S)
POP_CLIP_RE = re.compile(r"<pop_clip\s*/>")
CLIP_SHADE_EVENT_RE = re.compile(
    r"<clip_path\b[^>]*>.*?</clip_path>|<fill_shade\b[^>]*/>|<pop_clip\s*/>", re.S
)
CLIP_TEXT_EVENT_RE = re.compile(
    r"<clip_path\b[^>]*>.*?</clip_path>|"
    r"<(?P<text_op>fill_text|ignore_text)\b[^>]*>.*?</(?P=text_op)>|"
    r"<pop_clip\s*/>",
    re.S,
)
POINT_RE = re.compile(r'<(?:moveto|lineto|curveto)[^>]*x="([-0-9.e]+)" y="([-0-9.e]+)"')
PATH_COMMAND_RE = re.compile(
    r"<(?:(?P<point>moveto|lineto)\b(?P<attrs>[^>]*)|(?P<close>closepath)\s*)/>"
)
X_RE = re.compile(r'\bx="([-0-9.e]+)"')
Y_RE = re.compile(r'\by="([-0-9.e]+)"')
TRANSFORM_RE = re.compile(
    r'transform="([-0-9.e]+) ([-0-9.e]+) ([-0-9.e]+) ([-0-9.e]+) ([-0-9.e]+) ([-0-9.e]+)"'
)
TRM_RE = re.compile(r'trm="([-0-9.e]+) ')
ALPHA_RE = re.compile(r'alpha="([-0-9.e]+)"')
COLOR_RE = re.compile(r'color="([-0-9.e ]+)"')
MEDIABOX_RE = re.compile(r'\bmediabox="([-0-9.e ]+)"')

LINE_Y_TOLERANCE_PT = 0.6
# A normal inter-word gap is far smaller than this. Combined with a text-paint
# operation boundary, this isolates independently positioned chart/table text
# that mutool happens to merge only because the objects share a baseline.
SAFE_TOPOLOGY_GAP_PT = 24.0
# Unequal rectangle groups contain unmatched operations by definition. Limit
# their pair candidates so a nearby rule cannot be assigned to an unrelated
# shape elsewhere on the page. Equal-cardinality groups remain fully paired so
# a large translation is reported as geometry instead of disappearing.
RECT_AMBIGUOUS_MATCH_RADIUS_PT = 24.0
# When one side contains extra paint operations, a corresponding area may be
# split slightly differently, but a several-times thinner/wider primitive is
# not defensibly the same rectangle. Keep that operation unmatched instead of
# manufacturing a large geometry delta.
RECT_AMBIGUOUS_MAX_EXTENT_RATIO = 1.5
RECT_RULE_MAX_THICKNESS_PT = 2.0
RECT_RULE_MIN_LENGTH_PT = 4.0
RECT_SAMPLE_LIMIT = 5
OPAQUE_ALPHA = 0.98
INVISIBLE_ALPHA = 0.02
LOW_CONTRAST_CHANNEL_DELTA = 0.04
VISIBLE_FILL_MIN_EDGE_PT = 0.5
VISIBLE_FILL_MIN_AREA_PT2 = 0.5
VISIBLE_FILL_MAX_RULE_THICKNESS_PT = 2.0
VISIBLE_FILL_MIN_RULE_LENGTH_PT = 4.0
VISIBLE_FILL_COLOR_DELTA = 0.08
VISIBLE_FILL_MATCH_TOLERANCE_PT = 0.5
# MuPDF can quantise a page edge and a page-sized paint a few hundredths of a
# point apart. This matches the established Word trace noise floor while still
# requiring geometric coverage rather than overlap.
PAINT_COVER_TOLERANCE_PT = 0.12
XML_ESCAPES = {"&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'"}


@dataclass(frozen=True)
class Glyph:
    x: float
    y: float
    unicode: str
    size: float
    advance: float
    paint_index: int = -1
    color: tuple[float, float, float] | None = None
    alpha: float = 1.0
    needs_path_ink: bool = False
    paint_window_start: int = -1
    paint_window_end: int = -1
    visible_clip: tuple[float, float, float, float] | None = None


@dataclass(frozen=True)
class Paint:
    """One later operation that can determine whether text remains visible."""

    index: int
    kind: str  # "flat" | "image" | "shade" | "path"
    x0: float
    y0: float
    x1: float
    y1: float
    alpha: float
    color: tuple[float, float, float] | None = None


@dataclass(frozen=True)
class Rect:
    kind: str  # "fill" | "stroke"
    geometry_kind: str  # "rectangle" | "line" | "other"
    x0: float
    y0: float
    x1: float
    y1: float
    color: tuple[float, float, float] | None = None
    alpha: float = 1.0

    @property
    def center(self) -> tuple[float, float]:
        return ((self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2)

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    @property
    def bbox(self) -> list[float]:
        return [self.x0, self.y0, self.x1, self.y1]


@dataclass(frozen=True)
class CanonicalRect:
    rect: Rect
    source_indices: tuple[int, ...]


@dataclass(frozen=True)
class VisibleFillOcclusion:
    """A later flat fill that changes a thin earlier rule's visible colour."""

    bbox: tuple[float, float, float, float]
    under_color: tuple[float, float, float]
    over_color: tuple[float, float, float]

    @property
    def area(self) -> float:
        return (self.bbox[2] - self.bbox[0]) * (self.bbox[3] - self.bbox[1])


@dataclass
class Line:
    y: float
    glyphs: list[Glyph] = field(default_factory=list)
    visibility: str = "painted"

    @property
    def text(self) -> str:
        return "".join(g.unicode for g in self.glyphs)

    @property
    def key(self) -> str:
        """Whitespace-free text; GT exports carry stray trailing space glyphs."""
        return "".join(g.unicode for g in self.glyphs if not g.unicode.isspace())

    @property
    def visible_glyphs(self) -> list[Glyph]:
        """Glyphs whose ink defines the line's comparable visual extent.

        Native exporters and Typst disagree about terminal space glyphs. Their
        advances position no later ink, so including them invents a width or
        left-edge delta while every painted glyph agrees (issue #1482).
        Internal spaces still contribute through the following glyph's origin.
        """
        return [glyph for glyph in self.glyphs if not glyph.unicode.isspace()]

    @property
    def x0(self) -> float:
        return min(glyph.x for glyph in self.visible_glyphs)

    @property
    def x1(self) -> float:
        return max(glyph.x + glyph.advance for glyph in self.visible_glyphs)

    @property
    def width(self) -> float:
        return self.x1 - self.x0


@dataclass
class PageLayout:
    lines: list[Line]
    rects: list[Rect]
    paints: list[Paint] = field(default_factory=list)
    media_box: tuple[float, float, float, float] | None = None


def unescape(text: str) -> str:
    for entity, char in XML_ESCAPES.items():
        text = text.replace(entity, char)
    return text


def parse_transform(attrs: str) -> tuple[float, float, float, float, float, float] | None:
    match = TRANSFORM_RE.search(attrs)
    if not match:
        return None
    return tuple(float(value) for value in match.groups())  # type: ignore[return-value]


def parse_alpha(attrs: str) -> float:
    match = ALPHA_RE.search(attrs)
    return float(match.group(1)) if match else 1.0


def parse_rgb(attrs: str) -> tuple[float, float, float] | None:
    """Return a comparable RGB colour for trace operations when available."""
    match = COLOR_RE.search(attrs)
    if not match:
        return None
    values = [float(value) for value in match.group(1).split()]
    if len(values) == 1:
        return (values[0], values[0], values[0])
    if len(values) == 3:
        return (values[0], values[1], values[2])
    if len(values) == 4:
        cyan, magenta, yellow, black = values
        return (
            1.0 - min(1.0, cyan + black),
            1.0 - min(1.0, magenta + black),
            1.0 - min(1.0, yellow + black),
        )
    return None


def transformed_bbox(
    transform: tuple[float, float, float, float, float, float],
    points: list[tuple[float, float]],
) -> tuple[float, float, float, float]:
    a, b, c, d, e, f = transform
    transformed = [(a * x + c * y + e, b * x + d * y + f) for x, y in points]
    xs = [point[0] for point in transformed]
    ys = [point[1] for point in transformed]
    return min(xs), min(ys), max(xs), max(ys)


def axis_aligned_rect_bbox(
    attrs: str, body: str
) -> tuple[float, float, float, float] | None:
    """Return a bbox only for one closed, ordered axis-aligned rectangle.

    Merely visiting all four bounding-box corners is not enough: a bow-tie or
    a second subpath does not prove that the bounding box is fully painted.
    """
    commands: list[tuple[str, tuple[float, float] | None]] = []
    cursor = 0
    for match in PATH_COMMAND_RE.finditer(body):
        if body[cursor : match.start()].strip():
            return None
        cursor = match.end()
        if match.group("close"):
            commands.append(("closepath", None))
            continue
        point_attrs = match.group("attrs")
        x_match = X_RE.search(point_attrs)
        y_match = Y_RE.search(point_attrs)
        if not x_match or not y_match:
            return None
        commands.append(
            (
                match.group("point"),
                (float(x_match.group(1)), float(y_match.group(1))),
            )
        )
    if body[cursor:].strip():
        return None
    if not commands or commands[0][0] != "moveto" or commands[-1][0] != "closepath":
        return None
    if any(command != "lineto" for command, _ in commands[1:-1]):
        return None

    points = [point for _, point in commands[:-1] if point is not None]
    if len(points) == 5 and points[-1] == points[0]:
        points.pop()
    if len(points) != 4:
        return None

    transform = parse_transform(attrs) or (1, 0, 0, 1, 0, 0)
    a, b, c, d, e, f = transform
    transformed = [(a * x + c * y + e, b * x + d * y + f) for x, y in points]
    rounded = [(round(x, 6), round(y, 6)) for x, y in transformed]
    xs = {x for x, _ in rounded}
    ys = {y for _, y in rounded}
    corners = {(x, y) for x in xs for y in ys}
    if len(xs) != 2 or len(ys) != 2 or set(rounded) != corners:
        return None
    for start, end in zip(rounded, rounded[1:] + rounded[:1]):
        same_x = start[0] == end[0]
        same_y = start[1] == end[1]
        if same_x == same_y:
            return None
    return transformed_bbox((1, 0, 0, 1, 0, 0), transformed)


def axis_aligned_line_bbox(
    attrs: str, body: str
) -> tuple[float, float, float, float] | None:
    """Return the bbox for one straight horizontal or vertical path."""
    points: list[tuple[float, float]] = []
    cursor = 0
    saw_move = False
    for match in PATH_COMMAND_RE.finditer(body):
        if body[cursor : match.start()].strip():
            return None
        cursor = match.end()
        if match.group("close"):
            continue
        command = match.group("point")
        if command == "moveto":
            if saw_move:
                return None
            saw_move = True
        elif not saw_move:
            return None
        point_attrs = match.group("attrs")
        x_match = X_RE.search(point_attrs)
        y_match = Y_RE.search(point_attrs)
        if not x_match or not y_match:
            return None
        points.append((float(x_match.group(1)), float(y_match.group(1))))
    if body[cursor:].strip() or not saw_move:
        return None

    transform = parse_transform(attrs) or (1, 0, 0, 1, 0, 0)
    a, b, c, d, e, f = transform
    transformed = {
        (round(a * x + c * y + e, 6), round(b * x + d * y + f, 6))
        for x, y in points
    }
    if len(transformed) < 2:
        return None
    xs = {point[0] for point in transformed}
    ys = {point[1] for point in transformed}
    if (len(xs) == 1) == (len(ys) == 1):
        return None
    return transformed_bbox((1, 0, 0, 1, 0, 0), list(transformed))


def intersect_bboxes(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> tuple[float, float, float, float]:
    x0 = max(first[0], second[0])
    y0 = max(first[1], second[1])
    x1 = min(first[2], second[2])
    y1 = min(first[3], second[3])
    if x0 >= x1 or y0 >= y1:
        return (x0, y0, x0, y0)
    return (x0, y0, x1, y1)


def parse_media_box(attrs: str) -> tuple[float, float, float, float] | None:
    """Return the page's device-space media box when mutool provides one."""
    match = MEDIABOX_RE.search(attrs)
    if match is None:
        return None
    values = [float(value) for value in match.group(1).split()]
    if len(values) != 4:
        return None
    x0, y0, x1, y1 = values
    return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))


def text_visible_clips(
    content: str,
    media_box: tuple[float, float, float, float] | None,
) -> dict[int, tuple[float, float, float, float] | None]:
    """Map text operation offsets to conservative active visual bounds.

    A rectangular clip can prove that text outside it cannot paint. Unknown
    clip shapes only narrow the real visible region, so retaining the last
    known outer bound is conservative and cannot hide a visual finding.
    """
    active_clip = media_box
    stack: list[tuple[float, float, float, float] | None] = []
    clips: dict[int, tuple[float, float, float, float] | None] = {}

    for event in CLIP_TEXT_EVENT_RE.finditer(content):
        operation = event.group(0)
        clip_match = CLIP_PATH_RE.fullmatch(operation)
        if clip_match:
            stack.append(active_clip)
            bbox = axis_aligned_rect_bbox(*clip_match.groups())
            if bbox is not None:
                active_clip = (
                    bbox
                    if active_clip is None
                    else intersect_bboxes(active_clip, bbox)
                )
            continue

        if TEXT_RE.fullmatch(operation):
            clips[event.start()] = active_clip
            continue

        if POP_CLIP_RE.fullmatch(operation):
            active_clip = stack.pop() if stack else media_box

    return clips


def color_delta(
    first: tuple[float, float, float], second: tuple[float, float, float]
) -> float:
    return max(abs(channel - other) for channel, other in zip(first, second))


def merge_visible_fill_occlusions(
    occlusions: list[VisibleFillOcclusion],
) -> list[VisibleFillOcclusion]:
    """Union adjacent cells without preserving harmless operation splits."""

    coordinate_tolerance = 1e-6
    merged = list(occlusions)
    changed = True
    while changed:
        changed = False
        for first_index, first in enumerate(merged):
            for second_index in range(first_index + 1, len(merged)):
                second = merged[second_index]
                if (
                    color_delta(first.under_color, second.under_color)
                    > coordinate_tolerance
                    or color_delta(first.over_color, second.over_color)
                    > coordinate_tolerance
                ):
                    continue
                same_y_span = (
                    abs(first.bbox[1] - second.bbox[1]) <= coordinate_tolerance
                    and abs(first.bbox[3] - second.bbox[3]) <= coordinate_tolerance
                )
                same_x_span = (
                    abs(first.bbox[0] - second.bbox[0]) <= coordinate_tolerance
                    and abs(first.bbox[2] - second.bbox[2]) <= coordinate_tolerance
                )
                x_connected = (
                    first.bbox[0] <= second.bbox[2] + coordinate_tolerance
                    and second.bbox[0] <= first.bbox[2] + coordinate_tolerance
                )
                y_connected = (
                    first.bbox[1] <= second.bbox[3] + coordinate_tolerance
                    and second.bbox[1] <= first.bbox[3] + coordinate_tolerance
                )
                if not (
                    (same_y_span and x_connected) or (same_x_span and y_connected)
                ):
                    continue
                merged[first_index] = VisibleFillOcclusion(
                    bbox=(
                        min(first.bbox[0], second.bbox[0]),
                        min(first.bbox[1], second.bbox[1]),
                        max(first.bbox[2], second.bbox[2]),
                        max(first.bbox[3], second.bbox[3]),
                    ),
                    under_color=first.under_color,
                    over_color=first.over_color,
                )
                del merged[second_index]
                changed = True
                break
            if changed:
                break
    return merged


def visible_fill_occlusions(paints: list[Paint]) -> list[VisibleFillOcclusion]:
    """Return material later-fill intrusions into thin earlier rules.

    Mutool records vector geometry before raster antialiasing, so point and
    area floors can reject trace slivers without depending on DPI. Restricting
    the earlier paint to a long, thin rule also avoids treating ordinary
    stacked cell backgrounds as border damage. Same-colour operation splitting
    has no final visible-colour change and is ignored (issue #1476).
    """

    opaque_flat_fills = [
        paint
        for paint in paints
        if paint.kind == "flat"
        and paint.alpha >= OPAQUE_ALPHA
        and paint.color is not None
    ]
    occlusion_cells: list[VisibleFillOcclusion] = []
    for earlier in opaque_flat_fills:
        earlier_width = earlier.x1 - earlier.x0
        earlier_height = earlier.y1 - earlier.y0
        if (
            min(earlier_width, earlier_height) > VISIBLE_FILL_MAX_RULE_THICKNESS_PT
            or max(earlier_width, earlier_height) < VISIBLE_FILL_MIN_RULE_LENGTH_PT
        ):
            continue
        assert earlier.color is not None
        earlier_bbox = (earlier.x0, earlier.y0, earlier.x1, earlier.y1)
        later_fills = [
            paint
            for paint in opaque_flat_fills
            if paint.index > earlier.index
            and bboxes_overlap(
                earlier_bbox,
                (paint.x0, paint.y0, paint.x1, paint.y1),
            )
        ]
        if not later_fills:
            continue

        clipped_overlaps = [
            intersect_bboxes(
                earlier_bbox,
                (paint.x0, paint.y0, paint.x1, paint.y1),
            )
            for paint in later_fills
        ]
        x_coordinates = sorted(
            {earlier.x0, earlier.x1}
            | {coordinate for bbox in clipped_overlaps for coordinate in (bbox[0], bbox[2])}
        )
        y_coordinates = sorted(
            {earlier.y0, earlier.y1}
            | {coordinate for bbox in clipped_overlaps for coordinate in (bbox[1], bbox[3])}
        )
        for x0, x1 in zip(x_coordinates, x_coordinates[1:]):
            for y0, y1 in zip(y_coordinates, y_coordinates[1:]):
                if x0 >= x1 or y0 >= y1:
                    continue
                midpoint = ((x0 + x1) / 2, (y0 + y1) / 2)
                covering_fills = [
                    paint
                    for paint in later_fills
                    if paint.x0 <= midpoint[0] <= paint.x1
                    and paint.y0 <= midpoint[1] <= paint.y1
                ]
                if not covering_fills:
                    continue
                top_fill = max(covering_fills, key=lambda paint: paint.index)
                assert top_fill.color is not None
                if color_delta(earlier.color, top_fill.color) < VISIBLE_FILL_COLOR_DELTA:
                    continue
                occlusion_cells.append(
                    VisibleFillOcclusion(
                        bbox=(x0, y0, x1, y1),
                        under_color=earlier.color,
                        over_color=top_fill.color,
                    )
                )

    return [
        occlusion
        for occlusion in merge_visible_fill_occlusions(occlusion_cells)
        if min(
            occlusion.bbox[2] - occlusion.bbox[0],
            occlusion.bbox[3] - occlusion.bbox[1],
        )
        >= VISIBLE_FILL_MIN_EDGE_PT
        and occlusion.area >= VISIBLE_FILL_MIN_AREA_PT2
    ]


def visible_fill_occlusion_mismatches(
    gt_paints: list[Paint], out_paints: list[Paint]
) -> list[dict]:
    """Pair equivalent occlusions and report only final-coverage differences."""

    gt_occlusions = visible_fill_occlusions(gt_paints)
    out_occlusions = visible_fill_occlusions(out_paints)
    unclaimed_output = list(out_occlusions)
    mismatches: list[tuple[str, VisibleFillOcclusion]] = []
    for gt_occlusion in gt_occlusions:
        best: VisibleFillOcclusion | None = None
        best_distance = VISIBLE_FILL_MATCH_TOLERANCE_PT
        for candidate in unclaimed_output:
            if (
                color_delta(gt_occlusion.under_color, candidate.under_color)
                > LOW_CONTRAST_CHANNEL_DELTA
                or color_delta(gt_occlusion.over_color, candidate.over_color)
                > LOW_CONTRAST_CHANNEL_DELTA
            ):
                continue
            distance = max(
                abs(gt_value - out_value)
                for gt_value, out_value in zip(gt_occlusion.bbox, candidate.bbox)
            )
            if distance <= best_distance:
                best = candidate
                best_distance = distance
        if best is None:
            mismatches.append(("ground_truth", gt_occlusion))
        else:
            unclaimed_output.remove(best)
    mismatches.extend(("output", occlusion) for occlusion in unclaimed_output)

    def report(side: str, occlusion: VisibleFillOcclusion) -> dict:
        return {
            "side": side,
            "bbox": [round(value, 6) for value in occlusion.bbox],
            "area": round(occlusion.area, 6),
            "under_color": [round(value, 6) for value in occlusion.under_color],
            "over_color": [round(value, 6) for value in occlusion.over_color],
        }

    return [report(side, occlusion) for side, occlusion in mismatches]


def clipped_shade_paints(content: str) -> list[Paint]:
    """Collect shadings clipped by one closed axis-aligned rectangle.

    A shading has no useful geometric bounds in mutool's trace; the active
    clip is the only conservative device-space extent available. Unknown or
    non-rectangular outer clips poison nested clips rather than letting a
    smaller rectangle overclaim coverage. Non-extended shadings are skipped
    because their colour ramp may not paint the entire clip (issue #1450).
    """

    no_clip = object()
    unknown_clip = object()
    active_clip: object | tuple[float, float, float, float] = no_clip
    stack: list[object | tuple[float, float, float, float]] = []
    paints: list[Paint] = []

    for event in CLIP_SHADE_EVENT_RE.finditer(content):
        operation = event.group(0)
        clip_match = CLIP_PATH_RE.fullmatch(operation)
        if clip_match:
            stack.append(active_clip)
            bbox = axis_aligned_rect_bbox(*clip_match.groups())
            if bbox is None or active_clip is unknown_clip:
                active_clip = unknown_clip
            elif active_clip is no_clip:
                active_clip = bbox
            else:
                active_clip = intersect_bboxes(active_clip, bbox)
            continue

        shade_match = FILL_SHADE_RE.fullmatch(operation)
        if shade_match:
            attrs = shade_match.group(1)
            if isinstance(active_clip, tuple) and re.search(r'\bextend="1\s+1"', attrs):
                paints.append(
                    Paint(
                        index=event.start(),
                        kind="shade",
                        x0=active_clip[0],
                        y0=active_clip[1],
                        x1=active_clip[2],
                        y1=active_clip[3],
                        alpha=parse_alpha(attrs),
                    )
                )
            continue

        if POP_CLIP_RE.fullmatch(operation):
            active_clip = stack.pop() if stack else no_clip

    return paints


def paint_covers(
    paint: Paint,
    bbox: tuple[float, float, float, float],
    tolerance: float = PAINT_COVER_TOLERANCE_PT,
) -> bool:
    x0, y0, x1, y1 = bbox
    return (
        paint.x0 <= x0 + tolerance
        and paint.y0 <= y0 + tolerance
        and paint.x1 >= x1 - tolerance
        and paint.y1 >= y1 - tolerance
    )


def glyph_bbox(glyph: Glyph) -> tuple[float, float, float, float]:
    """A conservative device-space ink box around one horizontal glyph."""
    return (
        glyph.x,
        glyph.y - glyph.size,
        glyph.x + max(glyph.advance, glyph.size * 0.25),
        glyph.y + glyph.size * 0.25,
    )


def bboxes_overlap(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> bool:
    return (
        first[0] < second[2]
        and first[2] > second[0]
        and first[1] < second[3]
        and first[3] > second[1]
    )


def ignored_text_path_inks(glyph: Glyph, paints: list[Paint]) -> list[Paint]:
    """Compact preceding paths used as an ``ignore_text`` visibility cue.

    This does not identify a font program or prove that the path is a glyph.
    It is a conservative trace-order heuristic observed on issue #1407's
    ground truth: only a path since the preceding text operation and within
    half an em of the ignored glyph is accepted. Pathless invisible/OCR text
    remains hidden; pixel-difference inspection remains the fallback for an
    ambiguous nearby path.
    """
    bbox = glyph_bbox(glyph)
    allowance = glyph.size * 0.5
    return [
        paint
        for paint in paints
        if paint.kind == "path"
        and paint.alpha > INVISIBLE_ALPHA
        and glyph.paint_window_start < paint.index < glyph.paint_window_end
        and paint.x0 >= bbox[0] - allowance
        and paint.y0 >= bbox[1] - allowance
        and paint.x1 <= bbox[2] + allowance
        and paint.y1 <= bbox[3] + allowance
        and bboxes_overlap(
            (paint.x0, paint.y0, paint.x1, paint.y1),
            bbox,
        )
    ]


def glyph_visibility(glyph: Glyph, paints: list[Paint]) -> str:
    if glyph.alpha <= INVISIBLE_ALPHA:
        return "hidden"
    bbox = glyph_bbox(glyph)
    if glyph.visible_clip is not None:
        bbox = intersect_bboxes(bbox, glyph.visible_clip)
        if bbox[0] >= bbox[2] or bbox[1] >= bbox[3]:
            return "hidden"
    own_path_inks = (
        ignored_text_path_inks(glyph, paints) if glyph.needs_path_ink else []
    )
    if glyph.needs_path_ink and not own_path_inks:
        return "hidden"
    if any(
        paint.kind in {"flat", "image", "shade"}
        and paint.index > glyph.paint_index
        and paint.alpha >= OPAQUE_ALPHA
        and paint_covers(paint, bbox)
        for paint in paints
    ):
        return "hidden"

    background = next(
        (
            paint
            for paint in reversed(paints)
            if paint.index < glyph.paint_index
            and paint_covers(paint, bbox)
        ),
        None,
    )
    if (
        glyph.color is not None
        and background is not None
        and background.kind == "flat"
        and background.alpha >= OPAQUE_ALPHA
        and background.color is not None
    ):
        channel_delta = max(
            abs(channel - backdrop)
            for channel, backdrop in zip(glyph.color, background.color)
        )
        if channel_delta <= LOW_CONTRAST_CHANNEL_DELTA:
            return "low_contrast"
    return "painted"


def classify_line_visibility(line: Line, paints: list[Paint]) -> str:
    states = [
        glyph_visibility(glyph, paints)
        for glyph in line.glyphs
        if not glyph.unicode.isspace()
    ]
    if not states:
        return "painted"
    if all(state == "hidden" for state in states):
        return "hidden"
    if all(state in {"hidden", "low_contrast"} for state in states):
        return "low_contrast"
    return "painted"


def parse_trace(trace_xml: str) -> list[PageLayout]:
    pages: list[PageLayout] = []
    for page_match in PAGE_RE.finditer(trace_xml):
        page_attrs, content = page_match.groups()
        media_box = parse_media_box(page_attrs)
        visible_clips = text_visible_clips(content, media_box)
        glyphs: list[Glyph] = []
        rotated_lines: list[Line] = []
        text_operations = list(TEXT_RE.finditer(content))
        for operation_index, op_match in enumerate(text_operations):
            op_kind, op_attrs, op_body = op_match.groups()
            paint_window_start = (
                text_operations[operation_index - 1].end()
                if operation_index > 0
                else 0
            )
            paint_window_end = op_match.start()
            transform = parse_transform(op_attrs)
            if transform is None:
                continue
            a, b, c, d, e, f = transform
            color = parse_rgb(op_attrs)
            alpha = parse_alpha(op_attrs)
            transformed_run: list[Glyph] = []
            for span_attrs, span_body in SPAN_RE.findall(op_body):
                trm = TRM_RE.search(span_attrs)
                size_units = float(trm.group(1)) if trm else 0.0
                size_pt = abs(size_units) * (a * a + b * b) ** 0.5
                for unicode_char, gx, gy, adv in GLYPH_RE.findall(span_body):
                    glyph_x = float(gx)
                    glyph_y = float(gy)
                    transformed_run.append(
                        Glyph(
                            x=a * glyph_x + c * glyph_y + e,
                            y=b * glyph_x + d * glyph_y + f,
                            unicode=unescape(unicode_char),
                            size=size_pt,
                            advance=abs(float(adv) * size_units * a),
                            paint_index=op_match.start(),
                            color=color,
                            alpha=alpha,
                            needs_path_ink=op_kind == "ignore_text",
                            paint_window_start=paint_window_start,
                            paint_window_end=paint_window_end,
                            visible_clip=visible_clips.get(op_match.start(), media_box),
                        )
                    )
            if not transformed_run:
                continue
            if abs(b) > 1e-9 or abs(c) > 1e-9:
                # A rotated/skewed fill_text is already one visual run. Its
                # glyphs have different device y coordinates by construction;
                # feeding them into horizontal baseline bucketing fragments
                # one label into many lines and invents off-page shifts.
                rotated_lines.append(
                    Line(
                        y=min(glyph.y for glyph in transformed_run),
                        glyphs=transformed_run,
                    )
                )
            else:
                glyphs.extend(transformed_run)
        rects: list[Rect] = []
        paints: list[Paint] = []
        for op_match in PATH_RE.finditer(content):
            kind, op_attrs, op_body = op_match.groups()
            transform = parse_transform(op_attrs) or (1, 0, 0, 1, 0, 0)
            a, b, c, d, e, f = transform
            xs: list[float] = []
            ys: list[float] = []
            for px, py in POINT_RE.findall(op_body):
                point_x = float(px)
                point_y = float(py)
                xs.append(a * point_x + c * point_y + e)
                ys.append(b * point_x + d * point_y + f)
            if xs:
                bbox = (min(xs), min(ys), max(xs), max(ys))
                flat_bbox = axis_aligned_rect_bbox(op_attrs, op_body)
                line_bbox = (
                    axis_aligned_line_bbox(op_attrs, op_body)
                    if kind == "stroke_path"
                    else None
                )
                geometry_kind = (
                    "rectangle"
                    if flat_bbox is not None
                    else "line"
                    if line_bbox is not None
                    else "other"
                )
                rects.append(
                    Rect(
                        kind="fill" if kind == "fill_path" else "stroke",
                        geometry_kind=geometry_kind,
                        x0=bbox[0],
                        y0=bbox[1],
                        x1=bbox[2],
                        y1=bbox[3],
                        color=parse_rgb(op_attrs),
                        alpha=parse_alpha(op_attrs),
                    )
                )
                if kind == "fill_path":
                    paints.append(
                        Paint(
                            index=op_match.start(),
                            kind="path",
                            x0=bbox[0],
                            y0=bbox[1],
                            x1=bbox[2],
                            y1=bbox[3],
                            alpha=parse_alpha(op_attrs),
                            color=parse_rgb(op_attrs),
                        )
                    )
                if kind == "fill_path" and flat_bbox is not None:
                    paints.append(
                        Paint(
                            index=op_match.start(),
                            kind="flat",
                            x0=flat_bbox[0],
                            y0=flat_bbox[1],
                            x1=flat_bbox[2],
                            y1=flat_bbox[3],
                            alpha=parse_alpha(op_attrs),
                            color=parse_rgb(op_attrs),
                        )
                    )
        for op_match in FILL_IMAGE_RE.finditer(content):
            op_attrs = op_match.group(1)
            transform = parse_transform(op_attrs)
            if transform is None:
                continue
            a, b, c, d, _, _ = transform
            # A rotated image's axis-aligned bbox can include areas it never
            # paints, so only use rectangular page/background images here.
            if abs(b) > 1e-9 or abs(c) > 1e-9:
                continue
            x0, y0, x1, y1 = transformed_bbox(
                transform, [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)]
            )
            paints.append(
                Paint(
                    index=op_match.start(),
                    kind="image",
                    x0=x0,
                    y0=y0,
                    x1=x1,
                    y1=y1,
                    alpha=parse_alpha(op_attrs),
                )
            )
        paints.extend(clipped_shade_paints(content))
        paints.sort(key=lambda paint: paint.index)
        lines = build_lines(glyphs)
        lines.extend(line for line in rotated_lines if line.key)
        for line in lines:
            line.visibility = classify_line_visibility(line, paints)
        lines.sort(key=lambda line: (line.y, line.x0))
        pages.append(
            PageLayout(
                lines=lines,
                rects=rects,
                paints=paints,
                media_box=media_box,
            )
        )
    return pages


def build_lines(glyphs: list[Glyph], y_tolerance: float = LINE_Y_TOLERANCE_PT) -> list[Line]:
    lines: list[Line] = []
    for glyph in sorted(glyphs, key=lambda g: (g.y, g.x)):
        target = None
        for line in lines:
            if abs(line.y - glyph.y) <= y_tolerance:
                target = line
                break
        if target is None:
            target = Line(y=glyph.y)
            lines.append(target)
        target.glyphs.append(glyph)
    for line in lines:
        line.glyphs.sort(key=lambda g: g.x)
        line.y = statistics.median(g.y for g in line.glyphs)
    lines = [line for line in lines if line.key]
    lines.sort(key=lambda line: (line.y, line.x0))
    return lines


def match_lines(
    gt_lines: list[Line], out_lines: list[Line]
) -> tuple[list[tuple[Line, Line]], list[Line], list[Line]]:
    """Pair every exact-text line by minimum spatial distance.

    A sequence matcher can align repeated labels to the wrong occurrence, and
    the older render harness discarded repeats entirely. Grouping by text and
    assigning in x/y space preserves every chart title, legend, tick, and table
    value as an independently auditable instance.
    """
    gt_groups: dict[str, list[Line]] = {}
    out_groups: dict[str, list[Line]] = {}
    for line in gt_lines:
        gt_groups.setdefault(line.key, []).append(line)
    for line in out_lines:
        out_groups.setdefault(line.key, []).append(line)

    matches: list[tuple[Line, Line]] = []
    matched_gt: set[int] = set()
    matched_out: set[int] = set()
    for key, references in gt_groups.items():
        candidates = out_groups.get(key, [])
        references = sorted(references, key=lambda line: (line.y, line.x0))
        candidates = sorted(candidates, key=lambda line: (line.y, line.x0))
        for reference_index, candidate_index in minimum_cost_pairs(
            [(line.x0, line.y) for line in references],
            [(line.x0, line.y) for line in candidates],
        ):
            reference = references[reference_index]
            candidate = candidates[candidate_index]
            matches.append((reference, candidate))
            matched_gt.add(id(reference))
            matched_out.add(id(candidate))
    matches.sort(key=lambda pair: (pair[0].y, pair[0].x0))
    missing = [line for line in gt_lines if id(line) not in matched_gt]
    extra = [line for line in out_lines if id(line) not in matched_out]
    return matches, missing, extra


def split_distant_text_objects(line: Line) -> list[Line]:
    """Split a trace line only at unambiguously independent text objects.

    ``build_lines`` intentionally buckets every horizontal glyph on a shared
    baseline. That can merge an axis tick at the left of a chart with a legend
    label hundreds of points away. A boundary is safe only when it crosses
    both a PDF text-paint operation and a large visible gap. Styled runs inside
    one text operation and ordinary word spacing therefore remain intact.
    """
    glyphs = line.visible_glyphs
    if not glyphs:
        return []

    groups: list[list[Glyph]] = [[glyphs[0]]]
    for previous, glyph in zip(glyphs, glyphs[1:]):
        gap = glyph.x - (previous.x + previous.advance)
        minimum_gap = max(
            SAFE_TOPOLOGY_GAP_PT,
            2.0 * max(previous.size, glyph.size),
        )
        if previous.paint_index != glyph.paint_index and gap >= minimum_gap:
            groups.append([])
        groups[-1].append(glyph)

    return [
        Line(
            y=statistics.median(glyph.y for glyph in group),
            glyphs=group,
            visibility=line.visibility,
        )
        for group in groups
    ]


def ordered_unique_segment_match(
    joined_line: Line, candidates: list[Line]
) -> tuple[list[Line], list[Line]] | None:
    """Return one-to-many split/join matches when order and count are exact."""
    segments = split_distant_text_objects(joined_line)
    if len(segments) < 2:
        return None

    selected: list[Line] = []
    for segment in segments:
        same_text = [line for line in candidates if line.key == segment.key]
        # Ambiguous occurrences can hide a duplicated label, so leave the
        # whole group unmatched for the normal missing/extra audit.
        if len(same_text) != 1:
            return None
        selected.append(same_text[0])
    if len({id(line) for line in selected}) != len(selected):
        return None
    if any(len(split_distant_text_objects(line)) != 1 for line in selected):
        return None
    ordered = sorted(selected, key=lambda line: (line.y, line.x0))
    if [id(line) for line in selected] != [id(line) for line in ordered]:
        # A split/join changes PDF object topology, never reading order.
        return None
    return segments, selected


def take_topology_equivalents(
    missing: list[Line], extra: list[Line]
) -> tuple[list[list[tuple[Line, Line]]], dict, list[Line], list[Line]]:
    """Match safe one-to-many line groups without suppressing geometry checks."""
    match_groups: list[list[tuple[Line, Line]]] = []
    groups: list[tuple[int, int, str]] = []
    remaining_missing = list(missing)
    remaining_extra = list(extra)

    for joined_gt in list(remaining_missing):
        result = ordered_unique_segment_match(joined_gt, remaining_extra)
        if result is None:
            continue
        gt_segments, out_lines = result
        match_groups.append(list(zip(gt_segments, out_lines)))
        groups.append((1, len(out_lines), " + ".join(line.key for line in gt_segments)))
        remaining_missing.remove(joined_gt)
        for line in out_lines:
            remaining_extra.remove(line)

    for joined_out in list(remaining_extra):
        result = ordered_unique_segment_match(joined_out, remaining_missing)
        if result is None:
            continue
        out_segments, gt_lines = result
        match_groups.append(list(zip(gt_lines, out_segments)))
        groups.append((len(gt_lines), 1, " + ".join(line.key for line in out_segments)))
        remaining_extra.remove(joined_out)
        for line in gt_lines:
            remaining_missing.remove(line)

    info = {
        "groups": len(groups),
        "gt_lines": sum(gt_count for gt_count, _, _ in groups),
        "out_lines": sum(out_count for _, out_count, _ in groups),
        "samples": [sample[:120] for _, _, sample in groups[:5]],
    }
    return match_groups, info, remaining_missing, remaining_extra


def take_wrap_differences(
    missing: list[Line], extra: list[Line]
) -> tuple[list[str], list[Line], list[Line]]:
    """Pull re-wrapped paragraphs out of the missing/extra pools.

    A wrap difference is text that exists on both sides but breaks at a
    different point, so joined runs of unmatched lines carry the same
    characters. Greedy prefix consumption keeps it linear and is sufficient
    for the unmatched lines left after exact-text spatial pairing.
    """
    wraps: list[str] = []
    remaining_missing = list(missing)
    remaining_extra = list(extra)
    while remaining_missing and remaining_extra:
        gt_join = "".join(line.key for line in remaining_missing)
        out_join = "".join(line.key for line in remaining_extra)
        if not gt_join or gt_join != out_join:
            break
        wraps.append(remaining_missing[0].key[:60])
        remaining_missing.clear()
        remaining_extra.clear()
    return wraps, remaining_missing, remaining_extra


def take_reflows(missing: list[Line], extra: list[Line]) -> tuple[dict, list[Line], list[Line]]:
    """Classify unmatched lines whose characters all survive as reflow, not loss.

    Same multiset of characters on both sides means the text is present but
    grouped into different lines (baseline splits, cell order) — the audit's
    invoice wrap-flip and budget baseline-split cases. Real loss keeps its
    asymmetric remainder in missing/extra.
    """
    gt_chars = sorted("".join(line.key for line in missing))
    out_chars = sorted("".join(line.key for line in extra))
    if gt_chars and gt_chars == out_chars:
        info = {
            "gt_lines": len(missing),
            "out_lines": len(extra),
            "samples": [line.key[:60] for line in missing[:3]],
        }
        return info, [], []
    return {"gt_lines": 0, "out_lines": 0, "samples": []}, missing, extra


def same_rect_paint(first: Rect, second: Rect) -> bool:
    if (
        first.kind != second.kind
        or first.geometry_kind != second.geometry_kind
        or abs(first.alpha - second.alpha) > 1e-6
    ):
        return False
    if first.color is None or second.color is None:
        return first.color is second.color
    return all(
        abs(left - right) <= 1e-6
        for left, right in zip(first.color, second.color)
    )


def merge_rect_groups(
    groups: list[CanonicalRect], *, horizontal: bool, tolerance: float
) -> list[CanonicalRect]:
    """Merge collinear same-paint pieces without filling an L-shaped gap."""

    def can_merge(first: Rect, second: Rect) -> bool:
        if not same_rect_paint(first, second):
            return False
        same_bbox = all(
            abs(left - right) <= tolerance
            for left, right in zip(first.bbox, second.bbox)
        )
        if same_bbox:
            return True
        if horizontal:
            same_band = (
                abs(first.y0 - second.y0) <= tolerance
                and abs(first.y1 - second.y1) <= tolerance
            )
            intervals_touch = (
                abs(first.x1 - second.x0) <= tolerance
                or abs(second.x1 - first.x0) <= tolerance
            )
        else:
            same_band = (
                abs(first.x0 - second.x0) <= tolerance
                and abs(first.x1 - second.x1) <= tolerance
            )
            intervals_touch = (
                abs(first.y1 - second.y0) <= tolerance
                or abs(second.y1 - first.y0) <= tolerance
            )
        return same_band and intervals_touch

    ordered = sorted(
        groups,
        key=lambda group: (
            group.rect.kind,
            group.rect.y0 if horizontal else group.rect.x0,
            group.rect.y1 if horizontal else group.rect.x1,
            group.rect.x0 if horizontal else group.rect.y0,
        ),
    )
    merged: list[CanonicalRect] = []
    for group in ordered:
        target_index = next(
            (
                index
                for index, existing in enumerate(merged)
                if can_merge(existing.rect, group.rect)
            ),
            None,
        )
        if target_index is None:
            merged.append(group)
            continue
        existing = merged[target_index]
        merged[target_index] = CanonicalRect(
            rect=Rect(
                kind=existing.rect.kind,
                geometry_kind=existing.rect.geometry_kind,
                x0=min(existing.rect.x0, group.rect.x0),
                y0=min(existing.rect.y0, group.rect.y0),
                x1=max(existing.rect.x1, group.rect.x1),
                y1=max(existing.rect.y1, group.rect.y1),
                color=existing.rect.color,
                alpha=existing.rect.alpha,
            ),
            source_indices=tuple(sorted(existing.source_indices + group.source_indices)),
        )
    return merged


def canonical_rects(rects: list[Rect], tolerance: float) -> list[CanonicalRect]:
    groups = [
        CanonicalRect(rect=rect, source_indices=(index,))
        for index, rect in enumerate(rects)
        if rect.geometry_kind != "other"
    ]
    groups = merge_rect_groups(groups, horizontal=True, tolerance=tolerance)
    groups = merge_rect_groups(groups, horizontal=False, tolerance=tolerance)
    return sorted(
        groups,
        key=lambda group: (
            group.rect.kind,
            group.rect.y0,
            group.rect.x0,
            group.rect.height,
            group.rect.width,
        ),
    )


def rect_feature(rect: Rect) -> tuple[float, float, float, float]:
    """Represent both placement and extent for minimum-cost assignment."""
    return (rect.center[0], rect.center[1], rect.width / 2, rect.height / 2)


def rect_match_topology(rect: Rect) -> str:
    """Separate thin boundary operations from painted areas."""
    if rect.geometry_kind != "rectangle":
        return rect.geometry_kind
    if (
        rect.height <= RECT_RULE_MAX_THICKNESS_PT
        and rect.width >= RECT_RULE_MIN_LENGTH_PT
    ):
        return "horizontal-rule"
    if (
        rect.width <= RECT_RULE_MAX_THICKNESS_PT
        and rect.height >= RECT_RULE_MIN_LENGTH_PT
    ):
        return "vertical-rule"
    return "area"


def same_rect_match_bucket(first: Rect, second: Rect) -> bool:
    return same_rect_paint(first, second) and rect_match_topology(
        first
    ) == rect_match_topology(second)


def ambiguous_rect_pair_allowed(reference: Rect, candidate: Rect) -> bool:
    center_delta = max(
        abs(candidate.center[0] - reference.center[0]),
        abs(candidate.center[1] - reference.center[1]),
    )
    if center_delta > RECT_AMBIGUOUS_MATCH_RADIUS_PT:
        return False
    if min(reference.width, reference.height, candidate.width, candidate.height) <= 0:
        return False
    extent_ratio = max(
        reference.width / candidate.width,
        candidate.width / reference.width,
        reference.height / candidate.height,
        candidate.height / reference.height,
    )
    return extent_ratio <= RECT_AMBIGUOUS_MAX_EXTENT_RATIO


def match_rect_groups(
    references: list[CanonicalRect], candidates: list[CanonicalRect]
) -> list[tuple[CanonicalRect, CanonicalRect]]:
    """Match rectangles without forcing unequal groups across paint/topology."""
    if len(references) == len(candidates):
        return [
            (references[reference_index], candidates[candidate_index])
            for reference_index, candidate_index in minimum_cost_pairs(
                [rect_feature(group.rect) for group in references],
                [rect_feature(group.rect) for group in candidates],
            )
        ]

    matches: list[tuple[CanonicalRect, CanonicalRect]] = []
    remaining_references = list(references)
    remaining_candidates = list(candidates)
    while remaining_references:
        seed = remaining_references[0]
        reference_bucket = [
            group
            for group in remaining_references
            if same_rect_match_bucket(seed.rect, group.rect)
        ]
        candidate_bucket = [
            group
            for group in remaining_candidates
            if same_rect_match_bucket(seed.rect, group.rect)
        ]
        remaining_references = [
            group
            for group in remaining_references
            if not same_rect_match_bucket(seed.rect, group.rect)
        ]
        remaining_candidates = [
            group
            for group in remaining_candidates
            if not same_rect_match_bucket(seed.rect, group.rect)
        ]
        if not candidate_bucket:
            continue

        allowed_pairs = [
            [
                ambiguous_rect_pair_allowed(reference.rect, candidate.rect)
                for candidate in candidate_bucket
            ]
            for reference in reference_bucket
        ]
        matches.extend(
            (reference_bucket[reference_index], candidate_bucket[candidate_index])
            for reference_index, candidate_index in minimum_cost_pairs(
                [rect_feature(group.rect) for group in reference_bucket],
                [rect_feature(group.rect) for group in candidate_bucket],
                allowed_pairs=allowed_pairs,
            )
        )
    return matches


def rect_sample(gt_group: CanonicalRect, out_group: CanonicalRect) -> dict:
    gt_rect = gt_group.rect
    out_rect = out_group.rect
    dx = out_rect.x0 - gt_rect.x0
    dy = out_rect.y0 - gt_rect.y0
    dwidth = out_rect.width - gt_rect.width
    dheight = out_rect.height - gt_rect.height
    edges = {
        "left": dx,
        "top": dy,
        "right": out_rect.x1 - gt_rect.x1,
        "bottom": out_rect.y1 - gt_rect.y1,
    }
    deltas = [dx, dy, dwidth, dheight, *edges.values()]
    return {
        "kind": gt_rect.kind,
        "geometry_kind": gt_rect.geometry_kind,
        "gt_indices": list(gt_group.source_indices),
        "out_indices": list(out_group.source_indices),
        "gt_bbox": gt_rect.bbox,
        "out_bbox": out_rect.bbox,
        "dx": dx,
        "dy": dy,
        "dwidth": dwidth,
        "dheight": dheight,
        "edges": edges,
        "center_dx": out_rect.center[0] - gt_rect.center[0],
        "center_dy": out_rect.center[1] - gt_rect.center[1],
        "max_abs_delta": max(abs(delta) for delta in deltas),
    }


def diff_page(
    gt: PageLayout,
    out: PageLayout,
    noise_floor: float = 0.12,
    large_shift: float = 5.0,
    fine_shift: float | None = None,
) -> dict:
    exact_matches, missing, extra = match_lines(gt.lines, out.lines)
    hidden_missing = [line for line in missing if line.visibility == "hidden"]
    hidden_extra = [line for line in extra if line.visibility == "hidden"]
    missing = [line for line in missing if line.visibility != "hidden"]
    extra = [line for line in extra if line.visibility != "hidden"]
    topology_groups, topology, missing, extra = take_topology_equivalents(missing, extra)
    topology_matches = [pair for group in topology_groups for pair in group]
    for gt_line, out_line in topology_matches:
        # The original joined line can mix visible and occluded distant
        # objects. Reclassify each synthetic segment so topology equivalence
        # cannot suppress an independently real visibility mismatch.
        gt_line.visibility = classify_line_visibility(gt_line, gt.paints)
        out_line.visibility = classify_line_visibility(out_line, out.paints)
    matches = exact_matches + topology_matches
    matches.sort(key=lambda pair: (pair[0].y, pair[0].x0))
    wraps, missing, extra = take_wrap_differences(missing, extra)
    reflow, missing, extra = take_reflows(missing, extra)

    dys = [out_line.y - gt_line.y for gt_line, out_line in matches]
    dx0s = [out_line.x0 - gt_line.x0 for gt_line, out_line in matches]
    width_pcts = [
        (out_line.width - gt_line.width) / gt_line.width * 100
        for gt_line, out_line in matches
        if gt_line.width > 1.0
    ]
    deviant_lines = sum(
        1
        for gt_line, out_line in exact_matches
        if abs(out_line.y - gt_line.y) > noise_floor
    ) + sum(
        any(abs(out_line.y - gt_line.y) > noise_floor for gt_line, out_line in group)
        for group in topology_groups
    )

    gt_occurrences: dict[str, list[Line]] = {}
    for line, _ in matches:
        gt_occurrences.setdefault(line.key, []).append(line)
    occurrence_by_id: dict[int, tuple[int, int]] = {}
    for repeated in gt_occurrences.values():
        ordered_occurrences = sorted(repeated, key=lambda line: (line.y, line.x0))
        for index, line in enumerate(ordered_occurrences, start=1):
            occurrence_by_id[id(line)] = (index, len(ordered_occurrences))

    def instance_label(line: Line) -> str:
        occurrence, occurrence_count = occurrence_by_id[id(line)]
        label = line.key[:60]
        if occurrence_count > 1:
            label = f"{label} [{occurrence}/{occurrence_count}]"
        return label

    instances: list[dict] = []
    for gt_line, out_line in matches:
        width_pct = (
            (out_line.width - gt_line.width) / gt_line.width * 100
            if gt_line.width > 1.0
            else 0.0
        )
        instances.append(
            {
                "label": instance_label(gt_line),
                "gt_x": gt_line.x0,
                "gt_y": gt_line.y,
                "out_x": out_line.x0,
                "out_y": out_line.y,
                "dx": out_line.x0 - gt_line.x0,
                "dy": out_line.y - gt_line.y,
                "width_pct": width_pct,
            }
        )
    visibility_mismatches = [
        {
            "label": instance_label(gt_line),
            "gt": gt_line.visibility,
            "out": out_line.visibility,
        }
        for gt_line, out_line in matches
        if (gt_line.visibility == "painted") != (out_line.visibility == "painted")
    ]
    visible_fill_mismatches = visible_fill_occlusion_mismatches(gt.paints, out.paints)
    large_shifts = [
        instance
        for instance in instances
        if abs(instance["dx"]) > large_shift or abs(instance["dy"]) > large_shift
    ]
    fine_shifts = (
        [
            instance
            for instance in instances
            if abs(instance["dx"]) > fine_shift or abs(instance["dy"]) > fine_shift
        ]
        if fine_shift is not None
        else []
    )

    pitch_deltas: list[float] = []
    matched_gt = {id(gt_line) for gt_line, _ in matches}
    ordered = [pair for pair in matches]
    ordered.sort(key=lambda pair: pair[0].y)
    for (gt_a, out_a), (gt_b, out_b) in zip(ordered, ordered[1:]):
        if id(gt_a) in matched_gt and id(gt_b) in matched_gt:
            pitch_deltas.append((out_b.y - out_a.y) - (gt_b.y - gt_a.y))

    worst_dy_index = max(range(len(dys)), key=lambda i: abs(dys[i]), default=None)

    def stats(values: list[float]) -> dict:
        if not values:
            return {"mean_abs": 0.0, "worst": 0.0}
        return {
            "mean_abs": sum(abs(v) for v in values) / len(values),
            "worst": max(values, key=abs),
        }

    canonical_gt_rects = canonical_rects(gt.rects, noise_floor)
    canonical_out_rects = canonical_rects(out.rects, noise_floor)
    rect_matches: list[tuple[CanonicalRect, CanonicalRect]] = []
    for kind, geometry_kind in (
        ("fill", "rectangle"),
        ("stroke", "rectangle"),
        ("stroke", "line"),
    ):
        references = [
            group
            for group in canonical_gt_rects
            if (group.rect.kind, group.rect.geometry_kind) == (kind, geometry_kind)
        ]
        candidates = [
            group
            for group in canonical_out_rects
            if (group.rect.kind, group.rect.geometry_kind) == (kind, geometry_kind)
        ]
        rect_matches.extend(match_rect_groups(references, candidates))

    rect_samples = [
        rect_sample(reference, candidate) for reference, candidate in rect_matches
    ]
    rect_samples.sort(
        key=lambda sample: (-sample["max_abs_delta"], sample["gt_indices"])
    )
    rect_geometry_threshold = fine_shift if fine_shift is not None else large_shift
    rect_geometry_mismatches = [
        sample
        for sample in rect_samples
        if sample["max_abs_delta"] > rect_geometry_threshold
    ]
    rect_center_deltas = [
        max(abs(sample["center_dx"]), abs(sample["center_dy"]))
        for sample in rect_samples
    ]

    rect_gt_ids = {id(reference) for reference, _ in rect_matches}
    rect_out_ids = {id(candidate) for _, candidate in rect_matches}

    dy_stats = stats(dys)
    return {
        "lines": {
            "gt": len(gt.lines),
            "out": len(out.lines),
            "matched": len(exact_matches) + topology["groups"],
            "missing": len(missing),
            "extra": len(extra),
            "deviant": deviant_lines,
            "missing_text": [line.key[:60] for line in missing[:5]],
            "extra_text": [line.key[:60] for line in extra[:5]],
        },
        "baseline": {
            "mean_abs_dy": dy_stats["mean_abs"],
            "worst_dy": abs(dy_stats["worst"]),
            "worst_dy_signed": dy_stats["worst"],
            "worst_line": (
                matches[worst_dy_index][0].key[:60] if worst_dy_index is not None else ""
            ),
        },
        "dx0": stats(dx0s),
        "width": {
            "mean_abs_pct": stats(width_pcts)["mean_abs"],
            "worst_pct": abs(stats(width_pcts)["worst"]),
        },
        "instances": {
            "compared": len(matches),
            "large_shift_threshold": large_shift,
            "large_shift_count": len(large_shifts),
            "large_shifts": large_shifts,
            "fine_shift_threshold": fine_shift,
            "fine_shift_count": len(fine_shifts),
            "fine_shifts": fine_shifts,
        },
        "visibility": {
            "mismatch_count": len(visibility_mismatches),
            "mismatches": visibility_mismatches,
            "unmatched_hidden_gt": len(hidden_missing),
            "unmatched_hidden_out": len(hidden_extra),
            "unmatched_hidden_gt_text": [
                line.key[:60] for line in hidden_missing[:5]
            ],
            "unmatched_hidden_out_text": [
                line.key[:60] for line in hidden_extra[:5]
            ],
        },
        "visible_fills": {
            "mismatch_count": len(visible_fill_mismatches),
            "mismatches": visible_fill_mismatches,
        },
        "pitch": {"pairs": len(pitch_deltas), "worst_delta": abs(stats(pitch_deltas)["worst"])},
        "wraps": {"count": len(wraps), "samples": wraps[:5]},
        "topology": topology,
        "reflow": reflow,
        "rects": {
            "gt_count": len(gt.rects),
            "out_count": len(out.rects),
            "canonical_gt_count": len(canonical_gt_rects),
            "canonical_out_count": len(canonical_out_rects),
            "matched": len(rect_matches),
            "unmatched_gt": sum(
                id(group) not in rect_gt_ids for group in canonical_gt_rects
            ),
            "unmatched_out": sum(
                id(group) not in rect_out_ids for group in canonical_out_rects
            ),
            "mean_center_delta": stats(rect_center_deltas)["mean_abs"],
            "geometry_threshold": rect_geometry_threshold,
            "geometry_mismatch_count": len(rect_geometry_mismatches),
            "geometry_mismatch_samples": rect_geometry_mismatches[:RECT_SAMPLE_LIMIT],
            "samples": rect_samples[:RECT_SAMPLE_LIMIT],
            "x": stats([sample["dx"] for sample in rect_samples]),
            "y": stats([sample["dy"] for sample in rect_samples]),
            "width": stats([sample["dwidth"] for sample in rect_samples]),
            "height": stats([sample["dheight"] for sample in rect_samples]),
            "edges": {
                edge: stats([sample["edges"][edge] for sample in rect_samples])
                for edge in ("left", "top", "right", "bottom")
            },
        },
        "noise_floor": noise_floor,
    }


def render_reading(vectors: list[dict]) -> str:
    lines = ["## Reading", ""]
    for index, vector in enumerate(vectors, start=1):
        page_notes: list[str] = []
        line_info = vector["lines"]
        if line_info["missing"] or line_info["extra"]:
            page_notes.append(
                f"content differs: {line_info['missing']} GT line(s) unmatched "
                f"(e.g. {line_info['missing_text'][:1]}), {line_info['extra']} extra"
            )
        if vector["wraps"]["count"]:
            page_notes.append(
                f"{vector['wraps']['count']} paragraph(s) wrap at a different word "
                "(text present, break point moved) — check advances, kerning, or spacing"
            )
        if vector["topology"]["groups"]:
            page_notes.append(
                f"{vector['topology']['groups']} distant text-object group(s) use "
                "different PDF line splits/joins with content order preserved; "
                "their individual anchors remain position-audited"
            )
        if vector["reflow"]["gt_lines"]:
            page_notes.append(
                f"{vector['reflow']['gt_lines']} GT line(s) re-grouped into "
                f"{vector['reflow']['out_lines']} output line(s) with no text lost — "
                "baseline splits or a moved wrap, not missing content"
            )
        if line_info["deviant"]:
            page_notes.append(
                f"{line_info['deviant']}/{line_info['matched']} matched lines sit past the "
                f"{vector['noise_floor']}pt floor; worst {vector['baseline']['worst_dy_signed']:+.2f}pt "
                f"on '{vector['baseline']['worst_line']}'"
            )
        if vector["instances"]["large_shift_count"]:
            examples = "; ".join(
                f"'{item['label']}' dx {item['dx']:+.2f}pt, dy {item['dy']:+.2f}pt"
                for item in sorted(
                    vector["instances"]["large_shifts"],
                    key=lambda value: max(abs(value["dx"]), abs(value["dy"])),
                    reverse=True,
                )
            )
            page_notes.append(
                f"{vector['instances']['large_shift_count']} matched text instance(s) move "
                f"past {vector['instances']['large_shift_threshold']:.2f}pt: {examples}. "
                "These are layout differences, not antialiasing; inspect and track each one"
            )
        if vector["instances"]["fine_shift_count"]:
            examples = "; ".join(
                f"'{item['label']}' dx {item['dx']:+.2f}pt, dy {item['dy']:+.2f}pt"
                for item in sorted(
                    vector["instances"]["fine_shifts"],
                    key=lambda value: max(abs(value["dx"]), abs(value["dy"])),
                    reverse=True,
                )
            )
            page_notes.append(
                f"{vector['instances']['fine_shift_count']} matched text instance(s) move "
                f"past the fine-detail {vector['instances']['fine_shift_threshold']:.2f}pt "
                f"tolerance: {examples}. Trace-derived anchor movement is geometry, not "
                "font antialiasing; inspect and track each one"
            )
        if vector["visibility"]["mismatch_count"]:
            examples = "; ".join(
                f"'{item['label']}' GT {item['gt']} vs output {item['out']}"
                for item in vector["visibility"]["mismatches"]
            )
            page_notes.append(
                f"{vector['visibility']['mismatch_count']} matched text instance(s) differ in "
                f"painted visibility: {examples}. Check z-order, opacity, or foreground/background "
                "contrast"
            )
        hidden_gt = vector["visibility"].get("unmatched_hidden_gt", 0)
        hidden_out = vector["visibility"].get("unmatched_hidden_out", 0)
        if hidden_gt or hidden_out:
            page_notes.append(
                f"{hidden_gt} GT and {hidden_out} output unmatched trace line(s) are "
                "proven non-painted by the trace visibility analysis and "
                "excluded from visual missing/extra findings"
            )
        if vector["visible_fills"]["mismatch_count"]:
            examples = "; ".join(
                f"{item['side']} bbox {item['bbox']} {item['under_color']} -> {item['over_color']}"
                for item in vector["visible_fills"]["mismatches"][:3]
            )
            page_notes.append(
                f"{vector['visible_fills']['mismatch_count']} visible fill occlusion(s) "
                f"differ after paint order: {examples}. Check border/fill z-order"
            )
        if vector["pitch"]["worst_delta"] > vector["noise_floor"]:
            page_notes.append(
                f"line pitch drifts up to {vector['pitch']['worst_delta']:.2f}pt — "
                "line-height model rather than block placement"
            )
        if vector["rects"]["gt_count"] != vector["rects"]["out_count"]:
            page_notes.append(
                f"rect census differs: GT {vector['rects']['gt_count']} vs "
                f"output {vector['rects']['out_count']} raw draw ops; touching same-paint "
                "axis-aligned same-paint normalization yields "
                f"{vector['rects']['canonical_gt_count']} vs "
                f"{vector['rects']['canonical_out_count']} comparable rectangle/line "
                "primitives. Raw count differences can be non-rectangular paths or "
                "operation splitting, so confirm unmatched geometry visually before "
                "reading it as missing ink"
            )
        if vector["rects"]["geometry_mismatch_count"]:
            examples = "; ".join(
                f"{item['kind']} GT {item['gt_bbox']} -> output {item['out_bbox']} "
                f"(dx {item['dx']:+.2f}, dy {item['dy']:+.2f}, "
                f"dw {item['dwidth']:+.2f}, dh {item['dheight']:+.2f}pt)"
                for item in vector["rects"]["geometry_mismatch_samples"]
            )
            page_notes.append(
                f"{vector['rects']['geometry_mismatch_count']} matched rectangle geometry "
                f"deviation(s) exceed {vector['rects']['geometry_threshold']:.2f}pt: "
                f"{examples}. Inspect the recorded source operation indices and corresponding "
                "render cluster"
            )
        if not page_notes:
            page_notes.append("no deviation past the noise floor")
        lines.append(f"- page {index}: " + "; ".join(page_notes))
    return "\n".join(lines)


def audit_failures(vectors: list[dict]) -> int:
    """Count material layout or paint findings that require visual disposition."""
    return sum(
        (
            vector["instances"]["fine_shift_count"]
            if vector["instances"]["fine_shift_threshold"] is not None
            else vector["instances"]["large_shift_count"]
        )
        + vector["visibility"]["mismatch_count"]
        + vector["visible_fills"]["mismatch_count"]
        + vector["rects"]["geometry_mismatch_count"]
        + vector["lines"]["missing"]
        + vector["lines"]["extra"]
        + vector["wraps"]["count"]
        + int(bool(vector["reflow"]["gt_lines"] or vector["reflow"]["out_lines"]))
        for vector in vectors
    )


def run_mutool(pdf: Path) -> str:
    mutool = shutil.which("mutool")
    if mutool is None:
        sys.exit("mutool is required (brew install mupdf-tools)")
    result = subprocess.run(
        [mutool, "draw", "-F", "trace", str(pdf)],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("gt", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--page", type=int, help="1-based page to compare (default: all)")
    parser.add_argument(
        "--noise-floor",
        type=float,
        default=0.12,
        help="pt threshold under which a delta is measurement noise (0.12 Word GT, 0.5 Excel GT)",
    )
    parser.add_argument(
        "--large-shift",
        type=float,
        default=5.0,
        metavar="PT",
        help=(
            "flag any matched text instance or rectangle geometry component that moves "
            "by more than this (default: 5pt)"
        ),
    )
    parser.add_argument(
        "--fine-shift",
        type=float,
        metavar="PT",
        help=(
            "enable the fine-detail audit gate and flag matched text x/y or rectangle "
            "geometry beyond PT; does not replace the 5pt coarse text summary"
        ),
    )
    parser.add_argument(
        "--audit",
        action="store_true",
        help=(
            "exit nonzero on missing/extra/reflowed text, changed wraps, "
            "a text/fill visibility mismatch, an active fine/coarse text or "
            "rectangle geometry shift, or a page-count mismatch"
        ),
    )
    parser.add_argument("--json", action="store_true", help="emit the deviation vectors as JSON")
    args = parser.parse_args()

    if args.fine_shift is not None and args.fine_shift < args.noise_floor:
        parser.error("--fine-shift must be at least --noise-floor")

    gt_pages = parse_trace(run_mutool(args.gt))
    out_pages = parse_trace(run_mutool(args.output))

    # A trace that yields no pages means the parser did not understand mutool's
    # output, not that the two files agree. Reporting an empty diff here would
    # be indistinguishable from a clean comparison.
    for label, path, pages in (("GT", args.gt, gt_pages), ("output", args.output, out_pages)):
        if not pages:
            sys.exit(f"no pages parsed from the {label} trace ({path}) — cannot compare")

    if args.page is not None:
        gt_pages = gt_pages[args.page - 1 : args.page]
        out_pages = out_pages[args.page - 1 : args.page]

    count = min(len(gt_pages), len(out_pages))
    vectors = [
        diff_page(
            gt_pages[i],
            out_pages[i],
            noise_floor=args.noise_floor,
            large_shift=args.large_shift,
            fine_shift=args.fine_shift,
        )
        for i in range(count)
    ]

    if args.json:
        print(
            json.dumps(
                {"pages": vectors, "gt_pages": len(gt_pages), "out_pages": len(out_pages)},
                indent=1,
            )
        )
        if args.audit and (audit_failures(vectors) or len(gt_pages) != len(out_pages)):
            return 1
        return 0

    if len(gt_pages) != len(out_pages):
        print(f"PAGE COUNT MISMATCH: GT {len(gt_pages)} vs output {len(out_pages)}\n")
    for index, vector in enumerate(vectors, start=1):
        line_info = vector["lines"]
        fine_shift_summary = (
            f"fine shifts {vector['instances']['fine_shift_count']} | "
            if vector["instances"]["fine_shift_threshold"] is not None
            else ""
        )
        print(
            f"page {index}: {line_info['matched']} matched "
            f"({line_info['missing']} missing, {line_info['extra']} extra, "
            f"{vector['wraps']['count']} re-wrapped) | "
            f"dy mean {vector['baseline']['mean_abs_dy']:.2f}pt worst "
            f"{vector['baseline']['worst_dy_signed']:+.2f}pt | "
            f"pitch worst {vector['pitch']['worst_delta']:.2f}pt | "
            f"width worst {vector['width']['worst_pct']:.1f}% | "
            f"large shifts {vector['instances']['large_shift_count']} | "
            f"{fine_shift_summary}"
            f"visibility {vector['visibility']['mismatch_count']} | "
            f"visible fills {vector['visible_fills']['mismatch_count']} | "
            f"rects {vector['rects']['gt_count']}/{vector['rects']['out_count']} "
            f"(geometry {vector['rects']['geometry_mismatch_count']})"
        )
    print()
    print(render_reading(vectors))
    failures = audit_failures(vectors)
    if args.audit and (failures or len(gt_pages) != len(out_pages)):
        print()
        print(
            f"AUDIT FAILED: {failures} unresolved layout/paint finding(s) require visual "
            "inspection and an issue reference before the audit can be marked as matching."
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
