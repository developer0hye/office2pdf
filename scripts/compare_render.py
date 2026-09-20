#!/usr/bin/env python3
"""Compare a rendered PDF against a native Office export across several axes.

No single measure is trustworthy alone, and each one's blind spot is
another's strength:

- **Geometry** catches position, size, and pitch. It is what actually
  changes when a layout bug is fixed, and it is the only axis that can
  distinguish "moved to the right place" from "moved somewhere else".
  Blind to colour and to elements that are absent entirely.
- **Histogram** catches fill colour, recolouring, and missing elements,
  because it counts what is drawn without caring where. Blind to position,
  size, and font: those keep the ink total the same.
- **Pixel difference** is the catch-all that notices what the other two
  were not looking for. It is the weakest signal of the three: `AE` counts
  differing pixels without weighing how different they are, so it scores a
  layout shift and a colour inversion alike, and it can *rise* when a fix is
  correct but the element is still displaced by an unrelated defect.
- **Diff clusters** localise the pixel axis: the 5%-fuzz mask's contiguous
  regions, each with a bounding box in points. A bare count let #1029's
  squared-off panel corners survive an audit — two ~5,100pt² regions hiding
  inside one number. Every listed cluster must be dispositioned: an accepted
  rendering difference, an evidence-backed GT-exporter difference, or an issue.

Read them together. A fix that improves geometry while leaving the
histogram flat has moved something without changing what is drawn, which is
usually exactly what a positioning fix should do.

Usage:
    compare_render.py GT.pdf OUTPUT.pdf [--page N] [--dpi 150]
        [--fine-shift PT] [--audit]
        [--cluster-report PATH] [--cluster-dispositions PATH] [--strict-clusters]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from spatial_match import minimum_cost_pairs
try:
    from reference_exporter_differences import validate_reference_difference_document
except ModuleNotFoundError:  # Imported as scripts.compare_render in unit tests.
    from scripts.reference_exporter_differences import (
        validate_reference_difference_document,
    )

PAGE_RE = re.compile(r'<page width="[\d.]+" height="[\d.]+">(.*?)</page>', re.S)
WORD_RE = re.compile(
    r'<word xMin="([\d.-]+)" yMin="([\d.-]+)" xMax="([\d.-]+)" yMax="([\d.-]+)">(.*?)</word>',
    re.S,
)
FILL_TEXT_RE = re.compile(
    r'<fill_text[^>]*transform="([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) '
    r'([-0-9.]+) ([-0-9.]+)"[^>]*>(.*?)</fill_text>',
    re.S,
)
# mutool 1.23.x opens a page as `<page mediabox="...">`; later builds add a
# `number` attribute. Splitting on the numbered form alone yields zero pages and
# silently drops the geometry axis.
TRACE_PAGE_RE = re.compile(r"<page\b")
GLYPH_RE = re.compile(r'<g unicode="([^"]*)"(?: glyph="[^"]*")? x="([-0-9.]+)" y="([-0-9.]+)"')
HISTOGRAM_BINS = 32
# Every ImageMagick tool the colour and pixel axes reach for. Named here so the
# availability check and the call sites cannot drift apart.
IMAGEMAGICK_TOOLS = ("convert", "identify", "compare")
# A diff region smaller than this is glyph antialiasing, not a defect: a 20pt²
# cluster is roughly one 4.5pt-square blob. Everything at or above it is
# listed for disposition.
CLUSTER_MIN_AREA_PT2 = 20.0
CLUSTER_REPORT_LIMIT = 12
# One contiguous region this large is never rasterisation noise. #1029's
# squared-off panel corners measured ~5,100pt² each while every axis said the
# pages agreed; this threshold is what turns such a region into a finding.
CLUSTER_DOMINANT_AREA_PT2 = 500.0
CLUSTER_AUDIT_SCHEMA_VERSION = 1
ACCEPTED_RENDERING_CLASSES = frozenset(
    {
        "glyph-edge-rasterization",
        "photo-resampling",
        "gradient-rasterization",
        "shape-edge-antialiasing",
    }
)
COMPONENT_LINE_RE = re.compile(
    r"^\s*\d+:\s+(?P<w>\d+)x(?P<h>\d+)\+(?P<x>\d+)\+(?P<y>\d+)\s+"
    r"(?P<cx>[\d.eE+-]+),(?P<cy>[\d.eE+-]+)\s+(?P<area>[\d.eE+]+)\s+"
    r"(?P<color>\S.*)$"
)


@dataclass(frozen=True)
class TextLine:
    page: int
    x_min: float
    y_min: float
    text: str


@dataclass(frozen=True)
class TextLineMatch:
    reference: TextLine
    candidate: TextLine
    occurrence: int
    occurrences: int

    @property
    def label(self) -> str:
        if self.occurrences == 1:
            return self.reference.text
        return f"{self.reference.text} [{self.occurrence}/{self.occurrences}]"

    @property
    def dx(self) -> float:
        return self.candidate.x_min - self.reference.x_min

    @property
    def dy(self) -> float:
        return self.candidate.y_min - self.reference.y_min


@dataclass(frozen=True)
class DiffCluster:
    """One contiguous region of differing pixels, in page points."""

    x_pt: float
    y_pt: float
    width_pt: float
    height_pt: float
    area_pt2: float
    region: str


def diff_cluster_id(page: int, cluster: DiffCluster) -> str:
    """Return a deterministic ID for one page-local cluster geometry.

    The point values are quantized to 0.01pt so harmless float formatting does
    not change the ID, while a one-pixel movement at the minimum 150 DPI still
    creates a new ID that requires fresh visual review.
    """
    canonical = "|".join(
        (
            str(page),
            f"{cluster.x_pt:.2f}",
            f"{cluster.y_pt:.2f}",
            f"{cluster.width_pt:.2f}",
            f"{cluster.height_pt:.2f}",
            f"{cluster.area_pt2:.2f}",
        )
    )
    digest = hashlib.sha256(canonical.encode("ascii")).hexdigest()[:12]
    return f"p{page}-{digest}"


def _validated_cluster_dispositions(
    cluster_ids: set[str],
    disposition_document: object,
    *,
    page: int,
    reference_difference_document: object = None,
) -> tuple[dict[str, dict[str, str]], list[str], set[str], set[str]]:
    """Validate explicit-ID groups and return dispositions plus schema errors."""
    dispositions: dict[str, dict[str, str]] = {}
    errors: list[str] = []
    duplicate_ids: set[str] = set()
    reference_differences, reference_errors = validate_reference_difference_document(
        reference_difference_document
    ) if reference_difference_document is not None else ({}, [])
    errors.extend(reference_errors)

    if disposition_document is None:
        groups: object = []
    elif not isinstance(disposition_document, dict):
        return {}, ["cluster disposition document must be an object"], set(), set()
    else:
        extra_top_keys = set(disposition_document) - {
            "schema_version",
            "groups",
            "renderer_observations",
        }
        if extra_top_keys:
            errors.append(
                "cluster disposition document has unsupported fields: "
                + ", ".join(sorted(extra_top_keys))
            )
        if disposition_document.get("schema_version") != CLUSTER_AUDIT_SCHEMA_VERSION:
            errors.append(
                f"cluster disposition schema_version must be {CLUSTER_AUDIT_SCHEMA_VERSION}"
            )
        groups = disposition_document.get("groups")

    if not isinstance(groups, list):
        errors.append("cluster disposition groups must be a list")
        groups = []

    for index, group in enumerate(groups, start=1):
        if not isinstance(group, dict):
            errors.append(f"cluster disposition group {index} must be an object")
            continue
        extra_group_keys = set(group) - {
            "kind",
            "class",
            "issue",
            "difference_id",
            "cluster_ids",
            "note",
        }
        if extra_group_keys:
            errors.append(
                f"cluster disposition group {index} has unsupported fields "
                f"{', '.join(sorted(extra_group_keys))}; blanket page, region, or bbox "
                "selectors are not allowed — enumerate cluster_ids"
            )
        raw_ids = group.get("cluster_ids")
        if (
            not isinstance(raw_ids, list)
            or not raw_ids
            or any(not isinstance(cluster_id, str) or not cluster_id for cluster_id in raw_ids)
        ):
            errors.append(
                f"cluster disposition group {index} must enumerate one or more cluster_ids"
            )
            continue

        note = group.get("note")
        if note is not None and (not isinstance(note, str) or not note.strip()):
            errors.append(f"cluster disposition group {index} note must be non-empty text")

        kind = group.get("kind")
        disposition: dict[str, str] | None = None
        if kind == "accepted-rendering":
            accepted_class = group.get("class")
            if accepted_class not in ACCEPTED_RENDERING_CLASSES:
                allowed = ", ".join(sorted(ACCEPTED_RENDERING_CLASSES))
                errors.append(
                    f"cluster disposition group {index} class must be one of: {allowed}"
                )
            elif "issue" in group or "difference_id" in group:
                errors.append(
                    f"cluster disposition group {index} accepted-rendering cannot set "
                    "issue or difference_id"
                )
            else:
                disposition = {"kind": kind, "class": str(accepted_class)}
        elif kind == "issue":
            issue = group.get("issue")
            if not isinstance(issue, str) or re.fullmatch(r"#[1-9]\d*", issue) is None:
                errors.append(
                    f"cluster disposition group {index} issue must be a reference such as #123"
                )
            elif "class" in group or "difference_id" in group:
                errors.append(
                    f"cluster disposition group {index} issue cannot set class or "
                    "difference_id"
                )
            else:
                disposition = {"kind": kind, "issue": issue}
        elif kind == "reference-exporter-difference":
            difference_id = group.get("difference_id")
            difference = reference_differences.get(difference_id)
            if not isinstance(difference_id, str) or difference is None:
                errors.append(
                    f"cluster disposition group {index} difference_id must name a "
                    "validated reference exporter difference"
                )
            elif difference.get("kind") != "render-clusters":
                errors.append(
                    f"cluster disposition group {index} difference_id must name a "
                    "render-clusters difference"
                )
            elif difference.get("page") != page:
                errors.append(
                    f"cluster disposition group {index} reference difference is for "
                    f"page {difference.get('page')}, not page {page}"
                )
            elif set(raw_ids) != set(difference.get("render_cluster_ids", [])):
                errors.append(
                    f"cluster disposition group {index} must use the exact cluster IDs "
                    f"declared by reference difference {difference_id}"
                )
            elif "class" in group or "issue" in group:
                errors.append(
                    f"cluster disposition group {index} reference exporter difference "
                    "cannot set class or issue"
                )
            else:
                disposition = {
                    "kind": kind,
                    "difference_id": difference_id,
                }
        else:
            errors.append(
                f"cluster disposition group {index} kind must be accepted-rendering, "
                "issue, or reference-exporter-difference"
            )

        if disposition is not None and isinstance(note, str):
            disposition["note"] = note.strip()
        for cluster_id in raw_ids:
            if cluster_id in dispositions:
                duplicate_ids.add(cluster_id)
                continue
            if disposition is not None:
                dispositions[cluster_id] = disposition

    if duplicate_ids:
        errors.append(
            "cluster IDs may appear in only one disposition group: "
            + ", ".join(sorted(duplicate_ids))
        )
    unknown_ids = set(dispositions) - cluster_ids
    if unknown_ids:
        errors.append(
            "dispositions reference clusters absent from the current render: "
            + ", ".join(sorted(unknown_ids))
        )
    return dispositions, errors, unknown_ids, duplicate_ids


def _validated_renderer_observations(
    disposition_document: object,
) -> tuple[list[dict[str, object]], list[str]]:
    """Validate bounded observations that never disposition a material cluster."""
    if disposition_document is None:
        return [], []
    if not isinstance(disposition_document, dict):
        return [], []
    raw_observations = disposition_document.get("renderer_observations", [])
    if not isinstance(raw_observations, list):
        return [], ["renderer_observations must be a list"]

    observations: list[dict[str, object]] = []
    errors: list[str] = []
    for index, observation in enumerate(raw_observations, start=1):
        if not isinstance(observation, dict):
            errors.append(f"renderer observation {index} must be an object")
            continue
        extra_keys = set(observation) - {"class", "bbox_pt", "note"}
        if extra_keys:
            errors.append(
                f"renderer observation {index} has unsupported fields: "
                + ", ".join(sorted(extra_keys))
            )
        accepted_class = observation.get("class")
        if accepted_class not in ACCEPTED_RENDERING_CLASSES:
            errors.append(f"renderer observation {index} has an unsupported class")
            continue
        note = observation.get("note")
        if not isinstance(note, str) or not note.strip():
            errors.append(f"renderer observation {index} note must be non-empty text")
            continue
        bbox = observation.get("bbox_pt")
        if not isinstance(bbox, dict) or set(bbox) != {"x", "y", "width", "height"}:
            errors.append(
                f"renderer observation {index} must use an exact x/y/width/height bbox_pt"
            )
            continue
        values = [bbox[name] for name in ("x", "y", "width", "height")]
        if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values):
            errors.append(f"renderer observation {index} bbox_pt values must be numeric")
            continue
        if bbox["x"] < 0 or bbox["y"] < 0 or bbox["width"] <= 0 or bbox["height"] <= 0:
            errors.append(
                f"renderer observation {index} bbox_pt must have non-negative origins "
                "and positive dimensions"
            )
            continue
        observations.append(
            {
                "class": accepted_class,
                "bbox_pt": {
                    name: round(float(bbox[name]), 4)
                    for name in ("x", "y", "width", "height")
                },
                "note": note.strip(),
            }
        )
    return observations, errors


def build_cluster_audit_report(
    clusters: list[DiffCluster] | None,
    *,
    page: int,
    dpi: int,
    disposition_document: object,
    reference_difference_document: object = None,
    strict: bool,
) -> dict[str, object]:
    """Build the complete machine-readable strict cluster audit for one page."""
    census_available = clusters is not None
    current_clusters = clusters or []
    cluster_ids = {diff_cluster_id(page, cluster) for cluster in current_clusters}
    dispositions, errors, unknown_ids, duplicate_ids = _validated_cluster_dispositions(
        cluster_ids,
        disposition_document,
        page=page,
        reference_difference_document=reference_difference_document,
    )
    renderer_observations, observation_errors = _validated_renderer_observations(
        disposition_document
    )
    errors.extend(observation_errors)
    if not census_available:
        errors.append("diff cluster census is unavailable")
    undispositioned_ids = sorted(cluster_ids - set(dispositions))
    records: list[dict[str, object]] = []
    for cluster in current_clusters:
        cluster_id = diff_cluster_id(page, cluster)
        records.append(
            {
                "id": cluster_id,
                "bbox_pt": {
                    "x": round(cluster.x_pt, 4),
                    "y": round(cluster.y_pt, 4),
                    "width": round(cluster.width_pt, 4),
                    "height": round(cluster.height_pt, 4),
                },
                "area_pt2": round(cluster.area_pt2, 4),
                "region": cluster.region,
                "disposition": dispositions.get(cluster_id),
            }
        )
    passed = (
        census_available
        and not errors
        and not undispositioned_ids
        and not unknown_ids
        and not duplicate_ids
    )
    return {
        "schema_version": CLUSTER_AUDIT_SCHEMA_VERSION,
        "page": page,
        "dpi": dpi,
        "fuzz_percent": 5,
        "minimum_area_pt2": CLUSTER_MIN_AREA_PT2,
        "strict": strict,
        "clusters": records,
        "undispositioned_cluster_ids": undispositioned_ids,
        "unknown_disposition_cluster_ids": sorted(unknown_ids),
        "duplicate_disposition_cluster_ids": sorted(duplicate_ids),
        "errors": errors,
        "renderer_observations": renderer_observations,
        "summary": {
            "total": len(records),
            "dispositioned": len(records) - len(undispositioned_ids),
            "undispositioned": len(undispositioned_ids),
            "unknown": len(unknown_ids),
            "duplicate": len(duplicate_ids),
        },
        "passed": passed,
    }


def load_cluster_disposition_document(path: Path | None) -> object:
    if path is None:
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"could not read cluster dispositions from {path}: {exc}") from exc


def load_reference_difference_document(path: Path | None) -> object:
    if path is None:
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(
            f"could not read reference exporter differences from {path}: {exc}"
        ) from exc


def write_cluster_audit_report(path: Path, report: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def report_cluster_audit(report: dict[str, object], path: Path) -> None:
    summary = report["summary"]
    assert isinstance(summary, dict)
    print("## Strict diff-cluster audit")
    print(f"  machine report: {path}")
    print(
        f"  {summary['dispositioned']}/{summary['total']} current cluster(s) dispositioned"
    )
    for error in report["errors"]:
        print(f"  ERROR: {error}")
    for cluster_id in report["undispositioned_cluster_ids"]:
        print(f"  UNDISPOSITIONED: {cluster_id}")
    print("  PASS" if report["passed"] else "  FAIL")


def render_page(pdf: Path, page: int, dpi: int, out_dir: Path, role: str) -> Path:
    """Rasterise one page, returning the PNG path.

    `role` names the output, because the GT and the candidate usually share
    a file stem: rendering both under that stem made the second overwrite
    the first, and every comparison then ran an image against itself and
    reported a perfect match.
    """
    prefix = out_dir / role
    subprocess.run(
        ["pdftoppm", "-r", str(dpi), "-png", "-f", str(page), "-l", str(page),
         str(pdf), str(prefix)],
        check=True,
        capture_output=True,
    )
    pages = sorted(out_dir.glob(f"{role}-*.png"))
    if not pages:
        raise SystemExit(f"{pdf}: page {page} did not render")
    return pages[0]


def has_mutool() -> bool:
    """Whether `mutool` is on PATH (it ships in `mupdf-tools`)."""
    return shutil.which("mutool") is not None


def imagemagick_command(tool: str) -> list[str] | None:
    """Argv prefix invoking an ImageMagick tool, or None if it is unavailable.

    ImageMagick 7 dispatches every tool through a single `magick` binary;
    ImageMagick 6 installs them under their own names and has no `magick` at
    all. Hardcoding the IM7 spelling left the colour and pixel axes dying with
    FileNotFoundError on an IM6 host, after the geometry axis had already
    printed a report that looked whole.
    """
    if shutil.which("magick") is not None:
        # `magick convert` is deprecated in 7.1 and warns; plain conversion is
        # the bare dispatcher, so only the named subtools take an argument.
        return ["magick"] if tool == "convert" else ["magick", tool]
    if shutil.which(tool) is not None:
        return [tool]
    return None


def has_imagemagick() -> bool:
    """Whether every tool the colour and pixel axes need can be invoked."""
    return all(imagemagick_command(tool) is not None for tool in IMAGEMAGICK_TOOLS)


def require_vision_artifact_dependencies(artifacts_dir: Path | None) -> None:
    """Fail rather than silently omit requested model-vision evidence."""
    if artifacts_dir is not None and not has_imagemagick():
        raise SystemExit(
            "--artifacts-dir requires ImageMagick to preserve full pages, the "
            "pixel diff, and matched crops; install imagemagick and rerun"
        )


def baseline_lines(pdf: Path) -> list[TextLine]:
    """Text-line anchors from `mutool draw -F trace` affine coordinates.

    A `<fill_text>` carries the complete text-space matrix, so each glyph maps
    to `x = a * gx + c * gy + tx` and `y = b * gx + d * gy + ty`. Office
    exports on macOS often use scaled or rotated text spaces while ours are
    commonly plain translations; the same affine formula covers both.

    For non-rotated text, baselines jitter by fractions of a point inside one
    visual line, so rows are bucketed to 1pt before being joined. A rotated or
    skewed `<fill_text>` is already one visual run and cannot share a horizontal
    baseline bucket; it stays intact and uses the minimum transformed glyph
    coordinates as its comparable spatial anchor.
    """
    trace = subprocess.run(
        ["mutool", "draw", "-F", "trace", "-o", "-", str(pdf)],
        capture_output=True,
        text=True,
    )
    if trace.returncode != 0:
        return []
    lines: list[TextLine] = []
    for page_index, page in enumerate(TRACE_PAGE_RE.split(trace.stdout)[1:]):
        rows: dict[int, list[tuple[float, float, str]]] = {}
        rotated_lines: list[TextLine] = []
        for match in FILL_TEXT_RE.finditer(page):
            scale_x = float(match.group(1))
            shear_y = float(match.group(2))
            shear_x = float(match.group(3))
            scale_y = float(match.group(4))
            translate_x, translate_y = float(match.group(5)), float(match.group(6))
            transformed_glyphs: list[tuple[float, float, str]] = []
            for glyph in GLYPH_RE.finditer(match.group(7)):
                char = glyph.group(1)
                if not char.strip():
                    continue
                glyph_x, glyph_y = float(glyph.group(2)), float(glyph.group(3))
                x = translate_x + scale_x * glyph_x + shear_x * glyph_y
                y = translate_y + shear_y * glyph_x + scale_y * glyph_y
                transformed_glyphs.append((x, y, char))
            if abs(shear_x) > 1e-9 or abs(shear_y) > 1e-9:
                text = re.sub(
                    r"\s+", " ", "".join(glyph[2] for glyph in transformed_glyphs)
                ).strip()
                if text:
                    rotated_lines.append(
                        TextLine(
                            page_index,
                            min(glyph[0] for glyph in transformed_glyphs),
                            min(glyph[1] for glyph in transformed_glyphs),
                            text,
                        )
                    )
            else:
                for x, baseline, char in transformed_glyphs:
                    rows.setdefault(round(baseline), []).append((x, baseline, char))
        for key in sorted(rows):
            # Same-origin ligature continuations retain their trace text order.
            glyphs = sorted(rows[key], key=lambda glyph: glyph[:2])
            text = re.sub(r"\s+", " ", "".join(glyph[2] for glyph in glyphs)).strip()
            if text:
                # The 1pt bucket only groups glyphs into a line; reporting its
                # key as the position would quantise every measurement to 1pt
                # and hide the sub-point row-pitch differences this axis exists
                # to find. Report the line's own baseline instead.
                lines.append(
                    TextLine(
                        page_index,
                        min(g[0] for g in glyphs),
                        min(g[1] for g in glyphs),
                        text,
                    )
                )
        lines.extend(rotated_lines)
    return lines


def descriptor_box_lines(pdf: Path) -> list[TextLine]:
    """Fallback line tops from `pdftotext -bbox`, used only without `mutool`.

    `yMin` is each glyph's *font-descriptor box*, not its ink or its baseline.
    The two PDFs always embed different subsets, so the drift this yields
    carries an error proportional to font size — on the newsletter mock it
    reported +2.90pt for a 22pt heading whose baseline is really 1.07pt the
    other way, which is how #501 came to be filed against a defect that did
    not exist (issue #505).
    """
    xml = subprocess.run(
        ["pdftotext", "-bbox", str(pdf), "-"],
        capture_output=True,
        text=True,
        check=True,
    ).stdout
    lines: list[TextLine] = []
    for page_index, page in enumerate(PAGE_RE.findall(xml)):
        rows: dict[int, list[tuple[float, float, str]]] = {}
        for match in WORD_RE.finditer(page):
            x_min, y_min = float(match.group(1)), float(match.group(2))
            rows.setdefault(round(y_min), []).append((x_min, y_min, match.group(5)))
        for key in sorted(rows):
            words = sorted(rows[key])
            text = re.sub(r"\s+", " ", " ".join(word[2] for word in words)).strip()
            if text:
                lines.append(
                    TextLine(page_index, min(w[0] for w in words),
                             min(w[1] for w in words), text)
                )
    return lines


def text_lines(pdf: Path) -> list[TextLine]:
    """Lines for the geometry axis, by true baseline where `mutool` allows."""
    if has_mutool():
        lines = baseline_lines(pdf)
        if lines:
            return lines
    return descriptor_box_lines(pdf)


def match_text_line_instances(
    gt_lines: list[TextLine], other_lines: list[TextLine]
) -> list[TextLineMatch]:
    """Match equal text per page, including every repeated occurrence."""
    gt_groups: dict[tuple[int, str], list[TextLine]] = defaultdict(list)
    other_groups: dict[tuple[int, str], list[TextLine]] = defaultdict(list)
    for line in gt_lines:
        gt_groups[(line.page, line.text)].append(line)
    for line in other_lines:
        other_groups[(line.page, line.text)].append(line)

    matches: list[TextLineMatch] = []
    for key, references in gt_groups.items():
        candidates = other_groups.get(key, [])
        references = sorted(references, key=lambda line: (line.y_min, line.x_min))
        candidates = sorted(candidates, key=lambda line: (line.y_min, line.x_min))
        for reference_index, candidate_index in minimum_cost_pairs(
            [(line.x_min, line.y_min) for line in references],
            [(line.x_min, line.y_min) for line in candidates],
        ):
            matches.append(
                TextLineMatch(
                    reference=references[reference_index],
                    candidate=candidates[candidate_index],
                    occurrence=reference_index + 1,
                    occurrences=len(references),
                )
            )
    return sorted(
        matches,
        key=lambda match: (
            match.reference.page,
            match.reference.y_min,
            match.reference.x_min,
        ),
    )


def page_text_lines(pdf: Path, page: int) -> list[TextLine]:
    """Return only the requested 1-based page's text instances."""
    return [line for line in text_lines(pdf) if line.page == page - 1]


def report_geometry(
    gt: Path,
    other: Path,
    page: int = 1,
    large_shift: float = 5.0,
    fine_shift: float | None = None,
) -> dict[str, float]:
    """Vertical and horizontal drift of spatially matched text instances."""
    gt_text_lines = page_text_lines(gt, page)
    other_text_lines = page_text_lines(other, page)
    matches = match_text_line_instances(gt_text_lines, other_text_lines)
    dy = [match.dy for match in matches]
    dx = [match.dx for match in matches]
    other_text_pages: dict[str, set[int]] = defaultdict(set)
    for line in other_text_lines:
        other_text_pages[line.text].add(line.page)
    matched_reference_ids = {id(match.reference) for match in matches}
    page_mismatch = sum(
        1
        for line in gt_text_lines
        if id(line) not in matched_reference_ids
        and line.text in other_text_pages
        and line.page not in other_text_pages[line.text]
    )

    print("## Geometry — position, size, pitch")
    if not has_mutool():
        print("  APPROXIMATE: mutool absent, so positions come from font-descriptor")
        print("  boxes rather than baselines. The error scales with font size and can")
        print("  invert the sign. Install mupdf-tools before trusting these numbers.")
    if not dy:
        print("  no text instances matched; compare pages manually")
        return {}
    mad_y = sum(abs(value) for value in dy) / len(dy)
    mad_x = sum(abs(value) for value in dx) / len(dx)
    coverage = len(dy) / len(gt_text_lines) if gt_text_lines else 0.0
    worst_dy = max(matches, key=lambda match: abs(match.dy))
    worst_dx = max(matches, key=lambda match: abs(match.dx))
    large_matches = [
        match for match in matches if abs(match.dx) > large_shift or abs(match.dy) > large_shift
    ]
    fine_matches = (
        [
            match
            for match in matches
            if abs(match.dx) > fine_shift or abs(match.dy) > fine_shift
        ]
        if fine_shift is not None
        else []
    )
    print(f"  matched instances  {len(dy)} of {len(gt_text_lines)} "
          f"({coverage * 100:.0f}% of the GT's text lines)")
    print(
        f"  vertical   MAD {mad_y:7.2f}pt   worst {worst_dy.dy:+8.2f}pt  "
        f"{worst_dy.label[:60]}"
    )
    print(
        f"  horizontal MAD {mad_x:7.2f}pt   worst {worst_dx.dx:+8.2f}pt  "
        f"{worst_dx.label[:60]}"
    )
    print(f"  large instance shifts (>{large_shift:.2f}pt): {len(large_matches)}")
    for match in sorted(
        large_matches, key=lambda item: max(abs(item.dx), abs(item.dy)), reverse=True
    ):
        print(
            f"    page {match.reference.page + 1}: {match.label[:52]}  "
            f"dx {match.dx:+.2f}pt  dy {match.dy:+.2f}pt"
        )
    if fine_shift is not None:
        print(f"  fine-detail instance shifts (>{fine_shift:.2f}pt): {len(fine_matches)}")
        for match in sorted(
            fine_matches, key=lambda item: max(abs(item.dx), abs(item.dy)), reverse=True
        ):
            print(
                f"    page {match.reference.page + 1}: {match.label[:52]}  "
                f"dx {match.dx:+.2f}pt  dy {match.dy:+.2f}pt"
            )
    if page_mismatch:
        print(f"  on a different page: {page_mismatch} line(s) — pagination differs")
    result = {
        "mad_y": mad_y,
        "mad_x": mad_x,
        "page_mismatch": float(page_mismatch),
        "matched": float(len(dy)),
        "coverage": coverage,
        "worst_dx": worst_dx.dx,
        "worst_dy": worst_dy.dy,
        "large_shift_count": float(len(large_matches)),
        "large_shift_threshold": large_shift,
    }
    if fine_shift is not None:
        result.update(
            {
                "fine_shift_count": float(len(fine_matches)),
                "fine_shift_threshold": fine_shift,
            }
        )
    return result


def histogram(png: Path) -> tuple[list[int], int]:
    """Per-channel binned colour counts, flattened R|G|B, and the pixel total."""
    txt = subprocess.run(
        [*(imagemagick_command("convert") or []), str(png), "-depth", "8", "txt:-"],
        capture_output=True,
        text=True,
        check=True,
    ).stdout
    counts = [0] * (3 * HISTOGRAM_BINS)
    total = 0
    step = 256 // HISTOGRAM_BINS
    for line in txt.splitlines():
        head, _, rest = line.partition("#")
        if not head or len(rest) < 6:
            continue
        try:
            channels = [int(rest[i : i + 2], 16) for i in (0, 2, 4)]
        except ValueError:
            continue
        for index, value in enumerate(channels):
            counts[index * HISTOGRAM_BINS + min(value // step, HISTOGRAM_BINS - 1)] += 1
        total += 1
    return counts, total


def ink_fraction(counts: list[int], total: int) -> float:
    """Share of pixels that are not near-white, averaged over the channels."""
    if total == 0:
        return 0.0
    dark = sum(
        counts[channel * HISTOGRAM_BINS + b]
        for channel in range(3)
        for b in range(HISTOGRAM_BINS - 2)
    )
    return dark / (3.0 * total)


def report_histogram(gt_png: Path, other_png: Path) -> dict[str, float]:
    """Colour-distribution agreement, independent of where the pixels sit."""
    gt_counts, gt_total = histogram(gt_png)
    counts, total = histogram(other_png)
    gt_sum = sum(gt_counts) or 1
    other_sum = sum(counts) or 1
    reference = [value / gt_sum for value in gt_counts]
    candidate = [value / other_sum for value in counts]
    intersection = sum(min(a, b) for a, b in zip(reference, candidate))
    # Bin-wise agreement punishes a one-level shift as hard as a recolour: a
    # smooth gradient dithers by a channel step or two between renderers, and
    # every one of those pixels lands in a neighbouring bin. Three decks
    # scored 0.9745-0.9860 on intersection with their gradients pixel-identical
    # to within +-2 per channel. Comparing the *cumulative* distributions
    # instead measures how far colour has to move, so a one-bin shift costs
    # almost nothing while a genuine recolour still shows.
    shift = cumulative_distance(reference, candidate)
    gt_ink = ink_fraction(gt_counts, gt_total)
    ink = ink_fraction(counts, total)

    print("## Histogram — fill colour, recolouring, missing elements")
    print(f"  intersection       {intersection:.4f}   (1.0000 = identical distribution)")
    print(f"  colour shift       {shift:.4f}   (0.0000 = identical; tolerates dithering)")
    print(f"  ink coverage       {ink * 100:6.3f}%  against GT {gt_ink * 100:6.3f}%"
          f"   ({(ink - gt_ink) * 100:+.3f}%)")
    return {
        "intersection": intersection,
        "shift": shift,
        "ink_delta": (ink - gt_ink) * 100.0,
    }


def cumulative_distance(reference: list[float], candidate: list[float]) -> float:
    """Mean per-channel distance between the cumulative distributions.

    Insensitive to a colour landing one bin either side of where it did in
    the reference, which is what renderer dithering produces, while still
    growing with a real change in what colour is present.
    """
    channels = 3
    total = 0.0
    for channel in range(channels):
        start = channel * HISTOGRAM_BINS
        run_reference = 0.0
        run_candidate = 0.0
        for index in range(start, start + HISTOGRAM_BINS):
            run_reference += reference[index]
            run_candidate += candidate[index]
            total += abs(run_reference - run_candidate)
    return total / (channels * HISTOGRAM_BINS)


def report_pixels(gt_png: Path, other_png: Path, out_dir: Path) -> Path:
    """Whole-page difference, as a coarse catch-all.

    Returns the size-normalised GT so the cluster census compares the same
    pair of images the counts above were measured on.
    """
    normalised = out_dir / "gt-normalised.png"
    size = subprocess.run(
        [*(imagemagick_command("identify") or []), "-format", "%wx%h", str(other_png)],
        capture_output=True, text=True, check=True,
    ).stdout.strip()
    # pdftoppm can differ by a pixel between two PDFs of the same paper size,
    # which would otherwise make every comparison fail outright.
    subprocess.run(
        [*(imagemagick_command("convert") or []), str(gt_png),
         "-background", "white", "-extent", size, str(normalised)],
        check=True, capture_output=True,
    )

    print("## Pixel difference — coarse catch-all, read last")
    for label, args in (
        ("AE  5% fuzz", ["-metric", "AE", "-fuzz", "5%"]),
        ("AE  1% fuzz", ["-metric", "AE", "-fuzz", "1%"]),
        ("RMSE       ", ["-metric", "RMSE"]),
    ):
        result = subprocess.run(
            [*(imagemagick_command("compare") or []), *args,
             str(normalised), str(other_png), "null:"],
            capture_output=True, text=True,
        )
        print(f"  {label}      {result.stderr.strip()}")
    return normalised


def component_is_white(color: str) -> bool:
    """Whether a connected-component mean colour reads as mask-white.

    The mask is binary, but `area-threshold` merges sub-threshold specks into
    their neighbour, leaving near-pure means such as `gray(254.7)`; percentages
    appear when a build reports srgba. Thresholding at half intensity accepts
    both without enumerating ImageMagick's colour spellings.
    """
    value = re.search(r"([\d.]+)", color)
    if value is None:
        return False
    return float(value.group(1)) > (50.0 if "%" in color else 127.5)


def page_region(cx: float, cy: float, width: float, height: float) -> str:
    """Name the page third a centroid falls in, e.g. `bottom-right`."""
    column = "left" if cx < width / 3 else "right" if cx > 2 * width / 3 else "center"
    row = "top" if cy < height / 3 else "bottom" if cy > 2 * height / 3 else "middle"
    if row == "middle":
        return "center" if column == "center" else column
    return row if column == "center" else f"{row}-{column}"


def parse_diff_clusters(
    verbose: str, dpi: int, min_area_pt2: float = CLUSTER_MIN_AREA_PT2
) -> list[DiffCluster]:
    """White connected components of the diff mask, largest first, in points.

    The page extent comes from the union of every component's bounding box:
    the black background component normally spans the full page, and when it
    does not, the union still bounds everything that was compared.
    """
    scale = 72.0 / dpi
    components: list[tuple[float, float, float, float, float, float, float]] = []
    page_width = page_height = 0.0
    for line in verbose.splitlines():
        match = COMPONENT_LINE_RE.match(line)
        if match is None:
            continue
        x, y = float(match.group("x")), float(match.group("y"))
        w, h = float(match.group("w")), float(match.group("h"))
        page_width = max(page_width, x + w)
        page_height = max(page_height, y + h)
        if not component_is_white(match.group("color")):
            continue
        components.append(
            (x, y, w, h, float(match.group("cx")), float(match.group("cy")),
             float(match.group("area")))
        )
    clusters = [
        DiffCluster(
            x_pt=x * scale,
            y_pt=y * scale,
            width_pt=w * scale,
            height_pt=h * scale,
            area_pt2=area * scale * scale,
            region=page_region(cx, cy, page_width, page_height),
        )
        for x, y, w, h, cx, cy, area in components
        if area * scale * scale >= min_area_pt2
    ]
    clusters.sort(key=lambda cluster: cluster.area_pt2, reverse=True)
    return clusters


def diff_cluster_census(
    gt_png: Path, other_png: Path, out_dir: Path, dpi: int
) -> list[DiffCluster] | None:
    """Label the 5%-fuzz diff mask's contiguous regions, or None if unsupported.

    A single-pixel morphological open drops isolated antialiasing specks first,
    and 4-connectivity keeps diagonally-touching glyph noise from chaining into
    one page-spanning blob the way 8-connectivity does. `area-threshold` merges
    the sub-25px remnants so the verbose listing stays bounded.
    """
    mask = out_dir / "diff-mask.png"
    # `compare` exits 1 whenever the images differ; only a missing mask is
    # a failure.
    subprocess.run(
        [*(imagemagick_command("compare") or []), "-metric", "AE", "-fuzz", "5%",
         "-compose", "src", "-highlight-color", "white", "-lowlight-color",
         "black", str(gt_png), str(other_png), str(mask)],
        capture_output=True,
    )
    if not mask.is_file():
        return None
    census = subprocess.run(
        [*(imagemagick_command("convert") or []), str(mask),
         "-morphology", "Open", "Square:1",
         "-define", "connected-components:verbose=true",
         "-define", "connected-components:area-threshold=25",
         "-connected-components", "4", "null:"],
        capture_output=True,
        text=True,
    )
    if census.returncode != 0:
        return None
    # The objects listing has moved between stdout and stderr across
    # ImageMagick releases; parse both.
    return parse_diff_clusters(census.stdout + census.stderr, dpi)


def report_diff_clusters(clusters: list[DiffCluster] | None, page: int = 1) -> None:
    """List every contiguous diff region large enough to need a disposition."""
    print("## Diff clusters — where the differing pixels sit")
    if clusters is None:
        print("  SKIPPED: this ImageMagick build lacks -connected-components;")
        print("  inspect the diff image by eye and disposition each region.")
        return
    if not clusters:
        print(f"  none of {CLUSTER_MIN_AREA_PT2:.0f}pt² or more — what differs is")
        print("  dispersed specks (glyph rasterisation and antialiasing).")
        return
    print("  Disposition every cluster: an accepted rendering difference (glyph")
    print("  rasterisation), verified GT-exporter evidence, or an issue reference.")
    print("  A cluster hugging a shape's")
    print("  corner or edge while colour and position agree is outline geometry —")
    print("  compare the shape's path, not its fill or its box (#1029).")
    for index, cluster in enumerate(clusters[:CLUSTER_REPORT_LIMIT], start=1):
        print(f"  {index:>2}. {diff_cluster_id(page, cluster)}  "
              f"{cluster.width_pt:6.1f} x {cluster.height_pt:6.1f}pt "
              f"at ({cluster.x_pt:6.1f}, {cluster.y_pt:6.1f})  "
              f"area {cluster.area_pt2:7.0f}pt2  {cluster.region}")
    if len(clusters) > CLUSTER_REPORT_LIMIT:
        print(f"  … and {len(clusters) - CLUSTER_REPORT_LIMIT} more of "
              f"{CLUSTER_MIN_AREA_PT2:.0f}pt2 or more")


def shift_crop_box(
    match: TextLineMatch, dpi: int, image_width: int, image_height: int
) -> tuple[int, int, int, int]:
    """Matched page-space crop containing both locations of one shifted line."""
    scale = dpi / 72.0
    text_extent_pt = max(72.0, min(240.0, len(match.reference.text) * 8.0))
    left = max(0, round((min(match.reference.x_min, match.candidate.x_min) - 24.0) * scale))
    right = min(
        image_width,
        round(
            (max(match.reference.x_min, match.candidate.x_min) + text_extent_pt + 24.0)
            * scale
        ),
    )
    top = max(0, round((min(match.reference.y_min, match.candidate.y_min) - 32.0) * scale))
    bottom = min(
        image_height,
        round((max(match.reference.y_min, match.candidate.y_min) + 24.0) * scale),
    )
    return left, top, max(1, right - left), max(1, bottom - top)


def preserve_vision_artifacts(
    gt_png: Path,
    other_png: Path,
    artifacts_dir: Path,
    page: int,
    dpi: int,
    large_matches: list[TextLineMatch],
) -> list[Path]:
    """Persist full pages, a pixel diff, and matched crops for model vision."""
    artifacts_dir.mkdir(parents=True, exist_ok=True)
    gt_artifact = artifacts_dir / f"page-{page}-gt.png"
    output_artifact = artifacts_dir / f"page-{page}-output.png"
    side_by_side = artifacts_dir / f"page-{page}-side-by-side.png"
    diff_artifact = artifacts_dir / f"page-{page}-diff-5pct.png"

    size = subprocess.run(
        [*(imagemagick_command("identify") or []), "-format", "%wx%h", str(other_png)],
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()
    width, height = (int(value) for value in size.split("x", maxsplit=1))
    subprocess.run(
        [
            *(imagemagick_command("convert") or []),
            str(gt_png),
            "-background",
            "white",
            "-extent",
            size,
            str(gt_artifact),
        ],
        check=True,
        capture_output=True,
    )
    shutil.copy2(other_png, output_artifact)
    subprocess.run(
        [
            *(imagemagick_command("convert") or []),
            str(gt_artifact),
            str(output_artifact),
            "+append",
            str(side_by_side),
        ],
        check=True,
        capture_output=True,
    )
    diff = subprocess.run(
        [
            *(imagemagick_command("compare") or []),
            "-metric",
            "AE",
            "-fuzz",
            "5%",
            str(gt_artifact),
            str(output_artifact),
            str(diff_artifact),
        ],
        capture_output=True,
        text=True,
    )
    if diff.returncode not in (0, 1):
        raise SystemExit(f"ImageMagick failed to create {diff_artifact}: {diff.stderr.strip()}")

    paths = [gt_artifact, output_artifact, side_by_side, diff_artifact]
    for index, match in enumerate(large_matches, start=1):
        slug = re.sub(r"[^a-z0-9]+", "-", match.label.lower()).strip("-") or "text"
        crop_path = artifacts_dir / f"page-{page}-shift-{index:02d}-{slug[:48]}.png"
        left, top, crop_width, crop_height = shift_crop_box(match, dpi, width, height)
        crop = f"{crop_width}x{crop_height}+{left}+{top}"
        gt_crop = artifacts_dir / f".gt-crop-{index:02d}.png"
        output_crop = artifacts_dir / f".output-crop-{index:02d}.png"
        for source, destination in (
            (gt_artifact, gt_crop),
            (output_artifact, output_crop),
        ):
            subprocess.run(
                [
                    *(imagemagick_command("convert") or []),
                    str(source),
                    "-crop",
                    crop,
                    "+repage",
                    str(destination),
                ],
                check=True,
                capture_output=True,
            )
        subprocess.run(
            [
                *(imagemagick_command("convert") or []),
                str(gt_crop),
                str(output_crop),
                "+append",
                str(crop_path),
            ],
            check=True,
            capture_output=True,
        )
        gt_crop.unlink()
        output_crop.unlink()
        paths.append(crop_path)

    print("## Vision artifacts — open every image with Codex/Claude vision")
    print("  Numeric output does not complete the visual audit.")
    for path in paths:
        print(f"  {path}")
    return paths


def diagnose(
    geometry: dict[str, float],
    histogram_result: dict[str, float] | None,
    clusters: list[DiffCluster] | None = None,
) -> None:
    """Say what the combination of axes means, and what to look at next.

    This is the point of running all three: a single number invites the
    wrong conclusion. A pixel count that rises can accompany a correct fix,
    and one that does not move at all can hide a large geometric
    improvement. The pattern across axes is what identifies the defect
    class.

    `histogram_result` is None when ImageMagick is absent. An axis that did not
    run must never read as an axis that agreed, so silence there is reported
    rather than folded into the verdict.
    """
    colour_measured: bool = histogram_result is not None
    histogram_result = histogram_result or {}
    print("## Reading")
    if not geometry:
        print("  Geometry could not be measured, so the other axes stand alone.")
        print("  Compare the pages by eye before trusting them.")
        return

    # A drift figure averaged over a handful of lines is noise. Korean
    # sheets match poorly because word segmentation differs between the two
    # PDFs, and one such page reported 4.11pt of "drift" from nine matched
    # lines out of sixty — a number with no meaning that would have sent the
    # next investigation chasing a row-height bug that is not there.
    matched: float = geometry.get("matched", 0.0)
    coverage: float = geometry.get("coverage", 0.0)
    if matched < 10 or coverage < 0.25:
        print(f"  Only {matched:.0f} lines matched ({coverage * 100:.0f}% of the GT).")
        print("  Treat the geometry figures as unreliable: too few samples, and")
        print("  the ones that matched are not a random selection. Compare the")
        print("  pages by eye, or measure a specific element directly, before")
        print("  drawing any conclusion from the drift above.")
        print()

    mad_y: float = geometry["mad_y"]
    mad_x: float = geometry["mad_x"]
    pages_differ: bool = geometry["page_mismatch"] > 0
    intersection: float = histogram_result.get("intersection", 1.0)
    ink_delta: float = histogram_result.get("ink_delta", 0.0)
    large_shift_count = int(geometry.get("large_shift_count", 0.0))
    large_shift_threshold = geometry.get("large_shift_threshold", 5.0)
    fine_shift_count = int(geometry.get("fine_shift_count", 0.0))
    fine_shift_threshold = geometry.get("fine_shift_threshold")

    # Thresholds are deliberately loose: they route attention, they do not
    # decide correctness. A point of drift is invisible; ten is not.
    drifts_vertically: bool = mad_y > 2.0
    drifts_horizontally: bool = mad_x > 1.0
    # Judge colour on the shift, not the bin-wise intersection: the latter
    # flags smooth gradients that are pixel-identical to within dithering.
    # Measured separation on this corpus: renderer dithering across a smooth
    # gradient reaches 0.0003, and the half-width cell borders of #487
    # reached 0.0016 before the fix and 0.0004 after it.
    colour_differs: bool = colour_measured and histogram_result.get("shift", 0.0) > 0.001
    ink_differs: bool = colour_measured and abs(ink_delta) > 0.2

    findings: list[str] = []
    if fine_shift_count and fine_shift_threshold is not None:
        findings.append(
            f"{fine_shift_count} matched text instance(s) move more than the "
            f"fine-detail {fine_shift_threshold:.2f}pt tolerance. Trace-derived "
            "anchor movement is geometry, not font antialiasing; inspect and track "
            "each named instance above."
        )
    if large_shift_count:
        findings.append(
            f"{large_shift_count} matched text instance(s) move more than "
            f"{large_shift_threshold:.2f}pt. These named element-level differences "
            "remain valid even when the page has too few lines for aggregate MAD "
            "to be representative; inspect and track each one above."
        )
    if not colour_measured:
        findings.append(
            "The colour and pixel axes did not run — ImageMagick is absent, so "
            "a wrong fill, a recolour, or a missing element cannot be seen "
            "here at all. Geometry below stands alone; compare the pages by "
            "eye before concluding they agree."
        )
    if pages_differ:
        findings.append(
            "Pagination differs — content sits on the wrong page. Fix this "
            "first: every per-line measurement below it is contaminated by "
            "the accumulated drift that pushed it over."
        )
    if drifts_vertically:
        findings.append(
            f"Vertical drift {mad_y:.2f}pt — line advance, paragraph spacing, "
            "or row height. Compare consecutive-line pitch against the GT "
            "rather than absolute positions, so a constant offset near the "
            "top does not read as a spacing bug."
        )
    if drifts_horizontally:
        findings.append(
            f"Horizontal drift {mad_x:.2f}pt — indent, column width, or "
            "margin. If it grows across the page it is per-column and "
            "cumulative; if it is constant it is an indent or margin."
        )
    if colour_differs:
        findings.append(
            f"Colour distribution differs (shift {histogram_result.get('shift', 0.0):.4f}) — "
            "a fill, theme colour, or shading is wrong, or an element is "
            "missing entirely. Position measurements will not show this."
        )
    if ink_differs and not colour_differs:
        findings.append(
            f"Ink coverage is off by {ink_delta:+.3f}% while the colour "
            "distribution matches — the right things are drawn in the right "
            "colours but at the wrong size, or a font renders at a different "
            "weight."
        )
    dominant = [
        cluster
        for cluster in (clusters or [])
        if cluster.area_pt2 >= CLUSTER_DOMINANT_AREA_PT2
    ]
    if dominant:
        largest = dominant[0]
        findings.append(
            f"A contiguous diff cluster of {largest.area_pt2:.0f}pt2 sits at "
            f"({largest.x_pt:.0f}, {largest.y_pt:.0f}) ({largest.region})"
            + (f", plus {len(dominant) - 1} more of {CLUSTER_DOMINANT_AREA_PT2:.0f}pt2 or larger"
               if len(dominant) > 1 else "")
            + " — a structural difference: a missing element, a displaced "
            "block, or a shape outline (rounded corner, curved edge) drawn as "
            "its bounding box. The per-line geometry above only measures text, "
            "so it cannot see this; disposition the cluster before concluding "
            "the pages agree."
        )

    if not findings:
        print("  No axis shows a material difference. What remains is font")
        print("  rasterisation and antialiasing; inspect crops at full")
        print("  resolution before concluding anything is wrong.")
        return
    for index, finding in enumerate(findings, start=1):
        print(f"  {index}. {finding}")

    if (
        colour_measured
        and not colour_differs
        and not ink_differs
        and (drifts_vertically or drifts_horizontally)
    ):
        print()
        print("  Colour and ink are unchanged while geometry moves: this is a")
        print("  pure layout difference. A pixel count may rise even as the")
        print("  fix is correct, because a displaced element that grows")
        print("  toward its true size overlaps GT less, not more.")


def report_matched_lines(gt: Path, other: Path, page: int = 1) -> None:
    """Per-instance positions for every text line matched spatially.

    Aggregate drift says a page is wrong; this says which line. Pairing the two
    PDFs by hand — taking the topmost line, or grepping for a prefix — silently
    matches the wrong line and produces impossible numbers, so the pairing here
    is the same duplicate-safe spatial match the geometry axis already trusts.
    """
    rows = match_text_line_instances(page_text_lines(gt, page), page_text_lines(other, page))

    print("## Matched lines — x/y position of each spatial text instance")
    if not rows:
        print("  no text instances matched")
        return
    print(f"  {'page':>4} {'GT x':>8} {'out x':>8} {'dx':>8} "
          f"{'GT y':>8} {'out y':>8} {'dy':>8}  text instance")
    for match in rows:
        reference = match.reference
        candidate = match.candidate
        print(
            f"  {reference.page + 1:>4} {reference.x_min:8.2f} {candidate.x_min:8.2f} "
            f"{match.dx:+8.2f} {reference.y_min:8.2f} {candidate.y_min:8.2f} "
            f"{match.dy:+8.2f}  {match.label[:52]}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("gt", type=Path, help="native Office export")
    parser.add_argument("output", type=Path, help="office2pdf output")
    parser.add_argument("--page", type=int, default=1)
    parser.add_argument("--dpi", type=int, default=150, help="at least 150")
    parser.add_argument(
        "--large-shift",
        type=float,
        default=5.0,
        metavar="PT",
        help="flag any matched text instance whose x or y moves by more than this (default: 5pt)",
    )
    parser.add_argument(
        "--fine-shift",
        type=float,
        metavar="PT",
        help=(
            "enable the fine-detail audit gate and flag matched text whose x or y "
            "moves by more than PT; does not replace the 5pt coarse summary"
        ),
    )
    parser.add_argument(
        "--audit",
        action="store_true",
        help="exit nonzero when the enabled per-instance text-shift gate finds movement",
    )
    parser.add_argument(
        "--artifacts-dir",
        type=Path,
        help="preserve full GT/output pages, 5%% diff, and gated-shift crops for model vision",
    )
    parser.add_argument(
        "--cluster-report",
        type=Path,
        help="write the complete machine-readable diff-cluster audit for this page",
    )
    parser.add_argument(
        "--cluster-dispositions",
        type=Path,
        help=(
            "JSON groups that explicitly map current cluster IDs to an accepted "
            "renderer class, verified GT-exporter evidence, or an open issue"
        ),
    )
    parser.add_argument(
        "--reference-differences",
        type=Path,
        help=(
            "evidence-backed reference exporter differences used by exact cluster "
            "dispositions"
        ),
    )
    parser.add_argument(
        "--strict-clusters",
        action="store_true",
        help="exit nonzero unless every material diff cluster has a valid explicit disposition",
    )
    parser.add_argument(
        "--lines",
        action="store_true",
        help="list every matched text instance's x/y position without hand-pairing "
        "repeated labels between the two PDFs",
    )
    args = parser.parse_args()

    if args.dpi < 150:
        raise SystemExit("--dpi must be at least 150; hairlines vanish below that")
    if args.fine_shift is not None and args.fine_shift <= 0:
        parser.error("--fine-shift must be greater than zero")
    if args.strict_clusters and args.cluster_report is None:
        parser.error("--strict-clusters requires --cluster-report")
    if args.cluster_dispositions is not None and args.cluster_report is None:
        parser.error("--cluster-dispositions requires --cluster-report")
    if args.reference_differences is not None and args.cluster_dispositions is None:
        parser.error("--reference-differences requires --cluster-dispositions")
    require_vision_artifact_dependencies(args.artifacts_dir)
    try:
        disposition_document = load_cluster_disposition_document(args.cluster_dispositions)
        reference_difference_document = load_reference_difference_document(
            args.reference_differences
        )
    except ValueError as exc:
        parser.error(str(exc))

    print(f"GT     {args.gt}")
    print(f"output {args.output}")
    print(f"page {args.page} at {args.dpi} DPI\n")

    geometry = report_geometry(
        args.gt,
        args.output,
        page=args.page,
        large_shift=args.large_shift,
        fine_shift=args.fine_shift,
    )
    print()
    if args.lines:
        report_matched_lines(args.gt, args.output, page=args.page)
        print()
    histogram_result: dict[str, float] | None = None
    clusters: list[DiffCluster] | None = None
    if has_imagemagick():
        with tempfile.TemporaryDirectory() as raw_dir:
            out_dir = Path(raw_dir)
            gt_png = render_page(args.gt, args.page, args.dpi, out_dir, "gt")
            other_png = render_page(args.output, args.page, args.dpi, out_dir, "candidate")
            histogram_result = report_histogram(gt_png, other_png)
            print()
            normalised = report_pixels(gt_png, other_png, out_dir)
            print()
            clusters = diff_cluster_census(normalised, other_png, out_dir, args.dpi)
            report_diff_clusters(clusters, args.page)
            if args.artifacts_dir is not None:
                matches = match_text_line_instances(
                    page_text_lines(args.gt, args.page),
                    page_text_lines(args.output, args.page),
                )
                page_matches = [
                    match
                    for match in matches
                    if abs(match.dx) > (args.fine_shift or args.large_shift)
                    or abs(match.dy) > (args.fine_shift or args.large_shift)
                ]
                print()
                preserve_vision_artifacts(
                    gt_png,
                    other_png,
                    args.artifacts_dir,
                    args.page,
                    args.dpi,
                    page_matches,
                )
    else:
        print("## Histogram and pixel difference — SKIPPED")
        print("  ImageMagick is absent: neither `magick` (7.x) nor all of "
              f"{', '.join(f'`{tool}`' for tool in IMAGEMAGICK_TOOLS)} (6.x)")
        print("  is on PATH. Install `imagemagick` to measure colour and ink.")
    cluster_audit_report: dict[str, object] | None = None
    if args.cluster_report is not None:
        cluster_audit_report = build_cluster_audit_report(
            clusters,
            page=args.page,
            dpi=args.dpi,
            disposition_document=disposition_document,
            reference_difference_document=reference_difference_document,
            strict=args.strict_clusters,
        )
        write_cluster_audit_report(args.cluster_report, cluster_audit_report)
        print()
        report_cluster_audit(cluster_audit_report, args.cluster_report)
    print()
    diagnose(geometry, histogram_result, clusters)
    audit_shift_count = (
        geometry.get("fine_shift_count", 0.0)
        if args.fine_shift is not None
        else geometry.get("large_shift_count", 0.0)
    )
    audit_failed = args.audit and bool(audit_shift_count)
    cluster_audit_failed = bool(
        args.strict_clusters
        and cluster_audit_report is not None
        and not cluster_audit_report["passed"]
    )
    if audit_failed:
        print()
        print(
            "AUDIT FAILED: text-instance shifts past the active gate are layout "
            "differences, not antialiasing; inspect and track every line above."
        )
    if cluster_audit_failed:
        print()
        print(
            "AUDIT FAILED: every material diff cluster must have a valid explicit-ID "
            "disposition; inspect the machine report above."
        )
    if audit_failed or cluster_audit_failed:
        raise SystemExit(1)


if __name__ == "__main__":
    sys.exit(main())
