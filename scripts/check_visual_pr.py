#!/usr/bin/env python3
"""Fail a pull request when its visual audit or evidence is incomplete."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import shutil
import struct
import subprocess
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

try:
    from reference_exporter_differences import validate_reference_difference_document
except ModuleNotFoundError:  # Imported as scripts.check_visual_pr in unit tests.
    from scripts.reference_exporter_differences import (
        validate_reference_difference_document,
    )


VISUAL_ROOT = Path("assets/bugfixes")
EVIDENCE_PATH = re.compile(
    r"^assets/bugfixes/issue-(?P<issue>\d+)/"
    r"(?P<name>gt|before|after|compare|native)\.(?P<ext>[^/]+)$"
)
LAYOUT_AUDIT_PATH = re.compile(
    r"^assets/bugfixes/issue-(?P<issue>\d+)/layout-audit\.json$"
)
RENDER_CLUSTER_REPORT_PATH = re.compile(
    r"^assets/bugfixes/issue-(?P<issue>\d+)/render-clusters-page-(?P<page>\d+)\.json$"
)
REFERENCE_DIFFERENCE_PATH = re.compile(
    r"^assets/bugfixes/issue-(?P<issue>\d+)/reference-exporter-differences\.json$"
)
REFERENCE_DIFFERENCE_ID = re.compile(r"ref:([a-z0-9]+(?:-[a-z0-9]+)*)")
# These bounds match diff_page's missing_text/extra_text diagnostic labels.
LAYOUT_TEXT_SAMPLE_LIMIT = 20
LAYOUT_TEXT_LABEL_LIMIT = 60
# Prose living under assets/bugfixes/ carries no pixels, so it is neither evidence
# to validate nor a rendered change to audit. Without this the file documenting the
# evidence convention could never be edited (#539). Image suffixes stay outside the
# set so a stray screenshot is still rejected rather than silently waved through.
BOOKKEEPING_SUFFIXES = frozenset({".md", ".txt"})
AUDIT_ROWS = (
    "Page count/order",
    "Element presence",
    "Position/size",
    "Rotation/flip",
    "Fill",
    "Stroke/border",
    "Shape outline geometry",
    "Text content",
    "Font family/weight/style",
    "Text color",
    "Alignment",
    "Line/paragraph spacing",
    "Clipping/overflow",
)
INSPECTION_ITEMS = (
    "Rendered all evidence at 150 DPI or higher",
    "Stored progressive JPEG quality 86 assets with metadata stripped",
    "Used Codex/Claude vision to inspect the full GT/output pages, diff, and matched crops",
    "Inspected matched region crops at full resolution",
    "Ran compare_layout.py --audit --fine-shift PT and dispositioned every fine/large "
    "text-instance shift, rectangle geometry deviation, painted-text visibility mismatch, "
    "and visible-fill occlusion",
    "Ran compare_render.py --cluster-report PATH --strict-clusters and dispositioned "
    "every material 5% fuzz diff cluster by explicit ID",
    "Inventoried hairlines and border dash styles",
    "Inventoried font weight, italic, and underline emphasis",
)
FIX_PREVIEW_LABELS = ("GT", "Before", "After")
DEFECT_PREVIEW_LABELS = ("Compare",)
ALLOWED_RESULTS = (
    "Matches GT",
    "Fixed",
    "No deviation observed",
    "Reference difference:",
)
VISION_WORDS = re.compile(
    r"(?i)\b(?:page|diff|crop|text|title|label|line|shape|image|chart|table|"
    r"position|align(?:ment|ed)?|offset|colour|color|fill|stroke|border|font|"
    r"spacing|clip(?:ping|ped)?|overflow|rotation|size|weight)\b"
)
ACCEPTED_RENDERING_CLASSES = frozenset(
    {
        "glyph-edge-rasterization",
        "photo-resampling",
        "gradient-rasterization",
        "shape-edge-antialiasing",
    }
)


@dataclass(frozen=True)
class JpegInfo:
    width: int
    height: int
    progressive: bool
    density_dpi: tuple[float, float] | None
    metadata_markers: tuple[str, ...]


def extract_section(body: str, heading: str) -> str:
    pattern = re.compile(
        rf"(?ms)^{re.escape(heading)}\s*$\n(?P<section>.*?)(?=^##\s|\Z)"
    )
    match = pattern.search(body)
    return match.group("section") if match else ""


def checked(section: str, label: str) -> bool:
    return bool(re.search(rf"(?mi)^- \[[xX]\] {re.escape(label)}\s*$", section))


def field(section: str, name: str) -> str | None:
    match = re.search(rf"(?mi)^- {re.escape(name)}:\s*(.*?)\s*$", section)
    if not match:
        return None
    value = match.group(1).strip().strip("`")
    if not value or "<!--" in value:
        return None
    return value


def audit_table(section: str) -> dict[str, str]:
    rows: dict[str, str] = {}
    for line in section.splitlines():
        if not line.startswith("|"):
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if len(cells) == 2 and cells[0] in AUDIT_ROWS:
            rows[cells[0]] = cells[1]
    return rows


def _reference_difference_ids(value: str) -> list[str] | None:
    """Parse unique comma-separated exact IDs from one deviation-table row."""

    if not value.startswith("Reference difference:"):
        return None
    references = [
        reference.strip()
        for reference in value.removeprefix("Reference difference:").split(",")
    ]
    matches = [
        REFERENCE_DIFFERENCE_ID.fullmatch(reference) for reference in references
    ]
    if not references or not all(matches) or len(references) != len(set(references)):
        return None
    return [match.group(1) for match in matches if match is not None]


def rendered_preview_urls(section: str, labels: tuple[str, ...]) -> dict[str, str]:
    rendered_markup = re.sub(r"(?s)<!--.*?-->", "", section)
    rendered_markup = re.sub(r"(?s)```.*?```", "", rendered_markup)
    rendered_markup = re.sub(r"`[^`\n]*`", "", rendered_markup)
    urls: dict[str, str] = {}
    for label in labels:
        match = re.search(
            rf"!\[{re.escape(label)}\]\((?P<url>https?://[^\s)]+)\)",
            rendered_markup,
        )
        if match:
            urls[label] = match.group("url")
    return urls


def is_visual_asset(path: str) -> bool:
    """Report whether a changed path is an image under the visual evidence root."""
    if not path.startswith(f"{VISUAL_ROOT.as_posix()}/"):
        return False
    return PurePosixPath(path).suffix.lower() not in BOOKKEEPING_SUFFIXES


def validate_pr_body(body: str, changed_paths: list[str]) -> list[str]:
    errors: list[str] = []
    impact = extract_section(body, "## Visual impact")
    no_change = checked(impact, "No rendered PDF change")
    visual = checked(impact, "Rendered PDF change or visual evidence added")
    visual_assets_changed = any(is_visual_asset(path) for path in changed_paths)

    if no_change == visual:
        errors.append("Select exactly one Visual impact checkbox.")
        return errors

    if no_change:
        reason = field(impact, "Reason")
        if not reason:
            errors.append("Explain why the PR has no rendered PDF change in Visual impact > Reason.")
        if visual_assets_changed:
            errors.append("assets/bugfixes changes require 'Rendered PDF change or visual evidence added'.")
        return errors

    audit = extract_section(body, "## Visual audit")
    if not audit:
        return ["Rendered changes require a ## Visual audit section."]

    issue_value = field(audit, "Issue")
    issue_match = re.fullmatch(r"#(\d+)", issue_value or "")
    if not issue_match:
        errors.append("Visual audit > Issue must be a single issue reference such as #123.")
        issue_number = None
    else:
        issue_number = issue_match.group(1)

    for name in ("Fixture", "Page(s)"):
        if not field(audit, name):
            errors.append(f"Visual audit > {name} is required.")

    reference_differences = field(audit, "Reference exporter differences")
    if not reference_differences:
        errors.append(
            "Visual audit > Reference exporter differences must be `None` or the "
            "issue-local machine report path."
        )
    reference_artifact_changed = any(
        REFERENCE_DIFFERENCE_PATH.fullmatch(path)
        or (match := EVIDENCE_PATH.fullmatch(path)) is not None
        and match.group("name") == "native"
        for path in changed_paths
    )
    if reference_artifact_changed and reference_differences == "None":
        errors.append(
            "Changing native/reference-exporter evidence requires its issue-local "
            "Reference exporter differences report."
        )

    renderer = field(audit, "Renderer and DPI")
    dpi_match = re.search(r"(?i)(\d+(?:\.\d+)?)\s*DPI", renderer or "")
    if not dpi_match or float(dpi_match.group(1)) < 150:
        errors.append("Visual audit > Renderer and DPI must record at least 150 DPI.")

    mode = field(audit, "Evidence mode")
    if mode not in {"fix", "defect"}:
        errors.append("Visual audit > Evidence mode must be `fix` or `defect`.")
    elif issue_number:
        expected = (
            ("GT", "gt"),
            ("Before", "before"),
            ("After", "after"),
        ) if mode == "fix" else (("Compare", "compare"),)
        for field_name, basename in expected:
            expected_path = f"assets/bugfixes/issue-{issue_number}/{basename}.jpg"
            if field(audit, field_name) != expected_path:
                errors.append(f"Visual audit > {field_name} must be `{expected_path}`.")

    preview_labels: tuple[str, ...] = ()
    if mode == "fix":
        preview_labels = FIX_PREVIEW_LABELS
    elif mode == "defect":
        preview_labels = DEFECT_PREVIEW_LABELS
    preview_urls = rendered_preview_urls(audit, preview_labels)
    for label in preview_labels:
        if label not in preview_urls:
            errors.append(
                f"Visual audit must include a rendered preview for {label}: "
                f"![{label}](https://...)."
            )
    if len(preview_urls) == len(preview_labels) and len(set(preview_urls.values())) != len(
        preview_labels
    ):
        errors.append("Each rendered preview must reference a different evidence image.")

    follow_up_value = field(audit, "New follow-up issues found in this audit")
    if not follow_up_value:
        errors.append("Visual audit must summarize newly discovered follow-up issues or say None.")
    elif follow_up_value != "None":
        follow_up_issues = {int(number) for number in re.findall(r"#(\d+)", follow_up_value)}
        if not follow_up_issues:
            errors.append("New follow-up issues must use issue references such as #123, or None.")
        else:
            remaining_issues = remaining_issue_numbers(body)
            unclassified = follow_up_issues - remaining_issues
            if unclassified:
                references = ", ".join(f"#{number}" for number in sorted(unclassified))
                errors.append(
                    f"New follow-up issues must also classify a Remaining deviation: {references}."
                )

    vision_findings = field(audit, "Model vision findings")
    descriptive_words = re.findall(r"[A-Za-z]{3,}", vision_findings or "")
    if (
        not vision_findings
        or len(vision_findings) < 20
        or len(descriptive_words) < 4
        or not VISION_WORDS.search(vision_findings)
    ):
        errors.append(
            "Visual audit > Model vision findings must record a substantive direct "
            "inspection of the full pages, diff, and crops; numeric metrics alone do not count."
        )

    for item in INSPECTION_ITEMS:
        if not checked(audit, item):
            errors.append(f"Required visual inspection is not checked: {item}.")

    rows = audit_table(audit)
    for row in AUDIT_ROWS:
        result = rows.get(row, "")
        if not result or "<!--" in result:
            errors.append(f"Deviation audit row is incomplete: {row}.")
        elif result.startswith("Remaining:"):
            if not re.search(r"#\d+", result):
                errors.append(f"Remaining deviation must reference an issue: {row}.")
        elif result.startswith("Reference difference:"):
            if _reference_difference_ids(result) is None:
                errors.append(
                    "Reference exporter differences must use unique exact ref:<id> "
                    f"values separated by commas: {row}."
                )
        elif not result.startswith(ALLOWED_RESULTS):
            errors.append(
                f"Deviation audit row '{row}' must start with Matches GT, Fixed, "
                "No deviation observed, Reference difference: ref:<id>, or "
                "Remaining: #N."
            )

    if not visual_assets_changed:
        errors.append("Rendered visual changes require evidence under assets/bugfixes/issue-<number>/.")

    return errors


def remaining_issue_numbers(body: str) -> set[int]:
    audit = extract_section(body, "## Visual audit")
    numbers: set[int] = set()
    for result in audit_table(audit).values():
        if result.startswith("Remaining:"):
            numbers.update(int(number) for number in re.findall(r"#(\d+)", result))
    return numbers


def audit_reference_difference_ids(body: str) -> set[str]:
    audit = extract_section(body, "## Visual audit")
    difference_ids: set[str] = set()
    for result in audit_table(audit).values():
        references = _reference_difference_ids(result)
        if references is not None:
            difference_ids.update(references)
    return difference_ids


def layout_audit_categories(report: object) -> dict[str, bool]:
    """Return the material finding categories recorded by compare_layout JSON."""

    if not isinstance(report, dict):
        raise ValueError("the top level must be an object")
    pages = report.get("pages")
    gt_pages = report.get("gt_pages")
    out_pages = report.get("out_pages")
    if (
        not isinstance(pages, list)
        or type(gt_pages) is not int
        or type(out_pages) is not int
        or gt_pages <= 0
        or out_pages <= 0
    ):
        raise ValueError("pages must be a list and page counts must be positive integers")
    if len(pages) != min(gt_pages, out_pages):
        raise ValueError("page vectors must cover every comparable page")

    has_text_flow_findings = False
    has_visible_fill_findings = False
    has_rect_geometry_findings = False
    has_large_shifts = False
    has_fine_shifts = False
    for page_number, page in enumerate(pages, start=1):
        if not isinstance(page, dict):
            raise ValueError(f"page {page_number} must be an object")
        try:
            line_counts = page["lines"]
            wraps = page["wraps"]
            reflow = page["reflow"]
            instances = page["instances"]
            visibility = page.get("visibility", {"mismatch_count": 0})
            visible_fills = page.get("visible_fills", {"mismatch_count": 0})
            rects = page.get("rects", {"geometry_mismatch_count": 0})
            text_flow_counts = (
                line_counts["missing"],
                line_counts["extra"],
                wraps["count"],
                reflow["gt_lines"],
                reflow["out_lines"],
                visibility["mismatch_count"],
            )
            visible_fill_count = visible_fills["mismatch_count"]
            rect_geometry_count = rects.get("geometry_mismatch_count", 0)
            large_shift_count = instances["large_shift_count"]
            fine_shift_count = instances["fine_shift_count"]
        except (KeyError, TypeError) as exc:
            raise ValueError(f"page {page_number} is missing compare_layout fields") from exc
        counts = (
            *text_flow_counts,
            visible_fill_count,
            rect_geometry_count,
            large_shift_count,
            fine_shift_count,
        )
        if any(type(count) is not int or count < 0 for count in counts):
            raise ValueError(f"page {page_number} finding counts must be non-negative integers")

        has_text_flow_findings |= any(text_flow_counts)
        has_visible_fill_findings |= visible_fill_count > 0
        has_rect_geometry_findings |= rect_geometry_count > 0
        has_large_shifts |= large_shift_count > 0
        has_fine_shifts |= fine_shift_count > 0

    return {
        "page count": gt_pages != out_pages,
        "text flow": has_text_flow_findings,
        "visible fills": has_visible_fill_findings,
        "rectangle geometry": has_rect_geometry_findings,
        "large shifts": has_large_shifts,
        "fine shifts": has_fine_shifts,
    }


def layout_audit_fine_shift_threshold(report: object) -> float:
    """Return the single positive fine-detail threshold recorded on every page."""

    if not isinstance(report, dict) or not isinstance(report.get("pages"), list):
        raise ValueError("the report must contain page vectors")
    thresholds: list[float] = []
    for page_number, page in enumerate(report["pages"], start=1):
        try:
            threshold = page["instances"]["fine_shift_threshold"]
        except (KeyError, TypeError) as exc:
            raise ValueError(
                f"page {page_number} is missing the fine-detail threshold"
            ) from exc
        if isinstance(threshold, bool) or not isinstance(threshold, (int, float)):
            raise ValueError(f"page {page_number} fine-detail threshold must be numeric")
        if threshold <= 0:
            raise ValueError(f"page {page_number} fine-detail threshold must be positive")
        thresholds.append(float(threshold))
    if not thresholds:
        raise ValueError("the report has no fine-detail threshold")
    if any(abs(threshold - thresholds[0]) > 1e-9 for threshold in thresholds[1:]):
        raise ValueError("fine-detail threshold differs between pages")
    return thresholds[0]


def disposition_issue_numbers(value: str | None) -> set[int] | None:
    if value == "Pass":
        return set()
    if not value or not re.fullmatch(r"#\d+(?:\s*,\s*#\d+)*", value):
        return None
    return {int(number) for number in re.findall(r"#(\d+)", value)}


def layout_dispositions(value: str | None) -> tuple[set[int], set[str]] | None:
    """Parse a layout category's open-issue and evidence-backed dispositions."""

    if value == "Pass":
        return set(), set()
    if not value:
        return None
    tokens = [token.strip() for token in value.split(",")]
    if not tokens or any(
        re.fullmatch(r"(?:#[1-9]\d*|ref:[a-z0-9]+(?:-[a-z0-9]+)*)", token) is None
        for token in tokens
    ):
        return None
    issues = {
        int(token[1:]) for token in tokens if token.startswith("#")
    }
    references = {
        token.removeprefix("ref:") for token in tokens if token.startswith("ref:")
    }
    return issues, references


def _current_visibility_findings(
    report: dict[str, object], declared_pages: set[int]
) -> set[tuple[int, str, str, str, int]]:
    """Return occurrence-bounded painted-text visibility findings."""

    pages = report.get("pages")
    if not isinstance(pages, list) or len(pages) != len(declared_pages):
        raise ValueError(
            "declared Page(s) must map one-to-one to the layout report page vectors"
        )
    findings: set[tuple[int, str, str, str, int]] = set()
    for page_number, page_report in zip(sorted(declared_pages), pages, strict=True):
        if not isinstance(page_report, dict):
            raise ValueError(f"page {page_number} must be an object")
        visibility = page_report.get("visibility")
        if not isinstance(visibility, dict):
            continue
        raw_mismatches = visibility.get("mismatches", [])
        if not isinstance(raw_mismatches, list):
            raise ValueError(f"page {page_number} visibility.mismatches must be a list")
        occurrences: dict[tuple[str, str, str], int] = {}
        for mismatch in raw_mismatches:
            if not isinstance(mismatch, dict) or set(mismatch) != {"label", "gt", "out"}:
                raise ValueError(
                    f"page {page_number} visibility mismatch must contain exactly "
                    "label, gt, and out"
                )
            key = (str(mismatch["label"]), str(mismatch["gt"]), str(mismatch["out"]))
            occurrence = occurrences.get(key, 0) + 1
            occurrences[key] = occurrence
            findings.add((page_number, *key, occurrence))
    return findings


def _reference_visibility_key(
    difference: dict[str, object],
) -> tuple[int, str, str, str, int] | None:
    finding = difference.get("layout_finding")
    page = difference.get("page")
    if type(page) is not int or not isinstance(finding, dict):
        return None
    try:
        return (
            page,
            str(finding["label"]),
            str(finding["gt"]),
            str(finding["out"]),
            int(finding["occurrence"]),
        )
    except (KeyError, TypeError, ValueError):
        return None


def _current_rasterized_text_findings(
    report: dict[str, object], declared_pages: set[int]
) -> set[tuple[int, str, str, int]]:
    """Return exact labels and occurrences from compare_layout's missing/extra lists."""

    pages = report.get("pages")
    if not isinstance(pages, list) or len(pages) != len(declared_pages):
        raise ValueError(
            "declared Page(s) must map one-to-one to the layout report page vectors"
        )
    findings: set[tuple[int, str, str, int]] = set()
    for page_number, page_report in zip(sorted(declared_pages), pages, strict=True):
        if not isinstance(page_report, dict):
            raise ValueError(f"page {page_number} must be an object")
        lines = page_report.get("lines")
        if not isinstance(lines, dict):
            raise ValueError(f"page {page_number} lines must be an object")
        for report_field, count_field, side in (
            ("extra_text", "extra", "gt"),
            ("missing_text", "missing", "out"),
        ):
            labels = lines.get(report_field)
            count = lines.get(count_field)
            if not isinstance(labels, list):
                raise ValueError(f"page {page_number} lines.{report_field} must be a list")
            if type(count) is not int or count < 0:
                raise ValueError(f"page {page_number} lines.{count_field} must be non-negative")
            sample_count = min(count, LAYOUT_TEXT_SAMPLE_LIMIT)
            if len(labels) != sample_count:
                raise ValueError(
                    f"page {page_number} lines.{report_field} must list the first "
                    f"{sample_count} of {count} findings"
                )
            occurrences: dict[str, int] = {}
            for index, label in enumerate(labels, start=1):
                if (
                    not isinstance(label, str)
                    or not label.strip()
                    or len(label) > LAYOUT_TEXT_LABEL_LIMIT
                ):
                    raise ValueError(
                        f"page {page_number} lines.{report_field} entry {index} "
                        "must be a non-empty label of at most "
                        f"{LAYOUT_TEXT_LABEL_LIMIT} characters"
                    )
                occurrence = occurrences.get(label, 0) + 1
                occurrences[label] = occurrence
                findings.add((page_number, label, side, occurrence))
    return findings


def _reference_rasterized_text_key(
    difference: dict[str, object],
) -> tuple[int, str, str, int] | None:
    finding = difference.get("layout_finding")
    page = difference.get("page")
    if type(page) is not int or not isinstance(finding, dict):
        return None
    label = finding.get("label")
    side = finding.get("side")
    occurrence = finding.get("occurrence")
    if (
        not isinstance(label, str)
        or not isinstance(side, str)
        or type(occurrence) is not int
    ):
        return None
    return page, label, side, occurrence


def _shift_label_occurrence(label: str) -> tuple[str, int]:
    """Split compare_layout's repeated-label suffix into text and occurrence."""

    match = re.fullmatch(
        r"(?P<label>.+?) \[(?P<index>[1-9]\d*)/(?P<count>[1-9]\d*)\]",
        label,
    )
    if match is None:
        return label, 1
    occurrence = int(match.group("index"))
    if occurrence > int(match.group("count")):
        raise ValueError(f"invalid repeated shift label: {label}")
    return match.group("label"), occurrence


def _current_shift_findings(
    report: dict[str, object], declared_pages: set[int], category: str
) -> set[tuple[int, str, int, float, float]]:
    """Return exact, occurrence-bounded shifts from the requested audit list."""

    pages = report.get("pages")
    if not isinstance(pages, list) or len(pages) != len(declared_pages):
        raise ValueError(
            "declared Page(s) must map one-to-one to the layout report page vectors"
        )
    if category not in {"large shifts", "fine shifts"}:
        raise ValueError(f"unsupported shift category: {category}")
    list_name = "large_shifts" if category == "large shifts" else "fine_shifts"
    count_name = (
        "large_shift_count" if category == "large shifts" else "fine_shift_count"
    )
    findings: set[tuple[int, str, int, float, float]] = set()
    selectors: set[tuple[int, str, int]] = set()
    for page_number, page_report in zip(sorted(declared_pages), pages, strict=True):
        if not isinstance(page_report, dict):
            raise ValueError(f"page {page_number} must be an object")
        instances = page_report.get("instances")
        if not isinstance(instances, dict):
            raise ValueError(f"page {page_number} instances must be an object")
        raw_findings = instances.get(list_name)
        count = instances.get(count_name)
        if not isinstance(raw_findings, list):
            raise ValueError(f"page {page_number} instances.{list_name} must be a list")
        if type(count) is not int or count != len(raw_findings):
            raise ValueError(
                f"page {page_number} instances.{count_name} must match {list_name}"
            )
        for index, finding in enumerate(raw_findings, start=1):
            if not isinstance(finding, dict):
                raise ValueError(
                    f"page {page_number} {list_name} entry {index} must be an object"
                )
            label = finding.get("label")
            if not isinstance(label, str) or not label.strip():
                raise ValueError(
                    f"page {page_number} {list_name} entry {index} needs a label"
                )
            coordinates: list[float] = []
            for coordinate in ("dx", "dy"):
                value = finding.get(coordinate)
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    raise ValueError(
                        f"page {page_number} {list_name} entry {index} "
                        f"{coordinate} must be numeric"
                    )
                try:
                    numeric_value = float(value)
                except OverflowError as exc:
                    raise ValueError(
                        f"page {page_number} {list_name} entry {index} "
                        f"{coordinate} must be finite"
                    ) from exc
                if not math.isfinite(numeric_value):
                    raise ValueError(
                        f"page {page_number} {list_name} entry {index} "
                        f"{coordinate} must be finite"
                    )
                coordinates.append(numeric_value)
            base_label, occurrence = _shift_label_occurrence(label)
            selector = (page_number, base_label, occurrence)
            if selector in selectors:
                raise ValueError(
                    f"page {page_number} {list_name} repeats an exact label occurrence"
                )
            selectors.add(selector)
            findings.add((page_number, base_label, occurrence, *coordinates))
    return findings


def _reference_shift_key(
    difference: dict[str, object],
) -> tuple[int, str, int, float, float] | None:
    finding = difference.get("layout_finding")
    page = difference.get("page")
    if type(page) is not int or not isinstance(finding, dict):
        return None
    label = finding.get("label")
    occurrence = finding.get("occurrence")
    dx = finding.get("dx")
    dy = finding.get("dy")
    if (
        not isinstance(label, str)
        or type(occurrence) is not int
        or isinstance(dx, bool)
        or not isinstance(dx, (int, float))
        or isinstance(dy, bool)
        or not isinstance(dy, (int, float))
    ):
        return None
    try:
        dx_value = float(dx)
        dy_value = float(dy)
    except OverflowError:
        return None
    if not math.isfinite(dx_value) or not math.isfinite(dy_value):
        return None
    return page, label, occurrence, dx_value, dy_value


def decoded_pixel_delta(before: Path, after: Path) -> int:
    """Return ImageMagick's exact decoded-pixel AE count for two images."""

    if shutil.which("magick") is not None:
        command = ["magick", "compare"]
    elif shutil.which("compare") is not None:
        command = ["compare"]
    else:
        raise RuntimeError(
            "ImageMagick is required to verify text-layer-only pixel equality"
        )

    result = subprocess.run(
        [*command, "-metric", "AE", str(before), str(after), "null:"],
        check=False,
        capture_output=True,
        text=True,
    )
    metric_output = result.stderr.strip() or result.stdout.strip()
    metric_match = re.search(
        r"(?:^|\s)(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(?:\s|$)",
        metric_output,
    )
    if result.returncode not in {0, 1} or not metric_match:
        detail = metric_output or f"exit status {result.returncode}"
        raise RuntimeError(f"ImageMagick could not compare the evidence: {detail}")
    return int(float(metric_match.group(1)))


def validate_reference_exporter_differences(
    body: str, changed_paths: list[str], root: Path
) -> tuple[dict[str, object] | None, list[str]]:
    """Load exact GT-exporter dispositions and verify their local image evidence."""

    impact = extract_section(body, "## Visual impact")
    if checked(impact, "No rendered PDF change"):
        return None, []
    audit = extract_section(body, "## Visual audit")
    if field(audit, "Evidence mode") != "fix":
        return None, []
    issue_match = re.fullmatch(r"#(\d+)", field(audit, "Issue") or "")
    if issue_match is None:
        return None, []

    declared_path = field(audit, "Reference exporter differences")
    if declared_path == "None":
        return None, []
    issue_number = issue_match.group(1)
    expected_path = (
        f"assets/bugfixes/issue-{issue_number}/reference-exporter-differences.json"
    )
    errors: list[str] = []
    if declared_path != expected_path:
        return None, [
            "Visual audit > Reference exporter differences must be `None` or "
            f"`{expected_path}`."
        ]
    if expected_path not in changed_paths:
        errors.append(
            f"{expected_path}: reference exporter difference report must be changed "
            "in this pull request."
        )

    try:
        document = json.loads((root / expected_path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return None, [f"{expected_path}: invalid reference exporter differences: {exc}."]

    reference_registry, schema_errors = validate_reference_difference_document(document)
    errors.extend(f"{expected_path}: {error}." for error in schema_errors)
    if schema_errors or not isinstance(document, dict):
        return None, errors

    audit_ids = audit_reference_difference_ids(body)
    undocumented_ids = set(reference_registry) - audit_ids
    unknown_ids = audit_ids - set(reference_registry)
    if undocumented_ids:
        errors.append(
            f"{expected_path}: every difference must appear in one `Reference "
            "difference: ref:<id>` deviation row; missing "
            + ", ".join(sorted(undocumented_ids))
            + "."
        )
    if unknown_ids:
        errors.append(
            f"Visual audit deviation rows reference differences absent from "
            f"{expected_path}: "
            + ", ".join(sorted(unknown_ids))
            + "."
        )

    expected_native_path = f"assets/bugfixes/issue-{issue_number}/native.jpg"
    if field(audit, "Native") != expected_native_path:
        errors.append(
            f"Visual audit > Native must be `{expected_native_path}` when reference "
            "exporter differences are declared."
        )
    if "Native" not in rendered_preview_urls(audit, ("Native",)):
        errors.append(
            "Visual audit must render the native-export evidence as "
            "![Native](https://...)."
        )

    native_ground_truth = document.get("reference_export") is None
    expected_evidence = {"native_export": expected_native_path}
    ground_truth_path = f"assets/bugfixes/issue-{issue_number}/gt.jpg"
    if not native_ground_truth:
        expected_evidence["reference_export"] = ground_truth_path
    verified_evidence: dict[str, Path] = {}
    for export_name, evidence_path in expected_evidence.items():
        export = document[export_name]
        assert isinstance(export, dict)
        if export.get("evidence_path") != evidence_path:
            errors.append(
                f"{expected_path}: {export_name}.evidence_path must be `{evidence_path}`."
            )
            continue
        absolute_path = root / evidence_path
        if not absolute_path.is_file():
            errors.append(f"{evidence_path}: {export_name} evidence is missing.")
            continue
        verified_evidence[export_name] = absolute_path
        errors.extend(validate_jpeg(absolute_path))
        actual_hash = hashlib.sha256(absolute_path.read_bytes()).hexdigest()
        if actual_hash != export.get("evidence_sha256"):
            errors.append(
                f"{expected_path}: {export_name} evidence SHA-256 does not match "
                f"{evidence_path}."
            )

    if native_ground_truth:
        absolute_ground_truth_path = root / ground_truth_path
        if not absolute_ground_truth_path.is_file():
            errors.append(f"{ground_truth_path}: native ground-truth evidence is missing.")
        else:
            ground_truth_errors = validate_jpeg(absolute_ground_truth_path)
            errors.extend(ground_truth_errors)
            if not ground_truth_errors:
                verified_evidence["ground_truth"] = absolute_ground_truth_path

    if native_ground_truth and {"ground_truth", "native_export"} <= set(verified_evidence):
        try:
            evidence_delta = decoded_pixel_delta(
                verified_evidence["ground_truth"],
                verified_evidence["native_export"],
            )
        except RuntimeError as exc:
            errors.append(f"GT/native evidence comparison failed: {exc}.")
        else:
            if evidence_delta != 0:
                errors.append(
                    "GT/native evidence must be pixel-identical when reference_export is null."
                )
    elif not native_ground_truth and set(verified_evidence) == set(expected_evidence):
        try:
            evidence_delta = decoded_pixel_delta(
                verified_evidence["reference_export"],
                verified_evidence["native_export"],
            )
        except RuntimeError as exc:
            errors.append(f"Reference/native evidence comparison failed: {exc}.")
        else:
            if evidence_delta == 0:
                errors.append(
                    "Reference/native evidence must contain a decoded-pixel difference."
                )

    return document, errors


def validate_layout_audit(
    body: str,
    changed_paths: list[str],
    root: Path,
    *,
    reference_differences: dict[str, object] | None = None,
) -> list[str]:
    """Tie machine-detected layout failures to classified open issue references."""

    impact = extract_section(body, "## Visual impact")
    if checked(impact, "No rendered PDF change"):
        return []

    audit = extract_section(body, "## Visual audit")
    if field(audit, "Evidence mode") != "fix":
        return []
    issue_match = re.fullmatch(r"#(\d+)", field(audit, "Issue") or "")
    if not issue_match:
        return []

    issue_number = issue_match.group(1)
    expected_path = f"assets/bugfixes/issue-{issue_number}/layout-audit.json"
    errors: list[str] = []
    if field(audit, "Layout audit report") != expected_path:
        errors.append(f"Visual audit > Layout audit report must be `{expected_path}`.")
    if expected_path not in changed_paths:
        errors.append(f"{expected_path}: layout audit report must be changed in this pull request.")
    expected_after = f"assets/bugfixes/issue-{issue_number}/after.jpg"
    text_layer_only = field(audit, "Text-layer-only") == "Yes"
    if text_layer_only and field(audit, "Pixel delta") != "0":
        errors.append("Visual audit > Pixel delta must be 0 for a text-layer-only fix.")
    if expected_after not in changed_paths and not text_layer_only:
        errors.append(
            f"{expected_path}: {expected_after} must be changed with the layout audit report."
        )
    if text_layer_only:
        before_path = root / f"assets/bugfixes/issue-{issue_number}/before.jpg"
        after_path = root / expected_after
        missing_evidence = False
        for evidence_path in (before_path, after_path):
            if not evidence_path.is_file():
                errors.append(
                    f"{evidence_path.relative_to(root)}: current evidence is required."
                )
                missing_evidence = True
        if not missing_evidence:
            evidence_errors = [
                *validate_jpeg(before_path),
                *validate_jpeg(after_path),
            ]
            errors.extend(evidence_errors)
            if not evidence_errors:
                try:
                    actual_delta = decoded_pixel_delta(before_path, after_path)
                except RuntimeError as exc:
                    errors.append(f"Text-layer-only evidence comparison failed: {exc}.")
                else:
                    if actual_delta != 0:
                        errors.append(
                            "Text-layer-only evidence must have zero decoded-pixel "
                            f"delta, got {actual_delta}."
                        )

    report_path = root / expected_path
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
        findings = layout_audit_categories(report)
        report_fine_shift_threshold = layout_audit_fine_shift_threshold(report)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        errors.append(f"{expected_path}: invalid layout audit report: {exc}.")
        return errors

    threshold_value = field(audit, "Fine-detail threshold")
    threshold_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*pt", threshold_value or "", re.I)
    if not threshold_match:
        errors.append(
            "Visual audit > Fine-detail threshold is required in points, for example `0.5pt`."
        )
    elif abs(float(threshold_match.group(1)) - report_fine_shift_threshold) > 1e-9:
        errors.append(
            "Visual audit > Fine-detail threshold must match the layout report's "
            f"{report_fine_shift_threshold:g}pt threshold."
        )

    remaining_issues = remaining_issue_numbers(body)
    reference_registry, reference_errors = (
        validate_reference_difference_document(reference_differences)
        if reference_differences is not None
        else ({}, [])
    )
    errors.extend(f"Reference exporter differences: {error}." for error in reference_errors)
    declared_pages = compared_pages(field(audit, "Page(s)"))
    current_visibility: set[tuple[int, str, str, str, int]] | None = None
    current_rasterized_text: set[tuple[int, str, str, int]] | None = None
    current_shifts: dict[str, set[tuple[int, str, int, float, float]]] = {}
    fields = {
        "page count": "Layout audit page count",
        "text flow": "Layout audit text flow",
        "visible fills": "Layout audit visible fills",
        "rectangle geometry": "Layout audit rectangle geometry",
        "large shifts": "Layout audit large shifts",
        "fine shifts": "Layout audit fine shifts",
    }
    for category, has_findings in findings.items():
        field_name = fields[category]
        disposition = layout_dispositions(field(audit, field_name))
        if has_findings:
            if disposition is None or not any(disposition):
                errors.append(
                    f"Visual audit > {field_name} has {category} findings and requires "
                    "one or more issue references or exact `ref:<id>` dispositions."
                )
                continue
            disposition_issues, disposition_references = disposition
            unclassified = disposition_issues - remaining_issues
            if unclassified:
                references = ", ".join(f"#{number}" for number in sorted(unclassified))
                errors.append(
                    f"Visual audit > {field_name} references must also appear in a "
                    f"Remaining deviation row: {references}."
                )
            if not disposition_references:
                continue
            reference_kinds = {
                "text flow": {"painted-text-visibility", "rasterized-text"},
                "large shifts": {"text-shift"},
                "fine shifts": {"text-shift"},
            }.get(category)
            if reference_kinds is None:
                errors.append(
                    f"Visual audit > {field_name} does not support reference exporter "
                    "differences; use an open issue."
                )
                continue
            if declared_pages is None:
                errors.append(
                    "Visual audit > Page(s) must be valid before reference exporter "
                    "differences can be matched."
                )
                continue
            current_findings: set[tuple[object, ...]]
            if category == "text flow":
                if current_visibility is None:
                    try:
                        current_visibility = _current_visibility_findings(
                            report, declared_pages
                        )
                    except ValueError as exc:
                        errors.append(f"{expected_path}: {exc}.")
                        current_visibility = set()
                if current_rasterized_text is None:
                    try:
                        current_rasterized_text = _current_rasterized_text_findings(
                            report, declared_pages
                        )
                    except ValueError as exc:
                        errors.append(f"{expected_path}: {exc}.")
                        current_rasterized_text = set()
                current_findings = set(current_visibility) | set(current_rasterized_text)
            else:
                if category not in current_shifts:
                    try:
                        current_shifts[category] = _current_shift_findings(
                            report, declared_pages, category
                        )
                    except ValueError as exc:
                        errors.append(f"{expected_path}: {exc}.")
                        current_shifts[category] = set()
                current_findings = set(current_shifts[category])
            referenced_keys: set[tuple[object, ...]] = set()
            referenced_rasterized_text: set[tuple[int, str, str, int]] = set()
            for difference_id in sorted(disposition_references):
                difference = reference_registry.get(difference_id)
                if difference is None:
                    errors.append(
                        f"Visual audit > {field_name} ref:{difference_id} does not name "
                        "a validated reference exporter difference."
                    )
                    continue
                if difference.get("kind") not in reference_kinds:
                    expected_kind = (
                        next(iter(reference_kinds))
                        if len(reference_kinds) == 1
                        else None
                    )
                    errors.append(
                        f"Visual audit > {field_name} ref:{difference_id} is not a "
                        + (
                            f"{expected_kind} difference."
                            if expected_kind is not None
                            else f"supported {category} difference."
                        )
                    )
                    continue
                if category == "text flow":
                    if difference.get("kind") == "painted-text-visibility":
                        reference_key = _reference_visibility_key(difference)
                    else:
                        reference_key = _reference_rasterized_text_key(difference)
                else:
                    reference_key = _reference_shift_key(difference)
                if reference_key is None or reference_key not in current_findings:
                    finding_name = (
                        "rasterized-text"
                        if difference.get("kind") == "rasterized-text"
                        else "visibility" if category == "text flow" else "shift"
                    )
                    errors.append(
                        f"Visual audit > {field_name} ref:{difference_id} does not match "
                        f"an exact current {finding_name} finding."
                    )
                    continue
                referenced_keys.add(reference_key)
                if difference.get("kind") == "rasterized-text":
                    referenced_rasterized_text.add(reference_key)

            if not disposition_issues:
                if category == "text flow":
                    other_text_flow_count = 0
                    referenced_by_page_and_side: dict[tuple[int, str], int] = {}
                    for page_number, _, side, _ in referenced_rasterized_text:
                        count_name = "extra" if side == "gt" else "missing"
                        selector = (page_number, count_name)
                        referenced_by_page_and_side[selector] = (
                            referenced_by_page_and_side.get(selector, 0) + 1
                        )
                    for page_number, page in zip(
                        sorted(declared_pages), report["pages"], strict=True
                    ):
                        lines = page["lines"]
                        other_text_flow_count += max(
                            0,
                            lines["missing"]
                            - referenced_by_page_and_side.get((page_number, "missing"), 0),
                        )
                        other_text_flow_count += max(
                            0,
                            lines["extra"]
                            - referenced_by_page_and_side.get((page_number, "extra"), 0),
                        )
                        other_text_flow_count += (
                            page["wraps"]["count"]
                            + page["reflow"]["gt_lines"]
                            + page["reflow"]["out_lines"]
                        )
                    if other_text_flow_count:
                        errors.append(
                            f"Visual audit > {field_name} has text-flow findings "
                            "outside the verified visibility differences and requires "
                            "an open issue."
                        )
                uncovered = current_findings - referenced_keys
                if uncovered:
                    finding_name = "visibility" if category == "text flow" else "shifts"
                    errors.append(
                        f"Visual audit > {field_name} has {finding_name} findings not "
                        "covered by exact reference exporter differences; use an "
                        "open issue."
                    )
        elif disposition != (set(), set()):
            errors.append(
                f"Visual audit > {field_name} has no findings and must be Pass."
            )

    return errors


def compared_pages(value: str | None) -> set[int] | None:
    """Parse `1`, `1, 3`, and `1-3` page declarations."""
    if not value:
        return None
    pages: set[int] = set()
    for token in value.split(","):
        token = token.strip()
        match = re.fullmatch(r"([1-9]\d*)(?:\s*-\s*([1-9]\d*))?", token)
        if match is None:
            return None
        start = int(match.group(1))
        end = int(match.group(2) or start)
        if end < start:
            return None
        pages.update(range(start, end + 1))
    return pages or None


def _render_cluster_report_paths(value: str | None) -> list[str] | None:
    if not value:
        return None
    paths = re.findall(
        r"assets/bugfixes/issue-\d+/render-clusters-page-\d+\.json", value
    )
    residue = re.sub(
        r"assets/bugfixes/issue-\d+/render-clusters-page-\d+\.json", "", value
    )
    residue = residue.replace("`", "").replace(",", "").strip()
    if not paths or residue:
        return None
    return paths


def _validate_render_cluster_report(
    report: object,
    *,
    path: str,
    expected_page: int,
    remaining_issues: set[int],
    reference_differences: dict[str, object] | None = None,
) -> list[str]:
    errors: list[str] = []
    if not isinstance(report, dict):
        return [f"{path}: render cluster report must be an object."]
    if report.get("schema_version") != 1:
        errors.append(f"{path}: schema_version must be 1.")
    if report.get("page") != expected_page:
        errors.append(f"{path}: report page must be {expected_page}.")
    dpi = report.get("dpi")
    if isinstance(dpi, bool) or not isinstance(dpi, (int, float)) or dpi < 150:
        errors.append(f"{path}: report DPI must be at least 150.")
    if report.get("fuzz_percent") != 5:
        errors.append(f"{path}: report must use the 5% fuzz cluster sweep.")
    minimum_area = report.get("minimum_area_pt2")
    if (
        isinstance(minimum_area, bool)
        or not isinstance(minimum_area, (int, float))
        or minimum_area <= 0
    ):
        errors.append(f"{path}: minimum_area_pt2 must be positive.")
        minimum_area = 0.0
    if report.get("strict") is not True:
        errors.append(f"{path}: strict must be true.")
    if report.get("passed") is not True:
        errors.append(f"{path}: strict cluster audit did not pass.")

    list_fields = (
        "undispositioned_cluster_ids",
        "unknown_disposition_cluster_ids",
        "duplicate_disposition_cluster_ids",
        "errors",
    )
    for field_name in list_fields:
        value = report.get(field_name)
        if not isinstance(value, list):
            errors.append(f"{path}: {field_name} must be a list.")
        elif value:
            detail = ", ".join(str(item) for item in value)
            errors.append(f"{path}: {field_name} must be empty, got {detail}.")

    renderer_observations = report.get("renderer_observations")
    if not isinstance(renderer_observations, list):
        errors.append(f"{path}: renderer_observations must be a list.")
    else:
        for index, observation in enumerate(renderer_observations, start=1):
            if not isinstance(observation, dict):
                errors.append(f"{path}: renderer observation {index} must be an object.")
                continue
            if observation.get("class") not in ACCEPTED_RENDERING_CLASSES:
                errors.append(
                    f"{path}: renderer observation {index} has an unsupported class."
                )
            note = observation.get("note")
            if not isinstance(note, str) or not note.strip():
                errors.append(
                    f"{path}: renderer observation {index} needs an inspection note."
                )
            bbox = observation.get("bbox_pt")
            if not isinstance(bbox, dict) or set(bbox) != {
                "x",
                "y",
                "width",
                "height",
            }:
                errors.append(
                    f"{path}: renderer observation {index} needs an exact bounded bbox_pt."
                )
                continue
            values = [bbox[name] for name in ("x", "y", "width", "height")]
            if any(
                isinstance(value, bool) or not isinstance(value, (int, float))
                for value in values
            ) or bbox["x"] < 0 or bbox["y"] < 0 or bbox["width"] <= 0 or bbox["height"] <= 0:
                errors.append(
                    f"{path}: renderer observation {index} bbox_pt values are invalid."
                )

    clusters = report.get("clusters")
    if not isinstance(clusters, list):
        errors.append(f"{path}: clusters must be a list.")
        return errors

    seen_ids: set[str] = set()
    issue_dispositions: set[int] = set()
    reference_registry, reference_errors = (
        validate_reference_difference_document(reference_differences)
        if reference_differences is not None
        else ({}, [])
    )
    errors.extend(f"{path}: reference exporter differences: {error}." for error in reference_errors)
    referenced_clusters: dict[str, set[str]] = {}
    dispositioned_count = 0
    for index, cluster in enumerate(clusters, start=1):
        if not isinstance(cluster, dict):
            errors.append(f"{path}: cluster {index} must be an object.")
            continue
        cluster_id = cluster.get("id")
        if not isinstance(cluster_id, str) or re.fullmatch(
            rf"p{expected_page}-[0-9a-f]{{12}}", cluster_id
        ) is None:
            errors.append(f"{path}: cluster {index} has an invalid stable ID.")
            cluster_id = f"cluster {index}"
        elif cluster_id in seen_ids:
            errors.append(f"{path}: duplicate cluster ID {cluster_id}.")
        else:
            seen_ids.add(cluster_id)

        bbox = cluster.get("bbox_pt")
        bbox_is_valid = False
        if not isinstance(bbox, dict) or any(
            isinstance(bbox.get(name), bool)
            or not isinstance(bbox.get(name), (int, float))
            for name in ("x", "y", "width", "height")
        ):
            errors.append(f"{path}: {cluster_id} has an invalid bbox_pt.")
        elif bbox["width"] <= 0 or bbox["height"] <= 0:
            errors.append(f"{path}: {cluster_id} bbox dimensions must be positive.")
        else:
            bbox_is_valid = True
        area = cluster.get("area_pt2")
        area_is_valid = True
        if (
            isinstance(area, bool)
            or not isinstance(area, (int, float))
            or area < minimum_area
        ):
            errors.append(
                f"{path}: {cluster_id} area_pt2 must meet the report minimum."
            )
            area_is_valid = False
        if (
            isinstance(cluster_id, str)
            and cluster_id in seen_ids
            and bbox_is_valid
            and area_is_valid
        ):
            canonical = "|".join(
                (
                    str(expected_page),
                    f"{bbox['x']:.2f}",
                    f"{bbox['y']:.2f}",
                    f"{bbox['width']:.2f}",
                    f"{bbox['height']:.2f}",
                    f"{area:.2f}",
                )
            )
            expected_id = (
                f"p{expected_page}-"
                f"{hashlib.sha256(canonical.encode('ascii')).hexdigest()[:12]}"
            )
            if cluster_id != expected_id:
                errors.append(
                    f"{path}: {cluster_id} does not match its page/bbox/area stable ID "
                    f"{expected_id}."
                )

        disposition = cluster.get("disposition")
        if not isinstance(disposition, dict):
            errors.append(f"{path}: {cluster_id} is undispositioned.")
            continue
        if disposition.get("kind") == "accepted-rendering":
            if set(disposition) - {"kind", "class", "note"}:
                errors.append(
                    f"{path}: {cluster_id} accepted-rendering disposition has "
                    "unsupported fields."
                )
                continue
            if disposition.get("class") not in ACCEPTED_RENDERING_CLASSES:
                errors.append(
                    f"{path}: {cluster_id} has an unsupported accepted-rendering class."
                )
                continue
        elif disposition.get("kind") == "issue":
            if set(disposition) - {"kind", "issue", "note"}:
                errors.append(
                    f"{path}: {cluster_id} issue disposition has unsupported fields."
                )
                continue
            issue = disposition.get("issue")
            issue_match = re.fullmatch(r"#([1-9]\d*)", issue or "")
            if issue_match is None:
                errors.append(f"{path}: {cluster_id} has an invalid issue disposition.")
                continue
            issue_dispositions.add(int(issue_match.group(1)))
        elif disposition.get("kind") == "reference-exporter-difference":
            if set(disposition) - {"kind", "difference_id", "note"}:
                errors.append(
                    f"{path}: {cluster_id} reference exporter disposition has "
                    "unsupported fields."
                )
                continue
            difference_id = disposition.get("difference_id")
            difference = reference_registry.get(difference_id)
            if not isinstance(difference_id, str) or difference is None:
                errors.append(
                    f"{path}: {cluster_id} reference exporter disposition does not "
                    "name validated evidence."
                )
                continue
            if (
                difference.get("kind") != "render-clusters"
                or difference.get("page") != expected_page
            ):
                errors.append(
                    f"{path}: {cluster_id} reference exporter disposition has the "
                    "wrong kind or page."
                )
                continue
            declared_ids = difference.get("render_cluster_ids", [])
            if cluster_id not in declared_ids:
                errors.append(
                    f"{path}: {cluster_id} is not listed by reference difference "
                    f"{difference_id}."
                )
                continue
            referenced_clusters.setdefault(difference_id, set()).add(cluster_id)
        else:
            errors.append(f"{path}: {cluster_id} has an invalid disposition kind.")
            continue
        dispositioned_count += 1

    summary = report.get("summary")
    expected_summary = {
        "total": len(clusters),
        "dispositioned": dispositioned_count,
        "undispositioned": 0,
        "unknown": 0,
        "duplicate": 0,
    }
    if not isinstance(summary, dict) or any(
        summary.get(name) != count for name, count in expected_summary.items()
    ):
        errors.append(f"{path}: summary does not match the validated clusters.")

    unclassified_issues = issue_dispositions - remaining_issues
    if unclassified_issues:
        references = ", ".join(f"#{number}" for number in sorted(unclassified_issues))
        errors.append(
            f"{path}: issue dispositions must also appear in a Remaining deviation row: "
            f"{references}."
        )
    for difference_id, used_ids in sorted(referenced_clusters.items()):
        declared_ids = set(reference_registry[difference_id]["render_cluster_ids"])
        if used_ids != declared_ids:
            errors.append(
                f"{path}: reference difference {difference_id} must disposition its "
                "exact declared cluster set."
            )
    return errors


def validate_render_cluster_audits(
    body: str,
    changed_paths: list[str],
    root: Path,
    *,
    reference_differences: dict[str, object] | None = None,
) -> list[str]:
    """Require a fresh passing strict cluster report for every visual-fix page."""
    impact = extract_section(body, "## Visual impact")
    if checked(impact, "No rendered PDF change"):
        return []
    audit = extract_section(body, "## Visual audit")
    if field(audit, "Evidence mode") != "fix":
        return []
    issue_match = re.fullmatch(r"#(\d+)", field(audit, "Issue") or "")
    if issue_match is None:
        return []
    issue_number = issue_match.group(1)
    pages = compared_pages(field(audit, "Page(s)"))
    if pages is None:
        return ["Visual audit > Page(s) must use page numbers or ranges such as `1, 3-5`."]

    declared_paths = _render_cluster_report_paths(field(audit, "Render cluster reports"))
    if declared_paths is None:
        return [
            "Visual audit > Render cluster reports must list one strict report path per page."
        ]
    expected_paths = {
        f"assets/bugfixes/issue-{issue_number}/render-clusters-page-{page}.json": page
        for page in pages
    }
    errors: list[str] = []
    declared_set = set(declared_paths)
    missing_declarations = set(expected_paths) - declared_set
    extra_declarations = declared_set - set(expected_paths)
    for path in sorted(missing_declarations):
        errors.append(f"Visual audit > Render cluster reports is missing page {expected_paths[path]}.")
    for path in sorted(extra_declarations):
        errors.append(f"Visual audit > Render cluster reports has an unexpected path: {path}.")

    remaining_issues = remaining_issue_numbers(body)
    for path, page in expected_paths.items():
        if path not in changed_paths:
            errors.append(
                f"{path}: strict render cluster report for page {page} must be changed "
                "in this pull request."
            )
            continue
        try:
            report = json.loads((root / path).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(f"{path}: invalid render cluster report: {exc}.")
            continue
        errors.extend(
            _validate_render_cluster_report(
                report,
                path=path,
                expected_page=page,
                remaining_issues=remaining_issues,
                reference_differences=reference_differences,
            )
        )
    return errors


def validate_open_issues(issue_numbers: set[int], repository: str, token: str | None) -> list[str]:
    if not issue_numbers:
        return []
    if not token:
        return ["GITHUB_TOKEN is required to verify remaining visual issues."]

    errors: list[str] = []
    for number in sorted(issue_numbers):
        url = f"https://api.github.com/repos/{repository}/issues/{number}"
        request = urllib.request.Request(
            url,
            headers={
                "Accept": "application/vnd.github+json",
                "Authorization": f"Bearer {token}",
                "User-Agent": "office2pdf-visual-pr-gate",
                "X-GitHub-Api-Version": "2022-11-28",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=15) as response:
                issue = json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            errors.append(f"Could not verify remaining visual issue #{number}: {exc}.")
            continue
        if "pull_request" in issue:
            errors.append(f"Remaining reference #{number} is a pull request, not an issue.")
        elif issue.get("state") != "open":
            errors.append(f"Remaining visual issue #{number} is not open.")
    return errors


def read_jpeg_info(path: Path) -> JpegInfo:
    data = path.read_bytes()
    if not data.startswith(b"\xff\xd8"):
        raise ValueError("not a JPEG file")

    offset = 2
    width = height = None
    progressive = False
    density: tuple[float, float] | None = None
    metadata: list[str] = []

    while offset + 1 < len(data):
        if data[offset] != 0xFF:
            offset += 1
            continue
        while offset < len(data) and data[offset] == 0xFF:
            offset += 1
        if offset >= len(data):
            break
        marker = data[offset]
        offset += 1
        if marker in {0xD8, 0xD9, 0x01} or 0xD0 <= marker <= 0xD7:
            continue
        if offset + 2 > len(data):
            raise ValueError("truncated JPEG segment")
        length = struct.unpack(">H", data[offset : offset + 2])[0]
        if length < 2 or offset + length > len(data):
            raise ValueError("invalid JPEG segment length")
        payload = data[offset + 2 : offset + length]
        offset += length

        if marker == 0xE0 and payload.startswith(b"JFIF\x00") and len(payload) >= 12:
            units = payload[7]
            x_density = struct.unpack(">H", payload[8:10])[0]
            y_density = struct.unpack(">H", payload[10:12])[0]
            if units == 1:
                density = (float(x_density), float(y_density))
            elif units == 2:
                density = (x_density * 2.54, y_density * 2.54)
        elif marker in {0xE1, 0xE2, 0xED, 0xFE}:
            metadata.append({0xE1: "APP1", 0xE2: "APP2", 0xED: "APP13", 0xFE: "COM"}[marker])
        elif marker in {0xC0, 0xC1, 0xC2} and len(payload) >= 5:
            height = struct.unpack(">H", payload[1:3])[0]
            width = struct.unpack(">H", payload[3:5])[0]
            progressive = marker == 0xC2
        elif marker == 0xDA:
            break

    if width is None or height is None:
        raise ValueError("JPEG dimensions were not found")
    return JpegInfo(width, height, progressive, density, tuple(metadata))


def validate_jpeg(path: Path) -> list[str]:
    try:
        info = read_jpeg_info(path)
    except (OSError, ValueError) as exc:
        return [f"{path}: {exc}"]

    errors: list[str] = []
    if not info.progressive:
        errors.append(f"{path}: evidence must be a progressive JPEG.")
    if not info.density_dpi or min(info.density_dpi) < 150:
        errors.append(f"{path}: JPEG density must be at least 150 DPI.")
    if info.metadata_markers:
        markers = ", ".join(info.metadata_markers)
        errors.append(f"{path}: metadata was not stripped ({markers}).")
    if path.name == "compare.jpg" and info.width % 2:
        errors.append(f"{path}: side-by-side comparison width must split into equal panels.")
    return errors


def validate_evidence(changed_paths: list[str], root: Path) -> list[str]:
    errors: list[str] = []
    touched: dict[str, set[str]] = {}

    for raw_path in changed_paths:
        if (
            LAYOUT_AUDIT_PATH.fullmatch(raw_path)
            or RENDER_CLUSTER_REPORT_PATH.fullmatch(raw_path)
            or REFERENCE_DIFFERENCE_PATH.fullmatch(raw_path)
        ):
            continue
        if not is_visual_asset(raw_path):
            continue
        match = EVIDENCE_PATH.fullmatch(raw_path)
        if not match:
            errors.append(
                f"{raw_path}: visual evidence must be gt.jpg, before.jpg, after.jpg, "
                "compare.jpg, or native.jpg under assets/bugfixes/issue-<number>/."
            )
            continue
        issue = match.group("issue")
        name = match.group("name")
        if match.group("ext").lower() != "jpg":
            errors.append(f"{raw_path}: visual evidence must use the .jpg extension.")
            continue
        touched.setdefault(issue, set()).add(name)

    for issue, names in touched.items():
        issue_dir = root / VISUAL_ROOT / f"issue-{issue}"
        required = {"gt", "before", "after"} if names & {"gt", "before", "after"} else set()
        standalone = ({"compare"} if "compare" in names else set()) | (
            {"native"} if "native" in names else set()
        )
        for name in sorted(required | standalone):
            path = issue_dir / f"{name}.jpg"
            if not path.is_file():
                errors.append(f"{path}: required evidence file is missing.")
            else:
                errors.extend(validate_jpeg(path))

    return errors


def git_changed_paths(base: str, head: str, root: Path) -> list[str]:
    result = subprocess.run(
        ["git", "diff", "--name-only", "--diff-filter=ACMR", base, head],
        cwd=root,
        check=True,
        capture_output=True,
        text=True,
    )
    return [line for line in result.stdout.splitlines() if line]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--event", type=Path, required=True)
    parser.add_argument("--base", required=True)
    parser.add_argument("--head", required=True)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    args = parser.parse_args()

    event = json.loads(args.event.read_text(encoding="utf-8"))
    body = event.get("pull_request", {}).get("body") or ""
    changed_paths = git_changed_paths(args.base, args.head, args.root)
    errors = validate_pr_body(body, changed_paths)
    errors.extend(validate_evidence(changed_paths, args.root))
    reference_differences, reference_errors = validate_reference_exporter_differences(
        body, changed_paths, args.root
    )
    errors.extend(reference_errors)
    errors.extend(
        validate_layout_audit(
            body,
            changed_paths,
            args.root,
            reference_differences=reference_differences,
        )
    )
    errors.extend(
        validate_render_cluster_audits(
            body,
            changed_paths,
            args.root,
            reference_differences=reference_differences,
        )
    )
    errors.extend(
        validate_open_issues(
            remaining_issue_numbers(body),
            args.repository,
            os.environ.get("GITHUB_TOKEN"),
        )
    )

    if errors:
        for error in errors:
            print(f"::error::{error}")
        return 1

    print("Visual pull request contract is complete.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
