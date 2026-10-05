#!/usr/bin/env python3
"""Validate evidence-backed differences for a supplied ground-truth export."""

from __future__ import annotations

import math
import re


REFERENCE_DIFFERENCE_SCHEMA_VERSION = 1
SHA256_RE = re.compile(r"[0-9a-f]{64}")
DIFFERENCE_ID_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*")
HTTPS_RE = re.compile(r"https://[^\s]+")
NATIVE_OFFICE_APPLICATIONS = frozenset(
    {"Microsoft Word", "Microsoft PowerPoint", "Microsoft Excel"}
)
VISIBILITY_STATES = frozenset({"hidden", "low_contrast", "painted"})


def _unexpected_fields(value: dict[str, object], allowed: set[str]) -> list[str]:
    return sorted(set(value) - allowed)


def _validate_sha(value: object, field: str, errors: list[str]) -> None:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        errors.append(f"{field} must be a lowercase 64-character SHA-256")


def _validate_nonempty(value: object, field: str, errors: list[str]) -> None:
    if not isinstance(value, str) or not value.strip():
        errors.append(f"{field} must be non-empty text")


def _is_finite_number(value: int | float) -> bool:
    try:
        return math.isfinite(value)
    except OverflowError:
        return False


def _validate_export(
    value: object, field: str, *, native: bool, errors: list[str]
) -> None:
    if not isinstance(value, dict):
        errors.append(f"{field} must be an object")
        return
    allowed = {
        "application",
        "version",
        "platform",
        "pdf_sha256",
        "evidence_path",
        "evidence_sha256",
    }
    extra = _unexpected_fields(value, allowed)
    if extra:
        errors.append(f"{field} has unsupported fields: {', '.join(extra)}")
    missing = sorted(allowed - set(value))
    if missing:
        errors.append(f"{field} is missing fields: {', '.join(missing)}")
    for name in ("application", "version", "platform", "evidence_path"):
        _validate_nonempty(value.get(name), f"{field}.{name}", errors)
    if native and value.get("application") not in NATIVE_OFFICE_APPLICATIONS:
        allowed_apps = ", ".join(sorted(NATIVE_OFFICE_APPLICATIONS))
        errors.append(f"{field}.application must be one of: {allowed_apps}")
    if not native and value.get("application") in NATIVE_OFFICE_APPLICATIONS:
        errors.append(f"{field}.application must identify the non-native GT exporter")
    _validate_sha(value.get("pdf_sha256"), f"{field}.pdf_sha256", errors)
    _validate_sha(value.get("evidence_sha256"), f"{field}.evidence_sha256", errors)


def validate_reference_difference_document(
    document: object,
) -> tuple[dict[str, dict[str, object]], list[str]]:
    """Return validated differences by ID plus strict schema errors.

    The structured provenance is deliberately part of the same document as the
    exact finding selectors. A note or issue comment by itself cannot create an
    accepted disposition.
    """

    if not isinstance(document, dict):
        return {}, ["reference exporter difference document must be an object"]

    errors: list[str] = []
    allowed_top = {
        "schema_version",
        "source",
        "reference_export",
        "native_export",
        "verification_url",
        "differences",
    }
    extra = _unexpected_fields(document, allowed_top)
    if extra:
        errors.append(
            "reference exporter difference document has unsupported fields: "
            + ", ".join(extra)
        )
    if document.get("schema_version") != REFERENCE_DIFFERENCE_SCHEMA_VERSION:
        errors.append(
            "reference exporter difference schema_version must be "
            f"{REFERENCE_DIFFERENCE_SCHEMA_VERSION}"
        )

    source = document.get("source")
    if not isinstance(source, dict):
        errors.append("source must be an object")
    else:
        extra_source = _unexpected_fields(source, {"url", "sha256"})
        if extra_source:
            errors.append(f"source has unsupported fields: {', '.join(extra_source)}")
        if set(source) != {"url", "sha256"}:
            errors.append("source must contain exactly url and sha256")
        url = source.get("url")
        if not isinstance(url, str) or HTTPS_RE.fullmatch(url) is None:
            errors.append("source.url must be an HTTPS URL")
        _validate_sha(source.get("sha256"), "source.sha256", errors)

    if document.get("reference_export") is not None:
        _validate_export(
            document["reference_export"],
            "reference_export",
            native=False,
            errors=errors,
        )
    _validate_export(
        document.get("native_export"),
        "native_export",
        native=True,
        errors=errors,
    )

    if isinstance(source, dict) and isinstance(source.get("url"), str):
        source_path = source["url"].split("?", 1)[0].lower()
        expected_application = next(
            (
                application
                for suffix, application in (
                    (".docx", "Microsoft Word"),
                    (".pptx", "Microsoft PowerPoint"),
                    (".xlsx", "Microsoft Excel"),
                )
                if source_path.endswith(suffix)
            ),
            None,
        )
        native_export = document.get("native_export")
        if (
            expected_application is not None
            and isinstance(native_export, dict)
            and native_export.get("application") != expected_application
        ):
            errors.append(
                f"native_export.application must be {expected_application} for the "
                "source file type"
            )

    verification_url = document.get("verification_url")
    if (
        not isinstance(verification_url, str)
        or HTTPS_RE.fullmatch(verification_url) is None
        or re.search(r"/(?:issues|pull)/\d+#issuecomment-\d+$", verification_url) is None
    ):
        errors.append(
            "verification_url must be an HTTPS GitHub issue or pull-request comment URL"
        )

    raw_differences = document.get("differences")
    if not isinstance(raw_differences, list) or not raw_differences:
        errors.append("differences must be a non-empty list")
        raw_differences = []

    registry: dict[str, dict[str, object]] = {}
    shift_selectors: set[tuple[int, str, int]] = set()
    rasterized_text_selectors: set[tuple[int, str, str, int]] = set()
    native_ground_truth = document.get("reference_export") is None
    for index, difference in enumerate(raw_differences, start=1):
        prefix = f"difference {index}"
        if not isinstance(difference, dict):
            errors.append(f"{prefix} must be an object")
            continue
        difference_id = difference.get("id")
        if (
            not isinstance(difference_id, str)
            or DIFFERENCE_ID_RE.fullmatch(difference_id) is None
        ):
            errors.append(f"{prefix} id must be a lowercase hyphenated identifier")
            continue
        if difference_id in registry:
            errors.append(f"difference id appears more than once: {difference_id}")
            continue
        page = difference.get("page")
        if type(page) is not int or page <= 0:
            errors.append(f"{prefix} page must be a positive integer")

        kind = difference.get("kind")
        if native_ground_truth and kind != "rasterized-text":
            errors.append(
                f"{prefix} kind must be rasterized-text when reference_export is null"
            )
        if kind == "painted-text-visibility":
            allowed = {"id", "page", "kind", "layout_finding"}
            extra_difference = _unexpected_fields(difference, allowed)
            if extra_difference:
                errors.append(
                    f"{prefix} has unsupported fields for {kind}: "
                    + ", ".join(extra_difference)
                )
            finding = difference.get("layout_finding")
            if not isinstance(finding, dict):
                errors.append(f"{prefix} layout_finding must be an object")
            else:
                expected_fields = {"label", "gt", "out", "occurrence"}
                if set(finding) != expected_fields:
                    errors.append(
                        f"{prefix} layout_finding must contain exactly label, gt, out, "
                        "and occurrence"
                    )
                _validate_nonempty(finding.get("label"), f"{prefix} layout_finding.label", errors)
                if finding.get("gt") not in VISIBILITY_STATES:
                    errors.append(f"{prefix} layout_finding.gt has an invalid visibility state")
                if finding.get("out") not in VISIBILITY_STATES:
                    errors.append(f"{prefix} layout_finding.out has an invalid visibility state")
                if finding.get("gt") == finding.get("out"):
                    errors.append(f"{prefix} layout_finding must describe different states")
                if type(finding.get("occurrence")) is not int or finding["occurrence"] <= 0:
                    errors.append(f"{prefix} layout_finding.occurrence must be positive")
        elif kind == "render-clusters":
            allowed = {"id", "page", "kind", "render_cluster_ids"}
            extra_difference = _unexpected_fields(difference, allowed)
            if extra_difference:
                errors.append(
                    f"{prefix} has unsupported fields for {kind}: "
                    + ", ".join(extra_difference)
                )
            cluster_ids = difference.get("render_cluster_ids")
            if (
                not isinstance(cluster_ids, list)
                or not cluster_ids
                or any(
                    not isinstance(cluster_id, str)
                    or re.fullmatch(rf"p{page}-[0-9a-f]{{12}}", cluster_id) is None
                    for cluster_id in cluster_ids
                )
            ):
                errors.append(
                    f"{prefix} render_cluster_ids must be non-empty stable IDs for page {page}"
                )
            elif len(cluster_ids) != len(set(cluster_ids)):
                errors.append(f"{prefix} render_cluster_ids must be unique")
        elif kind == "text-shift":
            allowed = {"id", "page", "kind", "layout_finding"}
            extra_difference = _unexpected_fields(difference, allowed)
            if extra_difference:
                errors.append(
                    f"{prefix} has unsupported fields for {kind}: "
                    + ", ".join(extra_difference)
                )
            finding = difference.get("layout_finding")
            if not isinstance(finding, dict):
                errors.append(f"{prefix} layout_finding must be an object")
            else:
                expected_fields = {"label", "occurrence", "dx", "dy"}
                if set(finding) != expected_fields:
                    errors.append(
                        f"{prefix} layout_finding must contain exactly label, "
                        "occurrence, dx, and dy"
                    )
                label = finding.get("label")
                _validate_nonempty(label, f"{prefix} layout_finding.label", errors)
                occurrence = finding.get("occurrence")
                if type(occurrence) is not int or occurrence <= 0:
                    errors.append(
                        f"{prefix} layout_finding.occurrence must be positive"
                    )
                if (
                    type(page) is int
                    and page > 0
                    and isinstance(label, str)
                    and label.strip()
                    and type(occurrence) is int
                    and occurrence > 0
                ):
                    selector = (page, label, occurrence)
                    if selector in shift_selectors:
                        errors.append(
                            f"{prefix} text-shift selector appears more than once"
                        )
                    shift_selectors.add(selector)
                for coordinate in ("dx", "dy"):
                    measurement = finding.get(coordinate)
                    if isinstance(measurement, bool) or not isinstance(
                        measurement, (int, float)
                    ):
                        errors.append(
                            f"{prefix} layout_finding.{coordinate} must be numeric"
                        )
                    elif not _is_finite_number(measurement):
                        errors.append(
                            f"{prefix} layout_finding.{coordinate} must be finite"
                        )
        elif kind == "rasterized-text":
            allowed = {"id", "page", "kind", "layout_finding"}
            extra_difference = _unexpected_fields(difference, allowed)
            if extra_difference:
                errors.append(
                    f"{prefix} has unsupported fields for {kind}: "
                    + ", ".join(extra_difference)
                )
            finding = difference.get("layout_finding")
            if not isinstance(finding, dict):
                errors.append(f"{prefix} layout_finding must be an object")
            else:
                expected_fields = {"label", "side", "occurrence"}
                if set(finding) != expected_fields:
                    errors.append(
                        f"{prefix} layout_finding must contain exactly label, side, "
                        "and occurrence"
                    )
                label = finding.get("label")
                side = finding.get("side")
                occurrence = finding.get("occurrence")
                _validate_nonempty(label, f"{prefix} layout_finding.label", errors)
                if not isinstance(side, str) or side not in {"gt", "out"}:
                    errors.append(
                        f"{prefix} layout_finding.side must be gt or out"
                    )
                if type(occurrence) is not int or occurrence <= 0:
                    errors.append(
                        f"{prefix} layout_finding.occurrence must be positive"
                    )
                if (
                    type(page) is int
                    and page > 0
                    and isinstance(label, str)
                    and label.strip()
                    and isinstance(side, str)
                    and side in {"gt", "out"}
                    and type(occurrence) is int
                    and occurrence > 0
                ):
                    selector = (page, label, side, occurrence)
                    if selector in rasterized_text_selectors:
                        errors.append(
                            f"{prefix} rasterized-text selector appears more than once"
                        )
                    rasterized_text_selectors.add(selector)
        else:
            errors.append(
                f"{prefix} kind must be painted-text-visibility, render-clusters, "
                "text-shift, or rasterized-text"
            )

        registry[difference_id] = difference

    return registry, errors
