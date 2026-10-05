import io
import hashlib
import json
import re
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.check_visual_pr import (
    AUDIT_ROWS,
    INSPECTION_ITEMS,
    decoded_pixel_delta,
    read_jpeg_info,
    validate_evidence,
    validate_layout_audit,
    validate_open_issues,
    validate_pr_body,
    validate_reference_exporter_differences,
    validate_render_cluster_audits,
)
from scripts.reference_exporter_differences import (
    validate_reference_difference_document,
)


ROOT = Path(__file__).resolve().parents[2]


def visual_body(result_overrides=None, include_previews=True):
    results = {row: "Matches GT" for row in AUDIT_ROWS}
    results.update(result_overrides or {})
    follow_ups = sorted(
        {number for result in results.values() for number in re.findall(r"#\d+", result)}
    )
    follow_up_value = ", ".join(follow_ups) if follow_ups else "None"
    inspections = "\n".join(f"- [x] {item}" for item in INSPECTION_ITEMS)
    rows = "\n".join(f"| {row} | {results[row]} |" for row in AUDIT_ROWS)
    previews = (
        """### Visual comparison

| GT | Before | After |
| --- | --- | --- |
| ![GT](https://example.com/gt.jpg) | ![Before](https://example.com/before.jpg) | ![After](https://example.com/after.jpg) |

"""
        if include_previews
        else ""
    )
    return f"""## Visual impact

- [ ] No rendered PDF change
- [x] Rendered PDF change or visual evidence added
- Reason:

## Visual audit

- Issue: #186
- Fixture: tests/fixtures/xlsx/pr_186_contributor_acceptance.xlsx
- Page(s): 1
- Renderer and DPI: pdftoppm, 150 DPI
- Evidence mode: `fix`
- Layout audit report: `assets/bugfixes/issue-186/layout-audit.json`
- Render cluster reports: `assets/bugfixes/issue-186/render-clusters-page-1.json`
- Reference exporter differences: None
- Fine-detail threshold: 0.5pt
- Layout audit page count: Pass
- Layout audit text flow: Pass
- Layout audit visible fills: Pass
- Layout audit rectangle geometry: Pass
- Layout audit large shifts: Pass
- Layout audit fine shifts: Pass
- New follow-up issues found in this audit: {follow_up_value}
- Model vision findings: Full pages, pixel diff, and matched crops were opened; no untracked visual deviation remains.
- GT: `assets/bugfixes/issue-186/gt.jpg`
- Before: `assets/bugfixes/issue-186/before.jpg`
- After: `assets/bugfixes/issue-186/after.jpg`
- Native: None

{previews}
### Required inspection

{inspections}

### Deviation audit

| Check | Result |
| --- | --- |
{rows}
"""


def layout_report(
    *,
    gt_pages: int = 1,
    out_pages: int = 1,
    missing: int = 0,
    extra: int = 0,
    wraps: int = 0,
    reflow_gt: int = 0,
    reflow_out: int = 0,
    large_shifts: int = 0,
    fine_shifts: int = 0,
    fine_threshold: float | None = 0.5,
    visibility: int = 0,
    visible_fills: int = 0,
    rect_geometry: int = 0,
    large_shift_entries: list[dict[str, object]] | None = None,
    fine_shift_entries: list[dict[str, object]] | None = None,
    missing_text: list[str] | None = None,
    extra_text: list[str] | None = None,
) -> dict[str, object]:
    if missing_text is None:
        missing_text = [f"missing line {index}" for index in range(1, min(missing, 5) + 1)]
    if extra_text is None:
        extra_text = [f"extra line {index}" for index in range(1, min(extra, 5) + 1)]
    return {
        "pages": [
            {
                "lines": {
                    "missing": missing,
                    "extra": extra,
                    "missing_text": missing_text,
                    "extra_text": extra_text,
                },
                "wraps": {"count": wraps},
                "reflow": {"gt_lines": reflow_gt, "out_lines": reflow_out},
                "instances": {
                    "large_shift_count": large_shifts,
                    "large_shifts": large_shift_entries or [],
                    "fine_shift_count": fine_shifts,
                    "fine_shifts": fine_shift_entries or [],
                    "fine_shift_threshold": fine_threshold,
                },
                "visibility": {"mismatch_count": visibility},
                "visible_fills": {"mismatch_count": visible_fills},
                "rects": {"geometry_mismatch_count": rect_geometry},
            }
        ],
        "gt_pages": gt_pages,
        "out_pages": out_pages,
    }


def validate_report(body, report, reference_differences=None):
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        report_path = root / "assets/bugfixes/issue-186/layout-audit.json"
        report_path.parent.mkdir(parents=True)
        report_path.write_text(json.dumps(report), encoding="utf-8")
        return validate_layout_audit(
            body,
            [
                "assets/bugfixes/issue-186/after.jpg",
                "assets/bugfixes/issue-186/layout-audit.json",
            ],
            root,
            reference_differences=reference_differences,
        )


def render_cluster_report(
    *,
    page=1,
    strict=True,
    passed=True,
    disposition=None,
    undispositioned=None,
    renderer_observations=None,
):
    disposition = disposition or {
        "kind": "accepted-rendering",
        "class": "glyph-edge-rasterization",
        "note": "Inspected at full resolution.",
    }
    undispositioned = undispositioned or []
    canonical_id_input = f"{page}|10.00|20.00|8.00|8.00|40.00"
    cluster_id = f"p{page}-{hashlib.sha256(canonical_id_input.encode('ascii')).hexdigest()[:12]}"
    cluster = {
        "id": cluster_id,
        "bbox_pt": {"x": 10.0, "y": 20.0, "width": 8.0, "height": 8.0},
        "area_pt2": 40.0,
        "region": "top-left",
        "disposition": disposition,
    }
    return {
        "schema_version": 1,
        "page": page,
        "dpi": 300,
        "fuzz_percent": 5,
        "minimum_area_pt2": 20.0,
        "strict": strict,
        "clusters": [cluster],
        "undispositioned_cluster_ids": undispositioned,
        "unknown_disposition_cluster_ids": [],
        "duplicate_disposition_cluster_ids": [],
        "errors": [],
        "renderer_observations": renderer_observations or [],
        "summary": {
            "total": 1,
            "dispositioned": 0 if undispositioned else 1,
            "undispositioned": len(undispositioned),
            "unknown": 0,
            "duplicate": 0,
        },
        "passed": passed,
    }


def validate_cluster_reports(body, reports, reference_differences=None):
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        issue_dir = root / "assets/bugfixes/issue-186"
        issue_dir.mkdir(parents=True)
        changed_paths = []
        for page, report in reports.items():
            relative = f"assets/bugfixes/issue-186/render-clusters-page-{page}.json"
            (root / relative).write_text(json.dumps(report), encoding="utf-8")
            changed_paths.append(relative)
        return validate_render_cluster_audits(
            body,
            changed_paths,
            root,
            reference_differences=reference_differences,
        )


def reference_difference_document(
    *,
    cluster_ids: list[str] | None = None,
    shift_differences: list[dict[str, object]] | None = None,
    native_ground_truth: bool = False,
    rasterized_text_differences: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    differences = []
    if not native_ground_truth:
        differences.append(
            {
                "id": "page-9-slide-number-visibility",
                "page": 9,
                "kind": "painted-text-visibility",
                "layout_finding": {
                    "label": "9",
                    "gt": "hidden",
                    "out": "painted",
                    "occurrence": 1,
                },
            }
        )
    if cluster_ids is not None:
        differences.append(
            {
                "id": "page-1-title-native-match",
                "page": 1,
                "kind": "render-clusters",
                "render_cluster_ids": cluster_ids,
            }
        )
    differences.extend(shift_differences or [])
    differences.extend(rasterized_text_differences or [])
    return {
        "schema_version": 1,
        "source": {
            "url": "https://github.com/developer0hye/office2pdf/files/123/source.pptx",
            "sha256": "1" * 64,
        },
        "reference_export": (
            None
            if native_ground_truth
            else {
                "application": "LibreOffice Impress",
                "version": "26.2.5.2",
                "platform": "macOS 26.6.2",
                "pdf_sha256": "2" * 64,
                "evidence_path": "assets/bugfixes/issue-186/gt.jpg",
                "evidence_sha256": "3" * 64,
            }
        ),
        "native_export": {
            "application": "Microsoft PowerPoint",
            "version": "16.112.3",
            "platform": "macOS 26.6.2 build 25G83",
            "pdf_sha256": "4" * 64,
            "evidence_path": "assets/bugfixes/issue-186/native.jpg",
            "evidence_sha256": "5" * 64,
        },
        "verification_url": (
            "https://github.com/developer0hye/office2pdf/issues/1421"
            "#issuecomment-5471464526"
        ),
        "differences": differences,
    }


def rasterized_text_body(difference_id: str, *, page: int = 13) -> str:
    native_path = "assets/bugfixes/issue-186/native.jpg"
    manifest_path = "assets/bugfixes/issue-186/reference-exporter-differences.json"
    body = visual_body(
        {"Element presence": f"Reference difference: ref:{difference_id}"}
    )
    return (
        body.replace("- Page(s): 1", f"- Page(s): {page}")
        .replace(
            "Reference exporter differences: None",
            f"Reference exporter differences: `{manifest_path}`",
        )
        .replace(
            "Layout audit text flow: Pass",
            f"Layout audit text flow: ref:{difference_id}",
        )
        .replace("- Native: None", f"- Native: `{native_path}`")
        + f"\n![Native](https://example.com/native.jpg)\n"
    )


def rasterized_text_difference(
    *,
    difference_id: str = "page-13-slide-number-rasterized",
    page: int = 13,
    label: str = "13",
    side: str = "gt",
    occurrence: int = 1,
) -> dict[str, object]:
    return {
        "id": difference_id,
        "page": page,
        "kind": "rasterized-text",
        "layout_finding": {
            "label": label,
            "side": side,
            "occurrence": occurrence,
        },
    }


def defect_body():
    return (
        visual_body()
        .replace("- Evidence mode: `fix`", "- Evidence mode: `defect`")
        .replace(
            """- GT: `assets/bugfixes/issue-186/gt.jpg`
- Before: `assets/bugfixes/issue-186/before.jpg`
- After: `assets/bugfixes/issue-186/after.jpg`""",
            "- Compare: `assets/bugfixes/issue-186/compare.jpg`",
        )
        .replace(
            """| GT | Before | After |
| --- | --- | --- |
| ![GT](https://example.com/gt.jpg) | ![Before](https://example.com/before.jpg) | ![After](https://example.com/after.jpg) |""",
            """| Compare |
| --- |
| ![Compare](https://example.com/compare.jpg) |""",
        )
    )


class PullRequestBodyTests(unittest.TestCase):
    def test_non_visual_pr_requires_reason(self):
        body = """## Visual impact

- [x] No rendered PDF change
- [ ] Rendered PDF change or visual evidence added
- Reason: Documentation-only workflow change
"""
        self.assertEqual(validate_pr_body(body, ["README.md"]), [])

    def test_complete_visual_audit_passes(self):
        errors = validate_pr_body(
            visual_body({"Fill": "Remaining: #328"}),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertEqual(errors, [])

    def test_visual_audit_requires_rendered_evidence_previews(self):
        errors = validate_pr_body(
            visual_body(include_previews=False),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertTrue(any("rendered preview" in error for error in errors))

    def test_visual_audit_requires_layout_finding_disposition(self):
        required = (
            "- [x] Ran compare_layout.py --audit --fine-shift PT and dispositioned every "
            "fine/large text-instance shift, rectangle geometry deviation, painted-text "
            "visibility mismatch, and visible-fill occlusion"
        )
        errors = validate_pr_body(
            visual_body().replace(required, required.replace("[x]", "[ ]")),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertTrue(any("painted-text visibility mismatch" in error for error in errors))

    def test_visual_audit_requires_complete_strict_cluster_command(self):
        required = (
            "- [x] Ran compare_render.py --cluster-report PATH --strict-clusters and "
            "dispositioned every material 5% fuzz diff cluster by explicit ID"
        )
        errors = validate_pr_body(
            visual_body().replace(required, required.replace("[x]", "[ ]")),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertTrue(any("--cluster-report PATH" in error for error in errors))

    def test_visual_audit_requires_model_vision_inspection(self):
        required = (
            "- [x] Used Codex/Claude vision to inspect the full GT/output pages, "
            "diff, and matched crops"
        )
        errors = validate_pr_body(
            visual_body().replace(required, required.replace("[x]", "[ ]")),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertTrue(any("Codex/Claude vision" in error for error in errors))

    def test_visual_audit_requires_substantive_model_vision_findings(self):
        body = visual_body().replace(
            "Model vision findings: Full pages, pixel diff, and matched crops were opened; "
            "no untracked visual deviation remains.",
            "Model vision findings: None",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("substantive direct inspection" in error for error in errors))

    def test_visual_audit_rejects_numeric_only_model_vision_findings(self):
        body = visual_body().replace(
            "Model vision findings: Full pages, pixel diff, and matched crops were opened; "
            "no untracked visual deviation remains.",
            "Model vision findings: 123456789012345678901234567890",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("numeric metrics alone" in error for error in errors))

    def test_visual_audit_requires_distinct_preview_urls(self):
        body = visual_body().replace(
            "https://example.com/before.jpg",
            "https://example.com/gt.jpg",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("different evidence image" in error for error in errors))

    def test_commented_preview_does_not_count_as_rendered(self):
        body = visual_body().replace(
            "![GT](https://example.com/gt.jpg)",
            "<!-- ![GT](https://example.com/gt.jpg) -->",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("rendered preview for GT" in error for error in errors))

    def test_backticked_preview_does_not_count_as_rendered(self):
        body = visual_body().replace(
            "![GT](https://example.com/gt.jpg)",
            "`![GT](https://example.com/gt.jpg)`",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("rendered preview for GT" in error for error in errors))

    def test_defect_audit_requires_only_rendered_compare_preview(self):
        errors = validate_pr_body(
            defect_body(),
            ["assets/bugfixes/issue-186/compare.jpg"],
        )
        self.assertEqual(errors, [])

    def test_remaining_deviation_requires_issue(self):
        errors = validate_pr_body(
            visual_body({"Fill": "Remaining: still different"}),
            ["assets/bugfixes/issue-186/after.jpg"],
        )
        self.assertTrue(any("Remaining deviation must reference an issue" in error for error in errors))

    def test_new_follow_up_must_classify_a_remaining_deviation(self):
        body = visual_body().replace(
            "New follow-up issues found in this audit: None",
            "New follow-up issues found in this audit: #328",
        )
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("must also classify a Remaining deviation" in error for error in errors))

    def test_visual_assets_cannot_be_marked_non_visual(self):
        body = """## Visual impact

- [x] No rendered PDF change
- [ ] Rendered PDF change or visual evidence added
- Reason: Documentation only
"""
        errors = validate_pr_body(body, ["assets/bugfixes/issue-186/after.jpg"])
        self.assertTrue(any("assets/bugfixes changes require" in error for error in errors))

    def test_layout_report_cannot_be_marked_non_visual(self):
        body = """## Visual impact

- [x] No rendered PDF change
- [ ] Rendered PDF change or visual evidence added
- Reason: Report only
"""
        errors = validate_pr_body(
            body,
            ["assets/bugfixes/issue-186/layout-audit.json"],
        )
        self.assertTrue(any("assets/bugfixes changes require" in error for error in errors))

    def test_evidence_documentation_is_not_a_rendered_change(self):
        body = """## Visual impact

- [x] No rendered PDF change
- [ ] Rendered PDF change or visual evidence added
- Reason: Documents the evidence convention
"""
        self.assertEqual(validate_pr_body(body, ["assets/bugfixes/README.md"]), [])

    def test_visual_audit_requires_reference_difference_declaration(self):
        body = visual_body().replace("- Reference exporter differences: None\n", "")

        errors = validate_pr_body(
            body, ["assets/bugfixes/issue-186/after.jpg"]
        )

        self.assertTrue(any("Reference exporter differences" in error for error in errors))

    def test_reference_difference_row_requires_an_exact_id(self):
        body = visual_body({"Element presence": "Reference difference: PowerPoint matched"})

        errors = validate_pr_body(
            body, ["assets/bugfixes/issue-186/after.jpg"]
        )

        self.assertTrue(any("exact ref:<id>" in error for error in errors))

    def test_native_evidence_cannot_be_added_without_a_difference_report(self):
        errors = validate_pr_body(
            visual_body(),
            [
                "assets/bugfixes/issue-186/after.jpg",
                "assets/bugfixes/issue-186/native.jpg",
            ],
        )

        self.assertTrue(any("native/reference-exporter evidence" in error for error in errors))


class ReferenceExporterDifferenceTests(unittest.TestCase):
    def prepare_native_ground_truth_evidence(
        self, root: Path, *, pixel_identical: bool = True
    ) -> tuple[str, dict[str, object]]:
        issue_dir = root / "assets/bugfixes/issue-186"
        issue_dir.mkdir(parents=True)
        ground_truth_image = ROOT / "assets/bugfixes/issue-1497/gt.jpg"
        native_image = (
            ground_truth_image
            if pixel_identical
            else ROOT / "assets/bugfixes/issue-1497/after.jpg"
        )
        shutil.copyfile(ground_truth_image, issue_dir / "gt.jpg")
        shutil.copyfile(native_image, issue_dir / "native.jpg")
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
        )
        document["native_export"]["evidence_sha256"] = hashlib.sha256(
            native_image.read_bytes()
        ).hexdigest()
        relative = "assets/bugfixes/issue-186/reference-exporter-differences.json"
        (root / relative).write_text(json.dumps(document), encoding="utf-8")
        return relative, document

    def test_decoded_pixel_delta_accepts_scientific_notation(self):
        with (
            patch("scripts.check_visual_pr.shutil.which", return_value="/usr/bin/magick"),
            patch("scripts.check_visual_pr.subprocess.run") as run,
        ):
            run.return_value.returncode = 1
            run.return_value.stderr = "6.83993e+06 (0.759992)"
            run.return_value.stdout = ""

            delta = decoded_pixel_delta(Path("gt.jpg"), Path("native.jpg"))

        self.assertEqual(delta, 6_839_930)

    def test_native_ground_truth_rasterized_text_selector_is_valid(self) -> None:
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
        )

        registry, errors = validate_reference_difference_document(document)

        self.assertEqual(errors, [])
        self.assertIsNone(document["reference_export"])
        self.assertEqual(
            registry["page-13-slide-number-rasterized"]["kind"], "rasterized-text"
        )

    def test_native_ground_truth_may_omit_reference_export(self) -> None:
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
        )
        document.pop("reference_export")

        registry, errors = validate_reference_difference_document(document)

        self.assertEqual(errors, [])
        self.assertEqual(
            registry["page-13-slide-number-rasterized"]["kind"], "rasterized-text"
        )

    def test_non_native_reference_can_record_rasterized_text(self) -> None:
        document = reference_difference_document(
            rasterized_text_differences=[rasterized_text_difference()]
        )

        registry, errors = validate_reference_difference_document(document)

        self.assertEqual(errors, [])
        self.assertEqual(
            registry["page-13-slide-number-rasterized"]["kind"], "rasterized-text"
        )

    def test_rasterized_text_selector_rejects_invalid_side_and_duplicate_occurrence(
        self,
    ) -> None:
        first = rasterized_text_difference()
        invalid_side = rasterized_text_difference(
            difference_id="page-13-slide-number-invalid-side", side="both"
        )
        duplicate = rasterized_text_difference(
            difference_id="page-13-slide-number-copy"
        )
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[first, invalid_side, duplicate],
        )

        _, errors = validate_reference_difference_document(document)

        self.assertTrue(any("side must be gt or out" in error for error in errors))
        self.assertTrue(any("selector appears more than once" in error for error in errors))

    def test_native_ground_truth_does_not_allow_other_exporter_differences(
        self,
    ) -> None:
        difference = {
            "id": "page-13-slide-number-shift",
            "page": 13,
            "kind": "text-shift",
            "layout_finding": {
                "label": "13",
                "occurrence": 1,
                "dx": 1.0,
                "dy": 0.0,
            },
        }
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
            shift_differences=[difference],
        )

        _, errors = validate_reference_difference_document(document)

        self.assertTrue(
            any("kind must be rasterized-text when reference_export is null" in error
                for error in errors)
        )

    def prepare_evidence(self, root: Path) -> tuple[str, dict[str, object]]:
        issue_dir = root / "assets/bugfixes/issue-186"
        issue_dir.mkdir(parents=True)
        reference_image = ROOT / "assets/bugfixes/issue-1497/gt.jpg"
        native_image = ROOT / "assets/bugfixes/issue-1497/after.jpg"
        shutil.copyfile(reference_image, issue_dir / "gt.jpg")
        shutil.copyfile(native_image, issue_dir / "native.jpg")
        reference_hash = hashlib.sha256(reference_image.read_bytes()).hexdigest()
        native_hash = hashlib.sha256(native_image.read_bytes()).hexdigest()
        document = reference_difference_document()
        document["reference_export"]["evidence_sha256"] = reference_hash
        document["native_export"]["evidence_sha256"] = native_hash
        relative = "assets/bugfixes/issue-186/reference-exporter-differences.json"
        (root / relative).write_text(json.dumps(document), encoding="utf-8")
        return relative, document

    def body(self) -> str:
        return (
            visual_body()
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "| Element presence | Matches GT |",
                "| Element presence | Reference difference: "
                "ref:page-9-slide-number-visibility |",
            )
            .replace(
                "- Native: None",
                "- Native: `assets/bugfixes/issue-186/native.jpg`",
            )
            + "\n![Native](https://example.com/native.jpg)\n"
        )

    def test_structured_source_reference_and_native_evidence_pass(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, expected = self.prepare_evidence(root)

            document, errors = validate_reference_exporter_differences(
                self.body(), [path, "assets/bugfixes/issue-186/native.jpg"], root
            )

        self.assertEqual(errors, [])
        self.assertEqual(document, expected)

    def test_free_form_provenance_cannot_bypass_the_gate(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, document = self.prepare_evidence(root)
            document["native_export"] = "PowerPoint matched when inspected"
            (root / path).write_text(json.dumps(document), encoding="utf-8")

            _, errors = validate_reference_exporter_differences(
                self.body(), [path], root
            )

        self.assertTrue(any("native_export must be an object" in error for error in errors))

    def test_native_evidence_hash_must_match_the_tracked_image(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, document = self.prepare_evidence(root)
            document["native_export"]["evidence_sha256"] = "0" * 64
            (root / path).write_text(json.dumps(document), encoding="utf-8")

            _, errors = validate_reference_exporter_differences(
                self.body(), [path], root
            )

        self.assertTrue(any("native_export evidence SHA-256" in error for error in errors))

    def test_native_ground_truth_accepts_pixel_identical_gt_and_native_evidence(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, expected = self.prepare_native_ground_truth_evidence(root)

            document, errors = validate_reference_exporter_differences(
                rasterized_text_body("page-13-slide-number-rasterized"),
                [
                    path,
                    "assets/bugfixes/issue-186/gt.jpg",
                    "assets/bugfixes/issue-186/native.jpg",
                ],
                root,
            )

        self.assertEqual(errors, [])
        self.assertEqual(document, expected)

    def test_native_ground_truth_rejects_different_gt_and_native_evidence(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, _ = self.prepare_native_ground_truth_evidence(
                root, pixel_identical=False
            )

            _, errors = validate_reference_exporter_differences(
                rasterized_text_body("page-13-slide-number-rasterized"),
                [path, "assets/bugfixes/issue-186/native.jpg"],
                root,
            )

        self.assertTrue(
            any("GT/native evidence must be pixel-identical" in error for error in errors)
        )

    def test_native_ground_truth_rasterized_text_passes_evidence_and_layout_gates(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, document = self.prepare_native_ground_truth_evidence(root)
            layout_path = "assets/bugfixes/issue-186/layout-audit.json"
            absolute_layout_path = root / layout_path
            absolute_layout_path.parent.mkdir(parents=True, exist_ok=True)
            absolute_layout_path.write_text(
                json.dumps(layout_report(extra=1, extra_text=["13"])),
                encoding="utf-8",
            )
            changed_paths = [
                manifest_path,
                "assets/bugfixes/issue-186/gt.jpg",
                "assets/bugfixes/issue-186/native.jpg",
                "assets/bugfixes/issue-186/after.jpg",
                layout_path,
            ]
            body = rasterized_text_body("page-13-slide-number-rasterized")

            _, evidence_errors = validate_reference_exporter_differences(
                body, changed_paths, root
            )
            layout_errors = validate_layout_audit(
                body,
                changed_paths,
                root,
                reference_differences=document,
            )

        self.assertEqual(evidence_errors, [])
        self.assertEqual(layout_errors, [])

    def test_reference_and_native_images_must_record_a_pixel_difference(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path, document = self.prepare_evidence(root)
            gt_path = root / "assets/bugfixes/issue-186/gt.jpg"
            native_path = root / "assets/bugfixes/issue-186/native.jpg"
            shutil.copyfile(gt_path, native_path)
            document["native_export"]["evidence_sha256"] = hashlib.sha256(
                native_path.read_bytes()
            ).hexdigest()
            (root / path).write_text(json.dumps(document), encoding="utf-8")

            _, errors = validate_reference_exporter_differences(
                self.body(), [path], root
            )

        self.assertTrue(any("decoded-pixel difference" in error for error in errors))

    def test_text_shift_difference_has_an_exact_finite_selector(self):
        document = reference_difference_document(
            shift_differences=[
                {
                    "id": "page-10-august-bergquist-shift",
                    "page": 10,
                    "kind": "text-shift",
                    "layout_finding": {
                        "label": "AugustBergquist",
                        "occurrence": 1,
                        "dx": 12.840519999999998,
                        "dy": -1.0339500000000044,
                    },
                }
            ]
        )

        registry, errors = validate_reference_difference_document(document)

        self.assertEqual(errors, [])
        self.assertEqual(
            registry["page-10-august-bergquist-shift"]["kind"], "text-shift"
        )

    def test_text_shift_rejects_non_finite_measurements_and_duplicate_selectors(self):
        shift = {
            "id": "page-10-august-bergquist-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "AugustBergquist",
                "occurrence": 1,
                "dx": float("inf"),
                "dy": -1.0,
            },
        }
        duplicate = {**shift, "id": "page-10-august-bergquist-shift-copy"}
        document = reference_difference_document(shift_differences=[shift, duplicate])

        _, errors = validate_reference_difference_document(document)

        self.assertTrue(any("dx must be finite" in error for error in errors))
        self.assertTrue(
            any("selector appears more than once" in error for error in errors)
        )


class LayoutAuditTests(unittest.TestCase):
    def test_clean_report_accepts_pass_dispositions(self):
        self.assertEqual(validate_report(visual_body(), layout_report()), [])

    def test_page_count_failure_rejects_pass_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(gt_pages=2, out_pages=1),
        )
        self.assertTrue(any("page count" in error and "issue reference" in error for error in errors))

    def test_text_flow_failure_rejects_none_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(missing=2, extra=5, wraps=1, reflow_gt=2, reflow_out=1),
        )
        self.assertTrue(any("text flow" in error and "issue reference" in error for error in errors))

    def test_large_shift_failure_rejects_pass_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(large_shifts=5),
        )
        self.assertTrue(any("large shifts" in error and "issue reference" in error for error in errors))

    def test_fine_shift_failure_rejects_pass_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(fine_shifts=2),
        )
        self.assertTrue(any("fine shifts" in error and "issue reference" in error for error in errors))

    def test_fine_shift_failure_accepts_tracked_position_issue(self):
        body = visual_body({"Position/size": "Remaining: #328"}).replace(
            "Layout audit fine shifts: Pass",
            "Layout audit fine shifts: #328",
        )
        self.assertEqual(validate_report(body, layout_report(fine_shifts=2)), [])

    def test_exact_reference_shift_is_accepted_for_large_and_fine_categories(self):
        difference = {
            "id": "page-10-august-bergquist-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "AugustBergquist",
                "occurrence": 1,
                "dx": 12.840519999999998,
                "dy": -1.0339500000000044,
            },
        }
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 10")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit large shifts: Pass",
                "Layout audit large shifts: ref:page-10-august-bergquist-shift",
            )
            .replace(
                "Layout audit fine shifts: Pass",
                "Layout audit fine shifts: ref:page-10-august-bergquist-shift",
            )
        )
        finding = {
            "label": "AugustBergquist",
            "dx": 12.840519999999998,
            "dy": -1.0339500000000044,
        }
        report = layout_report(
            large_shifts=1,
            fine_shifts=1,
            large_shift_entries=[finding],
            fine_shift_entries=[finding],
        )

        self.assertEqual(
            validate_report(
                body,
                report,
                reference_difference_document(shift_differences=[difference]),
            ),
            [],
        )

    def test_reference_shift_must_match_page_label_occurrence_and_measurements(self):
        difference = {
            "id": "page-10-sales-second-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "Sales",
                "occurrence": 2,
                "dx": -27.0,
                "dy": 0.0,
            },
        }
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 10")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit large shifts: Pass",
                "Layout audit large shifts: ref:page-10-sales-second-shift",
            )
        )
        report = layout_report(
            large_shifts=1,
            large_shift_entries=[{"label": "Sales [2/2]", "dx": -27.0, "dy": 0.0}],
        )

        self.assertEqual(
            validate_report(
                body,
                report,
                reference_difference_document(shift_differences=[difference]),
            ),
            [],
        )

        wrong_findings = []
        for name, value in (
            ("page", 11),
            ("label", "Other label"),
            ("occurrence", 1),
            ("dx", -26.999),
            ("dy", 0.001),
        ):
            wrong_difference = json.loads(json.dumps(difference))
            field_name = "page" if name == "page" else "layout_finding"
            target_name = name if name == "page" else name
            if name == "page":
                wrong_difference[field_name] = value
            else:
                wrong_difference[field_name][target_name] = value
            errors = validate_report(
                body,
                report,
                reference_difference_document(shift_differences=[wrong_difference]),
            )
            wrong_findings.append(
                any(
                    "does not match an exact current shift" in error
                    for error in errors
                )
            )

        self.assertEqual(wrong_findings, [True] * 5)

    def test_reference_shift_does_not_cover_unreferenced_shift_without_issue(self):
        difference = {
            "id": "page-10-august-bergquist-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "AugustBergquist",
                "occurrence": 1,
                "dx": 12.840519999999998,
                "dy": -1.0339500000000044,
            },
        }
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 10")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit large shifts: Pass",
                "Layout audit large shifts: ref:page-10-august-bergquist-shift",
            )
        )
        report = layout_report(
            large_shifts=2,
            large_shift_entries=[
                {
                    "label": "AugustBergquist",
                    "dx": 12.840519999999998,
                    "dy": -1.0339500000000044,
                },
                {"label": "VictoriaLindqvist", "dx": 9.985, "dy": -1.0135},
            ],
        )

        errors = validate_report(
            body,
            report,
            reference_difference_document(shift_differences=[difference]),
        )

        self.assertTrue(
            any("not covered by exact reference" in error for error in errors)
        )

    def test_reference_shift_can_coexist_with_an_open_issue_for_another_shift(self):
        difference = {
            "id": "page-10-august-bergquist-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "AugustBergquist",
                "occurrence": 1,
                "dx": 12.840519999999998,
                "dy": -1.0339500000000044,
            },
        }
        body = (
            visual_body({"Position/size": "Remaining: #328"})
            .replace("- Page(s): 1", "- Page(s): 10")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit large shifts: Pass",
                "Layout audit large shifts: ref:page-10-august-bergquist-shift, #328",
            )
        )
        report = layout_report(
            large_shifts=2,
            large_shift_entries=[
                {
                    "label": "AugustBergquist",
                    "dx": 12.840519999999998,
                    "dy": -1.0339500000000044,
                },
                {"label": "VictoriaLindqvist", "dx": 9.985, "dy": -1.0135},
            ],
        )

        self.assertEqual(
            validate_report(
                body,
                report,
                reference_difference_document(shift_differences=[difference]),
            ),
            [],
        )

    def test_shift_reference_requires_text_shift_manifest_kind(self):
        body = (
            visual_body()
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit large shifts: Pass",
                "Layout audit large shifts: ref:page-9-slide-number-visibility",
            )
        )
        report = layout_report(
            large_shifts=1,
            large_shift_entries=[{"label": "9", "dx": 12.0, "dy": 0.0}],
        )

        errors = validate_report(body, report, reference_difference_document())

        self.assertTrue(any("not a text-shift difference" in error for error in errors))

    def test_shift_reference_cannot_disposition_geometry_or_wrong_difference_kind(self):
        difference = {
            "id": "page-10-august-bergquist-shift",
            "page": 10,
            "kind": "text-shift",
            "layout_finding": {
                "label": "AugustBergquist",
                "occurrence": 1,
                "dx": 12.840519999999998,
                "dy": -1.0339500000000044,
            },
        }
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 10")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit rectangle geometry: Pass",
                "Layout audit rectangle geometry: ref:page-10-august-bergquist-shift",
            )
        )
        document = reference_difference_document(shift_differences=[difference])
        report = layout_report(rect_geometry=1)

        errors = validate_report(body, report, document)

        self.assertTrue(
            any("does not support reference exporter" in error for error in errors)
        )

    def test_fine_detail_threshold_is_required_in_pr_metadata(self):
        body = visual_body().replace("- Fine-detail threshold: 0.5pt\n", "")
        errors = validate_report(body, layout_report())
        self.assertTrue(any("Fine-detail threshold" in error and "required" in error for error in errors))

    def test_fine_detail_threshold_must_match_machine_report(self):
        body = visual_body().replace(
            "Fine-detail threshold: 0.5pt",
            "Fine-detail threshold: 0.75pt",
        )
        errors = validate_report(body, layout_report())
        self.assertTrue(any("Fine-detail threshold" in error and "0.5" in error for error in errors))

    def test_machine_report_must_record_fine_detail_threshold(self):
        errors = validate_report(
            visual_body(),
            layout_report(fine_threshold=None),
        )
        self.assertTrue(any("fine-detail threshold" in error.lower() for error in errors))

    def test_visibility_failure_rejects_pass_text_flow_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(visibility=1),
        )
        self.assertTrue(
            any("text flow" in error and "issue reference" in error for error in errors)
        )

    def test_page_9_visibility_difference_accepts_exact_native_evidence(self):
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 9")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit text flow: Pass",
                "Layout audit text flow: ref:page-9-slide-number-visibility",
            )
        )
        report = layout_report(visibility=1)
        report["pages"][0]["visibility"]["mismatches"] = [
            {"label": "9", "gt": "hidden", "out": "painted"}
        ]

        self.assertEqual(
            validate_report(body, report, reference_difference_document()),
            [],
        )

    def test_rasterized_gt_text_can_disposition_the_exact_extra_line(self) -> None:
        difference_id = "page-13-slide-number-rasterized"
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
        )
        report = layout_report(extra=1, extra_text=["13"])

        errors = validate_report(
            rasterized_text_body(difference_id), report, document
        )

        self.assertEqual(errors, [])

    def test_rasterized_output_text_can_disposition_the_exact_missing_line(
        self,
    ) -> None:
        difference_id = "page-13-output-text-rasterized"
        difference = rasterized_text_difference(
            difference_id=difference_id, side="out"
        )
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[difference],
        )
        report = layout_report(missing=1, missing_text=["13"])

        errors = validate_report(
            rasterized_text_body(difference_id), report, document
        )

        self.assertEqual(errors, [])

    def test_rasterized_text_reference_must_match_page_label_side_and_occurrence(
        self,
    ) -> None:
        cases = (
            (14, "13", "gt", 1, ["13"]),
            (13, "12", "gt", 1, ["13"]),
            (13, "13", "out", 1, ["13"]),
            (13, "13", "gt", 2, ["13"]),
        )
        for page, label, side, occurrence, extra_text in cases:
            with self.subTest(
                page=page, label=label, side=side, occurrence=occurrence
            ):
                difference_id = "page-13-slide-number-rasterized"
                difference = rasterized_text_difference(
                    label=label,
                    side=side,
                    occurrence=occurrence,
                )
                document = reference_difference_document(
                    native_ground_truth=True,
                    rasterized_text_differences=[difference],
                )
                body = rasterized_text_body(difference_id, page=page)
                report = layout_report(extra=1, extra_text=extra_text)

                errors = validate_report(body, report, document)

                self.assertTrue(
                    any(
                        "does not match an exact current rasterized-text finding"
                        in error
                        for error in errors
                    ),
                    errors,
                )

    def test_rasterized_text_reference_leaves_other_text_findings_open(self) -> None:
        difference_id = "page-13-slide-number-rasterized"
        document = reference_difference_document(
            native_ground_truth=True,
            rasterized_text_differences=[rasterized_text_difference()],
        )
        report = layout_report(
            extra=2,
            extra_text=["13", "unrelated line"],
            wraps=1,
        )

        errors = validate_report(
            rasterized_text_body(difference_id), report, document
        )

        self.assertTrue(
            any("requires an open issue" in error for error in errors), errors
        )

    def test_reference_difference_cannot_hide_a_different_visibility_finding(self):
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 9")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit text flow: Pass",
                "Layout audit text flow: ref:page-9-slide-number-visibility",
            )
        )
        report = layout_report(visibility=1)
        report["pages"][0]["visibility"]["mismatches"] = [
            {"label": "10", "gt": "hidden", "out": "painted"}
        ]

        errors = validate_report(body, report, reference_difference_document())

        self.assertTrue(any("does not match" in error for error in errors))

    def test_reference_difference_does_not_cover_real_text_flow_defects(self):
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 9")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit text flow: Pass",
                "Layout audit text flow: ref:page-9-slide-number-visibility",
            )
        )
        report = layout_report(visibility=1, missing=1)
        report["pages"][0]["lines"]["missing_text"] = ["real converter defect"]
        report["pages"][0]["visibility"]["mismatches"] = [
            {"label": "9", "gt": "hidden", "out": "painted"}
        ]

        errors = validate_report(body, report, reference_difference_document())

        self.assertTrue(any("open issue" in error for error in errors))

    def test_reference_difference_can_coexist_with_a_tracked_real_defect(self):
        body = (
            visual_body({"Line/paragraph spacing": "Remaining: #328"})
            .replace("- Page(s): 1", "- Page(s): 9")
            .replace(
                "Reference exporter differences: None",
                "Reference exporter differences: "
                "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
            )
            .replace(
                "Layout audit text flow: Pass",
                "Layout audit text flow: ref:page-9-slide-number-visibility, #328",
            )
        )
        report = layout_report(visibility=1, missing=1)
        report["pages"][0]["lines"]["missing_text"] = ["real converter defect"]
        report["pages"][0]["visibility"]["mismatches"] = [
            {"label": "9", "gt": "hidden", "out": "painted"}
        ]

        self.assertEqual(
            validate_report(body, report, reference_difference_document()),
            [],
        )

    def test_visible_fill_failure_rejects_pass_visible_fill_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(visible_fills=1),
        )
        self.assertTrue(
            any("visible fills" in error and "issue reference" in error for error in errors)
        )

    def test_visible_fill_failure_accepts_remaining_fill_issue(self):
        body = visual_body({"Fill": "Remaining: #328"}).replace(
            "Layout audit visible fills: Pass",
            "Layout audit visible fills: #328",
        )
        self.assertEqual(validate_report(body, layout_report(visible_fills=1)), [])

    def test_rect_geometry_failure_rejects_pass_disposition(self):
        errors = validate_report(
            visual_body(),
            layout_report(rect_geometry=1),
        )
        self.assertTrue(
            any(
                "rectangle geometry" in error and "issue reference" in error
                for error in errors
            )
        )

    def test_rect_geometry_failure_accepts_remaining_position_issue(self):
        body = visual_body({"Position/size": "Remaining: #328"}).replace(
            "Layout audit rectangle geometry: Pass",
            "Layout audit rectangle geometry: #328",
        )
        self.assertEqual(validate_report(body, layout_report(rect_geometry=1)), [])

    def test_failed_categories_accept_remaining_issue_references(self):
        body = visual_body(
            {
                "Position/size": "Remaining: #328",
                "Line/paragraph spacing": "Remaining: #329",
            }
        )
        body = body.replace(
            "Layout audit text flow: Pass",
            "Layout audit text flow: #329",
        ).replace(
            "Layout audit large shifts: Pass",
            "Layout audit large shifts: #328, #329",
        )
        errors = validate_report(
            body,
            layout_report(missing=1, reflow_gt=2, reflow_out=1, large_shifts=3),
        )
        self.assertEqual(errors, [])

    def test_failed_category_rejects_issue_absent_from_remaining_rows(self):
        body = visual_body().replace(
            "Layout audit large shifts: Pass",
            "Layout audit large shifts: #328",
        )
        errors = validate_report(body, layout_report(large_shifts=1))
        self.assertTrue(any("must also appear in a Remaining" in error for error in errors))

    def test_clean_category_rejects_issue_disposition(self):
        body = visual_body({"Position/size": "Remaining: #328"}).replace(
            "Layout audit large shifts: Pass",
            "Layout audit large shifts: #328",
        )
        errors = validate_report(body, layout_report())
        self.assertTrue(any("has no findings and must be Pass" in error for error in errors))

    def test_fix_audit_requires_changed_machine_report(self):
        with tempfile.TemporaryDirectory() as directory:
            errors = validate_layout_audit(visual_body(), [], Path(directory))
        self.assertTrue(any("must be changed in this pull request" in error for error in errors))

    def test_fix_audit_report_must_accompany_fresh_image_evidence(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report_path = root / "assets/bugfixes/issue-186/layout-audit.json"
            report_path.parent.mkdir(parents=True)
            report_path.write_text(json.dumps(layout_report()), encoding="utf-8")
            errors = validate_layout_audit(
                visual_body(),
                ["assets/bugfixes/issue-186/layout-audit.json"],
                root,
            )
        self.assertTrue(any("after.jpg must be changed" in error for error in errors))

    def test_fix_audit_report_requires_changed_after_image(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report_path = root / "assets/bugfixes/issue-186/layout-audit.json"
            report_path.parent.mkdir(parents=True)
            report_path.write_text(json.dumps(layout_report()), encoding="utf-8")
            errors = validate_layout_audit(
                visual_body(),
                [
                    "assets/bugfixes/issue-186/gt.jpg",
                    "assets/bugfixes/issue-186/layout-audit.json",
                ],
                root,
            )
        self.assertTrue(any("after.jpg must be changed" in error for error in errors))

    def test_text_layer_only_fix_accepts_identical_decoded_before_and_after_images(self):
        body = visual_body().replace(
            "- Evidence mode: `fix`",
            "- Evidence mode: `fix`\n- Text-layer-only: Yes\n- Pixel delta: 0",
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            issue_dir = root / "assets/bugfixes/issue-186"
            issue_dir.mkdir(parents=True)
            (issue_dir / "layout-audit.json").write_text(
                json.dumps(layout_report()), encoding="utf-8"
            )
            evidence = (ROOT / "assets/bugfixes/issue-186/after.jpg").read_bytes()
            (issue_dir / "before.jpg").write_bytes(evidence)
            (issue_dir / "after.jpg").write_bytes(evidence)
            errors = validate_layout_audit(
                body,
                ["assets/bugfixes/issue-186/layout-audit.json"],
                root,
            )
        self.assertEqual(errors, [])

    def test_text_layer_only_fix_rejects_different_decoded_before_and_after_images(self):
        body = visual_body().replace(
            "- Evidence mode: `fix`",
            "- Evidence mode: `fix`\n- Text-layer-only: Yes\n- Pixel delta: 0",
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            issue_dir = root / "assets/bugfixes/issue-186"
            issue_dir.mkdir(parents=True)
            (issue_dir / "layout-audit.json").write_text(
                json.dumps(layout_report()), encoding="utf-8"
            )
            (issue_dir / "before.jpg").write_bytes(
                (ROOT / "assets/bugfixes/issue-186/gt.jpg").read_bytes()
            )
            (issue_dir / "after.jpg").write_bytes(
                (ROOT / "assets/bugfixes/issue-186/after.jpg").read_bytes()
            )
            errors = validate_layout_audit(
                body,
                ["assets/bugfixes/issue-186/layout-audit.json"],
                root,
            )
        self.assertTrue(any("decoded-pixel delta" in error for error in errors))

    def test_text_layer_only_fix_requires_zero_pixel_delta(self):
        body = visual_body().replace(
            "- Evidence mode: `fix`",
            "- Evidence mode: `fix`\n- Text-layer-only: Yes\n- Pixel delta: 1",
        )
        errors = validate_report(body, layout_report())
        self.assertTrue(any("Pixel delta must be 0" in error for error in errors))

    def test_text_layer_only_fix_requires_current_before_image(self):
        body = visual_body().replace(
            "- Evidence mode: `fix`",
            "- Evidence mode: `fix`\n- Text-layer-only: Yes\n- Pixel delta: 0",
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            issue_dir = root / "assets/bugfixes/issue-186"
            issue_dir.mkdir(parents=True)
            (issue_dir / "layout-audit.json").write_text(
                json.dumps(layout_report()), encoding="utf-8"
            )
            (issue_dir / "after.jpg").write_bytes(
                (ROOT / "assets/bugfixes/issue-186/after.jpg").read_bytes()
            )
            errors = validate_layout_audit(
                body,
                ["assets/bugfixes/issue-186/layout-audit.json"],
                root,
            )
        self.assertTrue(any("before.jpg: current evidence is required" in error for error in errors))

    def test_fix_audit_requires_issue_scoped_report_path(self):
        body = visual_body().replace(
            "assets/bugfixes/issue-186/layout-audit.json",
            "assets/bugfixes/issue-999/layout-audit.json",
        )
        errors = validate_report(body, layout_report())
        self.assertTrue(any("Layout audit report must be" in error for error in errors))

    def test_malformed_report_is_rejected(self):
        errors = validate_report(visual_body(), {"pages": "not-a-list"})
        self.assertTrue(any("invalid layout audit report" in error for error in errors))

    def test_incomplete_page_vectors_are_rejected(self):
        errors = validate_report(
            visual_body(),
            layout_report(gt_pages=2, out_pages=2),
        )
        self.assertTrue(any("invalid layout audit report" in error for error in errors))


class RenderClusterAuditTests(unittest.TestCase):
    def test_strict_dispositioned_report_passes(self):
        self.assertEqual(
            validate_cluster_reports(visual_body(), {1: render_cluster_report()}),
            [],
        )

    def test_fix_requires_render_cluster_report_field(self):
        body = visual_body().replace(
            "- Render cluster reports: `assets/bugfixes/issue-186/render-clusters-page-1.json`\n",
            "",
        )

        errors = validate_cluster_reports(body, {1: render_cluster_report()})

        self.assertTrue(any("Render cluster reports" in error for error in errors))

    def test_report_must_be_changed_in_the_pull_request(self):
        errors = validate_cluster_reports(visual_body(), {})

        self.assertTrue(any("must be changed" in error for error in errors))

    def test_non_strict_report_is_rejected(self):
        errors = validate_cluster_reports(
            visual_body(), {1: render_cluster_report(strict=False)}
        )

        self.assertTrue(any("strict" in error.lower() for error in errors))

    def test_undispositioned_cluster_is_rejected(self):
        report = render_cluster_report(
            passed=False,
            disposition=None,
            undispositioned=["placeholder"],
        )
        cluster_id = report["clusters"][0]["id"]
        report["undispositioned_cluster_ids"] = [cluster_id]
        report["clusters"][0]["disposition"] = None

        errors = validate_cluster_reports(visual_body(), {1: report})

        self.assertTrue(any(cluster_id in error for error in errors))

    def test_report_rejects_id_that_does_not_match_cluster_geometry(self):
        report = render_cluster_report()
        report["clusters"][0]["id"] = "p1-0123456789ab"

        errors = validate_cluster_reports(visual_body(), {1: report})

        self.assertTrue(any("does not match" in error for error in errors))

    def test_issue_disposition_must_be_a_remaining_open_issue_reference(self):
        report = render_cluster_report(
            disposition={"kind": "issue", "issue": "#328"}
        )

        errors = validate_cluster_reports(visual_body(), {1: report})

        self.assertTrue(any("#328" in error and "Remaining" in error for error in errors))

    def test_reference_exporter_disposition_uses_exact_manifest_cluster(self):
        report = render_cluster_report()
        cluster_id = report["clusters"][0]["id"]
        report["clusters"][0]["disposition"] = {
            "kind": "reference-exporter-difference",
            "difference_id": "page-1-title-native-match",
        }
        body = visual_body().replace(
            "Reference exporter differences: None",
            "Reference exporter differences: "
            "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
        )

        errors = validate_cluster_reports(
            body,
            {1: report},
            reference_difference_document(cluster_ids=[cluster_id]),
        )

        self.assertEqual(errors, [])

    def test_reference_exporter_disposition_rejects_unlisted_cluster(self):
        report = render_cluster_report()
        report["clusters"][0]["disposition"] = {
            "kind": "reference-exporter-difference",
            "difference_id": "page-1-title-native-match",
        }
        body = visual_body().replace(
            "Reference exporter differences: None",
            "Reference exporter differences: "
            "`assets/bugfixes/issue-186/reference-exporter-differences.json`",
        )

        errors = validate_cluster_reports(
            body,
            {1: report},
            reference_difference_document(cluster_ids=["p1-0123456789ab"]),
        )

        self.assertTrue(any("not listed" in error for error in errors))

    def test_every_declared_page_requires_its_own_report(self):
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 1-2")
            .replace(
                "render-clusters-page-1.json`",
                "render-clusters-page-1.json`, `assets/bugfixes/issue-186/"
                "render-clusters-page-2.json`",
            )
        )

        errors = validate_cluster_reports(body, {1: render_cluster_report(page=1)})

        self.assertTrue(any("page 2" in error for error in errors))

    def test_all_declared_pages_pass_with_distinct_reports(self):
        body = (
            visual_body()
            .replace("- Page(s): 1", "- Page(s): 1-2")
            .replace(
                "render-clusters-page-1.json`",
                "render-clusters-page-1.json`, `assets/bugfixes/issue-186/"
                "render-clusters-page-2.json`",
            )
        )

        errors = validate_cluster_reports(
            body,
            {1: render_cluster_report(page=1), 2: render_cluster_report(page=2)},
        )

        self.assertEqual(errors, [])

    def test_valid_below_floor_renderer_observation_passes(self):
        observation = {
            "class": "gradient-rasterization",
            "bbox_pt": {"x": 10.0, "y": 20.0, "width": 80.0, "height": 40.0},
            "note": "Full-resolution crop shows only gradient decomposition.",
        }

        errors = validate_cluster_reports(
            visual_body(),
            {1: render_cluster_report(renderer_observations=[observation])},
        )

        self.assertEqual(errors, [])

    def test_renderer_observation_requires_a_bounded_region(self):
        observation = {
            "class": "shape-edge-antialiasing",
            "note": "Blanket page claim.",
        }

        errors = validate_cluster_reports(
            visual_body(),
            {1: render_cluster_report(renderer_observations=[observation])},
        )

        self.assertTrue(any("renderer observation" in error.lower() for error in errors))


class EvidenceTests(unittest.TestCase):
    def test_render_cluster_report_is_machine_evidence_not_an_image(self):
        self.assertEqual(
            validate_evidence(
                ["assets/bugfixes/issue-186/render-clusters-page-1.json"], ROOT
            ),
            [],
        )

    def test_repository_evidence_is_progressive_150_dpi_jpeg(self):
        path = ROOT / "assets/bugfixes/issue-186/gt.jpg"
        info = read_jpeg_info(path)
        self.assertTrue(info.progressive)
        self.assertEqual(info.density_dpi, (150.0, 150.0))
        self.assertEqual(info.metadata_markers, ())

    def test_changed_trio_validates_all_three_files(self):
        errors = validate_evidence(
            ["assets/bugfixes/issue-186/after.jpg"],
            ROOT,
        )
        self.assertEqual(errors, [])

    def test_png_evidence_is_rejected(self):
        errors = validate_evidence(
            ["assets/bugfixes/issue-186/after.png"],
            ROOT,
        )
        self.assertTrue(any(".jpg extension" in error for error in errors))

    def test_evidence_readme_is_not_evidence(self):
        self.assertEqual(validate_evidence(["assets/bugfixes/README.md"], ROOT), [])

    def test_layout_audit_report_is_validated_separately_from_jpegs(self):
        self.assertEqual(
            validate_evidence(
                ["assets/bugfixes/issue-186/layout-audit.json"],
                ROOT,
            ),
            [],
        )

    def test_issue_directory_notes_are_not_evidence(self):
        self.assertEqual(
            validate_evidence(["assets/bugfixes/issue-186/notes.txt"], ROOT),
            [],
        )

    def test_unnamed_image_is_still_rejected(self):
        errors = validate_evidence(["assets/bugfixes/issue-186/extra.jpg"], ROOT)
        self.assertTrue(any("visual evidence must be" in error for error in errors))


class OpenIssueTests(unittest.TestCase):
    @patch("scripts.check_visual_pr.urllib.request.urlopen")
    def test_remaining_issue_must_be_open(self, urlopen):
        response = io.BytesIO(b'{"state":"closed"}')
        urlopen.return_value = response
        errors = validate_open_issues({328}, "developer0hye/office2pdf", "token")
        self.assertEqual(errors, ["Remaining visual issue #328 is not open."])


if __name__ == "__main__":
    unittest.main()
