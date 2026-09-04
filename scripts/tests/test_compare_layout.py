"""Unit tests for the trace-based layout differ.

All fixtures are synthetic ``mutool draw -F trace`` fragments that mirror the
real output format (numberless ``<page>`` opening tag, device transform on each
op, glyph coordinates in text space, sizes in trm units), so the parser is
exercised on the same shapes it will meet in production without shelling out to
mutool.
"""

from __future__ import annotations

import io
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import compare_layout


def trace_document(
    *pages: str,
    numbered: bool = False,
    mediabox: str = "0 0 595.2 841.92",
) -> str:
    """Wrap page bodies in a trace document.

    Defaults to the numberless `<page mediabox="...">` mutool 1.23.x emits, so
    the shared fixtures exercise the shape the parser actually meets. Pass
    ``numbered=True`` for the attribute later builds add, or override
    ``mediabox`` for a fixture with different page geometry.
    """
    body = "\n".join(
        "<page {}mediabox=\"{}\">\n{}\n</page>".format(
            f'number="{i + 1}" ' if numbered else "", mediabox, content
        )
        for i, content in enumerate(pages)
    )
    return f'<?xml version="1.0"?>\n<document filename="x.pdf">\n{body}\n</document>'


def text_op(
    words: list[tuple[str, float]],
    baseline_y: float,
    size_units: float = 44.0,
    scale: float = 0.24,
    color: str = "0 0 0",
    font: str = "AAAAAA+ArialMT",
) -> str:
    """One fill_text whose glyphs sit at ``baseline_y`` (device pt).

    ``words`` is a list of (character, device_x) pairs; coordinates are
    converted back into text space so the parser has to apply the transform.
    """
    offset = 841.92
    glyph_lines = []
    for char, device_x in words:
        text_x = device_x / scale
        text_y = (offset - baseline_y) / scale
        glyph_lines.append(
            f'<g unicode="{char}" glyph="1" x="{text_x:.4f}" y="{text_y:.4f}" adv=".5"/>'
        )
    glyphs = "\n".join(glyph_lines)
    return (
        f'<fill_text colorspace="ICCBased(RGB,sRGB IEC61966-2.1)" color="{color}" '
        f'ri="1" bp="1" op="0" opm="0" transform="{scale} 0 0 -{scale} 0 {offset}">\n'
        f'<span font="{font}" wmode="0" bidi="0" trm="{size_units} 0 0 {size_units}">\n'
        f"{glyphs}\n</span>\n</fill_text>"
    )


def line_of(text: str, x0: float, baseline_y: float, pitch: float = 6.0) -> str:
    words = [(char, x0 + i * pitch) for i, char in enumerate(text)]
    return text_op(words, baseline_y)


def ignored_text_op(
    words: list[tuple[str, float]], baseline_y: float, size_units: float = 44.0
) -> str:
    """Synthetic ``ignore_text`` geometry for path-correlation tests."""
    return (
        text_op(words, baseline_y, size_units=size_units)
        .replace("<fill_text ", "<ignore_text ", 1)
        .replace("</fill_text>", "</ignore_text>", 1)
    )


def rect_op(
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    kind: str = "fill_path",
    color: str = ".8 .8 .8",
    alpha: float = 1.0,
) -> str:
    extra = ' winding="nonzero"' if kind == "fill_path" else ' linewidth="1"'
    return (
        f'<{kind}{extra} colorspace="ICCBased(RGB,sRGB IEC61966-2.1)" color="{color}" '
        f'alpha="{alpha}" ri="1" bp="1" op="0" opm="0" '
        f'transform="1 0 0 -1 0 841.92">\n'
        f'<moveto x="{x0}" y="{841.92 - y0}"/>\n'
        f'<lineto x="{x1}" y="{841.92 - y0}"/>\n'
        f'<lineto x="{x1}" y="{841.92 - y1}"/>\n'
        f'<lineto x="{x0}" y="{841.92 - y1}"/>\n'
        f"<closepath/>\n</{kind}>"
    )


def line_op(
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    color: str = ".8 .8 .8",
) -> str:
    return "\n".join(
        [
            '<stroke_path linewidth="1" colorspace="ICCBased(RGB,sRGB IEC61966-2.1)" '
            f'color="{color}" alpha="1" transform="1 0 0 1 0 0">',
            f'<moveto x="{x0}" y="{y0}"/>',
            f'<lineto x="{x1}" y="{y1}"/>',
            "</stroke_path>",
        ]
    )


def image_op() -> str:
    return (
        '<fill_image alpha="1" colorspace="DeviceRGB" ri="1" bp="1" op="0" opm="0" '
        'transform="595.2 0 0 841.92 0 0" width="1280" height="720"/>'
    )


def clipped_shade_op(
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    alpha: float = 1.0,
    extend: str = "1 1",
    clip_transform: str = "1 0 0 1 0 0",
) -> str:
    return "\n".join(
        [
            f'<clip_path winding="nonzero" transform="{clip_transform}">',
            f'<moveto x="{x0}" y="{y0}"/>',
            f'<lineto x="{x1}" y="{y0}"/>',
            f'<lineto x="{x1}" y="{y1}"/>',
            f'<lineto x="{x0}" y="{y1}"/>',
            "<closepath/>",
            "</clip_path>",
            f'<fill_shade alpha="{alpha}" transform="1 0 0 1 0 0" '
            f'type="linear" extend="{extend}" start="0 0" end="1 1"/>',
            "<pop_clip/>",
        ]
    )


def clipped_text_op(
    content: str,
    x0: float,
    y0: float,
    x1: float,
    y1: float,
) -> str:
    """Wrap text in one rectangular device-space clip."""
    return "\n".join(
        [
            '<clip_path winding="nonzero" transform="1 0 0 1 0 0">',
            f'<moveto x="{x0}" y="{y0}"/>',
            f'<lineto x="{x1}" y="{y0}"/>',
            f'<lineto x="{x1}" y="{y1}"/>',
            f'<lineto x="{x0}" y="{y1}"/>',
            "<closepath/>",
            "</clip_path>",
            content,
            "<pop_clip/>",
        ]
    )


class PageElementTest(unittest.TestCase):
    """The ``<page>`` opening tag differs across mutool releases.

    1.23.x emits ``<page mediabox="...">`` with no ``number`` attribute, so a
    parser that requires one measures nothing at all — and reports that as
    "no differences" rather than as a failure.
    """

    def test_page_without_number_attribute_is_parsed(self) -> None:
        pages = compare_layout.parse_trace(
            trace_document(text_op([("A", 72.0)], baseline_y=100.0))
        )
        self.assertEqual(len(pages), 1)
        self.assertEqual(pages[0].media_box, (0.0, 0.0, 595.2, 841.92))
        self.assertEqual(pages[0].lines[0].text, "A")

    def test_pages_without_number_keep_document_order(self) -> None:
        pages = compare_layout.parse_trace(
            trace_document(
                text_op([("A", 72.0)], baseline_y=100.0),
                text_op([("B", 72.0)], baseline_y=100.0),
            )
        )
        self.assertEqual([p.lines[0].text for p in pages], ["A", "B"])

    def test_page_with_number_attribute_still_parses(self) -> None:
        pages = compare_layout.parse_trace(
            trace_document(text_op([("A", 72.0)], baseline_y=100.0), numbered=True)
        )
        self.assertEqual(len(pages), 1)
        self.assertEqual(pages[0].lines[0].text, "A")


class ParseTraceTest(unittest.TestCase):
    def test_glyphs_carry_device_coordinates_and_pt_sizes(self) -> None:
        doc = trace_document(text_op([("A", 72.0), ("B", 78.0)], baseline_y=100.0))
        pages = compare_layout.parse_trace(doc)
        self.assertEqual(len(pages), 1)
        line = pages[0].lines[0]
        self.assertEqual(line.text, "AB")
        self.assertAlmostEqual(line.glyphs[0].x, 72.0, places=3)
        self.assertAlmostEqual(line.glyphs[0].y, 100.0, places=3)
        # trm 44 under a 0.24 transform is a 10.56pt glyph.
        self.assertAlmostEqual(line.glyphs[0].size, 10.56, places=3)
        # adv .5 em at that size is 5.28pt.
        self.assertAlmostEqual(line.glyphs[0].advance, 5.28, places=3)

    def test_glyphs_on_one_baseline_group_into_one_line_in_x_order(self) -> None:
        page = "\n".join(
            [
                text_op([("B", 80.0)], baseline_y=100.0),
                text_op([("A", 72.0)], baseline_y=100.2),
                text_op([("C", 90.0)], baseline_y=300.0),
            ]
        )
        pages = compare_layout.parse_trace(trace_document(page))
        texts = [line.text for line in pages[0].lines]
        self.assertEqual(texts, ["AB", "C"])

    def test_ignore_text_with_compact_preceding_path_is_measured_as_painted_text(
        self,
    ) -> None:
        page = "\n".join(
            [
                rect_op(71.0, 90.0, 78.0, 102.0, color="0 0 0"),
                ignored_text_op([("A", 72.0)], baseline_y=100.0),
            ]
        )

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.text, "A")
        self.assertAlmostEqual(line.x0, 72.0, places=3)
        self.assertAlmostEqual(line.y, 100.0, places=3)
        self.assertEqual(line.visibility, "painted")

    def test_ignore_text_without_preceding_path_remains_hidden(self) -> None:
        page = ignored_text_op([("OCR", 72.0)], baseline_y=100.0)

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.text, "OCR")
        self.assertEqual(line.visibility, "hidden")

    def test_ignore_text_after_transparent_path_remains_hidden(self) -> None:
        page = "\n".join(
            [
                rect_op(71.0, 90.0, 78.0, 102.0, color="0 0 0", alpha=0.0),
                ignored_text_op([("OCR", 72.0)], baseline_y=100.0),
            ]
        )

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.text, "OCR")
        self.assertEqual(line.visibility, "hidden")

    def test_later_image_still_hides_path_correlated_ignore_text(self) -> None:
        path_correlated_text = "\n".join(
            [
                rect_op(71.0, 90.0, 78.0, 102.0, color="0 0 0"),
                ignored_text_op([("A", 72.0)], baseline_y=100.0),
            ]
        )
        gt = "\n".join([path_correlated_text, image_op()])
        out = "\n".join([image_op(), path_correlated_text])

        vector = compare_layout.diff_page(
            compare_layout.parse_trace(trace_document(gt))[0],
            compare_layout.parse_trace(trace_document(out))[0],
        )

        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(vector["visibility"]["mismatches"][0]["gt"], "hidden")
        self.assertEqual(vector["visibility"]["mismatches"][0]["out"], "painted")

    def test_rotated_fill_text_uses_the_full_affine_transform_and_stays_one_run(self) -> None:
        page = """<fill_text transform="1 2 3 4 5 6">
          <span font="AAAAAA+ArialMT" trm="10 0 0 10">
            <g unicode="A" glyph="A" x="7" y="11" adv=".5"/>
            <g unicode="B" glyph="B" x="8" y="11" adv=".5"/>
          </span>
        </fill_text>"""
        lines = compare_layout.parse_trace(trace_document(page))[0].lines
        self.assertEqual(len(lines), 1)
        self.assertEqual(lines[0].text, "AB")
        self.assertAlmostEqual(lines[0].x0, 45.0)
        self.assertAlmostEqual(lines[0].y, 64.0)

    def test_rotated_whitespace_only_run_remains_excluded(self) -> None:
        page = """<fill_text transform="1 2 3 4 5 6">
          <span font="AAAAAA+ArialMT" trm="10 0 0 10">
            <g unicode=" " glyph="space" x="7" y="11" adv=".5"/>
          </span>
        </fill_text>"""

        lines = compare_layout.parse_trace(trace_document(page))[0].lines

        self.assertEqual(lines, [])

    def test_path_bbox_uses_the_full_affine_transform(self) -> None:
        page = """<fill_path transform="1 2 3 4 5 6">
          <moveto x="7" y="11"/>
          <lineto x="8" y="12"/>
        </fill_path>"""
        rect = compare_layout.parse_trace(trace_document(page))[0].rects[0]
        self.assertEqual((rect.x0, rect.y0, rect.x1, rect.y1), (45.0, 64.0, 49.0, 70.0))
        self.assertEqual(rect.geometry_kind, "other")

    def test_rects_capture_device_bbox_and_kind(self) -> None:
        page = "\n".join(
            [rect_op(69.36, 792.96, 525.84, 793.44), rect_op(10, 20, 30, 21, "stroke_path")]
        )
        pages = compare_layout.parse_trace(trace_document(page))
        rects = pages[0].rects
        self.assertEqual(len(rects), 2)
        fill = next(r for r in rects if r.kind == "fill")
        self.assertAlmostEqual(fill.x0, 69.36, places=2)
        self.assertAlmostEqual(fill.y1, 793.44, places=2)
        self.assertEqual({r.kind for r in rects}, {"fill", "stroke"})
        self.assertEqual({r.geometry_kind for r in rects}, {"rectangle"})

    def test_axis_aligned_stroke_line_is_geometry_comparable(self) -> None:
        rect = compare_layout.parse_trace(
            trace_document(line_op(74.183, 137.254, 339.165, 137.254))
        )[0].rects[0]

        self.assertEqual(rect.geometry_kind, "line")
        self.assertEqual(rect.bbox, [74.183, 137.254, 339.165, 137.254])


class MatchAndDiffTest(unittest.TestCase):
    def diff(self, gt_page: str, out_page: str, **kwargs) -> dict:
        gt = compare_layout.parse_trace(trace_document(gt_page))[0]
        out = compare_layout.parse_trace(trace_document(out_page))[0]
        return compare_layout.diff_page(gt, out, **kwargs)

    def test_identical_pages_report_no_deviation(self) -> None:
        page = "\n".join(
            [line_of("hello", 72, 100), line_of("world", 72, 112), rect_op(70, 90, 300, 91)]
        )
        vector = self.diff(page, page)
        self.assertEqual(vector["lines"]["matched"], 2)
        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["lines"]["deviant"], 0)
        self.assertAlmostEqual(vector["baseline"]["mean_abs_dy"], 0.0, places=6)
        self.assertEqual(vector["rects"]["gt_count"], vector["rects"]["out_count"])

    def test_shifted_line_reports_dy_and_deviant_count(self) -> None:
        gt = "\n".join([line_of("hello", 72, 100), line_of("world", 72, 112)])
        out = "\n".join([line_of("hello", 72, 100), line_of("world", 72, 114)])
        vector = self.diff(gt, out)
        self.assertEqual(vector["lines"]["deviant"], 1)
        self.assertAlmostEqual(vector["baseline"]["worst_dy"], 2.0, places=3)
        worst = vector["baseline"]["worst_line"]
        self.assertIn("world", worst)

    def test_sub_noise_floor_shift_is_not_deviant(self) -> None:
        gt = line_of("hello", 72, 100)
        out = line_of("hello", 72, 100.08)
        vector = self.diff(gt, out, noise_floor=0.12)
        self.assertEqual(vector["lines"]["deviant"], 0)
        # The raw statistic still carries the measurement.
        self.assertGreater(vector["baseline"]["worst_dy"], 0.0)

    def test_fine_shift_gate_ignores_trace_noise_but_fails_visible_movement(self) -> None:
        gt = "\n".join([line_of("noise", 72, 100), line_of("moved", 72, 120)])
        out = "\n".join([line_of("noise", 72, 100.12), line_of("moved", 72, 120.75)])

        vector = self.diff(gt, out, fine_shift=0.5)

        self.assertEqual(vector["instances"]["large_shift_count"], 0)
        self.assertEqual(vector["instances"]["fine_shift_threshold"], 0.5)
        self.assertEqual(vector["instances"]["fine_shift_count"], 1)
        self.assertEqual(vector["instances"]["fine_shifts"][0]["label"], "moved")
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_coarse_audit_stays_clean_without_fine_shift_mode(self) -> None:
        gt = line_of("moved", 72, 120)
        out = line_of("moved", 72, 121.1)

        vector = self.diff(gt, out, large_shift=5.0)

        self.assertEqual(vector["instances"]["large_shift_threshold"], 5.0)
        self.assertEqual(vector["instances"]["large_shift_count"], 0)
        self.assertIsNone(vector["instances"]["fine_shift_threshold"])
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_missing_and_extra_lines_are_counted(self) -> None:
        gt = "\n".join([line_of("alpha", 72, 100), line_of("beta", 72, 112)])
        out = line_of("alpha", 72, 100)
        vector = self.diff(gt, out)
        self.assertEqual(vector["lines"]["missing"], 1)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertIn("beta", vector["lines"]["missing_text"][0])
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_wrap_difference_is_detected_not_reported_missing(self) -> None:
        gt = "\n".join([line_of("abcdef", 72, 100), line_of("ghi", 72, 112)])
        out = "\n".join([line_of("abc", 72, 100), line_of("defghi", 72, 112)])
        vector = self.diff(gt, out)
        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["wraps"]["count"], 1)
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_issue_1446_distant_objects_split_across_baselines_are_equivalent(
        self,
    ) -> None:
        # Page 8 of GENERAL SERVICES.pptx: the reference gives the 400
        # value-axis tick and Net Profit legend entry one shared baseline, so
        # mutool exposes one `400NetProfit` line. office2pdf puts the same two
        # distant objects on separate baselines. These coordinates are pinned
        # from the current #1220 attachment comparison. The PDF line topology
        # is not a wrap/reflow defect, but both position deltas stay auditable.
        gt = "\n".join(
            [
                line_of("400", 148.614, 400.201),
                line_of("NetProfit", 753.985, 400.201),
            ]
        )
        out = "\n".join(
            [
                line_of("400", 149.626, 391.195),
                line_of("NetProfit", 742.994, 425.407),
            ]
        )

        gt_page = compare_layout.parse_trace(
            trace_document(gt, mediabox="0 0 960 540")
        )[0]
        out_page = compare_layout.parse_trace(
            trace_document(out, mediabox="0 0 960 540")
        )[0]
        vector = compare_layout.diff_page(gt_page, out_page)

        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["lines"]["matched"], 1)
        self.assertEqual(vector["lines"]["deviant"], 1)
        self.assertEqual(vector["instances"]["compared"], 2)
        self.assertEqual(vector["wraps"]["count"], 0)
        self.assertEqual(vector["reflow"]["gt_lines"], 0)
        self.assertEqual(
            vector["topology"],
            {
                "groups": 1,
                "gt_lines": 1,
                "out_lines": 2,
                "samples": ["400 + NetProfit"],
            },
        )
        self.assertEqual(
            [item["label"] for item in vector["instances"]["large_shifts"]],
            ["400", "NetProfit"],
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 2)

    def test_distant_objects_joined_only_in_output_are_equivalent(self) -> None:
        gt = "\n".join(
            [line_of("left", 72, 99), line_of("right", 500, 101)]
        )
        out = "\n".join(
            [line_of("left", 72, 100), line_of("right", 500, 100)]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["topology"]["groups"], 1)
        self.assertEqual(vector["topology"]["gt_lines"], 2)
        self.assertEqual(vector["topology"]["out_lines"], 1)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_one_text_operation_with_a_large_gap_remains_a_wrap_finding(self) -> None:
        gt = text_op(
            [(char, 72 + index * 6) for index, char in enumerate("left")]
            + [(char, 500 + index * 6) for index, char in enumerate("right")],
            baseline_y=100,
        )
        out = "\n".join(
            [line_of("left", 72, 99), line_of("right", 500, 101)]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["topology"]["groups"], 0)
        self.assertEqual(vector["wraps"]["count"], 1)
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_split_join_segments_keep_independent_visibility_audits(self) -> None:
        gt = "\n".join(
            [
                line_of("left", 72, 100),
                line_of("right", 500, 100),
                rect_op(490, 80, 570, 110),
            ]
        )
        out = "\n".join(
            [
                rect_op(490, 80, 570, 110),
                line_of("left", 72, 99),
                line_of("right", 500, 101),
            ]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["topology"]["groups"], 1)
        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(
            vector["visibility"]["mismatches"][0],
            {"label": "right", "gt": "hidden", "out": "painted"},
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_distant_split_join_does_not_hide_reordered_objects(self) -> None:
        gt = "\n".join(
            [line_of("left", 72, 100), line_of("right", 500, 100)]
        )
        out = "\n".join(
            [line_of("right", 500, 90), line_of("left", 72, 110)]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["topology"]["groups"], 0)
        self.assertGreater(compare_layout.audit_failures([vector]), 0)

    def test_distant_split_join_does_not_hide_a_duplicate_object(self) -> None:
        gt = "\n".join(
            [line_of("left", 72, 100), line_of("right", 500, 100)]
        )
        out = "\n".join(
            [
                line_of("left", 72, 99),
                line_of("right", 500, 101),
                line_of("right", 500, 120),
            ]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["topology"]["groups"], 0)
        self.assertGreater(vector["lines"]["extra"], 0)
        self.assertGreater(compare_layout.audit_failures([vector]), 0)

    def test_distant_split_join_does_not_hide_absent_text(self) -> None:
        gt = "\n".join(
            [line_of("left", 72, 100), line_of("right", 500, 100)]
        )
        out = line_of("left", 72, 99)

        vector = self.diff(gt, out)

        self.assertEqual(vector["topology"]["groups"], 0)
        self.assertEqual(vector["lines"]["missing"], 1)

    def test_real_extra_text_remains_an_audit_failure(self) -> None:
        gt = line_of("Slide", 72, 100)
        out = "\n".join(
            [line_of("Slide", 72, 100), line_of("Sensitivity:Internal", 72, 120)]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["lines"]["extra"], 1)
        self.assertIn("Sensitivity:Internal", vector["lines"]["extra_text"])
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_unmatched_text_fully_below_media_box_is_not_a_visual_finding(self) -> None:
        vector = self.diff(line_of("Outside", 72, 860), "")

        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["visibility"]["unmatched_hidden_gt"], 1)
        self.assertEqual(
            vector["visibility"]["unmatched_hidden_gt_text"], ["Outside"]
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_unmatched_text_partly_inside_media_box_remains_a_finding(self) -> None:
        vector = self.diff(line_of("A", -3, 100), "")

        self.assertEqual(vector["lines"]["missing"], 1)
        self.assertIn("A", vector["lines"]["missing_text"])

    def test_unmatched_text_fully_outside_active_clip_is_not_a_visual_finding(
        self,
    ) -> None:
        gt = clipped_text_op(line_of("Outside clip", 72, 150), 60, 80, 200, 110)

        vector = self.diff(gt, "")

        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["visibility"]["unmatched_hidden_gt"], 1)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_unmatched_text_partly_inside_active_clip_remains_a_finding(self) -> None:
        gt = clipped_text_op(
            line_of("Clipped overflow", 72, 100), 60, 95, 200, 102
        )

        vector = self.diff(gt, "")

        self.assertEqual(vector["lines"]["missing"], 1)
        self.assertIn("Clippedoverflow", vector["lines"]["missing_text"])

    def test_page_clipped_text_covered_in_its_visible_area_is_not_missing(self) -> None:
        # Mirrors issue #1416: the baseline extends below the page, while a
        # later page-sized image hides the only area that could paint.
        gt = "\n".join([line_of("Internal", 72, 846), image_op()])

        vector = self.diff(gt, image_op())

        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["visibility"]["unmatched_hidden_gt"], 1)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_unmatched_hidden_output_text_is_not_a_visual_finding(self) -> None:
        out = "\n".join([line_of("Covered", 72, 100), image_op()])

        vector = self.diff(image_op(), out)

        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["visibility"]["unmatched_hidden_out"], 1)
        self.assertEqual(
            vector["visibility"]["unmatched_hidden_out_text"], ["Covered"]
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_pop_clip_restores_page_bounds_for_following_text(self) -> None:
        gt = "\n".join(
            [
                clipped_text_op(
                    line_of("Outside clip", 72, 150), 60, 80, 200, 110
                ),
                line_of("Visible", 72, 200),
            ]
        )

        vector = self.diff(gt, line_of("Visible", 72, 200))

        self.assertEqual(vector["lines"]["matched"], 1)
        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["visibility"]["unmatched_hidden_gt"], 1)

    def test_reordered_content_is_classified_reflow_not_loss(self) -> None:
        # A table row whose cells share one baseline in GT but split across two
        # baselines in the output: same characters, different line grouping.
        gt = line_of("labelvalue", 72, 100)
        out = "\n".join([line_of("value", 120, 99), line_of("label", 72, 103)])
        vector = self.diff(gt, out)
        self.assertEqual(vector["lines"]["missing"], 0)
        self.assertEqual(vector["lines"]["extra"], 0)
        self.assertEqual(vector["reflow"]["gt_lines"], 1)
        self.assertEqual(vector["reflow"]["out_lines"], 2)
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_real_text_loss_is_still_reported_missing(self) -> None:
        gt = "\n".join([line_of("kept", 72, 100), line_of("lost", 72, 112)])
        out = "\n".join([line_of("kept", 72, 100), line_of("other", 72, 112)])
        vector = self.diff(gt, out)
        self.assertEqual(vector["lines"]["missing"], 1)
        self.assertEqual(vector["lines"]["extra"], 1)
        self.assertEqual(vector["reflow"]["gt_lines"], 0)

    def test_pitch_delta_between_consecutive_matched_lines(self) -> None:
        gt = "\n".join([line_of("one", 72, 100), line_of("two", 72, 112), line_of("three", 72, 124)])
        out = "\n".join([line_of("one", 72, 100), line_of("two", 72, 113), line_of("three", 72, 126)])
        vector = self.diff(gt, out)
        self.assertAlmostEqual(vector["pitch"]["worst_delta"], 1.0, places=3)
        self.assertEqual(vector["pitch"]["pairs"], 2)

    def test_width_drift_is_relative(self) -> None:
        gt = line_of("wide", 72, 100, pitch=6.0)
        out = line_of("wide", 72, 100, pitch=6.3)
        vector = self.diff(gt, out)
        self.assertGreater(vector["width"]["worst_pct"], 3.0)

    def test_trailing_spaces_do_not_inflate_visible_width(self) -> None:
        gt = line_of("ALL", 72, 100)
        out = line_of("ALL  ", 72, 100)

        vector = self.diff(gt, out)

        self.assertEqual(vector["lines"]["matched"], 1)
        self.assertAlmostEqual(vector["width"]["worst_pct"], 0.0, places=6)

    def test_leading_spaces_do_not_move_the_visible_start(self) -> None:
        gt = line_of("ALL", 72, 100)
        out = line_of("  ALL", 60, 100)

        vector = self.diff(gt, out)

        self.assertEqual(vector["lines"]["matched"], 1)
        self.assertAlmostEqual(vector["dx0"]["worst"], 0.0, places=6)

    def test_internal_spaces_still_contribute_to_visible_width(self) -> None:
        page = line_of("A B", 72, 100)
        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertAlmostEqual(line.width, 17.28, places=3)

    def test_whitespace_only_line_remains_excluded(self) -> None:
        page = line_of("   ", 72, 100)
        lines = compare_layout.parse_trace(trace_document(page))[0].lines

        self.assertEqual(lines, [])

    def test_rect_census_reports_count_delta(self) -> None:
        gt = "\n".join([rect_op(70, 90, 300, 91), rect_op(70, 110, 300, 111, "stroke_path")])
        out = rect_op(70, 90, 300, 91)
        vector = self.diff(gt, out)
        self.assertEqual(vector["rects"]["gt_count"], 2)
        self.assertEqual(vector["rects"]["out_count"], 1)

    def test_equal_center_rect_size_change_reports_geometry_and_fails_audit(self) -> None:
        gt = rect_op(10.0, 20.0, 110.0, 40.0)
        out = rect_op(0.0, 15.0, 120.0, 45.0)

        vector = self.diff(gt, out, large_shift=5.0)
        rects = vector["rects"]

        self.assertEqual(rects["matched"], 1)
        self.assertEqual(rects["geometry_mismatch_count"], 1)
        self.assertAlmostEqual(rects["x"]["worst"], -10.0)
        self.assertAlmostEqual(rects["y"]["worst"], -5.0)
        self.assertAlmostEqual(rects["width"]["worst"], 20.0)
        self.assertAlmostEqual(rects["height"]["worst"], 10.0)
        self.assertAlmostEqual(rects["edges"]["left"]["worst"], -10.0)
        self.assertAlmostEqual(rects["edges"]["top"]["worst"], -5.0)
        self.assertAlmostEqual(rects["edges"]["right"]["worst"], 10.0)
        self.assertAlmostEqual(rects["edges"]["bottom"]["worst"], 5.0)
        self.assertEqual(rects["geometry_mismatch_samples"][0]["kind"], "fill")
        self.assertEqual(
            rects["geometry_mismatch_samples"][0]["gt_bbox"],
            [10.0, 20.0, 110.0, 40.0],
        )
        self.assertEqual(
            rects["geometry_mismatch_samples"][0]["out_bbox"],
            [0.0, 15.0, 120.0, 45.0],
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_translated_equal_size_rect_reports_position_and_edge_deltas(self) -> None:
        gt = rect_op(10.0, 20.0, 110.0, 30.0, "stroke_path")
        out = rect_op(18.0, 17.0, 118.0, 27.0, "stroke_path")

        rects = self.diff(gt, out, large_shift=5.0)["rects"]

        self.assertEqual(rects["geometry_mismatch_count"], 1)
        self.assertAlmostEqual(rects["x"]["mean_abs"], 8.0)
        self.assertAlmostEqual(rects["y"]["mean_abs"], 3.0)
        self.assertAlmostEqual(rects["width"]["worst"], 0.0)
        self.assertAlmostEqual(rects["height"]["worst"], 0.0)
        self.assertAlmostEqual(rects["edges"]["left"]["worst"], 8.0)
        self.assertAlmostEqual(rects["edges"]["right"]["worst"], 8.0)
        self.assertAlmostEqual(rects["edges"]["top"]["worst"], -3.0)
        self.assertAlmostEqual(rects["edges"]["bottom"]["worst"], -3.0)

    def test_rect_matching_uses_size_when_centers_are_identical(self) -> None:
        gt = "\n".join(
            [rect_op(0.0, 0.0, 100.0, 10.0), rect_op(30.0, 0.0, 70.0, 10.0)]
        )
        out = "\n".join(
            [rect_op(29.0, 0.0, 71.0, 10.0), rect_op(1.0, 0.0, 99.0, 10.0)]
        )

        rects = self.diff(gt, out, fine_shift=0.5)["rects"]

        self.assertEqual(rects["matched"], 2)
        self.assertEqual(rects["geometry_mismatch_count"], 2)
        self.assertEqual(
            [sample["gt_bbox"] for sample in rects["geometry_mismatch_samples"]],
            [[0.0, 0.0, 100.0, 10.0], [30.0, 0.0, 70.0, 10.0]],
        )
        self.assertEqual(
            [sample["out_bbox"] for sample in rects["geometry_mismatch_samples"]],
            [[1.0, 0.0, 99.0, 10.0], [29.0, 0.0, 71.0, 10.0]],
        )

    def test_same_color_rect_operation_splitting_is_informational(self) -> None:
        gt = rect_op(10.0, 20.0, 110.0, 21.0, color=".8 .2 .2")
        out = "\n".join(
            [
                rect_op(10.0, 20.0, 60.0, 21.0, color=".8 .2 .2"),
                rect_op(60.0, 20.0, 110.0, 21.0, color=".8 .2 .2"),
            ]
        )

        vector = self.diff(gt, out, fine_shift=0.5)
        rects = vector["rects"]

        self.assertEqual((rects["gt_count"], rects["out_count"]), (1, 2))
        self.assertEqual(
            (rects["canonical_gt_count"], rects["canonical_out_count"]),
            (1, 1),
        )
        self.assertEqual(rects["matched"], 1)
        self.assertEqual(rects["geometry_mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_boundary_bleed_completes_nominal_rectangle_coverage(self) -> None:
        gt = rect_op(50.0, 54.0, 474.0, 145.0, color=".85 .71 .73")
        out = "\n".join(
            [
                rect_op(50.0, 54.0, 473.0, 144.0, color=".85 .71 .73"),
                rect_op(50.0, 143.75, 474.0, 145.0, color=".85 .71 .73"),
                rect_op(472.75, 54.0, 474.0, 145.0, color=".85 .71 .73"),
            ]
        )

        vector = self.diff(gt, out, noise_floor=0.5, fine_shift=0.5)
        rects = vector["rects"]

        self.assertEqual(
            (rects["canonical_gt_count"], rects["canonical_out_count"]),
            (1, 1),
        )
        self.assertEqual(rects["matched"], 1)
        self.assertEqual(rects["geometry_mismatch_count"], 0)

    def test_boundary_bleed_accepts_one_noise_floor_corner_gap(self) -> None:
        gt = rect_op(50.0, 144.0, 474.0, 418.0, color=".97 .94 .94")
        out = "\n".join(
            [
                rect_op(50.0, 144.0, 473.0, 417.0, color=".97 .94 .94"),
                rect_op(50.0, 416.75, 454.0, 418.0, color=".97 .94 .94"),
                rect_op(453.0, 416.75, 474.0, 418.0, color=".97 .94 .94"),
                rect_op(452.75, 145.0, 454.0, 418.0, color=".97 .94 .94"),
                rect_op(472.75, 145.0, 474.0, 418.0, color=".97 .94 .94"),
            ]
        )

        vector = self.diff(gt, out, noise_floor=0.5, fine_shift=0.5)
        rects = vector["rects"]

        self.assertEqual(
            (rects["canonical_gt_count"], rects["canonical_out_count"]),
            (1, 1),
        )
        self.assertEqual(rects["geometry_mismatch_count"], 0)

    def test_l_shaped_same_paint_component_stays_split(self) -> None:
        page = compare_layout.parse_trace(
            trace_document(
                "\n".join(
                    [
                        rect_op(0.0, 0.0, 100.0, 10.0, color=".8 .2 .2"),
                        rect_op(0.0, 10.0, 10.0, 100.0, color=".8 .2 .2"),
                    ]
                )
            )
        )[0]

        rects = compare_layout.canonical_rects(page.rects, tolerance=0.5)

        self.assertEqual(len(rects), 2)

    def test_boundary_bleed_does_not_hide_a_material_corner_gap(self) -> None:
        page = compare_layout.parse_trace(
            trace_document(
                "\n".join(
                    [
                        rect_op(0.0, 0.0, 98.0, 98.0, color=".8 .2 .2"),
                        rect_op(0.0, 98.0, 98.0, 100.0, color=".8 .2 .2"),
                        rect_op(98.0, 0.0, 100.0, 98.0, color=".8 .2 .2"),
                    ]
                )
            )
        )[0]

        rects = compare_layout.canonical_rects(page.rects, tolerance=0.5)

        self.assertEqual(len(rects), 2)
        self.assertNotIn(
            [0.0, 0.0, 100.0, 100.0], [rect.rect.bbox for rect in rects]
        )

    def test_issue_1418_page_9_title_rule_is_a_rect_geometry_failure(self) -> None:
        gt = line_op(74.183, 137.254, 339.165, 137.254)
        out = line_op(74.215, 133.679, 327.632, 133.679)

        vector = self.diff(gt, out, large_shift=5.0)
        rects = vector["rects"]

        self.assertEqual(rects["matched"], 1)
        self.assertEqual(rects["geometry_mismatch_count"], 1)
        sample = rects["geometry_mismatch_samples"][0]
        self.assertAlmostEqual(sample["dy"], -3.575, places=3)
        self.assertAlmostEqual(sample["dwidth"], -11.565, places=3)
        self.assertAlmostEqual(sample["edges"]["right"], -11.533, places=3)
        self.assertEqual(compare_layout.audit_failures([vector]), 1)
        self.assertIn("rectangle geometry", compare_layout.render_reading([vector]))

    def test_unequal_rect_groups_do_not_pair_distant_unrelated_shapes(self) -> None:
        gt = "\n".join(
            [
                rect_op(10.0, 20.0, 110.0, 30.0),
                rect_op(10.0, 50.0, 110.0, 60.0),
            ]
        )
        out = rect_op(400.0, 400.0, 500.0, 410.0)

        rects = self.diff(gt, out)["rects"]

        self.assertEqual(rects["matched"], 0)
        self.assertEqual((rects["unmatched_gt"], rects["unmatched_out"]), (2, 1))
        self.assertEqual(rects["geometry_mismatch_count"], 0)

    def test_unequal_rect_groups_do_not_pair_area_fill_with_nearby_rule(self) -> None:
        gt = "\n".join(
            [
                rect_op(0.0, 0.0, 200.0, 25.0, color=".9 .9 .9"),
                rect_op(300.0, 50.0, 366.0, 75.0, color=".2 .2 .2"),
            ]
        )
        out = "\n".join(
            [
                # Its centre is within the ambiguity radius of the first GT
                # fill, but this thin band is not the same painted area.
                rect_op(65.0, 0.5, 135.0, 1.75, color=".9 .9 .9"),
                rect_op(300.1, 50.1, 366.1, 75.1, color=".2 .2 .2"),
                # Matching geometry alone cannot override different paint.
                rect_op(0.0, 0.0, 200.0, 25.0, color=".4 .4 .4"),
                rect_op(500.0, 500.0, 600.0, 510.0, color=".9 .9 .9"),
            ]
        )

        rects = self.diff(gt, out, fine_shift=0.5)["rects"]

        self.assertEqual(rects["matched"], 1)
        self.assertEqual((rects["unmatched_gt"], rects["unmatched_out"]), (1, 3))
        self.assertEqual(rects["geometry_mismatch_count"], 0)

    def test_unequal_rect_groups_reject_incompatible_area_extents(self) -> None:
        gt = "\n".join(
            [
                rect_op(0.0, 0.0, 200.0, 25.0, color=".9 .9 .9"),
                rect_op(300.0, 50.0, 366.0, 75.0, color=".2 .2 .2"),
            ]
        )
        out = "\n".join(
            [
                # The centres are close, but doubling the height means this is
                # not defensibly the same painted area.
                rect_op(5.0, -12.5, 195.0, 37.5, color=".9 .9 .9"),
                rect_op(300.1, 50.1, 366.1, 75.1, color=".2 .2 .2"),
                rect_op(500.0, 500.0, 600.0, 510.0, color=".9 .9 .9"),
            ]
        )

        rects = self.diff(gt, out, fine_shift=0.5)["rects"]

        self.assertEqual(rects["matched"], 1)
        self.assertEqual((rects["unmatched_gt"], rects["unmatched_out"]), (1, 2))
        self.assertEqual(rects["geometry_mismatch_count"], 0)

    def test_non_rectangular_path_bounds_do_not_become_rect_geometry(self) -> None:
        def triangle(x0: float, x1: float) -> str:
            midpoint = (x0 + x1) / 2
            return "\n".join(
                [
                    '<fill_path winding="nonzero" color=".8 .2 .2" alpha="1" '
                    'transform="1 0 0 1 0 0">',
                    f'<moveto x="{x0}" y="20"/>',
                    f'<lineto x="{x1}" y="20"/>',
                    f'<lineto x="{midpoint}" y="40"/>',
                    "<closepath/>",
                    "</fill_path>",
                ]
            )

        vector = self.diff(triangle(10.0, 110.0), triangle(0.0, 120.0))
        rects = vector["rects"]

        self.assertEqual((rects["gt_count"], rects["out_count"]), (1, 1))
        self.assertEqual(
            (rects["canonical_gt_count"], rects["canonical_out_count"]),
            (0, 0),
        )
        self.assertEqual(rects["geometry_mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_later_different_fill_cutting_rule_reports_visible_fill_occlusion(
        self,
    ) -> None:
        # Issue #1475: a pale body-cell fill starts 1pt too high and is painted
        # after the long rose title rule. Every constituent rectangle exists,
        # so a rect census cannot reveal the wrong final visible colour.
        rose_rule = rect_op(
            50.0,
            143.75,
            474.0,
            145.0,
            color=".85490199 .7137255 .7294118",
        )
        native_body = rect_op(
            452.75,
            145.0,
            454.0,
            184.0,
            color=".972549 .9372549 .9411765",
        )
        overpainting_body = rect_op(
            452.75,
            144.0,
            454.0,
            184.0,
            color=".972549 .9372549 .9411765",
        )

        vector = self.diff(
            "\n".join([rose_rule, native_body]),
            "\n".join([rose_rule, overpainting_body]),
        )

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 1)
        self.assertEqual(
            vector["visible_fills"]["mismatches"][0],
            {
                "side": "output",
                "bbox": [452.75, 144.0, 454.0, 145.0],
                "area": 1.25,
                "under_color": [0.854902, 0.713726, 0.729412],
                "over_color": [0.972549, 0.937255, 0.941176],
            },
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 1)
        self.assertIn("visible fill", compare_layout.render_reading([vector]).lower())

    def test_same_color_operation_splitting_is_not_a_visible_fill_mismatch(self) -> None:
        rule_color = ".85490199 .7137255 .7294118"
        cover_color = ".972549 .9372549 .9411765"
        rule = rect_op(50.0, 143.75, 474.0, 145.0, color=rule_color)
        gt = "\n".join(
            [rule, rect_op(452.75, 144.0, 454.0, 184.0, color=cover_color)]
        )
        output = "\n".join(
            [
                rule,
                rect_op(452.75, 144.0, 453.25, 184.0, color=cover_color),
                rect_op(453.25, 144.0, 454.0, 184.0, color=cover_color),
            ]
        )

        vector = self.diff(gt, output)

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_repainting_rule_color_restores_final_visible_fill(self) -> None:
        rule = rect_op(10.0, 20.0, 110.0, 21.0, color=".8 .2 .2")
        different_cover = rect_op(50.0, 20.0, 51.0, 21.0, color=".2 .8 .2")
        restored_rule = rect_op(50.0, 20.0, 51.0, 21.0, color=".8 .2 .2")

        vector = self.diff(rule, "\n".join([rule, different_cover, restored_rule]))

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_subthreshold_fill_overlap_is_ignored_as_vector_noise(self) -> None:
        rule = rect_op(10.0, 20.0, 110.0, 21.0, color=".8 .2 .2")
        tiny_overlap = rect_op(50.0, 20.8, 50.2, 40.0, color=".2 .8 .2")

        vector = self.diff(rule, "\n".join([rule, tiny_overlap]))

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_low_color_delta_overlap_is_ignored_as_trace_noise(self) -> None:
        rule = rect_op(10.0, 20.0, 110.0, 21.0, color=".80 .20 .20")
        near_same_cover = rect_op(50.0, 20.0, 51.0, 40.0, color=".84 .23 .24")

        vector = self.diff(rule, "\n".join([rule, near_same_cover]))

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_duplicate_cover_operations_do_not_duplicate_visible_coverage(self) -> None:
        rule = rect_op(10.0, 20.0, 110.0, 21.0, color=".8 .2 .2")
        cover = rect_op(50.0, 20.0, 51.0, 40.0, color=".2 .8 .2")

        vector = self.diff("\n".join([rule, cover]), "\n".join([rule, cover, cover]))

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_matching_different_color_occlusion_is_not_a_mismatch(self) -> None:
        rule = rect_op(10.0, 20.0, 110.0, 21.0, color=".8 .2 .2")
        intentional_cover = rect_op(50.0, 20.0, 51.0, 40.0, color=".2 .8 .2")
        page = "\n".join([rule, intentional_cover])

        vector = self.diff(page, page)

        self.assertEqual(vector["visible_fills"]["mismatch_count"], 0)
        self.assertEqual(compare_layout.audit_failures([vector]), 0)

    def test_repeated_labels_are_spatially_matched_and_large_shifts_are_named(self) -> None:
        gt = "\n".join(
            [line_of("Sales", 337, 133), line_of("Sales", 553, 286)]
        )
        out = "\n".join(
            [line_of("Sales", 457, 134), line_of("Sales", 526, 286)]
        )

        vector = self.diff(gt, out, large_shift=5.0)

        self.assertEqual(vector["lines"]["matched"], 2)
        self.assertEqual(vector["instances"]["large_shift_count"], 2)
        self.assertEqual(
            [item["label"] for item in vector["instances"]["large_shifts"]],
            ["Sales [1/2]", "Sales [2/2]"],
        )
        self.assertEqual(
            [round(item["dx"]) for item in vector["instances"]["large_shifts"]],
            [120, -27],
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 2)

    def test_same_text_hidden_by_later_image_reports_visibility_mismatch(self) -> None:
        text = line_of("Slide9", 72, 100)
        gt = "\n".join([text, image_op()])
        out = "\n".join([image_op(), text])

        vector = self.diff(gt, out)

        self.assertEqual(vector["lines"]["matched"], 1)
        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(
            vector["visibility"]["mismatches"][0],
            {"label": "Slide9", "gt": "hidden", "out": "painted"},
        )
        self.assertEqual(compare_layout.audit_failures([vector]), 1)

    def test_same_text_hidden_by_later_clipped_shading_reports_visibility_mismatch(self) -> None:
        text = line_of("Covered", 72, 100)
        shade = clipped_shade_op(60, 80, 150, 110)

        vector = self.diff("\n".join([text, shade]), "\n".join([shade, text]))

        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(
            vector["visibility"]["mismatches"][0],
            {"label": "Covered", "gt": "hidden", "out": "painted"},
        )

    def test_translucent_clipped_shading_does_not_hide_text(self) -> None:
        text = line_of("Tinted", 72, 100)
        shade = clipped_shade_op(60, 80, 150, 110, alpha=0.5)

        vector = self.diff("\n".join([text, shade]), "\n".join([shade, text]))

        self.assertEqual(vector["visibility"]["mismatch_count"], 0)

    def test_noncovering_clipped_shading_does_not_hide_text(self) -> None:
        text = line_of("Clear", 72, 100)
        shade = clipped_shade_op(200, 80, 300, 110)

        vector = self.diff("\n".join([text, shade]), "\n".join([shade, text]))

        self.assertEqual(vector["visibility"]["mismatch_count"], 0)

    def test_clipped_shading_painted_before_text_does_not_hide_it(self) -> None:
        page = "\n".join([clipped_shade_op(60, 80, 150, 110), line_of("Above", 72, 100)])

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.visibility, "painted")

    def test_nonextended_or_rotated_clipped_shading_is_not_overclaimed(self) -> None:
        text = line_of("Conservative", 72, 100)
        nonextended = clipped_shade_op(60, 80, 180, 110, extend="0 0")
        rotated = clipped_shade_op(
            0,
            0,
            120,
            30,
            clip_transform="0.7071 0.7071 -0.7071 0.7071 75 65",
        )

        for shade in (nonextended, rotated):
            page = "\n".join([text, shade])
            line = compare_layout.parse_trace(trace_document(page))[0].lines[0]
            self.assertEqual(line.visibility, "painted")

    def test_bow_tie_clip_does_not_overclaim_shading_coverage(self) -> None:
        text = line_of("Crossed", 72, 100)
        bow_tie = """<clip_path winding="nonzero" transform="1 0 0 1 0 0">
          <moveto x="60" y="80"/><lineto x="150" y="110"/>
          <lineto x="150" y="80"/><lineto x="60" y="110"/><closepath/>
        </clip_path>
        <fill_shade alpha="1" transform="1 0 0 1 0 0"
          type="linear" extend="1 1" start="0 0" end="1 1"/>
        <pop_clip/>"""

        line = compare_layout.parse_trace(trace_document("\n".join([text, bow_tie])))[0].lines[0]

        self.assertEqual(line.visibility, "painted")

    def test_multiple_clip_subpaths_do_not_overclaim_shading_coverage(self) -> None:
        text = line_of("Subpaths", 72, 100)
        multiple = """<clip_path winding="nonzero" transform="1 0 0 1 0 0">
          <moveto x="60" y="80"/><lineto x="150" y="80"/>
          <lineto x="150" y="110"/><lineto x="60" y="110"/><closepath/>
          <moveto x="70" y="90"/><lineto x="80" y="90"/>
          <lineto x="80" y="100"/><lineto x="70" y="100"/><closepath/>
        </clip_path>
        <fill_shade alpha="1" transform="1 0 0 1 0 0"
          type="linear" extend="1 1" start="0 0" end="1 1"/>
        <pop_clip/>"""

        line = compare_layout.parse_trace(trace_document("\n".join([text, multiple])))[0].lines[0]

        self.assertEqual(line.visibility, "painted")

    def test_nested_shading_keeps_the_outer_clip_intersection(self) -> None:
        text = line_of("A", 72, 100)
        partial_outer = """<clip_path winding="nonzero" transform="1 0 0 1 0 0">
          <moveto x="60" y="80"/><lineto x="74" y="80"/>
          <lineto x="74" y="110"/><lineto x="60" y="110"/><closepath/>
        </clip_path>"""
        full_inner = clipped_shade_op(60, 80, 150, 110)
        page = "\n".join([text, partial_outer, full_inner, "<pop_clip/>"])

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.visibility, "painted")

    def test_unknown_outer_clip_prevents_a_nested_rectangle_from_overclaiming(self) -> None:
        text = line_of("A", 72, 100)
        triangular_outer = """<clip_path winding="nonzero" transform="1 0 0 1 0 0">
          <moveto x="60" y="80"/><lineto x="150" y="80"/>
          <lineto x="60" y="110"/><closepath/>
        </clip_path>"""
        full_inner = clipped_shade_op(60, 80, 150, 110)
        page = "\n".join([text, triangular_outer, full_inner, "<pop_clip/>"])

        line = compare_layout.parse_trace(trace_document(page))[0].lines[0]

        self.assertEqual(line.visibility, "painted")

    def test_same_color_text_on_flat_fill_reports_low_contrast(self) -> None:
        background = rect_op(0, 0, 595.2, 841.92, color=".8 .8 .8")
        gt = "\n".join(
            [
                background,
                line_of("Muted", 72, 100).replace(
                    'color="0 0 0"', 'color=".8 .8 .8"'
                ),
            ]
        )
        out = "\n".join([background, line_of("Muted", 72, 100)])

        vector = self.diff(gt, out)

        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(
            vector["visibility"]["mismatches"][0],
            {"label": "Muted", "gt": "low_contrast", "out": "painted"},
        )

    def test_later_opaque_rectangle_hides_text(self) -> None:
        text = line_of("Covered", 72, 100)
        cover = rect_op(60, 80, 150, 110)

        vector = self.diff("\n".join([text, cover]), "\n".join([cover, text]))

        self.assertEqual(vector["visibility"]["mismatch_count"], 1)
        self.assertEqual(vector["visibility"]["mismatches"][0]["gt"], "hidden")

    def test_later_translucent_rectangle_does_not_hide_text(self) -> None:
        text = line_of("Tinted", 72, 100)
        tint = rect_op(60, 80, 150, 110, alpha=0.5)

        vector = self.diff("\n".join([text, tint]), "\n".join([tint, text]))

        self.assertEqual(vector["visibility"]["mismatch_count"], 0)

    def test_minor_text_color_substitution_is_not_a_visibility_mismatch(self) -> None:
        background = rect_op(0, 0, 595.2, 841.92, color="1 1 1")
        gt = "\n".join([background, line_of("Stable", 72, 100)])
        out = "\n".join(
            [
                background,
                line_of("Stable", 72, 100).replace(
                    'color="0 0 0"', 'color=".02 .02 .02"'
                ),
            ]
        )

        vector = self.diff(gt, out)

        self.assertEqual(vector["visibility"]["mismatch_count"], 0)


class ReadingTest(unittest.TestCase):
    def test_reading_mentions_wrap_and_missing_content(self) -> None:
        gt = compare_layout.parse_trace(
            trace_document("\n".join([line_of("abcdef", 72, 100), line_of("ghi", 72, 112)]))
        )[0]
        out = compare_layout.parse_trace(
            trace_document("\n".join([line_of("abc", 72, 100), line_of("defghi", 72, 112)]))
        )[0]
        vector = compare_layout.diff_page(gt, out)
        reading = compare_layout.render_reading([vector])
        self.assertIn("wrap", reading.lower())

    def test_reading_names_a_large_repeated_label_shift(self) -> None:
        gt = compare_layout.parse_trace(
            trace_document("\n".join([line_of("Sales", 337, 133), line_of("Sales", 553, 286)]))
        )[0]
        out = compare_layout.parse_trace(
            trace_document("\n".join([line_of("Sales", 457, 134), line_of("Sales", 526, 286)]))
        )[0]

        reading = compare_layout.render_reading(
            [compare_layout.diff_page(gt, out, large_shift=5.0)]
        )

        self.assertIn("Sales [1/2]", reading)
        self.assertIn("+120.00pt", reading)

    def test_reading_explains_safe_split_join_topology(self) -> None:
        gt = compare_layout.parse_trace(
            trace_document(
                "\n".join(
                    [line_of("left", 72, 100), line_of("right", 500, 100)]
                )
            )
        )[0]
        out = compare_layout.parse_trace(
            trace_document(
                "\n".join(
                    [line_of("left", 72, 99), line_of("right", 500, 101)]
                )
            )
        )[0]

        reading = compare_layout.render_reading([compare_layout.diff_page(gt, out)])

        self.assertIn("line splits/joins", reading)
        self.assertIn("position-audited", reading)

    def test_reading_explains_unmatched_hidden_trace_lines(self) -> None:
        gt = compare_layout.parse_trace(
            trace_document(line_of("Outside", 72, 860))
        )[0]
        out = compare_layout.parse_trace(trace_document(""))[0]

        reading = compare_layout.render_reading([compare_layout.diff_page(gt, out)])

        self.assertIn("1 GT and 0 output unmatched trace line", reading)
        self.assertIn("excluded from visual missing/extra findings", reading)


class CompareLayoutCliTest(unittest.TestCase):
    def test_rect_geometry_makes_json_audit_fail(self) -> None:
        traces = [
            trace_document(rect_op(10.0, 20.0, 110.0, 30.0)),
            trace_document(rect_op(18.0, 20.0, 118.0, 30.0)),
        ]
        stdout = io.StringIO()
        with (
            patch.object(compare_layout, "run_mutool", side_effect=traces),
            patch.object(
                sys,
                "argv",
                ["compare_layout.py", "gt.pdf", "output.pdf", "--json", "--audit"],
            ),
            patch("sys.stdout", stdout),
        ):
            result = compare_layout.main()

        report = json.loads(stdout.getvalue())
        self.assertEqual(result, 1)
        self.assertEqual(report["pages"][0]["rects"]["geometry_mismatch_count"], 1)

    def test_audit_exits_nonzero_for_issue_1475_visible_fill_occlusion(self) -> None:
        rose_rule = rect_op(
            50.0,
            143.75,
            474.0,
            145.0,
            color=".85490199 .7137255 .7294118",
        )
        native_body = rect_op(
            452.75,
            145.0,
            454.0,
            184.0,
            color=".972549 .9372549 .9411765",
        )
        overpainting_body = rect_op(
            452.75,
            144.0,
            454.0,
            184.0,
            color=".972549 .9372549 .9411765",
        )
        traces = [
            trace_document("\n".join([rose_rule, native_body])),
            trace_document("\n".join([rose_rule, overpainting_body])),
        ]

        stdout = io.StringIO()
        with (
            patch.object(compare_layout, "run_mutool", side_effect=traces),
            patch.object(
                sys,
                "argv",
                ["compare_layout.py", "gt.pdf", "output.pdf", "--json", "--audit"],
            ),
            patch("sys.stdout", stdout),
        ):
            result = compare_layout.main()

        report = json.loads(stdout.getvalue())
        self.assertEqual(result, 1)
        self.assertEqual(report["pages"][0]["visible_fills"]["mismatch_count"], 1)

    def test_fine_shift_mode_makes_json_audit_fail_below_coarse_threshold(self) -> None:
        traces = [
            trace_document(line_of("moved", 72, 100)),
            trace_document(line_of("moved", 72, 100.75)),
        ]
        stdout = io.StringIO()
        with (
            patch.object(compare_layout, "run_mutool", side_effect=traces),
            patch.object(
                sys,
                "argv",
                [
                    "compare_layout.py",
                    "gt.pdf",
                    "output.pdf",
                    "--json",
                    "--audit",
                    "--fine-shift",
                    "0.5",
                ],
            ),
            patch("sys.stdout", stdout),
        ):
            result = compare_layout.main()

        report = json.loads(stdout.getvalue())
        self.assertEqual(result, 1)
        self.assertEqual(report["pages"][0]["instances"]["large_shift_count"], 0)
        self.assertEqual(report["pages"][0]["instances"]["fine_shift_count"], 1)


if __name__ == "__main__":
    unittest.main()
