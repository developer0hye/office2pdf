#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import difflib
import json
import subprocess
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any

FORMATS: tuple[str, ...] = ("docx", "pptx")
ORIGINAL_DIRNAME = "原文档"
STANDARD_DIRNAME = "standard"
TARGET_DIRNAME = "target"
REPORT_DIRNAME = "report"


class CompareError(RuntimeError):
    pass


@dataclass(frozen=True)
class CompareCase:
    format_name: str
    source_path: Path
    relative_stem: Path
    standard_pdf: Path
    target_pdf: Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compare PDFs under project/standard and project/target using the "
            "project/原文档 layout."
        )
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Project directory containing 原文档/, standard/, target/, report/.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=160,
        help="Rendering DPI for pixel comparison (default: 160).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory for report files (default: <project-root>/report).",
    )
    return parser.parse_args()


def require_binary(name: str) -> None:
    result = subprocess.run(
        ["/usr/bin/env", "bash", "-lc", f"command -v {name}"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode != 0:
        raise CompareError(f"Required binary not found: {name}")


def ensure_layout(project_root: Path) -> None:
    required_dirs = [
        project_root / ORIGINAL_DIRNAME / "docx",
        project_root / ORIGINAL_DIRNAME / "pptx",
        project_root / STANDARD_DIRNAME / "docx",
        project_root / STANDARD_DIRNAME / "pptx",
        project_root / TARGET_DIRNAME / "docx",
        project_root / TARGET_DIRNAME / "pptx",
        project_root / REPORT_DIRNAME,
    ]
    missing = [str(path) for path in required_dirs if not path.exists()]
    if missing:
        raise CompareError("Missing required directories:\n" + "\n".join(missing))


def relative_display(path: Path, root: Path) -> str:
    return str(path.relative_to(root))


def discover_cases(project_root: Path) -> list[CompareCase]:
    cases: list[CompareCase] = []
    original_root = project_root / ORIGINAL_DIRNAME
    standard_root = project_root / STANDARD_DIRNAME
    target_root = project_root / TARGET_DIRNAME

    for format_name in FORMATS:
        source_root = original_root / format_name
        for source_path in sorted(source_root.rglob(f"*.{format_name}")):
            relative_source = source_path.relative_to(source_root)
            relative_stem = relative_source.with_suffix("")
            standard_pdf = standard_root / format_name / relative_stem.with_suffix(".pdf")
            target_pdf = target_root / format_name / relative_stem.with_suffix(".pdf")
            cases.append(
                CompareCase(
                    format_name=format_name,
                    source_path=source_path,
                    relative_stem=relative_stem,
                    standard_pdf=standard_pdf,
                    target_pdf=target_pdf,
                )
            )
    return cases


def run_capture(command: list[str]) -> str:
    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode != 0:
        raise CompareError(
            f"Command failed ({result.returncode}): {' '.join(command)}\n{result.stderr.strip()}"
        )
    return result.stdout


def parse_pdf_pages(pdf_path: Path) -> int:
    text = run_capture(["pdfinfo", str(pdf_path)])
    for line in text.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    raise CompareError(f"Could not read page count from {pdf_path}")


def extract_text(pdf_path: Path) -> str:
    return run_capture(["pdftotext", "-enc", "UTF-8", str(pdf_path), "-"])


def extract_word_boxes(pdf_path: Path) -> dict[int, list[dict[str, Any]]]:
    text = run_capture(["pdftotext", "-bbox-layout", "-enc", "UTF-8", str(pdf_path), "-"])
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        raise CompareError(f"Could not parse bbox output for {pdf_path}: {exc}") from exc

    page_words: dict[int, list[dict[str, Any]]] = {}
    page_index = 0
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "page":
            page_index += 1
            page_words.setdefault(page_index, [])
        elif tag == "word" and page_index > 0:
            content = " ".join((element.text or "").split())
            if not content:
                continue
            attrs = element.attrib
            try:
                bbox = (
                    float(attrs["xMin"]),
                    float(attrs["yMin"]),
                    float(attrs["xMax"]),
                    float(attrs["yMax"]),
                )
            except KeyError:
                continue
            page_words.setdefault(page_index, []).append({"text": content, "bbox": bbox})

    for words in page_words.values():
        words.sort(key=lambda item: (item["bbox"][1], item["bbox"][0], item["text"]))
    return page_words


def bbox_center(box: tuple[float, float, float, float]) -> tuple[float, float]:
    return ((box[0] + box[2]) * 0.5, (box[1] + box[3]) * 0.5)


def bbox_error(left: tuple[float, float, float, float], right: tuple[float, float, float, float]) -> float:
    return (
        abs(left[0] - right[0])
        + abs(left[1] - right[1])
        + abs(left[2] - right[2])
        + abs(left[3] - right[3])
    ) / 4.0


def compare_word_sets(
    standard_words: list[dict[str, Any]],
    target_words: list[dict[str, Any]],
) -> tuple[float, float, int]:
    if not standard_words and not target_words:
        return 1.0, 1.0, 0

    target_by_text: dict[str, list[int]] = {}
    for index, word in enumerate(target_words):
        target_by_text.setdefault(word["text"], []).append(index)

    used_indexes: set[int] = set()
    bbox_errors: list[float] = []
    matched = 0

    for standard_word in standard_words:
        candidates = target_by_text.get(standard_word["text"], [])
        if not candidates:
            continue

        sx, sy = bbox_center(standard_word["bbox"])
        best_index: int | None = None
        best_distance: float | None = None
        for candidate_index in candidates:
            if candidate_index in used_indexes:
                continue
            tx, ty = bbox_center(target_words[candidate_index]["bbox"])
            distance = ((sx - tx) ** 2 + (sy - ty) ** 2) ** 0.5
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_index = candidate_index

        if best_index is None:
            continue

        used_indexes.add(best_index)
        matched += 1
        bbox_errors.append(
            bbox_error(standard_word["bbox"], target_words[best_index]["bbox"])
        )

    recall = matched / len(standard_words) if standard_words else 1.0
    precision = matched / len(target_words) if target_words else 1.0
    if recall + precision == 0:
        word_f1 = 0.0
    else:
        word_f1 = (2.0 * recall * precision) / (recall + precision)

    if bbox_errors:
        average_error = sum(bbox_errors) / len(bbox_errors)
        word_bbox_score = max(0.0, 1.0 - min(1.0, average_error / 40.0))
    else:
        word_bbox_score = 0.0

    return word_f1, word_bbox_score, matched


def read_ppm(ppm_path: Path) -> tuple[int, int, bytes]:
    raw = ppm_path.read_bytes()
    if not raw.startswith(b"P6"):
        raise CompareError(f"Unsupported PPM format: {ppm_path}")

    index = 2
    tokens: list[bytes] = []
    while len(tokens) < 3:
        while index < len(raw) and raw[index] in b" \t\r\n":
            index += 1
        if index >= len(raw):
            raise CompareError(f"Truncated PPM header: {ppm_path}")
        if raw[index] == 35:
            while index < len(raw) and raw[index] not in b"\r\n":
                index += 1
            continue
        end = index
        while end < len(raw) and raw[end] not in b" \t\r\n":
            end += 1
        tokens.append(raw[index:end])
        index = end

    width = int(tokens[0])
    height = int(tokens[1])
    max_value = int(tokens[2])
    if max_value != 255:
        raise CompareError(f"Unsupported PPM max value {max_value}: {ppm_path}")

    while index < len(raw) and raw[index] in b" \t\r\n":
        index += 1

    pixel_bytes = raw[index:]
    expected = width * height * 3
    if len(pixel_bytes) < expected:
        raise CompareError(f"Truncated PPM data: {ppm_path}")
    return width, height, pixel_bytes[:expected]


def write_ppm(ppm_path: Path, width: int, height: int, pixel_bytes: bytes) -> None:
    header = f"P6\n{width} {height}\n255\n".encode("ascii")
    ppm_path.write_bytes(header + pixel_bytes)


def render_pdf_to_ppm(pdf_path: Path, output_dir: Path, prefix: str, dpi: int) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_prefix = output_dir / prefix
    subprocess.run(
        ["pdftoppm", "-r", str(dpi), str(pdf_path), str(output_prefix)],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    pages = sorted(output_dir.glob(f"{prefix}-*.ppm"))
    if not pages:
        raise CompareError(f"No rendered pages produced for {pdf_path}")
    return pages


def compare_ppm_pages(
    standard_ppm: Path,
    target_ppm: Path,
    diff_ppm: Path,
) -> float:
    standard_width, standard_height, standard_pixels = read_ppm(standard_ppm)
    target_width, target_height, target_pixels = read_ppm(target_ppm)
    width = max(standard_width, target_width)
    height = max(standard_height, target_height)

    total = width * height * 3
    diff_bytes = bytearray(total)
    total_abs = 0

    for y in range(height):
        for x in range(width):
            for channel in range(3):
                output_index = (y * width + x) * 3 + channel
                if x < standard_width and y < standard_height:
                    standard_index = (y * standard_width + x) * 3 + channel
                    standard_value = standard_pixels[standard_index]
                else:
                    standard_value = 255
                if x < target_width and y < target_height:
                    target_index = (y * target_width + x) * 3 + channel
                    target_value = target_pixels[target_index]
                else:
                    target_value = 255
                delta = abs(standard_value - target_value)
                total_abs += delta
                if channel == 0:
                    diff_bytes[output_index] = delta
                else:
                    diff_bytes[output_index] = 0

    write_ppm(diff_ppm, width, height, bytes(diff_bytes))
    return total_abs / (width * height * 3 * 255)


def markdown_escape(value: str) -> str:
    return value.replace("|", "\\|")


def compare_case(case: CompareCase, project_root: Path, output_dir: Path, dpi: int) -> dict[str, Any]:
    source_display = relative_display(case.source_path, project_root)
    standard_display = relative_display(case.standard_pdf, project_root)
    target_display = relative_display(case.target_pdf, project_root)
    artifact_dir = output_dir / "artifacts" / case.format_name / case.relative_stem
    artifact_dir.mkdir(parents=True, exist_ok=True)

    if not case.standard_pdf.exists():
        return {
            "format": case.format_name,
            "source": source_display,
            "standard_pdf": standard_display,
            "target_pdf": target_display,
            "status": "missing_standard",
            "artifact_dir": relative_display(artifact_dir, project_root),
        }

    if not case.target_pdf.exists():
        return {
            "format": case.format_name,
            "source": source_display,
            "standard_pdf": standard_display,
            "target_pdf": target_display,
            "status": "missing_target",
            "artifact_dir": relative_display(artifact_dir, project_root),
        }

    standard_pages = parse_pdf_pages(case.standard_pdf)
    target_pages = parse_pdf_pages(case.target_pdf)

    standard_text = extract_text(case.standard_pdf)
    target_text = extract_text(case.target_pdf)
    text_similarity = difflib.SequenceMatcher(None, standard_text, target_text).ratio()
    (artifact_dir / "standard.txt").write_text(standard_text, encoding="utf-8")
    (artifact_dir / "target.txt").write_text(target_text, encoding="utf-8")

    standard_words_by_page = extract_word_boxes(case.standard_pdf)
    target_words_by_page = extract_word_boxes(case.target_pdf)

    page_limit = max(standard_pages, target_pages)
    page_details: list[dict[str, Any]] = []
    word_f1_values: list[float] = []
    word_bbox_values: list[float] = []

    render_dir = artifact_dir / "renders"
    standard_ppms = render_pdf_to_ppm(case.standard_pdf, render_dir, "standard", dpi)
    target_ppms = render_pdf_to_ppm(case.target_pdf, render_dir, "target", dpi)

    pixel_losses: list[float] = []
    for page_number in range(1, page_limit + 1):
        standard_page_words = standard_words_by_page.get(page_number, [])
        target_page_words = target_words_by_page.get(page_number, [])
        word_f1, word_bbox_score, matched_words = compare_word_sets(
            standard_page_words,
            target_page_words,
        )
        word_f1_values.append(word_f1)
        word_bbox_values.append(word_bbox_score)

        if page_number <= len(standard_ppms) and page_number <= len(target_ppms):
            diff_ppm = render_dir / f"diff-{page_number:02d}.ppm"
            pixel_loss = compare_ppm_pages(
                standard_ppms[page_number - 1],
                target_ppms[page_number - 1],
                diff_ppm,
            )
        else:
            pixel_loss = 1.0
        pixel_losses.append(pixel_loss)

        page_details.append(
            {
                "page": page_number,
                "pixel_loss": pixel_loss,
                "word_f1": word_f1,
                "word_bbox_score": word_bbox_score,
                "matched_words": matched_words,
                "standard_words": len(standard_page_words),
                "target_words": len(target_page_words),
            }
        )

    mean_pixel_loss = sum(pixel_losses) / len(pixel_losses) if pixel_losses else 1.0
    mean_word_f1 = sum(word_f1_values) / len(word_f1_values) if word_f1_values else 0.0
    mean_word_bbox = sum(word_bbox_values) / len(word_bbox_values) if word_bbox_values else 0.0
    element_score = 0.65 * mean_word_f1 + 0.35 * mean_word_bbox
    overall_loss = 0.5 * mean_pixel_loss + 0.5 * (1.0 - element_score)

    worst_page = max(page_details, key=lambda item: item["pixel_loss"])

    return {
        "format": case.format_name,
        "source": source_display,
        "standard_pdf": standard_display,
        "target_pdf": target_display,
        "status": "ok",
        "page_count_standard": standard_pages,
        "page_count_target": target_pages,
        "text_length_standard": len(standard_text),
        "text_length_target": len(target_text),
        "text_similarity": text_similarity,
        "word_f1": mean_word_f1,
        "word_bbox_score": mean_word_bbox,
        "pixel_mean_loss": mean_pixel_loss,
        "overall_loss": overall_loss,
        "worst_page": worst_page,
        "artifact_dir": relative_display(artifact_dir, project_root),
        "page_details": page_details,
    }


def build_markdown_report(report: dict[str, Any]) -> str:
    lines: list[str] = []
    lines.append("# PDF Comparison Report")
    lines.append("")
    lines.append(f"Generated at: {report['generated_at']}")
    lines.append("")
    lines.append("## Summary")
    lines.append("")
    summary = report["summary"]
    lines.append(f"- Total cases: {summary['total_cases']}")
    lines.append(f"- Comparable cases: {summary['comparable_cases']}")
    lines.append(f"- Missing standard PDFs: {summary['missing_standard']}")
    lines.append(f"- Missing target PDFs: {summary['missing_target']}")
    lines.append(f"- Dataset score: {summary['dataset_score']:.6f}")
    lines.append(f"- Dataset loss: {summary['dataset_loss']:.6f}")
    lines.append("")
    lines.append("## Worst Cases")
    lines.append("")
    lines.append("| Source | Format | Loss | Pixel | Word F1 | Text Ratio | Status |")
    lines.append("| --- | --- | ---: | ---: | ---: | ---: | --- |")
    for case in summary["worst_cases"]:
        lines.append(
            "| {source} | {format_name} | {overall_loss:.6f} | {pixel_loss:.6f} | {word_f1:.6f} | {text_similarity:.6f} | {status} |".format(
                source=markdown_escape(case["source"]),
                format_name=case["format"],
                overall_loss=case.get("overall_loss", 1.0),
                pixel_loss=case.get("pixel_mean_loss", 1.0),
                word_f1=case.get("word_f1", 0.0),
                text_similarity=case.get("text_similarity", 0.0),
                status=case["status"],
            )
        )
    lines.append("")
    lines.append("## Per-file Details")
    lines.append("")
    for case in report["cases"]:
        lines.append(f"### {case['source']}")
        lines.append("")
        lines.append(f"- Status: {case['status']}")
        lines.append(f"- Standard PDF: {case['standard_pdf']}")
        lines.append(f"- Target PDF: {case['target_pdf']}")
        lines.append(f"- Artifact dir: {case['artifact_dir']}")
        if case["status"] == "ok":
            lines.append(f"- Overall loss: {case['overall_loss']:.6f}")
            lines.append(f"- Pixel mean loss: {case['pixel_mean_loss']:.6f}")
            lines.append(f"- Word F1: {case['word_f1']:.6f}")
            lines.append(f"- Word bbox score: {case['word_bbox_score']:.6f}")
            lines.append(f"- Text similarity: {case['text_similarity']:.6f}")
            lines.append(
                f"- Page count: {case['page_count_standard']} / {case['page_count_target']}"
            )
            lines.append(
                f"- Worst page: {case['worst_page']['page']} (pixel={case['worst_page']['pixel_loss']:.6f})"
            )
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def generate_report(project_root: Path, dpi: int, output_dir: Path) -> dict[str, Any]:
    ensure_layout(project_root)
    for binary in ("pdfinfo", "pdftoppm", "pdftotext"):
        require_binary(binary)

    cases = discover_cases(project_root)
    case_reports = [compare_case(case, project_root, output_dir, dpi) for case in cases]

    comparable_cases = [case for case in case_reports if case["status"] == "ok"]
    dataset_loss = (
        sum(case["overall_loss"] for case in comparable_cases) / len(comparable_cases)
        if comparable_cases
        else 1.0
    )
    worst_cases = sorted(
        case_reports,
        key=lambda item: item.get("overall_loss", 1.0),
        reverse=True,
    )[:10]

    report = {
        "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(),
        "project_root": str(project_root),
        "summary": {
            "total_cases": len(case_reports),
            "comparable_cases": len(comparable_cases),
            "missing_standard": sum(1 for case in case_reports if case["status"] == "missing_standard"),
            "missing_target": sum(1 for case in case_reports if case["status"] == "missing_target"),
            "dataset_loss": dataset_loss,
            "dataset_score": max(0.0, 1.0 - dataset_loss),
            "worst_cases": worst_cases,
        },
        "cases": case_reports,
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (output_dir / "report.md").write_text(
        build_markdown_report(report),
        encoding="utf-8",
    )
    return report


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    output_dir = (args.output_dir or (project_root / REPORT_DIRNAME)).resolve()
    try:
        report = generate_report(project_root=project_root, dpi=args.dpi, output_dir=output_dir)
    except CompareError as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    except subprocess.CalledProcessError as exc:
        print(f"[error] Subprocess failed: {exc}", file=sys.stderr)
        return 1

    print("[ok] comparison completed")
    print(f"dataset_score: {report['summary']['dataset_score']:.6f}")
    print(f"dataset_loss: {report['summary']['dataset_loss']:.6f}")
    print(f"report_json: {output_dir / 'report.json'}")
    print(f"report_md: {output_dir / 'report.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
