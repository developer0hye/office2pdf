from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT_PATH = Path(__file__).resolve().parents[1] / "compare_pdfs.py"
SPEC = importlib.util.spec_from_file_location("compare_pdfs", SCRIPT_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def build_pdf(text: str) -> bytes:
    escaped = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    stream = f"BT /F1 24 Tf 72 720 Td ({escaped}) Tj ET".encode("latin-1")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    parts = [b"%PDF-1.4\n"]
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(sum(len(part) for part in parts))
        parts.append(f"{index} 0 obj\n".encode("ascii"))
        parts.append(obj)
        parts.append(b"\nendobj\n")
    xref_offset = sum(len(part) for part in parts)
    parts.append(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    parts.append(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        parts.append(f"{offset:010d} 00000 n \n".encode("ascii"))
    parts.append(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("ascii")
    )
    return b"".join(parts)


class ComparePdfsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.project_root = Path(self.temp_dir.name) / "project"
        for relative in [
            "原文档/docx",
            "原文档/pptx",
            "standard/docx",
            "standard/pptx",
            "target/docx",
            "target/pptx",
            "report",
        ]:
            (self.project_root / relative).mkdir(parents=True, exist_ok=True)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def write_case(self, format_name: str, stem: str, standard_text: str, target_text: str) -> None:
        (self.project_root / "原文档" / format_name / f"{stem}.{format_name}").write_bytes(b"fixture")
        (self.project_root / "standard" / format_name / f"{stem}.pdf").write_bytes(build_pdf(standard_text))
        (self.project_root / "target" / format_name / f"{stem}.pdf").write_bytes(build_pdf(target_text))

    def test_identical_pdf_pair_scores_zero_loss(self) -> None:
        self.write_case("docx", "same", "Hello identical world", "Hello identical world")
        report = MODULE.generate_report(
            project_root=self.project_root,
            dpi=72,
            output_dir=self.project_root / "report",
        )
        self.assertEqual(report["summary"]["total_cases"], 1)
        self.assertEqual(report["cases"][0]["status"], "ok")
        self.assertAlmostEqual(report["cases"][0]["overall_loss"], 0.0, places=6)

    def test_different_pdf_pair_reports_nonzero_loss(self) -> None:
        self.write_case("pptx", "different", "Alpha headline", "Beta headline")
        report = MODULE.generate_report(
            project_root=self.project_root,
            dpi=72,
            output_dir=self.project_root / "report",
        )
        case = report["cases"][0]
        self.assertEqual(case["status"], "ok")
        self.assertGreater(case["overall_loss"], 0.0)
        diff_page = self.project_root / case["artifact_dir"] / "renders" / "diff-01.ppm"
        self.assertTrue(diff_page.exists())


if __name__ == "__main__":
    unittest.main()
