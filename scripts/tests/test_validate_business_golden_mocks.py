import copy
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from scripts import validate_business_golden_mocks as validator


class BusinessGoldenMockValidatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(validator.MANIFEST_PATH.read_text(encoding="utf-8"))

    def setUp(self) -> None:
        self.original_corpus_root = validator.CORPUS_ROOT
        self.original_manifest_path = validator.MANIFEST_PATH

    def tearDown(self) -> None:
        validator.CORPUS_ROOT = self.original_corpus_root
        validator.MANIFEST_PATH = self.original_manifest_path

    def test_repository_corpus_passes(self) -> None:
        validator.validate_manifest()

    def test_rejects_wrong_case_count(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            manifest_path = Path(temporary_directory) / "manifest.json"
            manifest_path.write_text(json.dumps({"cases": []}), encoding="utf-8")
            validator.MANIFEST_PATH = manifest_path

            with self.assertRaisesRegex(AssertionError, "exactly 30 unique"):
                validator.validate_manifest()

    def test_rejects_duplicate_case_ids(self) -> None:
        malformed = copy.deepcopy(self.manifest)
        malformed["cases"][1]["id"] = malformed["cases"][0]["id"]
        with tempfile.TemporaryDirectory() as temporary_directory:
            manifest_path = Path(temporary_directory) / "manifest.json"
            manifest_path.write_text(json.dumps(malformed), encoding="utf-8")
            validator.MANIFEST_PATH = manifest_path

            with self.assertRaisesRegex(AssertionError, "exactly 30 unique"):
                validator.validate_manifest()

    def test_rejects_unsafe_path(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            validator.CORPUS_ROOT = Path(temporary_directory)
            with self.assertRaisesRegex(AssertionError, "escapes corpus root"):
                validator.resolve_corpus_path("../outside.docx")

    def test_rejects_manifest_file_stem_mismatch(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        case["expected_pdf"] = self.manifest["cases"][1]["expected_pdf"]

        with self.assertRaisesRegex(AssertionError, "stems differ"):
            validator.validate_case(case, self.manifest)

    def test_rejects_source_hash_mismatch(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        case["source_sha256"] = "0" * 64

        with self.assertRaisesRegex(AssertionError, "source SHA-256 mismatch"):
            validator.validate_case(case, self.manifest)

    def test_rejects_corrupt_ooxml(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        with tempfile.TemporaryDirectory() as temporary_directory:
            corpus_root = Path(temporary_directory)
            source = corpus_root / "sources" / "docx" / "corrupt.docx"
            expected = corpus_root / "expected" / "docx" / "corrupt.pdf"
            source.parent.mkdir(parents=True)
            expected.parent.mkdir(parents=True)
            source.write_bytes(b"not an OOXML archive")
            shutil.copy2(
                self.original_corpus_root / case["expected_pdf"],
                expected,
            )
            case["source"] = "sources/docx/corrupt.docx"
            case["expected_pdf"] = "expected/docx/corrupt.pdf"
            case["source_sha256"] = validator.sha256(source)
            case["expected_sha256"] = validator.sha256(expected)
            validator.CORPUS_ROOT = corpus_root

            with self.assertRaisesRegex(Exception, "File is not a zip file"):
                validator.validate_case(case, self.manifest)

    def test_rejects_unreadable_pdf(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        with tempfile.TemporaryDirectory() as temporary_directory:
            corpus_root = Path(temporary_directory)
            source = corpus_root / "sources" / "docx" / "unreadable.docx"
            expected = corpus_root / "expected" / "docx" / "unreadable.pdf"
            source.parent.mkdir(parents=True)
            expected.parent.mkdir(parents=True)
            shutil.copy2(self.original_corpus_root / case["source"], source)
            expected.write_bytes(b"not a PDF")
            case["source"] = "sources/docx/unreadable.docx"
            case["expected_pdf"] = "expected/docx/unreadable.pdf"
            case["source_sha256"] = validator.sha256(source)
            case["expected_sha256"] = validator.sha256(expected)
            validator.CORPUS_ROOT = corpus_root

            with self.assertRaises(subprocess.CalledProcessError):
                validator.validate_case(case, self.manifest)

    def test_rejects_missing_representative_content(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        case["representative_text"] = "text that is intentionally absent"

        with self.assertRaisesRegex(AssertionError, "representative text not found"):
            validator.validate_case(case, self.manifest)

    def test_rejects_missing_closest_existing_fixture(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        case["closest_existing_fixtures"] = ["docx/does-not-exist.docx"]

        with self.assertRaisesRegex(AssertionError, "closest existing fixture not found"):
            validator.validate_case(case, self.manifest)

    def test_rejects_incorrect_feature_inventory(self) -> None:
        case = copy.deepcopy(self.manifest["cases"][0])
        case["structure"]["table_count"] += 1

        with self.assertRaisesRegex(AssertionError, "table_count mismatch"):
            validator.validate_case(case, self.manifest)


if __name__ == "__main__":
    unittest.main()
