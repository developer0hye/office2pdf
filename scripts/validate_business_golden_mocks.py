#!/usr/bin/env python3
"""Validate the repository-owned business golden mock corpus."""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORPUS_ROOT = PROJECT_ROOT / "tests" / "golden_mocks" / "business"
FIXTURES_ROOT = PROJECT_ROOT / "tests" / "fixtures"
MANIFEST_PATH = CORPUS_ROOT / "manifest.json"
EXPECTED_COUNTS = {"docx": 10, "pptx": 10, "xlsx": 10}
SOURCE_SUFFIXES = {"docx": ".docx", "pptx": ".pptx", "xlsx": ".xlsx"}
PPTX_SLIDE_PATTERN = re.compile(r"ppt/slides/slide[0-9]+\.xml$")
SPREADSHEET_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def fail(message: str) -> None:
    raise AssertionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve_corpus_path(relative_path: str) -> Path:
    candidate = (CORPUS_ROOT / relative_path).resolve()
    try:
        candidate.relative_to(CORPUS_ROOT.resolve())
    except ValueError:
        fail(f"path escapes corpus root: {relative_path}")
    return candidate


def resolve_fixture_path(relative_path: str) -> Path:
    candidate = (FIXTURES_ROOT / relative_path).resolve()
    try:
        candidate.relative_to(FIXTURES_ROOT.resolve())
    except ValueError:
        fail(f"path escapes fixture root: {relative_path}")
    return candidate


def pdfinfo(path: Path) -> dict[str, str]:
    completed = subprocess.run(
        ["pdfinfo", str(path)],
        check=True,
        capture_output=True,
        text=True,
    )
    fields: dict[str, str] = {}
    for line in completed.stdout.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        fields[key.strip()] = value.strip()
    return fields


def validate_ooxml(path: Path) -> zipfile.ZipFile:
    archive = zipfile.ZipFile(path)
    bad_member = archive.testzip()
    if bad_member is not None:
        fail(f"corrupt OOXML member in {path}: {bad_member}")
    if "[Content_Types].xml" not in archive.namelist():
        fail(f"missing [Content_Types].xml in {path}")
    return archive


def xml_text(archive: zipfile.ZipFile, member_names: list[str]) -> str:
    text_parts: list[str] = []
    for member_name in member_names:
        root = ElementTree.fromstring(archive.read(member_name))
        text_parts.extend(root.itertext())
    return "".join(text_parts)


def validate_common_content(case: dict[str, object], archive: zipfile.ZipFile) -> None:
    source_format = str(case["format"])
    if source_format == "docx":
        members = [name for name in archive.namelist() if name.startswith("word/") and name.endswith(".xml")]
    elif source_format == "pptx":
        members = [name for name in archive.namelist() if PPTX_SLIDE_PATTERN.fullmatch(name)]
    else:
        members = [
            name
            for name in archive.namelist()
            if name == "xl/sharedStrings.xml" or re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", name)
        ]
    visible_text = xml_text(archive, members)
    representative_text = str(case["representative_text"])
    if representative_text not in visible_text:
        fail(f"{case['id']}: representative text not found: {representative_text}")

    all_xml = "".join(
        archive.read(name).decode("utf-8", errors="ignore")
        for name in archive.namelist()
        if name.endswith(".xml")
    )
    for font_name in case["required_fonts"]:
        if str(font_name) not in all_xml:
            fail(f"{case['id']}: required font is not declared in OOXML: {font_name}")


def validate_docx(case: dict[str, object], archive: zipfile.ZipFile) -> None:
    root = ElementTree.fromstring(archive.read("word/document.xml"))
    structure = case["structure"]
    body_paragraph_count = len(root.findall("./w:body/w:p", WORD_NS))
    table_count = len(root.findall(".//w:tbl", WORD_NS))
    image_count = len(root.findall(".//w:drawing", WORD_NS))
    page_break_count = sum(
        1
        for element in root.findall(".//w:br", WORD_NS)
        if element.attrib.get(f"{{{WORD_NS['w']}}}type") == "page"
    )
    expected = {
        "paragraph_count": body_paragraph_count,
        "table_count": table_count,
        "image_count": image_count,
        "page_break_count": page_break_count,
    }
    for field_name, actual_value in expected.items():
        if int(structure[field_name]) != actual_value:
            fail(
                f"{case['id']}: {field_name} mismatch; expected {structure[field_name]}, "
                f"found {actual_value}"
            )


def validate_pptx(case: dict[str, object], archive: zipfile.ZipFile) -> None:
    slide_names = sorted(
        name for name in archive.namelist() if PPTX_SLIDE_PATTERN.fullmatch(name)
    )
    expected_slides = int(case["structure"]["slide_count"])
    if len(slide_names) != expected_slides:
        fail(f"{case['id']}: expected {expected_slides} slides, found {len(slide_names)}")
    if expected_slides != int(case["expected_pages"]):
        fail(f"{case['id']}: slide count must match expected PDF page count")
    for slide_name in slide_names:
        root = ElementTree.fromstring(archive.read(slide_name))
        if root.attrib.get("show") == "0":
            fail(f"{case['id']}: hidden slide is not allowed: {slide_name}")

    slide_xml = "".join(archive.read(name).decode("utf-8") for name in slide_names)
    table_count = len(re.findall(r"<a:tbl(?:\s|>)", slide_xml))
    media_count = sum(1 for name in archive.namelist() if name.startswith("ppt/media/") and not name.endswith("/"))
    bullet_count = len(re.findall(r"<a:bu(?:Char|AutoNum)(?:\s|>)", slide_xml))
    round_rect_count = len(re.findall(r"<a:prstGeom[^>]+prst=\"roundRect\"", slide_xml))
    chevron_count = len(re.findall(r"<a:prstGeom[^>]+prst=\"chevron\"", slide_xml))
    structure = case["structure"]
    exact_counts = {"table_count": table_count, "media_count": media_count}
    minimum_counts = {
        "bullet_count_min": bullet_count,
        "round_rect_count_min": round_rect_count,
        "chevron_count_min": chevron_count,
    }
    for field_name, actual_value in exact_counts.items():
        if int(structure[field_name]) != actual_value:
            fail(f"{case['id']}: {field_name} mismatch")
    for field_name, actual_value in minimum_counts.items():
        declared_value = int(structure.get(field_name, 0))
        if actual_value < declared_value:
            fail(f"{case['id']}: {field_name} requires {declared_value}, found {actual_value}")


def validate_xlsx(case: dict[str, object], archive: zipfile.ZipFile) -> None:
    workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    sheet_names = [
        sheet.attrib["name"]
        for sheet in workbook_root.findall("main:sheets/main:sheet", SPREADSHEET_NS)
    ]
    expected_sheet_names = case["structure"]["sheet_order"]
    if sheet_names != expected_sheet_names:
        fail(
            f"{case['id']}: sheet order mismatch; expected {expected_sheet_names}, "
            f"found {sheet_names}"
        )
    page_sequence = case["structure"]["page_sequence"]
    if len(page_sequence) != int(case["expected_pages"]):
        fail(f"{case['id']}: page_sequence must describe every expected PDF page")

    worksheet_names = sorted(
        name
        for name in archive.namelist()
        if re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", name)
    )
    worksheet_roots = [ElementTree.fromstring(archive.read(name)) for name in worksheet_names]
    formula_count = sum(len(root.findall(".//main:f", SPREADSHEET_NS)) for root in worksheet_roots)
    merged_range_count = sum(
        len(root.findall(".//main:mergeCell", SPREADSHEET_NS)) for root in worksheet_roots
    )
    structure = case["structure"]
    if formula_count != int(structure["formula_count"]):
        fail(f"{case['id']}: formula_count mismatch")
    if merged_range_count != int(structure["merged_range_count"]):
        fail(f"{case['id']}: merged_range_count mismatch")

    worksheet_xml = "".join(archive.read(name).decode("utf-8") for name in worksheet_names)
    for rule_type in structure["conditional_formatting"]:
        if str(rule_type) not in worksheet_xml:
            fail(f"{case['id']}: conditional-formatting type not found: {rule_type}")

    page_setup = worksheet_roots[0].find("main:pageSetup", SPREADSHEET_NS)
    if page_setup is None:
        fail(f"{case['id']}: first worksheet has no pageSetup")
    if page_setup.attrib.get("orientation") != structure["orientation"]:
        fail(f"{case['id']}: orientation mismatch")
    if int(page_setup.attrib.get("fitToWidth", "0")) != int(structure["fit_to_width"]):
        fail(f"{case['id']}: fit_to_width mismatch")
    if int(page_setup.attrib.get("fitToHeight", "0")) != int(structure["fit_to_height"]):
        fail(f"{case['id']}: fit_to_height mismatch")

    declared_print_area = structure.get("print_area")
    if declared_print_area is not None:
        workbook_text = xml_text(archive, ["xl/workbook.xml"])
        if str(declared_print_area) not in workbook_text.replace("$", ""):
            fail(f"{case['id']}: print area not found: {declared_print_area}")


def validate_case(case: dict[str, object], manifest: dict[str, object]) -> None:
    case_id = str(case["id"])
    source_format = str(case["format"])
    source = resolve_corpus_path(str(case["source"]))
    expected_pdf = resolve_corpus_path(str(case["expected_pdf"]))

    if source.suffix.lower() != SOURCE_SUFFIXES[source_format]:
        fail(f"{case_id}: invalid source extension: {source.suffix}")
    if expected_pdf.suffix.lower() != ".pdf":
        fail(f"{case_id}: expected artifact must be a PDF")
    if source.stem != expected_pdf.stem:
        fail(f"{case_id}: source and expected PDF stems differ")
    if not source.is_file() or not expected_pdf.is_file():
        fail(f"{case_id}: missing source or expected PDF")
    if sha256(source) != case["source_sha256"]:
        fail(f"{case_id}: source SHA-256 mismatch")
    if sha256(expected_pdf) != case["expected_sha256"]:
        fail(f"{case_id}: expected PDF SHA-256 mismatch")

    if len(case.get("features", [])) < 3:
        fail(f"{case_id}: declare at least three concrete features")
    closest_existing_fixtures = case.get("closest_existing_fixtures", [])
    if not closest_existing_fixtures or not case.get("new_coverage"):
        fail(f"{case_id}: missing existing-fixture gap analysis")
    for fixture_path in closest_existing_fixtures:
        if not resolve_fixture_path(str(fixture_path)).is_file():
            fail(f"{case_id}: closest existing fixture not found: {fixture_path}")

    with validate_ooxml(source) as archive:
        validate_common_content(case, archive)
        if source_format == "docx":
            validate_docx(case, archive)
        elif source_format == "pptx":
            validate_pptx(case, archive)
        elif source_format == "xlsx":
            validate_xlsx(case, archive)

    info = pdfinfo(expected_pdf)
    if int(info.get("Pages", "0")) != int(case["expected_pages"]):
        fail(f"{case_id}: expected PDF page count mismatch")
    provenance = manifest["provenance"]
    export_profile = provenance["export_profiles"][source_format]
    expected_producer = str(provenance["pdf_producer"])
    if info.get("Producer") != expected_producer:
        fail(f"{case_id}: PDF producer mismatch")
    expected_creator = export_profile.get("pdf_creator")
    if expected_creator is not None and info.get("Creator") != expected_creator:
        fail(f"{case_id}: PDF creator mismatch")


def validate_directory_inventory(cases: list[dict[str, object]]) -> None:
    for source_format, expected_count in EXPECTED_COUNTS.items():
        declared_sources = {
            str(case["source"])
            for case in cases
            if str(case["format"]) == source_format
        }
        declared_pdfs = {
            str(case["expected_pdf"])
            for case in cases
            if str(case["format"]) == source_format
        }
        source_directory = CORPUS_ROOT / "sources" / source_format
        expected_directory = CORPUS_ROOT / "expected" / source_format
        actual_sources = {
            path.relative_to(CORPUS_ROOT).as_posix()
            for path in source_directory.glob(f"*{SOURCE_SUFFIXES[source_format]}")
        }
        actual_pdfs = {
            path.relative_to(CORPUS_ROOT).as_posix()
            for path in expected_directory.glob("*.pdf")
        }
        if len(actual_sources) != expected_count or actual_sources != declared_sources:
            fail(f"{source_format}: source directory does not match manifest inventory")
        if len(actual_pdfs) != expected_count or actual_pdfs != declared_pdfs:
            fail(f"{source_format}: expected-PDF directory does not match manifest inventory")


def validate_manifest() -> None:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    cases = manifest.get("cases")
    if not isinstance(cases, list):
        fail("manifest cases must be a list")

    case_ids = [str(case["id"]) for case in cases]
    if len(case_ids) != 30 or len(set(case_ids)) != 30:
        fail("manifest must contain exactly 30 unique case IDs")

    format_counts = Counter(str(case["format"]) for case in cases)
    if dict(format_counts) != EXPECTED_COUNTS:
        fail(f"expected format counts {EXPECTED_COUNTS}, found {dict(format_counts)}")

    validate_directory_inventory(cases)

    for case in cases:
        validate_case(case, manifest)

    print("validated 30 business golden mocks (10 DOCX, 10 PPTX, 10 XLSX)")


def main() -> int:
    try:
        validate_manifest()
    except (
        AssertionError,
        KeyError,
        OSError,
        ValueError,
        subprocess.CalledProcessError,
        zipfile.BadZipFile,
    ) as error:
        print(f"business golden mock validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
