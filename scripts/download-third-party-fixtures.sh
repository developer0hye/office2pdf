#!/bin/bash
# Download test fixtures from LibreOffice core and Apache POI
# Uses sparse checkout to avoid cloning full repos (~3GB+ each)
#
# Output structure:
#   tests/fixtures/docx/{libreoffice,poi}/
#   tests/fixtures/pptx/{libreoffice,poi}/
#   tests/fixtures/xlsx/{libreoffice,poi}/

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
FIXTURES_DIR="$PROJECT_ROOT/tests/fixtures"
TMP_DIR="$(mktemp -d)"

cleanup() {
  echo "Cleaning up temp directory..."
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

# ── Create output directories ────────────────────────────────────────

mkdir -p "$FIXTURES_DIR/docx/libreoffice"
mkdir -p "$FIXTURES_DIR/docx/poi"
mkdir -p "$FIXTURES_DIR/pptx/libreoffice"
mkdir -p "$FIXTURES_DIR/pptx/poi"
mkdir -p "$FIXTURES_DIR/xlsx/libreoffice"
mkdir -p "$FIXTURES_DIR/xlsx/poi"

# ── Helper: count files ──────────────────────────────────────────────

count_files() {
  local dir="$1"
  local ext="$2"
  find "$dir" -maxdepth 1 -name "*.$ext" 2>/dev/null | wc -l | tr -d ' '
}

# ════════════════════════════════════════════════════════════════════
#  1. LibreOffice Core
# ════════════════════════════════════════════════════════════════════

echo ""
echo "============================================================"
echo "  Downloading LibreOffice Core test fixtures..."
echo "  (sparse checkout - only test data directories)"
echo "============================================================"

LO_DIR="$TMP_DIR/libreoffice-core"

git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/LibreOffice/core.git "$LO_DIR" 2>&1 | tail -3

cd "$LO_DIR"
git sparse-checkout set \
  sw/qa/extras/ooxmlexport/data \
  sw/qa/extras/ooxmlimport/data \
  sd/qa/unit/data/pptx \
  sc/qa/unit/data/xlsx 2>&1 | tail -3

# Copy DOCX files (from two directories, skip non-docx)
echo "Copying LibreOffice DOCX fixtures..."
find sw/qa/extras/ooxmlexport/data -maxdepth 1 -name "*.docx" -exec cp {} "$FIXTURES_DIR/docx/libreoffice/" \;
find sw/qa/extras/ooxmlimport/data -maxdepth 1 -name "*.docx" -exec cp {} "$FIXTURES_DIR/docx/libreoffice/" \;
echo "  DOCX: $(count_files "$FIXTURES_DIR/docx/libreoffice" docx) files"

# Copy PPTX files
echo "Copying LibreOffice PPTX fixtures..."
find sd/qa/unit/data/pptx -maxdepth 1 -name "*.pptx" -exec cp {} "$FIXTURES_DIR/pptx/libreoffice/" \;
echo "  PPTX: $(count_files "$FIXTURES_DIR/pptx/libreoffice" pptx) files"

# Copy XLSX files (may have subdirectories, flatten)
echo "Copying LibreOffice XLSX fixtures..."
find sc/qa/unit/data/xlsx -maxdepth 1 -name "*.xlsx" -exec cp {} "$FIXTURES_DIR/xlsx/libreoffice/" \;
echo "  XLSX: $(count_files "$FIXTURES_DIR/xlsx/libreoffice" xlsx) files"

cd "$PROJECT_ROOT"

# ════════════════════════════════════════════════════════════════════
#  2. Apache POI
# ════════════════════════════════════════════════════════════════════

echo ""
echo "============================================================"
echo "  Downloading Apache POI test fixtures..."
echo "============================================================"

POI_DIR="$TMP_DIR/apache-poi"

git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/apache/poi.git "$POI_DIR" 2>&1 | tail -3

cd "$POI_DIR"
git sparse-checkout set \
  test-data/document \
  test-data/slideshow \
  test-data/spreadsheet 2>&1 | tail -3

# Copy DOCX files
echo "Copying Apache POI DOCX fixtures..."
find test-data/document -maxdepth 1 -name "*.docx" -exec cp {} "$FIXTURES_DIR/docx/poi/" \;
echo "  DOCX: $(count_files "$FIXTURES_DIR/docx/poi" docx) files"

# Copy PPTX files
echo "Copying Apache POI PPTX fixtures..."
find test-data/slideshow -maxdepth 1 -name "*.pptx" -exec cp {} "$FIXTURES_DIR/pptx/poi/" \;
echo "  PPTX: $(count_files "$FIXTURES_DIR/pptx/poi" pptx) files"

# Copy XLSX files
echo "Copying Apache POI XLSX fixtures..."
find test-data/spreadsheet -maxdepth 1 -name "*.xlsx" -exec cp {} "$FIXTURES_DIR/xlsx/poi/" \;
echo "  XLSX: $(count_files "$FIXTURES_DIR/xlsx/poi" xlsx) files"

cd "$PROJECT_ROOT"

# ════════════════════════════════════════════════════════════════════
#  Summary
# ════════════════════════════════════════════════════════════════════

echo ""
echo "============================================================"
echo "  Download Complete!"
echo "============================================================"
echo ""
echo "  LibreOffice:"
echo "    DOCX: $(count_files "$FIXTURES_DIR/docx/libreoffice" docx)"
echo "    PPTX: $(count_files "$FIXTURES_DIR/pptx/libreoffice" pptx)"
echo "    XLSX: $(count_files "$FIXTURES_DIR/xlsx/libreoffice" xlsx)"
echo ""
echo "  Apache POI:"
echo "    DOCX: $(count_files "$FIXTURES_DIR/docx/poi" docx)"
echo "    PPTX: $(count_files "$FIXTURES_DIR/pptx/poi" pptx)"
echo "    XLSX: $(count_files "$FIXTURES_DIR/xlsx/poi" xlsx)"
echo ""

TOTAL_DOCX=$(($(count_files "$FIXTURES_DIR/docx/libreoffice" docx) + $(count_files "$FIXTURES_DIR/docx/poi" docx)))
TOTAL_PPTX=$(($(count_files "$FIXTURES_DIR/pptx/libreoffice" pptx) + $(count_files "$FIXTURES_DIR/pptx/poi" pptx)))
TOTAL_XLSX=$(($(count_files "$FIXTURES_DIR/xlsx/libreoffice" xlsx) + $(count_files "$FIXTURES_DIR/xlsx/poi" xlsx)))
GRAND_TOTAL=$((TOTAL_DOCX + TOTAL_PPTX + TOTAL_XLSX))

echo "  Total new fixtures: $GRAND_TOTAL"
echo "    DOCX: $TOTAL_DOCX  |  PPTX: $TOTAL_PPTX  |  XLSX: $TOTAL_XLSX"
echo ""
echo "  Location: tests/fixtures/{docx,pptx,xlsx}/{libreoffice,poi}/"
