#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
project_root="$(cd "$script_dir/../.." && pwd)"
corpus_root="$project_root/tests/golden_mocks/business"
stage_root="${1:-$project_root/target/business-golden-office-export}"

command -v osascript >/dev/null
command -v pdfunite >/dev/null
command -v pdfinfo >/dev/null

mkdir -p "$stage_root/docx" "$stage_root/pptx" "$stage_root/xlsx-sheets" "$stage_root/xlsx"

word_args=("$stage_root/docx")
while IFS= read -r -d '' source; do
    stem="$(basename "${source%.docx}")"
    word_args+=("$stem" "$source")
done < <(find "$corpus_root/sources/docx" -maxdepth 1 -type f -name '*.docx' -print0 | sort -z)

powerpoint_args=("$stage_root/pptx")
while IFS= read -r -d '' source; do
    stem="$(basename "${source%.pptx}")"
    powerpoint_args+=("$stem" "$source")
done < <(find "$corpus_root/sources/pptx" -maxdepth 1 -type f -name '*.pptx' -print0 | sort -z)

excel_args=("$stage_root/xlsx-sheets")
while IFS= read -r -d '' source; do
    stem="$(basename "${source%.xlsx}")"
    excel_args+=("$stem" "$source")
done < <(find "$corpus_root/sources/xlsx" -maxdepth 1 -type f -name '*.xlsx' -print0 | sort -z)

osascript "$script_dir/export_word_pdfs.applescript" "${word_args[@]}"
osascript "$script_dir/export_powerpoint_pdfs.applescript" "${powerpoint_args[@]}"
osascript "$script_dir/export_excel_pdfs.applescript" "${excel_args[@]}"

while IFS= read -r -d '' source; do
    stem="$(basename "${source%.xlsx}")"
    sheet_pdfs=("$stage_root/xlsx-sheets/$stem"-sheet-*.pdf)
    if [[ ! -e "${sheet_pdfs[0]}" ]]; then
        echo "no native Excel sheet PDFs found for $stem" >&2
        exit 1
    fi
    pdfunite "${sheet_pdfs[@]}" "$stage_root/xlsx/$stem.pdf"
done < <(find "$corpus_root/sources/xlsx" -maxdepth 1 -type f -name '*.xlsx' -print0 | sort -z)

{
    echo "exported_at=$(date -u +%Y-%m-%dT%H:%M:%SZ)"
    echo "macos=$(sw_vers -productVersion) build $(sw_vers -buildVersion)"
    for application in "Microsoft Word" "Microsoft PowerPoint" "Microsoft Excel"; do
        version="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "/Applications/$application.app/Contents/Info.plist")"
        echo "$application=$version"
    done
    find "$stage_root/docx" "$stage_root/pptx" "$stage_root/xlsx" -type f -name '*.pdf' -print0 \
        | sort -z \
        | xargs -0 shasum -a 256
} > "$stage_root/provenance.txt"

find "$stage_root/docx" "$stage_root/pptx" "$stage_root/xlsx" -type f -name '*.pdf' -print0 \
    | sort -z \
    | xargs -0 -n 1 pdfinfo >/dev/null

echo "staged native Office PDFs and provenance at $stage_root"
