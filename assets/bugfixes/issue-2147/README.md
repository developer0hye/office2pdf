# Table-style line spacing defect

Public synthetic case15 from [docxide-pdf](https://github.com/sverrejb/docxide-pdf/tree/1bb140b662a69f8786015c1b532d37d6084d0934/tests/fixtures/cases/case15), Apache-2.0. No internal document or font files are submitted.

`compare.jpg`: native Word reference on the left, main `8acc920a2088a9e1c38f3385c23f537ba49f7449` output on the right, page 1 of 1, pdftoppm at 150 DPI. Original 1275x1650 panels are retained without upscaling; JPEG quality 86, progressive, metadata stripped and 150 DPI density.

- Source SHA256: `b5b9db400d391d466e54164c0a044cd46ab040fd4459cce66accc6c5632dbad1`
- Reference SHA256: `cf483f616c2e17713274d2202373dc420e3cc8c78dcb485f399a498d66ac09da`; Creator: Microsoft Word 2024; source-aware integrity PASS.
- Output SHA256: `699990feae006dc0392386cc2ffac56f5fc28013b43c9fbe8432132c248897d4`
- Baseline converter SHA256: `41159e2a1cf26f2a90800089c6fb87e64c417db31269ac939aed3e18048f8d2b`

Reproduce with the pinned public input/reference and a CLI built from the baseline using `cargo build --locked --profile ci -p office2pdf-cli`. Run the baseline executable as `/path/to/baseline/office2pdf input.docx -o output.pdf`, `python3 scripts/check_gt_integrity.py reference.pdf --source input.docx --json`, `python3 scripts/compare_layout.py reference.pdf output.pdf --json --audit --fine-shift 0.5`, and `python3 scripts/compare_text_layer.py reference.pdf output.pdf --json`.

Model vision inspected the full native/main/current pages, their diffs and all five full-width table triplets at original 150 DPI. Automatic rows are taller than Word and progressively move later tables down (worst baseline 23.2934pt, pitch 2.4219pt); both fixed 36pt rows keep their heights. Black body text, blue bold titles, merged-cell divisions and red/black/blue border regions are present and readable. No italic, underline, CJK or dashed borders occur. This is defect evidence, not a passed fidelity audit.

Related: #2147 (TableGrid line240/auto is omitted by the table-style parser); #2148 (independent 5.4pt table-origin error); #1874 (small initial-title placement/advance differences). The character multiset and 51 matched instances agree, but normalized extraction order moves the vertically aligned second-row first-cell label earlier; that observation remains open under #2149. The historical comparison above remains the before-only defect evidence from PR #2150.

## Table-style spacing correction

`gt.jpg`, `before.jpg` and `after.jpg` retain the original full-page 150 DPI panels. The after PDF uses the same source/reference hashes above. Table-style line spacing now inherits through table styles and active conditional regions, below explicit named/direct paragraph formatting.

Fresh native comparison: all 51 text instances map, normalized extracted content matches all 493 characters, and no missing/reflowed/wrapped text, painted-text visibility mismatch or fill occlusion is reported. Mean absolute baseline displacement falls from 11.6638pt to 1.0564pt; worst displacement falls from 23.2934pt to 3.2984pt. The exact 36pt rows retain their heights. Full pages, diff, all 66 current cluster crops and five full-width GT/before/after table triplets were inspected at original 150 DPI. The strict report passes with all 66 current IDs linked to open #2147/#2148; this classifies remaining findings rather than proving visual parity.

- After PDF SHA256: `69690b9738cd9f0657d027190aa1f07d6bf175cba67424d0469adf4caefe7b85`
- After converter SHA256: `f03c573d4356e44dec578919321ff4cf265629690e690a311f549b26e92b641f`

This is not a completed fidelity audit: the material table-origin offset remains under #2148 and residual vertical placement remains under #2147 pending root isolation. Small font advances, rule quantization and initial-title metrics remain under #1874. The #2149 extraction-order observation now matches locally; publication and merge verification are still required before closing it. No geometry finding is waived as antialiasing.
