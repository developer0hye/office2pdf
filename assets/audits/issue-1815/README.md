# Issue #1815 current-runtime re-audit

This bundle rechecks the seven bottom-aligned Malgun Gothic rows from the synthetic XLSX fixture against the retained native Excel measurements and page image. The current main commit is `4408f9c`; its first parent is the audited runtime commit `e1537dee`. The intervening merge, PR #2030, added documentation and evidence without changing runtime source or configuration.

## Result

All seven Malgun Gothic baselines at 8–14pt now match the retained native coordinates exactly (`dy = 0.000pt`); their horizontal anchors also match. Across all 22 labels, the largest current non-space glyph-origin delta is `0.0321pt`. The seven one-point residual rows—Gulim and Batang at 8–10pt plus Arial 4pt—remain tracked by open issue #1814.

The current PDF contains one Letter page, all 22 searchable labels in order, and all 23 thin gray horizontal rules. Side-by-side 150-DPI pages and paired crops for all 22 rows show no independent page-flow, clipping, fill, or rule-geometry deviation. The 5%-fuzz pixel comparison reports 6,260 changed pixels out of 2,103,750 (`0.00297564`); visual inspection localizes these to glyph rasterization and the seven #1814 rows.

## Evidence and limits

`recheck.json` records source/runtime/PDF hashes, every label baseline and glyph-origin summary, row-rule positions, the pixel metric, and the visual dispositions. `current-output.pdf` is reproducible from the committed fixture and audited runtime. The display JPEGs are progressive quality 86 at 150 DPI; the current raster and pixel-diff PNGs preserve the exact pixels used for the AE measurement. The native page JPEG is copied byte-for-byte from the reviewed issue evidence.

The native Excel PDF bytes are unavailable locally; only its recorded SHA-256, native raster, and trace-derived baseline census remain. No Office desktop app is installed. Therefore this focused re-audit could not rerun `compare_layout.py`, `compare_render.py`, or `compare_text_layer.py` against the native PDF, or produce a fresh native Excel export. Current PDF extraction and trace measurements were checked against the retained 22-label native census. This is a recheck of the Malgun seat defect, not a whole-document equivalence claim.

## Reproduction

```sh
CARGO_TARGET_DIR=/private/tmp/office2pdf-1815-target CARGO_INCREMENTAL=0 cargo test --locked --profile ci -p office2pdf --lib small_malgun_sheet_baselines_match_native_isolated_rows -- --nocapture
CARGO_TARGET_DIR=/private/tmp/office2pdf-1815-target CARGO_INCREMENTAL=0 cargo run --locked --profile ci -p office2pdf-cli -- tests/visual_audits/issue-1815/source.xlsx -o assets/audits/issue-1815/current-output.pdf
pdftoppm -r 150 -png -singlefile assets/audits/issue-1815/current-output.pdf /tmp/issue-1815-current-page-1
magick compare -metric AE -fuzz 5% assets/audits/issue-1815/native-page-1.jpg /tmp/issue-1815-current-page-1.png assets/audits/issue-1815/pixel-diff-5pct.png
```
