# Custom footnote marks: native audit

- Input: `tests/fixtures/docx/docxide/custom_footnote_marks.docx`, independently authored.
- GT: Microsoft Word 16.113.4, exported from its own internal container;
  pdfinfo reports Quartz PDFContext and one letter-size page. GT integrity passes.
- Before: office2pdf 456e9338463345f405560efa831e495d1d9b29b1.
- After: office2pdf 02d44bb408442712d8938e1dceddc652599827a3.
- All page images are direct 150 DPI Poppler renders, progressive quality 86 JPEGs.
- No-patch re-zip, `1`/`true` and same-run/next-run/merged-text controls: zero decoded pixel differences.
- TDD: custom-numbering regression fails before repair; after repair all 4370 Rust
  tests and 12 Node/Wasm tests pass, strict Clippy passes.
- Source cause: the reader dropped `w:customMarkFollows`; the parser then rendered
  each custom note as an automatic reference. The fixed reader retains the flag.
- Current normalized searchable text equals GT; all 8 text instances match.
- Four automatic numeric labels (1/2 in body and entries) replace the previous
  six labels (1/2/3/4 in body,2/4 in entries); custom text remains in both places.
- Six PDF links resolve to the single page: four body-to-note destinations and
  two automatic-entry backlinks. Native link rectangles are not established
  by the authored topology probes.

## Model vision findings

The full GT, before, after and pixel diffs show the extra numeric body labels
and numbering gaps gone. Source `**` and `†`, baseline alignment, body font and all
trailing text survive. Full-resolution matched crops and each of seven current
cluster crops show residual note-area placement and indentation differences.
Note entries remain inset 9.35–9.82pt; their baselines are 10.08–30.65pt below GT.
All seven exact current cluster IDs are assigned to open #2138 in the strict
report. No exporter differences are approved. Note paragraph/area layout remains
unfinished; this fixture is not declared fully equivalent. Mutool trace measures
the separator at 144pt by 0.48pt in GT versus 140.4pt by 0.55pt in output; the
3.6pt length and 0.07pt thickness differences are sub-material under #1874.

## Corpus scope

All 393 inputs convert; only the custom-mark input changes versus the marker
base. Its missing/extra tokens fall 2/2 to 0/0; normalized order still differs.
The other 392 output PDF hashes are unchanged. The corpus retains 63 page-count
mismatches and 160 normalized matches. Its changed file's full page and diff were
inspected; matched crops, exact-ID cluster review and separate header/table/wrap/
footer defects remain pending. Other draft-PR corpus image audits are pending.

## Provenance hashes

Input SHA256: `90ef6de0377b88e4c8495c21b5fcf4ac1cee055ad2f70d2ec8c972d11c88348b`.

- probe/pdfs/base.pdf: `c068ebf5c6c9cba9e9261a80e3a28e468edfb0f3ef61f29d2d70c49940c7cbd8`
- before-marker-repair.pdf: `4889febe57073089565c9607053a96953d4bb5660c53cbee91ad85eaaed3ccd2`
- after.pdf: `09bfe4d4d730f1ce0bfeea0273b32fcff8fc47a8b6b03bce9899007821c85ea7`
