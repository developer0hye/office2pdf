# docxide-pdf

- Repository: https://github.com/sverrejb/docxide-pdf
- Inspected revision: `1bb140b662a69f8786015c1b532d37d6084d0934` (0.18.4)
- License: Apache-2.0; fixture and font redistribution rights require separate review.
- Published comparison: https://sverrejb.github.io/docxide-pdf/

Direct OOXML-to-PDF renderer with Word references. Useful implementations:
`src/docx/alt_chunk.rs` (HTML/MHTML imports),
`sections.rs` (page borders, line numbering, vertical alignment), `numbering.rs`
(restart/override/legal-numbering semantics), and the preset-geometry engine.
Our Typst backend handles complex-script shaping; reproduce gaps before copying patterns.
Pinned corpus: 393 inputs/391 references; re-run current code against release-only comparisons.

```sh
python3 scripts/audit_docxide_corpus.py --corpus-root /path/to/pinned/docxide-pdf \
  --converter /path/to/current/office2pdf --output /path/to/local/audit
```

The runner includes upstream skipped cases and records missing references,
conversion failures/timeouts, source/reference/output hashes, reference metadata,
page counts, and normalized text findings. It never executes fixture generators.
Token identities are hashed; source-document text stays outside the report.
The baseline is `tests/compatibility/docxide-pdf/baseline.json`.
Initial main `39bcb289`: 390 comparisons, two missing references, one decoding
failure, 64 page-count differences, and 27 empty text layers; rerun for current results.
Custom-XML text is retained. Alternate-prefix body/header/footer runs use
URI-based parsing backported from [docx-rs#934](https://github.com/bokuweb/docx-rs/pull/934).

This is triage, not a visual pass: missing searchable text can still be painted
as an image/path, and extraction order or font availability can change scores.
Fixes need fresh native GT, all comparison axes, image inspection and issue dispositions.
HTML expansion follows its CSS mapping, using Rust [html5ever](https://github.com/servo/html5ever) and
[mailparse](https://github.com/staktrace/mailparse); docx-rs and contextual scans read the expanded XML.
Section line numbering maps starts, intervals and restarts to Typst margin labels; case70 native exports check fonts/clearance, compiled tests section suppression.
An empty body `w:framePr` paragraph leaves no flow line per `is_empty_wrapping_frame` (`src/pdf/mod.rs` at `adb8c2b8`), checked against Word's `cont-first-first-fits` and `empty-frame-control` (#2161); `section_break_spacing` is the lead for #2232.
Section alignment maps `sections.rs` properties into Typst page flow. Fresh
Word case69 probes calibrate trailing spacing and continuous-section uniformity;
`both` and mixed automatic pagination warn; mixed explicit shared pages use top.
Numbering follows placeholder formatting and legal decimalization. Native case71
inline probes check ancestor counters and literals; case72 probes calibrate
restarts, suppression, implicit counters and skipped-level origins. MS-OI29500 2.1.282
confirms Word ignores restart rules inside complete level overrides.
Marker clearance/suffixes follow `pdf/list_label.rs`. Native probes calibrate
font/size/tab/wrap differences and inline marker ascent versus body descent/leading.
Damaged-PNG recovery follows `src/pdf/images.rs`; locked `png` 0.18.1 recovers samples, calibrated by Word CRC/truncation probes.
Picture crops follow `src/docx/images.rs`; raw XML retains per-picture offsets independently of [docx-rs#936](https://github.com/bokuweb/docx-rs/pull/936). An inline picture beside text rides the line like a glyph, as `inline_line_advance` in `src/pdf/layout.rs` (`fe7b8bea`) lays out: picture plus text descent (#2227).
SmartArt follows package relationships into cached DrawingML and separate text transforms; native outlines stay, compatible others use Annex D, and unsupported outlines/path paints warn and fall back.
Do not copy its `embedded_fonts.rs` font key: reversing the mixed-endian GUID corrupts the first table record. Ours follows LibreOffice's `FontTable.cxx` (#2200).
