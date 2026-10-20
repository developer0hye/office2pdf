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
An empty body `w:framePr` paragraph leaves no flow line per `is_empty_wrapping_frame` (`src/pdf/mod.rs` at `adb8c2b8`), checked against Word's `cont-first-first-fits` and `empty-frame-control` (#2161), and for a group of them sharing one frame against `empty-margin-frame-14-*` (#2231); in compatibility mode 15 such a group's frame at the top margin keeps its box, which a full-width table on its page starts below, per `empty_frame_table_bands` (`fe7b8bea`; ours does not stop at visible text before the group, which Word's `empty-margin-frame-visible-body-before-frame` moves too), checked against `empty-margin-frame-15-*` (#2231); its `fits` check reserves no box whose group may land past the section's first page, so it leaves the rows `empty-margin-frame-long-prefix` carries onto page 2 at the top margin, where Word and ours start them below the box (#2231); a text-anchored frame with a stated height and a nonzero `w:y` is lifted into its frame, which text wraps around, per `paragraph_offset_frame` (`fe7b8bea`), checked against Word's `paragraph-offset-frame-*` (#2231); an empty section mark keeps its line only as its section's first block, or in compatibility mode 15 after an inline table before a continuous section, per `keeps_section_mark_line` and `render_paragraph_block` (`fe7b8bea`), checked against Word's `page-header-ordinary-marker`, `paragraph-offset-frame-no-frame` and `section-mark-after-table-*` (#2232); its `section_break_spacing` is the lead for the mark-`w:after` absorption #2232 still owes. Body text beside a `w:tblpPr` table follows `FloatZone` there: a paragraph is checked at its first line and restarts below the table when the room beside it is under 72pt (`min_wrap_w`, empirical; combined room for a `BothSides` zone, the wider side for `Largest`). Ours always takes the wider side (#2229), and keeps the table's extent in force in a following continuous section, as Word's `section-mark-after-table-float-*` show (#2369). Its `has_clear_break` (`src/docx/runs.rs`) marks a paragraph holding a `w:clear="all"` break; ours keeps each clearing break's position and side, and starts the text after it below the tables it clears (#2230).
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
