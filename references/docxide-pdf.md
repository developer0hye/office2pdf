# docxide-pdf

- Repository: https://github.com/sverrejb/docxide-pdf
- Inspected revision: `1bb140b662a69f8786015c1b532d37d6084d0934` (0.18.4)
- License: Apache-2.0; fixture and font redistribution rights require separate review.
- Published comparison: https://sverrejb.github.io/docxide-pdf/

Direct OOXML-to-PDF renderer with a Word-exported reference corpus. Useful
implementation references: `src/docx/alt_chunk.rs` (HTML/MHTML imports),
`sections.rs` (page borders, line numbering, vertical alignment), `numbering.rs`
(restart/override/legal-numbering semantics), and the preset-geometry engine.
Our Typst backend handles complex-script shaping; reproduce gaps before copying patterns.

Pinned corpus: 393 inputs/391 references; re-run current code against release-only comparisons.

```sh
python3 scripts/audit_docxide_corpus.py --corpus-root /path/to/pinned/docxide-pdf \
  --converter /path/to/current/office2pdf --output /path/to/local/audit
```

The runner includes skipped cases, records failures/timeouts, hashes, reference
metadata, page counts and text findings; it never executes fixture generators.
Token identities are hashed; source text stays outside reports.
Baseline (`tests/compatibility/docxide-pdf/baseline.json`, main `39bcb289`): 390
comparisons, two missing references, one image failure, 64 page-count differences, 27 empty text layers.
Custom-XML text is retained. Alternate-prefix body/header/footer runs use
URI-based parsing backported from [docx-rs#934](https://github.com/bokuweb/docx-rs/pull/934).

This is triage, not a visual pass: missing searchable text can still be painted
as an image/path, and extraction order or font availability can change scores.
HTML expansion follows its CSS mapping, using Rust [html5ever](https://github.com/servo/html5ever) and
[mailparse](https://github.com/staktrace/mailparse); docx-rs and contextual scans read the expanded XML.
Section line numbering maps starts, intervals and restarts to Typst margin labels.
Case70 native exports check fonts/clearance; compiled tests cover section suppression.
Section alignment maps `sections.rs` properties into Typst page flow. Fresh
Word case69 probes calibrate trailing spacing and continuous-section uniformity;
`both` and mixed automatic pagination warn; mixed explicit shared pages use top.
Numbering follows placeholder formatting and legal decimalization. Native case71
inline probes check ancestor counters and literals; case72 probes calibrate
restarts, suppression, implicit counters and skipped-level origins. MS-OI29500 2.1.282
confirms Word ignores restart rules inside complete level overrides.
Marker clearance/suffixes follow `pdf/list_label.rs`. Native probes calibrate
font/size/tab/wrap differences and inline marker ascent versus body descent/leading.
Damaged-PNG handling follows `src/pdf/images.rs`: preserve the drawing when
decoding fails. A direct `png` dependency reuses locked 0.18.1 to recover samples;
native Word probes check CRC and truncated-media behavior.
EMF translation follows `src/pdf/emf.rs` record semantics: retain Unicode text,
fonts, explicit UTF-16 advances, drawing state, lines, rectangles and clipping.
Glyph-index text and world transforms remain unsupported.
