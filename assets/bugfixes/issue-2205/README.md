# DOCX standard-ligature defect (#2205)

Public synthetic case8 from [docxide-pdf](https://github.com/sverrejb/docxide-pdf/tree/1bb140b662a69f8786015c1b532d37d6084d0934/tests/fixtures/cases/case8),
Apache-2.0. No internal document or font files are submitted.

`compare.jpg` is the defect evidence: the Word reference on the left, main
`64f1d00c16294ddeb1296bf70ab93c6d4566c0aa` output on the right. `gt.jpg`,
`before.jpg` and `after.jpg` are the fix evidence: full Letter page 1 of 1 from
`pdftoppm` at 150 DPI, progressive JPEG quality 86 with metadata stripped and
the density reset.

- Source SHA-256: `b7f679d3cecd44c82db4125b3b1f47d913b1fc5c113794ba54dbe47b17207219`
- Reference SHA-256: `fe4f5428ce5c14d906c1601c27d8586c1167b2b661007d5fba1a2f6371065dd6`;
  `Creator: Microsoft Word`; source-aware integrity check passes.
- Before: main `faeca2bc` release CLI; PDF SHA-256
  `05b65e4754e85536bd614018864f25d7a7229c174fb5caa734b85463c3678990`
  (the defect evidence's output).
- After: the fix's release CLI; PDF SHA-256
  `cb94c9aeb4612b51f9d889416c643ce02d4545098d0e658805ed2f996a870d21`.

No run in case8 states `w14:ligatures`, so Word draws none. In
`culpa qui officia deserunt mollit anim id`, Word draws `f`, `f` and `i` as three
pixel cells. Before, the face's `fi` ligature fused two of them, so the rest of
the line ended one 10pt cell early (width -2.44%, worst line width 2.4%). After,
the line draws every cell where Word does; the worst line width is 0.2%, and no
glyph is a ligature continuation. The text layers are identical.

Remaining: every line sits 0.50-0.54pt high, and the right-aligned Arial
paragraph's first line is 0.32pt right (#1874). Those offsets are the 14 fine
shifts in `layout-audit.json` and the 322 glyph-edge clusters in
`render-clusters-page-1.json`.
