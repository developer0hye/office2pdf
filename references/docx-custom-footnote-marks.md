# DOCX custom footnote marks

- [Microsoft semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footnoteendnotereferencetype.custommarkfollows):
  a true `w:customMarkFollows` suppresses the automatic reference mark and does
  not advance the automatic numbering sequence. Missing/false values use it.
- The independently authored fixture is
  `tests/fixtures/docx/docxide/custom_footnote_marks.docx`.
- Its body alternates custom `**`, automatic, custom `†`, automatic references;
  native Microsoft Word exports automatic markers `1, 2` in body and entries.
- Native one-factor XML probes change only the flag or following text topology.
  Re-zip controls, `1` versus `true`, and same-run/next-run/merged-text variants
  have zero decoded pixel differences at 150 DPI.
- [Before comparison](docx-custom-footnote-marks/compare.jpg): Word on the left,
  office2pdf at `456e9338463345f405560efa831e495d1d9b29b1` on the right.
  Extra body labels `1**` and `3†` also move the automatic markers to `2, 4`.
- The repaired reader retains the flag; the parser keeps the following text
  child's own style. Typst's content numbering function keeps the internal link,
  and a counter update removes the custom note's automatic counter step.
- Fresh compiled tests verify numbering, searchable text, text-node/run
  boundaries and independent baseline/superscript/subscript styles.
- Current authored-fixture PDF links resolve to note destinations; automatic
  entries retain their backlinks. This is output-link validation, not proof of
  native Word's link rectangles for every source topology.
- Note indentation and paragraph/area layout remain under #2138; the corpus
  still requires full visual audits. Text equality is not layout parity.
