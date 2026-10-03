# Issue #1920: DOCX list marker continuity

The tracked fixture `tests/fixtures/docx/libreoffice/tdf149711.docx` contains a
`w:bookmarkEnd` between the third and fourth visible numbered paragraphs after
Word accepts its move revisions. A fresh Microsoft Word for Mac 16.113.3 export
on macOS 26.6.2 (build 25G83), made with
`scripts/macos/export_word_pdfs.applescript`, renders four consecutive items.
`check_gt_integrity.py` reports one page and `invalid: false`.

At baseline `dfad2798d233b25b8e962cadd919cdb87be05e5f`, the last item is
13.15 pt below the Word baseline. With the parser fix, `compare_layout.py
--audit --fine-shift 0.5` matches all four lines and the worst baseline delta is
0.14 pt. There are no missing or extra lines, changed wraps, visibility
mismatches, visible-fill differences, or rectangle geometry findings.

The text-layer comparison reports a sequence mismatch because Word's PDF
extracts all four list labels before the item text, while office2pdf interleaves
each label and paragraph. The full-page and matched text-block crops show the
same visible label and paragraph on each line, in the order produced by the
source after revisions are accepted.

At 150 DPI, the strict render audit finds no material diff clusters. The
remaining 5% pixel differences are dispersed glyph-edge rasterization and
antialiasing; the full pages, diff, and matched crops were inspected.
