# Header tabs resolve against their stops

`source.docx` is a synthetic A4 package in Arial 10.5pt. Every header and footer paragraph states Word's Header-style stops directly: centre at 4153 twips and right at 8306. Header paragraph 1 is `<tab>{Title}` with a bottom border, header paragraph 2 is `Confidential<tab>Draft 2`, and the footer is `Left<tab>Centre<tab>Right`.

Word advances each tab to the next stop past the pen. It centres `{Title}` at 282.87pt and `Draft 2` at 281.89pt on the 297.65pt centre stop. Before the fix, a paragraph with one tab and a right stop last was laid out as a left/right grid, so both went to the right margin (475.73pt and 473.79pt). Header and footer tabs now go through the body's measured tab engine, and both land at 282.87pt and 281.89pt. The two-tab footer was already correct and is unchanged.

`compare_layout.py --audit --fine-shift 0.5` has no large shift after the fix (three before). All nine render clusters and the remaining fine shift are vertical and belong to #1824. The bordered first header line and its rule sit 2.49pt high, and the second header line sits 20.31pt low, overprinting the first body line. Without the border, the same header seats both lines within #1640's line gap, so that displacement is the bordered paragraph's own. The normalized text layer holds the same 175 characters; Word writes the footer before the body.

Full 150-DPI pages, the header-band crops at 300 DPI and the difference image were inspected. Body text differs only by glyph-edge antialiasing.
