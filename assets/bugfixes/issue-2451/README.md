# DOCX tab after a line break defect (#2451)

The reproducer is an authored synthetic DOCX and is not committed; its parts are
listed below. It is a Letter page with 1in margins and `w:defaultTabStop w:val="720"`.
docDefaults set Liberation Serif 12pt and single spacing with `w:after="0"`. Each
paragraph has `w:spacing w:after="240"` and one run with that font. The first
three paragraphs are the control lines:

```xml
<w:t xml:space="preserve">Name:</w:t><w:tab/><w:t xml:space="preserve">Jane Example, Customer Services</w:t>
<w:t xml:space="preserve">Telephone:</w:t><w:tab/><w:t xml:space="preserve">555 0100</w:t>
<w:tab/><w:t xml:space="preserve">Leading tab</w:t>
```

The fourth paragraph holds the same three lines in one run, joined by `<w:br/>`.

`compare.jpg` is the defect evidence. The left half is the reference, a
**LibreOffice 24.2 export (not native Word)**: `Producer: LibreOffice 24.2`, and
the source-aware integrity check passes. The right half is the main
`9ce9f78c50102a89fce6f36ae275cecbca9b8e8c` release CLI output. Both halves are the
full Letter page 1 of 1 from `pdftoppm` at 150 DPI, saved as progressive JPEG
quality 86 with metadata stripped and the density reset.

In the fourth paragraph, every tab after the `w:br` advances from where the
previous line ended. LibreOffice puts `555 0100` at 144.10pt and `Leading tab` at
108.10pt, the same stops as the control lines. office2pdf puts them at 142.69pt
and 88.71pt. `compare_layout.py --audit --fine-shift 0.5` matches all 6 lines
and reports exactly those two shifts, -1.41pt and -19.39pt; the figures include
the reference's constant +0.1pt offset. All nine 5% render clusters sit on those
two lines. The text layers are identical.
