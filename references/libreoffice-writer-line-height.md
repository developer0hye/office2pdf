# LibreOffice Writer line height

- **URL:** https://github.com/LibreOffice/core/blob/022c34f162f94d59499ec48e48edbfd075f583b0/sw/source/core/text/itrform2.cxx#L2445-L2469
- **Inspected revision:** `022c34f162f94d59499ec48e48edbfd075f583b0` (2026-10-19)
- **Language:** C++
- **License:** MPL-2.0

## What it does

`SwTextFormatter::CalcRealHeight` sizes each formatted line from the
paragraph's line-spacing rule:

- `SvxLineSpaceRule::Fix` (Word's `w:lineRule="exact"`) sets the line to the
  stated height and its ascent to `(4 * nLineHeight) / 5`, whatever the
  font's own ascent. A font taller than either part marks the line for
  clipping.
- `SvxLineSpaceRule::Min` (`atLeast`) only raises a shorter line to the
  stated height and leaves the line's ascent unchanged.

## Why relevant

- Word's own exports agree with the 4/5 split for exact lines: docxide-pdf
  references seat Aptos 12pt 128.08pt and 9.50pt below the tops of 160pt and
  12pt exact header lines, Times New Roman 14pt 12.68pt into a 16pt one, and
  Arial Narrow Bold 14.63pt into an 18pt one (issue #2214).
- Its at-least rule is a different branch, which is why office2pdf keeps
  `LineSpacing::AtLeast` apart from `LineSpacing::Exact`.
- office2pdf applies the split to header lines in
  `crates/office2pdf/src/render/typst_gen_text.rs`
  (`word_exact_line_box_pt`). Body lines still centre an exact line's extra
  height in `word_line_box_em`.
