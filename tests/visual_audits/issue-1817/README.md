# Explicit worksheet print-scale evidence

Synthetic source for proposed issue 1817. `source.xlsx` declares
`pageSetup scale="85"`; `base.xlsx` omits scale (native 100%). The one-factor
probe reproduces native 85% and 78% outputs and passes its unchanged re-zip
control. No fit-to-page flag is set.

Native row 6 adjacent label origins are 91.8pt apart at 85%, 84.24pt at 78%,
and 108pt at 100%. Exact main at `bd443885` retains 108pt for the 85% source.
Both the table and text remain at 100%, producing a larger table and growing
position drift. All 45 synthetic labels and 16 horizontal rules remain present.
The extracted non-whitespace character multisets agree; extraction order differs.

Full pages, pixel diff and all 32 emitted shifted-region crops were inspected
at 150 DPI. Black text and solid gray rules remain; no new fill, rotation or
missing element was seen. Existing alignment/descent/glyph findings remain
tracked in #1721, #1814, #1815 and #1659. This is defect evidence, not a visual pass.
