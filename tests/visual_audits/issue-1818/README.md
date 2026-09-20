# Default worksheet padding under print scaling

Synthetic fit-to-page source at 85%, with a 78% variant and native re-zip control.
No explicit print percentage is needed: exact main `3022531c` already reproduces
this defect through the existing fit-to-page path.

Native first-column text starts at 51.85pt at 85%, and 52.26pt at 78%.
The respective grid origins are 49.30pt and 49.92pt, leaving insets 2.55pt and
2.34pt: the 3pt inset at 100% inset multiplied by the print scale. Exact main starts
both at 53pt. Its scaling helper adjusts explicit cell padding but omits the
table's default padding. That omission accounts for 0.45pt at 85% and 0.66pt
at 78%. A separate grid/text origin counter-shift accounts for the remaining
horizontal offset (0.70pt and0.08pt); do not attribute the entire shift to padding.

Full pages, pixel diff and shifted-region crops were inspected on the
pixel-identical explicit-percentage candidate comparison. This fit-to-page
comparison uses exact-main output; decoded GT/output/diff images all match
that inspected pair exactly. All 45 labels and 16 solid gray rules remain, with
right-shifted text and existing baseline/alignment deviations. Raw rectangle
counts differ because native emits fills and output uses three overlapping
strokes per rule; these counts alone do not demonstrate missing rules.
This evidence does not claim visual parity or a completed fix.

Related findings: default-cell descent #1814, alignment #1721, glyph advances
#1659, and explicit percentage parsing #1817. The independent origin
counter-shift still needs its own issue before a fix can claim a complete audit.
