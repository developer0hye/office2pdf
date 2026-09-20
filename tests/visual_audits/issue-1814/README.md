# Default cell-style descent probe

Synthetic one-page matrix: five cell fonts, three row heights (15, 19.5 and 30pt), and bottom/top/center alignment. The first cell-style record is unused by the 45 cells; its duplicate styles A1.

`probe.json` changes only `cellXfs[0].fontId`. Native Excel keeps cell fonts, top/center baselines and all row boundaries unchanged, but changes bottom baselines. The default Malgun Gothic 11 style gives a 4pt minimum; Arial 11 allows 2pt, while Calibri 11 and Trebuchet MS 14 allow 3pt. These measurements establish the style dependency, not a complete size/family formula.

The unmodified converter places the roomy Arial row 6 baseline at 208pt instead of 206pt. Roomy rows isolate this floor error from the tight-row alignment override in #1721. `native-baselines.json` records every native label and row boundary; `before-baselines.json` compares the base with the exact-main output. Reference hashes and export metadata are in `provenance.json`.

The comparison also contains separate tight-row, top-seat and glyph-origin differences; this evidence does not claim visual parity or a completed fix. No user-facing API or usage changed.
