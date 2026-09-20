# Small worksheet margin origin

Synthetic unscaled Letter worksheet with three alignment labels in Malgun Gothic 11pt. The explicit left margin is 0.25in. Native Excel places all three first glyphs at x22pt; current main places them at x21pt. The filing image shows native left and converter right at 150 DPI.

The controlled margin sweep holds native x at 22pt for 0–0.25in, then native and converter both place text at x24/31/53 for 0.3/0.4/0.7in. This rules out a universal extra point of cell padding. The exact minimum-origin mechanism is unproven; printer configuration and other paper sizes/scales need investigation before choosing an implementation. The explicit-zero fallback in #1812 is separate.

Both re-zip controls passed and every PDF contains exactly the expected three labels on one page. The current main PDF is byte-identical to the earlier controlled converter baseline. Vertical font/alignment differences are outside this filing's scope. This is initial issue evidence, not a completed fix audit. README usage and public API are unchanged.
