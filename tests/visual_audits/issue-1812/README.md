# Explicit zero worksheet margin

Synthetic one-page workbook with three labelled alignment samples in Malgun Gothic 11pt and an explicit `pageMargins/@left="0"`. The filing comparison is native Excel (left) versus the candidate recorded in `filing-evidence.json` (right), both at 150 DPI.

The converter places the first glyph at 53pt because `sheet_print_margins` treats zero as absent and selects the default left margin. Native Excel places it at 22pt. A controlled left-margin sweep showed native x22 for 0–0.25in and x24/31/53 for 0.3/0.4/0.7in; the converter matches the latter three. Native minimum printable-origin behavior is separate from the parser's explicit-zero/default conflation. Fixing zero presence must not imply hardcoding the native minimum origin.

The native and converter re-zip controls passed. Every export contained exactly the three expected labels on one page. This is filing evidence, not a completed visual-fix audit; the one-point font/alignment differences remain outside this issue's scope. README usage and public API are unchanged.
