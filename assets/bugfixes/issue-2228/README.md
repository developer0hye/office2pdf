# Issue 2228 visual evidence

A line under `w:lineRule="exact"` now keeps its stated height when it holds an
inline picture taller than the line. The picture still stands on the baseline,
and only the part below the line's top is painted. Before, the picture raised
the line to its own height and pushed everything after it down.

Source: docxide-pdf case
[`cell-inline-exact`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/cell-inline-exact)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (Apache-2.0, not committed). The
paragraph sits in a borderless single-cell table with `w:spacing w:line="360"
w:lineRule="exact"` (18pt) and holds `Left `, a 72×36pt inline picture and
` End` in Times New Roman 14pt. Its `reference.pdf` reports `Creator: Microsoft
Word` and one A4 page, and `check_gt_integrity.py --source` finds no
corruption.

`gt.jpg`, `before.jpg` and `after.jpg` are page 1 of 1 from `pdftoppm` at
150 DPI, saved as progressive JPEG quality 86 without metadata and with the
density reset.

SHA-256:

- Source: `64c3d63c2a32ac695a410f06c694a2e258ae761a5c66b85c5a9d5b4891b051c7`
- Word reference: `6c3cb3ddf695fce16167f4fad22391fac2bbdee378786558f74d275dca3cc1f1`
- Before output (main `95e1c780`): `d4ce724af3e4cd2c8a564d74d359ed06eedd3ff1f763974eb1ce995955a4a3fc`
- After output: `674c6c14552bb3a59faa74adb4b3c0cd3e23e8bfb2a7d5936a40a8c062caeb77`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and
`render-clusters-page-1.json` from `compare_render.py --dpi 150
--fine-shift 0.5 --strict-clusters`. The text layer's normalized content
matches Word's.

Positions from `mutool trace` (picture rect and clip, device space) and
`mutool draw -F stext` (baselines):

| Item | Word | Before | After |
| --- | --- | --- | --- |
| Picture painted | y 50.40–86.40 | y 72.00–108.00 | y 49.07–85.07 |
| Picture visible (clip) | y 72.00–86.40 | whole picture | y 72.00–85.07 |
| `Left`/`End` baseline | 86.53 | 108.00 | 85.07 |
| `Following paragraph` baseline | 103.03 | 126.00 | 103.07 |

The following paragraph is now 0.04pt from Word's, where it was 22.97pt low.
The line's top and the clip's top edge match Word.

The remaining 1.46pt is the exact line's baseline seat. In a table cell,
office2pdf puts the baseline of an expanded exact line at the face's ascent
(13.07pt below the line top here). Word puts it 0.8 of the line down (14.4pt).
So `Left`/`End` sit 1.46pt high, and the visible strip of the picture ends
1.33pt higher than Word's (13.07pt shown against 14.4pt). The seat rule is the
same for text-only exact lines and is recorded as sub-material under #1874.
