# Issue 2227 visual evidence

An inline picture that shares its paragraph with text now rides that text's
line, in run order. Its bottom edge sits on the baseline the text shares, a
picture taller than the text grows the line upward, and the next line still
follows one ordinary line pitch below. Pictures with no text beside them keep
the picture-block path (#499, #2201).

Source: docxide-pdf case
[`cell-inline-body`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/cell-inline-body)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (Apache-2.0, not committed). Its
`reference.pdf` reports `Creator: Microsoft Word` and one A4 page, and
`check_gt_integrity.py --source` finds no corruption.

`gt.jpg`, `before.jpg` and `after.jpg` are page 1 of 1 from `pdftoppm` at
150 DPI, saved as progressive JPEG quality 86 without metadata and with the
density reset.

SHA-256:

- Source: `dc2b4699c030a0a30774c82e18124c4eb8109b33e8aa05cd7cc96b2823e4b4df`
- Word reference: `26a8cc7e4a54ca7392328f18983df3350de76c2b5adff6e4d84df8c2a9e59312`
- Before output (main `8597ee78`): `12b7112bd425e57d67ba93dd81e49095509aa0c1bf03a781c91c8fb6e5996ea5`
- After output: `3984816b8a4ae3b4c90315e2a7b1d3136d915913d5c32f42e0e04e9b4e298454`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and
`render-clusters-page-1.json` from `compare_render.py --dpi 150
--fine-shift 0.5 --strict-clusters`. The layout audit reports no findings, and
the text layer's normalized content matches Word's.

All twelve reported cases, from `mutool trace` picture origins (top-left x, y)
and `mutool draw -F stext` glyph positions (`End` x / baseline; `Following`
baseline):

| Case | Word | Before | After | Remaining |
| --- | --- | --- | --- | --- |
| `cell-inline-body` | (98.82, 72); End 174.33/108.03; Foll 124.03 | (72.10, 72); End 98.92/121.07; Foll 137.17 | (98.92, 72); End 174.42/108.00; Foll 124.10 | — |
| `cell-inline-nested` | as `body` | (72.20, 72); Foll 148.67 | (99.02, 72); End 174.52/108.00; Foll 135.60 | #2238 |
| `cell-inline-tabbed` | (108.00, 72); End 183.58; Foll 124.03 | (72.10, 72); End 111.60/121.07; Foll 137.17 | (111.60, 72); End 187.10/108.00; Foll 124.10 | #2344 |
| `cell-inline-header` | (98.82, 28.35); End 174.33/64.28; body 85.03 | (72.10, 28.35); End 98.92/77.42 over the body line; body 85.07 | (98.92, 28.35); End 174.42/64.35; body 85.07 | — |
| `cell-inline-exact` | (98.82, 50.40); End 86.53; Foll 103.03 | (72.10, 72); End 121.07; Foll 139.07 | (98.92, 72); End 174.42/108.00; Foll 126.00 | #2228 |
| `leading-picture-image-fits` | (82.50, 72); End 72.03/121.03 | (72, 72); End 86.00/121.07 | (82.50, 72); End 72.00/124.10 | #2343 |
| `leading-picture-image-no-spaces` | (72, 72); End 72.03/121.03 | (72, 72); End 72.00/121.07 | (72, 72); End 72.00/121.07 | — |
| `leading-picture-image-overflow` | (72, 88.10); End 72.03/137.05 | (72, 72); End 117.50/121.07 | (72, 72); End 72.00/121.07 | #2343 |
| `picture-padding-mixed` | (72, 90.11); Foll 162.30 | (72.10, 80.62); Foll 159.88 | (72.10, 80.62); Foll 151.25 | #2343 |
| `picture-padding-no-placeholder` | as `mixed` | as `mixed` | as `mixed` | #2343 |
| `picture-padding-small-picture-font` | as `mixed` | as `mixed` | as `mixed` | #2343 |
| `picture-padding-no-padding` (control) | (72, 81.05); Foll 153.30 | (72.10, 80.62); Foll 151.25 | unchanged | — |

Text before a picture now stays before it and on its line. Two `End` and
`Following` positions move away from Word: `leading-picture-image-fits` (End
121.07 → 124.10 against Word's 121.03) and the three `picture-padding-*` cases
(Following 159.88 → 151.25 against 162.30). Main matched more closely only
because the lifted picture printed the spaces as an extra line below it. With
the spaces on the picture's line, #2343's whitespace line metrics decide those
lines: Typst gives a spaces-only line no height and lets leading spaces add
descent, and Word does neither.
