# Issue 2226 visual evidence

Tabs that open a Word body paragraph now advance to its stops, counted from
the margin past the first-line indent. A tab whose left or default stop leaves
no room for the next word before the line's end starts the next line with that
word, as in Word. A line left holding only tabs keeps their runs' height.

Source: docxide-pdf case
[`leading-tab-four-tabs`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/leading-tab-four-tabs)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (Apache-2.0, not committed). Its
`reference.pdf` reports `Creator: Microsoft Word` and one A4 page, and
`check_gt_integrity.py --source` finds no corruption.

`gt.jpg`, `before.jpg` and `after.jpg` are page 1 of 1 from `pdftoppm` at
150 DPI, saved as progressive JPEG quality 86 without metadata and with the
density reset.

SHA-256:

- Source: `2c1ff0b83fbe73ff5bb7217cc34b3adf60df3e685669ad510b2023c9ceff86a9`
- Word reference: `33c504b738860037ec7bf49900c961e8655253685fa1e245a89005daa2f1387d`
- Before output (main `aed6c59a`): `fb6484a059e36429c99973d7817f79019b144614143952fe74a34eed7090d392`
- After output: `88d9ea5ab74df33c0ec2041f0d0b20508ec94cd8f88589276b9dda2a862143ae`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and
`render-clusters-page-1.json` from `compare_render.py --dpi 150
--fine-shift 0.5 --strict-clusters`.

`Public greeting text` moves from (72.00, 72.59) to (268.08, 104.79), on
Word's third line at (268.38, 104.80); `Public body` follows it to within
0.07pt of Word. The text layer's normalized content matches Word's.

The other `leading-tab-*` cases, `mutool draw -F stext` line starts (x, top)
of `Public greeting text`:

| Case | Word | Before | After |
| --- | --- | --- | --- |
| `leading-tab-one-tab` | 268.38, 72.55 | 295.03, 72.59 | 268.08, 72.59 |
| `leading-tab-two-tabs` | 268.38, 88.55 | 72.00, 72.59 | 268.08, 88.69 |
| `leading-tab-three-tabs` | 268.38, 88.55 | 72.00, 72.59 | 268.08, 88.69 |
| `leading-tab-small-first` | 268.38, 84.05 | 72.00, 72.59 | 268.08, 84.09 |
| `leading-tab-large-first` | 268.38, 95.55 | 72.00, 72.59 | 268.08, 95.59 |
| `leading-tab-right-stop` | 393.67, 72.55 | 420.56, 72.59 (wraps `text`) | 393.61, 72.59 |
| `leading-tab-visible-prefix` | 268.38, 88.55 | 72.00, 88.69 | 268.08, 88.69 |

In `leading-tab-visible-prefix`, `Prefix` moves from 72.00 to 98.95; Word has
99.05.
