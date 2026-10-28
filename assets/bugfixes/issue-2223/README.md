# Issue 2223 visual evidence

A tab in a Word list item's body now advances to the paragraph's next stop,
counted from the margin like any body tab. Page 3's `3.2 PREPARATION:` starts
with a tab after a `3.%1` number whose 21.6pt hanging indent already reaches
the paragraph's own 21.6pt stop. The text now lands on the 36pt default stop,
as in Word, where main left it at the hanging indent.

Source: docxide-pdf scraped
[`construction_quality_control_spec`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/scraped/construction_quality_control_spec)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (a public document, not
committed). Its `reference.pdf` reports `Creator: Microsoft Word` and three
Letter pages, and `check_gt_integrity.py --source` finds no corruption.

`compare.jpg` is the defect evidence: the Word reference on the left and
office2pdf at filing time on the right. `gt.jpg`, `before.jpg` and `after.jpg`
are the fix evidence: page 3 of 3 from `pdftoppm` at 150 DPI, saved as
progressive JPEG quality 86 without metadata and with the density reset.

SHA-256:

- Source: `b2c0d12a5dd997088ba597e35f4b5665876cffcc2088db3cc448c2bd06153e06`
- Word reference: `5f75e91c24dccfcb63e7d6a987a48ac975891d92cfb3230456839b0b0b6df610`
- Before output (main `150584f0`): `cffec4ed3d3b3da1751d3bd604bef5ccc45d2c1f7d0a8d1a3fd9251e379e3bf7`
- After output: `f3c91f91551fa30a071658e7b8d41f819e83ae18411a96a60c9f9efce89ed824`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5` over all three pages, and
each `render-clusters-page-<N>.json` from `compare_render.py --dpi 150
--fine-shift 0.5 --strict-clusters`.

`PREPARATION:` moves from x 86.40pt to 100.80pt, Word's position; its start is
within 0.02pt of Word's on both axes. The other 134 lines are unchanged from
main, and the text layer's normalized content matches Word's.

Remaining, identical on main:

- The footer's first line is 3.06pt above Word's (#2326).
- `DMS 2023 Edition` is 0.81pt left, and body words drift up to 1.79pt along
  long tracked lines (#1874).
