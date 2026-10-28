# Issue 2221 visual evidence

A `w:lvlText` whose literal text holds a Typst counting symbol now prints that
text as written: `1.%1` gives `1.1` … `1.9` where main printed `1` … `9`.

Source: docxide-pdf scraped
[`construction_quality_control_spec`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/scraped/construction_quality_control_spec)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (a public document, not
committed). Its `reference.pdf` reports `Creator: Microsoft Word`, three Letter
pages, and `check_gt_integrity.py --source` finds no corruption.

`compare.jpg` is the defect evidence: the Word reference on the left and
office2pdf at filing time on the right. `gt.jpg`, `before.jpg` and `after.jpg`
are the fix evidence: page 1 of 3 from `pdftoppm` at 150 DPI, saved as
progressive JPEG quality 86 without metadata and with the density reset.

SHA-256:

- Source: `b2c0d12a5dd997088ba597e35f4b5665876cffcc2088db3cc448c2bd06153e06`
- Word reference: `5f75e91c24dccfcb63e7d6a987a48ac975891d92cfb3230456839b0b0b6df610`
- Before output (main `0772fa70`): `446e3d575a25a4661c755c7390149668e4b394b1b8b9508e033dda92b4ec75f7`
- After output: `8dbeb1c0ea7cfb44e07731aa4e7c8e6ab091f9bf196148f051d89dcc1f27dcd6`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and each
`render-clusters-page-<N>.json` from `compare_render.py --dpi 150 --fine-shift
0.5 --strict-clusters`.

All 9 numbered headings on the 3 pages now match Word: the labels sit within
0.03pt horizontally and 0.23pt vertically, and the text layer has the same
non-whitespace characters as Word (main lacked nine `1` and nine `.`).

Remaining, identical on main:

- The footer's first line is 33.28pt high and overprints the last body line
  (#2222).
- On page 3, `PREPARATION:` starts 14.40pt left of Word (#2223).
- `DMS 2023 Edition` is 0.81pt left, and body words drift up to 1.76pt along
  long tracked lines (#1874).
