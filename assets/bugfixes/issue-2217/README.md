# Issue 2217 visual evidence

A table row with neither `w:cantSplit` nor a height floor starts in the space
left on the page and continues on the next page, as in Word.

Source: `tests/fixtures/docx/docxide/first_header_row_overflow.docx`, the
unmodified docxide-pdf
[`first-header-no-minimum-overflow`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/first-header-no-minimum-overflow)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (synthetic, Apache-2.0). Its
`reference.pdf` reports `Creator: Microsoft Word`, three 500 × 500pt pages,
and embeds `Aptos` and `ArialMT`; `check_gt_integrity.py --source` passes.

SHA-256:

- Source: `8f6e84eae20261a6959255bf1133242cf405b97a19226d5a65da706a350ee29c`
- Word reference: `b3a7af30c5c86821400b9838266afe39929511bc93a346d1df01e93987dcbf93`
- Before output (main `70ebaa0a`): `3be6b8cde15e366aaa1330217811445e72d51921078274b53a6b470d68768c61`
- After output: `0355772cdc54774adcdc84c9a85f182c973a13a3bfa550e551f4426cd772f9ca`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and each
`render-clusters-page-N.json` from `compare_render.py --dpi 150 --fine-shift
0.5 --strict-clusters`. The JPEGs are page 1 rendered by pdftoppm at 150 DPI
and saved as progressive quality 86 without metadata.

`Continuation` lines per page, from `mutool draw -F stext` baselines (the
source has 60):

| page | Word | before | after |
|---|---|---|---|
| 1 | `Header`, `Before`, `Minimum`, `Second`, 15 | `Header`, `Before` | `Header`, `Before`, `Minimum`, `Second`, 15 |
| 2 | 26 | `Header`, `Minimum`, `Second`, 17, then 43 overprinted on one baseline at 475.3pt, below the bottom margin | `Header`, 18 |
| 3 | 19 | — | `Header`, 18 |
| 4 | — | — | `Header`, 9 |

Page 1 matches Word: all 19 lines match, with baselines within 0.26pt.
Remaining, each with its own issue:

- From page 2 on, office2pdf repeats the first-page header (#2163). The row
  continues at 168.5pt instead of Word's 36.4pt top margin, so a page holds 18
  of Word's 26 lines and the row runs onto a fourth page.
- `Header` is painted in the Noto Sans stand-in for unembedded Aptos (#2237).
- A row with a `w:trHeight` floor still moves whole to the next page
  (`first-header-minimum-overflow`); it stays open under #2217.
