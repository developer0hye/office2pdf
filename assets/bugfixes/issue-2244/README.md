# Issue 2244 visual evidence

A tab leader (`w:leader` `dot`, `hyphen` or `underscore`) is drawn in the
formatting of the run that holds the `w:tab`: its face, size and colour. It
used to take Typst's ambient face and size (Libertinus Serif 11pt).

Source: `tests/fixtures/docx/docxide/tab_leaders.docx`, the unmodified
docxide-pdf
[`case10`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/case10)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (synthetic, Apache-2.0). Its
`reference.pdf` reports `Creator: Microsoft Word`, has three Letter pages, and
embeds Aptos and Calibri Bold. `check_gt_integrity.py --source` passes.

SHA-256:

- Source: `2684464295123d0e9cde655926b26dc91d1ded4636a8a04e427f38761b3b4f26`
- Word reference: `9df1100edf1fe1262a0f147109ebb46b7b2f35715474981352401af5b63df7fb`
- Before output (main `518b2ed2`): `94112d73a2b4c2c34cb291d951b391f4a5946c800f8454a4c5e041cc0c7c223f`
- After output: `70d27e28d807c554cddc382b60a353bf2ae31e288d493e48b52aef19df743391`

Both CLIs were built with `cargo build --locked --release -p office2pdf-cli`
in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5` over all three pages.
Each `render-clusters-page-N.json` comes from `compare_render.py --dpi 150
--fine-shift 0.5 --strict-clusters`. The JPEGs are page 1 rendered by pdftoppm
at 150 DPI and saved as progressive quality 86 without metadata. Pages 2 and 3
are pixel-identical before and after.

Leader rows under "Tab Leaders" (page 1), from `mutool draw -F stext`
(dot count, first–last dot x in pt):

| row | Word | before | after |
|---|---|---|---|
| Introduction | Aptos 12, 93, 154.08–476.15 | Libertinus 11, 131, 160.69–476.72 | Noto Sans 12, 99, 160.69–475.92 |
| Background | Aptos 12, 93, 154.08–476.15 | Libertinus 11, 132, 158.51–476.72 | Noto Sans 12, 99, 158.51–475.92 |
| Methods | Aptos 12, 96, 136.55–469.15 | Libertinus 11, 137, 139.79–469.85 | Noto Sans 12, 103, 139.79–469.06 |
| Results | Aptos 12, 98, 129.55–469.20 | Libertinus 11, 141, 130.57–469.85 | Noto Sans 12, 106, 130.57–469.06 |
| Conclusion | Aptos 12, 92, 150.58–469.15 | Libertinus 11, 132, 152.05–469.85 | Noto Sans 12, 99, 152.05–469.06 |

Noto Sans is this host's stand-in for Aptos. Converted with the Aptos subset
Word embedded in the reference (`--font-path`), the after build draws the dots
in Aptos 12 and ends each row at 476.16pt or 469.75pt. Every other line on the
three pages then lands within 0.4pt of Word. Without the subset, those lines
sit up to 25pt low (#2237).

Remaining, each with its own issue:

- The stand-in's narrower period (3.216pt against Aptos's 3.428pt) and its
  taller line pitch move and widen every line (#2237).
- Word sets the dots on one 3.501pt grid that all rows share. office2pdf
  repeats them at the glyph advance and justifies each row, so even in Aptos
  a row holds one or two more dots and the dot columns drift between rows
  (#2397).
- Page 2's script digits are the stand-in's designed superior and inferior
  figures, where Word scales the regular digits to 7pt (#2398).
- In Aptos, page 3's last paragraph breaks after `bold/`, where Word keeps
  `bold/italic` together (#2297).

Corpus scan: release CLIs from main `518b2ed2` and this branch, run on 213
tracked DOCX fixtures, 393 docxide-pdf cases and 159 scraped documents. The 26
inputs that fail on main fail the same way. 12 outputs change, and in every
one the leader marks take their tab run's face and size. In 6 of them,
leader rows also change height: their leader box no longer carries the 11pt
Libertinus line. The clearest case is the defined-terms index of
`western_australia_owner_drivers_act`. Word paces its lines 11.50pt apart.
There, office2pdf moves each leader onto a line of its own, which is a separate
wrap. Main set that line 12.43pt below the one above and the next line 11.72pt
below it; this branch paces every line at 11.50pt. The
scan also found #2399: a `TOC` field with a cached result printed twice. Its
computed outline still draws Libertinus dots, because it has no `w:tab` run.
