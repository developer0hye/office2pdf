# Issue 2247 visual evidence

A missing Tahoma is still painted in Liberation Sans (#2235), but its lines
are now laid out on Tahoma's own `hhea`: 2049/-423/0 at 2048 units per em, a
1.2070em line with its baseline 1.0005em down. Before, they used Liberation
Sans's 1854/-434/67: a 1.1499em line with its baseline 0.9380em down. Both
Tahoma weights declare the same values in the subsets Microsoft Word exports
embed (this case's reference, and docxide-pdf's
`rehab_centre_physio_posting`, which embeds Tahoma Bold too).

Source: `tests/fixtures/docx/docxide/page_header_ordinary_marker.docx`, the
unmodified docxide-pdf
[`page-header-ordinary-marker`](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/page-header-ordinary-marker)
at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed` (synthetic, Apache-2.0). Its
`reference.pdf` reports `Creator: Microsoft Word`, one A4 page, and embeds
only `Tahoma`. `check_gt_integrity.py --source` finds no corruption.

SHA-256:

- Source: `b40d3b1501b5057e9fcc7b54dc4496f3e3f0f8c4351667f413eaf043024bfe65`
- Word reference: `b7c98fc786e4da606a5f963c183583c6e8d70ede9a800ba7bcd40e9f0cc7b4ae`
- Before output (main `241fdedc`): `c229a2d4530b20a8977d8f7975642fb727fe188e7315fb98e4c6f888efa848f1`
- After output: `c18b9a675ba5ccc255534f4d5ea7315c1f5027fb90fdc45a520f15b269baff07`

Both CLIs were built with `cargo build --locked --profile ci -p
office2pdf-cli` in separate target directories. `layout-audit.json` comes from
`compare_layout.py --json --audit --fine-shift 0.5`, and
`render-clusters-page-1.json` from `compare_render.py --dpi 150 --fine-shift
0.5 --strict-clusters`. The JPEGs are page 1 rendered by pdftoppm at 150 DPI
and saved as progressive quality 86 without metadata.

From `mutool draw -F stext` (baseline below the page top). The oracle is
main converting the same document with Word's own Tahoma subset on
`--font-path`:

| line | Word | before | after | oracle |
|---|---|---|---|---|
| `HEADER 1` | 38.25pt | 37.73pt | 38.35pt | 38.35pt |
| `Following section` | 94.03pt | 92.88pt | 94.08pt | 94.08pt |

Controls, measured the same way:

- docxide-pdf `picture-padding-no-padding`: the picture's top moves from
  80.62pt to 81.05pt (Word: 81.05pt), and `Following` from 151.25pt to
  153.16pt (Word: 153.30pt).
- The fixture with its body paragraph replaced by 40 single lines: line 40
  moves from 541.34pt to 564.82pt, exactly where the oracle puts it.
- docxide-pdf `empty-margin-frame-15-two-empty` and
  `empty-margin-frame-long-prefix` (tracked as `empty_frame_box.docx` and
  `empty_frame_box_carried_rows.docx`), whose empty frame paragraphs are set
  in Tahoma. The first table line below the frame (page 1 of the first, page
  2 of the second) moves from 1.08pt to 0.06pt off Word's baseline. The
  greeting moves from 1.02pt to 0.12pt off in the first and from 1.33pt to
  0.18pt off in the second.

Remaining, with its own issue: Liberation Sans's glyphs and advances. Its
capitals set `HEADER 1` 10.1% wide, and `Following section` runs 1.6% wide
(#2405). All five render clusters are those glyph differences, on matching
baselines.
