# Style-only trailing rows evidence

Synthetic source for issue #2475, written by openpyxl 3.1.5. `source.xlsx`
holds `ID`/`Description` and 110 `ITEM N` rows in A1:B111, then rows 112-1061
carry only an Arial 10 font and the `@` text format on A:E: no value, border
or fill. The dimension is `A1:E1061`. Landscape US Letter, no gridlines, no
scaling. SHA-256
`0ed46c6819e709c69426d9148d97622b6053aec971614ecfd56c626ce267a555`.

Ground truth is a Microsoft Excel for Mac 16.115 export (macOS 27.2, build
26B123) made with `scripts/macos/export_excel_pdfs.applescript`, PDF SHA-256
`38a75f2459a4464379ec19bd89b75df2a83cf5407421edd226a39be95ec8e3fc`. It prints
five pages: rows 1-27, 28-54, 55-81, 82-108 and 109-111. Before the fix the
converter printed those five pages and then 35 blank ones for rows 112-1061.

`cluster-dispositions-page-N.json` hold the strict diff-cluster dispositions
for the render reports in `assets/bugfixes/issue-2475/`. Every cluster is a
sub-pixel glyph advance drift inside a label, tracked under #1874.

The same rule was measured on one-factor edits of the workbook attached to
#2472 (`A1:Y1061` declared, ink through row 111). Excel ends every export at
row 111:

| Variant | Excel pages |
| --- | ---: |
| As attached: fit to one page wide, gridlines | 4 |
| No scaling, gridlines | 8 |
| No scaling, gridlines and row/column headings | 8 |
| No scaling, no gridlines | 8 |

Two more probes, each a single-sheet openpyxl workbook with values in A1:A5:

- A manual row break at row 100, with bold-only cells to row 300: Excel prints
  1 page.
- A 400-row `=IF(A1="","",A1)` column (the formula is content even when its
  cached result is empty): Excel prints all 400 rows on 11 pages.
