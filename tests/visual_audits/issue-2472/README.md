# Scaled worksheet pagination evidence

Source for issue #2472. `source.xlsx` is the issue's `scale50.xlsx`: 400 rows
of `row N colA`..`row N colC` in three width-20 columns under
`<pageSetup scale="50"/>`, written by openpyxl 3.1.5 from the issue's own
script. SHA-256 `f29e3f10f49a4bfb5b3c271ebd653826776bd4c8515a3e1cd83f61fbda78fb0a`.

Ground truth is a Microsoft Excel for Mac 16.115 export (macOS 27.2, build
26B123) made with `scripts/macos/export_excel_pdfs.applescript`, PDF SHA-256
`471299709c752469de7ecd029a2bca26ca79fdbc19be925b78f27a1144154ea3`. It prints
six US Letter pages of 76, 76, 76, 76, 76 and 20 rows. Before the fix the
converter printed one page: rows 1-76 inside the printable area, rows 77-84 in
the bottom margin, row 85 cut at the paper edge, and rows 86-400 laid out below
the paper and lost.

The evidence in `assets/bugfixes/issue-2472/` compares page 1 as images and
every page in the layout and render-cluster reports.

The issue's other workbooks, exported the same way:

| Workbook | Native pages (rows per page) | Before | After |
| --- | --- | --- | --- |
| `fit1x2` (`fitToWidth=1 fitToHeight=2`) | 2 (200, 200) | 1 (223 rows) | 2 (200, 200) |
| `fit1x1` (`fitToWidth=1 fitToHeight=1`) | 2 (381, 19) | 1 (400) | 1 (400) |

`fit1x1` is a separate rule, tracked in #2474: Excel does not fit below its
10% minimum, so the sheet still needs a second page. The converter fits it at
9%.
