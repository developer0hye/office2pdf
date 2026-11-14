# Fit-to-page 10% floor evidence

Source for issue #2474. `source.xlsx` is the `fit1x1.xlsx` workbook from
#2472, written by openpyxl 3.1.5 from that issue's own script: 400 rows of
`row N colA`..`row N colC` in three width-20 columns, `fitToPage="1"` with
`fitToWidth="1" fitToHeight="1"`. SHA-256
`90afa6ea0a9a8cc05b1a3cb56ccd9f439415d807c108fc0a471f1bce2cdc8d07`.

Ground truth is a Microsoft Excel for Mac 16.115 export (macOS 27.2, build
26B123) made with `scripts/macos/export_excel_pdfs.applescript`, PDF SHA-256
`eeb367a4cb722a6e3f9a0b7af3e179c801162dc7fdfdf3db636134c44aa3c3e6`.

The sheet is 400 rows of 17pt, 6800pt, against 648pt of printable height, an
exact fit of 9.53%. Excel prints it at 10% (a 1.70pt row pitch), which leaves
381 rows on page 1 and 19 on page 2. Before the fix the converter fitted it at
9% and printed all 400 rows on one page.
