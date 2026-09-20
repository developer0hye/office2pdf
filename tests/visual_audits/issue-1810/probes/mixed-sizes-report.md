# Probe: header-multiline-mixed-sizes

- base: two-lines.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| first and second header font size | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 8-11 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 0.75 | -3.00 | 3.00 | 29.7 | 1 | 0 | 0 | 0/0 |
| 24-11 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 3.38 | +15.00 | 15.00 | 118.5 | 2 | 0 | 0 | 0/0 |
| 11-8 | 1 | 1/1 | 6 | 0 | 0 | 0 | 1 | 0.38 | -3.00 | 3.00 | 28.8 | 1 | 0 | 0 | 0/0 |
| 11-24 | 1 | 1/1 | 6 | 0 | 0 | 0 | 1 | 1.50 | +12.00 | 12.00 | 118.8 | 1 | 0 | 0 | 0/0 |
