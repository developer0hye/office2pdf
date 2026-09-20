# Probe: header-multiline-print-scale

- base: two-lines.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| print scale | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 60 | 1 | 1/1 | 6 | 0 | 0 | 0 | 6 | 12.30 | -22.60 | 23.60 | 40.0 | 7 | 0 | 0 | 0/0 |
| 80 | 1 | 1/1 | 6 | 0 | 0 | 0 | 5 | 5.95 | -11.00 | 11.20 | 20.0 | 7 | 0 | 0 | 0/0 |
| 120 | 1 | 1/1 | 6 | 0 | 0 | 0 | 6 | 6.00 | +11.00 | 12.40 | 20.0 | 7 | 0 | 0 | 0/0 |
