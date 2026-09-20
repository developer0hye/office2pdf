# Probe: header-multiline-font-sizes

- base: source.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| header font size | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 8 | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 2.00 | -9.00 | 9.00 | 29.7 | 3 | 0 | 0 | 0/0 |
| 12 | 1 | 1/1 | 7 | 0 | 0 | 0 | 2 | 0.33 | +2.00 | 2.00 | 15.2 | 2 | 0 | 0 | 0/0 |
| 14 | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 1.67 | +8.00 | 8.00 | 32.4 | 3 | 0 | 0 | 0/0 |
| 24 | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 9.00 | +42.00 | 42.00 | 118.8 | 3 | 0 | 0 | 0/0 |
