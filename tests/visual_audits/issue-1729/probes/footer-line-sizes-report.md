# Probe: footer-line-sizes

- base: base.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| footer size | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 8 | 1 | 1/1 | 6 | 0 | 0 | 0 | 1 | 0.38 | +3.00 | 3.00 | 31.9 | 0 | 0 | 0 | 0/0 |
| 12 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 0.38 | -2.00 | 2.00 | 9.2 | 0 | 0 | 0 | 0/0 |
| 14 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 0.62 | -4.00 | 4.00 | 29.9 | 0 | 0 | 0 | 0/0 |
| 24 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 2.62 | -18.00 | 18.00 | 116.2 | 1 | 0 | 0 | 0/0 |
