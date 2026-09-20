# Probe: footer-line-mixed-sizes

- base: base.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| first or second line font size | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| first8 | 1 | 1/1 | 6 | 0 | 0 | 0 | 0 | 0.00 | +0.00 | 0.00 | 29.8 | 0 | 0 | 0 | 0/0 |
| first24 | 1 | 1/1 | 6 | 0 | 0 | 0 | 1 | 0.38 | -3.00 | 3.00 | 116.2 | 0 | 0 | 0 | 0/0 |
| second8 | 1 | 1/1 | 6 | 0 | 0 | 0 | 1 | 0.38 | +3.00 | 3.00 | 31.9 | 0 | 0 | 0 | 0/0 |
| second24 | 1 | 1/1 | 6 | 0 | 0 | 0 | 2 | 2.25 | -15.00 | 15.00 | 114.3 | 1 | 0 | 0 | 0/0 |
