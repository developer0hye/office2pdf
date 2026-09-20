# Probe: footer-line-wrapping

- base: base.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| long footer line | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| first | 1 | 1/1 | 5 | 1 | 2 | 0 | 0 | 0.00 | +0.00 | 0.00 | 0.0 | 0 | 0 | 0 | 0/0 |
| last | 1 | 1/1 | 5 | 1 | 2 | 0 | 1 | 2.00 | -14.00 | 14.00 | 0.0 | 1 | 0 | 0 | 0/0 |
