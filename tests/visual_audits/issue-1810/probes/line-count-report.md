# Probe: sheet-header-line-count

- base: line-count-base.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| header line count | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 2 | 1 | 1/1 | 5 | 0 | 1 | 0 | 1 | 0.14 | -1.00 | 1.00 | 0.0 | 0 | 0 | 0 | 0/0 |
| 3 | 1 | 1/1 | 5 | 0 | 2 | 0 | 1 | 0.14 | -1.00 | 1.00 | 0.0 | 0 | 0 | 0 | 0/0 |
