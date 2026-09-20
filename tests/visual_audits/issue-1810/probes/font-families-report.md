# Probe: header-multiline-font-families

- base: source.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| header font family | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Calibri | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 0.33 | -1.00 | 1.00 | 7.4 | 0 | 0 | 0 | 0/0 |
| Aptos | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 0.33 | -1.00 | 1.00 | 1.3 | 0 | 0 | 0 | 0/0 |
| Times New Roman | 1 | 1/1 | 7 | 0 | 0 | 0 | 0 | 0.00 | +0.00 | 0.00 | 5.1 | 0 | 0 | 0 | 0/0 |
| Malgun Gothic | 1 | 1/1 | 7 | 0 | 0 | 0 | 3 | 1.33 | +7.00 | 7.00 | 8.6 | 1 | 0 | 0 | 0/0 |
