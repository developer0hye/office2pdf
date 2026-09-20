# Probe: header-multiline-sections

- base: two-lines.xlsx
- part: xl/worksheets/sheet1.xml
- backend: office
- control: layout-identical re-zip (bytes differ)

| section font size and line count | patched | pages | matched | missing | extra | wraps | deviant | mean abs dy (pt) | worst dy (pt) | worst pitch (pt) | worst width (%) | large shifts | visibility | rect geometry | rects |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| different-sizes | 1 | 1/1 | 4 | 2 | 6 | 0 | 0 | 0.00 | +0.00 | 0.00 | 0.0 | 0 | 0 | 0 | 0/0 |
| unequal-counts | 1 | 1/1 | 4 | 2 | 4 | 0 | 0 | 0.00 | +0.00 | 0.00 | 0.0 | 0 | 0 | 0 | 0/0 |
