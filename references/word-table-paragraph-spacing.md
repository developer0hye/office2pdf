# Word table paragraph spacing

- [Paragraph style hierarchy](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraphstyleid): table properties follow document defaults; paragraph styles and direct formatting take precedence over table properties.
- [Style inheritance](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.basedon): table styles inherit from table styles; missing parents and parents of another type are ignored.
- [Line units](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.spacingbetweenlines.line): automatic spacing uses 240 units per line; exact/minimum spacing uses twips. An omitted setting retains the inherited value.
- [Office conditional precedence](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/2ac331d4-cf1e-4fa0-8bca-6da74411e284): apply row bands, column bands, column edges, row edges, then corners.

The pinned docxide-pdf case15 declares TableGrid line240/auto and document-default line276/auto. Its Normal style and cell paragraphs do not state a higher-precedence line value. Baseline `8acc920a` drops the table property; its native comparison shows taller automatic rows. The parser now retains table line spacing with inheritance and formatting provenance. A fresh native comparison reduces mean baseline displacement from 11.6638pt to 1.0564pt and restores normalized extraction order; residual placement remains tracked separately, with worst baseline displacement 3.2984pt.

A source scan of referenced table styles and their parent chains finds line declarations in 34 corpus documents: all 46 declarations use automatic spacing (42 base, two first-row, two last-row). Conditional declarations may be inactive; this is potential coverage, not a fidelity result. Unreferenced style declarations do not establish affected scope.

Keep formatting provenance separate from resolved values: a paragraph style inheriting a document default must allow the table property to override it, while a paragraph style stating its own line spacing must keep it. Never infer provenance from a resolved numeric value.

The shared paragraph decoder currently maps `atLeast` into the IR's exact-spacing representation; preserving a minimum remains an existing IR limitation. The referenced table-style declarations in this corpus all use automatic spacing.
