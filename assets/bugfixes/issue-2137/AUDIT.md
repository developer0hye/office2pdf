# Footnote marker style audit

The independently authored `tests/fixtures/docx/docxide/unstyled_footnotes.docx`
contains synthetic text, no metadata, external links, revisions or embedded objects.
Source SHA-256: `dce042515059bda0cf55de4b018e58b5b09d0ba9d249ff0d9430c31587c7418a`.
Fresh native Microsoft Word 16.113.4 export SHA-256: `5a01602f3f3c1e095774bc0bbd3c5e5725d8374e12c6bb1ec74b5290798bdaa8`.
Before PDF SHA-256: `5f765ae86bdd8270031b5457690a2a450ffaf4bf166f193e5a62df5ebddb6e15`; baseline is #2139 head `0f76b919`.
After PDF SHA-256: `f7aaf7f3db30442185f3da1388faba164b61b91ca189a35592bcbf244746b372`. Same source and 150-DPI Poppler renderer.

Full pages, matched crops and all 15 current material cluster crops were inspected.
The normal body reference and both normal entry labels now use the source font,
size and baseline. Literal spaces after entry labels survive in paint and searchable
text. Both normal and empty notes retain numbering; internal link destinations resolve.
Normal and explicitly superscripted references are verified separately with compiled
frames and fresh native exports; three source faces/sizes are exercised by
`footnote_marker_scripts.docx`. No bold, italic, underline or CJK runs are present.
The only hairline is the solid separator. Black ink and white fill remain correct;
no missing, overlapping or clipped text was observed in this fixture.

Remaining: #2141 tracks generic synthesized superscript placement. The second body
marker remains 1.54994pt high even with its source style retained. One-factor native
probes show ordinary superscript text and footnote references share this placement
rule. Font/size and reference-kind probes retain layout-identical re-zip controls;
font/size controls also have zero decoded-pixel difference at 150 DPI.

Remaining: #2138 tracks note paragraph/area layout: 9.35pt leading inset, baselines
1.01688pt and 2.15998pt low, and 12.6631pt pitch versus Word's 11.52pt. All 15 material
render clusters are shifted note glyph fragments and map by exact ID to #2138.
The separator's 3.6pt length difference remains sub-material under #1874.

The pinned corpus still has 63 page-count differences. This fixture audit and
conversion/text triage do not establish corpus-wide visual parity. The 25 changed
corpus PDFs have all-page geometry/text reports; their full pixel/vision audit remains
pending. Third-party source documents, fonts and screenshots are kept outside Git.
