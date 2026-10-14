# Generic superscript placement

Independent synthetic source: `tests/fixtures/docx/docxide/footnote_marker_scripts.docx`.
Source SHA-256: `9dd8fb0d6e8d4e650145e928498bb71b2f82699559b854f15391124b3e0b363d`.
Fresh native Word 16.113.4 is left; the marker-style candidate is right.
Both full pages, pixel diffs and matched source-resolution crops were inspected at
150 DPI. The source has three normal control runs and explicitly superscripted
references in Arial 8pt, Times New Roman 14pt and Courier New 18pt.

Measured rise relative to normal text in the same paragraph:

| Face/source size | Word | Candidate |
| --- | ---: | ---: |
| Arial/8pt | 3.12000pt | 3.81641pt |
| Times New Roman/14pt | 5.04000pt | 6.34375pt |
| Courier New/18pt | 5.51994pt | 7.60251pt |

The last two body markers are visibly too high. The source face and vertical
alignment are retained; the candidate uses Typst's synthesized-script metrics.
Native one-factor probes replace only a reference with ordinary superscript text
and obtain identical script size/rise for Arial, Courier New and Cambria. Their
no-patch re-zip controls are layout-identical. Font/size probes also have zero
control pixel delta at 150 DPI. Native size and face probes do not support one
constant translation for all fonts/sizes; the exact rule remains unproven.

The note-area inset, compression and baselines visible in the lower panel remain
#2138. This is defect evidence, not a claim of whole-document parity. No external
or classified content, font binaries, metadata, revisions or embedded objects are
present in the authored package.
