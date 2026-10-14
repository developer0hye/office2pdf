# Native fixture audit

Microsoft Word 16.113.4 reference and current output were inspected at 150 DPI,
including full pages, the 5% diff, both geometry crops and all 21 cluster crops.

- Reference recovery: both note bodies and body references are present. All
  three body paragraphs remain on one page, with no clipped or hidden text.
- Marker formatting (#2137): the first baseline 12pt body marker is incorrectly
  superscripted; the second explicitly superscripted marker also differs in
  scale/advance. Entry labels are raised instead of baseline 10pt text.
- Note geometry (#2138): note-body insets, the two baselines and line pitch
  differ. The separator is lower; its 3.6pt length difference is sub-material
  (2.5%) and tracked in #1874. No dashed rules, shapes, rotated content, images,
  coloured fills, bold, italics or underlining occur in this synthetic fixture.
- The empty-first-note native variant retains both numbers after the parser
  fix; the same marker and paragraph defects remain. Its 12 current clusters
  were separately inspected and assigned to #2137 or #2138.

The strict report classifies current clusters; it does not claim visual parity.
The nine changed corpus documents have full geometry/text reports, but their
complete high-DPI visual audit remains pending. No corpus-wide visual pass is
claimed here.
