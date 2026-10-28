# A shape casts its outer shadow from what it paints

PowerPoint casts an `a:outerShdw` from the shape's painted alpha: the fill
and the outline, each at its own alpha, composited and only then blurred and
scaled by the shadow colour's alpha. An unfilled shape (`<a:noFill/>`, or a
fill at alpha 0) therefore casts only a ring under its outline. We used to
cast every closed shape's shadow as a solid silhouette, so a `noFill`
rectangle showed a 35% grey box through its empty interior.

## Native evidence

`gt.jpg` is the submitter's PowerPoint 2019 (Windows) export of
`tests/fixtures/pptx/issue_2158_nofill_shadow.pptx`. Each shadow there is a
black image under a soft mask. That mask is a ring with a zero interior,
peaking at 27/255 under slide 1's 0.75pt outline and at 89/255 (the full 35%)
under slide 2's 7pt outline.

`scripts/probes/issue-2158-shadow-paint-alpha.json` varies slide 2's paints,
exported by PowerPoint 16 for Mac (control gate passed). Flattened shadow
mask, interior / under the outline:

| Variant | Interior | Under the outline |
| --- | ---: | ---: |
| base (`noFill`) | 0 | 89 |
| opaque fill | 89 | 89 |
| 50% fill | 45 | 89 |
| alpha-0 fill | 0 | 89 (byte-identical mask to `noFill`) |
| 50% line, `noFill` | 0 | 45 |
| 50% fill, 50% line | 45 | 66 (= 89 x (1 - 0.5 x 0.5)) |
| opaque fill, 50% line | 89 | 45 outside, 89 inside |

## Measurements

Interior pixel at (315pt, 150pt), 72 DPI: PowerPoint (255,255,255) on both
slides; main (166,166,166); this fix (255,255,255). The 5%-fuzz render
diff is empty on both pages. Our conversions of all seven probe variants
match the native exports within 4 levels at the centre and across the outline
band. #2159's slide 2 (50% fill and line) moves from (148,169,202) to
(170,191,224) against PowerPoint's (166,190,224).
