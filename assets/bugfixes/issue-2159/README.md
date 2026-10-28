# Style-reference alpha paints with the theme fill and line

A `<p:style>` reference colour's `<a:alpha>` belongs to the paint that colour
supplies. When the `a:fillRef` colour wins the fill, its alpha masks the
resolved theme fill, flat or gradient; when the `a:lnRef` colour wins the
outline, its alpha strokes the outline. A fill or line declared in `p:spPr`
wins with its own alpha, or with none. The `Shape` IR gained `stroke_opacity`,
so a direct `<a:ln>` alpha on a shape now reaches the outline as well.

## Native evidence

`gt.jpg` is the submitter's PowerPoint 2019 (Windows) export of
`tests/fixtures/pptx/issue_2159_style_ref_alpha.pptx`. Slide 2's ellipse fill
is the theme gradient (fill style 3) under `/GS46`, a luminosity soft mask
whose shading is a constant 0.50196 on every function piece: one uniform alpha
over the whole resolved gradient, nothing baked per stop. Its outline strokes
under `/GS47 /CA .50196`. 0.50196 = 128/255, the byte our codegen writes.

## Measurements

Pixel at slide (315pt, 150pt), 72 DPI:

| Render | Slide 1 | Slide 2 |
| --- | --- | --- |
| PowerPoint | (140,180,245) | (166,190,224) |
| Before | (131,173,239) | (131,173,239) |
| After | (131,173,239) | (148,169,202) |

## Remaining

- Slide 2 interiors read about 20 levels darker: the theme `outerShdw` is cast
  at its full 35% under the translucent fill, while PowerPoint attenuates it
  by the fill alpha (shadow SMask max 89 on slide 1, 45 on slide 2). A copy
  with the theme shadow alpha halved converts to (170,191,224) at the same
  pixel. Tracked by #2158, the alpha-0 case of the same rule.
- Sub-material, recorded under #1874: the theme-shadow halo below opaque
  shapes is up to ~28 levels darker; PowerPoint's opaque theme gradient uses
  an eased 512-sample ramp and a lighter top stop (155,193,255 vs our
  164,196,255); the theme line's `shade`/`satMod` is not applied to the
  `lnRef` colour (74,126,187 vs 79,129,189).
- `<a:fontRef>` alpha is not handled: the deck has no text in those shapes, so
  it gives no native evidence.
