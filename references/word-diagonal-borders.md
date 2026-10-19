# Word diagonal cell-border styles

- [MS-OE376 2.1.174](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/575e93b1-372b-4dd9-866d-01239859206a) documents Word's diagonal-style restrictions. Word keeps dotted, dashed, dashSmallGap, dotDash, dotDotDash and double; other painted styles use a single solid line. An absent direction inherits from the table style.
- Keep this mapping separate from ordinary cell edges: their `none` means suppression and their compound styles can use multiple strokes.
- Native Word 16.113.4 (16.113.26100421), macOS 26.6.2, local **Best for printing** exports: an unstyled 26-paragraph cell spans two pages. Changing only `tl2br/@val` from `single` to `none` leaves both full-page 150-DPI renders pixel-identical. `nil` and removal produce no diagonal in this unstyled probe; an absent direct direction inherits when a table style supplies it. Reversing the element to `tr2bl` reverses the line on both fragments.
- A separate no-patch re-zip export matches the base's complete geometry and pixels on both pages. The base source SHA-256 is `6ca4990b206250db8fad0c003e445f5aacb1ec8f89e3f916c7aa4e00ac345f76`.
- Therefore retain `nil` suppression, but normalize diagonal `none` to solid. Native probes correct the earlier assumption that diagonal `none` suppresses like a cardinal edge.
- Further direct/inherited probes show `none` with `sz="0"` still paints, as does direct `single` with zero size; inherited `nil` remains suppressed. Sizes 0, 1, 2 and 3 print at 0.24pt in this 300-DPI-quantized export. Retain a nominal 0.25pt minimum without copying printer quantization into positive source widths. Sizes 96 and 104 print at 12pt and 12.96pt: do not impose an unsupported maximum clamp.
- The independently authored style probe and its no-patch re-zip control match complete geometry and both full-page pixel renders. The style probe source SHA-256 is `5e2de9766d1e0f6fbc3b805ac602484db5194e2b537a1eab9e80531dcc3008b0`.
- Tests cover direct and table-style declarations, both directions, widths/colors, unsupported compound styles, painted single-line output and unchanged cardinal suppression.

The existing cell-content `move` wrapper prevents this oversized-row probe from
flowing correctly in the converter. Native per-fragment line observations do
not establish split-cell fidelity or full-file visual acceptance.
