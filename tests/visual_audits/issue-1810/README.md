# Multiline worksheet header spacing

Initial defect evidence only. `source.xlsx` is a synthetic Letter worksheet with Arial 11 text and three explicit LF-separated lines in `oddHeader`. The native re-zip control passed.

Native header baselines are 51, 65 and 79pt. The converter baselines are -25.634766, 15.882629 and 57.400026pt: 41.517397pt advances instead of 14pt, with the first line outside the page. The comparison uses native Excel on the left and converter output on the right at 150 DPI.

Full pages, the difference image and all six matched shift crops were inspected. Native has three evenly spaced header lines; output hides the first and spreads the remaining two. Body labels, numbers and footer remain visible. The page has no images, borders, hairlines, bold/italic/underlined runs or rotations. Separate remaining findings are header placement (#1731), body baseline rounding (#1721), and header glyph origin differences (#1659). This evidence does not establish a fix or an overall visual pass.
