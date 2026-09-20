# Multiline worksheet header spacing

`source.xlsx` is a synthetic Letter worksheet with Arial 11 text and three explicit LF-separated lines in `oddHeader`. Native re-zip controls passed for the line-count, font, size, mixed-size, print-scale, section and wrapping probes.

Native header baselines are 51, 65 and 79pt. Before the fix they were -25.634766, 15.882629 and 57.400026pt: 41.517397pt advances instead of 14pt, with the first line outside the page. The current baselines are 55.400026, 69.400028 and 83.400028pt. All three lines are visible and normalized extracted text matches the reference.

Text-only headers now grow downward from a first-line box, using native line advances and whole-sheet-point descent rounding. Wrapped paragraphs reserve every visual line without moving the first anchor. Sections align at their tops and keep independent line counts; a single-line section sits one scaled sheet point below a multiline section with the same font metrics. Footers retain their existing last-line seats and bottom alignment. Image-bearing stories retain the existing flow layout; native image parity is not claimed.

Full 150-DPI GT/output pages, the difference image, all six matched shift crops and a full-scale footer pair were inspected. Header spacing and clipping are fixed. Header placement remains 4.40003pt low (#1731); body labels and values remain 1pt high (#1721); small header glyph-origin differences remain (#1659). There are no images, borders, hairlines, bold/italic/underlined runs or rotations in this page. The strict cluster report assigns every material cluster to an open issue; this is not an overall visual parity claim.

The `probes/` directory preserves synthetic inputs, one-factor specifications, native control reports and baselines. Compiled regressions cover font families/sizes, mixed sizes, print scales, independent sections and wrapping. README usage, setup and API remain unchanged.
