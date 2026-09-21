# A header rule hangs below its line instead of parting the story

`source.docx` is a synthetic A4 package in Arial 10.5pt whose header holds two paragraphs, the first ruled off with `<w:bottom w:val="single" w:sz="6" w:space="1"/>`, over two body lines. `w:header` is 851 twips (42.55pt) under a 72pt top margin.

A ruled paragraph is block-level, which split the story — its paragraphs are joined by a line break — into two Typst paragraphs with the default paragraph gap between them. The second header line landed 20.32pt low, over the first body line, and the taller measured story clamped away the band shift, lifting the first line 2.49pt. Stories are now laid out as one block per paragraph, each as tall as Word's line plus whatever its rules and their `w:space` reserve, which is the same term the band measurement already used.

Measured on native Word 16.113.1 across one-factor probes (`w:space` 0, 1, 8 and absent; 10.5pt and 18pt; Arial and Malgun Gothic). Word seats the ruled line where an unruled one sits, hangs the rule below the line's bottom edge, and starts the next paragraph under the rule:

| | native Word | before | after |
| --- | ---: | ---: | ---: |
| header line 1 | 52.56 | 50.07 | 52.06 |
| header line 2 | 66.24 | 86.56 | 65.88 |
| Malgun 10.5pt, ruled, line 1 | 56.16 | 50.09 | 56.07 |
| Malgun 10.5pt, ruled, line 2 | 76.08 | 95.07 | 75.98 |

The rule's own gap follows the same measurements: Word seats the following line identically for `w:space="0"` and for a `w:bottom` stating no `w:space` at all, so the 0.5pt hairline the code used to substitute for an absent gap is gone.

After the fix the layout audit reports no large shift (one of +20.32pt before) and one fine shift of 0.50pt, the story's first-baseline seat tracked in #1640, which the single render cluster — the rule 0.4pt high — also belongs to. The text layer is unchanged. A top rule still seats its text a cap height rather than an ascent from the rule, which is #1828.
