# docxide-pdf regression fixtures

`page_borders.docx` is the synthetic `cases/case68/input.docx` from
[docxide-pdf at 1bb140b6](https://github.com/sverrejb/docxide-pdf/blob/1bb140b662a69f8786015c1b532d37d6084d0934/tests/fixtures/cases/case68/input.docx),
licensed under Apache-2.0. It declares four 3pt blue page borders, 24pt from the
page edge. The default parser/render regression reads this tracked package.
`line_numbering.docx` is the synthetic `cases/case70/input.docx` from the same
revision and license. Its twelve Cambria body lines use continuous line numbering;
native Word displays labels 2–13. A default conversion regression reads it.
No fonts or reference PDFs are copied from the upstream corpus.

`line_numbering_page_restart.docx` changes only case70’s section line-number
restart to `newPage`; native Word still displays 2–13 for `start=1`.

`line_numbering_disabled.docx` removes case70’s line-number attributes.
Fresh native Word renders its twelve body lines without margin numbers.

`section_vertical_center.docx` is the unmodified synthetic case69.
`section_alignment_continuous.docx` duplicates its two paragraphs, adding a
continuous section boundary with the same centered alignment.
`section_alignment_multipage.docx` adds an explicit second centered page, then
a top-aligned continuous section sharing only that second page. All retain the
upstream Apache-2.0 attribution above; native calibration variants stay external.

`legal_numbering.docx` is the unmodified synthetic case71 from the same revision
and Apache-2.0 license. Native Word retains Roman parent labels while legal
child labels use decimal references. Native spacing probes and compiled
geometry tests cover the independent nested-list transition defect #2083.
Inline-prefix tests vary ancestor counters, starts, Roman/alphabetic/zero-padded
formats, legal flags and literal text against native Word probes.

`numbering_restarts.docx` is the unmodified synthetic case72 from the same
revision and Apache-2.0 license. Its children continue across parent items.
`numbering_restart_policy.docx` derives a three-level sequence from case72;
compiled PDF-text tests vary the restart trigger and complete level overrides.
Native calibration packages and PDFs remain outside git.

`numbering_skipped_levels.docx` removes the restart flag from the three-level
policy fixture. Compiled geometry tests compare its skipped child's marker/body
origins and face, then vary nesting and marker format to cover shallower returns.

`numbering_implicit_ancestors.docx` derives a four-level 0/2/3/1/2/0/2 sequence
from case72. PDF-text tests vary starts, level order and restart rules against
native Word probes for missing ancestors and later visible counters.

`numbering_restart_suppression.docx` sets the three-level policy fixture’s level 2
restart to 2. PDF tests check suppressed markers, numbering versus paragraph indents,
bullet/numbered levels and continuation spacing against native Word probes.

`numbering_out_of_range_restarts.docx` sets the four-level derivative’s level 2
restart to 8. PDF tests vary levels, starts, thresholds and complete overrides
against native Word probes for continuing physical counters.

`numbering_wide_markers.docx` derives from the same Apache-2.0 case72: child
start 7 exposes a marker wider than its hanging span. Native probes vary tab
intervals, custom left/numbering stops, starts, wrapping, marker separators and
level run properties. Inline-marker geometry tests cover 16pt/22pt labels,
body-controlled proportional advance and wrapping against native Word exports.

`alternate_namespace_runs.docx` is an authored synthetic report with two Arial
11pt lines, the first bold, using the `ns0` WordprocessingML prefix. Default
parser and compiled-layout tests vary prefixes and scoped bindings; fresh Word
exports verify the prefix-only control. No third-party content is included.

`valid_png_report.docx` is an authored Cambria report with an inline 120×60pt
PNG and surrounding text. `truncated_png_report.docx` changes only that PNG's
tail. Parser and conversion tests require the picture's extent and surrounding
text to survive damaged media; runtime-generated variants cover bad checksums,
RGB/alpha/grayscale samples and header/footer pictures. Native Word exports
remain outside git. Neither package contains third-party content.

`smartart_cached_drawing.docx` is synthetic case60 from the pinned Apache-2.0
revision above. It contains two cached diagrams with curved panels, rounded
boxes, arrows, bullets and mixed text styles; no photographs or font binaries.
Tests require searchable labels, frame extents and renamed-part lookup.

`smartart_picture_fill.docx` derives the first cached diagram from case59 at
that same Apache-2.0 revision. All five photographs are replaced by an authored
800×600 color-coordinate grid; no photographs or embedded font binaries remain. Tests
vary signed destination offsets and source crops while preserving the shape frame.

`floating_picture_stretch.docx` is independently authored under this repository's
Apache-2.0 license, including its four-band PNG. Four margin-relative floating
pictures vary frame proportions, rotation and an explicit source crop. A default
compiled-frame regression checks declared dimensions. Native Word packages and
PDFs remain external; source-crop and placement defects are tracked separately.


`unstyled_footnotes.docx` is independently authored under this repository's
Apache-2.0 license. Two unstyled references use nonsequential IDs; one shares
its run with surrounding text. The default parser and compiled-frame regression
requires both note bodies and the surrounding text to survive. Native Word
exports remain external.
