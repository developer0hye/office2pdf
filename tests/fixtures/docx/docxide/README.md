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

`first_header_exact_line.docx` is the unmodified synthetic
`cases/first-header-no-minimum-control/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/first-header-no-minimum-control),
Apache-2.0. Its first-page header is one 160pt exact line 6pt below the page
top. Native Word seats the header baseline at 134.08pt, 0.8 of the line down,
and starts the body at 166pt; compiled-layout regressions check both
baselines.

`first_header_small_control.docx` is the unmodified synthetic
`cases/first-header-small-control/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/first-header-small-control),
Apache-2.0. Like `first_header_exact_line.docx`, it has no styles part: its
header run states no face or size, and its Arial 12pt body lines state no
`w:line`. Native Word sets the run in Aptos 12pt, seats it 0.8 of its 12pt
exact line down at 15.50pt, and sets the body lines 16pt apart. Parser and
compiled-layout regressions check all three.

`first_header_row_overflow.docx` is the unmodified synthetic
`cases/first-header-no-minimum-overflow/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/first-header-no-minimum-overflow),
Apache-2.0. The 160pt first-page header is followed by `Before` and a one-cell
table row with neither `w:cantSplit` nor a height rule, holding `Minimum`,
`Second` and 60 `Continuation` paragraphs. Native Word starts the row on page 1
below `Before` and continues it over three pages. A compiled-layout regression
checks that the row starts on page 1 and paints every line once, above the
bottom margin.

`page_header_ordinary_marker.docx` is the unmodified synthetic
`cases/page-header-ordinary-marker/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/page-header-ordinary-marker),
Apache-2.0. Nothing in its styles cascade states a `w:sz`: `w:docDefaults`
names Times New Roman, and `Normal` names Tahoma with only `w:szCs="24"`.
Native Word sets the first-page header and the body text in Tahoma 10.
Compiled-layout regressions check that both are set at 10pt, and that both
paint in Tahoma or one of its sans stand-ins rather than the default serif.

`continuous_framed_section_mark.docx` is the unmodified synthetic
`cases/cont-first-first-fits/input.docx` from
[docxide-pdf at adb8c2b8](https://github.com/sverrejb/docxide-pdf/tree/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/tests/fixtures/cases/cont-first-first-fits),
Apache-2.0. Thirty-eight exact 16pt lines end in an empty paragraph carrying a
`w:framePr` and the section's `w:sectPr`; the continuous next section opens with
24pt before an exact 16pt line. Native Word puts that line's baseline 40pt below
the last filler line, and a compiled-layout regression checks both positions.

`empty_frame_control.docx` is the unmodified synthetic
`cases/empty-frame-control/input.docx` from the same revision and license: an
empty `w:framePr` paragraph, then "Following paragraph". Native Word sets the
following line on the first body baseline, 85.03pt, as if the frame were absent;
a compiled-layout regression checks it.

The eight `inline_picture_indent_*.docx` fixtures are the unmodified synthetic
`cases/inline-picture-indent-*/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases),
Apache-2.0, named by their `w:jc` and `w:ind` variant. Each holds one paragraph
with only a 180×90pt inline picture and 72pt side margins. Native Word puts the
picture where the paragraph's first line of text would start, inside its
indents; a compiled-layout regression checks the eight left edges.

`empty_frame_group.docx` is the unmodified synthetic
`cases/empty-margin-frame-14-two-empty/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/empty-margin-frame-14-two-empty),
Apache-2.0. A two-column table is followed by three empty paragraphs sharing
one `w:framePr`, the last carrying the section's `w:sectPr`, and then
"Public greeting". Native Word sets the greeting on the line after the table,
as if the frame were absent; a compiled-layout regression checks its baseline.

`empty_frame_box.docx` is the unmodified synthetic
`cases/empty-margin-frame-15-two-empty/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/empty-margin-frame-15-two-empty),
Apache-2.0: `empty_frame_group.docx` in compatibility mode 15. There native
Word keeps the empty frame's box at the top margin and starts the table below
it, 24pt lower; a compiled-layout regression checks the table's first line and
the greeting.

`offset_text_frame.docx` is the unmodified synthetic
`cases/paragraph-offset-frame-immediate/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/paragraph-offset-frame-immediate),
Apache-2.0. A one-cell table ("Recipient") is followed by a paragraph in a
text-anchored `w:framePr` frame with a stated height, hung 16.8pt (`w:y`) below
its place in the flow, and then by "Dear recipient". Native Word sets "Dear
recipient" on the line after the table and the framed line 16.75pt below it; a
compiled-layout regression checks both baselines.

`floating_table_wrap.docx` is the unmodified synthetic
`cases/leading-table-clear-leading-no-clear/input.docx` from
[docxide-pdf at adb8c2b8](https://github.com/sverrejb/docxide-pdf/tree/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/tests/fixtures/cases/leading-table-clear-leading-no-clear),
Apache-2.0. A 241pt `w:tblpPr` table with `w:rightFromText="180"` sits at the
left margin, followed by a centred greeting and a body paragraph. Native Word
lays both beside the table, in the room right of it. A compiled-layout
regression checks both lines against that room.

`floating_table_clear_break.docx` is the unmodified synthetic
`cases/leading-table-clear-leading-clear/input.docx` from
[docxide-pdf at fe7b8bea](https://github.com/sverrejb/docxide-pdf/tree/fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed/tests/fixtures/cases/leading-table-clear-leading-clear),
Apache-2.0. It differs from `floating_table_wrap.docx` only by `w:clear="all"`
on the break before the greeting. Native Word sets the greeting below the table,
centred on the whole column, and the body paragraph two lines below it at the
margin; a compiled-layout regression checks both lines.
