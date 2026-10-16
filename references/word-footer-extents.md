# Word footer extents

- ECMA-376, pgMar@bottom: nonnegative values reserve the greater of the declared
  bottom margin and footer extent; negative values fix the body boundary at the
  absolute distance, allowing overlap. [Specification text](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_pgMar_topic_ID0EDACT.html).
- [docxide-pdf, pinned implementation](https://github.com/sverrejb/docxide-pdf/blob/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/src/pdf/header_footer.rs#L676)
  selects the current footer and computes max(bottom, footer distance + height),
  unless the bottom margin is fixed.
- [Typst page footer descent](https://github.com/typst/typst/blob/v0.15.1/crates/typst-library/src/layout/page.rs)
  defaults to 30% of the bottom margin. Set it to zero when positioning the footer
  explicitly, or growing the body margin also moves the footer.

Implementation: `render/typst_gen_footer_extents.rs` measures the paragraph footer
content at the actual section width and uses the same content for painting.
It includes the last paragraph's after-spacing in the bottom-edge reservation.
First-page capacity differs through a signed parent-region float, so continuation
pages keep the default footer capacity and the first-page footnote region moves
with the selected footer capacity.
An explicitly empty first footer suppresses default-story painting and reservation.
Variants without an explicit edge distance keep their prior placement.
Paragraph shaping, empty paragraph marks, and ruled/wrapped paragraph line layout
remain the existing renderer's responsibility; measurement does not repair them.

Compiled regressions cover body clearance, upward growth at a fixed page-edge
anchor, wrapped lines, negative fixed margins, both first/default height orders,
and the first-page footnote region. This is layout behavior coverage, not native
visual acceptance of the entire upstream corpus.
