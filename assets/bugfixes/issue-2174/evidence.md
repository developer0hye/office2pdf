# Indented heading content loss

Related: #2174. This records a defect; it does not change conversion code.

- Input: independently authored, five-part synthetic DOCX from the inline
  reproduction in #2174; only three sentinel text lines, styles and internal
  package relationships. No metadata, revisions, fonts or external content.
- Source SHA-256: `3bb18d9f332791de057d545d52ed3464a02168d81b7ab69a1c13af69601c4a82`.
- GT: freshly exported with Microsoft Word 16.113.4 for macOS through Save As,
  PDF, Best for printing, on 2026-10-16. One Letter page; Quartz PDFContext.
- GT PDF SHA-256: `a97c910c0508e76a3846311fc0eedf418a69a947d0bc4288757420d476d94f07`.
- Output: locked main `2fcceeed51cc9dc317daf1fad17509a5ef4be28d`, one page.
- Output PDF SHA-256: `d0190d8e55792452a39f9063ff6ef5240f5937cb786a097f2ed775e05984aa18`.
- CLI SHA-256: `02efaeca5c64fe347eb7f1a004c55e6f78266d2520e078ce087708250ba36a03`.
- `compare.jpg`: Word on the left, converter on the right. Both panels are
  complete 150-DPI pages; progressive JPEG quality 86, metadata stripped.

GT integrity passes. Native extraction has 55 normalized characters; output
has 30. The middle heading is absent. AFTER SENTINEL moves up 13.90533pt;
BEFORE SENTINEL differs by only 0.02416pt vertically. No wrapping or rectangle,
visibility, or visible-fill finding occurs. The tiny first-line seat remains
in #1874; heading loss and downstream reflow remain in #2174.

Full GT/output/diff pages, the matched AFTER crop, and a full-resolution strip
covering all 29 current diff clusters were inspected. Native Word paints three
black Arial lines; output paints two. No border, hairline, dash, image, CJK,
italic, underline or separate bold run occurs. All 29 current material cluster
IDs are classified under #2174; none are waived as rasterization noise. The
layout audit at 0.5pt records the missing heading and downstream shift under
#2174; the strict cluster audit passes with all 29 IDs classified.

The emitter nests a block heading inside `#par(first-line-indent: ...)`.
Typst warns that the block was ignored, but conversion exits successfully.
Positive first-line and hanging indents both reproduce; removing only the
indent retains the heading, and released 0.8.1 retains all three lines.

Fresh Word verification applies only to this authored one-page fixture. The
benchmark's fresh 16.115 oracle PDFs were not available publicly. No complete
visual pass of a third-party benchmark document is claimed by this evidence.
