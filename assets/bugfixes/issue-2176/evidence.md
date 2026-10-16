# Regular headings rendered in bold

Related: #2176, #2174. This records a separate weight defect, not a fix.

- Source: independently authored three-line `indented_heading.docx` from
  [the nesting fix](https://github.com/developer0hye/office2pdf/blob/8f46b82370f02c8b2d496b675cd0d4ffb09a08a1/tests/fixtures/docx/indented_heading.docx).
  Five package parts; no metadata, revisions, embedded objects or external links.
- Source SHA-256: `3bb18d9f332791de057d545d52ed3464a02168d81b7ab69a1c13af69601c4a82`.
- GT: Microsoft Word 16.113.4 on macOS 26.6.2 build 25G83, freshly exported
  2026-10-16 through Save As PDF, Best for printing; one Letter page.
- GT PDF SHA-256: `a97c910c0508e76a3846311fc0eedf418a69a947d0bc4288757420d476d94f07`.
- Output: nesting-only fix `8f46b82370f02c8b2d496b675cd0d4ffb09a08a1`.
- Output PDF SHA-256: `78f3ae922e4f5f99981532a572df8f46c3cd57db68b59e09f21442ab179281b4`.
- Locked ci-profile CLI SHA-256: `85ae4234c6a4fbaf0a65691d148b61fec72226445ec86fcf14a0e0d3a49393d4`.
- `compare.jpg`: complete 150-DPI pages, Word left and converter right;
  progressive JPEG quality 86, metadata stripped.

Source-aware GT integrity and normalized text comparison pass. All three lines
are present and their horizontal origins agree. Worst baseline deviation is
0.106501pt, below the 0.12pt DOCX trace noise floor; the 0.5pt geometry audit
finds no shifted instance, wrap/reflow, rectangle, visibility or fill mismatch.

Full GT/output pages, diff and original 150-DPI heading crops were inspected.
The heading is visibly heavier in output. `pdffonts` confirms Arial-BoldMT
where native uses regular ArialMT. All eight current material cluster IDs
are classified under #2176, with none waived as rasterization noise. No border,
dash, fill, shape, image, CJK, italic or underline occurs in the fixture.

The zero-indent main control already paints the heading bold. The nesting fix
exposes this independent semantic-heading default; it does not establish full
heading fidelity. The ordinary body lines remain regular Arial. No corpus-wide
visual acceptance or conversion performance conclusion is claimed.
