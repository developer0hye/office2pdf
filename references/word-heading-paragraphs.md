# Word heading paragraphs

- Related: #2174.
- [Typst paragraph model](https://typst.app/docs/reference/model/par/): explicit
  paragraphs accept inline content; block content inside them is ignored.
- [Typst headings](https://typst.app/docs/reference/model/heading/) are blocks.
  Their inline text does not automatically become a paragraph, so paragraph
  indentation needs an explicit paragraph inside the heading.
- Emit `heading[par(first-line-indent: ...)[runs]]`, keeping the existing outer
  wrappers, alignment, contents metadata, and paragraph spacing. Putting
  the heading inside `par` silently drops its painted and searchable content.
- Audit heading weight separately: Typst's implicit bold can override a Word
  heading whose resolved run style has no bold setting.
- The independently authored `indented_heading.docx` fixture is the three-line
  reproduction from #2174. Word 16.113.4 paints all three lines; main before
  this fix loses the middle heading. No third-party document is redistributed.
- Compiled-frame tests cover heading levels, positive/hanging/zero indents,
  continuation lines and tabs. A public-conversion test checks PDF extraction;
  generated-source string assertions alone cannot catch discarded content.
