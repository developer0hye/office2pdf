# Word paragraph page breaks

- [Microsoft `pageBreakBefore` specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pagebreakbefore): an omitted setting inherits through the style hierarchy; a setting never stated is off. A true value starts the paragraph on a new page even across a keep-next chain.
- [python-docx paragraph formatting](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/text/paragraph-format.html#page-placement): pagination properties are tri-state; explicit false clears inheritance.
- [docx-rs issue #939](https://github.com/bokuweb/docx-rs/issues/939) and [upstream fix #940](https://github.com/bokuweb/docx-rs/pull/940): the published reader loses false pagination flags. Read current XML properties for compatibility until a released dependency preserves them.

The parser combines `wordWrap` and `pageBreakBefore` in one paragraph cursor and style scan. Only live, direct paragraph properties count; revision history, fallback text boxes and character-style properties cannot replace them. Resolve breaks through document defaults and same-type `basedOn` chains, then apply the direct paragraph override. Unresolved style IDs use the default paragraph style. Existing cycle handling discards cyclic ancestry.

Behavioral tests cover inheritance, lexical false forms, direct re-enabling, document/default styles, revision history, cycles and character-style isolation. A compiled-frame test checks actual page membership and verifies that a cleared break leaves following text on the inherited paragraph's page. These tests establish the rule, not corpus-wide visual acceptance.

Use `Block::PageBreakBefore` for paragraph requirements and retain `Block::PageBreak` for hard run breaks. [Typst weak page breaks](https://typst.app/docs/reference/layout/pagebreak/#parameters-weak) reuse an empty page. Both forms split column/continuous-section containers at the top level, preserving break strength. Compiled tests cover first-page, explicit-break, equal/unequal-column, continuous-section and intentional-blank-page boundaries.
