# Word aligned tabs in the right indent

- [Tab stop semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tabstop?view=openxml-3.0.1)
- [Pinned docxide-pdf layout](https://github.com/sverrejb/docxide-pdf/blob/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/src/pdf/layout.rs#L2552)
- [Typst measure](https://typst.app/docs/reference/layout/measure/)

The upstream renderer allows aligned fields to reach into the right indent,
up to the physical column edge. Its Word reference has short contents entries
with a hanging indent, a body tab, and a trailing right-aligned page-number tab.
The ordinary text wraps before that field's tab stop.

Our ordinary inline spacing made that field wrap to another line. Measuring
the whole prefix also selects a stop from its unwrapped width, not the cursor
on its last rendered line. Word tab positions also need the first-line inset
removed when the prefix is emitted inside an indented block.

For a left-aligned LTR flow paragraph with a final aligned stop in the right
indent, measure the final field once. If it fits the reserved indent and column,
emit the prefix with its original wrap width and locate the field from the
actual final-line cursor. A zero-width host permits its paint to extend into
the indent without making the ordinary text column wider. Render run sources
once so alternative layout projections do not repeat generation side effects.

Compiled controls cover short right/center/decimal fields, wrapped body text,
hanging/first-line indents, and a DOCX package. Full native visual comparisons
remain required. Wide or multiline fields, RTL/non-left paragraphs, nonterminal
overhangs, and stops beyond the physical column retain the existing layout;
their complete Word wrapping rules are outside this terminal-field change.
