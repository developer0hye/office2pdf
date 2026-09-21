# Outline-level paragraphs keep their horizontal placement

`source.docx` is a synthetic zh-CN Word report template on A4 with 90pt side margins and Arial throughout. Its Title and Subtitle are Word's built-in zh-CN styles: both centred, and carrying `w:outlineLvl` 0 and 1, so the parser resolves them as headings. A left-aligned Heading 1 is the control, and the Heading 2 style indents by `w:ind w:left="840"` (42pt).

Native Word places `{Title}` at 275.12pt, `{Subtitle}` at 262.09pt and `1.1 Scope` at 132.00pt. Before the fix all three started at the 90pt margin, because the heading branch of `generate_paragraph` never applied `w:jc` or `w:ind`. The centred Normal line and the left Heading 1 already matched. After the fix, headings take the body paragraph path, and all five lines match to within 0.001pt horizontally.

`compare_layout.py --audit --fine-shift 0.5` reports 11/11 matched lines, no large or fine shifts, and matching geometry for the Title's bottom rule. The strict render census at 300 DPI finds no material cluster (`cluster-dispositions-page-1.json` is empty), and the text layer is intact. Full 150-DPI pages, the 5% difference image and 300-DPI crops of the title block and the indented heading were inspected: only glyph-edge antialiasing differs.
