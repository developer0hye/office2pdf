# PowerPoint BulletFormat Docs

- **URL:** https://learn.microsoft.com/en-us/office/vba/api/powerpoint.bulletformat
- **Type:** Microsoft Learn API docs

## What it does

Documents PowerPoint's bullet object model: character, numbering, font, color, and relative size.

## Why relevant

- `UseTextFont` confirms bullets can inherit the first text character's font.
- `UseTextColor` confirms bullets can inherit the first text character's color.
- `RelativeSize` matches PPTX bullet-size semantics relative to text.
- Helps map OOXML `buChar` / `buAutoNum` / `buFontTx` / `buClrTx` / `buSzTx` behavior back to PowerPoint UI behavior.

## Key pages to consult

- `BulletFormat`
- `BulletFormat.UseTextFont`
- `BulletFormat.UseTextColor`
- `BulletFormat.RelativeSize`

## When to consult

- Preserving PowerPoint bullet markers instead of falling back to generic Typst bullets
- Verifying when bullet font/color/size should follow the paragraph text vs. use explicit overrides
