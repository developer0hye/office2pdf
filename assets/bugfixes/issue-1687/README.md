# A style-less table's cell margins

A DOCX table can state no `w:tblCellMar` in a package that resolves no table
style at all. office2pdf fell through to Typst's symmetric 5pt inset, so every
cell line sat 5pt low and each row added 10pt to everything below the table.

## What Word does

Native Word for Mac 16 exports of one-factor probe packages, all Arial 11pt with
a two-column borderless table whose `w:tblW` is 9360 dxa. Every number is read
off `mutool draw -F trace`: `left inset` is the first cell's text pen x minus
the 72pt page margin, and `top inset` is its baseline minus the baseline every
zero-margin variant shares. Word quantises both to its 0.24pt device grid.

| Probe | Table style | Style `w:tblCellMar` | Direct `w:tblCellMar` | Left | Top |
| --- | --- | --- | --- | --- | --- |
| a | none | — | — | 0.48 | 0 |
| g | none | — | — (plus `w:tblInd w:w="0"`) | 0.48 | 0 |
| j | none | — | empty `<w:tblCellMar/>` | 0.48 | 0 |
| o | none | — | `top="200"` only | 0.48 | 10.08 |
| f | none | — | all four, left `0` | 0.00 | 0 |
| k | none | — | all four, left `9` (0.45pt) | 0.48 | 0 |
| l | none | — | all four, left `12` (0.60pt) | 0.72 | 0 |
| m | none | — | all four, left `15` (0.75pt) | 0.72 | 0 |
| c | none | — | all four, left `108` (5.40pt) | 5.28 | 0 |
| d | none | — | all four, left `720` (36pt) | 36.00 | 0 |
| i | none | — | `left="108"` only | 5.28 | 0 |
| b | default, named `Normal Table` | `0/108/0/108` | — | 5.28 | 0 |
| r | default, named `Normal Table` | none | — | 5.28 | 0 |
| h | default, named `Normal Table` | `top="200"` only | — | 5.28 | 0 |
| q | default, named `Grid Probe` | none | — | 0.00 | 0 |
| p | default, named `Grid Probe` | `top="200"` only | — | 0.00 | 10.08 |
| e | none | — | — (plus 0.5pt `w:tblBorders`) | 0.72 | 0.48 |

The probes fix these rules:

- **The table grid sits on the page margin.** Probe f declares a left cell
  margin of 0 and its cell text pen lands on 72.00 exactly, so Word neither
  indents nor outdents a table that states no `w:tblInd`. Probe g confirms that
  stating `w:tblInd w:w="0"` changes nothing.
- **Nothing declared means no vertical inset.** Every variant that declares no
  `w:top` puts the cell baseline at the same y, whatever its style or
  horizontal margin. Only a declared `w:top` moves it (probes o and p, 200
  twips, +10.08pt).
- **Nothing declared and no table style means a 0.48pt horizontal inset.**
  Probes k and l bracket it: 9 twips still paints 0.48 and 12 twips already
  paints 0.72, so the unquantised default is between 9 and 11 twips. 0.48pt is
  what reaches the page. A direct `w:tblCellMar` overlays per side onto it —
  probe o states only `w:top` and keeps the 0.48pt left inset.
- **The built-in Normal Table's 108 twips needs a resolved table style, and one
  Word recognises by name.** Probes b, h and r all name the default style
  `Normal Table` and all paint 5.28pt (108 twips on the 0.24pt grid) whatever
  their XML declares — h's authored `w:top="200"` is discarded with it. Probes
  q and p name theirs `Grid Probe` and paint 0.00pt. That second rule is a
  separate root cause, tracked in #1884.

So the issue's vertical claim holds and its horizontal one does not: Word's
style-less fallback is 0pt top and bottom, and **0.48pt** — not 108 twips —
left and right.

## Fix

`convert_table` in `crates/office2pdf/src/parser/docx_tables.rs` gains a last
arm: a table whose margins neither a direct `w:tblCellMar` nor a resolved table
style states takes `WORD_STYLELESS_TABLE_CELL_MARGINS`. The table-style
resolver separately applies Word's `Normal Table` margins by recognized style
name, overlays selected-style margins on the package default, and uses zero
only for sides the applicable style cascade leaves unstated (issue #1884).

## Evidence page

The evidence page is a short handover memo: a bold title, an intro paragraph
stating `w:after="200"`, a borderless three-row two-column table that declares
no `w:tblStyle` and no `w:tblCellMar`, and a closing paragraph stating
`w:before="300"`. Its `styles.xml` defines `Normal`, `Body Text` and `caption`
and no table style. The GT is the native Word for Mac 16 export.

Against the GT, the before output puts each cell line 5.02pt, 15.18pt and
25.11pt low by row and the closing paragraph 30.16pt low, each cell line also
4.62pt right. The after output is within 0.21pt vertically and 0.10pt
horizontally everywhere; both residuals are below the materiality bar and are
tracked in #1874.

The script below regenerates the package part for part. The package used had
SHA-256 `ab54a8d92320f6ecc4d40821c2cbc433f5081135353bf38b213025ef901076ad`; a
re-run matches every part, but zip timestamps change its hash:

```python
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def paragraph(text, style=None, before=None, after=None, bold=False):
    properties = f'<w:pStyle w:val="{style}"/>' if style else ""
    spacing = ""
    if before is not None:
        spacing += f' w:before="{before}"'
    if after is not None:
        spacing += f' w:after="{after}"'
    if spacing:
        properties += f"<w:spacing{spacing}/>"
    run_properties = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return (
        f"<w:p><w:pPr>{properties}</w:pPr><w:r>{run_properties}"
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def table(rows, width, columns):
    grid = "".join(f'<w:gridCol w:w="{column}"/>' for column in columns)
    body = ""
    for row in rows:
        cells = ""
        for column, text in zip(columns, row):
            cells += (
                f'<w:tc><w:tcPr><w:tcW w:w="{column}" w:type="dxa"/></w:tcPr>'
                f"{paragraph(text)}</w:tc>"
            )
        body += f"<w:tr>{cells}</w:tr>"
    # No `w:tblStyle` and no `w:tblCellMar`: the cell margins come from Word's
    # style-less fallback, because this package defines no table style.
    return (
        "<w:tbl><w:tblPr>"
        f'<w:tblW w:w="{width}" w:type="dxa"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"
    )


def paragraph_style(style_id, name, based_on="Normal", is_default=False):
    default = ' w:default="1"' if is_default else ""
    parent = f'<w:basedOn w:val="{based_on}"/>' if based_on else ""
    return (
        f'<w:style w:type="paragraph"{default} w:styleId="{style_id}">'
        f'<w:name w:val="{name}"/>{parent}</w:style>'
    )


COLUMNS = (4680, 4680)
ROWS = (
    ("Warehouse lease", "Closes 31 March"),
    ("Night-shift hiring", "Two engineers, April"),
    ("Billing migration", "New cluster, May"),
)

body = (
    paragraph("Operations handover", bold=True)
    + paragraph("The table below lists the three items that carry over.", after=200)
    + table(ROWS, sum(COLUMNS), COLUMNS)
    + paragraph("Questions go to the shared tracker.", before=300)
)

styles = (
    f'<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "</w:pPrDefault></w:docDefaults>"
    + paragraph_style("Normal", "Normal", based_on=None, is_default=True)
    + paragraph_style("BodyText", "Body Text")
    + paragraph_style("Caption", "caption")
    + "</w:styles>"
)
settings = (
    f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
)
document = (
    f"<w:document {W}><w:body>{body}"
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" '
    'w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
)
parts = {
    "[Content_Types].xml": (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/word/document.xml" ContentType="{CT}.document.main+xml"/>'
        f'<Override PartName="/word/styles.xml" ContentType="{CT}.styles+xml"/>'
        f'<Override PartName="/word/settings.xml" ContentType="{CT}.settings+xml"/></Types>'
    ),
    "_rels/.rels": (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'
    ),
    "word/document.xml": document,
    "word/_rels/document.xml.rels": (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/></Relationships>'
    ),
    "word/styles.xml": styles,
    "word/settings.xml": settings,
}

if __name__ == "__main__":
    out = "issue-1687-evidence.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as package:
        for name, xml in parts.items():
            package.writestr(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml)
    print(out)
```

Each probe package in the table above is that same script with one factor
changed: `table()` gains the stated `w:tblCellMar` or `w:tblBorders`, and
`styles` gains the stated table style.
