# Apache POI

- **URL:** https://github.com/apache/poi
- **Stars:** ~4k
- **Language:** Java
- **License:** Apache-2.0

## What it does

Java OOXML implementation for Word, Excel, and PowerPoint.

## Why relevant

- `XSLFTableRow#getHeight()` and `setHeight()` map directly to `CTTableRow.h`.
- Useful as a control sample: it exposes PPTX row height as raw OOXML data, not a rendered layout result.
- Confirms that object-model libraries generally do not resolve PowerPoint's final table reflow for us.

## Key files to study

- `poi-ooxml/src/main/java/org/apache/poi/xslf/usermodel/XSLFTableRow.java`
- `poi-ooxml/src/main/java/org/apache/poi/xslf/usermodel/XSLFTableCell.java`

## When to consult

- Verifying how another mature OOXML library exposes PPTX table row/cell properties
- Cross-checking whether a behavior comes from raw XML vs. actual renderer layout
