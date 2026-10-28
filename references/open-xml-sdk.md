# Open-XML-SDK

- **URL:** https://github.com/dotnet/Open-XML-SDK
- **Stars:** ~4.4k
- **Language:** C#
- **License:** MIT

## What it does

Microsoft's official SDK for working with Office Open XML documents. Strongly-typed APIs for all OOXML elements — Word, Excel, and PowerPoint.

## Why relevant

- **Canonical reference** for OOXML types and relationships. When in doubt about what an OOXML element means or how it's structured, this is authoritative.
- Type hierarchy directly mirrors the ECMA-376 spec.
- Useful for cross-checking our parser's interpretation of OOXML attributes.

## Key files to study

- `DocumentFormat.OpenXml/GeneratedCode/` — Auto-generated types for all OOXML elements
- Type definitions for `WordprocessingML`, `SpreadsheetML`, `PresentationML`
- Word 2010 extensions Word itself writes, such as `w14:ligatures`: its
  [`LigaturesValues`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.ligaturesvalues)
  lists `none`, `all` and the standard/contextual/historical/discretional
  combinations in that fixed order, which `parse_ligatures_value` accepts (#2205)

## When to consult

- Verifying correct parsing of OOXML attributes and element structures
- Understanding element type relationships in the OOXML schema
- Cross-referencing with ECMA-376 spec
