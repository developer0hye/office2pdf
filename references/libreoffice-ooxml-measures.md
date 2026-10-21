# LibreOffice Writer OOXML measures

- **URL:** https://github.com/LibreOffice/core/blob/f7faa162c274df921fdb431ddab7c74d26f9c1cf/sw/source/writerfilter/ooxml/OOXMLPropertySet.cxx#L445-L514
- **Inspected revision:** `f7faa162c274df921fdb431ddab7c74d26f9c1cf` (2026-03-19)
- **Language:** C++
- **License:** MPL-2.0

## What it does

Writer's DOCX import converts every length attribute to an integer before
the document model sees it:

- `OOXMLValue::createUniversalMeasure(value, perPoint)` reads the leading
  number (`o3tl::toDouble`, 0 when there is none), scales it by the unit
  suffix (`pt`, `cm`, `mm`, `in`, `pc` or `pi`, 1pc = 1pi = 12pt; no
  suffix means the number is already in the target unit) and rounds with
  `std::round`. Twips pass `perPoint = 20`.
- `OOXMLValue::createMeasurementOrPercent` reads a value ending in `%` as
  fiftieths of a percent, truncated (`static_cast<int>(val * 50)`), and
  anything else as twips through `createUniversalMeasure`.
- `sw/source/writerfilter/ooxml/model.xml` maps `CT_TblWidth`'s `w` to
  `ST_MeasurementOrPercent`, so `w:tcW`, `w:tblW`, `w:tblInd`, the cell
  margins and the other table widths all take that path.

## Why relevant

- A Strict-conformance package writes every length as a universal measure
  (`w:tcW w:w="81.85pt"`). docx-rs reads table widths only as integers or
  leading numbers, so office2pdf rewrites each `CT_TblWidth` value with this
  rule before docx-rs reads the package
  (`crates/office2pdf/src/parser/docx_universal_measures.rs`, issue #2245).
