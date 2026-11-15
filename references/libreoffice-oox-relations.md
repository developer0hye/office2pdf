# LibreOffice oox relationship lookup

- **URL:** https://github.com/LibreOffice/core/blob/022c34f162f94d59499ec48e48edbfd075f583b0/oox/source/core/relations.cxx#L83-L164
- **Inspected revision:** `022c34f162f94d59499ec48e48edbfd075f583b0` (2026-10-18)
- **Language:** C++
- **License:** MPL-2.0

## What it does

`oox::core::Relations` finds a package part by its relationship type.
`getFragmentPathFromFirstTypeFromOfficeDoc` takes a bare type name such as
`styles` and tries it under the Transitional prefix
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/`, then the
Strict prefix `http://purl.oclc.org/ooxml/officeDocument/relationships/`, then
`http://schemas.microsoft.com/office/2006/relationships/`.
`getRelationsFromTypeFromOfficeDoc` collects relationships of either spelling.
Both compare types without regard to ASCII case.

## Why relevant

- docx-rs (`office2pdf-docx-rs` 0.1.0) resolves parts by the Transitional
  types only, so it found none of a Strict package's styles, numbering,
  settings, header, footer, image or hyperlink relationships (issue #2243).
- office2pdf follows LibreOffice's rule: the same type name under either
  prefix names the same part. `parser/docx_strict_relationships.rs` renames
  Strict `Type` values to the Transitional prefix before docx-rs reads the
  package.

## When to consult

- Adding a relationship lookup that must accept Strict packages.
- Deciding whether a Strict type maps one-to-one onto a Transitional one.
  A few differ in more than the prefix (Strict `extendedProperties` against
  Transitional `extended-properties`), and the prefix rule leaves those apart.
