# Prefixed worksheet reproduction

Pre-fix evidence for ISSUE_PENDING. `control.xlsx` uses the default SpreadsheetML namespace; `prefixed.xlsx` binds the same URI to `s` and prefixes the worksheet elements. Expanded element names, text and attributes are identical. All other ZIP parts are unchanged.

Native Excel renders the two sources with zero decoded-pixel difference at 150 DPI. The converter renders the control normally but reports success and produces a blank page for the prefixed source. Its dependency worksheet reader matches literal `e.name()` values such as `row` and `headerFooter`, which do not match `s:row` and `s:headerFooter`.

Model vision findings: inspected the complete 150-DPI native/output comparison and pixel diff. Both have one page; every native body label, value, header and footer is absent from the output. There are no shapes, borders, hairlines, rotated elements, bold/italic/underlined runs or color fills in this synthetic reference. Position, size, font and alignment cannot be compared because no output element survives. Layout/text extraction confirm missing content; no matched shift crops exist. No runtime fix or post-fix result is claimed.
