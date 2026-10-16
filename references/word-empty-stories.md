# Explicitly empty Word header/footer stories

- [Header reference semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.headerreference?view=openxml-3.0.1)
- [Footer reference semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footerreference?view=openxml-3.0.1)
- [docxide-pdf section resolution](https://github.com/sverrejb/docxide-pdf/blob/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/src/docx/sections.rs)

A reference selects a part even when that part has no painted content. Omitted
references have inheritance rules; an explicit empty part must not activate them.
The pinned upstream resolves default/first/even slots separately and keeps an
empty paragraph as a block. Its `headers_footers.rs` drops entirely zero-child
parts; our zero-child controls explicitly cover that distinction as well.

Our parser retains an empty referenced part as `Some(HeaderFooter)`. The renderer
chooses an explicit first-page story by presence, then checks its paintable content.
A resolved empty slot also clears the previous Typst section's decoration.
Default header and footer references inherit independently when omitted.

Package-to-render controls cover zero-child parts and empty paragraphs, first-page
suppression with continuation-page defaults, independent default inheritance,
and clearing absent IR slots across sections. These controls do not establish
native visual fidelity for a real document.

Separate limitations remain: inheritance of omitted first-page references,
legacy fallback from first/even stories into missing defaults, and even-page
story selection. Retaining explicit emptiness does not resolve those rules.
