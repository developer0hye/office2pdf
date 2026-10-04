# ONLYOFFICE Document Server

- **URL:** https://github.com/ONLYOFFICE/DocumentServer
- **Focus:** DOCX-compatible document editing and page layout
- **Section API:** https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiSection/Methods/SetType/

## Why relevant

- The official `ApiSection.SetType` example treats `continuous` as a section
  start mode that keeps the next section on the current page. It contrasts this
  with `nextPage`, `oddPage`, and `evenPage` starts.
- This supports keeping section-start semantics separate from unconditional
  page breaks in a document conversion pipeline.

## When to consult

- Comparing DOCX section-start and page-flow behavior.
