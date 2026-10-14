# docx-rs

- Upstream: https://github.com/bokuweb/docx-rs (MIT).
- Our immutable fork pin retains the existing parser repairs.
- Reader footnote references use the existing `RunChild::FootnoteReference`
  model, retaining IDs and their positions among text children.
- Upstream repair: https://github.com/bokuweb/docx-rs/pull/938.
- office2pdf resolves retained IDs against raw `word/footnotes.xml` runs;
  a character style is not evidence of a reference.
- Reference runs can contain surrounding text or occur in tables/hyperlinks.
- Empty normal notes retain their references and later numbering; separator
  records are excluded by type.
- Endnote references still use the existing style/cursor fallback because the
  reader does not retain their reference nodes. Unstyled endnotes remain a gap.
- The reader repair does not parse note bodies. Native Word probes and the
  pinned docxide-pdf corpus provide separate conversion evidence; text recovery
  does not establish marker, spacing, or page-layout fidelity.
