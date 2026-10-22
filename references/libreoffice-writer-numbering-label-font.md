# LibreOffice Writer numbering label font

- **URL:** https://github.com/LibreOffice/core/blob/f35aaec880d081205dc42ce36a4161c5acbbdefa/sw/source/core/text/txtfld.cxx#L329-L416
- **Also:** https://github.com/LibreOffice/core/blob/f35aaec880d081205dc42ce36a4161c5acbbdefa/sw/source/core/txtnode/thints.cxx#L1969-L1978
- **Inspected revision:** `f35aaec880d081205dc42ce36a4161c5acbbdefa` (2026-10-22)
- **Language:** C++
- **License:** MPL-2.0

## What it does

`checkApplyParagraphMarkFormatToNumbering` formats a DOCX list label with
the paragraph mark's run properties (`w:pPr/w:rPr`), which its comment notes
Word applies to the numbering although the spec scopes them to the mark:

- The label font starts from the paragraph's own character attributes (its
  style chain), then takes the numbering level's character format.
- The mark's properties go on top, except any the level's format sets: the
  level's `w:rPr` outranks the mark.
- `SwTextNode::IsIgnoredCharFormatForNumbering` drops the mark's underline and
  escapement (super/subscript and raised or lowered position), and shading. Small
  capitals are dropped too; all capitals stay.

## Why relevant

- Word's export of docxide-pdf's scraped `k4_regulations_synthesis_document`
  agrees: labels whose marks state Arial 10pt print in Arial 10pt while a run
  in the item is 11pt, and labels whose marks state nothing print in the
  document-default Calibri 11pt beside Arial 10pt runs (issue #2337).
- office2pdf resolves the label's formatting in `numbering_label_style`
  (`crates/office2pdf/src/parser/docx.rs`) and lays the level's `w:rPr` over
  a number's at render time (`level_label_styles` and `merge_marker_style` in
  `crates/office2pdf/src/render/typst_gen_lists.rs`).
- Word's exports disagree with it for bullets: `alfies_arc_adult_safeguarding_policy`
  and `streamnet_steering_committee_meeting` draw a bullet in its mark's face
  but not its bold, so bullets keep their runs' formatting until a native
  probe settles the rule.
