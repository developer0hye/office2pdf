# Reference Projects

Curated open-source projects relevant to office2pdf's architecture (OOXML → IR → Typst → PDF).

## OOXML Parsing

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [docx4j](https://github.com/plutext/docx4j) | Java | DOCX/PPTX/XLSX parsing + PDF export | [docx4j.md](docx4j.md) |
| [python-docx](https://github.com/python-openxml/python-docx) | Python | DOCX object model | [python-docx.md](python-docx.md) |
| [python-pptx](https://github.com/scanny/python-pptx) | Python | PPTX slide master/layout inheritance | [python-pptx.md](python-pptx.md) |
| [Apache POI](https://github.com/apache/poi) | Java | OOXML row/column/cell object model | [apache-poi.md](apache-poi.md) |
| [openpyxl](https://github.com/ericgazoni/openpyxl) | Python | XLSX reader/writer and indexed color palette | [openpyxl.md](openpyxl.md) |
| [excelize](https://github.com/qax-os/excelize) | Go | XLSX (most complete impl) | [excelize.md](excelize.md) |
| [Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) | C# | Microsoft's official OOXML SDK | [open-xml-sdk.md](open-xml-sdk.md) |
| [docxjs](https://github.com/VolodymyrBaydalka/docxjs) | TS | DOCX → HTML renderer | [docxjs.md](docxjs.md) |
| [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | TS | PPTX table text layout heuristics | [pptxgenjs.md](pptxgenjs.md) |

## Document Conversion & IR Design

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [html5ever / mailparse](https://github.com/servo/html5ever) | Rust | HTML and MIME decoding for DOCX imports | [html-import.md](html-import.md) |
| [docxide-pdf](https://github.com/sverrejb/docxide-pdf) | Rust | Word-based DOCX fixtures and direct PDF rendering | [docxide-pdf.md](docxide-pdf.md) |
| [pandoc](https://github.com/jgm/pandoc) | Haskell | Universal converter (DOCX reader + Typst writer) | [pandoc.md](pandoc.md) |
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | JS | DOCX → semantic HTML | [mammoth.md](mammoth.md) |
| [ONLYOFFICE Document Server](https://github.com/ONLYOFFICE/DocumentServer) | Mixed | DOCX document editing and section layout | [onlyoffice.md](onlyoffice.md) |
| [LibreOffice Writer](https://github.com/LibreOffice/core/blob/022c34f162f94d59499ec48e48edbfd075f583b0/sw/source/core/text/itrform2.cxx#L2445-L2469) | C++ | Exact and at-least line heights | [libreoffice-writer-line-height.md](libreoffice-writer-line-height.md) |
| [LibreOffice Writer OOXML import](https://github.com/LibreOffice/core/blob/f7faa162c274df921fdb431ddab7c74d26f9c1cf/sw/source/writerfilter/ooxml/OOXMLPropertySet.cxx#L445-L514) | C++ | Universal measures and percentages in OOXML lengths | [libreoffice-ooxml-measures.md](libreoffice-ooxml-measures.md) |

## Typst & PDF Generation

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [typst](https://github.com/typst/typst) | Rust | Typesetting engine (our backend) | [typst.md](typst.md) |
| [krilla](https://github.com/LaurenzV/krilla) | Rust | PDF generation (PDF/A + PDF/UA) | [krilla.md](krilla.md) |
| [hayro](https://github.com/LaurenzV/hayro) | Rust | PDF renderer (visual comparison testing) | [hayro.md](hayro.md) |
| [veraPDF](https://github.com/veraPDF/veraPDF-library) | Java | PDF/A and PDF/UA validation | [verapdf.md](verapdf.md) |
| [zune-image](https://github.com/etemesi254/zune-image) | Rust | JPEG decoding used through image and Typst | [zune-image.md](zune-image.md) |

## Math Equations

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [texmath](https://github.com/jgm/texmath) | Haskell | OMML ↔ Typst math conversion | [texmath.md](texmath.md) |
| [tex2typst](https://github.com/qwinsi/tex2typst) | TS | LaTeX ↔ Typst math mappings | [tex2typst.md](tex2typst.md) |

## Font Handling

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [rustybuzz](https://github.com/harfbuzz/rustybuzz) | Rust | Text shaping (used by Typst) | [rustybuzz.md](rustybuzz.md) |
| [allsorts](https://github.com/yeslogic/allsorts) | Rust | Font parsing, shaping, subsetting | [allsorts.md](allsorts.md) |
| [LibreOffice VCL.xcu](https://github.com/LibreOffice/core/blob/a11496e66e08c820df21e7985ba79e6a9bd6d7f0/officecfg/registry/data/org/openoffice/VCL.xcu) | XML/C++ | Missing-font substitution table | [libreoffice-vcl-font-substitution.md](libreoffice-vcl-font-substitution.md) |

## Chart Rendering

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [plotters](https://github.com/plotters-rs/plotters) | Rust | Data visualization (SVG/PNG) | [plotters.md](plotters.md) |
| [charts-rs](https://github.com/vicanso/charts-rs) | Rust | Chart rendering (more chart types) | [charts-rs.md](charts-rs.md) |

## Visual Audit Tools

| Project | Focus | Detail |
|---------|-------|--------|
| [ImageMagick](https://github.com/ImageMagick/ImageMagick) | Raw diff components and thin-rule preservation | [imagemagick-diff-census.md](imagemagick-diff-census.md) |

## OOXML Specification

| Resource | Type |
|----------|------|
| [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/) | Official spec (free download) |
| [officeopenxml.com](http://officeopenxml.com/) | Readable OOXML reference |
| [PowerPoint BulletFormat docs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.bulletformat) | PowerPoint bullet semantics reference ([detail](powerpoint-bulletformat.md)) |

## Native Excel Layout

| Reference | Focus | Detail |
|-----------|-------|--------|
| [Native Excel probes (#1728)](https://github.com/developer0hye/office2pdf/issues/1728) | Plain left header/footer anchors and scaled footer inset | [excel-plain-header-footer.md](excel-plain-header-footer.md) |

## Agent Harness

| Resource | Type |
|----------|------|
| [Meta-Harness (arXiv:2603.28052)](https://arxiv.org/abs/2603.28052) | Harness search from raw prior traces ([detail](meta-harness.md)) |
| [OpenAI Codex](https://github.com/openai/codex) | Headless CLI, JSONL events and structured reports ([detail](codex.md)) |
| [GitHub CLI](https://github.com/cli/cli) | Complete PR enumeration for progress detection ([detail](github-cli-pagination.md)) |
