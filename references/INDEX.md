# Reference Projects

Curated open-source projects relevant to office2pdf's architecture (OOXML → IR → Typst → PDF).

## OOXML Parsing

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [docx4j](https://github.com/plutext/docx4j) | Java | DOCX/PPTX/XLSX parsing + PDF export | [docx4j.md](docx4j.md) |
| [python-docx](https://github.com/python-openxml/python-docx) | Python | DOCX object model | [python-docx.md](python-docx.md) |
| [python-pptx](https://github.com/scanny/python-pptx) | Python | PPTX slide master/layout inheritance | [python-pptx.md](python-pptx.md) |
| [Apache POI](https://github.com/apache/poi) | Java | OOXML row/column/cell object model | [apache-poi.md](apache-poi.md) |
| [excelize](https://github.com/qax-os/excelize) | Go | XLSX (most complete impl) | [excelize.md](excelize.md) |
| [Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) | C# | Microsoft's official OOXML SDK | [open-xml-sdk.md](open-xml-sdk.md) |
| [docxjs](https://github.com/VolodymyrBaydalka/docxjs) | TS | DOCX → HTML renderer | [docxjs.md](docxjs.md) |
| [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | TS | PPTX table text layout heuristics | [pptxgenjs.md](pptxgenjs.md) |

## Document Conversion & IR Design

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [pandoc](https://github.com/jgm/pandoc) | Haskell | Universal converter (DOCX reader + Typst writer) | [pandoc.md](pandoc.md) |
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | JS | DOCX → semantic HTML | [mammoth.md](mammoth.md) |

## Typst & PDF Generation

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [typst](https://github.com/typst/typst) | Rust | Typesetting engine (our backend) | [typst.md](typst.md) |
| [krilla](https://github.com/LaurenzV/krilla) | Rust | PDF generation (PDF/A + PDF/UA) | [krilla.md](krilla.md) |
| [hayro](https://github.com/LaurenzV/hayro) | Rust | PDF renderer (visual comparison testing) | [hayro.md](hayro.md) |
| [veraPDF](https://github.com/veraPDF/veraPDF-library) | Java | PDF/A and PDF/UA validation | [verapdf.md](verapdf.md) |

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

## Chart Rendering

| Project | Lang | Focus | Detail |
|---------|------|-------|--------|
| [plotters](https://github.com/plotters-rs/plotters) | Rust | Data visualization (SVG/PNG) | [plotters.md](plotters.md) |
| [charts-rs](https://github.com/vicanso/charts-rs) | Rust | Chart rendering (more chart types) | [charts-rs.md](charts-rs.md) |

## OOXML Specification

| Resource | Type |
|----------|------|
| [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/) | Official spec (free download) |
| [officeopenxml.com](http://officeopenxml.com/) | Readable OOXML reference |
