//! Performance validation tests with tiered targets.
//!
//! Three tiers based on document complexity:
//!
//! | Tier   | Pages / Slides / Sheets | P95 Budget |
//! |--------|-------------------------|------------|
//! | Small  | < 10                    | 2s         |
//! | Medium | 10–50                   | 5s         |
//! | Large  | 50–100                  | 8s         |
//!
//! Budgets include CI variability. The first conversion in a process warms
//! the font cache (~1.5s one-time cost via `OnceLock`); subsequent conversions
//! are ~10-100× faster. Large-tier tests are `#[ignore]`d to avoid CI timeouts.
//!
//! All tiers use the P95 budget (generous for CI). Warm-cache measurements
//! are printed to stderr for diagnostic visibility.

use std::io::Cursor;
use std::sync::Once;
use std::time::{Duration, Instant};

use office2pdf::config::{ConvertOptions, Format};

// ── Tiered budgets (P95, including CI variance + cold-cache overhead) ────────

/// Small tier: < 10 pages. Handles cold font cache on first test in process.
const SMALL_BUDGET: Duration = Duration::from_secs(2);

/// Medium tier: 10–50 pages. XLSX tables are expensive in Typst; budget
/// includes CI variance (observed ~3.4s on Ubuntu CI for 10×8×50 XLSX).
const MEDIUM_BUDGET: Duration = Duration::from_secs(5);

/// Large tier: 50–100 pages. Only run locally via `cargo test -- --ignored`.
const LARGE_BUDGET: Duration = Duration::from_secs(8);

/// Warm-cache budget for font-cache validation tests.
const WARM_CACHE_BUDGET: Duration = Duration::from_secs(2);

// ── Font cache warm-up ──────────────────────────────────────────────────────

static WARM_UP: Once = Once::new();

/// Ensure the font cache is populated before measuring. The `OnceLock`-based
/// font cache in `render::pdf` is process-wide, so a single conversion
/// warms it for all subsequent tests in the same binary.
fn ensure_font_cache_warm() {
    WARM_UP.call_once(|| {
        let data = build_small_docx();
        let _ = office2pdf::convert_bytes(&data, Format::Docx, &ConvertOptions::default());
    });
}

// ── Small tier document builders ────────────────────────────────────────────

/// ~5-page DOCX with simple text paragraphs.
fn build_small_docx() -> Vec<u8> {
    let mut doc = docx_rs::Docx::new();
    for i in 0..15 {
        doc = doc.add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(format!(
                "Paragraph {i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
                 Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
            ))),
        );
    }
    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

/// 5-slide PPTX with one text box per slide.
fn build_small_pptx() -> Vec<u8> {
    build_pptx_n_slides(5)
}

/// 3-sheet XLSX with 3 columns × 10 rows each.
fn build_small_xlsx() -> Vec<u8> {
    build_xlsx_sheets(3, 3, 10)
}

// ── Medium tier document builders ───────────────────────────────────────────

/// ~20-page DOCX with text paragraphs, formatting, and tables.
fn build_medium_docx() -> Vec<u8> {
    let mut doc = docx_rs::Docx::new();
    for i in 0..60 {
        let mut run = docx_rs::Run::new().add_text(format!(
            "Paragraph {i}. Sed ut perspiciatis unde omnis iste natus error sit voluptatem \
             accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo \
             inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo."
        ));
        // Add formatting variation every 5th paragraph
        if i % 5 == 0 {
            run = run.bold();
        }
        if i % 7 == 0 {
            run = run.italic();
        }
        doc = doc.add_paragraph(docx_rs::Paragraph::new().add_run(run));

        // Insert a table every 20 paragraphs
        if i > 0 && i % 20 == 0 {
            let mut table = docx_rs::Table::new(vec![]);
            for r in 0..5 {
                let row = docx_rs::TableRow::new(vec![
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("R{r}C1"))),
                    ),
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("R{r}C2"))),
                    ),
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("R{r}C3"))),
                    ),
                ]);
                table = table.add_row(row);
            }
            doc = doc.add_table(table);
        }
    }
    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

/// 20-slide PPTX with text boxes and shapes.
fn build_medium_pptx() -> Vec<u8> {
    build_pptx_n_slides_with_shapes(20)
}

/// 10-sheet XLSX with 8 columns × 50 rows and cell formatting.
fn build_medium_xlsx() -> Vec<u8> {
    build_xlsx_sheets(10, 8, 50)
}

// ── Large tier document builders ────────────────────────────────────────────

/// ~50-page DOCX with paragraphs, tables, and varied formatting.
fn build_large_docx() -> Vec<u8> {
    let mut doc = docx_rs::Docx::new();
    for i in 0..150 {
        let mut run = docx_rs::Run::new().add_text(format!(
            "Paragraph {i}. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit \
             aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi \
             nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet."
        ));
        if i % 3 == 0 {
            run = run.bold();
        }
        if i % 5 == 0 {
            run = run.italic();
        }
        doc = doc.add_paragraph(docx_rs::Paragraph::new().add_run(run));

        // Insert a table every 15 paragraphs
        if i > 0 && i % 15 == 0 {
            let mut table = docx_rs::Table::new(vec![]);
            for r in 0..8 {
                let row = docx_rs::TableRow::new(vec![
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("T{i}R{r}C1"))),
                    ),
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("T{i}R{r}C2"))),
                    ),
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("T{i}R{r}C3"))),
                    ),
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new()
                            .add_run(docx_rs::Run::new().add_text(format!("T{i}R{r}C4"))),
                    ),
                ]);
                table = table.add_row(row);
            }
            doc = doc.add_table(table);
        }
    }
    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

/// 50-slide PPTX with text and shapes.
fn build_large_pptx() -> Vec<u8> {
    build_pptx_n_slides_with_shapes(50)
}

/// 20-sheet XLSX with 10 columns × 100 rows.
fn build_large_xlsx() -> Vec<u8> {
    build_xlsx_sheets(20, 10, 100)
}

// ── Shared PPTX builder ────────────────────────────────────────────────────

/// Build a PPTX with `n` slides, each containing a single text box.
fn build_pptx_n_slides(n: usize) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let opts: zip::write::FileOptions = zip::write::FileOptions::default();

    let mut slide_ct = String::new();
    for i in 1..=n {
        slide_ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
        ));
    }
    writer.start_file("[Content_Types].xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  {slide_ct}
</Types>"#
        )
        .as_bytes(),
    )
    .unwrap();

    writer.start_file("_rels/.rels", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    )
    .unwrap();

    let mut sid = String::new();
    for i in 1..=n {
        sid.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{i}"/>"#, 255 + i));
    }
    writer.start_file("ppt/presentation.xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>{sid}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        )
        .as_bytes(),
    )
    .unwrap();

    let mut srels = String::new();
    for i in 1..=n {
        srels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>"#
        ));
    }
    writer
        .start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {srels}
</Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    for i in 1..=n {
        writer
            .start_file(format!("ppt/slides/slide{i}.xml"), opts)
            .unwrap();
        std::io::Write::write_all(
            &mut writer,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TextBox {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="5943600"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Slide {i}: Lorem ipsum dolor sit amet.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#
            )
            .as_bytes(),
        )
        .unwrap();
    }

    writer.finish().unwrap().into_inner()
}

/// Build a PPTX with `n` slides, each containing a text box and a shape.
fn build_pptx_n_slides_with_shapes(n: usize) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let opts: zip::write::FileOptions = zip::write::FileOptions::default();

    let mut slide_ct = String::new();
    for i in 1..=n {
        slide_ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
        ));
    }
    writer.start_file("[Content_Types].xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  {slide_ct}
</Types>"#
        )
        .as_bytes(),
    )
    .unwrap();

    writer.start_file("_rels/.rels", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    )
    .unwrap();

    let mut sid = String::new();
    for i in 1..=n {
        sid.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{i}"/>"#, 255 + i));
    }
    writer.start_file("ppt/presentation.xml", opts).unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>{sid}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        )
        .as_bytes(),
    )
    .unwrap();

    let mut srels = String::new();
    for i in 1..=n {
        srels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>"#
        ));
    }
    writer
        .start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    std::io::Write::write_all(
        &mut writer,
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {srels}
</Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    for i in 1..=n {
        writer
            .start_file(format!("ppt/slides/slide{i}.xml"), opts)
            .unwrap();
        std::io::Write::write_all(
            &mut writer,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TextBox {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="2971800"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:rPr lang="en-US" b="1" sz="2400"/><a:t>Slide {i}: Performance test content</a:t></a:r></a:p>
          <a:p><a:r><a:t>Additional text for medium/large complexity testing.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Shape {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="3886200"/><a:ext cx="4114800" cy="2514600"/></a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#
            )
            .as_bytes(),
        )
        .unwrap();
    }

    writer.finish().unwrap().into_inner()
}

// ── Shared XLSX builder ─────────────────────────────────────────────────────

/// Build an XLSX with `sheets` sheets, each having `cols` × `rows` cells.
fn build_xlsx_sheets(sheets: usize, cols: u32, rows: u32) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name("Sheet1");
    fill_xlsx_sheet(sheet, 1, cols, rows);

    for s in 2..=sheets {
        let name = format!("Sheet{s}");
        book.new_sheet(&name).unwrap();
        let sheet = book.get_sheet_by_name_mut(&name).unwrap();
        fill_xlsx_sheet(sheet, s, cols, rows);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// Fill a worksheet with data. Column letters wrap (A-Z then AA, AB, ...).
fn fill_xlsx_sheet(
    sheet: &mut umya_spreadsheet::Worksheet,
    sheet_num: usize,
    cols: u32,
    rows: u32,
) {
    for row in 1..=rows {
        for col in 1..=cols {
            let coord = col_row_to_coord(col, row);
            sheet
                .get_cell_mut(coord.as_str())
                .set_value(format!("S{sheet_num}R{row}C{col}"));
        }
    }
}

/// Convert 1-indexed (col, row) to Excel coordinate string (e.g., "AA12").
fn col_row_to_coord(col: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        letters.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    format!("{letters}{row}")
}

// ── Helper: timed conversion with metrics output ────────────────────────────

/// Convert data and print per-stage metrics to stderr. Returns elapsed time.
fn timed_convert(data: &[u8], format: Format, label: &str) -> Duration {
    let opts = ConvertOptions::default();
    let start = Instant::now();
    let result = office2pdf::convert_bytes(data, format, &opts).unwrap();
    let elapsed = start.elapsed();
    if let Some(m) = result.metrics.as_ref() {
        eprintln!(
            "{label}: parse={:?} codegen={:?} compile={:?} total={:?} pages={}",
            m.parse_duration,
            m.codegen_duration,
            m.compile_duration,
            m.total_duration,
            m.page_count
        );
    }
    elapsed
}

// ═══════════════════════════════════════════════════════════════════════════
// Small tier tests (< 10 pages/slides/sheets) — P95 budget: 2s
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn perf_small_docx() {
    let data = build_small_docx();
    let elapsed = timed_convert(&data, Format::Docx, "Small DOCX (~5 pages)");
    assert!(
        elapsed < SMALL_BUDGET,
        "Small DOCX conversion took {elapsed:?}, exceeds {SMALL_BUDGET:?} budget"
    );
}

#[test]
fn perf_small_pptx() {
    let data = build_small_pptx();
    let elapsed = timed_convert(&data, Format::Pptx, "Small PPTX (5 slides)");
    assert!(
        elapsed < SMALL_BUDGET,
        "Small PPTX conversion took {elapsed:?}, exceeds {SMALL_BUDGET:?} budget"
    );
}

#[test]
fn perf_small_xlsx() {
    let data = build_small_xlsx();
    let elapsed = timed_convert(&data, Format::Xlsx, "Small XLSX (3 sheets)");
    assert!(
        elapsed < SMALL_BUDGET,
        "Small XLSX conversion took {elapsed:?}, exceeds {SMALL_BUDGET:?} budget"
    );
}

// ═══════════════════════════════════════════════════════════════════════════
// Medium tier tests (10–50 pages/slides/sheets) — P95 budget: 3s
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn perf_medium_docx() {
    let data = build_medium_docx();
    let elapsed = timed_convert(&data, Format::Docx, "Medium DOCX (~20 pages)");
    assert!(
        elapsed < MEDIUM_BUDGET,
        "Medium DOCX conversion took {elapsed:?}, exceeds {MEDIUM_BUDGET:?} budget"
    );
}

#[test]
fn perf_medium_pptx() {
    let data = build_medium_pptx();
    let elapsed = timed_convert(&data, Format::Pptx, "Medium PPTX (20 slides)");
    assert!(
        elapsed < MEDIUM_BUDGET,
        "Medium PPTX conversion took {elapsed:?}, exceeds {MEDIUM_BUDGET:?} budget"
    );
}

#[test]
fn perf_medium_xlsx() {
    let data = build_medium_xlsx();
    let elapsed = timed_convert(&data, Format::Xlsx, "Medium XLSX (10 sheets)");
    assert!(
        elapsed < MEDIUM_BUDGET,
        "Medium XLSX conversion took {elapsed:?}, exceeds {MEDIUM_BUDGET:?} budget"
    );
}

// ═══════════════════════════════════════════════════════════════════════════
// Large tier tests (50–100 pages/slides/sheets) — P95 budget: 8s
// These are #[ignore]d to avoid CI timeouts on GitHub Actions runners.
// Run locally with: cargo test -- --ignored
// ═══════════════════════════════════════════════════════════════════════════

#[test]
#[ignore]
fn perf_large_docx() {
    ensure_font_cache_warm();
    let data = build_large_docx();
    let elapsed = timed_convert(&data, Format::Docx, "Large DOCX (~50 pages)");
    assert!(
        elapsed < LARGE_BUDGET,
        "Large DOCX conversion took {elapsed:?}, exceeds {LARGE_BUDGET:?} budget"
    );
}

#[test]
#[ignore]
fn perf_large_pptx() {
    ensure_font_cache_warm();
    let data = build_large_pptx();
    let elapsed = timed_convert(&data, Format::Pptx, "Large PPTX (50 slides)");
    assert!(
        elapsed < LARGE_BUDGET,
        "Large PPTX conversion took {elapsed:?}, exceeds {LARGE_BUDGET:?} budget"
    );
}

#[test]
#[ignore]
fn perf_large_xlsx() {
    ensure_font_cache_warm();
    let data = build_large_xlsx();
    let elapsed = timed_convert(&data, Format::Xlsx, "Large XLSX (20 sheets)");
    assert!(
        elapsed < LARGE_BUDGET,
        "Large XLSX conversion took {elapsed:?}, exceeds {LARGE_BUDGET:?} budget"
    );
}

// ═══════════════════════════════════════════════════════════════════════════
// Font cache validation tests
// ═══════════════════════════════════════════════════════════════════════════

/// Verify per-stage metrics are populated and that the compile stage
/// benefits from font caching on repeated conversions.
#[test]
fn perf_font_cache_second_conversion_faster() {
    let data = build_small_docx();
    let opts = ConvertOptions::default();

    // First conversion: cold font cache (or warm if another test ran first)
    let result1 = office2pdf::convert_bytes(&data, Format::Docx, &opts).unwrap();
    let m1 = result1
        .metrics
        .as_ref()
        .expect("metrics should be populated");

    // Second conversion: guaranteed warm font cache
    let result2 = office2pdf::convert_bytes(&data, Format::Docx, &opts).unwrap();
    let m2 = result2
        .metrics
        .as_ref()
        .expect("metrics should be populated");

    eprintln!(
        "First conversion:  parse={:?} codegen={:?} compile={:?} total={:?}",
        m1.parse_duration, m1.codegen_duration, m1.compile_duration, m1.total_duration
    );
    eprintln!(
        "Second conversion: parse={:?} codegen={:?} compile={:?} total={:?}",
        m2.parse_duration, m2.codegen_duration, m2.compile_duration, m2.total_duration
    );

    // Second conversion total should be under the warm-cache budget
    assert!(
        m2.total_duration < WARM_CACHE_BUDGET,
        "Second DOCX conversion took {:?}, expected under {WARM_CACHE_BUDGET:?} with warm font cache",
        m2.total_duration
    );
}

/// After font cache is warm, conversions across all formats should be fast.
#[test]
fn perf_cross_format_cached_conversion() {
    let opts = ConvertOptions::default();

    // Warm up the font cache
    let docx_data = build_small_docx();
    let _ = office2pdf::convert_bytes(&docx_data, Format::Docx, &opts).unwrap();

    // PPTX with warm cache
    let pptx_data = build_small_pptx();
    let pptx_elapsed = timed_convert(&pptx_data, Format::Pptx, "PPTX (warm cache)");
    assert!(
        pptx_elapsed < WARM_CACHE_BUDGET,
        "PPTX conversion with warm cache took {pptx_elapsed:?}, expected under {WARM_CACHE_BUDGET:?}"
    );

    // XLSX with warm cache
    let xlsx_data = build_small_xlsx();
    let xlsx_elapsed = timed_convert(&xlsx_data, Format::Xlsx, "XLSX (warm cache)");
    assert!(
        xlsx_elapsed < WARM_CACHE_BUDGET,
        "XLSX conversion with warm cache took {xlsx_elapsed:?}, expected under {WARM_CACHE_BUDGET:?}"
    );
}

// ═══════════════════════════════════════════════════════════════════════════
// Tier builder unit tests — verify synthetic documents are well-formed
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn tier_builders_produce_valid_documents() {
    let opts = ConvertOptions::default();

    // Warm font cache once
    ensure_font_cache_warm();

    // Small tier: all formats should convert without error
    let small_docx = build_small_docx();
    let r = office2pdf::convert_bytes(&small_docx, Format::Docx, &opts).unwrap();
    assert!(!r.pdf.is_empty(), "Small DOCX should produce non-empty PDF");

    let small_pptx = build_small_pptx();
    let r = office2pdf::convert_bytes(&small_pptx, Format::Pptx, &opts).unwrap();
    assert!(!r.pdf.is_empty(), "Small PPTX should produce non-empty PDF");

    let small_xlsx = build_small_xlsx();
    let r = office2pdf::convert_bytes(&small_xlsx, Format::Xlsx, &opts).unwrap();
    assert!(!r.pdf.is_empty(), "Small XLSX should produce non-empty PDF");

    // Medium tier
    let medium_docx = build_medium_docx();
    let r = office2pdf::convert_bytes(&medium_docx, Format::Docx, &opts).unwrap();
    assert!(
        !r.pdf.is_empty(),
        "Medium DOCX should produce non-empty PDF"
    );

    let medium_pptx = build_medium_pptx();
    let r = office2pdf::convert_bytes(&medium_pptx, Format::Pptx, &opts).unwrap();
    assert!(
        !r.pdf.is_empty(),
        "Medium PPTX should produce non-empty PDF"
    );

    let medium_xlsx = build_medium_xlsx();
    let r = office2pdf::convert_bytes(&medium_xlsx, Format::Xlsx, &opts).unwrap();
    assert!(
        !r.pdf.is_empty(),
        "Medium XLSX should produce non-empty PDF"
    );
}

#[test]
fn col_row_to_coord_basic() {
    assert_eq!(col_row_to_coord(1, 1), "A1");
    assert_eq!(col_row_to_coord(26, 1), "Z1");
    assert_eq!(col_row_to_coord(27, 5), "AA5");
    assert_eq!(col_row_to_coord(28, 10), "AB10");
}
