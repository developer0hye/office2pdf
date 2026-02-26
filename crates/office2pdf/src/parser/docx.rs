use std::collections::HashMap;

use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, Block, BorderSide, CellBorder, Color, Document, FlowPage, ImageData, ImageFormat,
    LineSpacing, Margins, Metadata, Page, PageSize, Paragraph, ParagraphStyle, Run, StyleSheet,
    Table, TableCell, TableRow, TextStyle,
};
use crate::parser::Parser;

pub struct DocxParser;

/// Map from relationship ID → PNG image bytes.
type ImageMap = HashMap<String, Vec<u8>>;

/// Build a lookup map from the DOCX's embedded images.
/// docx-rs converts all images to PNG; we use the PNG bytes.
fn build_image_map(docx: &docx_rs::Docx) -> ImageMap {
    docx.images
        .iter()
        .map(|(id, _path, _image, png)| (id.clone(), png.0.clone()))
        .collect()
}

/// Convert EMU (English Metric Units) to points.
/// 1 inch = 914400 EMU, 1 inch = 72 points, so 1 pt = 12700 EMU.
fn emu_to_pt(emu: u32) -> f64 {
    emu as f64 / 12700.0
}

impl Parser for DocxParser {
    fn parse(&self, data: &[u8]) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX: {e}")))?;

        let (size, margins) = extract_page_setup(&docx.document.section_property);
        let images = build_image_map(&docx);
        let mut warnings = Vec::new();

        let mut content: Vec<Block> = Vec::new();
        for (idx, child) in docx.document.children.iter().enumerate() {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut blocks = Vec::new();
                match child {
                    docx_rs::DocumentChild::Paragraph(para) => {
                        convert_paragraph_blocks(para, &mut blocks, &images);
                    }
                    docx_rs::DocumentChild::Table(table) => {
                        blocks.push(Block::Table(convert_table(table, &images)));
                    }
                    _ => {}
                }
                blocks
            }));

            match result {
                Ok(blocks) => content.extend(blocks),
                Err(_) => {
                    warnings.push(ConvertWarning {
                        element: format!("Document element at index {idx}"),
                        reason: "element processing panicked; skipped".to_string(),
                    });
                }
            }
        }

        Ok((
            Document {
                metadata: Metadata::default(),
                pages: vec![Page::Flow(FlowPage {
                    size,
                    margins,
                    content,
                })],
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

/// Extract page size and margins from DOCX section properties.
fn extract_page_setup(section_prop: &docx_rs::SectionProperty) -> (PageSize, Margins) {
    let size = extract_page_size(&section_prop.page_size);
    let margins = extract_margins(&section_prop.page_margin);
    (size, margins)
}

/// Extract page size from docx-rs PageSize (which has private fields).
/// Uses serde serialization to access the private `w` and `h` fields.
/// Values in DOCX are in twips (1/20 of a point).
fn extract_page_size(page_size: &docx_rs::PageSize) -> PageSize {
    if let Ok(json) = serde_json::to_value(page_size) {
        let w = json.get("w").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let h = json.get("h").and_then(|v| v.as_f64()).unwrap_or(0.0);
        if w > 0.0 && h > 0.0 {
            return PageSize {
                width: w / 20.0,  // twips to points
                height: h / 20.0, // twips to points
            };
        }
    }
    PageSize::default()
}

/// Extract margins from docx-rs PageMargin.
/// PageMargin fields are public i32 values in twips.
fn extract_margins(page_margin: &docx_rs::PageMargin) -> Margins {
    Margins {
        top: page_margin.top as f64 / 20.0,
        bottom: page_margin.bottom as f64 / 20.0,
        left: page_margin.left as f64 / 20.0,
        right: page_margin.right as f64 / 20.0,
    }
}

/// Convert a docx-rs Paragraph to IR blocks, handling page breaks and inline images.
/// If the paragraph has `page_break_before`, a `Block::PageBreak` is emitted first.
/// Inline images within runs are extracted as separate `Block::Image` elements.
fn convert_paragraph_blocks(para: &docx_rs::Paragraph, out: &mut Vec<Block>, images: &ImageMap) {
    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    // Collect text runs and detect inline images
    let mut runs: Vec<Run> = Vec::new();
    let mut inline_images: Vec<Block> = Vec::new();

    for child in &para.children {
        if let docx_rs::ParagraphChild::Run(run) = child {
            // Check for inline images in this run
            for run_child in &run.children {
                if let docx_rs::RunChild::Drawing(drawing) = run_child
                    && let Some(img_block) = extract_drawing_image(drawing, images)
                {
                    inline_images.push(img_block);
                }
            }

            // Extract text from the run
            let text = extract_run_text(run);
            if !text.is_empty() {
                runs.push(Run {
                    text,
                    style: extract_run_style(&run.run_property),
                });
            }
        }
    }

    // Emit image blocks before the paragraph (inline images are block-level in our IR)
    out.extend(inline_images);

    out.push(Block::Paragraph(Paragraph {
        style: extract_paragraph_style(&para.property),
        runs,
    }));
}

/// Extract an image from a Drawing element if it contains a Pic with matching image data.
fn extract_drawing_image(drawing: &docx_rs::Drawing, images: &ImageMap) -> Option<Block> {
    let pic = match &drawing.data {
        Some(docx_rs::DrawingData::Pic(pic)) => pic,
        _ => return None,
    };

    // Look up image data by relationship ID
    let data = images.get(&pic.id)?;

    let (w_emu, h_emu) = pic.size;
    let width = if w_emu > 0 {
        Some(emu_to_pt(w_emu))
    } else {
        None
    };
    let height = if h_emu > 0 {
        Some(emu_to_pt(h_emu))
    } else {
        None
    };

    Some(Block::Image(ImageData {
        data: data.clone(),
        format: ImageFormat::Png, // docx-rs converts all images to PNG
        width,
        height,
    }))
}

/// Extract paragraph-level formatting from docx-rs ParagraphProperty.
fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
    let alignment = prop.alignment.as_ref().and_then(|j| match j.val.as_str() {
        "center" => Some(Alignment::Center),
        "right" | "end" => Some(Alignment::Right),
        "left" | "start" => Some(Alignment::Left),
        "both" | "justified" => Some(Alignment::Justify),
        _ => None,
    });

    let (indent_left, indent_right, indent_first_line) = extract_indent(&prop.indent);

    let (line_spacing, space_before, space_after) = extract_line_spacing(&prop.line_spacing);

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        space_before,
        space_after,
    }
}

/// Extract indentation values from docx-rs Indent.
/// All values in docx-rs are in twips; convert to points (÷20).
fn extract_indent(indent: &Option<docx_rs::Indent>) -> (Option<f64>, Option<f64>, Option<f64>) {
    let Some(indent) = indent else {
        return (None, None, None);
    };

    let left = indent.start.map(|v| v as f64 / 20.0);
    let right = indent.end.map(|v| v as f64 / 20.0);
    let first_line = indent.special_indent.map(|si| match si {
        docx_rs::SpecialIndentType::FirstLine(v) => v as f64 / 20.0,
        docx_rs::SpecialIndentType::Hanging(v) => -(v as f64 / 20.0),
    });

    (left, right, first_line)
}

/// Extract line spacing from docx-rs LineSpacing (private fields, use serde).
///
/// OOXML line spacing:
/// - Auto: `line` is in 240ths of a line (240=single, 360=1.5×, 480=double)
/// - Exact/AtLeast: `line` is in twips (÷20 → points)
/// - `before`/`after`: twips (÷20 → points)
fn extract_line_spacing(
    spacing: &Option<docx_rs::LineSpacing>,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
    let Some(spacing) = spacing else {
        return (None, None, None);
    };

    let json = match serde_json::to_value(spacing) {
        Ok(j) => j,
        Err(_) => return (None, None, None),
    };

    let space_before = json
        .get("before")
        .and_then(|v| v.as_f64())
        .map(|v| v / 20.0);
    let space_after = json.get("after").and_then(|v| v.as_f64()).map(|v| v / 20.0);

    let line_spacing = json.get("line").and_then(|line_val| {
        let line = line_val.as_f64()?;
        let rule = json.get("lineRule").and_then(|v| v.as_str());
        match rule {
            Some("exact") | Some("atLeast") => Some(LineSpacing::Exact(line / 20.0)),
            _ => {
                // Auto: 240ths of a line
                Some(LineSpacing::Proportional(line / 240.0))
            }
        }
    });

    (line_spacing, space_before, space_after)
}

/// Convert a docx-rs Table to an IR Table.
///
/// Handles:
/// - Column widths from the table grid (twips → points)
/// - Cell content (paragraphs with formatted text)
/// - Horizontal merging via gridSpan (colspan)
/// - Vertical merging via vMerge restart/continue (rowspan)
/// - Cell background color via shading
/// - Cell borders
fn convert_table(table: &docx_rs::Table, images: &ImageMap) -> Table {
    let column_widths: Vec<f64> = table.grid.iter().map(|&w| w as f64 / 20.0).collect();

    // First pass: extract raw rows with vmerge info for rowspan calculation
    let raw_rows = extract_raw_rows(table, images);

    // Second pass: resolve vertical merges into rowspan values and build IR rows
    let rows = resolve_vmerge_and_build_rows(&raw_rows);

    Table {
        rows,
        column_widths,
    }
}

/// Intermediate cell representation for vmerge resolution.
struct RawCell {
    content: Vec<Block>,
    col_span: u32,
    col_index: usize,
    vmerge: Option<String>, // "restart", "continue", or None
    border: Option<CellBorder>,
    background: Option<Color>,
}

/// Extract raw rows from a docx-rs Table, tracking column indices and vmerge state.
fn extract_raw_rows(table: &docx_rs::Table, images: &ImageMap) -> Vec<Vec<RawCell>> {
    let mut raw_rows = Vec::new();

    for table_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = table_child;
        let mut cells = Vec::new();
        let mut col_index: usize = 0;

        for row_child in &row.cells {
            let docx_rs::TableRowChild::TableCell(cell) = row_child;

            let prop_json = serde_json::to_value(&cell.property).ok();
            let grid_span = prop_json
                .as_ref()
                .and_then(|j| j.get("gridSpan"))
                .and_then(|v| v.as_u64())
                .unwrap_or(1) as u32;

            let vmerge = prop_json
                .as_ref()
                .and_then(|j| j.get("verticalMerge"))
                .and_then(|v| v.as_str())
                .map(String::from);

            let content = extract_cell_content(cell, images);
            let border = prop_json
                .as_ref()
                .and_then(|j| j.get("borders"))
                .and_then(extract_cell_borders);
            let background = prop_json
                .as_ref()
                .and_then(|j| j.get("shading"))
                .and_then(extract_cell_shading);

            cells.push(RawCell {
                content,
                col_span: grid_span,
                col_index,
                vmerge,
                border,
                background,
            });

            col_index += grid_span as usize;
        }

        raw_rows.push(cells);
    }

    raw_rows
}

/// Resolve vertical merges: compute rowspan for "restart" cells and skip "continue" cells.
fn resolve_vmerge_and_build_rows(raw_rows: &[Vec<RawCell>]) -> Vec<TableRow> {
    let mut rows = Vec::new();

    for (row_idx, raw_row) in raw_rows.iter().enumerate() {
        let mut cells = Vec::new();

        for raw_cell in raw_row {
            match raw_cell.vmerge.as_deref() {
                Some("continue") => {
                    // Skip continue cells — they are part of a vertical merge above
                    continue;
                }
                Some("restart") => {
                    // Count how many consecutive "continue" cells follow in the same column
                    let row_span = count_vmerge_span(raw_rows, row_idx, raw_cell.col_index);
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                    });
                }
                _ => {
                    // Normal cell (no vmerge)
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span: 1,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                    });
                }
            }
        }

        rows.push(TableRow {
            cells,
            height: None,
        });
    }

    rows
}

/// Count the vertical merge span starting from a "restart" cell.
/// Looks at rows below `start_row` for "continue" cells at the same column index.
fn count_vmerge_span(raw_rows: &[Vec<RawCell>], start_row: usize, col_index: usize) -> u32 {
    let mut span = 1u32;
    for row in raw_rows.iter().skip(start_row + 1) {
        let has_continue = row
            .iter()
            .any(|c| c.col_index == col_index && c.vmerge.as_deref() == Some("continue"));
        if has_continue {
            span += 1;
        } else {
            break;
        }
    }
    span
}

/// Extract cell content (paragraphs) from a docx-rs TableCell.
fn extract_cell_content(cell: &docx_rs::TableCell, images: &ImageMap) -> Vec<Block> {
    let mut blocks = Vec::new();
    for content in &cell.children {
        match content {
            docx_rs::TableCellContent::Paragraph(para) => {
                convert_paragraph_blocks(para, &mut blocks, images);
            }
            docx_rs::TableCellContent::Table(nested_table) => {
                blocks.push(Block::Table(convert_table(nested_table, images)));
            }
            _ => {}
        }
    }
    blocks
}

/// Extract cell borders from the serialized "borders" JSON object.
/// Border size in docx-rs is in eighths of a point; convert to points (÷8).
fn extract_cell_borders(borders_json: &serde_json::Value) -> Option<CellBorder> {
    if borders_json.is_null() {
        return None;
    }

    let extract_side = |key: &str| -> Option<BorderSide> {
        let side = borders_json.get(key)?;
        if side.is_null() {
            return None;
        }
        let border_type = side
            .get("borderType")
            .and_then(|v| v.as_str())
            .unwrap_or("none");
        if border_type == "none" || border_type == "nil" {
            return None;
        }
        let size = side.get("size").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let color_hex = side
            .get("color")
            .and_then(|v| v.as_str())
            .unwrap_or("000000");
        let color = parse_hex_color(color_hex).unwrap_or(Color::black());
        Some(BorderSide {
            width: size / 8.0, // eighths of a point → points
            color,
        })
    };

    let top = extract_side("top");
    let bottom = extract_side("bottom");
    let left = extract_side("left");
    let right = extract_side("right");

    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }

    Some(CellBorder {
        top,
        bottom,
        left,
        right,
    })
}

/// Extract background color from the serialized "shading" JSON object.
fn extract_cell_shading(shading_json: &serde_json::Value) -> Option<Color> {
    if shading_json.is_null() {
        return None;
    }
    let fill = shading_json.get("fill").and_then(|v| v.as_str())?;
    // Skip auto/transparent fills
    if fill == "auto" || fill == "FFFFFF" || fill.is_empty() {
        return None;
    }
    parse_hex_color(fill)
}

/// Extract inline text style from a docx-rs RunProperty.
///
/// docx-rs types with private fields serialize directly as their inner value
/// (e.g. Bold → `true`, Sz → `24`, Color → `"FF0000"`), not as `{"val": ...}`.
/// Strike has a public `val` field and can be accessed directly.
fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    TextStyle {
        bold: extract_bool_prop(&rp.bold),
        italic: extract_bool_prop(&rp.italic),
        underline: rp.underline.as_ref().and_then(|u| {
            let json = serde_json::to_value(u).ok()?;
            let val = json.as_str()?;
            if val == "none" { None } else { Some(true) }
        }),
        strikethrough: rp.strike.as_ref().map(|s| s.val),
        font_size: rp.sz.as_ref().and_then(|sz| {
            let json = serde_json::to_value(sz).ok()?;
            let half_points = json.as_f64()?;
            Some(half_points / 2.0)
        }),
        color: rp.color.as_ref().and_then(|c| {
            let json = serde_json::to_value(c).ok()?;
            let hex = json.as_str()?;
            parse_hex_color(hex)
        }),
        font_family: rp.fonts.as_ref().and_then(|f| {
            let json = serde_json::to_value(f).ok()?;
            // Prefer ascii font name, fall back to hi_ansi, east_asia, cs
            json.get("ascii")
                .or_else(|| json.get("hi_ansi"))
                .or_else(|| json.get("east_asia"))
                .or_else(|| json.get("cs"))
                .and_then(|v| v.as_str())
                .map(String::from)
        }),
    }
}

/// Extract a boolean property (Bold, Italic) via serde. Returns None if absent.
/// docx-rs serializes Bold/Italic directly as a boolean (e.g. `true`).
fn extract_bool_prop<T: serde::Serialize>(prop: &Option<T>) -> Option<bool> {
    prop.as_ref().and_then(|p| {
        let json = serde_json::to_value(p).ok()?;
        json.as_bool()
    })
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Extract text content from a docx-rs Run.
fn extract_run_text(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(_) => text.push('\n'),
            _ => {}
        }
    }
    text
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;
    use std::io::Cursor;

    /// Helper: build a minimal DOCX as bytes using docx-rs builder.
    fn build_docx_bytes(paragraphs: Vec<docx_rs::Paragraph>) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new();
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with custom page size and margins.
    fn build_docx_bytes_with_page_setup(
        paragraphs: Vec<docx_rs::Paragraph>,
        width_twips: u32,
        height_twips: u32,
        margin_top: i32,
        margin_bottom: i32,
        margin_left: i32,
        margin_right: i32,
    ) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new()
            .page_size(width_twips, height_twips)
            .page_margin(
                docx_rs::PageMargin::new()
                    .top(margin_top)
                    .bottom(margin_bottom)
                    .left(margin_left)
                    .right(margin_right),
            );
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    // ----- Basic parsing tests -----

    #[test]
    fn test_parse_empty_docx() {
        let data = build_docx_bytes(vec![]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        // An empty DOCX should produce a document with one FlowPage and no content blocks
        assert_eq!(doc.pages.len(), 1);
        match &doc.pages[0] {
            Page::Flow(page) => {
                assert!(page.content.is_empty());
            }
            _ => panic!("Expected FlowPage"),
        }
    }

    #[test]
    fn test_parse_single_paragraph() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello, world!")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert_eq!(para.runs.len(), 1);
                assert_eq!(para.runs[0].text, "Hello, world!");
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    #[test]
    fn test_parse_multiple_paragraphs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Third paragraph")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 3);

        let texts: Vec<&str> = page
            .content
            .iter()
            .map(|b| match b {
                Block::Paragraph(p) => p.runs[0].text.as_str(),
                _ => panic!("Expected Paragraph"),
            })
            .collect();
        assert_eq!(
            texts,
            vec!["First paragraph", "Second paragraph", "Third paragraph"]
        );
    }

    #[test]
    fn test_parse_paragraph_with_multiple_runs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hello, "))
                .add_run(docx_rs::Run::new().add_text("beautiful "))
                .add_run(docx_rs::Run::new().add_text("world!")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].text, "Hello, ");
        assert_eq!(para.runs[1].text, "beautiful ");
        assert_eq!(para.runs[2].text, "world!");
    }

    #[test]
    fn test_parse_empty_paragraph() {
        let data = build_docx_bytes(vec![docx_rs::Paragraph::new()]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // An empty paragraph should still be present (it may have no runs)
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert!(para.runs.is_empty());
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    // ----- Page setup tests -----

    #[test]
    fn test_default_page_size_is_used() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // docx-rs default: 11906 x 16838 twips (A4)
        // = 595.3 x 841.9 pt
        assert!(page.size.width > 0.0);
        assert!(page.size.height > 0.0);
    }

    #[test]
    fn test_custom_page_size_extracted() {
        // A5 page: 148mm x 210mm
        // In twips: 8392 x 11907 (1 pt = 20 twips)
        let width_twips: u32 = 8392;
        let height_twips: u32 = 11907;
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            width_twips,
            height_twips,
            1440,
            1440,
            1440,
            1440,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_width = width_twips as f64 / 20.0;
        let expected_height = height_twips as f64 / 20.0;
        assert!(
            (page.size.width - expected_width).abs() < 1.0,
            "Expected width ~{expected_width}, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - expected_height).abs() < 1.0,
            "Expected height ~{expected_height}, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_custom_margins_extracted() {
        // Margins: 0.5 inch = 720 twips = 36pt
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            12240,
            15840,
            720,
            720,
            720,
            720,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_margin = 720.0 / 20.0; // 36pt
        assert!(
            (page.margins.top - expected_margin).abs() < 1.0,
            "Expected top margin ~{expected_margin}, got {}",
            page.margins.top
        );
        assert!((page.margins.bottom - expected_margin).abs() < 1.0);
        assert!((page.margins.left - expected_margin).abs() < 1.0);
        assert!((page.margins.right - expected_margin).abs() < 1.0);
    }

    // ----- Error handling tests -----

    #[test]
    fn test_parse_invalid_data_returns_error() {
        let parser = DocxParser;
        let result = parser.parse(b"not a valid docx file");
        assert!(result.is_err());
        match result.unwrap_err() {
            ConvertError::Parse(_) => {}
            other => panic!("Expected Parse error, got: {other:?}"),
        }
    }

    // ----- Text style defaults -----

    #[test]
    fn test_parsed_runs_have_default_text_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        let run = &para.runs[0];
        // Plain text should have default style (all None)
        assert!(run.style.bold.is_none() || run.style.bold == Some(false));
        assert!(run.style.italic.is_none() || run.style.italic == Some(false));
        assert!(run.style.underline.is_none() || run.style.underline == Some(false));
    }

    #[test]
    fn test_parsed_paragraphs_have_default_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        // Default paragraph style should have no explicit alignment
        assert!(para.style.alignment.is_none());
    }

    // ----- Inline formatting tests (US-004) -----

    /// Helper: extract the first run from the first paragraph of a parsed document.
    fn first_run(doc: &Document) -> &Run {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        &para.runs[0]
    }

    #[test]
    fn test_bold_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bold text").bold()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_italic_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Italic text").italic()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.italic, Some(true));
    }

    #[test]
    fn test_underline_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Underlined text")
                    .underline("single"),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.underline, Some(true));
    }

    #[test]
    fn test_strikethrough_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Struck text").strike()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.strikethrough, Some(true));
    }

    #[test]
    fn test_font_size_extracted() {
        // docx-rs size is in half-points: 24 half-points = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Sized text").size(24)),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_size, Some(12.0));
    }

    #[test]
    fn test_font_color_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Red text").color("FF0000")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_font_family_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Arial text")
                    .fonts(docx_rs::RunFonts::new().ascii("Arial")),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_family, Some("Arial".to_string()));
    }

    #[test]
    fn test_combined_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Styled text")
                    .bold()
                    .italic()
                    .underline("single")
                    .strike()
                    .size(28) // 14pt
                    .color("0000FF")
                    .fonts(docx_rs::RunFonts::new().ascii("Courier")),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.italic, Some(true));
        assert_eq!(run.style.underline, Some(true));
        assert_eq!(run.style.strikethrough, Some(true));
        assert_eq!(run.style.font_size, Some(14.0));
        assert_eq!(run.style.color, Some(Color::new(0, 0, 255)));
        assert_eq!(run.style.font_family, Some("Courier".to_string()));
    }

    #[test]
    fn test_plain_text_has_no_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert!(run.style.bold.is_none());
        assert!(run.style.italic.is_none());
        assert!(run.style.underline.is_none());
        assert!(run.style.strikethrough.is_none());
        assert!(run.style.font_size.is_none());
        assert!(run.style.color.is_none());
        assert!(run.style.font_family.is_none());
    }

    // ----- Paragraph formatting tests (US-005) -----

    /// Helper: extract the first paragraph from a parsed document.
    fn first_paragraph(doc: &Document) -> &Paragraph {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph block"),
        }
    }

    /// Helper: get all blocks from the first page.
    fn all_blocks(doc: &Document) -> &[Block] {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        &page.content
    }

    #[test]
    fn test_paragraph_alignment_center() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Centered"))
                .align(docx_rs::AlignmentType::Center),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_paragraph_alignment_right() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Right"))
                .align(docx_rs::AlignmentType::Right),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Right));
    }

    #[test]
    fn test_paragraph_alignment_left() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Left"))
                .align(docx_rs::AlignmentType::Left),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Left));
    }

    #[test]
    fn test_paragraph_alignment_justify() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Justified"))
                .align(docx_rs::AlignmentType::Both),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Justify));
    }

    #[test]
    fn test_paragraph_indent_left() {
        // 720 twips = 36pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(Some(720), None, None, None),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
    }

    #[test]
    fn test_paragraph_indent_right() {
        // 360 twips = 18pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(None, None, Some(360), None),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_right, Some(18.0));
    }

    #[test]
    fn test_paragraph_indent_first_line() {
        // first line indent: 480 twips = 24pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First line indented"))
                .indent(
                    None,
                    Some(docx_rs::SpecialIndentType::FirstLine(480)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_first_line, Some(24.0));
    }

    #[test]
    fn test_paragraph_indent_hanging() {
        // hanging indent: 360 twips = 18pt (negative first-line indent)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hanging indent"))
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::Hanging(360)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(-18.0));
    }

    #[test]
    fn test_paragraph_line_spacing_auto() {
        // Auto line spacing: line=480 means 480/240 = 2.0 (double spacing)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Double spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(480),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 2.0).abs() < 0.01,
                    "Expected 2.0 (double spacing), got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_line_spacing_exact() {
        // Exact line spacing: line=240 twips = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Exact spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Exact)
                        .line(240),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Exact(pts)) => {
                assert!((pts - 12.0).abs() < 0.01, "Expected 12pt, got {pts}");
            }
            other => panic!("Expected Exact line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_space_before_after() {
        // before=240 twips = 12pt, after=120 twips = 6pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Spaced paragraph"))
                .line_spacing(docx_rs::LineSpacing::new().before(240).after(120)),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.space_before, Some(12.0));
        assert_eq!(para.style.space_after, Some(6.0));
    }

    #[test]
    fn test_paragraph_page_break_before() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before break")),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("After break"))
                .page_break_before(true),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let blocks = all_blocks(&doc);
        // Should have: Paragraph("Before break"), PageBreak, Paragraph("After break")
        assert_eq!(blocks.len(), 3, "Expected 3 blocks, got {}", blocks.len());
        assert!(matches!(&blocks[0], Block::Paragraph(_)));
        assert!(matches!(&blocks[1], Block::PageBreak));
        assert!(matches!(&blocks[2], Block::Paragraph(_)));
    }

    #[test]
    fn test_paragraph_combined_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Styled paragraph"))
                .align(docx_rs::AlignmentType::Center)
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::FirstLine(360)),
                    None,
                    None,
                )
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(360)
                        .before(120)
                        .after(60),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(18.0));
        assert_eq!(para.style.space_before, Some(6.0));
        assert_eq!(para.style.space_after, Some(3.0));
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 1.5).abs() < 0.01,
                    "Expected 1.5 spacing, got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_multiple_runs_with_different_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bold ").bold())
                .add_run(docx_rs::Run::new().add_text("Italic ").italic())
                .add_run(docx_rs::Run::new().add_text("Plain")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert!(para.runs[0].style.italic.is_none());
        assert!(para.runs[1].style.bold.is_none());
        assert_eq!(para.runs[1].style.italic, Some(true));
        assert!(para.runs[2].style.bold.is_none());
        assert!(para.runs[2].style.italic.is_none());
    }

    // ----- Table parsing tests (US-007) -----

    /// Helper: build a DOCX with a table using docx-rs builder.
    fn build_docx_with_table(table: docx_rs::Table) -> Vec<u8> {
        let docx = docx_rs::Docx::new().add_table(table);
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: extract the first table block from a parsed document.
    fn first_table(doc: &Document) -> &crate::ir::Table {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        for block in &page.content {
            if let Block::Table(t) = block {
                return t;
            }
        }
        panic!("No Table block found");
    }

    #[test]
    fn test_table_simple_2x2() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A1")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 3000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 2);
        assert_eq!(t.rows[0].cells.len(), 2);
        assert_eq!(t.rows[1].cells.len(), 2);

        // Check cell content
        let cell_text = |row: usize, col: usize| -> String {
            t.rows[row].cells[col]
                .content
                .iter()
                .filter_map(|b| match b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
                    _ => None,
                })
                .collect::<String>()
        };
        assert_eq!(cell_text(0, 0), "A1");
        assert_eq!(cell_text(0, 1), "B1");
        assert_eq!(cell_text(1, 0), "A2");
        assert_eq!(cell_text(1, 1), "B2");
    }

    #[test]
    fn test_table_column_widths_from_grid() {
        // Grid widths in twips: 2000, 3000 → 100pt, 150pt
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B")),
            ),
        ])])
        .set_grid(vec![2000, 3000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.column_widths.len(), 2);
        assert!(
            (t.column_widths[0] - 100.0).abs() < 0.1,
            "Expected 100pt, got {}",
            t.column_widths[0]
        );
        assert!(
            (t.column_widths[1] - 150.0).abs() < 0.1,
            "Expected 150pt, got {}",
            t.column_widths[1]
        );
    }

    #[test]
    fn test_table_cell_with_formatted_text() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Bold").bold())
                    .add_run(docx_rs::Run::new().add_text(" and italic").italic()),
            ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let para = match &cell.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph in cell"),
        };
        assert_eq!(para.runs.len(), 2);
        assert_eq!(para.runs[0].text, "Bold");
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert_eq!(para.runs[1].text, " and italic");
        assert_eq!(para.runs[1].style.italic, Some(true));
    }

    #[test]
    fn test_table_colspan_via_grid_span() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
                    )
                    .grid_span(2),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        // First row: one merged cell with colspan=2
        assert_eq!(t.rows[0].cells.len(), 1);
        assert_eq!(t.rows[0].cells[0].col_span, 2);

        // Second row: two normal cells
        assert_eq!(t.rows[1].cells.len(), 2);
        assert_eq!(t.rows[1].cells[0].col_span, 1);
    }

    #[test]
    fn test_table_rowspan_via_vmerge() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tall")),
                    )
                    .vertical_merge(docx_rs::VMergeType::Restart),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 3);

        // First row: the restart cell should have rowspan=3
        let tall_cell = &t.rows[0].cells[0];
        assert_eq!(tall_cell.row_span, 3);

        // Second and third rows: continue cells should be skipped
        // so rows[1] and rows[2] should have only 1 cell each (B2, B3)
        assert_eq!(t.rows[1].cells.len(), 1);
        assert_eq!(t.rows[2].cells.len(), 1);
    }

    #[test]
    fn test_table_cell_background_color() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Red bg")),
                )
                .shading(docx_rs::Shading::new().fill("FF0000")),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("No bg")),
            ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows[0].cells[0].background, Some(Color::new(255, 0, 0)));
        assert!(t.rows[0].cells[1].background.is_none());
    }

    #[test]
    fn test_table_cell_borders() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bordered")),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                        .size(16)
                        .color("FF0000"),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                        .size(8)
                        .color("0000FF"),
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected cell border");

        // Top: size=16 eighths → 2pt, color=FF0000
        let top = border.top.as_ref().expect("Expected top border");
        assert!(
            (top.width - 2.0).abs() < 0.01,
            "Expected 2pt, got {}",
            top.width
        );
        assert_eq!(top.color, Color::new(255, 0, 0));

        // Bottom: size=8 eighths → 1pt, color=0000FF
        let bottom = border.bottom.as_ref().expect("Expected bottom border");
        assert!(
            (bottom.width - 1.0).abs() < 0.01,
            "Expected 1pt, got {}",
            bottom.width
        );
        assert_eq!(bottom.color, Color::new(0, 0, 255));
    }

    #[test]
    fn test_table_cell_with_multiple_paragraphs() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 1")),
                )
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 2")),
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let paras: Vec<&str> = cell
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) if !p.runs.is_empty() => Some(p.runs[0].text.as_str()),
                _ => None,
            })
            .collect();
        assert!(paras.contains(&"Para 1"), "Expected 'Para 1' in cell");
        assert!(paras.contains(&"Para 2"), "Expected 'Para 2' in cell");
    }

    #[test]
    fn test_table_with_paragraph_before_and_after() {
        let data = {
            let docx = docx_rs::Docx::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before")),
                )
                .add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")),
                    ),
                ])]))
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After")),
                );
            let buf = Vec::new();
            let mut cursor = Cursor::new(buf);
            docx.build().pack(&mut cursor).unwrap();
            cursor.into_inner()
        };

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let blocks = all_blocks(&doc);

        // Should have: Paragraph("Before"), Table, Paragraph("After")
        assert!(
            blocks.len() >= 3,
            "Expected at least 3 blocks, got {}",
            blocks.len()
        );
        assert!(matches!(&blocks[0], Block::Paragraph(_)));
        let has_table = blocks.iter().any(|b| matches!(b, Block::Table(_)));
        assert!(has_table, "Expected a Table block");
    }

    #[test]
    fn test_table_colspan_and_rowspan_combined() {
        // 3x3 table with top-left 2x2 merged
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Big")),
                    )
                    .grid_span(2)
                    .vertical_merge(docx_rs::VMergeType::Restart),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .grid_span(2)
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C2")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A3")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C3")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        // First row: "Big" (colspan=2, rowspan=2) + "C1"
        let big_cell = &t.rows[0].cells[0];
        assert_eq!(big_cell.col_span, 2, "Expected colspan=2");
        assert_eq!(big_cell.row_span, 2, "Expected rowspan=2");

        // Second row: continue cell skipped, so only "C2"
        assert_eq!(t.rows[1].cells.len(), 1);

        // Third row: three normal cells
        assert_eq!(t.rows[2].cells.len(), 3);
    }

    #[test]
    fn test_table_empty_cells() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
            docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 1);
        assert_eq!(t.rows[0].cells.len(), 2);
        // Empty cells should still have content (possibly empty paragraphs)
        for cell in &t.rows[0].cells {
            assert_eq!(cell.col_span, 1);
            assert_eq!(cell.row_span, 1);
        }
    }

    // ── Image extraction tests ──────────────────────────────────────────

    /// Build a minimal valid 1×1 red pixel BMP image.
    /// BMP is trivially decodable by the `image` crate (no compression).
    fn make_test_bmp() -> Vec<u8> {
        let mut bmp = Vec::new();
        // BMP file header (14 bytes)
        bmp.extend_from_slice(b"BM"); // magic
        bmp.extend_from_slice(&58u32.to_le_bytes()); // file size
        bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
        bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset

        // BITMAPINFOHEADER (40 bytes)
        bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
        bmp.extend_from_slice(&1i32.to_le_bytes()); // width
        bmp.extend_from_slice(&1i32.to_le_bytes()); // height
        bmp.extend_from_slice(&1u16.to_le_bytes()); // color planes
        bmp.extend_from_slice(&24u16.to_le_bytes()); // bits per pixel (RGB)
        bmp.extend_from_slice(&0u32.to_le_bytes()); // compression (none)
        bmp.extend_from_slice(&4u32.to_le_bytes()); // image size (row aligned to 4 bytes)
        bmp.extend_from_slice(&0u32.to_le_bytes()); // x pixels/meter
        bmp.extend_from_slice(&0u32.to_le_bytes()); // y pixels/meter
        bmp.extend_from_slice(&0u32.to_le_bytes()); // total colors
        bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors

        // Pixel data: 1 pixel BGR (red = 00 00 FF) + 1 byte padding
        bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);

        bmp
    }

    /// Build a DOCX containing an inline image using Pic::new() which processes
    /// the image through the `image` crate, ensuring valid PNG output in the ZIP.
    fn build_docx_with_image(width_px: u32, height_px: u32) -> Vec<u8> {
        let bmp_data = make_test_bmp();
        // Use Pic::new() which decodes the BMP and re-encodes as PNG internally
        let pic = docx_rs::Pic::new(&bmp_data).size(width_px * 9525, height_px * 9525);
        let para_with_image = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic));
        let docx = docx_rs::Docx::new().add_paragraph(para_with_image);
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: find all Image blocks in a FlowPage.
    fn find_images(doc: &Document) -> Vec<&ImageData> {
        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        page.content
            .iter()
            .filter_map(|b| match b {
                Block::Image(img) => Some(img),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_docx_image_inline_basic() {
        let data = build_docx_with_image(100, 80);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        assert_eq!(images.len(), 1, "Expected exactly one image block");
        assert!(!images[0].data.is_empty(), "Image data should not be empty");
    }

    #[test]
    fn test_docx_image_format_is_png() {
        let data = build_docx_with_image(50, 50);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        assert_eq!(
            images[0].format,
            ImageFormat::Png,
            "Image format should be PNG"
        );
    }

    #[test]
    fn test_docx_image_dimensions() {
        // 100px × 80px → EMU: 100*9525=952500, 80*9525=762000
        // EMU to points: 952500/12700=75.0, 762000/12700=60.0
        let data = build_docx_with_image(100, 80);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        let img = images[0];

        let width = img.width.expect("Expected width");
        let height = img.height.expect("Expected height");

        assert!(
            (width - 75.0).abs() < 0.1,
            "Expected width ~75pt, got {width}"
        );
        assert!(
            (height - 60.0).abs() < 0.1,
            "Expected height ~60pt, got {height}"
        );
    }

    #[test]
    fn test_docx_image_with_text_paragraphs() {
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data);
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before image")),
            )
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)))
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After image")),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };

        let has_image = page.content.iter().any(|b| matches!(b, Block::Image(_)));
        assert!(has_image, "Expected an image block in the content");

        let has_before = page.content.iter().any(|b| match b {
            Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("Before")),
            _ => false,
        });
        assert!(has_before, "Expected 'Before image' text");

        let has_after = page.content.iter().any(|b| match b {
            Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("After")),
            _ => false,
        });
        assert!(has_after, "Expected 'After image' text");
    }

    #[test]
    fn test_docx_multiple_images() {
        let bmp_data = make_test_bmp();
        let pic1 = docx_rs::Pic::new(&bmp_data).size(100 * 9525, 100 * 9525);
        let pic2 = docx_rs::Pic::new(&bmp_data).size(200 * 9525, 150 * 9525);
        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic1)))
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic2)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        assert!(
            images.len() >= 2,
            "Expected at least 2 images, got {}",
            images.len()
        );
    }

    #[test]
    fn test_docx_image_data_contains_png_header() {
        let data = build_docx_with_image(50, 50);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        let img_data = &images[0].data;

        // docx-rs converts all images to PNG; verify PNG magic bytes
        assert!(
            img_data.len() >= 8 && img_data[0..4] == [0x89, 0x50, 0x4E, 0x47],
            "Image data should start with PNG magic bytes"
        );
    }

    #[test]
    fn test_docx_no_images_produces_no_image_blocks() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };

        let image_count = page
            .content
            .iter()
            .filter(|b| matches!(b, Block::Image(_)))
            .count();
        assert_eq!(image_count, 0, "Expected no image blocks");
    }

    #[test]
    fn test_docx_image_with_custom_emu_size() {
        // Create image with specific EMU size via .size() override
        // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data).size(2_540_000, 1_270_000);
        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data).unwrap();

        let images = find_images(&doc);
        assert!(!images.is_empty(), "Expected at least one image");
        let img = images[0];

        let width = img.width.expect("Expected width");
        let height = img.height.expect("Expected height");

        assert!(
            (width - 200.0).abs() < 0.1,
            "Expected width ~200pt, got {width}"
        );
        assert!(
            (height - 100.0).abs() < 0.1,
            "Expected height ~100pt, got {height}"
        );
    }
}
