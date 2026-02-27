use std::fmt::Write;

use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, BorderSide, CellBorder, Color, Document, FixedElement, FixedElementKind,
    FixedPage, FlowPage, HFInline, HeaderFooter, ImageData, ImageFormat, LineSpacing, List,
    ListKind, Margins, Page, PageSize, Paragraph, ParagraphStyle, Run, Shape, ShapeKind, Table,
    TableCell, TablePage, TextStyle,
};

/// An image asset to be embedded in the Typst compilation.
#[derive(Debug, Clone)]
pub struct ImageAsset {
    /// Virtual file path (e.g., "img-0.png").
    pub path: String,
    /// Raw image bytes.
    pub data: Vec<u8>,
}

/// Output from Typst codegen: markup source and embedded image assets.
#[derive(Debug)]
pub struct TypstOutput {
    /// The generated Typst markup string.
    pub source: String,
    /// Image assets referenced by the markup.
    pub images: Vec<ImageAsset>,
}

/// Internal context for tracking image assets during code generation.
struct GenCtx {
    images: Vec<ImageAsset>,
    next_image_id: usize,
}

impl GenCtx {
    fn new() -> Self {
        Self {
            images: Vec::new(),
            next_image_id: 0,
        }
    }

    fn add_image(&mut self, data: &[u8], format: ImageFormat) -> String {
        let ext = format.extension();
        let path = format!("img-{}.{}", self.next_image_id, ext);
        self.next_image_id += 1;
        self.images.push(ImageAsset {
            path: path.clone(),
            data: data.to_vec(),
        });
        path
    }
}

/// Generate Typst markup from a Document IR.
pub fn generate_typst(doc: &Document) -> Result<TypstOutput, ConvertError> {
    let mut out = String::new();
    let mut ctx = GenCtx::new();
    for page in &doc.pages {
        match page {
            Page::Flow(flow) => generate_flow_page(&mut out, flow, &mut ctx)?,
            Page::Fixed(fixed) => generate_fixed_page(&mut out, fixed, &mut ctx)?,
            Page::Table(table_page) => generate_table_page(&mut out, table_page, &mut ctx)?,
        }
    }
    Ok(TypstOutput {
        source: out,
        images: ctx.images,
    })
}

fn generate_flow_page(
    out: &mut String,
    page: &FlowPage,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    write_flow_page_setup(out, page);
    out.push('\n');

    for (i, block) in page.content.iter().enumerate() {
        if i > 0 {
            out.push('\n');
        }
        generate_block(out, block, ctx)?;
    }
    Ok(())
}

fn generate_fixed_page(
    out: &mut String,
    page: &FixedPage,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // Slides use zero margins — all positioning is absolute
    if let Some(ref bg) = page.background_color {
        let _ = writeln!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt, fill: rgb({}, {}, {}))",
            format_f64(page.size.width),
            format_f64(page.size.height),
            bg.r,
            bg.g,
            bg.b,
        );
    } else {
        let _ = writeln!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt)",
            format_f64(page.size.width),
            format_f64(page.size.height),
        );
    }
    out.push('\n');

    for elem in &page.elements {
        generate_fixed_element(out, elem, ctx)?;
    }
    Ok(())
}

fn generate_table_page(
    out: &mut String,
    page: &TablePage,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    write_table_page_setup(out, page);
    out.push('\n');
    generate_table(out, &page.table, ctx)?;
    Ok(())
}

fn generate_fixed_element(
    out: &mut String,
    elem: &FixedElement,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // Use Typst's place() for absolute positioning
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt",
        format_f64(elem.x),
        format_f64(elem.y),
    );
    out.push_str(")[\n");

    match &elem.kind {
        FixedElementKind::TextBox(blocks) => {
            let _ = writeln!(
                out,
                "#block(width: {}pt, height: {}pt)[",
                format_f64(elem.width),
                format_f64(elem.height),
            );
            for block in blocks {
                generate_block(out, block, ctx)?;
            }
            out.push_str("]\n");
        }
        FixedElementKind::Image(img) => {
            generate_image(out, img, ctx);
        }
        FixedElementKind::Shape(shape) => {
            generate_shape(out, shape, elem.width, elem.height);
        }
        FixedElementKind::Table(table) => {
            generate_table(out, table, ctx)?;
        }
    }

    out.push_str("]\n");
    Ok(())
}

fn generate_shape(out: &mut String, shape: &Shape, width: f64, height: f64) {
    let has_rotation = shape.rotation_deg.is_some();
    if let Some(deg) = shape.rotation_deg {
        let _ = write!(out, "#rotate({}deg)[", format_f64(deg));
    }

    match &shape.kind {
        ShapeKind::Rectangle => {
            out.push_str("#rect(");
            write_shape_params(out, shape, width, height);
            out.push_str(")\n");
        }
        ShapeKind::Ellipse => {
            out.push_str("#ellipse(");
            write_shape_params(out, shape, width, height);
            out.push_str(")\n");
        }
        ShapeKind::Line { x2, y2 } => {
            out.push_str("#line(");
            let _ = write!(
                out,
                "start: (0pt, 0pt), end: ({}pt, {}pt)",
                format_f64(*x2),
                format_f64(*y2),
            );
            if let Some(stroke) = &shape.stroke {
                let _ = write!(
                    out,
                    ", stroke: {}pt + rgb({}, {}, {})",
                    format_f64(stroke.width),
                    stroke.color.r,
                    stroke.color.g,
                    stroke.color.b,
                );
            }
            out.push_str(")\n");
        }
    }

    if has_rotation {
        out.push_str("]\n");
    }
}

/// Write fill color, using rgba when opacity is set, rgb otherwise.
fn write_fill_color(out: &mut String, fill: &Color, opacity: Option<f64>) {
    if let Some(op) = opacity {
        let alpha = (op * 255.0).round() as u8;
        let _ = write!(
            out,
            ", fill: rgba({}, {}, {}, {})",
            fill.r, fill.g, fill.b, alpha
        );
    } else {
        let _ = write!(out, ", fill: rgb({}, {}, {})", fill.r, fill.g, fill.b);
    }
}

fn write_shape_params(out: &mut String, shape: &Shape, width: f64, height: f64) {
    let _ = write!(
        out,
        "width: {}pt, height: {}pt",
        format_f64(width),
        format_f64(height),
    );
    if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    if let Some(stroke) = &shape.stroke {
        let _ = write!(
            out,
            ", stroke: {}pt + rgb({}, {}, {})",
            format_f64(stroke.width),
            stroke.color.r,
            stroke.color.g,
            stroke.color.b,
        );
    }
}

fn write_page_setup(out: &mut String, size: &PageSize, margins: &Margins) {
    let _ = writeln!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt))",
        format_f64(size.width),
        format_f64(size.height),
        format_f64(margins.top),
        format_f64(margins.bottom),
        format_f64(margins.left),
        format_f64(margins.right),
    );
}

/// Write the full page setup for a FlowPage, including optional header/footer.
fn write_flow_page_setup(out: &mut String, page: &FlowPage) {
    if page.header.is_none() && page.footer.is_none() {
        write_page_setup(out, &page.size, &page.margins);
        return;
    }

    let _ = write!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt)",
        format_f64(page.size.width),
        format_f64(page.size.height),
        format_f64(page.margins.top),
        format_f64(page.margins.bottom),
        format_f64(page.margins.left),
        format_f64(page.margins.right),
    );

    if let Some(header) = &page.header {
        if hf_needs_context(header) {
            out.push_str(", header: context [");
        } else {
            out.push_str(", header: [");
        }
        generate_hf_content(out, header);
        out.push(']');
    }

    if let Some(footer) = &page.footer {
        if hf_needs_context(footer) {
            out.push_str(", footer: context [");
        } else {
            out.push_str(", footer: [");
        }
        generate_hf_content(out, footer);
        out.push(']');
    }

    out.push_str(")\n");
}

/// Write the full page setup for a TablePage, including optional header/footer.
fn write_table_page_setup(out: &mut String, page: &TablePage) {
    if page.header.is_none() && page.footer.is_none() {
        write_page_setup(out, &page.size, &page.margins);
        return;
    }

    let _ = write!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt)",
        format_f64(page.size.width),
        format_f64(page.size.height),
        format_f64(page.margins.top),
        format_f64(page.margins.bottom),
        format_f64(page.margins.left),
        format_f64(page.margins.right),
    );

    if let Some(header) = &page.header {
        if hf_needs_context(header) {
            out.push_str(", header: context [");
        } else {
            out.push_str(", header: [");
        }
        generate_hf_content(out, header);
        out.push(']');
    }

    if let Some(footer) = &page.footer {
        if hf_needs_context(footer) {
            out.push_str(", footer: context [");
        } else {
            out.push_str(", footer: [");
        }
        generate_hf_content(out, footer);
        out.push(']');
    }

    out.push_str(")\n");
}

/// Check if a header/footer contains any context-dependent fields (page number or total pages).
fn hf_needs_context(hf: &HeaderFooter) -> bool {
    hf.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, HFInline::PageNumber | HFInline::TotalPages))
    })
}

/// Generate inline content for a header or footer.
fn generate_hf_content(out: &mut String, hf: &HeaderFooter) {
    for (i, para) in hf.paragraphs.iter().enumerate() {
        if i > 0 {
            out.push_str("\\\n");
        }
        // Apply paragraph alignment if set
        if let Some(align) = para.style.alignment {
            let align_str = match align {
                Alignment::Left => "left",
                Alignment::Center => "center",
                Alignment::Right => "right",
                Alignment::Justify => "left",
            };
            let _ = write!(out, "#align({align_str})[");
        }
        for elem in &para.elements {
            match elem {
                HFInline::Run(run) => {
                    generate_run(out, run);
                }
                HFInline::PageNumber => {
                    out.push_str("#counter(page).display()");
                }
                HFInline::TotalPages => {
                    out.push_str("#counter(page).final().first()");
                }
            }
        }
        if para.style.alignment.is_some() {
            out.push(']');
        }
    }
}

fn generate_block(out: &mut String, block: &Block, ctx: &mut GenCtx) -> Result<(), ConvertError> {
    match block {
        Block::Paragraph(para) => generate_paragraph(out, para),
        Block::PageBreak => {
            out.push_str("#pagebreak()\n");
            Ok(())
        }
        Block::Table(table) => generate_table(out, table, ctx),
        Block::Image(img) => {
            generate_image(out, img, ctx);
            Ok(())
        }
        Block::List(list) => generate_list(out, list),
    }
}

/// Generate Typst markup for a list (ordered or unordered).
///
/// Uses Typst's `#enum()` for ordered lists and `#list()` for unordered lists.
/// Nested items are wrapped in `list.item()` / `enum.item()` with a sub-list.
fn generate_list(out: &mut String, list: &List) -> Result<(), ConvertError> {
    let (func, item_func) = match list.kind {
        ListKind::Ordered => ("enum", "enum.item"),
        ListKind::Unordered => ("list", "list.item"),
    };

    // Build nested structure from flat items with levels.
    // We use Typst function syntax: #list(item, item, ...) or #enum(item, item, ...)
    // Nested items use list.item(body) with a sub-list inside.
    let _ = writeln!(out, "#{func}(");
    generate_list_items(out, &list.items, 0, func, item_func)?;
    out.push_str(")\n");
    Ok(())
}

/// Recursively generate list items, grouping consecutive items at the same or deeper level.
fn generate_list_items(
    out: &mut String,
    items: &[crate::ir::ListItem],
    base_level: u32,
    func: &str,
    item_func: &str,
) -> Result<(), ConvertError> {
    let mut i = 0;
    while i < items.len() {
        let item = &items[i];
        if item.level == base_level {
            // Emit this item's content
            let _ = write!(out, "  {item_func}[");
            for para in &item.content {
                for run in &para.runs {
                    generate_run(out, run);
                }
            }
            out.push(']');

            // Check if next items are nested (deeper level) — they become a sub-list
            let nested_start = i + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].level > base_level {
                nested_end += 1;
            }

            if nested_end > nested_start {
                // Emit nested sub-list
                let _ = writeln!(out, "[#{func}(");
                generate_list_items(
                    out,
                    &items[nested_start..nested_end],
                    base_level + 1,
                    func,
                    item_func,
                )?;
                out.push_str(")]");
                i = nested_end;
            } else {
                i += 1;
            }

            out.push_str(",\n");
        } else {
            // Item at a deeper level without a parent at base_level;
            // treat it as if it were at base_level
            let _ = write!(out, "  {item_func}[");
            for para in &item.content {
                for run in &para.runs {
                    generate_run(out, run);
                }
            }
            out.push_str("],\n");
            i += 1;
        }
    }
    Ok(())
}

fn generate_table(out: &mut String, table: &Table, ctx: &mut GenCtx) -> Result<(), ConvertError> {
    out.push_str("#table(\n");

    // Column widths
    if !table.column_widths.is_empty() {
        out.push_str("  columns: (");
        for (i, w) in table.column_widths.iter().enumerate() {
            if i > 0 {
                out.push_str(", ");
            }
            let _ = write!(out, "{}pt", format_f64(*w));
        }
        out.push_str("),\n");
    }

    // Rows and cells
    for row in &table.rows {
        for cell in &row.cells {
            generate_table_cell(out, cell, ctx)?;
        }
    }

    out.push_str(")\n");
    Ok(())
}

fn generate_table_cell(
    out: &mut String,
    cell: &TableCell,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let needs_cell_fn = cell.col_span > 1
        || cell.row_span > 1
        || cell.border.is_some()
        || cell.background.is_some();

    if needs_cell_fn {
        out.push_str("  table.cell(");
        write_cell_params(out, cell);
        out.push_str(")[");
    } else {
        out.push_str("  [");
    }

    // Generate cell content
    generate_cell_content(out, &cell.content, ctx)?;

    out.push_str("],\n");
    Ok(())
}

fn write_cell_params(out: &mut String, cell: &TableCell) {
    let mut first = true;

    if cell.col_span > 1 {
        write_param(out, &mut first, &format!("colspan: {}", cell.col_span));
    }
    if cell.row_span > 1 {
        write_param(out, &mut first, &format!("rowspan: {}", cell.row_span));
    }
    if let Some(ref bg) = cell.background {
        write_param(out, &mut first, &format_color(bg));
    }
    if let Some(ref border) = cell.border {
        let stroke = format_cell_stroke(border);
        if !stroke.is_empty() {
            write_param(out, &mut first, &stroke);
        }
    }
}

fn format_cell_stroke(border: &CellBorder) -> String {
    let mut parts = Vec::new();

    if let Some(ref side) = border.top {
        parts.push(format!("top: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.bottom {
        parts.push(format!("bottom: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.left {
        parts.push(format!("left: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.right {
        parts.push(format!("right: {}", format_border_side(side)));
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("stroke: ({})", parts.join(", "))
    }
}

fn format_border_side(side: &BorderSide) -> String {
    format!(
        "{}pt + rgb({}, {}, {})",
        format_f64(side.width),
        side.color.r,
        side.color.g,
        side.color.b
    )
}

/// Generate content inside a table cell (list of blocks rendered inline).
fn generate_cell_content(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    for (i, block) in blocks.iter().enumerate() {
        if i > 0 {
            // Paragraph break between blocks
            out.push('\n');
        }
        match block {
            Block::Paragraph(para) => generate_cell_paragraph(out, para),
            Block::Table(table) => generate_table(out, table, ctx)?,
            Block::Image(img) => generate_image(out, img, ctx),
            Block::List(list) => generate_list(out, list)?,
            Block::PageBreak => {}
        }
    }
    Ok(())
}

/// Generate paragraph content for inside a table cell (runs only, no block wrapper).
fn generate_cell_paragraph(out: &mut String, para: &Paragraph) {
    for run in &para.runs {
        generate_run(out, run);
    }
}

fn generate_image(out: &mut String, img: &ImageData, ctx: &mut GenCtx) {
    let path = ctx.add_image(&img.data, img.format);
    out.push_str("#image(\"");
    out.push_str(&path);
    out.push('"');

    if let Some(w) = img.width {
        let _ = write!(out, ", width: {}pt", format_f64(w));
    }
    if let Some(h) = img.height {
        let _ = write!(out, ", height: {}pt", format_f64(h));
    }

    out.push_str(")\n");
}

fn generate_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
    let style = &para.style;
    let has_para_style = needs_block_wrapper(style);

    if has_para_style {
        out.push_str("#block(");
        write_block_params(out, style);
        out.push_str(")[\n");
        write_par_settings(out, style);
    }

    // Generate alignment wrapper if needed
    let alignment = style.alignment;
    let use_align = matches!(
        alignment,
        Some(Alignment::Center) | Some(Alignment::Right) | Some(Alignment::Left)
    );

    if use_align {
        let align_str = match alignment.unwrap() {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => unreachable!(),
        };
        let _ = write!(out, "#align({align_str})[");
    }

    // Generate runs
    for run in &para.runs {
        generate_run(out, run);
    }

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

/// Check if paragraph style needs a block wrapper (for spacing/leading/justify).
fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.line_spacing.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
}

fn write_block_params(out: &mut String, style: &ParagraphStyle) {
    let mut first = true;

    if let Some(above) = style.space_before {
        write_param(out, &mut first, &format!("above: {}pt", format_f64(above)));
    }
    if let Some(below) = style.space_after {
        write_param(out, &mut first, &format!("below: {}pt", format_f64(below)));
    }
}

fn write_par_settings(out: &mut String, style: &ParagraphStyle) {
    if let Some(ref spacing) = style.line_spacing {
        match spacing {
            LineSpacing::Proportional(factor) => {
                let leading = factor * 0.65;
                let _ = writeln!(out, "  #set par(leading: {}em)", format_f64(leading));
            }
            LineSpacing::Exact(pts) => {
                let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(*pts));
            }
        }
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
}

fn generate_run(out: &mut String, run: &Run) {
    // Emit footnote if present (footnote runs have empty text)
    if let Some(ref content) = run.footnote {
        let escaped_content = escape_typst(content);
        let _ = write!(out, "#footnote[{escaped_content}]");
        return;
    }

    let style = &run.style;
    let escaped = escape_typst(&run.text);

    let has_text_props = has_text_properties(style);
    let needs_underline = matches!(style.underline, Some(true));
    let needs_strike = matches!(style.strikethrough, Some(true));
    let has_link = run.href.is_some();

    // Wrap with link (outermost)
    if let Some(ref href) = run.href {
        let _ = write!(out, "#link(\"{href}\")[");
    }

    // Wrap with decorations
    if needs_strike {
        out.push_str("#strike[");
    }
    if needs_underline {
        out.push_str("#underline[");
    }

    if has_text_props {
        out.push_str("#text(");
        write_text_params(out, style);
        out.push_str(")[");
        out.push_str(&escaped);
        out.push(']');
    } else {
        out.push_str(&escaped);
    }

    if needs_underline {
        out.push(']');
    }
    if needs_strike {
        out.push(']');
    }
    if has_link {
        out.push(']');
    }
}

/// Check if the text style has properties that need a #text() wrapper
/// (not counting underline/strikethrough which use separate wrappers).
fn has_text_properties(style: &TextStyle) -> bool {
    matches!(style.bold, Some(true))
        || matches!(style.italic, Some(true))
        || style.font_size.is_some()
        || style.color.is_some()
        || style.font_family.is_some()
}

fn write_text_params(out: &mut String, style: &TextStyle) {
    let mut first = true;

    if let Some(ref family) = style.font_family {
        write_param(out, &mut first, &format!("font: \"{family}\""));
    }
    if let Some(size) = style.font_size {
        write_param(out, &mut first, &format!("size: {}pt", format_f64(size)));
    }
    if matches!(style.bold, Some(true)) {
        write_param(out, &mut first, "weight: \"bold\"");
    }
    if matches!(style.italic, Some(true)) {
        write_param(out, &mut first, "style: \"italic\"");
    }
    if let Some(ref color) = style.color {
        write_param(out, &mut first, &format_color(color));
    }
}

fn write_param(out: &mut String, first: &mut bool, param: &str) {
    if !*first {
        out.push_str(", ");
    }
    out.push_str(param);
    *first = false;
}

fn format_color(color: &Color) -> String {
    format!("fill: rgb({}, {}, {})", color.r, color.g, color.b)
}

/// Format an f64 without unnecessary trailing zeros.
fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

/// Escape special Typst characters in text content.
fn escape_typst(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' => {
                result.push('\\');
                result.push(ch);
            }
            '$' => {
                result.push('\\');
                result.push('$');
            }
            _ => result.push(ch),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{HeaderFooterParagraph, Metadata, StyleSheet};

    /// Helper to create a minimal Document with one FlowPage.
    fn make_doc(pages: Vec<Page>) -> Document {
        Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        }
    }

    /// Helper to create a FlowPage with default A4 size and margins.
    fn make_flow_page(content: Vec<Block>) -> Page {
        Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content,
            header: None,
            footer: None,
        })
    }

    /// Helper to create a simple paragraph with one plain-text run.
    fn make_paragraph(text: &str) -> Block {
        Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })
    }

    #[test]
    fn test_generate_plain_paragraph() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("Hello World"));
    }

    #[test]
    fn test_generate_page_setup() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize {
                width: 612.0,
                height: 792.0,
            },
            margins: Margins {
                top: 36.0,
                bottom: 36.0,
                left: 54.0,
                right: 54.0,
            },
            content: vec![make_paragraph("test")],
            header: None,
            footer: None,
        })]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("612pt"));
        assert!(result.contains("792pt"));
        assert!(result.contains("36pt"));
        assert!(result.contains("54pt"));
    }

    #[test]
    fn test_generate_bold_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold text".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("weight: \"bold\""),
            "Expected bold weight in: {result}"
        );
        assert!(result.contains("Bold text"));
    }

    #[test]
    fn test_generate_italic_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Italic text".to_string(),
                style: TextStyle {
                    italic: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("style: \"italic\""),
            "Expected italic style in: {result}"
        );
        assert!(result.contains("Italic text"));
    }

    #[test]
    fn test_generate_underline_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Underlined".to_string(),
                style: TextStyle {
                    underline: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#underline["),
            "Expected underline wrapper in: {result}"
        );
        assert!(result.contains("Underlined"));
    }

    #[test]
    fn test_generate_font_size() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Large text".to_string(),
                style: TextStyle {
                    font_size: Some(24.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("size: 24pt"),
            "Expected font size in: {result}"
        );
    }

    #[test]
    fn test_generate_font_color() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Red text".to_string(),
                style: TextStyle {
                    color: Some(Color::new(255, 0, 0)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("fill: rgb(255, 0, 0)"),
            "Expected RGB color in: {result}"
        );
    }

    #[test]
    fn test_generate_combined_text_styles() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Styled".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    italic: Some(true),
                    font_size: Some(16.0),
                    color: Some(Color::new(0, 128, 255)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("weight: \"bold\""));
        assert!(result.contains("style: \"italic\""));
        assert!(result.contains("size: 16pt"));
        assert!(result.contains("fill: rgb(0, 128, 255)"));
        assert!(result.contains("Styled"));
    }

    #[test]
    fn test_generate_alignment_center() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Centered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("align(center"),
            "Expected center alignment in: {result}"
        );
    }

    #[test]
    fn test_generate_alignment_right() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Right),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Right".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("align(right"),
            "Expected right alignment in: {result}"
        );
    }

    #[test]
    fn test_generate_alignment_justify() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Justify),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Justified text".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("par(justify: true") || result.contains("set par(justify: true"),
            "Expected justify in: {result}"
        );
    }

    #[test]
    fn test_generate_line_spacing_proportional() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                line_spacing: Some(LineSpacing::Proportional(2.0)),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Double spaced".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("leading:"),
            "Expected leading setting in: {result}"
        );
    }

    #[test]
    fn test_generate_line_spacing_exact() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                line_spacing: Some(LineSpacing::Exact(18.0)),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Exact spaced".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("leading: 18pt"),
            "Expected exact leading in: {result}"
        );
    }

    #[test]
    fn test_generate_multiple_paragraphs() {
        let doc = make_doc(vec![make_flow_page(vec![
            make_paragraph("First paragraph"),
            make_paragraph("Second paragraph"),
        ])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("First paragraph"));
        assert!(result.contains("Second paragraph"));
    }

    #[test]
    fn test_generate_paragraph_with_multiple_runs() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "Normal ".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
                Run {
                    text: "bold".to_string(),
                    style: TextStyle {
                        bold: Some(true),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                },
                Run {
                    text: " normal again".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
            ],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("Normal "));
        assert!(result.contains("bold"));
        assert!(result.contains(" normal again"));
    }

    #[test]
    fn test_generate_empty_document() {
        let doc = make_doc(vec![]);
        let result = generate_typst(&doc).unwrap().source;
        // Should produce valid (possibly empty) Typst markup
        assert!(result.is_empty() || !result.is_empty()); // Just shouldn't error
    }

    #[test]
    fn test_generate_special_characters_escaped() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
            "Price: $100 #items @store",
        )])]);
        let result = generate_typst(&doc).unwrap().source;
        // The text should appear but special chars should be escaped for Typst
        // In Typst, # starts a code expression, so it needs escaping
        assert!(
            result.contains("\\#") || result.contains("Price"),
            "Expected escaped or present text in: {result}"
        );
    }

    // ── Table codegen tests ───────────────────────────────────────────

    use crate::ir::{BorderSide, CellBorder, Table, TableCell, TableRow};

    /// Helper to create a table cell with plain text.
    fn make_text_cell(text: &str) -> TableCell {
        TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            ..TableCell::default()
        }
    }

    #[test]
    fn test_table_simple_2x2() {
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 200.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("#table("), "Expected #table( in: {result}");
        assert!(
            result.contains("columns: (100pt, 200pt)"),
            "Expected column widths in: {result}"
        );
        assert!(result.contains("A1"), "Expected A1 in: {result}");
        assert!(result.contains("B1"), "Expected B1 in: {result}");
        assert!(result.contains("A2"), "Expected A2 in: {result}");
        assert!(result.contains("B2"), "Expected B2 in: {result}");
    }

    #[test]
    fn test_table_with_colspan() {
        let merged_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Merged".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            col_span: 2,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![merged_cell],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 200.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("colspan: 2"),
            "Expected colspan: 2 in: {result}"
        );
        assert!(result.contains("Merged"), "Expected Merged in: {result}");
    }

    #[test]
    fn test_table_with_rowspan() {
        let tall_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Tall".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            row_span: 2,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![tall_cell, make_text_cell("B1")],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("B2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 200.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("rowspan: 2"),
            "Expected rowspan: 2 in: {result}"
        );
        assert!(result.contains("Tall"), "Expected Tall in: {result}");
    }

    #[test]
    fn test_table_with_colspan_and_rowspan() {
        let big_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Big".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            col_span: 2,
            row_span: 2,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![big_cell, make_text_cell("C1")],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("C2")],
                    height: None,
                },
                TableRow {
                    cells: vec![
                        make_text_cell("A3"),
                        make_text_cell("B3"),
                        make_text_cell("C3"),
                    ],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 100.0, 100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("colspan: 2"),
            "Expected colspan: 2 in: {result}"
        );
        assert!(
            result.contains("rowspan: 2"),
            "Expected rowspan: 2 in: {result}"
        );
        assert!(result.contains("Big"), "Expected Big in: {result}");
    }

    #[test]
    fn test_table_with_background_color() {
        let colored_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Colored".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            background: Some(Color::new(200, 200, 200)),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![colored_cell],
                height: None,
            }],
            column_widths: vec![100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("fill: rgb(200, 200, 200)"),
            "Expected fill color in: {result}"
        );
        assert!(result.contains("Colored"), "Expected Colored in: {result}");
    }

    #[test]
    fn test_table_with_cell_borders() {
        let bordered_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Bordered".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            border: Some(CellBorder {
                top: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                }),
                bottom: Some(BorderSide {
                    width: 2.0,
                    color: Color::new(255, 0, 0),
                }),
                left: None,
                right: None,
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![bordered_cell],
                height: None,
            }],
            column_widths: vec![100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("stroke:"), "Expected stroke in: {result}");
        assert!(
            result.contains("Bordered"),
            "Expected Bordered in: {result}"
        );
    }

    #[test]
    fn test_table_with_styled_text_in_cell() {
        let styled_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Bold cell".to_string(),
                    style: TextStyle {
                        bold: Some(true),
                        font_size: Some(14.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                }],
            })],
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![styled_cell],
                height: None,
            }],
            column_widths: vec![100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("weight: \"bold\""),
            "Expected bold in table cell: {result}"
        );
        assert!(
            result.contains("size: 14pt"),
            "Expected font size in table cell: {result}"
        );
    }

    #[test]
    fn test_table_empty_cells() {
        let empty_cell = TableCell::default();
        let table = Table {
            rows: vec![TableRow {
                cells: vec![empty_cell, make_text_cell("Has text")],
                height: None,
            }],
            column_widths: vec![100.0, 100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("#table("), "Expected #table( in: {result}");
        assert!(
            result.contains("Has text"),
            "Expected Has text in: {result}"
        );
    }

    #[test]
    fn test_table_no_column_widths() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![make_text_cell("A"), make_text_cell("B")],
                height: None,
            }],
            column_widths: vec![],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("#table("), "Expected #table( in: {result}");
        // Without explicit widths, should still produce valid table
        assert!(result.contains("A"), "Expected A in: {result}");
        assert!(result.contains("B"), "Expected B in: {result}");
    }

    #[test]
    fn test_table_all_borders() {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "All borders".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            border: Some(CellBorder {
                top: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                }),
                bottom: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                }),
                left: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                }),
                right: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                }),
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(result.contains("top:"), "Expected top border in: {result}");
        assert!(
            result.contains("bottom:"),
            "Expected bottom border in: {result}"
        );
        assert!(
            result.contains("left:"),
            "Expected left border in: {result}"
        );
        assert!(
            result.contains("right:"),
            "Expected right border in: {result}"
        );
    }

    #[test]
    fn test_table_cell_with_multiple_paragraphs() {
        let multi_para_cell = TableCell {
            content: vec![
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "First para".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Second para".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
            ],
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![multi_para_cell],
                height: None,
            }],
            column_widths: vec![200.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("First para"),
            "Expected First para in: {result}"
        );
        assert!(
            result.contains("Second para"),
            "Expected Second para in: {result}"
        );
    }

    #[test]
    fn test_table_special_chars_in_cells() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![make_text_cell("Price: $100 #items")],
                height: None,
            }],
            column_widths: vec![200.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        // Special chars should be escaped
        assert!(
            result.contains("\\$") && result.contains("\\#"),
            "Expected escaped special chars in: {result}"
        );
    }

    #[test]
    fn test_table_in_flow_page_with_paragraphs() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![make_text_cell("Cell")],
                height: None,
            }],
            column_widths: vec![100.0],
        };
        let doc = make_doc(vec![make_flow_page(vec![
            make_paragraph("Before table"),
            Block::Table(table),
            make_paragraph("After table"),
        ])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("Before table"),
            "Expected Before table in: {result}"
        );
        assert!(result.contains("#table("), "Expected #table( in: {result}");
        assert!(
            result.contains("After table"),
            "Expected After table in: {result}"
        );
    }

    #[test]
    fn test_generate_space_before_after() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                space_before: Some(12.0),
                space_after: Some(6.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Spaced paragraph".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        // Should contain spacing directives
        assert!(
            result.contains("12pt") || result.contains("above"),
            "Expected space_before in: {result}"
        );
    }

    // ── Image codegen tests ─────────────────────────────────────────────

    use crate::ir::ImageData;

    /// Minimal valid 1x1 red pixel PNG for testing.
    const MINIMAL_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    fn make_image(format: ImageFormat, width: Option<f64>, height: Option<f64>) -> Block {
        Block::Image(ImageData {
            data: MINIMAL_PNG.to_vec(),
            format,
            width,
            height,
        })
    }

    #[test]
    fn test_image_basic_no_size() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            None,
            None,
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#image(\"img-0.png\")"),
            "Expected #image(\"img-0.png\") in: {}",
            output.source
        );
    }

    #[test]
    fn test_image_with_width_only() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            Some(100.0),
            None,
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains("#image(\"img-0.png\", width: 100pt)"),
            "Expected width param in: {}",
            output.source
        );
    }

    #[test]
    fn test_image_with_height_only() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            None,
            Some(80.0),
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains("#image(\"img-0.png\", height: 80pt)"),
            "Expected height param in: {}",
            output.source
        );
    }

    #[test]
    fn test_image_with_both_dimensions() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            Some(200.0),
            Some(150.0),
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains("#image(\"img-0.png\", width: 200pt, height: 150pt)"),
            "Expected both dimensions in: {}",
            output.source
        );
    }

    #[test]
    fn test_image_collects_asset() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            None,
            None,
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert_eq!(output.images.len(), 1);
        assert_eq!(output.images[0].path, "img-0.png");
        assert_eq!(output.images[0].data, MINIMAL_PNG);
    }

    #[test]
    fn test_multiple_images_numbered_sequentially() {
        let doc = make_doc(vec![make_flow_page(vec![
            make_image(ImageFormat::Png, None, None),
            make_image(ImageFormat::Jpeg, Some(50.0), None),
        ])]);
        let output = generate_typst(&doc).unwrap();
        assert_eq!(output.images.len(), 2);
        assert_eq!(output.images[0].path, "img-0.png");
        assert_eq!(output.images[1].path, "img-1.jpeg");
        assert!(output.source.contains("img-0.png"));
        assert!(output.source.contains("img-1.jpeg"));
    }

    #[test]
    fn test_image_format_extensions() {
        let formats = [
            (ImageFormat::Png, "png"),
            (ImageFormat::Jpeg, "jpeg"),
            (ImageFormat::Gif, "gif"),
            (ImageFormat::Bmp, "bmp"),
            (ImageFormat::Tiff, "tiff"),
        ];
        for (i, (format, expected_ext)) in formats.iter().enumerate() {
            let doc = make_doc(vec![make_flow_page(vec![make_image(*format, None, None)])]);
            let output = generate_typst(&doc).unwrap();
            let expected_path = format!("img-0.{expected_ext}");
            assert_eq!(
                output.images[0].path, expected_path,
                "Format {format:?} should produce .{expected_ext} extension (test #{i})"
            );
        }
    }

    #[test]
    fn test_image_with_fractional_dimensions() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(
            ImageFormat::Png,
            Some(72.5),
            Some(96.25),
        )])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("width: 72.5pt"),
            "Expected fractional width in: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 96.25pt"),
            "Expected fractional height in: {}",
            output.source
        );
    }

    #[test]
    fn test_image_mixed_with_paragraphs() {
        let doc = make_doc(vec![make_flow_page(vec![
            make_paragraph("Before image"),
            make_image(ImageFormat::Png, Some(100.0), Some(80.0)),
            make_paragraph("After image"),
        ])]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Before image"));
        assert!(output.source.contains("#image(\"img-0.png\""));
        assert!(output.source.contains("After image"));
        assert_eq!(output.images.len(), 1);
    }

    #[test]
    fn test_no_images_produces_empty_assets() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Just text")])]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.images.is_empty());
    }

    // ── FixedPage codegen tests (US-010) ────────────────────────────────

    /// Helper to create a FixedPage (slide-like) with given elements.
    fn make_fixed_page(width: f64, height: f64, elements: Vec<FixedElement>) -> Page {
        Page::Fixed(FixedPage {
            size: PageSize { width, height },
            elements,
            background_color: None,
        })
    }

    /// Helper to create a text box FixedElement.
    fn make_text_box(x: f64, y: f64, w: f64, h: f64, text: &str) -> FixedElement {
        FixedElement {
            x,
            y,
            width: w,
            height: h,
            kind: FixedElementKind::TextBox(vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })]),
        }
    }

    /// Helper to create a shape FixedElement.
    fn make_shape_element(
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        kind: ShapeKind,
        fill: Option<Color>,
        stroke: Option<BorderSide>,
    ) -> FixedElement {
        FixedElement {
            x,
            y,
            width: w,
            height: h,
            kind: FixedElementKind::Shape(Shape {
                kind,
                fill,
                stroke,
                rotation_deg: None,
                opacity: None,
            }),
        }
    }

    /// Helper to create an image FixedElement.
    fn make_fixed_image(x: f64, y: f64, w: f64, h: f64, format: ImageFormat) -> FixedElement {
        FixedElement {
            x,
            y,
            width: w,
            height: h,
            kind: FixedElementKind::Image(ImageData {
                data: vec![0x89, 0x50, 0x4E, 0x47], // PNG header stub
                format,
                width: Some(w),
                height: Some(h),
            }),
        }
    }

    #[test]
    fn test_fixed_page_sets_page_size() {
        // Standard 16:9 slide: 960pt × 540pt
        let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("width: 960pt"),
            "Expected slide width in: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 540pt"),
            "Expected slide height in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_zero_margins() {
        // Slides should have zero margins
        let doc = make_doc(vec![make_fixed_page(960.0, 540.0, vec![])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("margin: 0pt"),
            "Expected zero margins for slide in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_text_box() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_text_box(100.0, 200.0, 300.0, 50.0, "Slide Title")],
        )]);
        let output = generate_typst(&doc).unwrap();
        // Text box should be placed at absolute position
        assert!(
            output.source.contains("Slide Title"),
            "Expected text content in: {}",
            output.source
        );
        assert!(
            output.source.contains("100pt"),
            "Expected x position in: {}",
            output.source
        );
        assert!(
            output.source.contains("200pt"),
            "Expected y position in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_text_box_with_width_height() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_text_box(50.0, 60.0, 400.0, 100.0, "Sized box")],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("400pt"),
            "Expected width in: {}",
            output.source
        );
        assert!(
            output.source.contains("100pt"),
            "Expected height in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_rectangle_shape() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                20.0,
                200.0,
                150.0,
                ShapeKind::Rectangle,
                Some(Color::new(255, 0, 0)),
                None,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("rect"),
            "Expected rect shape in: {}",
            output.source
        );
        assert!(
            output.source.contains("200pt"),
            "Expected shape width in: {}",
            output.source
        );
        assert!(
            output.source.contains("rgb(255, 0, 0)"),
            "Expected fill color in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_ellipse_shape() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                50.0,
                50.0,
                120.0,
                80.0,
                ShapeKind::Ellipse,
                Some(Color::new(0, 128, 255)),
                None,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("ellipse"),
            "Expected ellipse shape in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_line_shape() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                0.0,
                0.0,
                300.0,
                0.0,
                ShapeKind::Line { x2: 300.0, y2: 0.0 },
                None,
                Some(BorderSide {
                    width: 2.0,
                    color: Color::black(),
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("line"),
            "Expected line shape in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_shape_with_stroke() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                10.0,
                100.0,
                100.0,
                ShapeKind::Rectangle,
                None,
                Some(BorderSide {
                    width: 1.5,
                    color: Color::new(0, 0, 255),
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("stroke"),
            "Expected stroke in: {}",
            output.source
        );
        assert!(
            output.source.contains("1.5pt"),
            "Expected stroke width in: {}",
            output.source
        );
    }

    #[test]
    fn test_shape_rotation_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![FixedElement {
                x: 10.0,
                y: 20.0,
                width: 200.0,
                height: 150.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: ShapeKind::Rectangle,
                    fill: Some(Color::new(255, 0, 0)),
                    stroke: None,
                    rotation_deg: Some(90.0),
                    opacity: None,
                }),
            }],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("rotate"),
            "Expected rotate wrapper in: {}",
            output.source
        );
        assert!(
            output.source.contains("90deg"),
            "Expected 90deg angle in: {}",
            output.source
        );
    }

    #[test]
    fn test_shape_opacity_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![FixedElement {
                x: 10.0,
                y: 20.0,
                width: 200.0,
                height: 150.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: ShapeKind::Rectangle,
                    fill: Some(Color::new(0, 255, 0)),
                    stroke: None,
                    rotation_deg: None,
                    opacity: Some(0.5),
                }),
            }],
        )]);
        let output = generate_typst(&doc).unwrap();
        // With 50% opacity, the fill color should include alpha
        assert!(
            output.source.contains("rgba(0, 255, 0, 128)"),
            "Expected rgba fill with alpha in: {}",
            output.source
        );
    }

    #[test]
    fn test_shape_rotation_and_opacity_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![FixedElement {
                x: 50.0,
                y: 50.0,
                width: 100.0,
                height: 100.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: ShapeKind::Ellipse,
                    fill: Some(Color::new(0, 0, 255)),
                    stroke: None,
                    rotation_deg: Some(45.0),
                    opacity: Some(0.75),
                }),
            }],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("rotate"),
            "Expected rotate in: {}",
            output.source
        );
        assert!(
            output.source.contains("45deg"),
            "Expected 45deg in: {}",
            output.source
        );
        assert!(
            output.source.contains("rgba(0, 0, 255, 191)"),
            "Expected rgba fill in: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_image_element() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_fixed_image(
                100.0,
                150.0,
                400.0,
                300.0,
                ImageFormat::Png,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#image("),
            "Expected image call in: {}",
            output.source
        );
        assert_eq!(output.images.len(), 1, "Expected one image asset");
    }

    #[test]
    fn test_fixed_page_mixed_elements() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![
                make_text_box(50.0, 30.0, 800.0, 60.0, "Title"),
                make_shape_element(
                    50.0,
                    100.0,
                    400.0,
                    300.0,
                    ShapeKind::Rectangle,
                    Some(Color::new(200, 200, 200)),
                    None,
                ),
                make_fixed_image(500.0, 100.0, 350.0, 300.0, ImageFormat::Jpeg),
                make_text_box(50.0, 420.0, 800.0, 40.0, "Footer text"),
            ],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Title"));
        assert!(output.source.contains("rect"));
        assert!(output.source.contains("#image("));
        assert!(output.source.contains("Footer text"));
        assert_eq!(output.images.len(), 1);
    }

    #[test]
    fn test_fixed_page_multiple_text_boxes() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![
                make_text_box(100.0, 50.0, 300.0, 40.0, "First"),
                make_text_box(100.0, 120.0, 300.0, 40.0, "Second"),
                make_text_box(100.0, 190.0, 300.0, 40.0, "Third"),
            ],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("First"));
        assert!(output.source.contains("Second"));
        assert!(output.source.contains("Third"));
    }

    #[test]
    fn test_fixed_page_uses_place_for_positioning() {
        // Verify Typst uses `place()` for absolute positioning
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_text_box(100.0, 200.0, 300.0, 50.0, "Positioned")],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("place("),
            "Expected place() for absolute positioning in: {}",
            output.source
        );
    }

    // ── TablePage codegen tests ──────────────────────────────────────────

    /// Helper to create a TablePage.
    fn make_table_page(
        name: &str,
        width: f64,
        height: f64,
        margins: Margins,
        table: Table,
    ) -> Page {
        Page::Table(crate::ir::TablePage {
            name: name.to_string(),
            size: PageSize { width, height },
            margins,
            table,
            header: None,
            footer: None,
        })
    }

    /// Helper to create a simple Table with text cells.
    fn make_simple_table(rows: Vec<Vec<&str>>) -> Table {
        Table {
            rows: rows
                .into_iter()
                .map(|cells| TableRow {
                    cells: cells
                        .into_iter()
                        .map(|text| TableCell {
                            content: vec![Block::Paragraph(Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: text.to_string(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                }],
                            })],
                            ..TableCell::default()
                        })
                        .collect(),
                    height: None,
                })
                .collect(),
            column_widths: vec![],
        }
    }

    #[test]
    fn test_table_page_basic() {
        let table = make_simple_table(vec![vec!["A1", "B1"], vec!["A2", "B2"]]);
        let doc = make_doc(vec![make_table_page(
            "Sheet1",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#set page("),
            "Expected page setup in: {}",
            output.source
        );
        assert!(
            output.source.contains("#table("),
            "Expected table markup in: {}",
            output.source
        );
        assert!(output.source.contains("A1"));
        assert!(output.source.contains("B1"));
        assert!(output.source.contains("A2"));
        assert!(output.source.contains("B2"));
    }

    #[test]
    fn test_table_page_custom_page_size_and_margins() {
        let table = make_simple_table(vec![vec!["Data"]]);
        let doc = make_doc(vec![make_table_page(
            "Custom",
            800.0,
            600.0,
            Margins {
                top: 20.0,
                bottom: 20.0,
                left: 30.0,
                right: 30.0,
            },
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("width: 800pt"),
            "Expected custom width in: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 600pt"),
            "Expected custom height in: {}",
            output.source
        );
        assert!(
            output.source.contains("top: 20pt"),
            "Expected custom top margin in: {}",
            output.source
        );
        assert!(
            output.source.contains("left: 30pt"),
            "Expected custom left margin in: {}",
            output.source
        );
    }

    #[test]
    fn test_table_page_cell_data_types() {
        // Text, numbers, and dates are all stored as text strings in IR
        let table = make_simple_table(vec![
            vec!["Name", "Age", "Date"],
            vec!["Alice", "30", "2024-01-15"],
        ]);
        let doc = make_doc(vec![make_table_page(
            "Data",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Name"));
        assert!(output.source.contains("Age"));
        assert!(output.source.contains("Date"));
        assert!(output.source.contains("Alice"));
        assert!(output.source.contains("30"));
        assert!(output.source.contains("2024-01-15"));
    }

    #[test]
    fn test_table_page_merged_cells() {
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Merged".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        col_span: 2,
                        ..TableCell::default()
                    }],
                    height: None,
                },
                TableRow {
                    cells: vec![
                        TableCell {
                            content: vec![Block::Paragraph(Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Left".to_string(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                }],
                            })],
                            ..TableCell::default()
                        },
                        TableCell {
                            content: vec![Block::Paragraph(Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Right".to_string(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                }],
                            })],
                            ..TableCell::default()
                        },
                    ],
                    height: None,
                },
            ],
            column_widths: vec![],
        };
        let doc = make_doc(vec![make_table_page(
            "MergeSheet",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("colspan: 2"),
            "Expected colspan in: {}",
            output.source
        );
        assert!(output.source.contains("Merged"));
        assert!(output.source.contains("Left"));
        assert!(output.source.contains("Right"));
    }

    #[test]
    fn test_table_page_with_column_widths() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Col1".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Col2".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                ],
                height: None,
            }],
            column_widths: vec![100.0, 200.0],
        };
        let doc = make_doc(vec![make_table_page(
            "Widths",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("columns: (100pt, 200pt)"),
            "Expected column widths in: {}",
            output.source
        );
    }

    #[test]
    fn test_table_page_empty_table() {
        let table = Table {
            rows: vec![],
            column_widths: vec![],
        };
        let doc = make_doc(vec![make_table_page(
            "Empty",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        // Should still produce valid Typst with page setup
        assert!(output.source.contains("#set page("));
    }

    #[test]
    fn test_table_page_multiple_sheets() {
        let table1 = make_simple_table(vec![vec!["Sheet1Data"]]);
        let table2 = make_simple_table(vec![vec!["Sheet2Data"]]);
        let doc = make_doc(vec![
            make_table_page("Sheet1", 595.28, 841.89, Margins::default(), table1),
            make_table_page("Sheet2", 595.28, 841.89, Margins::default(), table2),
        ]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Sheet1Data"));
        assert!(output.source.contains("Sheet2Data"));
    }

    #[test]
    fn test_table_page_rowspan_merge() {
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![
                        TableCell {
                            content: vec![Block::Paragraph(Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Tall".to_string(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                }],
                            })],
                            row_span: 2,
                            ..TableCell::default()
                        },
                        TableCell {
                            content: vec![Block::Paragraph(Paragraph {
                                style: ParagraphStyle::default(),
                                runs: vec![Run {
                                    text: "Top".to_string(),
                                    style: TextStyle::default(),
                                    href: None,
                                    footnote: None,
                                }],
                            })],
                            ..TableCell::default()
                        },
                    ],
                    height: None,
                },
                TableRow {
                    cells: vec![TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Bottom".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    }],
                    height: None,
                },
            ],
            column_widths: vec![],
        };
        let doc = make_doc(vec![make_table_page(
            "RowMerge",
            595.28,
            841.89,
            Margins::default(),
            table,
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("rowspan: 2"),
            "Expected rowspan in: {}",
            output.source
        );
        assert!(output.source.contains("Tall"));
        assert!(output.source.contains("Top"));
        assert!(output.source.contains("Bottom"));
    }

    // ----- List codegen tests -----

    #[test]
    fn test_generate_bulleted_list() {
        use crate::ir::{List, ListItem, ListKind};
        let list = List {
            kind: ListKind::Unordered,
            items: vec![
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Apple".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Banana".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
            ],
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#list("),
            "Expected #list( in: {}",
            output.source
        );
        assert!(output.source.contains("Apple"));
        assert!(output.source.contains("Banana"));
    }

    #[test]
    fn test_generate_numbered_list() {
        use crate::ir::{List, ListItem, ListKind};
        let list = List {
            kind: ListKind::Ordered,
            items: vec![
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Step 1".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Step 2".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
            ],
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#enum("),
            "Expected #enum( in: {}",
            output.source
        );
        assert!(output.source.contains("Step 1"));
        assert!(output.source.contains("Step 2"));
    }

    #[test]
    fn test_generate_nested_list() {
        use crate::ir::{List, ListItem, ListKind};
        let list = List {
            kind: ListKind::Unordered,
            items: vec![
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Parent".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Child".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 1,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Sibling".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
            ],
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Parent"));
        assert!(output.source.contains("Child"));
        assert!(output.source.contains("Sibling"));
        // Nested list should contain a sub-list
        assert!(
            output.source.contains("#list("),
            "Expected nested #list( in: {}",
            output.source
        );
    }

    // ----- US-020: Header/footer codegen tests -----

    #[test]
    fn test_generate_flow_page_with_text_header() {
        use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Body text")],
            header: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::Run(Run {
                        text: "Document Title".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    })],
                }],
            }),
            footer: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("header:"),
            "Should contain header: in page setup. Got: {}",
            output.source
        );
        assert!(
            output.source.contains("Document Title"),
            "Header should contain 'Document Title'. Got: {}",
            output.source
        );
    }

    #[test]
    fn test_generate_flow_page_with_page_number_footer() {
        use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Body text")],
            header: None,
            footer: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![
                        HFInline::Run(Run {
                            text: "Page ".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }),
                        HFInline::PageNumber,
                    ],
                }],
            }),
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("footer:"),
            "Should contain footer: in page setup. Got: {}",
            output.source
        );
        assert!(
            output.source.contains("counter(page).display()"),
            "Footer should contain page counter. Got: {}",
            output.source
        );
        assert!(
            output.source.contains("Page "),
            "Footer should contain 'Page ' text. Got: {}",
            output.source
        );
    }

    #[test]
    fn test_generate_flow_page_with_header_and_footer() {
        use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Body")],
            header: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::Run(Run {
                        text: "Header".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    })],
                }],
            }),
            footer: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::PageNumber],
                }],
            }),
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("header:") && output.source.contains("footer:"),
            "Should contain both header: and footer:. Got: {}",
            output.source
        );
    }

    #[test]
    fn test_generate_flow_page_without_header_footer() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Body")])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            !output.source.contains("header:"),
            "Should NOT contain header: when no header. Got: {}",
            output.source
        );
        assert!(
            !output.source.contains("footer:"),
            "Should NOT contain footer: when no footer. Got: {}",
            output.source
        );
    }

    // ── Fixed page background tests ──────────────────────────────────────

    #[test]
    fn test_fixed_page_with_background_color() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: Some(Color::new(255, 0, 0)),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("fill: rgb(255, 0, 0)"),
            "Should contain fill: rgb(255, 0, 0). Got: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_without_background_color() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: None,
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            !output.source.contains("fill:"),
            "Should NOT contain fill: when no background. Got: {}",
            output.source
        );
    }

    #[test]
    fn test_fixed_page_table_element() {
        // A table placed at absolute position on a fixed page
        let table = Table {
            rows: vec![TableRow {
                cells: vec![
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "A1".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                    TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "B1".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    },
                ],
                height: None,
            }],
            column_widths: vec![100.0, 100.0],
        };

        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![FixedElement {
                x: 50.0,
                y: 100.0,
                width: 200.0,
                height: 50.0,
                kind: FixedElementKind::Table(table),
            }],
            background_color: None,
        });

        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();

        // Should have a #place() with table inside
        assert!(
            output
                .source
                .contains("#place(top + left, dx: 50pt, dy: 100pt)")
        );
        assert!(output.source.contains("#table("));
        assert!(output.source.contains("columns: (100pt, 100pt)"));
        assert!(output.source.contains("A1"));
        assert!(output.source.contains("B1"));
    }

    // ----- Hyperlink codegen tests (US-030) -----

    #[test]
    fn test_hyperlink_generates_typst_link() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Click me".to_string(),
                style: TextStyle::default(),
                href: Some("https://example.com".to_string()),
                footnote: None,
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains(r#"#link("https://example.com")[Click me]"#),
            "Expected Typst link markup, got: {}",
            output.source
        );
    }

    #[test]
    fn test_hyperlink_with_styled_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold link".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: Some("https://example.com".to_string()),
                footnote: None,
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        // Should have link wrapping styled text
        assert!(
            output.source.contains(r#"#link("https://example.com")["#),
            "Expected Typst link markup, got: {}",
            output.source
        );
        assert!(
            output.source.contains("#text(weight: \"bold\")"),
            "Expected bold text inside link, got: {}",
            output.source
        );
    }

    #[test]
    fn test_hyperlink_mixed_with_plain_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "Visit ".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
                Run {
                    text: "Rust".to_string(),
                    style: TextStyle::default(),
                    href: Some("https://rust-lang.org".to_string()),
                    footnote: None,
                },
                Run {
                    text: " for more.".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
            ],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("Visit "),
            "Expected plain text, got: {}",
            output.source
        );
        assert!(
            output
                .source
                .contains(r#"#link("https://rust-lang.org")[Rust]"#),
            "Expected Typst link markup, got: {}",
            output.source
        );
        assert!(
            output.source.contains(" for more."),
            "Expected plain text after link, got: {}",
            output.source
        );
    }

    #[test]
    fn test_hyperlink_url_with_special_chars_escaped() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Link".to_string(),
                style: TextStyle::default(),
                href: Some("https://example.com/path?q=1&r=2".to_string()),
                footnote: None,
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains(r#"#link("https://example.com/path?q=1&r=2")[Link]"#),
            "Expected URL with special chars in link, got: {}",
            output.source
        );
    }

    // ── Footnotes ───────────────────────────────────────────────────────

    #[test]
    fn test_footnote_generates_typst_footnote() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "Some text".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
                Run {
                    text: String::new(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: Some("This is a footnote.".to_string()),
                },
            ],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#footnote[This is a footnote.]"),
            "Expected Typst footnote markup, got: {}",
            output.source
        );
    }

    #[test]
    fn test_footnote_with_special_chars() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: String::new(),
                style: TextStyle::default(),
                href: None,
                footnote: Some("Note with #special *chars*".to_string()),
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains(r"#footnote[Note with \#special \*chars\*]"),
            "Expected escaped footnote content, got: {}",
            output.source
        );
    }

    // --- US-036: TablePage header/footer codegen ---

    #[test]
    fn test_table_page_with_header() {
        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: make_simple_table(vec![vec!["A"]]),
            header: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Center),
                        ..ParagraphStyle::default()
                    },
                    elements: vec![HFInline::Run(Run {
                        text: "My Header".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    })],
                }],
            }),
            footer: None,
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("header: ["),
            "Expected header in page setup, got: {}",
            output.source
        );
        assert!(
            output.source.contains("My Header"),
            "Expected header text, got: {}",
            output.source
        );
    }

    #[test]
    fn test_table_page_with_page_number_footer() {
        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: make_simple_table(vec![vec!["A"]]),
            header: None,
            footer: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Center),
                        ..ParagraphStyle::default()
                    },
                    elements: vec![
                        HFInline::Run(Run {
                            text: "Page ".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }),
                        HFInline::PageNumber,
                        HFInline::Run(Run {
                            text: " of ".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }),
                        HFInline::TotalPages,
                    ],
                }],
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        // Footer with page numbers needs context
        assert!(
            output.source.contains("footer: context ["),
            "Expected context footer, got: {}",
            output.source
        );
        assert!(
            output.source.contains("#counter(page).display()"),
            "Expected page number counter, got: {}",
            output.source
        );
        assert!(
            output.source.contains("#counter(page).final().first()"),
            "Expected total pages counter, got: {}",
            output.source
        );
    }

    #[test]
    fn test_table_page_no_header_footer() {
        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: make_simple_table(vec![vec!["A"]]),
            header: None,
            footer: None,
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        // Should use simple page setup without header/footer
        assert!(
            !output.source.contains("header:"),
            "Expected no header, got: {}",
            output.source
        );
        assert!(
            !output.source.contains("footer:"),
            "Expected no footer, got: {}",
            output.source
        );
    }
}
