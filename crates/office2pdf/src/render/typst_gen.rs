use std::fmt::Write;
use std::io::Cursor;

use image::{GenericImageView, ImageFormat as RasterImageFormat};

use crate::config::ConvertOptions;
use crate::error::ConvertError;
use crate::ir::{
    Alignment, ArrowHead, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Chart,
    ChartType, Color, ColumnLayout, Document, FixedElement, FixedElementKind, FixedPage,
    FloatingImage, FloatingShape, FloatingTextBox, FlowPage, FrameAnchor, GradientFill, HFInline,
    HeaderFooter, HeaderFooterFrame, ImageCrop, ImageData, ImageFormat, Insets, LineSpacing, List,
    ListKind, Margins, MathEquation, Metadata, Page, PageSize, Paragraph, ParagraphStyle,
    PositionedTabAlignment, PositionedTabRelativeTo, Run, Shadow, Shape, ShapeKind, SheetPage,
    SmartArt, TabAlignment, TabLeader, TabStop, Table, TableCell, TableRow, TextBoxData,
    TextBoxVerticalAlign, TextDirection, TextStyle, VerticalTextAlign, WrapMode,
};

use self::diagrams::{generate_chart, generate_smartart};
use self::lists::{
    can_render_fixed_text_list_inline, common_text_style, generate_fixed_text_list, generate_list,
    write_common_text_settings, write_fixed_text_default_par_settings,
};
use self::shapes::{
    generate_shape, write_fill_color, write_gradient_fill, write_shape_stroke,
    write_text_box_shape_background,
};
use self::tables::generate_table;
use self::text::*;
use super::font_context::FontSearchContext;

#[path = "typst_gen_diagrams.rs"]
mod diagrams;
#[path = "typst_gen_lists.rs"]
mod lists;
#[path = "typst_gen_shapes.rs"]
mod shapes;
#[path = "typst_gen_tables.rs"]
mod tables;
#[path = "typst_gen_text.rs"]
mod text;

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

/// Maximum nesting depth for tables-within-tables, matching the parser limit.
const MAX_TABLE_DEPTH: usize = 64;
/// Typst's line box leaves more top leading than Word/LibreOffice text frames.
const FLOATING_TEXT_BOX_TOP_LEADING_COMPENSATION_PT: f64 = 6.0;

/// Internal context for tracking image assets during code generation.
struct GenCtx {
    images: Vec<ImageAsset>,
    next_image_id: usize,
    next_text_box_id: usize,
    table_depth: usize,
}

impl GenCtx {
    fn new() -> Self {
        Self {
            images: Vec::new(),
            next_image_id: 0,
            next_text_box_id: 0,
            table_depth: 0,
        }
    }

    fn add_image(&mut self, image: &ImageData) -> String {
        let (data, format) = preprocess_image_asset(image);
        let ext = format.extension();
        let id = self.next_image_id;
        self.next_image_id += 1;
        let path = format!("img-{id}.{ext}");
        self.images.push(ImageAsset {
            path: path.clone(),
            data,
        });
        path
    }

    fn next_text_box_id(&mut self) -> usize {
        let id = self.next_text_box_id;
        self.next_text_box_id += 1;
        id
    }
}

fn raster_image_format(format: ImageFormat) -> Option<RasterImageFormat> {
    match format {
        ImageFormat::Png => Some(RasterImageFormat::Png),
        ImageFormat::Jpeg => Some(RasterImageFormat::Jpeg),
        ImageFormat::Gif => Some(RasterImageFormat::Gif),
        ImageFormat::Bmp => Some(RasterImageFormat::Bmp),
        ImageFormat::Tiff => Some(RasterImageFormat::Tiff),
        ImageFormat::Svg => None,
    }
}

fn crop_to_pixels(crop: ImageCrop, width: u32, height: u32) -> Option<(u32, u32, u32, u32)> {
    let left = ((crop.left.clamp(0.0, 1.0) * width as f64).round() as u32).min(width);
    let top = ((crop.top.clamp(0.0, 1.0) * height as f64).round() as u32).min(height);
    let right = ((crop.right.clamp(0.0, 1.0) * width as f64).round() as u32).min(width);
    let bottom = ((crop.bottom.clamp(0.0, 1.0) * height as f64).round() as u32).min(height);
    if left + right >= width || top + bottom >= height {
        return None;
    }
    Some((left, top, width - left - right, height - top - bottom))
}

fn preprocess_image_asset(image: &ImageData) -> (Vec<u8>, ImageFormat) {
    let Some(crop) = image.crop.filter(|crop| !crop.is_empty()) else {
        return (image.data.clone(), image.format);
    };
    let Some(raster_format) = raster_image_format(image.format) else {
        return (image.data.clone(), image.format);
    };
    let Ok(decoded) = image::load_from_memory_with_format(&image.data, raster_format) else {
        return (image.data.clone(), image.format);
    };
    let (width, height) = decoded.dimensions();
    let Some((left, top, crop_width, crop_height)) = crop_to_pixels(crop, width, height) else {
        return (image.data.clone(), image.format);
    };

    let cropped = decoded.crop_imm(left, top, crop_width, crop_height);
    let mut encoded = Cursor::new(Vec::new());
    if cropped
        .write_to(&mut encoded, RasterImageFormat::Png)
        .is_ok()
    {
        (encoded.into_inner(), ImageFormat::Png)
    } else {
        (image.data.clone(), image.format)
    }
}

/// Resolve the effective page size, applying paper_size and landscape overrides.
fn resolve_page_size(original: &PageSize, options: &ConvertOptions) -> PageSize {
    let (mut w, mut h) = if let Some(ref ps) = options.paper_size {
        let (pw, ph) = ps.dimensions();
        (pw, ph)
    } else {
        (original.width, original.height)
    };

    if let Some(landscape) = options.landscape {
        let needs_swap = (landscape && w < h) || (!landscape && w > h);
        if needs_swap {
            std::mem::swap(&mut w, &mut h);
        }
    }

    PageSize {
        width: w,
        height: h,
    }
}

/// Emit `#set document(title: ..., author: ..., date: ...)` if metadata is present.
fn generate_document_metadata(out: &mut String, metadata: &Metadata) {
    let has_title = metadata.title.is_some();
    let has_author = metadata.author.is_some();
    let parsed_date = metadata.created.as_deref().and_then(parse_iso8601_date);
    if !has_title && !has_author && parsed_date.is_none() {
        return;
    }

    out.push_str("#set document(");
    let mut first = true;
    if let Some(ref title) = metadata.title {
        let _ = write!(out, "title: \"{}\"", escape_typst_string(title));
        first = false;
    }
    if let Some(ref author) = metadata.author {
        if !first {
            out.push_str(", ");
        }
        let _ = write!(out, "author: \"{}\"", escape_typst_string(author));
        first = false;
    }
    if let Some((year, month, day, hour, minute, second)) = parsed_date {
        if !first {
            out.push_str(", ");
        }
        let _ = write!(
            out,
            "date: datetime(year: {year}, month: {month}, day: {day}, \
             hour: {hour}, minute: {minute}, second: {second})"
        );
    }
    out.push_str(")\n");
}

/// Parse an ISO 8601 date string (e.g. `2024-06-15T10:30:00Z`) into components.
///
/// Returns `(year, month, day, hour, minute, second)` or `None` if unparseable.
fn parse_iso8601_date(s: &str) -> Option<(i32, u8, u8, u8, u8, u8)> {
    let s = s.trim();
    if s.len() < 10 {
        return None;
    }
    let year: i32 = s.get(0..4)?.parse().ok()?;
    if s.as_bytes().get(4)? != &b'-' {
        return None;
    }
    let month: u8 = s.get(5..7)?.parse().ok()?;
    if s.as_bytes().get(7)? != &b'-' {
        return None;
    }
    let day: u8 = s.get(8..10)?.parse().ok()?;

    // Validate ranges
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }

    if s.len() >= 19 && s.as_bytes().get(10) == Some(&b'T') {
        let hour: u8 = s.get(11..13)?.parse().ok()?;
        let minute: u8 = s.get(14..16)?.parse().ok()?;
        let second: u8 = s.get(17..19)?.parse().ok()?;
        Some((year, month, day, hour, minute, second))
    } else {
        Some((year, month, day, 0, 0, 0))
    }
}

/// Escape a string for use inside Typst double quotes.
fn escape_typst_string(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"")
}

/// Generate Typst markup from a Document IR.
pub fn generate_typst(doc: &Document) -> Result<TypstOutput, ConvertError> {
    generate_typst_with_options_and_font_context(doc, &ConvertOptions::default(), None)
}

/// Generate Typst markup from a Document IR with conversion options.
///
/// When `options.paper_size` is set, all pages use the specified paper size.
/// When `options.landscape` is set, page orientation is forced.
pub fn generate_typst_with_options(
    doc: &Document,
    options: &ConvertOptions,
) -> Result<TypstOutput, ConvertError> {
    generate_typst_with_options_and_font_context(doc, options, None)
}

pub(crate) fn generate_typst_with_options_and_font_context(
    doc: &Document,
    options: &ConvertOptions,
    font_context: Option<&FontSearchContext>,
) -> Result<TypstOutput, ConvertError> {
    super::font_subst::with_font_search_context(font_context, || {
        // Pre-allocate output string: ~2KB per page is a reasonable estimate
        let mut out = String::with_capacity(doc.pages.len() * 2048);

        // Emit document metadata (title/author) if present
        generate_document_metadata(&mut out, &doc.metadata);

        let mut ctx = GenCtx::new();
        for (index, page) in doc.pages.iter().enumerate() {
            if index > 0 {
                out.push_str("\n#pagebreak()\n");
            }
            match page {
                Page::Flow(flow) => generate_flow_page(&mut out, flow, &mut ctx, options)?,
                Page::Fixed(fixed) => generate_fixed_page(&mut out, fixed, &mut ctx, options)?,
                Page::Sheet(sheet_page) => {
                    generate_table_page(&mut out, sheet_page, &mut ctx, options)?;
                }
            }
        }
        Ok(TypstOutput {
            source: out,
            images: ctx.images,
        })
    })
}

fn generate_flow_page(
    out: &mut String,
    page: &FlowPage,
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    write_flow_page_setup(out, page, &size, ctx);
    out.push('\n');

    if let Some(ref cols) = page.columns {
        generate_flow_page_columns(out, &page.content, cols, ctx)?;
    } else {
        generate_blocks(out, &page.content, ctx)?;
    }
    Ok(())
}

/// Generate Typst markup for multi-column content.
///
/// Equal columns use `#columns(n, gutter: Xpt)[content]`.
/// Unequal columns use `#grid(columns: (W1pt, W2pt, ...), gutter: Xpt)` with
/// content split by `ColumnBreak` blocks into separate grid cells.
fn generate_flow_page_columns(
    out: &mut String,
    content: &[Block],
    cols: &ColumnLayout,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    if let Some(ref widths) = cols.column_widths {
        // Unequal columns: use grid with explicit column widths.
        // Split content at ColumnBreak boundaries.
        let _ = write!(out, "#grid(columns: (");
        for (i, w) in widths.iter().enumerate() {
            if i > 0 {
                out.push_str(", ");
            }
            let _ = write!(out, "{}pt", format_f64(*w));
        }
        let _ = write!(out, "), gutter: {}pt", format_f64(cols.spacing));
        out.push_str(")\n");

        // Split content by ColumnBreak into grid cells
        let segments = split_at_column_breaks(content);
        for segment in &segments {
            out.push('[');
            for (i, block) in segment.iter().enumerate() {
                if i > 0 {
                    out.push('\n');
                }
                generate_block(out, block, ctx)?;
            }
            out.push(']');
        }
        out.push('\n');
    } else {
        // Equal columns: use Typst columns()
        let _ = writeln!(
            out,
            "#columns({}, gutter: {}pt)[",
            cols.num_columns,
            format_f64(cols.spacing)
        );
        generate_blocks(out, content, ctx)?;
        out.push_str("\n]\n");
    }
    Ok(())
}

/// Split content blocks at ColumnBreak boundaries into segments.
fn split_at_column_breaks(content: &[Block]) -> Vec<Vec<&Block>> {
    let mut segments: Vec<Vec<&Block>> = vec![vec![]];
    for block in content {
        if matches!(block, Block::ColumnBreak) {
            segments.push(vec![]);
        } else if let Some(last) = segments.last_mut() {
            last.push(block);
        }
    }
    segments
}

fn generate_fixed_page(
    out: &mut String,
    page: &FixedPage,
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    // Slides use zero margins — all positioning is absolute
    if let Some(ref gradient) = page.background_gradient {
        let _ = write!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt, fill: ",
            format_f64(size.width),
            format_f64(size.height),
        );
        write_gradient_fill(out, gradient);
        let _ = writeln!(out, ")");
    } else if let Some(ref bg) = page.background_color {
        let _ = writeln!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt, fill: rgb({}, {}, {}))",
            format_f64(size.width),
            format_f64(size.height),
            bg.r,
            bg.g,
            bg.b,
        );
    } else {
        let _ = writeln!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt)",
            format_f64(size.width),
            format_f64(size.height),
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
    page: &SheetPage,
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    write_table_page_setup(out, page, &size, ctx);
    out.push('\n');

    if page.charts.is_empty() {
        generate_table(out, &page.table, ctx)?;
    } else {
        generate_table_with_charts(out, &page.table, &page.charts, ctx)?;
    }
    Ok(())
}

/// Render a table interleaved with charts at their anchor positions.
/// Splits the table into segments at chart anchor rows and emits charts between segments.
fn generate_table_with_charts(
    out: &mut String,
    table: &Table,
    charts: &[(u32, Chart)],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    use crate::ir::Table;

    // Sort charts by anchor row (should already be sorted, but ensure)
    let mut sorted_charts: Vec<&(u32, Chart)> = charts.iter().collect();
    sorted_charts.sort_by_key(|(row, _)| *row);

    let total_rows = table.rows.len();
    let mut row_start = 0usize;
    let mut chart_idx = 0;

    // Walk through rows and emit table segments + charts
    for row_end in 0..total_rows {
        let row_num = (row_end + 1) as u32; // 1-indexed row number

        // Emit all charts anchored at or before this row
        while chart_idx < sorted_charts.len() && sorted_charts[chart_idx].0 <= row_num {
            // Emit table segment up to and including this row
            if row_start <= row_end {
                let segment = Table {
                    rows: table.rows[row_start..=row_end].to_vec(),
                    column_widths: table.column_widths.clone(),
                    header_row_count: if row_start == 0 {
                        table.header_row_count.min(row_end + 1)
                    } else {
                        0
                    },
                    alignment: table.alignment,
                    default_cell_padding: table.default_cell_padding,
                    use_content_driven_row_heights: table.use_content_driven_row_heights,
                };
                generate_table(out, &segment, ctx)?;
                out.push('\n');
                row_start = row_end + 1;
            }
            // Emit the chart
            generate_chart(out, &sorted_charts[chart_idx].1);
            out.push('\n');
            chart_idx += 1;
        }
    }

    // Emit remaining rows after last chart
    if row_start < total_rows {
        let segment = Table {
            rows: table.rows[row_start..].to_vec(),
            column_widths: table.column_widths.clone(),
            header_row_count: if row_start == 0 {
                table.header_row_count.min(total_rows - row_start)
            } else {
                0
            },
            alignment: table.alignment,
            default_cell_padding: table.default_cell_padding,
            use_content_driven_row_heights: table.use_content_driven_row_heights,
        };
        generate_table(out, &segment, ctx)?;
        out.push('\n');
    }

    // Emit any remaining charts (anchored beyond last row, e.g., u32::MAX)
    while chart_idx < sorted_charts.len() {
        generate_chart(out, &sorted_charts[chart_idx].1);
        out.push('\n');
        chart_idx += 1;
    }

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
        FixedElementKind::TextBox(text_box) => generate_fixed_text_box(out, elem, text_box, ctx)?,
        FixedElementKind::Image(img) => {
            generate_image(out, img, ctx);
            // Render image border as a separate overlay so that #image()
            // dimensions are not affected by Typst's #box(stroke:) sizing.
            if let Some(ref stroke) = img.stroke {
                let _ = write!(
                    out,
                    "]\n#place(top + left, dx: {}pt, dy: {}pt)[\n",
                    format_f64(elem.x),
                    format_f64(elem.y),
                );
                let _ = write!(
                    out,
                    "#rect(width: {}pt, height: {}pt, fill: none, stroke: ",
                    format_f64(elem.width),
                    format_f64(elem.height),
                );
                shapes::write_image_border_stroke(out, stroke);
                out.push_str(")\n");
            }
        }
        FixedElementKind::Shape(shape) => {
            generate_shape(out, shape, elem.width, elem.height);
        }
        FixedElementKind::Table(table) => {
            generate_table(out, table, ctx)?;
        }
        FixedElementKind::SmartArt(smartart) => {
            generate_smartart(out, smartart, elem.width, elem.height);
        }
        FixedElementKind::Chart(chart) => {
            generate_chart(out, chart);
        }
    }

    out.push_str("]\n");
    Ok(())
}

fn generate_fixed_text_box(
    out: &mut String,
    elem: &FixedElement,
    text_box: &TextBoxData,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let outer_width_pt: f64 = elem.width.max(0.0);
    let outer_height_pt: f64 = elem.height.max(0.0);
    let inner_width_pt: f64 =
        (outer_width_pt - text_box.padding.left - text_box.padding.right).max(0.0);
    let inner_height_pt: f64 =
        (outer_height_pt - text_box.padding.top - text_box.padding.bottom).max(0.0);
    let text_box_id: usize = ctx.next_text_box_id();

    let has_custom_shape: bool = text_box.shape_kind.is_some();

    let _ = write!(
        out,
        "#block(width: {}pt, height: {}pt, inset: {}",
        format_f64(outer_width_pt),
        format_f64(outer_height_pt),
        format_insets(&text_box.padding),
    );
    if text_box.no_wrap {
        out.push_str(", clip: false");
    }
    // For non-rectangular shapes, render fill/stroke as a placed background shape.
    if has_custom_shape {
        // Transparent outer block — shape background is placed inside.
    } else {
        if let Some(fill) = &text_box.fill {
            write_fill_color(out, fill, text_box.opacity);
        }
        write_shape_stroke(out, &text_box.stroke);
    }
    out.push_str(")[\n");

    // Render non-rectangular shape background via #place overlay.
    if let Some(ref shape_kind) = text_box.shape_kind {
        write_text_box_shape_background(
            out,
            shape_kind,
            outer_width_pt,
            outer_height_pt,
            &text_box.padding,
            text_box.fill.as_ref(),
            text_box.opacity,
            &text_box.stroke,
        );
    }
    if let Some(paragraph) = single_line_fit_paragraph(text_box, inner_height_pt) {
        let mut raw_paragraph: Paragraph = paragraph.clone();
        raw_paragraph.style.alignment = None;
        let estimated_line_height_pt: f64 = estimate_single_line_height_pt(paragraph);

        let _ = writeln!(out, "  #let text_box_raw_{text_box_id} = [");
        out.push_str("  ");
        // The measured raw paragraph must stay unbreakable through Typst layout,
        // otherwise mixed-font headers can reflow again inside the scaled box.
        generate_fixed_text_paragraph(out, &raw_paragraph, true)?;
        out.push_str("  ]\n");

        let _ = writeln!(out, "  #let text_box_content_{text_box_id} = context {{");
        let _ = writeln!(
            out,
            "    let text_box_scale_width_{text_box_id} = ({}pt / calc.max(measure(text_box_raw_{text_box_id}).width, 1pt)) * 100%",
            format_f64(inner_width_pt),
        );
        let _ = writeln!(
            out,
            "    let text_box_scale_height_{text_box_id} = ({}pt / {}pt) * 100%",
            format_f64(inner_height_pt),
            format_f64(estimated_line_height_pt.max(1.0)),
        );
        let _ = writeln!(
            out,
            "    let text_box_scale_{text_box_id} = calc.min(100%, calc.min(text_box_scale_width_{text_box_id}, text_box_scale_height_{text_box_id}))",
        );
        let _ = writeln!(out, "    box(width: {}pt)[", format_f64(inner_width_pt),);
        if let Some(align_str) = fixed_text_box_alignment_name(paragraph.style.alignment) {
            let _ = writeln!(out, "      #align({align_str})[");
        }
        let _ = writeln!(
            out,
            "        #scale(x: text_box_scale_{text_box_id}, y: text_box_scale_{text_box_id}, origin: top + left, reflow: true)["
        );
        let _ = writeln!(out, "          #text_box_raw_{text_box_id}");
        out.push_str("        ]\n");
        if fixed_text_box_alignment_name(paragraph.style.alignment).is_some() {
            out.push_str("      ]\n");
        }
        out.push_str("    ]\n");
        out.push_str("  }\n");
    } else if let Some(paragraph) = wrapped_fit_paragraph(text_box) {
        let _ = writeln!(
            out,
            "  #let text_box_raw_{text_box_id} = block(width: {}pt)[",
            format_f64(inner_width_pt),
        );
        out.push_str("  ");
        generate_fixed_text_paragraph(out, paragraph, false)?;
        out.push_str("  ]\n");

        let _ = writeln!(out, "  #let text_box_content_{text_box_id} = context {{");
        let _ = writeln!(
            out,
            "    let text_box_scale_{text_box_id} = calc.min(100%, ({}pt / calc.max(measure(text_box_raw_{text_box_id}).height, 1pt)) * 100%)",
            format_f64(inner_height_pt),
        );
        let _ = writeln!(out, "    box(width: {}pt)[", format_f64(inner_width_pt),);
        let _ = writeln!(
            out,
            "      #scale(x: text_box_scale_{text_box_id}, y: text_box_scale_{text_box_id}, origin: top + left, reflow: true)["
        );
        let _ = writeln!(out, "        #text_box_raw_{text_box_id}");
        out.push_str("      ]\n");
        out.push_str("    ]\n");
        out.push_str("  }\n");
    } else {
        let _ = writeln!(
            out,
            "  #let text_box_content_{text_box_id} = block(width: {}pt)[",
            format_f64(inner_width_pt),
        );
        for (index, block) in text_box.content.iter().enumerate() {
            if index > 0 {
                out.push('\n');
            }
            out.push_str("  ");
            generate_fixed_text_box_block(out, block, ctx, Some(inner_width_pt), text_box.no_wrap)?;
        }
        out.push_str("  ]\n");
    }

    match text_box.vertical_align {
        TextBoxVerticalAlign::Top => {
            let _ = writeln!(out, "  #text_box_content_{text_box_id}");
        }
        TextBoxVerticalAlign::Center | TextBoxVerticalAlign::Bottom => {
            out.push_str("  #context {\n");
            let _ = writeln!(
                out,
                "    let text_box_slack_{text_box_id} = calc.max({}pt - measure(text_box_content_{text_box_id}).height, 0pt)",
                format_f64(inner_height_pt),
            );
            let spacer_expr = match text_box.vertical_align {
                TextBoxVerticalAlign::Center => format!("text_box_slack_{text_box_id} / 2"),
                TextBoxVerticalAlign::Bottom => format!("text_box_slack_{text_box_id}"),
                TextBoxVerticalAlign::Top => unreachable!(),
            };
            let _ = writeln!(out, "    let text_box_aligned_{text_box_id} = [");
            let _ = writeln!(out, "      #v({spacer_expr})");
            let _ = writeln!(out, "      #text_box_content_{text_box_id}");
            out.push_str("    ]\n");
            let _ = writeln!(out, "    text_box_aligned_{text_box_id}");
            out.push_str("  }\n");
        }
    }

    out.push_str("]\n");
    Ok(())
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
fn write_flow_page_setup(out: &mut String, page: &FlowPage, size: &PageSize, ctx: &mut GenCtx) {
    if page.header.is_none() && page.footer.is_none() {
        write_page_setup(out, size, &page.margins);
        return;
    }

    let _ = write!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt)",
        format_f64(size.width),
        format_f64(size.height),
        format_f64(page.margins.top),
        format_f64(page.margins.bottom),
        format_f64(page.margins.left),
        format_f64(page.margins.right),
    );

    if let Some(header) = &page.header
        && hf_has_flow_content(header)
    {
        if hf_needs_context(header) {
            out.push_str(", header: context [");
        } else {
            out.push_str(", header: [");
        }
        generate_flow_hf_content(out, header, ctx);
        out.push(']');
    }

    if let Some(footer) = &page.footer
        && hf_has_flow_content(footer)
    {
        let edge_offset = footer
            .distance_from_edge
            .map(|distance| (page.margins.bottom - distance).max(0.0))
            .unwrap_or(0.0);
        if hf_needs_stack_offset(footer) || edge_offset > 0.0 {
            out.push_str(", footer: context { let footer_content = block(width: 100%)[");
            generate_flow_hf_content(out, footer, ctx);
            out.push_str("]; move(dy: ");
            if hf_needs_stack_offset(footer) {
                out.push_str("-measure(footer_content).height / 2");
                if edge_offset > 0.0 {
                    let _ = write!(out, " - {}pt", format_f64(edge_offset));
                }
            } else {
                let _ = write!(out, "-{}pt", format_f64(edge_offset));
            }
            out.push_str(")[#footer_content] }");
        } else if hf_needs_context(footer) {
            out.push_str(", footer: context [");
            generate_flow_hf_content(out, footer, ctx);
            out.push(']');
        } else {
            out.push_str(", footer: [");
            generate_flow_hf_content(out, footer, ctx);
            out.push(']');
        }
    }

    if page
        .header
        .iter()
        .chain(page.footer.iter())
        .any(hf_has_page_anchored_frames)
    {
        out.push_str(", foreground: [");
        if let Some(header) = &page.header {
            generate_page_anchored_hf_frames(out, header, size.width, page.margins.right, ctx);
        }
        if let Some(footer) = &page.footer {
            generate_page_anchored_hf_frames(out, footer, size.width, page.margins.right, ctx);
        }
        out.push(']');
    }

    out.push_str(")\n");
}

fn is_page_anchored_frame(frame: &HeaderFooterFrame) -> bool {
    frame.horizontal_anchor == FrameAnchor::Page && frame.vertical_anchor == FrameAnchor::Page
}

fn hf_has_flow_content(header_footer: &HeaderFooter) -> bool {
    header_footer
        .paragraphs
        .iter()
        .any(hf_paragraph_has_flow_content)
}

fn hf_paragraph_has_content(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    !paragraph.elements.is_empty() || paragraph.border.is_some()
}

fn hf_paragraph_has_flow_content(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    hf_paragraph_has_content(paragraph)
        && paragraph
            .frame
            .as_ref()
            .is_none_or(|frame| !is_page_anchored_frame(frame))
}

fn hf_has_page_anchored_frames(header_footer: &HeaderFooter) -> bool {
    header_footer
        .paragraphs
        .iter()
        .any(|paragraph| paragraph.frame.as_ref().is_some_and(is_page_anchored_frame))
}

fn generate_flow_hf_content(out: &mut String, hf: &HeaderFooter, ctx: &mut GenCtx) {
    let mut is_first: bool = true;
    for paragraph in &hf.paragraphs {
        if paragraph.frame.as_ref().is_some_and(is_page_anchored_frame)
            || !hf_paragraph_has_content(paragraph)
        {
            continue;
        }
        if !is_first {
            out.push_str("\\\n");
        }
        generate_hf_styled_paragraph(out, paragraph, ctx);
        is_first = false;
    }
}

fn generate_page_anchored_hf_frames(
    out: &mut String,
    hf: &HeaderFooter,
    page_width: f64,
    right_margin: f64,
    ctx: &mut GenCtx,
) {
    let mut index: usize = 0;
    while index < hf.paragraphs.len() {
        let Some(frame) = hf.paragraphs[index].frame.as_ref() else {
            index += 1;
            continue;
        };
        if !is_page_anchored_frame(frame) {
            index += 1;
            continue;
        }
        let mut end: usize = index + 1;
        while end < hf.paragraphs.len() && hf.paragraphs[end].frame.as_ref() == Some(frame) {
            end += 1;
        }
        let x = frame.x.unwrap_or(0.0);
        let y = frame.y.unwrap_or(0.0);
        let _ = write!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt)[#block(",
            format_f64(x),
            format_f64(y)
        );
        if let Some(width) = frame.width {
            let _ = write!(out, "width: {}pt", format_f64(width));
        } else {
            let width = (page_width - x - right_margin).max(0.0);
            let _ = write!(out, "width: {}pt", format_f64(width));
        }
        out.push_str(")[#stack(dir: ttb, spacing: 4pt");
        for paragraph in &hf.paragraphs[index..end] {
            out.push_str(", [");
            if hf_paragraph_has_content(paragraph) {
                generate_hf_styled_paragraph(out, paragraph, ctx);
            } else {
                out.push_str("#box(height: 12pt)");
            }
            out.push(']');
        }
        out.push_str(")]]");
        index = end;
    }
}

/// Write the full page setup for a SheetPage, including optional header/footer.
fn write_table_page_setup(out: &mut String, page: &SheetPage, size: &PageSize, ctx: &mut GenCtx) {
    if page.header.is_none() && page.footer.is_none() {
        write_page_setup(out, size, &page.margins);
        return;
    }

    let _ = write!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt)",
        format_f64(size.width),
        format_f64(size.height),
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
        generate_hf_content(out, header, ctx);
        out.push(']');
    }

    if let Some(footer) = &page.footer {
        if hf_needs_stack_offset(footer) {
            out.push_str(", footer: context { let footer_content = block(width: 100%)[");
            generate_hf_content(out, footer, ctx);
            out.push_str("]; move(dy: -measure(footer_content).height / 2)[#footer_content] }");
        } else if hf_needs_context(footer) {
            out.push_str(", footer: context [");
            generate_hf_content(out, footer, ctx);
            out.push(']');
        } else {
            out.push_str(", footer: [");
            generate_hf_content(out, footer, ctx);
            out.push(']');
        }
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

fn hf_needs_stack_offset(hf: &HeaderFooter) -> bool {
    hf.paragraphs
        .iter()
        .filter(|paragraph| hf_paragraph_has_flow_content(paragraph))
        .count()
        > 1
        || hf
            .paragraphs
            .iter()
            .filter(|paragraph| hf_paragraph_has_flow_content(paragraph))
            .flat_map(|paragraph| &paragraph.elements)
            .any(|element| matches!(element, HFInline::Image(_)))
}

/// Generate inline content for a header or footer.
fn generate_hf_content(out: &mut String, hf: &HeaderFooter, ctx: &mut GenCtx) {
    for (i, para) in hf.paragraphs.iter().enumerate() {
        if i > 0 {
            out.push_str("\\\n");
        }
        generate_hf_styled_paragraph(out, para, ctx);
    }
}

fn generate_hf_styled_paragraph(
    out: &mut String,
    paragraph: &crate::ir::HeaderFooterParagraph,
    ctx: &mut GenCtx,
) {
    if let Some(align) = paragraph.style.alignment {
        let align_str = match align {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => "left",
        };
        let _ = write!(out, "#align({align_str})[");
    }
    if paragraph.style.direction == Some(TextDirection::Rtl) {
        out.push_str("#text(dir: rtl)[");
    }
    generate_hf_paragraph(out, paragraph, ctx);
    if paragraph.style.direction == Some(TextDirection::Rtl) {
        out.push(']');
    }
    if paragraph.style.alignment.is_some() {
        out.push(']');
    }
}

fn generate_hf_paragraph(
    out: &mut String,
    paragraph: &crate::ir::HeaderFooterParagraph,
    ctx: &mut GenCtx,
) {
    let right_tab = paragraph.elements.iter().position(|element| {
        matches!(
            element,
            HFInline::PositionedTab(tab)
                if tab.alignment == PositionedTabAlignment::Right
                    && tab.relative_to == PositionedTabRelativeTo::Margin
        )
    });
    let has_top_border = paragraph
        .border
        .as_ref()
        .and_then(|border| border.top.as_ref());

    if let Some(border) = has_top_border {
        out.push_str("#stack(dir: ttb, spacing: 0.5pt, ");
        write_hf_border_line(out, border, border.style == BorderLineStyle::Double);
        if border.style == BorderLineStyle::Double {
            out.push_str(", ");
            write_hf_border_line(out, border, false);
        }
        out.push_str(", [");
    }

    if let Some(index) = right_tab {
        out.push_str("#grid(columns: (1fr, auto), [");
        generate_hf_elements(out, &paragraph.elements[..index], ctx);
        out.push_str("], [");
        generate_hf_elements(out, &paragraph.elements[index + 1..], ctx);
        out.push_str("])");
    } else {
        generate_hf_elements(out, &paragraph.elements, ctx);
    }

    if has_top_border.is_some() {
        out.push_str("])");
    }
}

fn write_hf_border_line(out: &mut String, border: &BorderSide, is_primary_double: bool) {
    let width = if is_primary_double {
        border.width * 0.67
    } else if border.style == BorderLineStyle::Double {
        border.width * 0.17
    } else {
        border.width
    };
    let dash = border_line_style_to_typst(border.style);
    let _ = write!(
        out,
        "block(height: {}pt)[#line(length: 100%, stroke: (paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\"))]",
        format_f64(width),
        border.color.r,
        border.color.g,
        border.color.b,
        format_f64(width),
        if border.style == BorderLineStyle::Double {
            "solid"
        } else {
            dash
        }
    );
}

fn generate_hf_elements(out: &mut String, elements: &[HFInline], ctx: &mut GenCtx) {
    for element in elements {
        match element {
            HFInline::Run(run) => generate_run(out, run),
            HFInline::Image(image) => generate_image(out, image, ctx),
            HFInline::PageNumber => out.push_str("#counter(page).display()"),
            HFInline::TotalPages => out.push_str("#counter(page).final().first()"),
            HFInline::PositionedTab(_) => out.push_str("#h(1em)"),
        }
    }
}

/// Generate Typst markup for a sequence of blocks, separating each with a newline.
fn generate_blocks(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let mut index: usize = 0;
    while index < blocks.len() {
        if index > 0 {
            out.push('\n');
        }

        if is_zero_size_floating_anchor(&blocks[index]) {
            let consumed = generate_floating_anchor_group(out, &blocks[index..], ctx)?;
            index += consumed;
            continue;
        }

        generate_block(out, &blocks[index], ctx)?;
        index += 1;
    }

    Ok(())
}

fn is_zero_size_floating_anchor(block: &Block) -> bool {
    match block {
        Block::FloatingShape(shape) => matches!(
            shape.wrap_mode,
            WrapMode::Behind | WrapMode::InFront | WrapMode::None
        ),
        Block::FloatingTextBox(text_box) => matches!(
            text_box.wrap_mode,
            WrapMode::Behind | WrapMode::InFront | WrapMode::None
        ),
        _ => false,
    }
}

fn generate_floating_anchor_group(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<usize, ConvertError> {
    out.push_str("#box(width: 0pt, height: 0pt)[\n");
    let mut consumed: usize = 0;

    for block in blocks {
        if !is_zero_size_floating_anchor(block) {
            break;
        }

        match block {
            Block::FloatingShape(shape) => generate_floating_shape_overlay(out, shape),
            Block::FloatingTextBox(text_box) => {
                generate_floating_text_box_overlay(out, text_box, ctx)?;
            }
            _ => unreachable!("checked by is_zero_size_floating_anchor"),
        }
        consumed += 1;
    }

    out.push_str("]\n");
    Ok(consumed)
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
            let align_str: Option<&str> = match img.alignment {
                Some(Alignment::Center) => Some("center"),
                Some(Alignment::Right) => Some("right"),
                _ => None,
            };
            if let Some(align_str) = align_str {
                let _ = write!(out, "#align({align_str})[");
            }
            if let Some(ref stroke) = img.stroke {
                out.push_str("#box(stroke: ");
                shapes::write_image_border_stroke(out, stroke);
                out.push_str(")[");
                generate_image(out, img, ctx);
                out.push(']');
            } else {
                generate_image(out, img, ctx);
            }
            if align_str.is_some() {
                out.push(']');
            }
            out.push('\n');
            Ok(())
        }
        Block::InlineImages(images) => {
            out.push_str("#block(width: 100%)[\n");
            for (index, image) in images.iter().enumerate() {
                if index > 0 {
                    out.push(' ');
                }
                out.push_str("#box[");
                generate_image(out, image, ctx);
                out.pop();
                out.push(']');
            }
            out.push_str("\n]\n");
            Ok(())
        }
        Block::FloatingImage(fi) => {
            generate_floating_image(out, fi, ctx);
            Ok(())
        }
        Block::FloatingTextBox(ftb) => generate_floating_text_box(out, ftb, ctx),
        Block::FloatingShape(fs) => {
            generate_floating_shape(out, fs);
            Ok(())
        }
        Block::List(list) => generate_list(out, list),
        Block::MathEquation(math) => {
            generate_math_equation(out, math);
            Ok(())
        }
        Block::Chart(chart) => {
            generate_chart(out, chart);
            Ok(())
        }
        Block::ColumnBreak => {
            out.push_str("#colbreak()\n");
            Ok(())
        }
    }
}

/// Generate Typst markup for a math equation.
///
/// Display math is rendered as `$ content $` (on its own line, centered).
/// Inline math is rendered as `$content$`.
fn generate_math_equation(out: &mut String, math: &MathEquation) {
    if math.display {
        let _ = writeln!(out, "$ {} $", math.content);
    } else {
        let _ = write!(out, "${}$", math.content);
    }
}

fn format_insets(insets: &Insets) -> String {
    format!(
        "(top: {}pt, right: {}pt, bottom: {}pt, left: {}pt)",
        format_f64(insets.top),
        format_f64(insets.right),
        format_f64(insets.bottom),
        format_f64(insets.left),
    )
}

fn border_line_style_to_typst(style: BorderLineStyle) -> &'static str {
    match style {
        BorderLineStyle::Solid => "solid",
        BorderLineStyle::Dashed => "dashed",
        BorderLineStyle::Dotted => "dotted",
        BorderLineStyle::DashDot => "dash-dotted",
        BorderLineStyle::DashDotDot => "dash-dotted",
        BorderLineStyle::Double => "dashed",
        BorderLineStyle::None => "solid",
    }
}

fn generate_image(out: &mut String, img: &ImageData, ctx: &mut GenCtx) {
    let path = ctx.add_image(img);

    out.push_str("#image(\"");
    out.push_str(&path);
    out.push('"');

    if let Some(w) = img.width {
        let _ = write!(out, ", width: {}pt", format_f64(w));
    }
    if let Some(h) = img.height {
        let _ = write!(out, ", height: {}pt", format_f64(h));
    }

    // Typst defaults to fit: "cover" which preserves the image's native
    // aspect ratio.  When both width and height are specified (common for
    // PPTX slides), the image must fill its bounding box exactly — e.g.
    // after a non-uniform group transform the AR may differ from the
    // pixel data.  "stretch" ensures the rendered size matches.
    if img.width.is_some() && img.height.is_some() {
        out.push_str(", fit: \"stretch\"");
    }

    out.push_str(")\n");
}

/// Generate Typst markup for a floating image.
///
/// Uses `#place()` for absolute positioning. The wrap mode determines how text
/// interacts with the image:
/// - Behind/InFront/None: `#place()` with no text wrapping
/// - Square/Tight/TopAndBottom: `#place()` with `float: true` for best-effort text flow
fn generate_floating_image(out: &mut String, fi: &FloatingImage, ctx: &mut GenCtx) {
    let path = ctx.add_image(&fi.image);

    match fi.wrap_mode {
        WrapMode::TopAndBottom => {
            // Emit a block-level image — text above and below only
            out.push_str("#block(width: 100%)[\n");
            let _ = write!(
                out,
                "  #place(top + left, dx: {}pt, dy: 0pt)[",
                format_f64(fi.offset_x)
            );
            out.push_str("#image(\"");
            out.push_str(&path);
            out.push('"');
            if let Some(w) = fi.image.width {
                let _ = write!(out, ", width: {}pt", format_f64(w));
            }
            if let Some(h) = fi.image.height {
                let _ = write!(out, ", height: {}pt", format_f64(h));
            }
            out.push_str(")]\n");
            // Reserve vertical space equal to image height
            if let Some(h) = fi.image.height {
                let _ = writeln!(out, "  #v({}pt)", format_f64(h));
            }
            out.push_str("]\n");
        }
        WrapMode::Behind | WrapMode::InFront | WrapMode::None => {
            // Place the image at absolute position, no text wrapping
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt)[",
                format_f64(fi.offset_x),
                format_f64(fi.offset_y)
            );
            out.push_str("#image(\"");
            out.push_str(&path);
            out.push('"');
            if let Some(w) = fi.image.width {
                let _ = write!(out, ", width: {}pt", format_f64(w));
            }
            if let Some(h) = fi.image.height {
                let _ = write!(out, ", height: {}pt", format_f64(h));
            }
            out.push_str(")]\n");
        }
        WrapMode::Square | WrapMode::Tight => {
            // Best-effort text wrapping: use #place with float: true
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, float: true)[",
                format_f64(fi.offset_x),
                format_f64(fi.offset_y)
            );
            out.push_str("#image(\"");
            out.push_str(&path);
            out.push('"');
            if let Some(w) = fi.image.width {
                let _ = write!(out, ", width: {}pt", format_f64(w));
            }
            if let Some(h) = fi.image.height {
                let _ = write!(out, ", height: {}pt", format_f64(h));
            }
            out.push_str(")]\n");
        }
    }
}

fn generate_floating_text_box(
    out: &mut String,
    ftb: &FloatingTextBox,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    match ftb.wrap_mode {
        WrapMode::TopAndBottom => {
            out.push_str("#block(width: 100%)[\n");
            let _ = writeln!(
                out,
                "  #place(top + left, dx: {}pt, dy: 0pt)[",
                format_f64(ftb.offset_x)
            );
            generate_floating_text_box_content(out, ftb, ctx)?;
            out.push_str("  ]\n");
            if ftb.height > 0.0 {
                let _ = writeln!(out, "  #v({}pt)", format_f64(ftb.height));
            }
            out.push_str("]\n");
        }
        WrapMode::Behind | WrapMode::InFront | WrapMode::None => {
            // Anchor to the current flow position (the box's paragraph), not the
            // page, by wrapping `#place` in a zero-size box. Without this the
            // box piles at the page top, away from the shapes it belongs with
            // (issue #176).
            out.push_str("#box(width: 0pt, height: 0pt)[\n");
            generate_floating_text_box_overlay(out, ftb, ctx)?;
            out.push_str("]\n");
        }
        WrapMode::Square | WrapMode::Tight => {
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, float: true)[",
                format_f64(ftb.offset_x),
                format_f64(ftb.offset_y)
            );
            generate_floating_text_box_content(out, ftb, ctx)?;
            out.push_str("]\n");
        }
    }

    Ok(())
}

/// Generate Typst markup for a floating geometric shape (issue #176).
///
/// The DOCX anchor positions the shape relative to its paragraph (`positionV
/// relativeFrom="paragraph"`) and the text column (`positionH
/// relativeFrom="column"`), not the page. A bare `#place(top + left, …)` at the
/// document top level anchors to the page, piling every shape at the top. To
/// anchor to the current flow position instead, the `#place` is wrapped in a
/// zero-size `#box`, whose top-left sits exactly where the anchoring paragraph
/// is laid out. Word-processing shapes use `wrapNone`, so no float is needed.
fn generate_floating_shape(out: &mut String, fs: &FloatingShape) {
    out.push_str("#box(width: 0pt, height: 0pt)[\n");
    generate_floating_shape_overlay(out, fs);
    out.push_str("]\n");
}

fn generate_floating_shape_overlay(out: &mut String, fs: &FloatingShape) {
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(fs.offset_x),
        format_f64(fs.offset_y)
    );
    shapes::generate_shape(out, &fs.shape, fs.width, fs.height);
    out.push_str("]\n");
}

fn generate_floating_text_box_overlay(
    out: &mut String,
    ftb: &FloatingTextBox,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let _ = writeln!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(ftb.offset_x),
        format_f64(ftb.offset_y)
    );
    generate_floating_text_box_content(out, ftb, ctx)?;
    out.push_str("]\n");
    Ok(())
}

fn generate_floating_text_box_content(
    out: &mut String,
    ftb: &FloatingTextBox,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let inner_width: f64 = (ftb.width - ftb.padding.left - ftb.padding.right).max(0.0);
    let inner_height: f64 = (ftb.height - ftb.padding.top - ftb.padding.bottom).max(0.0);
    let inset: String = if ftb.padding == Insets::default() {
        "0pt".to_string()
    } else {
        format_insets(&ftb.padding)
    };
    let _ = writeln!(
        out,
        "#box(width: {}pt, height: {}pt, inset: {})[",
        format_f64(ftb.width),
        format_f64(ftb.height),
        inset,
    );

    if matches!(ftb.vertical_align, TextBoxVerticalAlign::Top) {
        let _ = writeln!(
            out,
            "#place(top + left, dy: -{}pt)[\n#block(width: {}pt)[",
            format_f64(FLOATING_TEXT_BOX_TOP_LEADING_COMPENSATION_PT),
            format_f64(inner_width)
        );
        for (index, block) in ftb.content.iter().enumerate() {
            if index > 0 {
                out.push('\n');
            }
            generate_fixed_text_box_block(out, block, ctx, Some(inner_width), false)?;
        }
        out.push_str("]\n]\n]\n");
        return Ok(());
    }

    let text_box_id: usize = ctx.next_text_box_id();
    let _ = writeln!(
        out,
        "#let floating_text_box_content_{text_box_id} = block(width: {}pt)[",
        format_f64(inner_width)
    );
    for (index, block) in ftb.content.iter().enumerate() {
        if index > 0 {
            out.push('\n');
        }
        generate_fixed_text_box_block(out, block, ctx, Some(inner_width), false)?;
    }
    out.push_str("]\n#context {\n");
    let _ = writeln!(
        out,
        "  let floating_text_box_slack_{text_box_id} = calc.max({}pt - measure(floating_text_box_content_{text_box_id}).height, 0pt)",
        format_f64(inner_height)
    );
    let spacer: String = match ftb.vertical_align {
        TextBoxVerticalAlign::Center => format!("floating_text_box_slack_{text_box_id} / 2"),
        TextBoxVerticalAlign::Bottom => format!("floating_text_box_slack_{text_box_id}"),
        TextBoxVerticalAlign::Top => unreachable!(),
    };
    out.push_str("  [\n");
    let _ = writeln!(out, "    #v({spacer})");
    let _ = writeln!(out, "    #floating_text_box_content_{text_box_id}");
    out.push_str("  ]\n");
    out.push_str("}\n]\n");
    Ok(())
}

fn single_line_fit_paragraph(text_box: &TextBoxData, inner_height_pt: f64) -> Option<&Paragraph> {
    if text_box.no_wrap && !text_box.auto_fit {
        return None;
    }
    let [Block::Paragraph(paragraph)] = text_box.content.as_slice() else {
        return None;
    };
    if paragraph.runs.is_empty() || paragraph_has_forced_breaks(paragraph) {
        return None;
    }

    let max_font_size_pt: f64 = paragraph_max_font_size_pt(paragraph);
    if max_font_size_pt <= 0.0 || inner_height_pt <= 0.0 {
        return None;
    }

    let has_mixed_font_sizes: bool = paragraph_has_mixed_font_sizes(paragraph);
    if has_mixed_font_sizes && inner_height_pt <= max_font_size_pt * 2.5 {
        return Some(paragraph);
    }

    let estimated_line_height_pt: f64 = estimate_single_line_height_pt(paragraph);
    if estimated_line_height_pt <= 0.0 {
        return None;
    }

    let is_short_box: bool = inner_height_pt <= estimated_line_height_pt * 2.0;
    if !is_short_box {
        return None;
    }

    let needs_single_line_fit: bool =
        text_box.auto_fit || inner_height_pt <= estimated_line_height_pt * 1.2;

    needs_single_line_fit.then_some(paragraph)
}

fn wrapped_fit_paragraph(text_box: &TextBoxData) -> Option<&Paragraph> {
    if text_box.no_wrap || matches!(text_box.vertical_align, TextBoxVerticalAlign::Top) {
        return None;
    }

    let [Block::Paragraph(paragraph)] = text_box.content.as_slice() else {
        return None;
    };

    (!paragraph.runs.is_empty() && !paragraph_has_forced_breaks(paragraph)).then_some(paragraph)
}

fn paragraph_has_forced_breaks(paragraph: &Paragraph) -> bool {
    paragraph.runs.iter().any(|run| {
        run.text
            .chars()
            .any(|ch| matches!(ch, '\n' | '\r' | '\u{000B}'))
    })
}

fn paragraph_has_mixed_font_sizes(paragraph: &Paragraph) -> bool {
    let mut first_size: Option<i64> = None;
    for run in &paragraph.runs {
        let size_pt: f64 = run.style.font_size.unwrap_or(12.0);
        let size_key: i64 = (size_pt * 100.0).round() as i64;
        match first_size {
            Some(first) if first != size_key => return true,
            None => first_size = Some(size_key),
            _ => {}
        }
    }
    false
}

fn estimate_single_line_height_pt(paragraph: &Paragraph) -> f64 {
    let max_font_size_pt: f64 = paragraph_max_font_size_pt(paragraph);
    let default_line_height_pt: f64 = max_font_size_pt * 1.2;

    match paragraph.style.line_spacing {
        Some(LineSpacing::Exact(points)) => default_line_height_pt.max(points),
        Some(LineSpacing::Proportional(factor)) => {
            default_line_height_pt.max(max_font_size_pt * factor)
        }
        None => default_line_height_pt,
    }
}

fn paragraph_max_font_size_pt(paragraph: &Paragraph) -> f64 {
    paragraph
        .runs
        .iter()
        .filter_map(|run| run.style.font_size)
        .fold(12.0, f64::max)
}

fn fixed_text_box_alignment_name(alignment: Option<Alignment>) -> Option<&'static str> {
    match alignment {
        Some(Alignment::Center) => Some("center"),
        Some(Alignment::Right) => Some("right"),
        Some(Alignment::Left) => Some("left"),
        _ => None,
    }
}

fn generate_fixed_text_box_block(
    out: &mut String,
    block: &Block,
    ctx: &mut GenCtx,
    available_width_pt: Option<f64>,
    no_wrap: bool,
) -> Result<(), ConvertError> {
    match block {
        Block::List(list) if can_render_fixed_text_list_inline(list) => {
            generate_fixed_text_list(out, list, true, available_width_pt)
        }
        Block::Paragraph(para) => generate_fixed_text_paragraph(out, para, no_wrap),
        _ => generate_block(out, block, ctx),
    }
}

fn generate_fixed_text_paragraph(
    out: &mut String,
    para: &Paragraph,
    no_wrap: bool,
) -> Result<(), ConvertError> {
    let style: &ParagraphStyle = &para.style;
    let needs_text_scope: bool = common_text_style(&para.runs).is_some();
    let has_para_style: bool = needs_block_wrapper(style) || needs_text_scope;

    if has_para_style {
        out.push_str("#block(");
        write_block_params(out, style);
        out.push_str(")[\n");
        write_par_settings(out, style);
        write_common_text_settings(out, &para.runs, "  ");
        write_fixed_text_default_par_settings(out, style, &para.runs, "  ");
    }

    let alignment = style.alignment;
    let use_align = matches!(
        alignment,
        Some(Alignment::Center) | Some(Alignment::Right) | Some(Alignment::Left)
    );

    // Use #block(width: 100%)[#set align(...); content] to ensure alignment
    // works reliably inside #context + measure() vertical centering.
    if use_align {
        let align_str = match alignment {
            Some(Alignment::Left) => "left",
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            _ => "left",
        };
        let _ = writeln!(out, "#block(width: 100%)[#set align({align_str})");
    }

    if no_wrap {
        out.push_str("#box[");
        generate_runs_with_tabs_no_wrap(out, &para.runs, style.tab_stops.as_deref());
    } else {
        generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());
    }
    if no_wrap {
        out.push(']');
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

#[cfg(test)]
#[path = "typst_gen_tests.rs"]
mod tests;
