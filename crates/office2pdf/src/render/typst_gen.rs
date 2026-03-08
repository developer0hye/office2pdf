use std::fmt::Write;
use std::io::Cursor;

use image::{GenericImageView, ImageFormat as RasterImageFormat};
use unicode_normalization::UnicodeNormalization;

use crate::config::ConvertOptions;
use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Chart, ChartType,
    Color, ColumnLayout, Document, FixedElement, FixedElementKind, FixedPage, FloatingImage,
    FloatingTextBox, FlowPage, GradientFill, HFInline, HeaderFooter, ImageCrop, ImageData,
    ImageFormat, Insets, LineSpacing, List, ListKind, Margins, MathEquation, Metadata, Page,
    PageSize, Paragraph, ParagraphStyle, Run, Shadow, Shape, ShapeKind, SmartArt, TabAlignment,
    TabLeader, TabStop, Table, TableCell, TablePage, TableRow, TextBoxData, TextBoxVerticalAlign,
    TextDirection, TextStyle, VerticalTextAlign, WrapMode,
};

use self::diagrams::{generate_chart, generate_smartart};
use self::lists::{
    can_render_fixed_text_list_inline, common_text_style, generate_fixed_text_list, generate_list,
    write_common_text_settings, write_fixed_text_default_par_settings,
};
use self::tables::generate_table;
use super::font_context::FontSearchContext;

#[path = "typst_gen_diagrams.rs"]
mod diagrams;
#[path = "typst_gen_lists.rs"]
mod lists;
#[path = "typst_gen_tables.rs"]
mod tables;

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
                Page::Table(table_page) => {
                    generate_table_page(&mut out, table_page, &mut ctx, options)?;
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
    write_flow_page_setup(out, page, &size);
    out.push('\n');

    if let Some(ref cols) = page.columns {
        generate_flow_page_columns(out, &page.content, cols, ctx)?;
    } else {
        for (i, block) in page.content.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            generate_block(out, block, ctx)?;
        }
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
        for (i, block) in content.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            generate_block(out, block, ctx)?;
        }
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
    page: &TablePage,
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    write_table_page_setup(out, page, &size);
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

    let _ = writeln!(
        out,
        "#block(width: {}pt, height: {}pt, inset: {})[",
        format_f64(outer_width_pt),
        format_f64(outer_height_pt),
        format_insets(&text_box.padding),
    );
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
        generate_fixed_text_box_block(out, block, ctx, Some(inner_width_pt))?;
    }
    out.push_str("  ]\n");

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

fn generate_shape(out: &mut String, shape: &Shape, width: f64, height: f64) {
    // Render shadow as offset duplicate before main shape
    if let Some(shadow) = &shape.shadow {
        write_shadow_shape(out, shape, width, height, shadow);
    }

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
            write_shape_stroke(out, &shape.stroke);
            out.push_str(")\n");
        }
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let radius = radius_fraction * width.min(height);
            out.push_str("#rect(");
            write_shape_params(out, shape, width, height);
            let _ = write!(out, ", radius: {}pt", format_f64(radius));
            out.push_str(")\n");
        }
        ShapeKind::Polygon { vertices } => {
            write_polygon(out, shape, width, height, vertices);
        }
    }

    if has_rotation {
        out.push_str("]\n");
    }
}

/// Render a shadow approximation as an offset duplicate shape with reduced opacity.
fn write_shadow_shape(out: &mut String, shape: &Shape, width: f64, height: f64, shadow: &Shadow) {
    let dir_rad = shadow.direction.to_radians();
    let dx = shadow.distance * dir_rad.cos();
    let dy = shadow.distance * dir_rad.sin();
    let alpha = (shadow.opacity * 255.0).round() as u8;

    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(dx),
        format_f64(dy),
    );

    match &shape.kind {
        ShapeKind::Line { .. } => {
            // Lines don't have meaningful shadows; skip
            out.push_str("]\n");
            return;
        }
        ShapeKind::Polygon { vertices } => {
            // Shadow for polygon: duplicate polygon with shadow color
            out.push_str("#polygon(");
            write_polygon_vertices(out, width, height, vertices);
            let _ = write!(
                out,
                ", fill: rgb({}, {}, {}, {})",
                shadow.color.r, shadow.color.g, shadow.color.b, alpha,
            );
            out.push_str(")]\n");
            return;
        }
        _ => {}
    }
    let shape_cmd = match &shape.kind {
        ShapeKind::Rectangle => "#rect(",
        ShapeKind::Ellipse => "#ellipse(",
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let _ = writeln!(
                out,
                "#rect(width: {}pt, height: {}pt, radius: {}pt, fill: rgb({}, {}, {}, {}))]",
                format_f64(width),
                format_f64(height),
                format_f64(radius_fraction * width.min(height)),
                shadow.color.r,
                shadow.color.g,
                shadow.color.b,
                alpha,
            );
            return;
        }
        // Line and Polygon are handled by early returns above; any future
        // variants gracefully skip the shadow rather than panicking.
        _ => {
            out.push_str("]\n");
            return;
        }
    };
    out.push_str(shape_cmd);
    let _ = write!(
        out,
        "width: {}pt, height: {}pt, fill: rgb({}, {}, {}, {})",
        format_f64(width),
        format_f64(height),
        shadow.color.r,
        shadow.color.g,
        shadow.color.b,
        alpha,
    );
    out.push_str(")]\n");
}

/// Write fill color, using rgb with 4 args when opacity is set, rgb with 3 args otherwise.
fn write_fill_color(out: &mut String, fill: &Color, opacity: Option<f64>) {
    if let Some(op) = opacity {
        let alpha = (op * 255.0).round() as u8;
        let _ = write!(
            out,
            ", fill: rgb({}, {}, {}, {})",
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
    if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
}

/// Write stroke parameter for shapes, handling dash patterns.
fn write_shape_stroke(out: &mut String, stroke: &Option<BorderSide>) {
    if let Some(stroke) = stroke {
        match stroke.style {
            BorderLineStyle::Solid | BorderLineStyle::None => {
                let _ = write!(
                    out,
                    ", stroke: {}pt + rgb({}, {}, {})",
                    format_f64(stroke.width),
                    stroke.color.r,
                    stroke.color.g,
                    stroke.color.b,
                );
            }
            _ => {
                let _ = write!(
                    out,
                    ", stroke: (paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\")",
                    stroke.color.r,
                    stroke.color.g,
                    stroke.color.b,
                    format_f64(stroke.width),
                    border_line_style_to_typst(stroke.style),
                );
            }
        }
    }
}

/// Write polygon vertex coordinates scaled to actual dimensions.
fn write_polygon_vertices(out: &mut String, width: f64, height: f64, vertices: &[(f64, f64)]) {
    for (i, (vx, vy)) in vertices.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        let _ = write!(
            out,
            "({}pt, {}pt)",
            format_f64(vx * width),
            format_f64(vy * height),
        );
    }
}

/// Generate a Typst `#polygon(...)` for an arbitrary polygon shape.
fn write_polygon(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    vertices: &[(f64, f64)],
) {
    out.push_str("#polygon(");
    write_polygon_vertices(out, width, height, vertices);
    if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
    out.push_str(")\n");
}

/// Write a Typst `gradient.linear(...)` expression.
///
/// Stops are sorted by position before rendering because Typst requires
/// gradient stop offsets to be in monotonic (non-decreasing) order.
/// The first stop is clamped to 0% and the last to 100% as Typst requires.
fn write_gradient_fill(out: &mut String, gradient: &GradientFill) {
    // Typst requires at least 2 stops for gradient.linear().
    // Fall back to solid fill if fewer than 2 stops.
    if gradient.stops.len() < 2 {
        if let Some(stop) = gradient.stops.first() {
            let _ = write!(
                out,
                "rgb({}, {}, {})",
                stop.color.r, stop.color.g, stop.color.b,
            );
        }
        return;
    }
    let mut sorted_stops = gradient.stops.clone();
    sorted_stops.sort_by(|a, b| {
        a.position
            .partial_cmp(&b.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    // Typst requires first stop at 0% and last stop at 100%.
    if let Some(first) = sorted_stops.first_mut() {
        first.position = 0.0;
    }
    if let Some(last) = sorted_stops.last_mut() {
        last.position = 1.0;
    }
    out.push_str("gradient.linear(");
    for (i, stop) in sorted_stops.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        let pos_pct = (stop.position * 100.0).round() as i64;
        let _ = write!(
            out,
            "(rgb({}, {}, {}), {}%)",
            stop.color.r, stop.color.g, stop.color.b, pos_pct,
        );
    }
    if gradient.angle.abs() > 0.001 {
        let _ = write!(out, ", angle: {}deg", format_f64(gradient.angle));
    }
    out.push(')');
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
fn write_flow_page_setup(out: &mut String, page: &FlowPage, size: &PageSize) {
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
fn write_table_page_setup(out: &mut String, page: &TablePage, size: &PageSize) {
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
        Block::FloatingImage(fi) => {
            generate_floating_image(out, fi, ctx);
            Ok(())
        }
        Block::FloatingTextBox(ftb) => generate_floating_text_box(out, ftb, ctx),
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
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt)[",
                format_f64(ftb.offset_x),
                format_f64(ftb.offset_y)
            );
            generate_floating_text_box_content(out, ftb, ctx)?;
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

fn generate_floating_text_box_content(
    out: &mut String,
    ftb: &FloatingTextBox,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let _ = writeln!(
        out,
        "#block(width: {}pt, height: {}pt)[",
        format_f64(ftb.width),
        format_f64(ftb.height)
    );
    for (index, block) in ftb.content.iter().enumerate() {
        if index > 0 {
            out.push('\n');
        }
        generate_fixed_text_box_block(out, block, ctx, Some(ftb.width))?;
    }
    out.push_str("]\n");
    Ok(())
}

fn generate_fixed_text_box_block(
    out: &mut String,
    block: &Block,
    ctx: &mut GenCtx,
    available_width_pt: Option<f64>,
) -> Result<(), ConvertError> {
    match block {
        Block::List(list) if can_render_fixed_text_list_inline(list) => {
            generate_fixed_text_list(out, list, true, available_width_pt)
        }
        Block::Paragraph(para) => generate_fixed_text_paragraph(out, para),
        _ => generate_block(out, block, ctx),
    }
}

fn generate_fixed_text_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
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

    if use_align {
        let align_str = match alignment {
            Some(Alignment::Left) => "left",
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            _ => "left",
        };
        let _ = write!(out, "#align({align_str})[");
    }

    generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

fn generate_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
    let style = &para.style;

    // Heading paragraphs: emit #heading(level: N)[content] for proper PDF structure tagging
    if let Some(level) = style.heading_level {
        let _ = write!(out, "#heading(level: {level})[");
        generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());
        out.push_str("]\n");
        return Ok(());
    }

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
        let align_str = match alignment {
            Some(Alignment::Left) => "left",
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            // Justify and None are excluded by use_align, but handle gracefully
            _ => "left",
        };
        let _ = write!(out, "#align({align_str})[");
    }

    generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

/// Check if paragraph style needs a block wrapper (for spacing/leading/justify/direction).
fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.line_spacing.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
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
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

/// Word's default tab stop interval (0.5 inch = 36pt).
const DEFAULT_TAB_WIDTH_PT: f64 = 36.0;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';

fn generate_runs_with_tabs(out: &mut String, runs: &[Run], tab_stops: Option<&[TabStop]>) {
    if !paragraph_contains_tabs(runs) {
        generate_runs(out, runs);
        return;
    }

    let segments: Vec<Vec<Run>> = split_runs_on_tabs(runs);
    out.push_str("#context {\n");

    for (index, segment) in segments.iter().enumerate() {
        let _ = write!(out, "  let tab_segment_{index} = [");
        generate_runs(out, segment);
        out.push_str("]\n");

        if index == 0 {
            out.push_str("  let tab_prefix_0 = tab_segment_0\n");
            continue;
        }

        let _ = writeln!(
            out,
            "  let tab_prefix_width_{index} = measure(tab_prefix_{}).width",
            index - 1
        );
        let _ = writeln!(
            out,
            "  let tab_segment_width_{index} = measure(tab_segment_{index}).width"
        );

        if let Some(anchor_runs) = extract_decimal_anchor_runs(segment) {
            let _ = write!(out, "  let tab_decimal_anchor_{index} = [");
            generate_runs(out, &anchor_runs);
            out.push_str("]\n");
            let _ = writeln!(
                out,
                "  let tab_decimal_width_{index} = measure(tab_decimal_anchor_{index}).width"
            );
        }

        let _ = writeln!(
            out,
            "  let tab_default_remainder_{index} = calc.rem-euclid(tab_prefix_width_{index}.abs.pt(), {})",
            format_f64(DEFAULT_TAB_WIDTH_PT)
        );
        let _ = writeln!(
            out,
            "  let tab_advance_{index} = {}",
            build_tab_advance_expr(index, segment, tab_stops)
        );
        let _ = writeln!(
            out,
            "  let tab_fill_{index} = {}",
            build_tab_fill_expr(index, tab_stops)
        );
        let _ = writeln!(
            out,
            "  let tab_prefix_{index} = [#tab_prefix_{}#tab_fill_{index}#tab_segment_{index}]",
            index - 1
        );
    }

    let _ = writeln!(out, "  tab_prefix_{}", segments.len() - 1);
    out.push('}');
}

fn paragraph_contains_tabs(runs: &[Run]) -> bool {
    runs.iter().any(|run| run.text.contains('\t'))
}

fn generate_runs(out: &mut String, runs: &[Run]) {
    for run in runs {
        generate_run(out, run);
    }
}

fn split_runs_on_tabs(runs: &[Run]) -> Vec<Vec<Run>> {
    let mut segments: Vec<Vec<Run>> = vec![Vec::new()];

    for run in runs {
        if run.footnote.is_some() || !run.text.contains('\t') {
            if run.footnote.is_some() || !run.text.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(run.clone());
            }
            continue;
        }

        for (index, part) in run.text.split('\t').enumerate() {
            if index > 0 {
                segments.push(Vec::new());
            }

            if !part.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(Run {
                        text: part.to_string(),
                        style: run.style.clone(),
                        href: run.href.clone(),
                        footnote: None,
                    });
            }
        }
    }

    segments
}

fn extract_decimal_anchor_runs(runs: &[Run]) -> Option<Vec<Run>> {
    let visible_text: String = runs
        .iter()
        .filter(|run| run.footnote.is_none())
        .map(|run| run.text.as_str())
        .collect();
    let separator_offset = find_decimal_separator_offset(&visible_text)?;

    let mut anchor_runs: Vec<Run> = Vec::new();
    let mut visible_offset: usize = 0;

    for run in runs {
        if let Some(content) = &run.footnote {
            anchor_runs.push(Run {
                text: String::new(),
                style: run.style.clone(),
                href: run.href.clone(),
                footnote: Some(content.clone()),
            });
            continue;
        }

        let run_end = visible_offset + run.text.len();
        if run_end <= separator_offset {
            if !run.text.is_empty() {
                anchor_runs.push(run.clone());
            }
            visible_offset = run_end;
            continue;
        }

        let offset = separator_offset.saturating_sub(visible_offset);
        if offset > 0 {
            anchor_runs.push(Run {
                text: run.text[..offset].to_string(),
                style: run.style.clone(),
                href: run.href.clone(),
                footnote: None,
            });
        }

        return Some(anchor_runs);
    }

    None
}

fn find_decimal_separator_offset(text: &str) -> Option<usize> {
    let separator = text.char_indices().rev().find(|(offset, ch)| {
        matches!(ch, '.' | ',')
            && has_ascii_digit_before(text, *offset)
            && has_ascii_digit_after(text, *offset + ch.len_utf8())
    })?;

    if is_grouped_integer(
        &text
            .chars()
            .filter(|ch| ch.is_ascii_digit() || matches!(ch, '.' | ','))
            .collect::<String>(),
        separator.1,
    ) {
        return None;
    }

    Some(separator.0)
}

fn has_ascii_digit_before(text: &str, offset: usize) -> bool {
    text[..offset].chars().rev().any(|ch| ch.is_ascii_digit())
}

fn has_ascii_digit_after(text: &str, offset: usize) -> bool {
    text[offset..].chars().any(|ch| ch.is_ascii_digit())
}

fn is_grouped_integer(text: &str, separator: char) -> bool {
    if text
        .chars()
        .any(|ch| matches!(ch, '.' | ',') && ch != separator)
    {
        return false;
    }

    let parts: Vec<&str> = text.split(separator).collect();
    parts.len() > 1
        && parts
            .iter()
            .all(|part| !part.is_empty() && part.chars().all(|ch| ch.is_ascii_digit()))
        && parts[1..].iter().all(|part| part.len() == 3)
}

fn build_tab_advance_expr(index: usize, segment: &[Run], tab_stops: Option<&[TabStop]>) -> String {
    let prefix_width_var = format!("tab_prefix_width_{index}");
    let segment_width_var = format!("tab_segment_width_{index}");
    let decimal_width_var =
        extract_decimal_anchor_runs(segment).map(|_| format!("tab_decimal_width_{index}"));
    let default_expr = build_default_tab_advance_expr(index);

    let Some(tab_stops) = tab_stops else {
        return default_expr;
    };

    if tab_stops.is_empty() {
        return default_expr;
    }

    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = format!(
            "calc.max(0pt, {}pt - {prefix_width_var} - {})",
            format_f64(stop.position),
            tab_alignment_offset_expr(stop, &segment_width_var, decimal_width_var.as_deref())
        );

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ {default_expr} }}");
    expr
}

fn build_tab_fill_expr(index: usize, tab_stops: Option<&[TabStop]>) -> String {
    let Some(tab_stops) = tab_stops else {
        return format!("h(tab_advance_{index})");
    };

    if tab_stops.is_empty() {
        return format!("h(tab_advance_{index})");
    }

    let prefix_width_var = format!("tab_prefix_width_{index}");
    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = tab_fill_content_expr(index, stop.leader);

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ h(tab_advance_{index}) }}");
    expr
}

fn tab_fill_content_expr(index: usize, leader: TabLeader) -> String {
    let leader_markup = match leader {
        TabLeader::None => return format!("h(tab_advance_{index})"),
        TabLeader::Dot => ".",
        TabLeader::Hyphen => "-",
        TabLeader::Underscore => "\\_",
    };

    format!("box(width: tab_advance_{index}, repeat[{leader_markup}])")
}

fn build_default_tab_advance_expr(index: usize) -> String {
    format!(
        "if tab_default_remainder_{index} == 0 {{ {}pt }} else {{ ({} - tab_default_remainder_{index}) * 1pt }}",
        format_f64(DEFAULT_TAB_WIDTH_PT),
        format_f64(DEFAULT_TAB_WIDTH_PT)
    )
}

fn tab_alignment_offset_expr(
    stop: &TabStop,
    segment_width_var: &str,
    decimal_width_var: Option<&str>,
) -> String {
    match stop.alignment {
        TabAlignment::Left => "0pt".to_string(),
        TabAlignment::Center => format!("{segment_width_var} / 2"),
        TabAlignment::Right => segment_width_var.to_string(),
        TabAlignment::Decimal => decimal_width_var.unwrap_or(segment_width_var).to_string(),
    }
}

fn generate_run(out: &mut String, run: &Run) {
    // Emit footnote if present (footnote runs have empty text)
    if let Some(ref content) = run.footnote {
        let escaped_content = escape_typst(content);
        let _ = write!(out, "#footnote[{escaped_content}]");
        return;
    }

    if run.text.contains(PPTX_SOFT_LINE_BREAK_CHAR) {
        write_run_with_soft_line_breaks(out, run);
        return;
    }

    write_run_segment(out, run, &run.text);
}

fn write_run_with_soft_line_breaks(out: &mut String, run: &Run) {
    let mut segment_start: usize = 0;

    for (offset, ch) in run.text.char_indices() {
        if ch != PPTX_SOFT_LINE_BREAK_CHAR {
            continue;
        }

        if segment_start < offset {
            write_run_segment(out, run, &run.text[segment_start..offset]);
        }
        out.push_str("#linebreak()");
        segment_start = offset + ch.len_utf8();
    }

    if segment_start < run.text.len() {
        write_run_segment(out, run, &run.text[segment_start..]);
    }
}

fn write_run_segment(out: &mut String, run: &Run, text: &str) {
    let style = &run.style;
    let escaped = escape_typst(text);

    let has_text_props = has_text_properties(style);
    let needs_underline = matches!(style.underline, Some(true));
    let needs_strike = matches!(style.strikethrough, Some(true));
    let has_link = run.href.is_some();
    let needs_highlight = style.highlight.is_some();
    let needs_super = matches!(style.vertical_align, Some(VerticalTextAlign::Superscript));
    let needs_sub = matches!(style.vertical_align, Some(VerticalTextAlign::Subscript));
    let needs_small_caps = matches!(style.small_caps, Some(true));
    let needs_all_caps = matches!(style.all_caps, Some(true));

    // Apply all-caps text transformation before escaping
    let escaped: String = if needs_all_caps {
        escape_typst(&text.to_uppercase())
    } else {
        escaped
    };

    // Wrap with link (outermost)
    if let Some(ref href) = run.href {
        let _ = write!(out, "#link(\"{href}\")[");
    }

    // Wrap with highlight
    if let Some(ref hl) = style.highlight {
        let _ = write!(out, "#highlight(fill: rgb({}, {}, {}))[", hl.r, hl.g, hl.b);
    }

    // Wrap with decorations
    if needs_strike {
        out.push_str("#strike[");
    }
    if needs_underline {
        out.push_str("#underline[");
    }

    // Wrap with vertical alignment
    if needs_super {
        out.push_str("#super[");
    }
    if needs_sub {
        out.push_str("#sub[");
    }

    // Wrap with small caps
    if needs_small_caps {
        out.push_str("#smallcaps[");
    }

    if has_text_props {
        out.push_str("#text(");
        write_text_params(out, style);
        out.push_str(")[");
        out.push_str(&escaped);
        out.push(']');
    } else {
        // Prevent `](` pattern: when previous output ends with an
        // unescaped `]` and this text starts with `(`, `.`, or `[`,
        // Typst would interpret it as function arguments / method call /
        // trailing content.  Wrap in `#[...]` to keep it in content mode.
        let needs_wrap = !escaped.is_empty()
            && out.ends_with(']')
            && !out.ends_with("\\]")
            && matches!(escaped.as_bytes()[0], b'(' | b'.' | b'[');
        if needs_wrap {
            out.push_str("#[");
            out.push_str(&escaped);
            out.push(']');
        } else {
            out.push_str(&escaped);
        }
    }

    if needs_small_caps {
        out.push(']');
    }
    if needs_sub {
        out.push(']');
    }
    if needs_super {
        out.push(']');
    }
    if needs_underline {
        out.push(']');
    }
    if needs_strike {
        out.push(']');
    }
    if needs_highlight {
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
        || style.letter_spacing.is_some()
}

fn inferred_font_weight(font_family: &str) -> Option<&'static str> {
    let lower = font_family.trim().to_ascii_lowercase();
    if lower.contains("extrabold") || lower.contains("extra bold") {
        Some("extrabold")
    } else if lower.contains("semibold") || lower.contains("semi bold") {
        Some("semibold")
    } else if lower.contains("medium") {
        Some("medium")
    } else if lower.contains("light") {
        Some("light")
    } else {
        None
    }
}

fn font_weight_rank(weight: &str) -> u8 {
    match weight {
        "light" => 1,
        "medium" => 2,
        "semibold" => 3,
        "bold" => 4,
        "extrabold" => 5,
        "black" => 6,
        _ => 0,
    }
}

fn effective_font_weight(style: &TextStyle) -> Option<&'static str> {
    let inferred = style.font_family.as_deref().and_then(inferred_font_weight);
    let explicit = matches!(style.bold, Some(true)).then_some("bold");
    match (explicit, inferred) {
        (Some(explicit), Some(inferred)) => {
            if font_weight_rank(explicit) >= font_weight_rank(inferred) {
                Some(explicit)
            } else {
                Some(inferred)
            }
        }
        (Some(explicit), None) => Some(explicit),
        (None, Some(inferred)) => Some(inferred),
        (None, None) => None,
    }
}

fn write_text_params(out: &mut String, style: &TextStyle) {
    let mut first = true;

    if let Some(ref family) = style.font_family {
        let font_value = super::font_subst::font_with_fallbacks(family);
        write_param(out, &mut first, &format!("font: {font_value}"));
    }
    if let Some(size) = style.font_size {
        write_param(out, &mut first, &format!("size: {}pt", format_f64(size)));
    }
    if let Some(weight) = effective_font_weight(style) {
        write_param(out, &mut first, &format!("weight: \"{weight}\""));
    }
    if matches!(style.italic, Some(true)) {
        write_param(out, &mut first, "style: \"italic\"");
    }
    if let Some(ref color) = style.color {
        write_param(out, &mut first, &format_color(color));
    }
    if let Some(spacing) = style.letter_spacing {
        write_param(
            out,
            &mut first,
            &format!("tracking: {}pt", format_f64(spacing)),
        );
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
/// Also normalizes text to Unicode NFC form to prevent decomposed characters
/// (e.g., Korean NFD jamo, combining diacritics) from causing issues in PDFs.
fn escape_typst(text: &str) -> String {
    let normalized_text: String = text.nfc().collect();
    let mut result = String::with_capacity(normalized_text.len());
    let mut chars = normalized_text.chars().peekable();
    let mut is_first_char = true;

    while let Some(ch) = chars.next() {
        let should_escape_list_prefix: bool = is_first_char
            && matches!(ch, '-' | '+')
            && chars.peek().is_some_and(|next| next.is_whitespace());

        match ch {
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' | '$' | '[' | ']' | '{'
            | '}'
                if !should_escape_list_prefix =>
            {
                result.push('\\');
                result.push(ch);
            }
            _ if should_escape_list_prefix => {
                result.push('\\');
                result.push(ch);
            }
            _ => result.push(ch),
        }

        is_first_char = false;
    }
    result
}

#[cfg(test)]
#[path = "typst_gen_tests.rs"]
mod tests;
