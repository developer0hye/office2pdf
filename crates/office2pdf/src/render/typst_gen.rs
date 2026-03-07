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
    TabLeader, TabStop, Table, TableCell, TablePage, TableRow, TextDirection, TextStyle,
    VerticalTextAlign, WrapMode,
};

use super::font_context::FontSearchContext;

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
    table_depth: usize,
}

impl GenCtx {
    fn new() -> Self {
        Self {
            images: Vec::new(),
            next_image_id: 0,
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
        FixedElementKind::TextBox(blocks) => {
            let _ = writeln!(
                out,
                "#block(width: {}pt, height: {}pt)[",
                format_f64(elem.width),
                format_f64(elem.height),
            );
            for (index, block) in blocks.iter().enumerate() {
                if index > 0 {
                    out.push('\n');
                }
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

/// Generate Typst markup for a chart with improved visual representation.
///
/// Renders charts in a bordered box with title header and type-specific
/// visual representation:
/// - Bar/Column: proportional visual bars
/// - Pie: percentage legend table
/// - Line: data table with trend indicators (↑↓→)
/// - Others: standard data table
fn generate_chart(out: &mut String, chart: &Chart) {
    // Open bordered box
    let _ = writeln!(
        out,
        "#block(stroke: 1pt + rgb(100, 100, 100), radius: 4pt, inset: 10pt, width: 100%)["
    );

    // Chart title in header
    let type_label = match &chart.chart_type {
        ChartType::Bar => "Bar Chart",
        ChartType::Column => "Column Chart",
        ChartType::Line => "Line Chart",
        ChartType::Pie => "Pie Chart",
        ChartType::Area => "Area Chart",
        ChartType::Scatter => "Scatter Chart",
        ChartType::Other(s) => s.as_str(),
    };

    if let Some(ref title) = chart.title {
        let escaped = escape_typst(title);
        let _ = writeln!(
            out,
            "#align(center)[#text(size: 14pt, weight: \"bold\")[{escaped}]]\n"
        );
    }
    let _ = writeln!(
        out,
        "#align(center)[#text(fill: rgb(100, 100, 100))[_{type_label}_]]\n"
    );

    if chart.series.is_empty() {
        out.push_str("]\n");
        return;
    }

    match &chart.chart_type {
        ChartType::Bar | ChartType::Column => generate_chart_bar(out, chart),
        ChartType::Pie => generate_chart_pie(out, chart),
        ChartType::Line => generate_chart_line(out, chart),
        _ => generate_chart_table(out, chart),
    }

    // Close bordered box
    out.push_str("]\n");
}

/// Generate bar chart with proportional visual bars.
fn generate_chart_bar(out: &mut String, chart: &Chart) {
    // Find max value across all series for proportional scaling
    let max_val = chart
        .series
        .iter()
        .flat_map(|s| s.values.iter())
        .copied()
        .fold(0.0_f64, f64::max);
    let max_val = if max_val == 0.0 { 1.0 } else { max_val };

    // Series color palette
    let colors = [
        "rgb(66, 133, 244)",
        "rgb(219, 68, 55)",
        "rgb(244, 180, 0)",
        "rgb(15, 157, 88)",
    ];

    for (row_idx, cat) in chart.categories.iter().enumerate() {
        let escaped_cat = escape_typst(cat);
        let _ = writeln!(out, "#text(weight: \"bold\")[{escaped_cat}]");
        for (s_idx, series) in chart.series.iter().enumerate() {
            let val = series.values.get(row_idx).copied().unwrap_or(0.0);
            let pct = (val / max_val * 100.0).round().min(100.0) as u32;
            let color = colors[s_idx % colors.len()];
            let _ = writeln!(
                out,
                "#box(width: {pct}%, height: 14pt, fill: {color}, radius: 2pt)[#text(size: 8pt, fill: white)[ {}]]",
                format_f64(val)
            );
        }
        let _ = writeln!(out);
    }

    // Legend for multiple series
    if chart.series.len() > 1 {
        let _ = writeln!(out);
        for (i, series) in chart.series.iter().enumerate() {
            let default_name = format!("Series {}", i + 1);
            let name = series.name.as_deref().unwrap_or(&default_name);
            let color = colors[i % colors.len()];
            let _ = writeln!(
                out,
                "#box(width: 10pt, height: 10pt, fill: {color}) #text(size: 9pt)[{name}] "
            );
        }
    }
}

/// Generate pie chart with percentage labels.
fn generate_chart_pie(out: &mut String, chart: &Chart) {
    let series = match chart.series.first() {
        Some(s) => s,
        None => return,
    };

    let total: f64 = series.values.iter().sum();
    let total = if total == 0.0 { 1.0 } else { total };

    let colors = [
        "rgb(66, 133, 244)",
        "rgb(219, 68, 55)",
        "rgb(244, 180, 0)",
        "rgb(15, 157, 88)",
        "rgb(171, 71, 188)",
        "rgb(0, 172, 193)",
    ];

    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: 3,");
    let _ = writeln!(out, "  [*Slice*], [*Value*], [*%*],");

    for (i, cat) in chart.categories.iter().enumerate() {
        let val = series.values.get(i).copied().unwrap_or(0.0);
        let pct = val / total * 100.0;
        let escaped_cat = escape_typst(cat);
        let color = colors[i % colors.len()];
        let _ = writeln!(
            out,
            "  [#box(width: 8pt, height: 8pt, fill: {color}) {escaped_cat}], [{}], [{:.1}%],",
            format_f64(val),
            pct
        );
    }

    let _ = writeln!(out, ")\n");
}

/// Generate line chart with trend indicators.
fn generate_chart_line(out: &mut String, chart: &Chart) {
    let col_count = 1 + chart.series.len();
    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: {col_count},");

    // Header row
    out.push_str("  [*Category*], ");
    for (i, series) in chart.series.iter().enumerate() {
        let default_name = format!("Series {}", i + 1);
        let name = series.name.as_deref().unwrap_or(&default_name);
        let _ = write!(out, "[*{name}*]");
        if i + 1 < chart.series.len() {
            out.push_str(", ");
        }
    }
    out.push_str(",\n");

    // Data rows with trend indicators
    for (row_idx, cat) in chart.categories.iter().enumerate() {
        let escaped_cat = escape_typst(cat);
        let _ = write!(out, "  [{escaped_cat}], ");
        for (s_idx, series) in chart.series.iter().enumerate() {
            let val = series.values.get(row_idx).copied().unwrap_or(0.0);
            let trend = if row_idx > 0 {
                let prev = series.values.get(row_idx - 1).copied().unwrap_or(0.0);
                if val > prev {
                    " ↑"
                } else if val < prev {
                    " ↓"
                } else {
                    " →"
                }
            } else {
                ""
            };
            let _ = write!(out, "[{}{}]", format_f64(val), trend);
            if s_idx + 1 < chart.series.len() {
                out.push_str(", ");
            }
        }
        out.push_str(",\n");
    }

    let _ = writeln!(out, ")\n");
}

/// Generate generic data table for chart types without specialized rendering.
fn generate_chart_table(out: &mut String, chart: &Chart) {
    let col_count = 1 + chart.series.len();
    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: {col_count},");

    // Header row
    out.push_str("  [*Category*], ");
    for (i, series) in chart.series.iter().enumerate() {
        let default_name = format!("Series {}", i + 1);
        let name = series.name.as_deref().unwrap_or(&default_name);
        let _ = write!(out, "[*{name}*]");
        if i + 1 < chart.series.len() {
            out.push_str(", ");
        }
    }
    out.push_str(",\n");

    // Data rows
    for (row_idx, cat) in chart.categories.iter().enumerate() {
        let escaped_cat = escape_typst(cat);
        let _ = write!(out, "  [{escaped_cat}], ");
        for (i, series) in chart.series.iter().enumerate() {
            let val = series.values.get(row_idx).copied().unwrap_or(0.0);
            let _ = write!(out, "[{}]", format_f64(val));
            if i + 1 < chart.series.len() {
                out.push_str(", ");
            }
        }
        out.push_str(",\n");
    }

    let _ = writeln!(out, ")\n");
}

/// Generate Typst markup for a SmartArt diagram.
///
/// Renders SmartArt as a visually distinct bordered box with:
/// - Hierarchy items (varying depths): indented tree with depth-based padding
/// - Flat items (all same depth): numbered steps with arrows
fn generate_smartart(out: &mut String, smartart: &SmartArt, width: f64, height: f64) {
    let _ = writeln!(
        out,
        "#block(width: {}pt, height: {}pt, stroke: 1pt + rgb(70, 130, 180), radius: 4pt, inset: 10pt, fill: rgb(245, 248, 255))[",
        format_f64(width),
        format_f64(height),
    );
    let _ = writeln!(
        out,
        "#align(center)[#text(size: 11pt, weight: \"bold\", fill: rgb(70, 130, 180))[SmartArt Diagram]]\n"
    );

    if smartart.items.is_empty() {
        out.push_str("]\n");
        return;
    }

    // Determine if hierarchy (varying depths) or flat (all same depth)
    let has_hierarchy = smartart.items.iter().any(|n| n.depth > 0);

    if has_hierarchy {
        generate_smartart_hierarchy(out, smartart);
    } else {
        generate_smartart_steps(out, smartart);
    }

    out.push_str("]\n");
}

/// Render hierarchical SmartArt as an indented tree.
fn generate_smartart_hierarchy(out: &mut String, smartart: &SmartArt) {
    for node in &smartart.items {
        let escaped = escape_typst(&node.text);
        if node.depth == 0 {
            let _ = writeln!(out, "#text(weight: \"bold\")[{escaped}]");
        } else {
            let indent = node.depth as f64 * 16.0;
            let _ = writeln!(
                out,
                "#pad(left: {}pt)[{} {escaped}]",
                format_f64(indent),
                if node.depth == 1 { "├" } else { "└" },
            );
        }
    }
}

/// Render flat SmartArt as numbered steps with arrows.
fn generate_smartart_steps(out: &mut String, smartart: &SmartArt) {
    for (i, node) in smartart.items.iter().enumerate() {
        let escaped = escape_typst(&node.text);
        let step_num = i + 1;
        let _ = writeln!(
            out,
            "#box(stroke: 0.5pt + rgb(70, 130, 180), radius: 3pt, inset: 6pt)[#text(weight: \"bold\")[{}. ] {escaped}]",
            step_num,
        );
        if i + 1 < smartart.items.len() {
            let _ = writeln!(out, "#align(center)[#text(size: 14pt)[↓]]");
        }
    }
}

/// Generate Typst markup for a list (ordered or unordered).
///
/// Uses Typst's `#enum()` for ordered lists and `#list()` for unordered lists.
/// Nested items are wrapped in `list.item()` / `enum.item()` with a sub-list.
struct EffectiveListStyle<'a> {
    kind: ListKind,
    numbering_pattern: Option<&'a str>,
    full_numbering: bool,
}

fn list_style_for_level<'a>(list: &'a List, level: u32) -> EffectiveListStyle<'a> {
    if let Some(style) = list.level_styles.get(&level) {
        EffectiveListStyle {
            kind: style.kind,
            numbering_pattern: style.numbering_pattern.as_deref(),
            full_numbering: style.full_numbering,
        }
    } else {
        EffectiveListStyle {
            kind: list.kind,
            numbering_pattern: None,
            full_numbering: false,
        }
    }
}

fn list_funcs(kind: ListKind) -> (&'static str, &'static str) {
    match kind {
        ListKind::Ordered => ("enum", "enum.item"),
        ListKind::Unordered => ("list", "list.item"),
    }
}

fn write_list_open(
    out: &mut String,
    prefix: &str,
    style: &EffectiveListStyle<'_>,
    start_at: Option<u32>,
) {
    let (func, _) = list_funcs(style.kind);
    let _ = write!(out, "{prefix}{func}(");

    if style.kind == ListKind::Ordered {
        if let Some(numbering_pattern) = style.numbering_pattern {
            let _ = write!(
                out,
                "numbering: \"{}\", ",
                escape_typst_string(numbering_pattern)
            );
        }
        if let Some(start_at) = start_at {
            let _ = write!(out, "start: {start_at}, ");
        }
        if style.full_numbering {
            out.push_str("full: true, ");
        }
    }

    out.push('\n');
}

fn generate_list(out: &mut String, list: &List) -> Result<(), ConvertError> {
    let style = list_style_for_level(list, 0);
    let start_at = list.items.first().and_then(|item| item.start_at);
    write_list_open(out, "#", &style, start_at);
    generate_list_items(out, list, &list.items, 0)?;
    out.push_str(")\n");
    Ok(())
}

/// Recursively generate list items, grouping consecutive items at the same or deeper level.
fn generate_list_items(
    out: &mut String,
    list: &List,
    items: &[crate::ir::ListItem],
    base_level: u32,
) -> Result<(), ConvertError> {
    let style = list_style_for_level(list, base_level);
    let (_, item_func) = list_funcs(style.kind);
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

            // Check if next items are nested (deeper level) — they become a sub-list
            let nested_start = i + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].level > base_level {
                nested_end += 1;
            }

            if nested_end > nested_start {
                // Emit nested sub-list inside the same content block
                let nested_style = list_style_for_level(list, base_level + 1);
                let nested_start_at = items[nested_start].start_at;
                write_list_open(out, " #", &nested_style, nested_start_at);
                generate_list_items(out, list, &items[nested_start..nested_end], base_level + 1)?;
                out.push(')');
                i = nested_end;
            } else {
                i += 1;
            }

            out.push_str("],\n");
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
    ctx.table_depth += 1;
    let result = match table.alignment {
        Some(Alignment::Center) => {
            out.push_str("#align(center)[\n");
            let result = generate_table_inner(out, table, ctx);
            out.push_str("]\n");
            result
        }
        Some(Alignment::Right) => {
            out.push_str("#align(right)[\n");
            let result = generate_table_inner(out, table, ctx);
            out.push_str("]\n");
            result
        }
        _ => generate_table_inner(out, table, ctx),
    };
    ctx.table_depth -= 1;
    result
}

fn generate_table_inner(
    out: &mut String,
    table: &Table,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    out.push_str("#table(\n");

    if let Some(padding) = table.default_cell_padding {
        let _ = writeln!(out, "  inset: {},", format_insets(&padding));
    }

    // Determine number of columns
    let num_cols = if !table.column_widths.is_empty() {
        table.column_widths.len()
    } else {
        // Infer from the maximum number of cells in any row
        table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    };

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
    } else if num_cols > 1 {
        // No explicit widths but multiple columns inferred — tell Typst the column count
        let _ = writeln!(out, "  columns: {num_cols},");
    }

    if table.rows.iter().any(|row| row.height.is_some()) {
        out.push_str("  rows: (");
        for (i, row) in table.rows.iter().enumerate() {
            if i > 0 {
                out.push_str(", ");
            }
            match row.height {
                Some(height) => {
                    let _ = write!(out, "{}pt", format_f64(height));
                }
                None => out.push_str("auto"),
            }
        }
        out.push_str("),\n");
    }

    // Rows and cells — clamp colspan to prevent exceeding available columns.
    // Also handle merge continuation cells: col_span=0 (hMerge) and row_span=0
    // (vMerge) are continuation markers that must not be emitted as Typst cells.
    // Track column occupancy from rowspans so we clamp colspans correctly.
    // rowspan_remaining[c] = N means column c is occupied for N more rows.
    let mut rowspan_remaining = vec![0usize; num_cols];
    let header_row_count = table.header_row_count.min(table.rows.len());

    if header_row_count > 0 {
        out.push_str("  table.header(\n");
        generate_table_rows(
            out,
            &table.rows[..header_row_count],
            num_cols,
            &mut rowspan_remaining,
            "    ",
            ctx,
        )?;
        out.push_str("  ),\n");
    }

    generate_table_rows(
        out,
        &table.rows[header_row_count..],
        num_cols,
        &mut rowspan_remaining,
        "  ",
        ctx,
    )?;

    out.push_str(")\n");
    Ok(())
}

fn generate_table_rows(
    out: &mut String,
    rows: &[TableRow],
    num_cols: usize,
    rowspan_remaining: &mut [usize],
    indent: &str,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    for row in rows {
        // Decrement rowspan counters at the start of each row.
        for rs in rowspan_remaining.iter_mut() {
            if *rs > 0 {
                *rs -= 1;
            }
        }

        let mut col_pos: usize = 0;
        for cell in &row.cells {
            if cell.col_span == 0 || cell.row_span == 0 {
                continue;
            }

            while col_pos < num_cols && rowspan_remaining[col_pos] > 0 {
                col_pos += 1;
            }
            if col_pos >= num_cols {
                break;
            }

            let remaining = num_cols - col_pos;
            let clamped_colspan = (cell.col_span as usize).min(remaining).max(1) as u32;
            generate_table_cell(out, cell, clamped_colspan, indent, ctx)?;

            if cell.row_span > 1 {
                for rs in rowspan_remaining
                    .iter_mut()
                    .skip(col_pos)
                    .take(clamped_colspan as usize)
                {
                    *rs = cell.row_span as usize;
                }
            }
            col_pos += clamped_colspan as usize;
        }

        while col_pos < num_cols {
            if rowspan_remaining[col_pos] == 0 {
                let _ = writeln!(out, "{indent}[],");
            }
            col_pos += 1;
        }
    }

    Ok(())
}

fn generate_table_cell(
    out: &mut String,
    cell: &TableCell,
    clamped_colspan: u32,
    indent: &str,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let needs_cell_fn = clamped_colspan > 1
        || cell.row_span > 1
        || cell.border.is_some()
        || cell.background.is_some()
        || cell.vertical_align.is_some()
        || cell.padding.is_some();

    if needs_cell_fn {
        out.push_str(indent);
        out.push_str("table.cell(");
        write_cell_params(out, cell, clamped_colspan);
        out.push_str(")[");
    } else {
        out.push_str(indent);
        out.push('[');
    }

    // Render DataBar: colored box at fill percentage
    if let Some(ref db) = cell.data_bar {
        let pct = db.fill_pct.clamp(0.0, 100.0);
        let _ = write!(
            out,
            "#box(width: 100%, height: 0.8em, fill: rgb(240, 240, 240))[#box(width: {}%, height: 100%, fill: rgb({}, {}, {}))]",
            format_f64(pct),
            db.color.r,
            db.color.g,
            db.color.b,
        );
    }

    // Render IconSet: prepend icon text
    if let Some(ref icon) = cell.icon_text {
        let _ = write!(out, "{} ", icon);
    }

    // Generate cell content
    generate_cell_content(out, &cell.content, ctx)?;

    out.push_str("],\n");
    Ok(())
}

fn write_cell_params(out: &mut String, cell: &TableCell, clamped_colspan: u32) {
    let mut first = true;

    if clamped_colspan > 1 {
        write_param(out, &mut first, &format!("colspan: {clamped_colspan}"));
    }
    if cell.row_span > 1 {
        write_param(out, &mut first, &format!("rowspan: {}", cell.row_span));
    }
    if let Some(ref bg) = cell.background {
        write_param(out, &mut first, &format_color(bg));
    }
    if let Some(ref padding) = cell.padding {
        write_param(
            out,
            &mut first,
            &format!("inset: {}", format_insets(padding)),
        );
    }
    if let Some(ref border) = cell.border {
        let stroke = format_cell_stroke(border);
        if !stroke.is_empty() {
            write_param(out, &mut first, &stroke);
        }
    }
    if let Some(ref va) = cell.vertical_align {
        let align_str: &str = match va {
            CellVerticalAlign::Top => "top",
            CellVerticalAlign::Center => "horizon",
            CellVerticalAlign::Bottom => "bottom",
        };
        write_param(out, &mut first, &format!("align: {align_str}"));
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

fn format_cell_stroke(border: &CellBorder) -> String {
    let mut parts = Vec::with_capacity(4);

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
    let base = format!(
        "{}pt + rgb({}, {}, {})",
        format_f64(side.width),
        side.color.r,
        side.color.g,
        side.color.b
    );
    match side.style {
        BorderLineStyle::Solid | BorderLineStyle::None => base,
        _ => format!(
            "(paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\")",
            side.color.r,
            side.color.g,
            side.color.b,
            format_f64(side.width),
            border_line_style_to_typst(side.style),
        ),
    }
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
            Block::Table(table) => {
                if ctx.table_depth < MAX_TABLE_DEPTH {
                    generate_table(out, table, ctx)?;
                }
                // Silently skip nested tables beyond MAX_TABLE_DEPTH
            }
            Block::Image(img) => generate_image(out, img, ctx),
            Block::FloatingImage(fi) => generate_floating_image(out, fi, ctx),
            Block::FloatingTextBox(ftb) => generate_floating_text_box(out, ftb, ctx)?,
            Block::List(list) => generate_list(out, list)?,
            Block::MathEquation(math) => generate_math_equation(out, math),
            Block::Chart(chart) => generate_chart(out, chart),
            Block::PageBreak | Block::ColumnBreak => {}
        }
    }
    Ok(())
}

/// Generate paragraph content for inside a table cell (runs only, no block wrapper).
fn generate_cell_paragraph(out: &mut String, para: &Paragraph) {
    generate_runs_with_tabs(out, &para.runs, para.style.tab_stops.as_deref());
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
        generate_block(out, block, ctx)?;
    }
    out.push_str("]\n");
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

    let style = &run.style;
    let escaped = escape_typst(&run.text);

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
        escape_typst(&run.text.to_uppercase())
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
    let mut result = String::with_capacity(text.len());
    for ch in text.nfc() {
        match ch {
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' | '$' | '[' | ']' | '{'
            | '}' => {
                result.push('\\');
                result.push(ch);
            }
            _ => result.push(ch),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{
        ChartSeries, ColumnLayout, GradientStop, HeaderFooterParagraph, ListItem, ListKind,
        ListLevelStyle, Metadata, SmartArtNode, StyleSheet,
    };
    use std::collections::BTreeMap;

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
            columns: None,
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
            columns: None,
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
    fn test_generate_letter_spacing() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Spaced text".to_string(),
                style: TextStyle {
                    letter_spacing: Some(2.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("tracking: 2pt"),
            "Expected tracking param in: {result}"
        );
    }

    #[test]
    fn test_generate_letter_spacing_negative() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Condensed".to_string(),
                style: TextStyle {
                    letter_spacing: Some(-0.5),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("tracking: -0.5pt"),
            "Expected negative tracking in: {result}"
        );
    }

    #[test]
    fn test_generate_tab_uses_measured_default_stops() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Name:\tValue".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#context {"),
            "Expected contextual tab rendering in: {result}"
        );
        assert!(
            result.contains("measure(tab_prefix_0).width"),
            "Expected tab spacing to measure the rendered prefix in: {result}"
        );
        assert!(
            result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 36)"),
            "Expected default tabs to advance to the next 36pt stop in: {result}"
        );
        assert!(
            !result.contains("#h(36pt)"),
            "Expected default tabs to avoid a hard-coded 36pt gap in: {result}"
        );
    }

    #[test]
    fn test_generate_tab_uses_next_explicit_stop_and_alignment() {
        use crate::ir::{TabAlignment, TabLeader, TabStop};

        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                tab_stops: Some(vec![
                    TabStop {
                        position: 72.0,
                        alignment: TabAlignment::Left,
                        leader: TabLeader::None,
                    },
                    TabStop {
                        position: 216.0,
                        alignment: TabAlignment::Right,
                        leader: TabLeader::Dot,
                    },
                ]),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Col1\tCol2\tCol3".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("if tab_prefix_width_1 < 72pt"),
            "Expected the first explicit stop to be chosen by measured width in: {result}"
        );
        assert!(
            result.contains("else if tab_prefix_width_2 < 216pt"),
            "Expected the next explicit stop to be selected after the first one in: {result}"
        );
        assert!(
            result.contains("216pt - tab_prefix_width_2 - tab_segment_width_2"),
            "Expected right-aligned tabs to subtract the following segment width in: {result}"
        );
    }

    #[test]
    fn test_generate_tab_falls_back_to_next_default_stop_after_explicit_tabs() {
        use crate::ir::{TabAlignment, TabLeader, TabStop};

        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                tab_stops: Some(vec![TabStop {
                    position: 100.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                }]),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "A\tB\tC".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("if tab_prefix_width_1 < 100pt"),
            "Expected the explicit stop to be used when it is still ahead of the prefix in: {result}"
        );
        assert!(
            result.contains("calc.rem-euclid(tab_prefix_width_2.abs.pt(), 36)"),
            "Expected tabs beyond explicit stops to use the next default stop in: {result}"
        );
    }

    #[test]
    fn test_generate_tab_leader_uses_repeat_fill() {
        use crate::ir::{TabAlignment, TabLeader, TabStop};

        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                tab_stops: Some(vec![TabStop {
                    position: 144.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::Dot,
                }]),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Heading\t12".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("box(width: tab_advance_1, repeat[.])"),
            "Expected dot tab leaders to render with Typst repeat fill in: {result}"
        );
    }

    #[test]
    fn test_generate_decimal_tab_uses_decimal_separator_not_thousands_separator() {
        use crate::ir::{TabAlignment, TabLeader, TabStop};

        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                tab_stops: Some(vec![TabStop {
                    position: 180.0,
                    alignment: TabAlignment::Decimal,
                    leader: TabLeader::None,
                }]),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Total\t1,234.56".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("let tab_decimal_anchor_1 = [1,234]"),
            "Expected decimal alignment to anchor after the thousands group in: {result}"
        );
    }

    #[test]
    fn test_generate_decimal_tab_handles_comma_decimal_locale() {
        use crate::ir::{TabAlignment, TabLeader, TabStop};

        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                tab_stops: Some(vec![TabStop {
                    position: 180.0,
                    alignment: TabAlignment::Decimal,
                    leader: TabLeader::None,
                }]),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Total\t1.234,56".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("let tab_decimal_anchor_1 = [1.234]"),
            "Expected decimal alignment to anchor on the locale decimal separator in: {result}"
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
        assert!(
            result.contains("First paragraph\n\nSecond paragraph"),
            "Expected paragraph break between flow paragraphs in: {result}"
        );
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

    use crate::ir::{BorderSide, CellBorder, Insets, Table, TableCell, TableRow};

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
            ..Table::default()
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
    fn test_table_with_default_cell_padding() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![make_text_cell("Padded")],
                height: None,
            }],
            column_widths: vec![100.0],
            header_row_count: 0,
            alignment: None,
            default_cell_padding: Some(Insets {
                top: 2.0,
                right: 3.0,
                bottom: 1.0,
                left: 4.0,
            }),
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;

        assert!(
            result.contains("inset: (top: 2pt, right: 3pt, bottom: 1pt, left: 4pt)"),
            "Expected table inset in: {result}"
        );
    }

    #[test]
    fn test_table_cell_with_padding_override() {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Inset".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            padding: Some(Insets {
                top: 5.0,
                right: 2.0,
                bottom: 3.0,
                left: 6.0,
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
            header_row_count: 0,
            alignment: None,
            default_cell_padding: Some(Insets {
                top: 1.0,
                right: 2.0,
                bottom: 3.0,
                left: 4.0,
            }),
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;

        assert!(
            result.contains("table.cell(inset: (top: 5pt, right: 2pt, bottom: 3pt, left: 6pt))"),
            "Expected cell inset override in: {result}"
        );
    }

    #[test]
    fn test_table_alignment_center_wraps_table() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![make_text_cell("Centered table")],
                height: None,
            }],
            column_widths: vec![100.0],
            header_row_count: 0,
            alignment: Some(Alignment::Center),
            default_cell_padding: None,
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;

        assert!(
            result.contains("#align(center)["),
            "Expected center wrapper in: {result}"
        );
        assert!(
            result.contains("#table("),
            "Expected table inside wrapper in: {result}"
        );
    }

    #[test]
    fn test_table_with_repeating_header_rows_uses_table_header() {
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![make_text_cell("Header 1"), make_text_cell("Header 2")],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 100.0],
            header_row_count: 1,
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;

        assert!(
            result.contains("table.header("),
            "Expected table.header wrapper in: {result}"
        );
        assert!(
            result.contains("Header 1") && result.contains("Body 1"),
            "Expected header and body cell content in: {result}"
        );
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
            ..Table::default()
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
            ..Table::default()
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
    fn test_table_with_explicit_row_sizes_and_cell_vertical_align() {
        let centered_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Centered".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            vertical_align: Some(CellVerticalAlign::Center),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![centered_cell, make_text_cell("B1")],
                    height: Some(36.0),
                },
                TableRow {
                    cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;

        assert!(
            result.contains("rows: (36pt, auto)"),
            "Expected explicit Typst row sizes in: {result}"
        );
        assert!(
            result.contains("align: horizon"),
            "Expected centered vertical alignment in: {result}"
        );
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
            ..Table::default()
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
            ..Table::default()
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
                    style: BorderLineStyle::Solid,
                }),
                bottom: Some(BorderSide {
                    width: 2.0,
                    color: Color::new(255, 0, 0),
                    style: BorderLineStyle::Solid,
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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
                    style: BorderLineStyle::Solid,
                }),
                bottom: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
                }),
                left: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
                }),
                right: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
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
            ..Table::default()
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
    fn test_table_dashed_border_codegen() {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Dashed".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            border: Some(CellBorder {
                top: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Dashed,
                }),
                bottom: Some(BorderSide {
                    width: 1.0,
                    color: Color::new(255, 0, 0),
                    style: BorderLineStyle::Dotted,
                }),
                left: None,
                right: None,
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("dash: \"dashed\""),
            "Expected dashed dash pattern in: {result}"
        );
        assert!(
            result.contains("dash: \"dotted\""),
            "Expected dotted dash pattern in: {result}"
        );
    }

    #[test]
    fn test_shape_dashed_stroke_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                10.0,
                100.0,
                100.0,
                ShapeKind::Rectangle,
                Some(Color::new(0, 128, 255)),
                Some(BorderSide {
                    width: 2.0,
                    color: Color::black(),
                    style: BorderLineStyle::Dashed,
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("dash: \"dashed\""),
            "Expected dashed stroke in: {}",
            output.source
        );
    }

    #[test]
    fn test_shape_dash_dot_stroke_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                10.0,
                100.0,
                100.0,
                ShapeKind::Ellipse,
                None,
                Some(BorderSide {
                    width: 1.0,
                    color: Color::new(0, 0, 255),
                    style: BorderLineStyle::DashDot,
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("dash: \"dash-dotted\""),
            "Expected dash-dotted stroke in: {}",
            output.source
        );
    }

    #[test]
    fn test_border_line_style_to_typst_mapping() {
        assert_eq!(border_line_style_to_typst(BorderLineStyle::Solid), "solid");
        assert_eq!(
            border_line_style_to_typst(BorderLineStyle::Dashed),
            "dashed"
        );
        assert_eq!(
            border_line_style_to_typst(BorderLineStyle::Dotted),
            "dotted"
        );
        assert_eq!(
            border_line_style_to_typst(BorderLineStyle::DashDot),
            "dash-dotted"
        );
        assert_eq!(
            border_line_style_to_typst(BorderLineStyle::DashDotDot),
            "dash-dotted"
        );
        assert_eq!(
            border_line_style_to_typst(BorderLineStyle::Double),
            "dashed"
        );
        assert_eq!(border_line_style_to_typst(BorderLineStyle::None), "solid");
    }

    #[test]
    fn test_solid_border_no_dash_param() {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Solid".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            border: Some(CellBorder {
                top: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Solid,
                }),
                bottom: None,
                left: None,
                right: None,
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        // Solid borders should use the simple format (no "dash:" parameter)
        assert!(
            !result.contains("dash:"),
            "Solid border should not have dash parameter in: {result}"
        );
        assert!(
            result.contains("1pt + rgb(0, 0, 0)"),
            "Expected simple solid format in: {result}"
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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

    use crate::ir::{ImageCrop, ImageData};

    /// Minimal valid 1x1 red pixel PNG for testing.
    const MINIMAL_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    fn make_quadrant_png() -> Vec<u8> {
        let mut image = image::RgbaImage::new(2, 2);
        image.put_pixel(0, 0, image::Rgba([255, 0, 0, 255]));
        image.put_pixel(1, 0, image::Rgba([0, 255, 0, 255]));
        image.put_pixel(0, 1, image::Rgba([0, 0, 255, 255]));
        image.put_pixel(1, 1, image::Rgba([255, 255, 0, 255]));

        let mut encoded = Cursor::new(Vec::new());
        image::DynamicImage::ImageRgba8(image)
            .write_to(&mut encoded, RasterImageFormat::Png)
            .unwrap();
        encoded.into_inner()
    }

    fn make_image(format: ImageFormat, width: Option<f64>, height: Option<f64>) -> Block {
        Block::Image(ImageData {
            data: MINIMAL_PNG.to_vec(),
            format,
            width,
            height,
            crop: None,
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
    fn test_image_crop_preprocesses_raster_asset() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
            data: make_quadrant_png(),
            format: ImageFormat::Png,
            width: Some(20.0),
            height: Some(20.0),
            crop: Some(ImageCrop {
                left: 0.5,
                top: 0.0,
                right: 0.0,
                bottom: 0.0,
            }),
        })])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output
                .source
                .contains("#image(\"img-0.png\", width: 20pt, height: 20pt)"),
            "Expected original display size in: {}",
            output.source
        );

        let cropped =
            image::load_from_memory_with_format(&output.images[0].data, RasterImageFormat::Png)
                .unwrap()
                .to_rgba8();
        assert_eq!(cropped.dimensions(), (1, 2));
        assert_eq!(cropped.get_pixel(0, 0).0, [0, 255, 0, 255]);
        assert_eq!(cropped.get_pixel(0, 1).0, [255, 255, 0, 255]);
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
            (ImageFormat::Svg, "svg"),
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
            background_gradient: None,
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
                gradient_fill: None,
                stroke,
                rotation_deg: None,
                opacity: None,
                shadow: None,
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
                crop: None,
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
    fn test_fixed_page_text_box_multiple_paragraphs_preserve_breaks() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![FixedElement {
                x: 100.0,
                y: 200.0,
                width: 300.0,
                height: 100.0,
                kind: FixedElementKind::TextBox(vec![
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "First item".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }),
                    Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Second item".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }),
                ]),
            }],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("First item\n\nSecond item"),
            "Expected paragraph break inside fixed text box in: {}",
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
                    style: BorderLineStyle::Solid,
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
                    style: BorderLineStyle::Solid,
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
                    gradient_fill: None,
                    stroke: None,
                    rotation_deg: Some(90.0),
                    opacity: None,
                    shadow: None,
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
                    gradient_fill: None,
                    stroke: None,
                    rotation_deg: None,
                    opacity: Some(0.5),
                    shadow: None,
                }),
            }],
        )]);
        let output = generate_typst(&doc).unwrap();
        // With 50% opacity, the fill color should include alpha
        assert!(
            output.source.contains("rgb(0, 255, 0, 128)"),
            "Expected rgb fill with alpha in: {}",
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
                    gradient_fill: None,
                    stroke: None,
                    rotation_deg: Some(45.0),
                    opacity: Some(0.75),
                    shadow: None,
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
            output.source.contains("rgb(0, 0, 255, 191)"),
            "Expected rgb fill with alpha in: {}",
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
            charts: vec![],
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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
            ..Table::default()
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
        use crate::ir::List;
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
                    start_at: None,
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
                    start_at: None,
                },
            ],
            level_styles: BTreeMap::new(),
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
            columns: None,
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
        use crate::ir::List;
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
                    start_at: Some(3),
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
                    start_at: None,
                },
            ],
            level_styles: BTreeMap::from([(
                0,
                ListLevelStyle {
                    kind: ListKind::Ordered,
                    numbering_pattern: Some("1.".to_string()),
                    full_numbering: false,
                },
            )]),
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
            columns: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#enum("),
            "Expected #enum( in: {}",
            output.source
        );
        assert!(output.source.contains("start: 3"));
        assert!(output.source.contains("numbering: \"1.\""));
        assert!(output.source.contains("Step 1"));
        assert!(output.source.contains("Step 2"));
    }

    #[test]
    fn test_generate_nested_list() {
        use crate::ir::List;
        let list = List {
            kind: ListKind::Ordered,
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
                    start_at: Some(1),
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
                    start_at: None,
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
                    start_at: None,
                },
            ],
            level_styles: BTreeMap::from([
                (
                    0,
                    ListLevelStyle {
                        kind: ListKind::Ordered,
                        numbering_pattern: Some("1.".to_string()),
                        full_numbering: false,
                    },
                ),
                (
                    1,
                    ListLevelStyle {
                        kind: ListKind::Unordered,
                        numbering_pattern: None,
                        full_numbering: false,
                    },
                ),
            ]),
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
            columns: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(output.source.contains("Parent"));
        assert!(output.source.contains("Child"));
        assert!(output.source.contains("Sibling"));
        assert!(output.source.contains("#enum("));
        assert!(
            output.source.contains("#list("),
            "Expected nested #list( in: {}",
            output.source
        );
    }

    #[test]
    fn test_nested_list_single_content_block() {
        use crate::ir::List;
        // A parent item with a nested child must produce a single content block:
        //   list.item[Parent #list(...)]
        // NOT two blocks:
        //   list.item[Parent][#list(...)]
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
                    start_at: None,
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
                    start_at: None,
                },
            ],
            level_styles: BTreeMap::new(),
        };
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(list)],
            header: None,
            footer: None,
            columns: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        // Must NOT have "][#list" — that would be two content blocks
        assert!(
            !output.source.contains("][#list"),
            "Nested list must be in a single content block, not double [...].\nGot: {}",
            output.source
        );
        // Must have the nested list inside the parent item's single block
        assert!(
            output.source.contains(" #list("),
            "Nested list should be inside the parent item's content block.\nGot: {}",
            output.source
        );
    }

    #[test]
    fn test_generate_nested_ordered_list_uses_full_numbering() {
        use crate::ir::List;
        let list = List {
            kind: ListKind::Ordered,
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
                    start_at: Some(1),
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
                    start_at: Some(1),
                },
            ],
            level_styles: BTreeMap::from([
                (
                    0,
                    ListLevelStyle {
                        kind: ListKind::Ordered,
                        numbering_pattern: Some("1.".to_string()),
                        full_numbering: false,
                    },
                ),
                (
                    1,
                    ListLevelStyle {
                        kind: ListKind::Ordered,
                        numbering_pattern: Some("1.a.".to_string()),
                        full_numbering: true,
                    },
                ),
            ]),
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::List(list)])]);
        let output = generate_typst(&doc).unwrap();

        assert!(output.source.contains("full: true"));
        assert!(output.source.contains("numbering: \"1.a.\""));
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
            columns: None,
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
            columns: None,
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
            columns: None,
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

    #[test]
    fn test_generate_typst_inserts_pagebreak_between_flow_pages() {
        let first = Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("First section")],
            header: None,
            footer: None,
            columns: None,
        });
        let second = Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Second section")],
            header: None,
            footer: None,
            columns: None,
        });

        let output = generate_typst(&make_doc(vec![first, second])).unwrap();
        let pagebreak_count = output.source.matches("#pagebreak()").count();

        assert_eq!(
            pagebreak_count, 1,
            "Expected exactly one page break between FlowPages. Got:\n{}",
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
            background_gradient: None,
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
            background_gradient: None,
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
            ..Table::default()
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
            background_gradient: None,
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
            charts: vec![],
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
            charts: vec![],
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
            charts: vec![],
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

    // --- Table page with interleaved charts ---

    #[test]
    fn test_table_page_with_chart_at_row() {
        use crate::ir::{Chart, ChartSeries, ChartType};

        let chart = Chart {
            chart_type: ChartType::Bar,
            title: Some("Sales".to_string()),
            categories: vec!["Q1".to_string(), "Q2".to_string()],
            series: vec![ChartSeries {
                name: Some("Revenue".to_string()),
                values: vec![100.0, 200.0],
            }],
        };

        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: make_simple_table(vec![
                vec!["Row 1"],
                vec!["Row 2"],
                vec!["Row 3"],
                vec!["Row 4"],
                vec!["Row 5"],
            ]),
            header: None,
            footer: None,
            charts: vec![(2, chart)], // Chart after row 2
        });

        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        let src = &output.source;

        // Should contain two separate #table blocks (split at chart position)
        let table_count = src.matches("#table(").count();
        assert_eq!(
            table_count, 2,
            "Expected 2 table segments (split at chart row), got {table_count}"
        );

        // Should contain chart rendering between table segments
        assert!(src.contains("Sales"), "Expected chart title in output");
    }

    #[test]
    fn test_table_page_with_chart_at_end() {
        use crate::ir::{Chart, ChartSeries, ChartType};

        let chart = Chart {
            chart_type: ChartType::Pie,
            title: Some("Pie".to_string()),
            categories: vec!["A".to_string()],
            series: vec![ChartSeries {
                name: None,
                values: vec![100.0],
            }],
        };

        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: make_simple_table(vec![vec!["Data"]]),
            header: None,
            footer: None,
            charts: vec![(u32::MAX, chart)], // Chart at end
        });

        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        let src = &output.source;

        // Table should appear before chart
        let table_pos = src.find("#table(").unwrap();
        let chart_pos = src.find("Pie").unwrap();
        assert!(table_pos < chart_pos, "Table should appear before chart");
    }

    // --- Paper size and landscape override tests ---

    #[test]
    fn test_paper_size_override_letter() {
        use crate::config::PaperSize;

        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
        let options = ConvertOptions {
            paper_size: Some(PaperSize::Letter),
            ..Default::default()
        };
        let output = generate_typst_with_options(&doc, &options).unwrap();
        assert!(
            output.source.contains("width: 612pt"),
            "Expected Letter width 612pt, got: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 792pt"),
            "Expected Letter height 792pt, got: {}",
            output.source
        );
    }

    #[test]
    fn test_landscape_override_swaps_dimensions() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
        let options = ConvertOptions {
            landscape: Some(true),
            ..Default::default()
        };
        let output = generate_typst_with_options(&doc, &options).unwrap();
        // A4 default is 595.28 x 841.89; landscape should swap to 841.89 x 595.28
        assert!(
            output.source.contains("width: 841.89pt"),
            "Expected landscape width 841.89pt, got: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 595.28pt"),
            "Expected landscape height 595.28pt, got: {}",
            output.source
        );
    }

    #[test]
    fn test_portrait_override_keeps_portrait() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
        let options = ConvertOptions {
            landscape: Some(false),
            ..Default::default()
        };
        let output = generate_typst_with_options(&doc, &options).unwrap();
        // A4 is already portrait, should remain unchanged
        assert!(
            output.source.contains("width: 595.28pt"),
            "Expected portrait width, got: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 841.89pt"),
            "Expected portrait height, got: {}",
            output.source
        );
    }

    #[test]
    fn test_paper_size_with_landscape() {
        use crate::config::PaperSize;

        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
        let options = ConvertOptions {
            paper_size: Some(PaperSize::Letter),
            landscape: Some(true),
            ..Default::default()
        };
        let output = generate_typst_with_options(&doc, &options).unwrap();
        // Letter landscape: 792 x 612
        assert!(
            output.source.contains("width: 792pt"),
            "Expected landscape Letter width 792pt, got: {}",
            output.source
        );
        assert!(
            output.source.contains("height: 612pt"),
            "Expected landscape Letter height 612pt, got: {}",
            output.source
        );
    }

    #[test]
    fn test_no_override_uses_original_size() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
        let options = ConvertOptions::default();
        let output = generate_typst_with_options(&doc, &options).unwrap();
        // Default A4 dimensions
        assert!(
            output.source.contains("width: 595.28pt"),
            "Expected A4 width, got: {}",
            output.source
        );
    }

    // ── Floating image codegen tests ──

    #[test]
    fn test_floating_image_square_wrap_codegen() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::FloatingImage(FloatingImage {
                    image: ImageData {
                        data: vec![0x89, 0x50, 0x4E, 0x47],
                        format: ImageFormat::Png,
                        width: Some(200.0),
                        height: Some(100.0),
                        crop: None,
                    },
                    wrap_mode: WrapMode::Square,
                    offset_x: 72.0,
                    offset_y: 36.0,
                })],
                header: None,
                footer: None,
                columns: None,
            })],
            styles: StyleSheet::default(),
        };

        let output = generate_typst(&doc).unwrap();
        // Square wrap should use #place with float: true
        assert!(
            output.source.contains("#place("),
            "Expected #place() for floating image, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("float: true"),
            "Expected float: true for square wrap, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("dx: 72pt"),
            "Expected dx: 72pt, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_floating_image_top_and_bottom_codegen() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::FloatingImage(FloatingImage {
                    image: ImageData {
                        data: vec![0x89, 0x50, 0x4E, 0x47],
                        format: ImageFormat::Png,
                        width: Some(150.0),
                        height: Some(75.0),
                        crop: None,
                    },
                    wrap_mode: WrapMode::TopAndBottom,
                    offset_x: 10.0,
                    offset_y: 0.0,
                })],
                header: None,
                footer: None,
                columns: None,
            })],
            styles: StyleSheet::default(),
        };

        let output = generate_typst(&doc).unwrap();
        // TopAndBottom should use a block with vertical space
        assert!(
            output.source.contains("#block("),
            "Expected #block() for topAndBottom wrap, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("#v(75pt)"),
            "Expected vertical space for image height, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_floating_image_behind_codegen() {
        let doc = Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::FloatingImage(FloatingImage {
                    image: ImageData {
                        data: vec![0x89, 0x50, 0x4E, 0x47],
                        format: ImageFormat::Png,
                        width: Some(100.0),
                        height: Some(50.0),
                        crop: None,
                    },
                    wrap_mode: WrapMode::Behind,
                    offset_x: 0.0,
                    offset_y: 0.0,
                })],
                header: None,
                footer: None,
                columns: None,
            })],
            styles: StyleSheet::default(),
        };

        let output = generate_typst(&doc).unwrap();
        // Behind should use #place without float
        assert!(
            output.source.contains("#place("),
            "Expected #place() for behind wrap, got:\n{}",
            output.source
        );
        assert!(
            !output.source.contains("float: true"),
            "Behind wrap should NOT use float, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_floating_text_box_square_wrap_codegen() {
        let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
            FloatingTextBox {
                content: vec![make_paragraph("Anchored box")],
                wrap_mode: WrapMode::Square,
                width: 200.0,
                height: 100.0,
                offset_x: 72.0,
                offset_y: 36.0,
            },
        )])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#place("),
            "Expected #place() for floating text box, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("float: true"),
            "Expected float: true for square-wrapped text box, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("dx: 72pt"),
            "Expected dx: 72pt, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("width: 200pt"),
            "Expected width: 200pt, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("height: 100pt"),
            "Expected height: 100pt, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Anchored box"),
            "Expected text box content, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_floating_text_box_top_and_bottom_codegen() {
        let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
            FloatingTextBox {
                content: vec![make_paragraph("Top box")],
                wrap_mode: WrapMode::TopAndBottom,
                width: 150.0,
                height: 60.0,
                offset_x: 10.0,
                offset_y: 0.0,
            },
        )])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#block(width: 100%)"),
            "Expected block wrapper for top-and-bottom text box, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("#v(60pt)"),
            "Expected reserved vertical space for text box height, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Top box"),
            "Expected text box content, got:\n{}",
            output.source
        );
    }

    // ── Math equation codegen tests ──

    #[test]
    fn test_codegen_display_math() {
        let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
            MathEquation {
                content: "frac(a, b)".to_string(),
                display: true,
            },
        )])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("$ frac(a, b) $"),
            "Expected display math '$ frac(a, b) $', got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_inline_math() {
        let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
            MathEquation {
                content: "x^2".to_string(),
                display: false,
            },
        )])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("$x^2$"),
            "Expected inline math '$x^2$', got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_complex_math() {
        let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
            MathEquation {
                content: "sum_(i=1)^n i".to_string(),
                display: true,
            },
        )])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("$ sum_(i=1)^n i $"),
            "Expected display math with sum, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_chart_bar_visual_bars() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
            chart_type: ChartType::Bar,
            title: Some("Sales Report".to_string()),
            categories: vec!["Q1".to_string(), "Q2".to_string()],
            series: vec![ChartSeries {
                name: Some("Revenue".to_string()),
                values: vec![100.0, 250.0],
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        // Wrapped in bordered box with header
        assert!(
            output.source.contains("stroke:"),
            "Expected bordered box, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Sales Report"),
            "Expected chart title in header, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Bar Chart"),
            "Expected chart type label, got:\n{}",
            output.source
        );
        // Bar chart should have visual bars (box with proportional width)
        assert!(
            output.source.contains("box(") || output.source.contains("#box("),
            "Expected visual bar boxes for bar chart, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Q1"),
            "Expected category label, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_chart_pie_percentages() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
            chart_type: ChartType::Pie,
            title: Some("Market Share".to_string()),
            categories: vec!["A".to_string(), "B".to_string()],
            series: vec![ChartSeries {
                name: None,
                values: vec![60.0, 40.0],
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("Pie Chart"),
            "Expected pie chart label, got:\n{}",
            output.source
        );
        // Pie chart should show percentages
        assert!(
            output.source.contains("60") && output.source.contains("%"),
            "Expected percentage in pie chart, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("40") && output.source.contains("%"),
            "Expected percentage in pie chart, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_chart_line_trend_indicators() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
            chart_type: ChartType::Line,
            title: Some("Trends".to_string()),
            categories: vec!["Jan".to_string(), "Feb".to_string(), "Mar".to_string()],
            series: vec![ChartSeries {
                name: Some("Sales".to_string()),
                values: vec![10.0, 20.0, 15.0],
            }],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("Line Chart"),
            "Expected line chart label, got:\n{}",
            output.source
        );
        // Line chart should have trend indicators (↑ or ↓)
        let has_trend = output.source.contains('↑')
            || output.source.contains('↓')
            || output.source.contains('→');
        assert!(
            has_trend,
            "Expected trend indicators in line chart, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_codegen_chart_empty_series() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
            chart_type: ChartType::Line,
            title: Some("Empty".to_string()),
            categories: vec![],
            series: vec![],
        })])]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("Line Chart"),
            "Expected line chart label, got:\n{}",
            output.source
        );
    }

    // ── SmartArt codegen tests ──────────────────────────────────────────

    /// Helper to create a SmartArtNode.
    fn sa_node(text: &str, depth: usize) -> SmartArtNode {
        SmartArtNode {
            text: text.to_string(),
            depth,
        }
    }

    #[test]
    fn test_smartart_codegen_flat_numbered_steps() {
        let doc = make_doc(vec![make_fixed_page(
            720.0,
            540.0,
            vec![FixedElement {
                x: 72.0,
                y: 100.0,
                width: 400.0,
                height: 300.0,
                kind: FixedElementKind::SmartArt(SmartArt {
                    items: vec![
                        sa_node("Step 1", 0),
                        sa_node("Step 2", 0),
                        sa_node("Step 3", 0),
                    ],
                }),
            }],
        )]);

        let output = generate_typst(&doc).unwrap();
        // Wrapped in bordered box
        assert!(
            output.source.contains("stroke:"),
            "Expected bordered box, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("SmartArt Diagram"),
            "Expected SmartArt header, got:\n{}",
            output.source
        );
        // Flat items → numbered steps with arrows
        assert!(
            output.source.contains("Step 1"),
            "Expected Step 1, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Step 2"),
            "Expected Step 2, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Step 3"),
            "Expected Step 3, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_smartart_codegen_hierarchy_indented_tree() {
        let doc = make_doc(vec![make_fixed_page(
            720.0,
            540.0,
            vec![FixedElement {
                x: 72.0,
                y: 100.0,
                width: 400.0,
                height: 300.0,
                kind: FixedElementKind::SmartArt(SmartArt {
                    items: vec![
                        sa_node("CEO", 0),
                        sa_node("VP Engineering", 1),
                        sa_node("VP Sales", 1),
                        sa_node("Dev Lead", 2),
                    ],
                }),
            }],
        )]);

        let output = generate_typst(&doc).unwrap();
        // Hierarchical items should use indentation
        assert!(
            output.source.contains("CEO"),
            "Expected CEO, got:\n{}",
            output.source
        );
        // Deeper items should have padding/indentation
        assert!(
            output.source.contains("pad"),
            "Expected indented items for hierarchy, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("VP Engineering"),
            "Expected VP Engineering, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains("Dev Lead"),
            "Expected Dev Lead, got:\n{}",
            output.source
        );
    }

    #[test]
    fn test_smartart_codegen_empty_items() {
        let doc = make_doc(vec![make_fixed_page(
            720.0,
            540.0,
            vec![FixedElement {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
                kind: FixedElementKind::SmartArt(SmartArt { items: vec![] }),
            }],
        )]);

        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("SmartArt Diagram"),
            "Expected SmartArt header even for empty SmartArt"
        );
    }

    #[test]
    fn test_smartart_codegen_special_chars() {
        let doc = make_doc(vec![make_fixed_page(
            720.0,
            540.0,
            vec![FixedElement {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
                kind: FixedElementKind::SmartArt(SmartArt {
                    items: vec![sa_node("Item #1", 0), sa_node("Price $10", 0)],
                }),
            }],
        )]);

        let output = generate_typst(&doc).unwrap();
        // # and $ should be escaped
        assert!(
            output.source.contains(r"\#"),
            "Expected escaped #, got:\n{}",
            output.source
        );
        assert!(
            output.source.contains(r"\$"),
            "Expected escaped $, got:\n{}",
            output.source
        );
    }

    // ── Gradient codegen tests (US-050) ─────────────────────────────────

    #[test]
    fn test_gradient_background_codegen() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: Some(Color::new(255, 0, 0)), // fallback
            background_gradient: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(255, 0, 0),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 255),
                    },
                ],
                angle: 90.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("gradient.linear("),
            "Should contain gradient.linear. Got: {}",
            output.source,
        );
        assert!(
            output.source.contains("(rgb(255, 0, 0), 0%)"),
            "Should contain first stop. Got: {}",
            output.source,
        );
        assert!(
            output.source.contains("(rgb(0, 0, 255), 100%)"),
            "Should contain second stop. Got: {}",
            output.source,
        );
        assert!(
            output.source.contains("angle: 90deg"),
            "Should contain angle. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_gradient_background_no_angle_codegen() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: None,
            background_gradient: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(255, 255, 255),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 0),
                    },
                ],
                angle: 0.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("gradient.linear("),
            "Should contain gradient.linear. Got: {}",
            output.source,
        );
        // Angle of 0 should NOT be emitted
        assert!(
            !output.source.contains("angle:"),
            "Should not contain angle for 0 degrees. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_gradient_shape_fill_codegen() {
        let elem = FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(255, 0, 0)), // fallback
                gradient_fill: Some(GradientFill {
                    stops: vec![
                        GradientStop {
                            position: 0.0,
                            color: Color::new(0, 128, 0),
                        },
                        GradientStop {
                            position: 1.0,
                            color: Color::new(0, 0, 128),
                        },
                    ],
                    angle: 45.0,
                }),
                stroke: None,
                rotation_deg: None,
                opacity: None,
                shadow: None,
            }),
        };
        let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("gradient.linear("),
            "Should contain gradient.linear for shape. Got: {}",
            output.source,
        );
        assert!(
            output.source.contains("(rgb(0, 128, 0), 0%)"),
            "Should contain first stop. Got: {}",
            output.source,
        );
        // Should NOT contain the fallback rgb fill since gradient takes precedence
        assert!(
            !output.source.contains("fill: rgb(255, 0, 0)"),
            "Should not contain fallback solid fill. Got: {}",
            output.source,
        );
    }

    // ── Shadow codegen tests ──────────────────────────────────────────

    #[test]
    fn test_shape_shadow_codegen() {
        use crate::ir::Shadow;
        let elem = FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(255, 0, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: None,
                shadow: Some(Shadow {
                    blur_radius: 4.0,
                    distance: 3.0,
                    direction: 45.0,
                    color: Color::new(0, 0, 0),
                    opacity: 0.5,
                }),
            }),
        };
        let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
        let output = generate_typst(&doc).unwrap();
        // Shadow should render as an offset duplicate with rgb fill (4 args for alpha)
        assert!(
            output.source.contains("rgb(0, 0, 0, 128)"),
            "Shadow should use rgb with alpha. Got: {}",
            output.source,
        );
        // The shadow shape should be placed before the main shape
        let shadow_pos = output.source.find("rgb(0, 0, 0, 128)");
        let main_pos = output.source.find("rgb(255, 0, 0)");
        assert!(
            shadow_pos < main_pos,
            "Shadow should appear before main shape in output",
        );
    }

    #[test]
    fn test_shape_no_shadow_no_extra_output() {
        let elem = FixedElement {
            x: 10.0,
            y: 20.0,
            width: 200.0,
            height: 150.0,
            kind: FixedElementKind::Shape(Shape {
                kind: ShapeKind::Rectangle,
                fill: Some(Color::new(255, 0, 0)),
                gradient_fill: None,
                stroke: None,
                rotation_deg: None,
                opacity: None,
                shadow: None,
            }),
        };
        let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
        let output = generate_typst(&doc).unwrap();
        // No shadow → no rgb(0, 0, 0, ...) for shadow color
        assert!(
            !output.source.contains("rgb(0, 0, 0,"),
            "No shadow should produce no rgb shadow. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_gradient_prefers_over_solid_fill() {
        // When both gradient_fill and fill are present, gradient should be used
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: Some(Color::new(128, 128, 128)),
            background_gradient: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(255, 0, 0),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 255),
                    },
                ],
                angle: 180.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        // Gradient should be used, not the solid fallback
        assert!(
            output.source.contains("gradient.linear("),
            "Gradient should be preferred. Got: {}",
            output.source,
        );
        assert!(
            !output.source.contains("fill: rgb(128, 128, 128)"),
            "Solid fallback should not appear. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_gradient_unsorted_stops_rendered_in_sorted_order() {
        // Gradient stops provided in reverse order should be sorted by position
        // before rendering — Typst requires monotonic offsets.
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: None,
            background_gradient: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 255),
                    },
                    GradientStop {
                        position: 0.5,
                        color: Color::new(0, 255, 0),
                    },
                    GradientStop {
                        position: 0.0,
                        color: Color::new(255, 0, 0),
                    },
                ],
                angle: 90.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        // Stops should appear in order: 0% (red), 50% (green), 100% (blue)
        let src = &output.source;
        let pos_red = src.find("(rgb(255, 0, 0), 0%)").expect("red stop missing");
        let pos_green = src
            .find("(rgb(0, 255, 0), 50%)")
            .expect("green stop missing");
        let pos_blue = src
            .find("(rgb(0, 0, 255), 100%)")
            .expect("blue stop missing");
        assert!(
            pos_red < pos_green && pos_green < pos_blue,
            "Stops should be in sorted order (0% < 50% < 100%). Got: {}",
            src,
        );
    }

    // ── DataBar / IconSet codegen tests ──────────────────────────────

    #[test]
    fn test_data_bar_codegen() {
        use crate::ir::DataBarInfo;
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "50".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            data_bar: Some(DataBarInfo {
                color: Color::new(0x63, 0x8E, 0xC6),
                fill_pct: 50.0,
            }),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table,
            header: None,
            footer: None,
            charts: vec![],
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("fill: rgb(99, 142, 198)"),
            "DataBar should contain bar color fill. Got: {}",
            output.source,
        );
        assert!(
            output.source.contains("width: 50%"),
            "DataBar should contain 50% width. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_icon_text_codegen() {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "90".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            icon_text: Some("↑".to_string()),
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let page = Page::Table(TablePage {
            name: "Sheet1".to_string(),
            size: PageSize::default(),
            margins: Margins::default(),
            table,
            header: None,
            footer: None,
            charts: vec![],
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("↑"),
            "Icon text should appear in output. Got: {}",
            output.source,
        );
    }

    #[test]
    fn test_table_colspan_clamped_to_available_columns() {
        // Table with 2 columns, but cell has col_span: 3 (exceeds available).
        // The codegen should clamp it to 2.
        let wide_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Wide".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            col_span: 3,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![wide_cell],
                    height: None,
                },
                TableRow {
                    cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                    height: None,
                },
            ],
            column_widths: vec![100.0, 200.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        // colspan should be clamped to 2 (number of columns), not 3
        assert!(
            result.contains("colspan: 2"),
            "Expected colspan clamped to 2, got: {result}"
        );
        assert!(
            !result.contains("colspan: 3"),
            "colspan: 3 should have been clamped, got: {result}"
        );
    }

    #[test]
    fn test_table_colspan_clamped_mid_row() {
        // Table with 3 columns, row has cell at col 1 + cell with col_span: 3 at col 2.
        // col_span should be clamped to 2 (3 - 1 = 2 remaining columns).
        let normal_cell = make_text_cell("A1");
        let wide_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Wide".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            col_span: 3,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![TableRow {
                cells: vec![normal_cell, wide_cell],
                height: None,
            }],
            column_widths: vec![100.0, 100.0, 100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        // At col position 1, col_span 3 exceeds 3 columns → clamped to 2
        assert!(
            result.contains("colspan: 2"),
            "Expected colspan clamped to 2, got: {result}"
        );
    }

    #[test]
    fn test_table_colspan_no_column_widths_inferred() {
        // Table without explicit column_widths — num_cols inferred from max cells in a row.
        let wide_cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Wide".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            col_span: 5,
            ..TableCell::default()
        };
        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![wide_cell],
                    height: None,
                },
                TableRow {
                    cells: vec![
                        make_text_cell("A"),
                        make_text_cell("B"),
                        make_text_cell("C"),
                    ],
                    height: None,
                },
            ],
            column_widths: vec![],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        // Inferred num_cols = 3 (max cells in any row), col_span 5 clamped to 3
        assert!(
            result.contains("colspan: 3"),
            "Expected colspan clamped to 3 (inferred columns), got: {result}"
        );
        assert!(
            !result.contains("colspan: 5"),
            "colspan: 5 should have been clamped, got: {result}"
        );
    }

    // ── Metadata codegen tests ─────────────────────────────────────────

    #[test]
    fn test_generate_typst_with_metadata_title_and_author() {
        let doc = Document {
            metadata: Metadata {
                title: Some("Test Title".to_string()),
                author: Some("Test Author".to_string()),
                ..Default::default()
            },
            pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Hello".to_string(),
                    style: TextStyle::default(),
                    footnote: None,
                    href: None,
                }],
                style: ParagraphStyle::default(),
            })])],
            styles: StyleSheet::default(),
        };
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#set document(title: \"Test Title\", author: \"Test Author\")"),
            "Expected document metadata in Typst output, got: {result}"
        );
    }

    #[test]
    fn test_generate_typst_with_metadata_title_only() {
        let doc = Document {
            metadata: Metadata {
                title: Some("Only Title".to_string()),
                ..Default::default()
            },
            pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Hello".to_string(),
                    style: TextStyle::default(),
                    footnote: None,
                    href: None,
                }],
                style: ParagraphStyle::default(),
            })])],
            styles: StyleSheet::default(),
        };
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#set document(title: \"Only Title\")"),
            "Expected title-only metadata in Typst output, got: {result}"
        );
    }

    #[test]
    fn test_generate_typst_without_metadata() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            !result.contains("#set document("),
            "Should not emit #set document when no metadata, got: {result}"
        );
    }

    #[test]
    fn test_generate_typst_with_metadata_created_date() {
        let doc = Document {
            metadata: Metadata {
                title: Some("Dated Doc".to_string()),
                created: Some("2024-06-15T10:30:00Z".to_string()),
                ..Default::default()
            },
            pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Hello".to_string(),
                    style: TextStyle::default(),
                    footnote: None,
                    href: None,
                }],
                style: ParagraphStyle::default(),
            })])],
            styles: StyleSheet::default(),
        };
        let result = generate_typst(&doc).unwrap().source;
        // When metadata has a created date, it should be emitted in Typst
        assert!(
            result.contains("date: datetime(year: 2024, month: 6, day: 15"),
            "Expected document date from metadata created field, got: {result}"
        );
    }

    #[test]
    fn test_generate_typst_with_metadata_date_only() {
        // When only the created date is set (no title/author), date should still appear
        let doc = Document {
            metadata: Metadata {
                created: Some("2023-12-25T08:00:00Z".to_string()),
                ..Default::default()
            },
            pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Hello".to_string(),
                    style: TextStyle::default(),
                    footnote: None,
                    href: None,
                }],
                style: ParagraphStyle::default(),
            })])],
            styles: StyleSheet::default(),
        };
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("date: datetime(year: 2023, month: 12, day: 25"),
            "Expected document date even without title/author, got: {result}"
        );
    }

    #[test]
    fn test_generate_typst_with_invalid_created_date() {
        // Invalid date string should be silently ignored
        let doc = Document {
            metadata: Metadata {
                title: Some("Bad Date Doc".to_string()),
                created: Some("not-a-date".to_string()),
                ..Default::default()
            },
            pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Hello".to_string(),
                    style: TextStyle::default(),
                    footnote: None,
                    href: None,
                }],
                style: ParagraphStyle::default(),
            })])],
            styles: StyleSheet::default(),
        };
        let result = generate_typst(&doc).unwrap().source;
        // Invalid date should not crash or produce a date field
        assert!(
            !result.contains("date: datetime("),
            "Invalid date should not produce document date, got: {result}"
        );
    }

    #[test]
    fn test_parse_iso8601_date_full() {
        let result = parse_iso8601_date("2024-06-15T10:30:45Z");
        assert_eq!(result, Some((2024, 6, 15, 10, 30, 45)));
    }

    #[test]
    fn test_parse_iso8601_date_date_only() {
        let result = parse_iso8601_date("2023-12-25");
        assert_eq!(result, Some((2023, 12, 25, 0, 0, 0)));
    }

    #[test]
    fn test_parse_iso8601_date_invalid() {
        assert_eq!(parse_iso8601_date("not-a-date"), None);
        assert_eq!(parse_iso8601_date(""), None);
        assert_eq!(parse_iso8601_date("2024"), None);
        assert_eq!(parse_iso8601_date("2024-13-01T00:00:00Z"), None); // month > 12
        assert_eq!(parse_iso8601_date("2024-00-01T00:00:00Z"), None); // month 0
    }

    // ── Extended geometry codegen tests (US-085) ──────────────────────────

    #[test]
    fn test_triangle_polygon_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                20.0,
                200.0,
                150.0,
                ShapeKind::Polygon {
                    vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
                },
                Some(Color::new(255, 0, 0)),
                None,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#polygon("),
            "Expected #polygon in: {}",
            output.source
        );
        // Check vertex at top-center: 0.5 * 200 = 100pt
        assert!(
            output.source.contains("100pt"),
            "Expected 100pt vertex x in: {}",
            output.source
        );
        assert!(
            output.source.contains("fill: rgb(255, 0, 0)"),
            "Expected fill in: {}",
            output.source
        );
    }

    #[test]
    fn test_rounded_rectangle_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                10.0,
                20.0,
                200.0,
                100.0,
                ShapeKind::RoundedRectangle {
                    radius_fraction: 0.1,
                },
                Some(Color::new(0, 0, 255)),
                None,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#rect("),
            "Expected #rect in: {}",
            output.source
        );
        assert!(
            output.source.contains("radius:"),
            "Expected radius parameter in: {}",
            output.source
        );
        // Radius: 0.1 * min(200, 100) = 10pt
        assert!(
            output.source.contains("radius: 10pt"),
            "Expected radius: 10pt in: {}",
            output.source
        );
    }

    #[test]
    fn test_arrow_polygon_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                0.0,
                0.0,
                300.0,
                150.0,
                ShapeKind::Polygon {
                    vertices: vec![
                        (0.0, 0.25),
                        (0.6, 0.25),
                        (0.6, 0.0),
                        (1.0, 0.5),
                        (0.6, 1.0),
                        (0.6, 0.75),
                        (0.0, 0.75),
                    ],
                },
                Some(Color::new(255, 136, 0)),
                None,
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#polygon("),
            "Expected #polygon for arrow in: {}",
            output.source
        );
        // Arrow tip at x=1.0*300=300pt, y=0.5*150=75pt
        assert!(
            output.source.contains("300pt"),
            "Expected 300pt (arrow tip) in: {}",
            output.source
        );
    }

    #[test]
    fn test_polygon_with_stroke_codegen() {
        let doc = make_doc(vec![make_fixed_page(
            960.0,
            540.0,
            vec![make_shape_element(
                0.0,
                0.0,
                100.0,
                100.0,
                ShapeKind::Polygon {
                    vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
                },
                None,
                Some(BorderSide {
                    width: 2.0,
                    color: Color::new(0, 0, 0),
                    style: BorderLineStyle::Solid,
                }),
            )],
        )]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("#polygon("),
            "Expected #polygon in: {}",
            output.source
        );
        assert!(
            output.source.contains("stroke: 2pt + rgb(0, 0, 0)"),
            "Expected stroke in: {}",
            output.source
        );
    }

    #[test]
    fn test_font_substitution_calibri_produces_fallback_list() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Calibri text".to_string(),
                style: TextStyle {
                    font_family: Some("Calibri".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"font: ("Calibri", "Carlito", "Liberation Sans")"#),
            "Expected font fallback list for Calibri in: {result}"
        );
    }

    #[test]
    fn test_font_substitution_arial_produces_fallback_list() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Arial text".to_string(),
                style: TextStyle {
                    font_family: Some("Arial".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"font: ("Arial", "Liberation Sans", "Arimo")"#),
            "Expected font fallback list for Arial in: {result}"
        );
    }

    #[test]
    fn test_font_substitution_unknown_font_no_fallback() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Custom text".to_string(),
                style: TextStyle {
                    font_family: Some("Helvetica".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"font: "Helvetica""#),
            "Unknown font should use simple quoted string in: {result}"
        );
        // Should NOT contain parenthesized array
        assert!(
            !result.contains("font: (\"Helvetica\""),
            "Unknown font should not use array syntax in: {result}"
        );
    }

    #[test]
    fn test_font_substitution_times_new_roman() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "TNR text".to_string(),
                style: TextStyle {
                    font_family: Some("Times New Roman".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"font: ("Times New Roman", "Liberation Serif", "Tinos")"#),
            "Expected font fallback list for Times New Roman in: {result}"
        );
    }

    #[test]
    fn test_font_family_infers_medium_weight_from_family_name() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Title".to_string(),
                style: TextStyle {
                    font_family: Some("Pretendard Medium".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"weight: "medium""#),
            "Expected medium weight inferred from family name in: {result}"
        );
    }

    #[test]
    fn test_font_family_infers_extrabold_weight_from_family_name() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Heading".to_string(),
                style: TextStyle {
                    font_family: Some("Pretendard ExtraBold".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains(r#"weight: "extrabold""#),
            "Expected extrabold weight inferred from family name in: {result}"
        );
    }

    #[test]
    fn test_generate_typst_prefers_office_font_fallback_order_when_context_present() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Title".to_string(),
                style: TextStyle {
                    font_family: Some("Pretendard".to_string()),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let context = FontSearchContext::for_test(
            Vec::new(),
            &["Apple SD Gothic Neo", "Malgun Gothic"],
            &["Malgun Gothic"],
            &[],
        );

        let output = generate_typst_with_options_and_font_context(
            &doc,
            &ConvertOptions::default(),
            Some(&context),
        )
        .unwrap();

        assert!(
            output
                .source
                .contains(r#"font: ("Pretendard", "Malgun Gothic", "Apple SD Gothic Neo""#),
            "Office-managed font should be emitted ahead of system fallback: {}",
            output.source
        );
    }

    // --- Heading level codegen tests (US-096) ---

    #[test]
    fn test_generate_heading_level_1() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(1),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Main Title".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#heading(level: 1)[Main Title]"),
            "H1 paragraph should emit #heading(level: 1): {result}"
        );
    }

    #[test]
    fn test_generate_heading_level_2() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(2),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Sub Section".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#heading(level: 2)[Sub Section]"),
            "H2 paragraph should emit #heading(level: 2): {result}"
        );
    }

    #[test]
    fn test_generate_heading_levels_3_to_6() {
        for level in 3..=6u8 {
            let text = format!("Heading {level}");
            let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    heading_level: Some(level),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: text.clone(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })])]);
            let result = generate_typst(&doc).unwrap().source;
            let expected = format!("#heading(level: {level})[{text}]");
            assert!(
                result.contains(&expected),
                "H{level} should emit {expected}: {result}"
            );
        }
    }

    #[test]
    fn test_generate_heading_with_styled_run() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(1),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Styled Heading".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    font_size: Some(24.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#heading(level: 1)"),
            "Heading with styling should still emit #heading: {result}"
        );
    }

    #[test]
    fn test_generate_regular_paragraph_no_heading() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Normal text".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            !result.contains("#heading"),
            "Regular paragraph should not emit #heading: {result}"
        );
    }

    // ── Unicode NFC normalization tests ──────────────────────────────

    #[test]
    fn test_escape_typst_normalizes_korean_nfd_to_nfc() {
        // Korean "한글" in NFD (decomposed jamo): ㅎ + ㅏ + ㄴ + ㄱ + ㅡ + ㄹ
        let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
        let nfc_korean = "한글";
        let result = escape_typst(nfd_korean);
        assert_eq!(
            result, nfc_korean,
            "NFD Korean jamo should be normalized to composed hangul"
        );
    }

    #[test]
    fn test_escape_typst_normalizes_combining_diacritics() {
        // "café" with combining acute accent (NFD): 'e' + combining acute
        let nfd_cafe = "cafe\u{0301}";
        let nfc_cafe = "caf\u{00E9}"; // é as precomposed
        let result = escape_typst(nfd_cafe);
        assert_eq!(
            result, nfc_cafe,
            "Combining diacritics should be normalized to NFC"
        );
    }

    #[test]
    fn test_escape_typst_nfc_with_special_chars() {
        // NFD text with Typst special chars: "café $5" with combining accent
        let nfd_input = "cafe\u{0301} \\$5";
        let result = escape_typst(nfd_input);
        // NFC normalization + Typst escaping
        assert!(
            result.contains("caf\u{00E9}"),
            "Should contain NFC-normalized é: {result}"
        );
        assert!(
            result.contains("\\$"),
            "Should still escape $ sign: {result}"
        );
    }

    #[test]
    fn test_generate_typst_nfc_korean_in_paragraph() {
        // NFD Korean in a full paragraph through the pipeline
        let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_korean)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("한글"),
            "Generated Typst should contain NFC-composed Korean: {result}"
        );
        assert!(
            !result.contains('\u{1112}'),
            "Generated Typst should not contain decomposed jamo: {result}"
        );
    }

    #[test]
    fn test_generate_typst_nfc_diacritics_in_paragraph() {
        // NFD "résumé" through the full pipeline
        let nfd_resume = "re\u{0301}sume\u{0301}";
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_resume)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("r\u{00E9}sum\u{00E9}"),
            "Generated Typst should contain NFC-composed résumé: {result}"
        );
    }

    #[test]
    fn test_escape_typst_already_nfc_unchanged() {
        // Already NFC text should pass through unchanged (minus Typst escaping)
        let nfc_text = "Hello 한글 café";
        let result = escape_typst(nfc_text);
        assert_eq!(result, nfc_text, "Already-NFC text should be unchanged");
    }

    // --- US-103: Multi-column section layout codegen tests ---

    #[test]
    fn test_generate_flow_page_with_equal_columns() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Column text")],
            header: None,
            footer: None,
            columns: Some(ColumnLayout {
                num_columns: 2,
                spacing: 36.0,
                column_widths: None,
            }),
        })]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#columns(2, gutter: 36pt)"),
            "Should contain columns() call. Got: {result}"
        );
        assert!(
            result.contains("Column text"),
            "Should contain the text content. Got: {result}"
        );
    }

    #[test]
    fn test_generate_flow_page_with_three_columns() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Three col text")],
            header: None,
            footer: None,
            columns: Some(ColumnLayout {
                num_columns: 3,
                spacing: 18.0,
                column_widths: None,
            }),
        })]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#columns(3, gutter: 18pt)"),
            "Should contain columns(3, ...). Got: {result}"
        );
    }

    #[test]
    fn test_generate_flow_page_with_unequal_columns() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Unequal col text")],
            header: None,
            footer: None,
            columns: Some(ColumnLayout {
                num_columns: 2,
                spacing: 36.0,
                column_widths: Some(vec![300.0, 150.0]),
            }),
        })]);
        let result = generate_typst(&doc).unwrap().source;
        // Unequal columns should use grid() with explicit widths
        assert!(
            result.contains("#grid(columns: (300pt, 150pt)"),
            "Unequal columns should use grid(). Got: {result}"
        );
    }

    #[test]
    fn test_generate_column_break() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![
                make_paragraph("Before break"),
                Block::ColumnBreak,
                make_paragraph("After break"),
            ],
            header: None,
            footer: None,
            columns: Some(ColumnLayout {
                num_columns: 2,
                spacing: 36.0,
                column_widths: None,
            }),
        })]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#colbreak()"),
            "Should contain colbreak(). Got: {result}"
        );
    }

    #[test]
    fn test_generate_no_columns_no_wrapper() {
        // Without column layout, content should not be wrapped in columns()
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Normal text")])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            !result.contains("#columns("),
            "Should not contain columns(). Got: {result}"
        );
        assert!(
            !result.contains("#grid(columns:"),
            "Should not contain grid(columns:). Got: {result}"
        );
    }

    // ── BiDi / RTL codegen tests ──────────────────────────────────────

    #[test]
    fn test_generate_rtl_paragraph() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                direction: Some(TextDirection::Rtl),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "مرحبا بالعالم".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#set text(dir: rtl)"),
            "RTL paragraph should emit #set text(dir: rtl). Got: {result}"
        );
    }

    #[test]
    fn test_generate_ltr_paragraph_no_direction() {
        // Normal LTR paragraph should NOT emit any text direction
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            !result.contains("dir: rtl"),
            "LTR paragraph should not emit dir: rtl. Got: {result}"
        );
    }

    #[test]
    fn test_generate_mixed_rtl_ltr_paragraphs() {
        let doc = make_doc(vec![make_flow_page(vec![
            Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    direction: Some(TextDirection::Rtl),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "مرحبا 123".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
            make_paragraph("English text"),
        ])]);
        let result = generate_typst(&doc).unwrap().source;
        // Should contain RTL setting for the Arabic paragraph
        assert!(
            result.contains("#set text(dir: rtl)"),
            "Should contain RTL direction for Arabic paragraph. Got: {result}"
        );
        // The Arabic text and English text should both appear
        assert!(result.contains("مرحبا 123"), "Arabic text should appear");
        assert!(
            result.contains("English text"),
            "English text should appear"
        );
    }

    // --- US-204: Codegen/render robustness tests ---

    #[test]
    fn test_codegen_robustness_zero_pages() {
        // An empty document with zero pages should produce valid Typst output
        let doc = make_doc(vec![]);
        let output = generate_typst(&doc).unwrap();
        // Should produce an empty (or near-empty) source without panicking
        assert!(output.images.is_empty());
    }

    #[test]
    fn test_codegen_robustness_flow_page_empty_content() {
        // A flow page with no content blocks should not panic
        let doc = make_doc(vec![make_flow_page(vec![])]);
        let output = generate_typst(&doc).unwrap();
        assert!(!output.source.is_empty());
    }

    #[test]
    fn test_generate_fixed_page_empty_elements() {
        // A fixed page with no elements should not panic
        let doc = make_doc(vec![Page::Fixed(FixedPage {
            size: PageSize::default(),
            elements: vec![],
            background_color: None,
            background_gradient: None,
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(!output.source.is_empty());
    }

    #[test]
    fn test_generate_table_page_empty_rows() {
        // A table page with zero rows should not panic
        let doc = make_doc(vec![Page::Table(TablePage {
            name: String::new(),
            size: PageSize::default(),
            margins: Margins::default(),
            table: Table {
                rows: vec![],
                column_widths: vec![],
                ..Table::default()
            },
            header: None,
            footer: None,
            charts: vec![],
        })]);
        let output = generate_typst(&doc).unwrap();
        assert!(!output.source.is_empty());
    }

    #[test]
    fn test_generate_paragraph_all_alignment_variants() {
        // All alignment variants (Left, Center, Right, Justify, None) should
        // produce valid Typst output without panicking.
        for alignment in [
            Some(Alignment::Left),
            Some(Alignment::Center),
            Some(Alignment::Right),
            Some(Alignment::Justify),
            None,
        ] {
            let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment,
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: format!("Alignment: {alignment:?}"),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })])]);
            let output = generate_typst(&doc);
            assert!(
                output.is_ok(),
                "Codegen should not fail for alignment {alignment:?}"
            );
        }
    }

    #[test]
    fn test_generate_shape_shadow_all_kinds() {
        // Shadow generation should handle all ShapeKind variants without panicking.
        let shadow = Shadow {
            blur_radius: 4.0,
            color: Color { r: 0, g: 0, b: 0 },
            opacity: 0.5,
            direction: 45.0,
            distance: 3.0,
        };

        let shape_kinds = vec![
            ShapeKind::Rectangle,
            ShapeKind::Ellipse,
            ShapeKind::Line { x2: 100.0, y2: 0.0 },
            ShapeKind::RoundedRectangle {
                radius_fraction: 0.1,
            },
            ShapeKind::Polygon {
                vertices: vec![(0.0, 0.0), (1.0, 0.0), (0.5, 1.0)],
            },
        ];

        for kind in shape_kinds {
            let doc = make_doc(vec![Page::Fixed(FixedPage {
                size: PageSize {
                    width: 960.0,
                    height: 540.0,
                },
                elements: vec![FixedElement {
                    x: 100.0,
                    y: 100.0,
                    width: 200.0,
                    height: 100.0,
                    kind: FixedElementKind::Shape(Shape {
                        kind: kind.clone(),
                        fill: Some(Color { r: 255, g: 0, b: 0 }),
                        gradient_fill: None,
                        stroke: None,
                        opacity: None,
                        shadow: Some(shadow.clone()),
                        rotation_deg: None,
                    }),
                }],
                background_color: None,
                background_gradient: None,
            })]);
            let output = generate_typst(&doc);
            assert!(
                output.is_ok(),
                "Codegen should not panic for shape kind {kind:?} with shadow"
            );
        }
    }

    #[test]
    fn test_column_break_with_empty_content() {
        // Column breaks on empty content should not panic
        let segments = split_at_column_breaks(&[]);
        assert_eq!(segments.len(), 1);
        assert!(segments[0].is_empty());
    }

    #[test]
    fn test_column_break_only_breaks() {
        // Content consisting only of column breaks should not panic
        let blocks = vec![Block::ColumnBreak, Block::ColumnBreak];
        let segments = split_at_column_breaks(&blocks);
        assert_eq!(segments.len(), 3);
        assert!(segments.iter().all(|s| s.is_empty()));
    }

    // --- US-315: text escaping for Typst-significant characters ---

    #[test]
    fn test_escape_typst_backslash() {
        assert_eq!(escape_typst("path\\to\\file"), "path\\\\to\\\\file");
    }

    #[test]
    fn test_escape_typst_hash() {
        assert_eq!(escape_typst("#hashtag"), "\\#hashtag");
    }

    #[test]
    fn test_escape_typst_dollar() {
        assert_eq!(escape_typst("$100"), "\\$100");
    }

    #[test]
    fn test_escape_typst_brackets() {
        assert_eq!(escape_typst("[content]"), "\\[content\\]");
    }

    #[test]
    fn test_escape_typst_braces() {
        assert_eq!(escape_typst("{code}"), "\\{code\\}");
    }

    #[test]
    fn test_escape_typst_all_special_chars() {
        let input = r"#*_`<>@\~/$[]{}";
        let result = escape_typst(input);
        // Every character should be escaped
        assert_eq!(result, "\\#\\*\\_\\`\\<\\>\\@\\\\\\~\\/\\$\\[\\]\\{\\}");
    }

    #[test]
    fn test_escape_typst_in_paragraph_output() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
            "Price: $100 path\\to",
        )])]);
        let output = generate_typst(&doc).unwrap().source;
        assert!(
            output.contains("\\$100"),
            "Dollar sign should be escaped in output: {output}"
        );
        assert!(
            output.contains("path\\\\to"),
            "Backslash should be escaped in output: {output}"
        );
    }

    // --- US-316: single-stop gradient fallback ---

    #[test]
    fn test_gradient_single_stop_fallback_to_solid() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: None,
            background_gradient: Some(GradientFill {
                stops: vec![GradientStop {
                    position: 0.5,
                    color: Color::new(255, 128, 0),
                }],
                angle: 0.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        // Should NOT contain gradient.linear (needs >= 2 stops)
        assert!(
            !output.source.contains("gradient.linear"),
            "Single-stop gradient should fall back to solid fill: {}",
            output.source,
        );
        // Should contain the solid fill color instead
        assert!(
            output.source.contains("rgb(255, 128, 0)"),
            "Single-stop gradient should use the stop color as solid fill: {}",
            output.source,
        );
    }

    #[test]
    fn test_gradient_two_stops_still_works() {
        let page = Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![],
            background_color: None,
            background_gradient: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(255, 0, 0),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 255),
                    },
                ],
                angle: 90.0,
            }),
        });
        let doc = make_doc(vec![page]);
        let output = generate_typst(&doc).unwrap();
        assert!(
            output.source.contains("gradient.linear"),
            "Two-stop gradient should still produce gradient.linear: {}",
            output.source,
        );
    }

    // --- US-382/383: unstyled run after styled run must not create `](` pattern ---

    #[test]
    fn test_unstyled_run_with_parens_after_styled_run() {
        // When a styled run is followed by an unstyled run starting with `(`,
        // the `](` pattern must not be interpreted as Typst function arguments.
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "bold text".to_string(),
                    style: TextStyle {
                        bold: Some(true),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                },
                Run {
                    text: "(parenthetical note)".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
            ],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        // The result must not contain `](` directly — it would be interpreted
        // as function arguments in Typst
        assert!(
            !result.contains("](\\(") || !result.contains("]("),
            "Unstyled text with parens after styled run must be wrapped safely. Got: {result}"
        );
        // Verify the output uses #[...] wrapper or other safe pattern
        assert!(
            result.contains("#[") || result.contains("\\("),
            "Unstyled text should be wrapped in #[...] to prevent syntax issues. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_superscript() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "2".to_string(),
                style: TextStyle {
                    vertical_align: Some(VerticalTextAlign::Superscript),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#super[2]"),
            "Superscript should use #super[...]. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_subscript() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "2".to_string(),
                style: TextStyle {
                    vertical_align: Some(VerticalTextAlign::Subscript),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#sub[2]"),
            "Subscript should use #sub[...]. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_small_caps() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle {
                    small_caps: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#smallcaps[Hello]"),
            "Small caps should use #smallcaps[...]. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_all_caps() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Hello World".to_string(),
                style: TextStyle {
                    all_caps: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("HELLO WORLD"),
            "All caps should uppercase the text. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_superscript_with_bold() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "n".to_string(),
                style: TextStyle {
                    vertical_align: Some(VerticalTextAlign::Superscript),
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#super[") && result.contains("weight: \"bold\""),
            "Superscript with bold should combine both. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_highlight_yellow() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Important".to_string(),
                style: TextStyle {
                    highlight: Some(Color::new(255, 255, 0)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#highlight(fill: rgb(255, 255, 0))[Important]"),
            "Highlight should use #highlight(fill: ...). Got: {result}"
        );
    }

    #[test]
    fn test_table_cell_vertical_align_center() {
        let table = Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Centered".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    vertical_align: Some(CellVerticalAlign::Center),
                    ..TableCell::default()
                }],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("align: horizon"),
            "Center vertical alignment should emit 'align: horizon'. Got: {result}"
        );
    }

    #[test]
    fn test_generate_run_highlight_with_bold() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold Highlight".to_string(),
                style: TextStyle {
                    highlight: Some(Color::new(0, 255, 0)),
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("#highlight(fill: rgb(0, 255, 0))["),
            "Should have highlight wrapper. Got: {result}"
        );
        assert!(
            result.contains("weight: \"bold\""),
            "Should have bold text. Got: {result}"
        );
    }

    #[test]
    fn test_table_cell_vertical_align_bottom() {
        let table = Table {
            rows: vec![TableRow {
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
                    vertical_align: Some(CellVerticalAlign::Bottom),
                    ..TableCell::default()
                }],
                height: None,
            }],
            column_widths: vec![100.0],
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            result.contains("align: bottom"),
            "Bottom vertical alignment should emit 'align: bottom'. Got: {result}"
        );
    }
}
