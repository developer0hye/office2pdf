use std::fmt::Write;
use std::io::Cursor;

use image::{GenericImageView, ImageFormat as RasterImageFormat};

use crate::config::ConvertOptions;
use crate::error::ConvertError;
use crate::ir::{
    Alignment, ArrowHead, AxisTickMark, BarBandLayout, BaselineShiftEm, Block, BorderLineStyle,
    BorderSide, CellBorder, CellVerticalAlign, Chart, ChartGrouping, ChartType, Color,
    ColumnLayout, Document, FixedElement, FixedElementKind, FixedPage, FloatingImage,
    FloatingShape, FloatingTextBox, FlowPage, FrameAnchor, GradientFill, HFInline, HeaderFooter,
    HeaderFooterFrame, IconShading, ImageCrop, ImageData, ImageFormat, ImageParagraphSpacing,
    Insets, LegendPosition, LineBox, LineJoin, LineSpacing, List, ListKind, Margins, MathEquation,
    Metadata, Page, PageNumberFormat, PageSize, PairKerning, Paragraph, ParagraphStyle,
    PatternFill, PatternPreset, PositionedTabAlignment, PositionedTabRelativeTo, Run, Shadow,
    Shape, ShapeKind, SheetPage, SmartArt, TabAlignment, TabLeader, TabStop, Table,
    TableBorderPaintModel, TableCell, TableOfContents, TableRow, TextBoxData, TextBoxVerticalAlign,
    TextDirection, TextStyle, VerticalTextAlign, WrapMode,
};

use self::diagrams::{
    generate_chart, generate_chart_in, generate_sheet_chart_in, generate_smartart,
};
use self::fmt::*;
use self::lists::{
    ListEojeolWrap, can_render_fixed_text_list_inline, common_text_style,
    fixed_text_paragraph_hanging_indent_pt, fixed_text_paragraph_inset, generate_fixed_text_list,
    generate_list, generate_list_with_spacing_model, write_common_text_settings,
    write_fixed_text_default_par_settings,
};
use self::shapes::{
    clamp_ring_corner_radius, generate_shape, shadow_alpha, shadow_outline_outset,
    shadow_silhouette_corner_radius, write_fill_color, write_gradient_fill, write_shape_stroke,
    write_text_box_shape_background,
};
use self::tables::generate_table;
use self::text::*;
use super::font_context::FontSearchContext;

#[path = "typst_gen_diagrams.rs"]
mod diagrams;
#[path = "typst_gen_fmt.rs"]
mod fmt;
#[path = "typst_gen_lists.rs"]
mod lists;
#[path = "typst_gen_shadow_outline.rs"]
mod shadow_outline;
#[path = "typst_gen_shapes.rs"]
mod shapes;
#[path = "typst_gen_tables.rs"]
mod tables;
#[path = "typst_gen_text.rs"]
mod text;

// The DOCX table min-content measurement routes East Asian codepoints to the
// run's `w:eastAsia` face the same way rendering does (issue #624).
pub(crate) use self::tables::header_row_count_covering_rowspans;
pub(crate) use self::text::is_cjk_like;
pub(crate) use self::text::{
    COMPACTED_SHEET_CELL_MIN_DESCENT_SEAT_PT, SHEET_CELL_MIN_DESCENT_SEAT_PT,
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

/// Maximum nesting depth for tables-within-tables, matching the parser limit.
const MAX_TABLE_DEPTH: usize = 64;
/// How far a justified line may squeeze its spaces, as a share of the space's
/// own width. Calibrated on the corpus in [`write_page_format_state`], which
/// is where the measurements behind it are recorded.
pub(super) const JUSTIFIED_SPACING_FLOOR: &str = "80%";

/// Typst's line box leaves more top leading than Word/LibreOffice text frames.
const FLOATING_TEXT_BOX_TOP_LEADING_COMPENSATION_PT: f64 = 6.0;

/// LibreOffice Writer's page-left WPS text origin sits 0.15pt inside the
/// DrawingML inset. The #1219 / PR #1407 footer isolates the seat on both
/// pages: its declared 20pt inset exports at x=20.15pt, while the run's width
/// and bottom baseline already agree (issue #1487).
const WRITER_PAGE_LEFT_TEXT_ORIGIN_SEAT_PT: f64 = 0.15;

/// The Typst state carrying the section's `w:pgNumType w:fmt`.
///
/// A `PAGE` field can take the format straight from `GenCtx`, because codegen
/// already knows which section wrote the field. A contents entry cannot: which
/// section an entry points into is only settled by the layout, so the format
/// has to travel with it (issue #605).
const PAGE_FORMAT_STATE: &str = "o2p-page-format";

/// The indent a contents entry takes per level below the first.
///
/// Word's built-in `TOC<n>` styles step the left indent by a flat amount; a
/// native export of the technical brief puts level 1 on the left margin at
/// 65.04pt and level 2 exactly 20pt in (issue #610).
const TOC_LEVEL_INDENT_PT: f64 = 20.0;

/// The indent a `TOC \a` caption-list entry takes. The same export puts the
/// figure and table lists' entries at 105.04pt against a 65.04pt margin — two
/// of the heading list's level steps (issue #611).
const CAPTION_LIST_INDENT_PT: f64 = 40.0;

/// The gap below a heading-list entry. Word prints a 24.96pt entry pitch on
/// the technical brief's contents page; the pitch is this gap plus the line
/// the entry's own font takes, so it is calibrated against that font.
const TOC_ENTRY_SPACING_PT: f64 = 17.8;

/// The same gap for a caption list. Word prints the same 24.96pt pitch there,
/// but its entries are Korean throughout where the contents page mixes
/// scripts, and the taller Korean line needs 1.4pt less gap to land on it
/// (issue #611).
const CAPTION_ENTRY_SPACING_PT: f64 = 16.4;

/// Internal context for tracking image assets during code generation.
struct GenCtx {
    /// Which flow section is being written, so a `<w:titlePg/>` header can
    /// label its own section's first page (issue #846).
    flow_section_index: usize,
    images: Vec<ImageAsset>,
    next_image_id: usize,
    next_text_box_id: usize,
    /// How many sheet drawing layers have been written. Each needs its own
    /// label, because the layer is a page foreground that has to recognise
    /// its own sheet's first printed page (issue #1168).
    next_sheet_drawing_layer_id: usize,
    table_depth: usize,
    /// Active section's Word document-grid line pitch, in points.
    line_grid_pitch: Option<f64>,
    /// `w:defaultTabStop` from the document settings, in points.
    document_default_tab_stop_pt: Option<f64>,
    /// Effective default tab stop interval, in points, for the active page.
    default_tab_width_pt: f64,
    /// True until the document's first page has been generated. Word keeps
    /// `w:spacing w:before` on the very first body paragraph, unlike the top of
    /// pages reached by a break.
    at_document_start: bool,
    /// The table row being generated's East Asian line answer. Decided once
    /// per row so every cell in it shares a baseline, which reading each
    /// cell's own text could not guarantee (issue #498).
    row_east_asian: RowEastAsianMetrics,
    /// The enclosing table's default vertical alignment: a cell that declares
    /// none takes this, and its paragraph codegen must know the effective
    /// answer to seat the line box (issue #618).
    table_default_vertical_align: Option<CellVerticalAlign>,
    /// Whether the enclosing table's box is positioned by an `#align(...)`
    /// wrapper. Typst inherits `align` into the cells, so a cell paragraph
    /// that declares none has to reset it or it inherits the table's own
    /// placement as text alignment (issue #843).
    table_box_is_aligned: bool,
    /// Whether the enclosing table rests bottom-aligned text on the descender
    /// line, i.e. is a spreadsheet ([`Table::seats_bottom_aligned_text_on_descender`]).
    table_seats_bottom_aligned_text_on_descender: bool,
    /// Whether that descender seat keeps Excel's minimum gap above the row's
    /// bottom boundary ([`Table::bottom_aligned_descent_floor_pt`], issues
    /// #1097 and #1199).
    table_bottom_aligned_descent_floor_pt: f64,
    /// The fit-to-page scale already folded into this table's sizes, from
    /// [`Table::print_scale`]. `None` on an unscaled sheet and off a sheet
    /// entirely; [`GenCtx::sheet_print_scale`] resolves the two apart.
    table_print_scale: Option<f64>,
    /// Visual translation applied to a fitted worksheet's table paint. Cell
    /// content takes the inverse so text that already matches Excel stays in
    /// its established seat while fills and boundary bands follow the scaled
    /// paper-space origin (issue #1538).
    sheet_paint_offset_pt: Option<(f64, f64)>,
    /// Whether the cell being generated seats its line box on the descender:
    /// the enclosing table is a spreadsheet and the cell's effective vertical
    /// alignment is bottom (issue #618).
    cell_seats_text_on_descender: bool,
    /// The current cell's effective vertical alignment. Word's compressed
    /// line box needs the resolved table default as well as an explicit cell
    /// value to reproduce the anchor-specific baseline seat (issue #1479).
    cell_vertical_align: Option<CellVerticalAlign>,
    /// The shared line the cell being generated seats on when its spreadsheet
    /// row is too tight for per-cell vertical alignment to differ: one metric
    /// family and size for the whole row, so every cell lands on one baseline
    /// as Excel prints it (issue #839). `None` outside that regime.
    cell_sheet_row_line: Option<SheetRowLine>,
    /// The fixed sheet track the cell being generated sits in, so its line
    /// seats on the baseline Excel prints rather than on the centre of the
    /// cell's own inset box (issue #1063). `None` outside that regime.
    cell_sheet_seat: Option<SheetCellSeat>,
    /// Whether emission is inside a spill cell's clipped wrapper (issue #811).
    in_spill_cell: bool,
    /// Numerals the active section's `PAGE` fields render in. A header is
    /// generated as part of its page's setup, so the section's `w:pgNumType
    /// w:fmt` reaches the field through the context rather than through the
    /// inline, which carries only the run properties (issue #582).
    page_number_format: PageNumberFormat,
    /// `w:docDefaults/w:rPrDefault` — the family and size a computed contents
    /// entry is laid out in, rather than the heading's own (issue #610).
    document_default_text: Option<crate::ir::TextStyle>,
    /// Whether the page being generated is a Word flow page, whose ordinary
    /// lines break Hangul only at eojeol boundaries (issue #626). Slides and
    /// sheets keep the engine's syllable breaking, which is what PowerPoint
    /// and Excel do.
    ///
    /// The flag reaches the body, the headers and footers, the tables and the
    /// lists of a flow page. It deliberately does NOT reach a *floating* text
    /// box: those go through [`generate_fixed_text_paragraph`], which pins
    /// `EojeolWrap::Syllable` because that path resolves neither the frame's
    /// fixed text edges nor the box's inner measure — see the note there.
    ///
    /// The flag is the flow page's *default*, not the last word. A paragraph
    /// carrying `w:wordWrap w:val="0"` — how a document asks Word for
    /// character-level Hangul breaking — overrides it, and
    /// [`paragraph_eojeol_wrap`] checks that before anything here (issue
    /// #730).
    breaks_hangul_at_eojeol: bool,
    /// The width one line of the current container has, in points, before a
    /// paragraph's own indents are taken off it: a flow page's text width,
    /// narrowed to the column width inside a table cell. `None` when no
    /// measure is known. An eojeol wider than this must not be framed — the
    /// frame would be pushed onto a line of its own and overflow it
    /// (issue #626).
    available_measure_pt: Option<f64>,
    /// Whether the table being generated is a slide's, and so paces its cell
    /// text on PowerPoint's flat 1.2em line rather than Word's hhea one.
    ///
    /// A slide's own text boxes already route through
    /// [`powerpoint_line_height_settings`], but a `<a:tbl>` reaches the shared
    /// table codegen, which gave its cells Word's model: 1.587em measured on
    /// `office2pdf_introduction_ko` slide 16, so multi-line cells grew and the
    /// table's bottom border moved down with them (issue #663).
    ///
    /// The target is PowerPoint's documented 1.2em, the same factor the slide's
    /// text boxes already use. No native PowerPoint export of that fixture is
    /// committed; a **LibreOffice** render of it advances 1.235em, which
    /// corroborates the direction and magnitude without being ground truth.
    table_uses_powerpoint_line_box: bool,
}

impl GenCtx {
    /// The scale a spreadsheet cell's declared-space metrics must be read
    /// through: `Some(1.0)` on an unscaled sheet, `Some(the fit-to-page
    /// factor)` on a fitted one, and `None` off a sheet. This covers the
    /// measured wrapped-line advance (#1163) and whole-point seats (#1238).
    fn sheet_print_scale(&self) -> Option<f64> {
        self.table_seats_bottom_aligned_text_on_descender
            .then(|| self.table_print_scale.unwrap_or(1.0))
    }

    fn new() -> Self {
        Self {
            flow_section_index: 0,
            images: Vec::new(),
            next_image_id: 0,
            next_text_box_id: 0,
            next_sheet_drawing_layer_id: 0,
            table_depth: 0,
            table_uses_powerpoint_line_box: false,
            line_grid_pitch: None,
            row_east_asian: RowEastAsianMetrics {
                has_east_asian_text: false,
                takes_east_asian_metrics: false,
            },
            table_default_vertical_align: None,
            table_box_is_aligned: false,
            table_seats_bottom_aligned_text_on_descender: false,
            table_bottom_aligned_descent_floor_pt: 0.0,
            table_print_scale: None,
            sheet_paint_offset_pt: None,
            cell_seats_text_on_descender: false,
            cell_vertical_align: None,
            cell_sheet_row_line: None,
            cell_sheet_seat: None,
            in_spill_cell: false,
            page_number_format: PageNumberFormat::default(),
            document_default_text: None,
            document_default_tab_stop_pt: None,
            default_tab_width_pt: DEFAULT_TAB_WIDTH_PT,
            at_document_start: true,
            breaks_hangul_at_eojeol: false,
            available_measure_pt: None,
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

    fn add_generated_svg(&mut self, data: Vec<u8>) -> String {
        let id = self.next_image_id;
        self.next_image_id += 1;
        let path = format!("img-{id}.svg");
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

/// Narrow an SVG's `viewBox` to the region an `a:srcRect` keeps.
///
/// A root `<svg>` clips to its viewport, so moving the viewBox is the whole
/// crop — the drawing itself is untouched, and nothing has to be rasterised
/// (issue #892). Returns `None` when the root states no usable `viewBox`,
/// leaving the asset alone rather than guessing at its user units.
fn crop_svg_view_box(data: &[u8], crop: ImageCrop) -> Option<Vec<u8>> {
    let text: &str = std::str::from_utf8(data).ok()?;
    let open_end: usize = text.find('>')?;
    let head: &str = &text[..open_end];
    if !head.trim_start().starts_with("<svg") {
        return None;
    }
    let attr_start: usize = head.find("viewBox=\"")? + "viewBox=\"".len();
    let attr_len: usize = head[attr_start..].find('"')?;
    let values: Vec<f64> = head[attr_start..attr_start + attr_len]
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|part| !part.is_empty())
        .map(str::parse::<f64>)
        .collect::<Result<Vec<f64>, _>>()
        .ok()?;
    let [x, y, width, height] = values[..] else {
        return None;
    };
    if !(width.is_finite() && height.is_finite() && width > 0.0 && height > 0.0) {
        return None;
    }

    let left: f64 = crop.left.clamp(0.0, 1.0);
    let top: f64 = crop.top.clamp(0.0, 1.0);
    let kept_w: f64 = width * (1.0 - left - crop.right.clamp(0.0, 1.0));
    let kept_h: f64 = height * (1.0 - top - crop.bottom.clamp(0.0, 1.0));
    if kept_w <= 0.0 || kept_h <= 0.0 {
        return None;
    }

    let replacement: String = format!(
        "viewBox=\"{} {} {} {}\"",
        format_f64(x + width * left),
        format_f64(y + height * top),
        format_f64(kept_w),
        format_f64(kept_h)
    );
    let mut out: String = String::with_capacity(text.len() + replacement.len());
    out.push_str(&text[..attr_start - "viewBox=\"".len()]);
    out.push_str(&replacement);
    out.push_str(&text[attr_start + attr_len + 1..]);

    // The viewport has to shrink with the viewBox. Left at its old size it
    // still describes the whole drawing, and `preserveAspectRatio` then meets
    // the smaller viewBox inside it — scaling by 1 and centring the content
    // rather than cropping it, which is a translation and no crop at all.
    let out: String = replace_svg_length(&out, "width", kept_w);
    let out: String = replace_svg_length(&out, "height", kept_h);
    Some(out.into_bytes())
}

/// Rewrite a root `<svg>` length attribute, keeping any unit suffix it
/// carries. Leaves the document alone when the attribute is absent, since a
/// root without one already takes its size from the viewBox.
fn replace_svg_length(text: &str, name: &str, value: f64) -> String {
    let open_end: usize = match text.find('>') {
        Some(index) => index,
        None => return text.to_string(),
    };
    let needle: String = format!("{name}=\"");
    let Some(offset) = text[..open_end].find(&needle) else {
        return text.to_string();
    };
    let start: usize = offset + needle.len();
    let Some(len) = text[start..].find('"') else {
        return text.to_string();
    };
    let old: &str = &text[start..start + len];
    let unit: &str = old.trim_start_matches(|c: char| c.is_ascii_digit() || c == '.' || c == '-');
    format!(
        "{}{}{}{}",
        &text[..start],
        format_f64(value),
        unit,
        &text[start + len..]
    )
}

fn preprocess_image_asset(image: &ImageData) -> (Vec<u8>, ImageFormat) {
    let Some(crop) = image.crop.filter(|crop| !crop.is_empty()) else {
        return (image.data.clone(), image.format);
    };
    if image.format == ImageFormat::Svg {
        return match crop_svg_view_box(&image.data, crop) {
            Some(cropped) => (cropped, ImageFormat::Svg),
            None => (image.data.clone(), image.format),
        };
    }
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
pub(crate) fn resolve_page_size(original: &PageSize, options: &ConvertOptions) -> PageSize {
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
pub(crate) fn escape_typst_string(s: &str) -> String {
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
// Only the wasm pipeline branch calls this at runtime; native builds use the
// font-context variant and reach this wrapper solely from tests.
#[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
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
    // The classes the source declares travel with it, so a face that states
    // itself sans is not guessed at by name (issue #891).
    let previous_classes =
        super::font_subst::set_declared_font_classes(doc.styles.declared_font_classes.clone());
    // `compatibilityMode` is declared once for the package and decides whether
    // a justified East Asian line may compress to seat one more token, so it
    // travels with the whole generation rather than with any one paragraph
    // (issue #1130).
    let legacy_justification: bool = matches!(
        doc.styles.word_compatibility_mode,
        Some(crate::ir::WordCompatibilityMode::Legacy)
    );
    let generated = super::font_subst::with_font_search_context(font_context, || {
        text::with_legacy_word_justification(legacy_justification, || {
            let first_pass: TypstOutput =
                text::with_rtl_shaping_exemption(false, || generate_pages(doc, options))?;
            // Whether the document shapes right-to-left is answered by the
            // markup itself, so nothing about the IR has to be modelled to find
            // it — see `with_rtl_shaping_exemption`. The pass costs one more
            // walk of the IR for the documents that do, and nothing for the
            // ones that do not.
            if !text::source_shapes_right_to_left(&first_pass.source) {
                return Ok(first_pass);
            }
            text::with_rtl_shaping_exemption(true, || generate_pages(doc, options))
        })
    });
    super::font_subst::set_declared_font_classes(previous_classes);
    generated
}

fn generate_pages(doc: &Document, options: &ConvertOptions) -> Result<TypstOutput, ConvertError> {
    // Pre-allocate output string: ~2KB per page is a reasonable estimate
    let mut out = String::with_capacity(doc.pages.len() * 2048);

    // Emit document metadata (title/author) if present
    generate_document_metadata(&mut out, &doc.metadata);
    write_page_format_state(&mut out);
    if doc.pages.iter().any(|page| matches!(page, Page::Fixed(_))) {
        write_powerpoint_ligature_state(&mut out);
        write_powerpoint_advance_grid_helpers(&mut out);
        out.push('\n');
    }

    let mut ctx = GenCtx::new();
    ctx.document_default_tab_stop_pt = doc.styles.default_tab_stop_pt;
    ctx.document_default_text = doc.styles.default_text.clone();
    for (index, page) in doc.pages.iter().enumerate() {
        if index > 0 {
            out.push_str("\n#pagebreak()\n");
        }
        match page {
            Page::Flow(flow) => {
                generate_flow_page(&mut out, flow, &mut ctx, options)?;
                ctx.flow_section_index += 1;
            }
            Page::Fixed(fixed) => generate_fixed_page(&mut out, fixed, &mut ctx, options)?,
            Page::Sheet(sheet_page) => {
                generate_table_page(&mut out, sheet_page, &mut ctx, options)?;
            }
        }
        ctx.at_document_start = false;
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
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    // The format has to be in place before the header is written, because the
    // header is part of the page setup and carries the `PAGE` field.
    if let Some(numbering) = page.page_numbering {
        ctx.page_number_format = numbering.format;
    }
    ctx.breaks_hangul_at_eojeol = true;
    // Word's text column: what `#set page(margin:)` below leaves between the
    // left and right margins.
    ctx.available_measure_pt =
        Some(size.width - page.margins.left - page.margins.right).filter(|measure| *measure > 0.0);
    write_flow_page_setup(out, page, &size, ctx);
    out.push('\n');
    // The marker sits at the section's first page, so a first-page header can
    // resolve which page that is without assuming the section starts the
    // document (issue #846).
    if page.first_header.is_some() || page.first_footer.is_some() {
        let _ = writeln!(
            out,
            "#metadata(none) <{}>",
            section_first_page_label(ctx.flow_section_index)
        );
    }
    // Word restarts the counter at the section boundary; Typst counts pages
    // from the document start, so the section states its own first number.
    if let Some(start) = page.page_numbering.and_then(|numbering| numbering.start) {
        let _ = writeln!(out, "#counter(page).update({start})");
    }
    // A `PAGE` field takes the format from `ctx` at codegen time, but a
    // contents entry cannot: which section an entry lands in is only known
    // once the document is laid out. Record the format in a Typst state so the
    // outline can read back whatever was in force where each entry resolved
    // (issue #605).
    if let Some(numbering) = page.page_numbering {
        let _ = writeln!(
            out,
            "#{PAGE_FORMAT_STATE}.update(\"{}\")",
            numbering.format.typst_pattern()
        );
    }
    // Only a snapping grid reaches the line model: a `w:docGrid` whose type is
    // `default` declares a pitch Word ignores for layout (issue #518). The bare
    // presence of the element still marks an East Asian edition for the tab
    // default below, which is a different question.
    ctx.line_grid_pitch = page.line_grid_pitch.filter(|_| page.line_grid_snaps_lines);
    // Absent w:defaultTabStop: East Asian Word editions (signalled by the
    // section's w:docGrid) default to 800 twips = 40pt where Western
    // editions use the ECMA 720 twips = 36pt (issue #393).
    ctx.default_tab_width_pt =
        ctx.document_default_tab_stop_pt
            .unwrap_or(if page.line_grid_pitch.is_some() {
                EAST_ASIAN_DEFAULT_TAB_WIDTH_PT
            } else {
                DEFAULT_TAB_WIDTH_PT
            });

    // Word keeps `w:spacing w:before` on the document's first body paragraph,
    // but Typst collapses leading block spacing at a page boundary, pulling the
    // first heading up to the top margin. Emit that gap as explicit vertical
    // space instead, and drop the block's own `above` so it is not counted
    // twice. Later page tops keep collapsing spacing, which is what Word does
    // after a page break.
    let leading_gap: Option<f64> = if ctx.at_document_start && page.columns.is_none() {
        match page.content.first() {
            Some(Block::Paragraph(paragraph)) => {
                paragraph.style.space_before.filter(|gap| *gap > 0.0)
            }
            _ => None,
        }
    } else {
        None
    };

    if let Some(gap) = leading_gap {
        let _ = writeln!(out, "#v({}pt, weak: false)", format_f64(gap));
        let Some(Block::Paragraph(first)) = page.content.first() else {
            unreachable!("leading_gap is only set for a leading paragraph")
        };
        let mut adjusted = first.clone();
        adjusted.style.space_before = None;
        generate_block(out, &Block::Paragraph(adjusted), ctx)?;
        if page.content.len() > 1 {
            out.push('\n');
            generate_blocks(out, &page.content[1..], ctx)?;
        }
    } else if let Some(ref cols) = page.columns {
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
    ctx.breaks_hangul_at_eojeol = false;
    ctx.available_measure_pt = None;
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
            "#set page(width: {}pt, height: {}pt, margin: 0pt, fill: {})",
            format_f64(size.width),
            format_f64(size.height),
            rgb(bg),
        );
    } else {
        let _ = writeln!(
            out,
            "#set page(width: {}pt, height: {}pt, margin: 0pt, fill: white)",
            format_f64(size.width),
            format_f64(size.height),
        );
    }
    out.push('\n');

    let occluding_images: Vec<(usize, &FixedElement)> = page
        .elements
        .iter()
        .enumerate()
        .filter(|(_, element)| is_full_page_opaque_image(element, size))
        .collect();

    with_powerpoint_advance_grid(true, || -> Result<(), ConvertError> {
        for (index, elem) in page.elements.iter().enumerate() {
            if is_frame_bounded_text_box(elem)
                && occluding_images.iter().any(|(cover_index, cover)| {
                    *cover_index > index && fixed_element_covers(cover, elem)
                })
            {
                continue;
            }
            generate_fixed_element(out, elem, size.height, ctx)?;
        }
        Ok(())
    })
}

/// Whether a raster can safely stand in for the full slide when deciding that
/// an earlier, frame-bounded text box contributes neither pixels nor searchable
/// PDF content.
///
/// PowerPoint drops master text covered by later slide artwork from its PDF
/// text layer. Keep this deliberately narrow: non-page-sized pictures and
/// transformed/clipped images may leave part of the text visible (issue #1432).
fn is_full_page_opaque_image(element: &FixedElement, page_size: PageSize) -> bool {
    const EDGE_TOLERANCE_PT: f64 = 0.5;

    let FixedElementKind::Image(image) = &element.kind else {
        return false;
    };
    if image.rotation_deg.is_some_and(|degrees| degrees != 0.0) || image.clip_shape.is_some() {
        return false;
    }
    if element.x > EDGE_TOLERANCE_PT
        || element.y > EDGE_TOLERANCE_PT
        || element.x + element.width < page_size.width - EDGE_TOLERANCE_PT
        || element.y + element.height < page_size.height - EDGE_TOLERANCE_PT
    {
        return false;
    }

    match image.format {
        ImageFormat::Jpeg => true,
        ImageFormat::Png | ImageFormat::Gif | ImageFormat::Bmp | ImageFormat::Tiff => {
            image::load_from_memory(&image.data)
                .ok()
                .is_some_and(|decoded| {
                    !decoded.color().has_alpha()
                        || decoded.to_rgba8().pixels().all(|pixel| pixel[3] == u8::MAX)
                })
        }
        ImageFormat::Svg => false,
    }
}

/// Restrict occlusion culling to upright, wrapping text boxes. A no-wrap box
/// deliberately paints beyond its frame, while either text or shape rotation
/// can move glyphs outside the axis-aligned bounds tested below.
fn is_frame_bounded_text_box(element: &FixedElement) -> bool {
    let FixedElementKind::TextBox(text_box) = &element.kind else {
        return false;
    };
    !text_box.no_wrap
        && !text_box
            .text_rotation_deg
            .is_some_and(|degrees| degrees != 0.0)
        && !text_box
            .shape_rotation_deg
            .is_some_and(|degrees| degrees != 0.0)
}

fn fixed_element_covers(cover: &FixedElement, covered: &FixedElement) -> bool {
    const GEOMETRY_TOLERANCE_PT: f64 = 0.01;

    cover.x <= covered.x + GEOMETRY_TOLERANCE_PT
        && cover.y <= covered.y + GEOMETRY_TOLERANCE_PT
        && cover.x + cover.width + GEOMETRY_TOLERANCE_PT >= covered.x + covered.width
        && cover.y + cover.height + GEOMETRY_TOLERANCE_PT >= covered.y + covered.height
}

fn generate_table_page(
    out: &mut String,
    page: &SheetPage,
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    ctx.breaks_hangul_at_eojeol = false;
    ctx.available_measure_pt = None;

    // `<printOptions horizontalCentered="1"/>` moves the whole printed sheet,
    // the drawings floating over the grid included, so the inset reaches both.
    // The grid takes it from the `#pad` below; the drawing layer sits outside
    // that flow and carries it in its own offsets instead.
    // The page margins themselves stay put: the header and footer keep their
    // own alignment, which the centering does not touch.
    let centering_inset_pt: Option<f64> = horizontal_centering_inset_pt(page, &size);

    // Excel paints a fitted sheet's grid against a paper box constructed in
    // declared sheet space and then scaled. Cell text has its own already
    // calibrated seats, so only the table paint takes this translation; the
    // cell writer counter-shifts its content by the inverse (#1538).
    let paint_offset_pt: Option<(f64, f64)> =
        scaled_sheet_paint_offset_pt(page, &size, centering_inset_pt);

    // Every glyph Excel prints on a sheet — the grid's cells and the text of
    // the drawings floating over it alike — advances on a whole-point grid in
    // sheet space (issues #1088, #1238). The drawing layer goes into the page
    // setup, so its scope has to open before that; the page setup itself states
    // no run text.
    // A drawing keeps its declared sizes and is scaled as a whole by its
    // placement, so its advance grid is already in sheet space. Cell runs,
    // below, have had the print scale folded into their sizes by the parser
    // and need that factor to recover the same coordinate system (#1238).
    let drawings: Option<SheetDrawingLayer> = with_sheet_advance_grid(Some(1.0), || {
        sheet_drawing_layer(page, centering_inset_pt, ctx)
    });

    write_table_page_setup(
        out,
        page,
        &size,
        ctx,
        drawings.as_ref().map(|layer| layer.foreground.as_str()),
    );
    out.push('\n');

    if let Some((dx_pt, dy_pt)) = paint_offset_pt {
        let _ = writeln!(
            out,
            "#move(dx: {}pt, dy: {}pt)[",
            format_f64(dx_pt),
            format_f64(dy_pt),
        );
    }

    ctx.sheet_paint_offset_pt = paint_offset_pt;

    if let Some(inset_pt) = centering_inset_pt {
        let _ = writeln!(out, "#pad(left: {}pt)[", format_f64(inset_pt));
    }

    if drawings.is_none() && page.charts.is_empty() {
        with_sheet_advance_grid(Some(page.table.print_scale.unwrap_or(1.0)), || {
            generate_table(out, &page.table, ctx)
        })?;
    } else {
        generate_sheet_grid(
            out,
            &page.table,
            &page.charts,
            drawings.as_ref().map(|layer| layer.marker.as_str()),
            ctx,
        )?;
    }

    if centering_inset_pt.is_some() {
        out.push_str("\n]\n");
    }
    if paint_offset_pt.is_some() {
        out.push_str("]\n");
    }
    ctx.sheet_paint_offset_pt = None;
    Ok(())
}

/// Translation from the converter's physical table origin to Excel's fitted
/// sheet-space paint origin.
///
/// On issue #1538's A3 page, the converter centres the 1,078.30pt grid at
/// `50 + 5.125 = 55.125pt`. Excel first snaps the paper box in sheet space:
/// its left edge is `floor(50 / .82) * .82 = 49.20pt`, its right edge is
/// `ceil((1190.55 - 50) / .82) * .82 = 1140.62pt`, and its scaled 1pt
/// centring bias leaves a 5.74pt inset. The fitted paint origin is therefore
/// 54.94pt, or -0.185pt from the converter's. The top origin similarly moves
/// from 54pt to `floor(54 / .82) * .82 = 53.30pt`.
fn scaled_sheet_paint_offset_pt(
    page: &SheetPage,
    size: &PageSize,
    physical_centering_inset_pt: Option<f64>,
) -> Option<(f64, f64)> {
    let scale: f64 = page
        .table
        .print_scale
        .filter(|scale| *scale > 0.0 && *scale < 1.0)?;
    let thousandth = |points: f64| -> f64 { (points * 1000.0).round() / 1000.0 };
    let left_pt: f64 = (page.margins.left / scale).floor() * scale;
    let right_pt: f64 = ((size.width - page.margins.right) / scale).ceil() * scale;
    let top_pt: f64 = (page.margins.top / scale).floor() * scale;
    let grid_width_pt: f64 = page.table.column_widths.iter().sum();
    let scaled_centering_inset_pt: f64 = if page.table.centers_between_print_margins {
        ((right_pt - left_pt - grid_width_pt) / 2.0 - HORIZONTAL_CENTERING_BIAS_PT * scale).max(0.0)
    } else {
        0.0
    };
    let physical_left_pt: f64 = page.margins.left + physical_centering_inset_pt.unwrap_or(0.0);
    Some((
        thousandth(left_pt + scaled_centering_inset_pt - physical_left_pt),
        thousandth(top_pt - page.margins.top),
    ))
}

/// Points Excel prints a horizontally centred sheet left of the exact centre
/// of its printable width.
///
/// Measured on fourteen native Excel-for-Mac exports of one-factor probe
/// workbooks (issue #1110). Each fills its whole print range solid, so the
/// export's fill rectangle names the printed grid box directly; `left` below
/// is that rectangle's left edge and `centre` the exact
/// `margin + (printable − grid) / 2`. The ten centred rows below come from
/// nine of the workbooks, one of which splits into two page-columns:
///
/// | probe | page | margins | grid | centre | left |
/// | --- | --- | --- | ---: | ---: | ---: |
/// | 5 × 60pt | A4 portrait | 0.7in | 300 | 147.5 | **146** |
/// | 5 × 60pt | A4 portrait | 1.0in | 300 | 147.5 | **146** |
/// | 5 × 60pt | A4 landscape | 0.7in | 300 | 271.0 | **270** |
/// | 5 × 60pt | A4 portrait | 0.5in/1.5in | 300 | 111.5 | **110** |
/// | 3 × 120pt | A4 portrait | 0.7in | 360 | 117.5 | **116** |
/// | 5 × 63pt | A4 portrait | 0.7in | 315 | 140.0 | **139** |
/// | 5 × 60pt + empty 60pt print-area column | A4 portrait | 0.7in | 360 | 117.5 | **116** |
/// | 20 × 180pt at the 0.20 fit scale | A4 landscape | 0.7in | 720 | 61.0 | **60** |
/// | 10 × 120pt, pages 1-2 (4 columns each) | A4 portrait | 0.7in | 480 | 57.5 | **56** |
/// | 10 × 120pt, page 3 (2 columns) | A4 portrait | 0.7in | 240 | 177.5 | **176** |
///
/// Every one lands on `floor(centre) − 1`. Three of the remaining five
/// exports are uncentred controls — the first, second and last workbooks above
/// re-exported with `horizontalCentered="0"` — and print flush at the left
/// margin instead, so the offset is the centering's alone; one carries
/// `verticalCentered` only and moves the grid down, not sideways; and the last
/// adds a trailing `<col width="10"/>` outside both the used range and any
/// print area, and prints identically to the first, so the width Excel centres
/// is the printed range and not every declared column.
///
/// Three further readings follow from the table: the asymmetric probe centres
/// between the *margins*, not on the page; the fit-to-page probe centres the
/// width the scale leaves; and a sheet split into column groups centres each
/// page by the columns on it. Excel's own whole-point page and margin geometry
/// is what makes the result an integer. The margins now reach the page on that
/// whole point (issue #1127), but this converter's page is still 595.28pt wide
/// where Excel's A4 is 595, so the flooring is not reproduced end to end and
/// only the measured 1pt is taken off; the residual stays under 0.7pt on every
/// probe above. A centred grid's left edge is `(page + left − right − grid)/2`,
/// so snapping both side margins by the same fraction leaves every reading in
/// the table exactly where it was.
const HORIZONTAL_CENTERING_BIAS_PT: f64 = 1.0;

/// Points to inset the printed sheet from its left margin, or `None` when the
/// sheet is not centred or its grid already fills the printable width.
fn horizontal_centering_inset_pt(page: &SheetPage, size: &PageSize) -> Option<f64> {
    if !page.table.centers_between_print_margins {
        return None;
    }
    let printable_width_pt: f64 = size.width - page.margins.left - page.margins.right;
    let grid_width_pt: f64 = page.table.column_widths.iter().sum();
    let inset_pt: f64 = (printable_width_pt - grid_width_pt) / 2.0 - HORIZONTAL_CENTERING_BIAS_PT;
    (inset_pt > 0.0).then_some(inset_pt)
}

/// A drawing overlaid on a sheet at its anchor's absolute coordinates.
enum SheetAnchor<'a> {
    Chart(&'a crate::ir::SheetChart),
    Image(&'a crate::ir::SheetImage),
    TextBox(&'a crate::ir::SheetTextBox),
}

/// Render a sheet's grid, under the marker its drawing layer is pinned to and
/// above the charts no drawing anchors.
fn generate_sheet_grid(
    out: &mut String,
    table: &Table,
    charts: &[crate::ir::SheetChart],
    drawing_marker: Option<&str>,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // The marker is what pins the drawing layer to this sheet's first printed
    // page; see [`sheet_drawing_layer`]. It leads the grid because the sheet's
    // content origin is where the drawings' offsets are measured from.
    if let Some(marker) = drawing_marker {
        out.push_str(marker);
    }

    // Fitted cell runs already carry printed font sizes, so recover the
    // declared sheet-space advance grid only while rendering the grid. A flow
    // chart below it keeps declared chart sizes and must not inherit the cell
    // scale (issue #1238), just like the floating drawing layer above.
    with_sheet_advance_grid(Some(table.print_scale.unwrap_or(1.0)), || {
        generate_table(out, table, ctx)
    })?;
    out.push('\n');

    // A chart no drawing anchors has no worksheet coordinates to overlay it
    // at, so it stays flow content and follows the grid.
    for sheet_chart in charts.iter().filter(|chart| chart.placement.is_none()) {
        generate_chart(out, &sheet_chart.chart);
        out.push('\n');
    }

    Ok(())
}

/// A sheet's floating drawings, as the two pieces of markup that place them.
struct SheetDrawingLayer {
    /// The `#set page(foreground: …)` value that paints them.
    foreground: String,
    /// The zero-height, labelled block the foreground recognises the sheet's
    /// first printed page by. Emitted at the top of the sheet's content.
    marker: String,
}

/// Build the layer that floats every worksheet drawing over the grid.
///
/// Excel overlays drawings on the grid at absolute worksheet coordinates.
/// Threading them through the row flow could never match that, because our
/// printed row heights are not Excel's: on the regression fixture the two rows
/// above the anchor occupy 47.3pt here against Excel's 36pt, leaving the shapes
/// 11pt low however carefully they were sequenced (issue #474). They are placed
/// from the sheet's content origin instead, reserving no flow height (issues
/// #459, #982, #1101).
///
/// The layer is the *page foreground* rather than flow content, because Excel
/// floats a drawing above the cells and Typst paints in document order. A
/// zero-height block of `#place`d drawings ahead of the grid reserves no space,
/// but it also keeps no z-order claim over what follows: every cell fill landed
/// on top of it, and a picture anchored inside a filled panel was painted and
/// then covered (issue #1168). Moving that block after the grid would paint in
/// the right order but anchor to the wrong place — a sheet taller than one page
/// breaks across regions, and a `#place` after the grid resolves against the
/// last of them, so the drawings would leave the page they belong on. The page
/// foreground paints above the whole body and is anchored to a page rather than
/// to a flow position, which is both properties at once.
///
/// A foreground applies to every page the sheet spans, so the markup queries
/// [`SheetDrawingLayer::marker`] and draws only on the page that carries it.
/// Its offsets are page-relative where the flow's were content-relative, so the
/// margins and any centering inset are folded into them here.
fn sheet_drawing_layer(
    page: &SheetPage,
    centering_inset_pt: Option<f64>,
    ctx: &mut GenCtx,
) -> Option<SheetDrawingLayer> {
    let placed_charts: Vec<&crate::ir::SheetChart> = page
        .charts
        .iter()
        .filter(|chart| chart.placement.is_some())
        .collect();
    if placed_charts.is_empty() && page.images.is_empty() && page.text_boxes.is_empty() {
        return None;
    }

    let label: String = format!("o2p-sheet-drawings-{}", ctx.next_sheet_drawing_layer_id);
    ctx.next_sheet_drawing_layer_id += 1;
    // Block-level, not a `box`: an inline box still makes its paragraph lay out
    // a line, which dropped the whole grid by 13.2pt — Typst's default 11pt
    // text at 1.2 leading, independent of the sheet's own font (issue #1101).
    let marker: String =
        format!("#block(width: 100%, height: 0pt, spacing: 0pt)[#metadata(none)<{label}>]");

    let left_pt: f64 = page.margins.left + centering_inset_pt.unwrap_or(0.0);
    let top_pt: f64 = page.margins.top;

    let mut foreground = String::new();
    // Typst has no z-index, so the page a foreground belongs to is decided by
    // introspection: the marker's own page against the one being painted.
    let _ = write!(
        foreground,
        "context {{ let marker = query(<{label}>); \
         if marker.len() != 0 and marker.first().location().page() == here().page() ["
    );
    for sheet_chart in placed_charts {
        let placement = sheet_chart
            .placement
            .expect("only placed charts reach the drawing layer");
        write_placed_sheet_drawing(
            &mut foreground,
            &SheetAnchor::Chart(sheet_chart),
            left_pt,
            top_pt + placement.y_offset_pt,
            ctx,
        );
    }
    for sheet_image in &page.images {
        write_placed_sheet_drawing(
            &mut foreground,
            &SheetAnchor::Image(sheet_image),
            left_pt,
            top_pt + sheet_image.y_offset_pt,
            ctx,
        );
    }
    for text_box in &page.text_boxes {
        write_placed_sheet_drawing(
            &mut foreground,
            &SheetAnchor::TextBox(text_box),
            left_pt,
            top_pt + text_box.y_offset_pt,
            ctx,
        );
    }
    foreground.push_str("] }");

    Some(SheetDrawingLayer { foreground, marker })
}

/// Place one drawing at `dy` from the page top, inside the drawing layer's
/// markup block. `left_pt` is what the drawing's own horizontal offset is
/// measured from.
fn write_placed_sheet_drawing(
    out: &mut String,
    anchor: &SheetAnchor,
    left_pt: f64,
    dy_pt: f64,
    ctx: &mut GenCtx,
) {
    let _ = write!(out, "#place(top + left, dy: {}pt)[", format_f64(dy_pt));
    write_placed_sheet_anchor(out, anchor, left_pt, dy_pt, ctx);
    out.push(']');
}

/// Place one drawing at its horizontal offset, inside the vertical `#place`
/// the caller opened.
fn write_placed_sheet_anchor(
    out: &mut String,
    anchor: &SheetAnchor,
    left_pt: f64,
    dy_pt: f64,
    ctx: &mut GenCtx,
) {
    match anchor {
        SheetAnchor::Chart(sheet_chart) => {
            let Some(placement) = sheet_chart.placement else {
                return;
            };
            // The anchor sizes the chart, the way a slide's graphicFrame
            // extent does (issue #548); rendering at the intrinsic size
            // instead left the anchored band empty beneath it (issue #982).
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt)[",
                format_f64(left_pt + placement.x_offset_pt),
            );
            // A fitted sheet prints its drawings shrunk whole, so the chart is
            // laid out at its full frame and the transform brings its text down
            // with its geometry (issue #1069). The corner it grows from is the
            // anchor's, which is where the enclosing `place` put it.
            let fitted: bool = placement.print_scale != 1.0;
            if fitted {
                // Excel's fit scale is a whole percent, so rounding here keeps
                // 0.82 from printing as 82.00000000000001%.
                let percent: String = format_f64((placement.print_scale * 1e6).round() / 1e4);
                let _ = write!(
                    out,
                    "#scale(x: {percent}%, y: {percent}%, origin: top + left)[",
                );
            }
            let sheet_frame_top_pt: f64 = if placement.print_scale > 0.0 {
                dy_pt / placement.print_scale
            } else {
                dy_pt
            };
            generate_sheet_chart_in(
                out,
                &sheet_chart.chart,
                (placement.width, placement.height),
                sheet_frame_top_pt,
            );
            if fitted {
                out.push(']');
            }
            out.push(']');
        }
        SheetAnchor::TextBox(text_box) => {
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt)[#box(width: {}pt, height: {}pt",
                format_f64(left_pt + text_box.x_offset_pt),
                format_f64(text_box.width),
                format_f64(text_box.height),
            );
            if let Some(fill) = text_box.fill {
                let _ = write!(out, ", fill: {}", rgb(&fill));
            }
            if let Some(ref border) = text_box.border {
                let _ = write!(
                    out,
                    ", stroke: {}pt + {}",
                    format_f64(border.width),
                    rgb(&border.color),
                );
            }
            out.push_str(", inset: 4pt)[");
            if text_box.vertical_center {
                out.push_str("#align(horizon)[");
            }
            for paragraph in &text_box.paragraphs {
                let _ = generate_block(out, &Block::Paragraph(paragraph.clone()), ctx);
            }
            if text_box.vertical_center {
                out.push(']');
            }
            out.push_str("]]");
        }
        SheetAnchor::Image(sheet_image) => {
            // A page-column window from drawing-width pagination: the image
            // clips at the printable edge and a continued copy starts at a
            // negative offset, as Excel prints it (issue #713). The clip box
            // needs the image's own size — an inner `place` occupies nothing,
            // so an auto-sized box would collapse and clip everything away.
            let clip = sheet_image
                .clip_width_pt
                .zip(sheet_image.image.height)
                .filter(|_| sheet_image.image.width.is_some());
            if let Some((clip_width, image_height)) = clip {
                let _ = write!(
                    out,
                    "#place(top + left, dx: {}pt)[#box(width: {}pt, height: {}pt, clip: true)[#place(top + left, dx: {}pt)[",
                    format_f64(left_pt),
                    format_f64(clip_width),
                    format_f64(image_height),
                    format_f64(sheet_image.x_offset_pt),
                );
                generate_image(out, &sheet_image.image, ctx);
                out.push_str("]]]");
            } else {
                let _ = write!(
                    out,
                    "#place(top + left, dx: {}pt)[",
                    format_f64(left_pt + sheet_image.x_offset_pt),
                );
                generate_image(out, &sheet_image.image, ctx);
                out.push(']');
            }
        }
    }
}

fn generate_fixed_element(
    out: &mut String,
    elem: &FixedElement,
    page_height_pt: f64,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let placement_y = fixed_raster_placement_y(elem, page_height_pt);
    // Use Typst's place() for absolute positioning
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt",
        format_f64(elem.x),
        format_f64(placement_y),
    );
    out.push_str(")[\n");

    match &elem.kind {
        FixedElementKind::TextBox(text_box) => generate_fixed_text_box(out, elem, text_box, ctx)?,
        FixedElementKind::Image(img) => {
            if let Some(ref shadow) = img.shadow {
                let dir_rad = shadow.direction.to_radians();
                let dx = shadow.distance * dir_rad.cos();
                let dy = shadow.distance * dir_rad.sin();
                if shadow.blur_radius > 0.0 {
                    // SVG filters are rasterised by Typst at four samples per
                    // output point, giving pictures the same continuous
                    // Gaussian ramp as geometric shapes (issue #1309).
                    shapes::write_blurred_shadow_asset(
                        out,
                        ctx,
                        &ShapeKind::Rectangle,
                        (elem.width, elem.height),
                        &img.stroke,
                        shadow,
                        (dx, dy),
                    );
                } else {
                    // A picture's `a:ln` straddles its frame exactly as a shape's
                    // does, so the silhouette PowerPoint casts is the frame grown
                    // by half that width (issue #1057).
                    let outline_outset: f64 = shadow_outline_outset(&img.stroke);
                    // …and it turns the frame's corners the way that `a:ln`'s
                    // join does, so the shadow duplicate carries the same arc
                    // rather than a mitre (issue #1138).
                    let silhouette_radius: f64 = shadow_silhouette_corner_radius(&img.stroke);
                    let alpha: u8 = shadow_alpha(shadow);
                    {
                        let expansion = outline_outset;
                        let layer_width: f64 = (elem.width + 2.0 * expansion).max(0.0);
                        let layer_height: f64 = (elem.height + 2.0 * expansion).max(0.0);
                        let radius: f64 =
                            clamp_ring_corner_radius(silhouette_radius, layer_width, layer_height);
                        let _ = writeln!(
                            out,
                            "#place(top + left, dx: {}pt, dy: {}pt, rect(width: {}pt, height: {}pt, radius: {}pt, fill: rgb({}, {}, {}, {})))",
                            format_f64(dx - expansion),
                            format_f64(dy - expansion),
                            format_f64(layer_width),
                            format_f64(layer_height),
                            format_f64(radius),
                            shadow.color.r,
                            shadow.color.g,
                            shadow.color.b,
                            alpha,
                        );
                    }
                }
            }
            // `a:xfrm/@rot` and `@flipH`/`@flipV` on a `p:pic` (issues #682,
            // #1017). PowerPoint mirrors and turns the frame about its own
            // centre, and the measured GT keeps that centre fixed while the
            // drawn extent grows to the rotated bounding box.
            //
            // `origin: center` cannot express that centre: Typst resolves it
            // against the frame it lays the body out in, and that frame is
            // clamped to the region. A picture box taller or wider than the
            // slide then turns about the slide's midpoint and lands
            // translated — the CONTOSO deck's 856.8pt artwork on a 540pt
            // slide moved 119pt off its seat (issue #1032). A `top + left`
            // origin sits at the frame's own corner, which no clamp can
            // move, so pivot there and carry the difference in an enclosing
            // `#move`. The element is absolutely placed, so no reflow is
            // needed and the box is left alone.
            let rotation = img.rotation_deg.filter(|deg| *deg != 0.0);
            let flipped = img.flip_h || img.flip_v;
            if rotation.is_some() || flipped {
                let (dx, dy): (f64, f64) = centre_pivot_shift(
                    img.width.unwrap_or(elem.width),
                    img.height.unwrap_or(elem.height),
                    rotation.unwrap_or(0.0),
                    img.flip_h,
                    img.flip_v,
                );
                let _ = write!(
                    out,
                    "#move(dx: {}pt, dy: {}pt)[",
                    format_f64(dx),
                    format_f64(dy)
                );
            }
            if let Some(deg) = rotation {
                let _ = write!(out, "#rotate({}deg, origin: top + left)[", format_f64(deg));
            }
            if flipped {
                let _ = write!(
                    out,
                    "#scale(x: {}, y: {}, origin: top + left)[",
                    if img.flip_h { "-100%" } else { "100%" },
                    if img.flip_v { "-100%" } else { "100%" },
                );
            }
            generate_image(out, img, ctx);
            if flipped {
                out.push(']');
            }
            if rotation.is_some() {
                out.push(']');
            }
            if rotation.is_some() || flipped {
                out.push(']');
            }
            // Render image border as a separate overlay so that #image()
            // dimensions are not affected by Typst's #box(stroke:) sizing.
            if let Some(ref stroke) = img.stroke {
                let _ = write!(
                    out,
                    "]\n#place(top + left, dx: {}pt, dy: {}pt)[\n",
                    format_f64(elem.x),
                    format_f64(placement_y),
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
            generate_shape(out, shape, elem.width, elem.height, ctx);
        }
        FixedElementKind::Table(table) => {
            let enclosing = ctx.table_uses_powerpoint_line_box;
            ctx.table_uses_powerpoint_line_box = true;
            let result = generate_table(out, table, ctx);
            ctx.table_uses_powerpoint_line_box = enclosing;
            result?;
        }
        FixedElementKind::SmartArt(smartart) => {
            generate_smartart(out, smartart, elem.width, elem.height);
        }
        FixedElementKind::Chart(chart) => {
            // A slide chart is laid out at its `<p:graphicFrame>` extent, the
            // way every other fixed element already honours its own.
            generate_chart_in(out, chart, Some((elem.width, elem.height)));
        }
    }

    out.push_str("]\n");
    Ok(())
}

/// How far a box mirrored and turned about its **top-left corner** has to move
/// to land where PowerPoint's **centre** pivot would have put it. Pictures pass
/// their mirror flags; a text box turns without mirroring and passes `false`.
///
/// With a `top + left` origin the mirror maps the box to `(±u, ±v)` and the
/// turn follows as `R`, so the composed map is `R(S(p))`. PowerPoint's is
/// `C + R(S(p) + m - C)`, where `C` is the box centre and `m` returns each
/// mirrored axis to the box (`width` when `flip_h`, `height` when `flip_v`).
/// Subtracting the two leaves `C + R(m) - R(C)`, independent of `p` — a plain
/// translation, which is why one `#move` can carry it (issue #1032).
fn centre_pivot_shift(
    width: f64,
    height: f64,
    rotation_deg: f64,
    flip_h: bool,
    flip_v: bool,
) -> (f64, f64) {
    let (centre_x, centre_y): (f64, f64) = (width / 2.0, height / 2.0);
    let (mirror_x, mirror_y): (f64, f64) = (
        if flip_h { width } else { 0.0 },
        if flip_v { height } else { 0.0 },
    );
    let (sin, cos): (f64, f64) = rotation_deg.to_radians().sin_cos();
    (
        centre_x + (cos * mirror_x - sin * mirror_y) - (cos * centre_x - sin * centre_y),
        centre_y + (sin * mirror_x + cos * mirror_y) - (sin * centre_x + cos * centre_y),
    )
}

/// Keep an upright raster's PDF-space bottom edge from rounding below its
/// exact DrawingML coordinate when Typst converts geometry to `f32`.
///
/// At a 324pt top and 183.6pt height on a 540pt slide, the exact bottom is
/// 32.4pt. Typst 0.14/krilla emits 32.399994pt, putting the source's bottom
/// hairline halfway into a second device row at 150 DPI. Moving the top to the
/// preceding `f32` only when that subtraction rounds downward makes the PDF
/// matrix land above the exact edge without a visible point-space offset
/// (issue #666).
fn fixed_raster_placement_y(elem: &FixedElement, page_height_pt: f64) -> f64 {
    let FixedElementKind::Image(image) = &elem.kind else {
        return elem.y;
    };
    if image.format == ImageFormat::Svg
        || image.rotation_deg.is_some_and(|degrees| degrees != 0.0)
        || image.crop.is_some_and(|crop| !crop.is_empty())
        || image.clip_shape.is_some()
        || elem.y <= 0.0
    {
        return elem.y;
    }

    let exact_bottom = page_height_pt - elem.y - elem.height;
    let f32_bottom = (page_height_pt as f32 - (elem.y as f32 + elem.height as f32)) as f64;
    if f32_bottom < exact_bottom {
        f64::from((elem.y as f32).next_down())
    } else {
        elem.y
    }
}

fn generate_fixed_text_box(
    out: &mut String,
    elem: &FixedElement,
    text_box: &TextBoxData,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // Shape rotation (`<a:xfrm rot>`): the content lays out in the unrotated
    // box and the whole result turns about its centre, so `reflow: false`
    // keeps the placed geometry where the slide puts it. This wraps the
    // vertical-text case below, which rotates the glyphs inside the box.
    //
    // `origin: center` cannot name that centre: Typst resolves it against the
    // frame it lays the body out in, and that frame is clamped to the region,
    // so a box taller or wider than the slide turns about the slide's midpoint
    // and lands translated. Pivot on the frame's own top-left corner, which no
    // clamp can move, and carry the difference in an enclosing `#move` — the
    // same correction #1032 made for a picture (issue #1078). A box that fits
    // the region is unaffected: the shift is exactly what `origin: center`
    // resolved to there.
    if let Some(rotation) = text_box.shape_rotation_deg.filter(|deg| *deg != 0.0) {
        let mut inner: TextBoxData = text_box.clone();
        inner.shape_rotation_deg = None;
        let (dx, dy): (f64, f64) = centre_pivot_shift(
            elem.width.max(0.0),
            elem.height.max(0.0),
            rotation,
            false,
            false,
        );
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[#rotate({}deg, origin: top + left, reflow: false)[",
            format_f64(dx),
            format_f64(dy),
            format_f64(rotation)
        );
        generate_fixed_text_box(out, elem, &inner, ctx)?;
        out.push_str("]]\n");
        return Ok(());
    }

    // Vertical text (`<a:bodyPr vert>`): lay the content out in a box with
    // swapped dimensions and rotate it around the element center; the outer
    // geometry stays unrotated, matching PowerPoint.
    if let Some(rotation) = text_box.text_rotation_deg
        && elem.width > 0.0
        && elem.height > 0.0
    {
        let mut inner: TextBoxData = text_box.clone();
        inner.text_rotation_deg = None;
        // Remap the insets into the rotated coordinate system: the side a
        // padding lands on after rotation must carry the original value
        // (e.g. for 270° the original top inset becomes the inner left).
        let padding = &text_box.padding;
        inner.padding = if (rotation - 270.0).abs() < 1.0 {
            crate::ir::Insets {
                left: padding.top,
                top: padding.right,
                right: padding.bottom,
                bottom: padding.left,
            }
        } else {
            crate::ir::Insets {
                left: padding.bottom,
                top: padding.left,
                right: padding.top,
                bottom: padding.right,
            }
        };
        let swapped_elem = FixedElement {
            x: elem.x,
            y: elem.y,
            width: elem.height,
            height: elem.width,
            kind: elem.kind.clone(),
        };
        // The outer #place pins the top-left of a width x height region;
        // center the swapped box on that region before rotating in place.
        //
        // The swap trades the box's two dimensions, so a box wider than the
        // slide is high lays out as a frame taller than the region and meets
        // the same clamp as the shape-rotation case above. Pivot on the corner
        // for the same reason, and fold both translations into the one `#move`
        // (issue #1078).
        let (pivot_dx, pivot_dy): (f64, f64) = centre_pivot_shift(
            swapped_elem.width,
            swapped_elem.height,
            rotation,
            false,
            false,
        );
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[#rotate({}deg, origin: top + left, reflow: false)[",
            format_f64((elem.width - elem.height) / 2.0 + pivot_dx),
            format_f64((elem.height - elem.width) / 2.0 + pivot_dy),
            format_f64(rotation)
        );
        generate_fixed_text_box(out, &swapped_elem, &inner, ctx)?;
        out.push_str("]]\n");
        return Ok(());
    }

    let outer_width_pt: f64 = elem.width.max(0.0);
    let outer_height_pt: f64 = elem.height.max(0.0);
    let original_inner_width_pt: f64 =
        (outer_width_pt - text_box.padding.left - text_box.padding.right).max(0.0);
    // A title saved with empty hard-break lines on both boundaries still has
    // one visible line in PowerPoint. Letting the inline fallback emit those
    // markers as full Typst lines can push that title below its centred banner
    // (issue #1334). A paragraph eligible for the measured hard-break stack
    // keeps that existing path: slide 6 of the same fixture is its control.
    let text_box_without_empty_boundary_lines: Option<TextBoxData> =
        powerpoint_centered_single_line_text_box(text_box, original_inner_width_pt);
    let text_box: &TextBoxData = text_box_without_empty_boundary_lines
        .as_ref()
        .unwrap_or(text_box);

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
        generate_fixed_text_paragraph(out, &raw_paragraph, Some(inner_width_pt), true, false)?;
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
        generate_fixed_text_paragraph(out, paragraph, Some(inner_width_pt), false, false)?;
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
            generate_fixed_text_box_block(
                out,
                block,
                ctx,
                Some(inner_width_pt),
                text_box.no_wrap,
                matches!(text_box.vertical_align, TextBoxVerticalAlign::Top),
            )?;
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
/// One footer story rendered as the value of `#set page(footer: …)`.
#[derive(Default)]
struct FlowFooterValue {
    markup: String,
    /// Whether this story needs `footer-descent: 0pt`.
    wants_zero_descent: bool,
}

/// Round a point measurement to hundredths, keeping float noise (62.35 - 35.4)
/// out of the emitted source.
fn round_to_hundredths(value_pt: f64) -> f64 {
    (value_pt * 100.0).round() / 100.0
}

/// The gap the footer's last drawn paragraph reserves below its last line, in
/// points.
///
/// Word bottom-anchors a footer story on `w:pgMar/@w:footer` and then keeps the
/// last paragraph's resolved `w:spacing w:after` between its last line and that
/// anchor, exactly as it keeps that gap below a body paragraph. A package
/// stating no `w:docDefaults/w:pPrDefault` resolves Word's built-in `Normal`
/// 8pt there, which is why an unreserved band printed a whole `w:after` low
/// (issue #1195).
///
/// The paragraph asked is the last one the story actually draws: a
/// page-anchored frame is positioned against the page rather than laid out
/// above the anchor, so its gap is not what sits there. A story stating no gap
/// at all — every format but DOCX — reserves nothing and keeps the band it had.
fn hf_trailing_space_after_pt(hf: &HeaderFooter) -> f64 {
    hf.paragraphs
        .iter()
        .rev()
        .find(|paragraph| hf_paragraph_is_emitted(paragraph))
        .and_then(|paragraph| paragraph.style.space_after)
        .filter(|gap| *gap > 0.0)
        .unwrap_or(0.0)
}

/// Build the `#set page(footer: …)` value for one footer story (issue #846).
fn flow_footer_value(footer: &HeaderFooter, page: &FlowPage, ctx: &mut GenCtx) -> FlowFooterValue {
    let mut value = String::new();
    // `w:sectPr/w:pgMar/@w:footer` is the distance from the bottom page
    // edge to the *bottom* of the footer, which then grows upward.
    // `footer-descent: 0pt` puts Typst's footer origin on the bottom margin
    // line, so a block spanning exactly that gap ends where Word's footer
    // ends; bottom-aligning the content inside it reproduces the upward
    // growth without having to measure the content. What the band's text
    // bottom edge has to match is set below, from the footer face's own line
    // box.
    let footer_band: Option<f64> = footer
        .distance_from_edge
        .map(|distance| round_to_hundredths(page.margins.bottom - distance))
        .filter(|band| *band > 0.0);
    if let Some(band) = footer_band {
        if hf_needs_context(footer) {
            value.push_str("context ");
        }
        // The band's bottom is the footer's bottom, so what sits between it
        // and the last baseline is that line's own sub-baseline share.
        // Typst's `"descender"` is its normalised one, which is the wrong
        // quantity for a face whose line box carries more below the
        // baseline than its descender does (issue #630).
        let bottom_edge: String = footer
            .paragraphs
            .last()
            .map(hf_paragraph_metric_runs)
            .and_then(|runs| text::word_line_box_descent_em(&runs))
            .map(|descent_em| format!("-{}em", format_f64(descent_em)))
            .unwrap_or_else(|| "\"descender\"".to_string());
        // Word reserves the last paragraph's own `w:spacing w:after` between
        // its last line and that anchor, so the drawn band stops that far above
        // `w:pgMar/@w:footer` rather than on it. A gap deeper than the band
        // itself would invert the block, so the band is as far as the story can
        // be lifted (issue #1195).
        let drawn_band: f64 =
            round_to_hundredths((band - hf_trailing_space_after_pt(footer)).max(0.0));
        let _ = write!(
            value,
            "block(width: 100%, height: {}pt)[#set text(bottom-edge: {bottom_edge}); #place(bottom, block(width: 100%)[",
            format_f64(drawn_band)
        );
        generate_flow_hf_content(&mut value, footer, ctx);
        value.push_str("])]");
        return FlowFooterValue {
            markup: value,
            wants_zero_descent: true,
        };
    }
    if hf_needs_stack_offset(footer) {
        value.push_str("context { let footer_content = block(width: 100%)[");
        generate_flow_hf_content(&mut value, footer, ctx);
        value.push_str("]; move(dy: -measure(footer_content).height / 2)[#footer_content] }");
    } else {
        if hf_needs_context(footer) {
            value.push_str("context ");
        }
        value.push('[');
        generate_flow_hf_content(&mut value, footer, ctx);
        value.push(']');
    }
    FlowFooterValue {
        markup: value,
        wants_zero_descent: false,
    }
}

/// One header story rendered as the value of `#set page(header: …)`.
#[derive(Default)]
struct FlowHeaderValue {
    /// The Typst content expression, ready to place after `header:`.
    markup: String,
    /// Whether this story needs `header-ascent: 0pt` to sit where Word puts it.
    wants_zero_ascent: bool,
}

/// Label a section's first page carries, so a first-page header can ask which
/// page it is on without assuming the section starts the document.
fn section_first_page_label(section_index: usize) -> String {
    format!("o2p-sec-{section_index}")
}

/// Typst expression for the page a section starts on (issue #846).
fn section_first_page_expression(section_index: usize) -> String {
    format!(
        "query(<{}>).first().location().page()",
        section_first_page_label(section_index)
    )
}

/// Build the `#set page(header: …)` value for one header story.
///
/// Split out of the page setup so `<w:titlePg/>`'s two stories are built the
/// same way and differ only in which one a page picks (issue #846).
fn flow_header_value(
    header: &HeaderFooter,
    page: &FlowPage,
    size: &PageSize,
    top_margin_pt: f64,
    ctx: &mut GenCtx,
) -> FlowHeaderValue {
    let mut out = String::new();
    // `w:sectPr/w:pgMar/@w:header` is the distance from the top page edge to
    // the *top* of the header, which then grows downward. Typst anchors
    // header content by its bottom, so anything added below the text — a
    // `w:pBdr` rule and its `w:space` gap — would otherwise push the text
    // up. `header-ascent: 0pt` puts the origin on the top margin line, and a
    // band of exactly that gap holds the content against Word's header top.
    let header_band: Option<f64> = header
        .distance_from_edge
        // Keep float noise out of the emitted source.
        .map(|distance| ((top_margin_pt - distance) * 100.0).round() / 100.0)
        .filter(|band| *band > 0.0);
    if let Some(band) = header_band {
        match header_band_shift_pt(header) {
            Some(shift) => {
                write_shifted_header_band(&mut out, header, band, shift, page, size, ctx)
            }
            None => {
                if hf_needs_context(header) {
                    out.push_str("context ");
                }
                let _ = write!(
                    out,
                    "block(width: 100%, height: {}pt)[#place(top, block(width: 100%)[",
                    format_f64(band)
                );
                generate_flow_hf_content(&mut out, header, ctx);
                out.push_str("])]");
            }
        }
        return FlowHeaderValue {
            markup: out,
            wants_zero_ascent: true,
        };
    }
    if hf_needs_context(header) {
        out.push_str("context [");
    } else {
        out.push('[');
    }
    generate_flow_hf_content(&mut out, header, ctx);
    out.push(']');
    FlowHeaderValue {
        markup: out,
        wants_zero_ascent: false,
    }
}

/// The top margin the page needs, in points.
///
/// Word's header band is `w:top - w:header`, and a header taller than it grows
/// the top margin rather than overprinting the body. Before this, a header of
/// four 12pt lines in a 26.95pt band interleaved its last two lines with the
/// body text; the reference export pushes the body below all four (issue #736).
///
/// Only growth is possible, and only for a story that outgrows its band: one
/// that fits leaves `w:top` alone, so the common case emits exactly what it
/// always did. A header whose face cannot be measured also leaves it alone
/// rather than guessing.
fn flow_page_top_margin_pt(page: &FlowPage) -> f64 {
    // One margin serves the whole section, so where `w:titlePg` gives the first
    // page its own story both have to fit it — taking the default story alone
    // would leave a taller first-page header overprinting page one, which is
    // the same defect one page in (#846 added that second story).
    [page.header.as_ref(), page.first_header.as_ref()]
        .into_iter()
        .flatten()
        .filter(|header| hf_has_flow_content(header))
        .filter_map(|header| {
            // The band runs from the `w:header` line down, so a story that
            // overflows needs the margin to reach the bottom of its content.
            let reach: f64 =
                header.distance_from_edge.unwrap_or(0.0) + hf_content_height_pt(header)?;
            // A story that fits its band is left exactly as it renders today,
            // clamp and all. Growing those too would move the body a fraction
            // of a point on ordinary documents — measured at 0.63pt for a
            // single 12pt Malgun line well inside its band — which is a
            // different change from stopping an overflow, and one no reference
            // export here justifies.
            if reach <= page.margins.top {
                return None;
            }
            // An overflowing story needs room for the ascent correction as
            // well, or `write_shifted_header_band` clamps the shift straight
            // back off: the two-line Malgun letterhead measured 41.4984pt of
            // content against a 41.5pt band, 0.0016pt of room for a 6.84pt
            // shift (#629 seats that baseline).
            Some(reach + header_band_shift_pt(header).unwrap_or(0.0).max(0.0))
        })
        .fold(page.margins.top, f64::max)
}

/// Height a header or footer story's lines take, in points.
///
/// One natural line per paragraph plus whatever its `w:pBdr` rules and their
/// `w:space` reserve. Wrapping is not modelled: a header paragraph that wraps
/// would measure short, which grows the margin less than it should rather than
/// more, so the failure stays on the side of the current behaviour.
fn hf_content_height_pt(hf: &HeaderFooter) -> Option<f64> {
    let mut total: f64 = 0.0;
    for paragraph in &hf.paragraphs {
        let runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
        total += text::word_line_advance_pt(&runs)?;
        if let Some(border) = paragraph.border.as_ref() {
            for (side, space) in [
                (border.top.as_ref(), paragraph.border_space.map(|i| i.top)),
                (
                    border.bottom.as_ref(),
                    paragraph.border_space.map(|i| i.bottom),
                ),
            ] {
                if let Some(side) = side {
                    total += side.width + space.filter(|gap| *gap > 0.0).unwrap_or(0.5);
                }
            }
        }
    }
    Some(total)
}

fn write_flow_page_setup(out: &mut String, page: &FlowPage, size: &PageSize, ctx: &mut GenCtx) {
    // A section may declare only a first-page story — `w:titlePg` with just a
    // `first` reference means pages after the first carry none — so the
    // shortcut has to ask about those too (issue #846).
    if page.header.is_none()
        && page.footer.is_none()
        && page.first_header.is_none()
        && page.first_footer.is_none()
    {
        write_page_setup(out, size, &page.margins);
        return;
    }

    // A header taller than `w:top - w:header` grows the margin instead of
    // overprinting the body (issue #736).
    let top_margin_pt: f64 = flow_page_top_margin_pt(page);
    let _ = write!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt)",
        format_f64(size.width),
        format_f64(size.height),
        format_f64(top_margin_pt),
        format_f64(page.margins.bottom),
        format_f64(page.margins.left),
        format_f64(page.margins.right),
    );

    // A section with `<w:titlePg/>` draws a different story on its first page,
    // so the header value becomes a choice made per page rather than one block
    // (issue #846). Both variants share the band and ascent, which come from
    // the section's own margins, so they are built the same way.
    let default_header = page
        .header
        .as_ref()
        .filter(|header| hf_has_flow_content(header));
    let first_header = page
        .first_header
        .as_ref()
        .filter(|header| hf_has_flow_content(header));
    if first_header.is_some() || default_header.is_some() {
        let default_value = default_header
            .map(|header| flow_header_value(header, page, size, top_margin_pt, ctx))
            .unwrap_or_default();
        let first_value = first_header
            .map(|header| flow_header_value(header, page, size, top_margin_pt, ctx))
            .unwrap_or_default();
        // The ascent is a page property, so it is set when either story wants
        // it; a story that does not is unaffected by it.
        if default_value.wants_zero_ascent || first_value.wants_zero_ascent {
            out.push_str(", header-ascent: 0pt");
        }
        match first_header {
            Some(_) => {
                let _ = write!(
                    out,
                    ", header: context {{ if here().page() == {} {{ {} }} else {{ {} }} }}",
                    section_first_page_expression(ctx.flow_section_index),
                    first_value.markup,
                    default_value.markup
                );
            }
            None => {
                let _ = write!(out, ", header: {}", default_value.markup);
            }
        }
    }

    // The footer takes the same first-page choice the header does (issue #846).
    let default_footer = page
        .footer
        .as_ref()
        .filter(|footer| hf_has_flow_content(footer));
    let first_footer = page
        .first_footer
        .as_ref()
        .filter(|footer| hf_has_flow_content(footer));
    if first_footer.is_some() || default_footer.is_some() {
        let default_value = default_footer
            .map(|footer| flow_footer_value(footer, page, ctx))
            .unwrap_or_default();
        let first_value = first_footer
            .map(|footer| flow_footer_value(footer, page, ctx))
            .unwrap_or_default();
        if default_value.wants_zero_descent || first_value.wants_zero_descent {
            out.push_str(", footer-descent: 0pt");
        }
        match first_footer {
            Some(_) => {
                let _ = write!(
                    out,
                    ", footer: context {{ if here().page() == {} {{ {} }} else {{ {} }} }}",
                    section_first_page_expression(ctx.flow_section_index),
                    first_value.markup,
                    default_value.markup
                );
            }
            None => {
                let _ = write!(out, ", footer: {}", default_value.markup);
            }
        }
    }

    // `<wp:anchor behindDoc="1">` puts the shape under the page's own content,
    // which is where a decorative banner belongs — over it, the banner would
    // hide the body text it sits behind (issue #961).
    for (layer, behind_text) in [("background", true), ("foreground", false)] {
        // `<w:titlePg/>` gives page one its own stories, and on this invoice
        // they are the only ones carrying a banner, so the layers are gated
        // per page exactly as the flow header and footer are (issue #961).
        let default_markup: String = page_anchored_layer_markup(
            page.header.as_ref(),
            page.footer.as_ref(),
            page,
            size,
            behind_text,
            ctx,
        );
        let first_stated: bool = page.first_header.is_some() || page.first_footer.is_some();
        let first_markup: String = match first_stated {
            true => page_anchored_layer_markup(
                page.first_header.as_ref().or(page.header.as_ref()),
                page.first_footer.as_ref().or(page.footer.as_ref()),
                page,
                size,
                behind_text,
                ctx,
            ),
            false => String::new(),
        };
        if default_markup.is_empty() && first_markup.is_empty() {
            continue;
        }
        match first_stated {
            true => {
                let _ = write!(
                    out,
                    ", {layer}: context {{ if here().page() == {} [{}] else [{}] }}",
                    section_first_page_expression(ctx.flow_section_index),
                    first_markup,
                    default_markup
                );
            }
            false => {
                let _ = write!(out, ", {layer}: [{default_markup}]");
            }
        }
    }

    out.push_str(")\n");
}

/// Where a `<wp:align>` puts a box of `extent` inside a reference frame of
/// `available`. An unstated alignment pins to the start, which is where an
/// absent offset already put it. An unknown extent counts as zero, so a
/// `Center` or `End` alignment still resolves — to the frame's midpoint or its
/// far edge, not to the start.
///
/// The result is not clamped to the frame: the invoice's header banner is
/// 609.12pt wide on a 595.28pt page, and centring it genuinely hangs 6.92pt
/// off each edge (issue #961).
fn aligned_offset(
    align: Option<crate::ir::FrameAlign>,
    available: f64,
    extent: Option<f64>,
) -> f64 {
    let extent: f64 = extent.unwrap_or(0.0);
    match align {
        Some(crate::ir::FrameAlign::Center) => (available - extent) / 2.0,
        Some(crate::ir::FrameAlign::End) => available - extent,
        _ => 0.0,
    }
}

/// Resolve the page-space x origin of an anchored header/footer text column.
///
/// An explicit `posOffset` is already a concrete coordinate and is left exact.
/// A page-left `<wp:align>` WPS box takes Writer's measured inner text-origin
/// seat in addition to its declared inset; keeping this correction here rather
/// than changing the inset preserves the box's parsed natural width.
fn page_anchored_hf_text_origin_x(frame: &HeaderFooterFrame, page_width: f64) -> f64 {
    let aligned_x: f64 = frame
        .x
        .unwrap_or_else(|| aligned_offset(frame.horizontal_align, page_width, frame.width));
    let writer_seat: f64 =
        if frame.x.is_none() && frame.horizontal_align == Some(crate::ir::FrameAlign::Start) {
            WRITER_PAGE_LEFT_TEXT_ORIGIN_SEAT_PT
        } else {
            0.0
        };
    aligned_x + frame.inset_left + writer_seat
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

/// One page layer's markup: the anchored shapes both stories put on it, plus
/// their framed paragraphs below the body in the background.
///
/// Empty when neither story draws anything there, which is the caller's signal
/// to leave the layer off the `#set page` entirely.
fn page_anchored_layer_markup(
    header: Option<&HeaderFooter>,
    footer: Option<&HeaderFooter>,
    page: &FlowPage,
    size: &PageSize,
    behind_text: bool,
    ctx: &mut GenCtx,
) -> String {
    let mut markup = String::new();
    for hf in header.into_iter().chain(footer) {
        generate_page_anchored_hf_shapes(&mut markup, hf, size, behind_text, ctx);
        // Word paints the main story after header/footer framed paragraphs, so
        // body drawings can cover those frames. Typst's page background is the
        // layer that preserves that cross-story order (issue #1408).
        if behind_text {
            generate_page_anchored_hf_frames(&mut markup, hf, size, page.margins.right, ctx);
        }
    }
    markup
}

/// Place a story's page-anchored shapes, each against the page rather than in
/// the story's flow.
///
/// Unlike a framed paragraph, the shape's own extent is known, so both axes
/// resolve here without the block's height entering into it (issue #961).
fn generate_page_anchored_hf_shapes(
    out: &mut String,
    hf: &HeaderFooter,
    page_size: &PageSize,
    behind_text: bool,
    ctx: &mut GenCtx,
) {
    for shape in &hf.shapes {
        if shape.behind_text != behind_text || !is_page_anchored_frame(&shape.frame) {
            continue;
        }
        let x: f64 = shape.frame.x.unwrap_or_else(|| {
            aligned_offset(
                shape.frame.horizontal_align,
                page_size.width,
                Some(shape.width),
            )
        });
        let y: f64 = shape.frame.y.unwrap_or_else(|| {
            aligned_offset(
                shape.frame.vertical_align,
                page_size.height,
                Some(shape.height),
            )
        });
        // The `#box` is not decoration: `#rotate` turns its body about the
        // frame it was laid out into, and inside a bare `#place` that frame is
        // the page's width. The invoice's footer band is 627.84pt on a 595.28pt
        // page, so a 180-degree turn about the page centre slid it 32.56pt off
        // the right edge (issue #961).
        let _ = write!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt)[#box(width: {}pt, height: {}pt)[",
            format_f64(x),
            format_f64(y),
            format_f64(shape.width),
            format_f64(shape.height)
        );
        generate_shape(out, &shape.shape, shape.width, shape.height, ctx);
        out.push_str("]]");
    }
}

fn generate_flow_hf_content(out: &mut String, hf: &HeaderFooter, ctx: &mut GenCtx) {
    // The story's paragraphs are joined by a `\\` line break below, so they are
    // one Typst paragraph and `par(leading:)` is what separates their lines.
    // Typst advances a line by `top-edge + bottom-edge + leading`, and the story
    // set no leading, so it took the 0.65em default: an 8pt Arial header advanced
    // 10.9305pt where Word advances 9.1992pt. Stating the remainder after the
    // edges leaves the first baseline — seated against the top edge by #629 —
    // exactly where it was, and corrects only the advance between lines (#735).
    //
    // Set once for the story rather than per paragraph: wrapping each paragraph
    // in its own content block makes it a block, and Typst then puts
    // `par(spacing:)` between them — a different and much larger gap.
    if let Some(leading) = hf
        .paragraphs
        .iter()
        .find(|paragraph| hf_paragraph_is_emitted(paragraph))
        .and_then(|paragraph| {
            let runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
            text::word_hf_line_leading_pt(&runs, 0.0)
        })
    {
        let _ = writeln!(out, "#set par(leading: {}pt)", format_f64(leading));
    }
    let mut is_first: bool = true;
    for paragraph in &hf.paragraphs {
        if !hf_paragraph_is_emitted(paragraph) {
            continue;
        }
        if !is_first {
            out.push_str("\\\n");
        }
        generate_hf_styled_paragraph(out, paragraph, ctx);
        is_first = false;
    }
}

/// Whether [`generate_flow_hf_content`] writes this paragraph into the story.
///
/// Page-anchored frames are drawn separately in the page background, and an
/// empty paragraph carrying neither content nor a border produces nothing.
/// Shared with the band placement so the two cannot disagree about which
/// paragraph comes first (issue #629).
fn hf_paragraph_is_emitted(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    !paragraph.frame.as_ref().is_some_and(is_page_anchored_frame)
        && hf_paragraph_has_content(paragraph)
}

/// Below this the shift is not worth an extra `context` block: it is under a
/// tenth of the 0.24pt grid the native exports this is calibrated against
/// quantise to, and skipping it keeps the plain band form for faces whose
/// ascent and cap height happen to coincide.
const MIN_HEADER_BAND_SHIFT_PT: f64 = 0.02;

/// How far a pinned header band must move for its first baseline to land where
/// Word puts it, in points, positive downward.
///
/// Word seats a header story's first baseline one font ascent below
/// `w:pgMar/@w:header`; the compiler seats it one cap height below the same
/// line. Shifting the band by the difference lands that baseline on Word's
/// without touching a single line box, so the story's own baseline-to-baseline
/// advance stays exactly what the compiler would produce — declaring the ascent
/// as a `top-edge` instead would widen *every* wrapped line of the paragraph and
/// stretch that advance (issue #629). Preserving it is still the point, and it
/// is what keeps this shift valid now that the advance itself is Word's: the
/// story states its leading separately (issue #735), so a placement fix and a
/// line-advance fix stay independent of each other.
///
/// The band is sized by the first paragraph the story actually emits, which is
/// the one whose top the header distance pins. `None` when that paragraph gives
/// nothing to measure: no styled text at all, a family that does not resolve,
/// or an image, whose height rather than the text's would set the compiler's
/// line ascent.
fn header_band_shift_pt(hf: &HeaderFooter) -> Option<f64> {
    let first: &crate::ir::HeaderFooterParagraph =
        hf.paragraphs.iter().find(|p| hf_paragraph_is_emitted(p))?;
    if first
        .elements
        .iter()
        .any(|element| matches!(element, HFInline::Image(_)))
    {
        return None;
    }
    let runs: Vec<Run> = hf_paragraph_metric_runs(first);
    let shift: f64 = self::text::word_header_band_shift_pt(&runs)?;
    // Keep float noise out of the emitted source.
    let shift: f64 = (shift * 10_000.0).round() / 10_000.0;
    (shift.abs() >= MIN_HEADER_BAND_SHIFT_PT).then_some(shift)
}

/// The runs a header paragraph's line metrics resolve against.
///
/// A `PAGE`/`NUMPAGES` field carries the run properties of the `w:r` holding it
/// and shapes as digits, so it folds in as a synthetic run: a header whose first
/// paragraph is nothing but a page number still has an ascent to seat, and
/// without this the decision would leak onto the second paragraph (issue #629).
/// Positioned tabs contribute no glyphs.
fn hf_paragraph_metric_runs(paragraph: &crate::ir::HeaderFooterParagraph) -> Vec<Run> {
    paragraph
        .elements
        .iter()
        .filter_map(|element| match element {
            HFInline::Run(run) => Some(run.clone()),
            HFInline::PageNumber(style) | HFInline::TotalPages(style) => Some(Run {
                text: "1".to_string(),
                style: style.clone(),
                href: None,
                footnote: None,
            }),
            HFInline::Image(_) | HFInline::PositionedTab(_) => None,
        })
        .collect()
}

/// Emit a pinned header band whose content is moved down by `shift` points.
///
/// The band this shifts within is sized by [`flow_page_top_margin_pt`], which
/// grows the top margin to hold the story's content *and* this shift, so the
/// cap below has slack to spare on the stories it grows for. It binds as before
/// on a story that fits its band, which that growth deliberately leaves alone. It still binds where the height could
/// not be measured — the margin then stays at `w:top` — and it remains the
/// backstop that keeps ink off the body if the two measurements ever disagree.
///
/// Word's growth was unconfirmed when this was written — every native `save as
/// … format PDF` died with AppleEvent -1712 — and is now measured: against a
/// reference export, four 12pt header lines in a 26.95pt band place the body
/// below all four rather than interleaving with them (issue #736).
fn write_shifted_header_band(
    out: &mut String,
    header: &HeaderFooter,
    band: f64,
    shift: f64,
    page: &FlowPage,
    size: &PageSize,
    ctx: &mut GenCtx,
) {
    // `measure` reports the height the story takes in the column it will be
    // laid out in; without the width it would measure in an infinite region and
    // report a wrapped story as a single line.
    let text_width: f64 = (size.width - page.margins.left - page.margins.right).max(0.0);
    // Keep float noise (595.28 - 70.85 - 70.85) out of the emitted source.
    let text_width: f64 = (text_width * 100.0).round() / 100.0;
    out.push_str("context { let header_content = block(width: 100%)[");
    generate_flow_hf_content(out, header, ctx);
    let _ = write!(
        out,
        "]; block(width: 100%, height: {band}pt)[#place(top, dy: calc.min({shift}pt, calc.max(0pt, {band}pt - measure(header_content, width: {text_width}pt).height)), header_content)] }}",
        band = format_f64(band),
        shift = format_f64(shift),
        text_width = format_f64(text_width),
    );
}

fn generate_page_anchored_hf_frames(
    out: &mut String,
    hf: &HeaderFooter,
    page_size: &PageSize,
    right_margin: f64,
    ctx: &mut GenCtx,
) {
    let page_width: f64 = page_size.width;
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
        // `<wp:align>` states the edge rather than an offset, and only the
        // renderer knows the page it is measured against (issue #847).
        let x: f64 = page_anchored_hf_text_origin_x(frame, page_size.width);
        // A box that seats its text at its own bottom edge is placed upward
        // from the page. In the issue #1370 reference, the final 8pt run's
        // baseline is one em above the bottom inset; putting the baseline on
        // the inset itself left that run 8.05pt too low.
        let (anchor, y): (&str, f64) = match frame.bottom_offset {
            Some(gap) => (
                "bottom + left",
                -bottom_seated_frame_baseline_offset_pt(&hf.paragraphs[index..end], gap),
            ),
            None => (
                "top + left",
                frame.y.unwrap_or_else(|| {
                    aligned_offset(frame.vertical_align, page_size.height, frame.height)
                }) + frame.inset_top,
            ),
        };
        // `<a:bodyPr wrap="none">` keeps the paragraph whole, so the box takes
        // its content's natural width instead of the column's. A `#box` shrinks
        // to its content where a `#block` fills the region, which is the whole
        // difference between one line and two (issue #967).
        let container: &str = match frame.wraps_text {
            true => "block",
            false => "box",
        };
        let _ = write!(
            out,
            "#place({anchor}, dx: {}pt, dy: {}pt)[#{container}(",
            format_f64(x),
            format_f64(y)
        );
        if frame.wraps_text {
            let width: f64 = frame
                .width
                .unwrap_or_else(|| (page_width - x - right_margin).max(0.0));
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

/// Distance from the page bottom to the last baseline of a bottom-seated WPS
/// text box.
///
/// DrawingML's `bodyPr anchor="b"` aligns the text area's bottom, not a
/// typographic baseline. The issue #1370 reference leaves the stated bottom
/// inset and then one em for its final one-line paragraph. The
/// framed-paragraph generator emits inline text whose box baseline is its
/// placement origin, so that em has to be part of the placement offset
/// explicitly.
///
/// The largest run in the final paragraph with measurable text approximates
/// that paragraph's final line box. A field is represented by a synthetic run
/// in [`hf_paragraph_metric_runs`], and an unstated size takes the same 11pt
/// default used by the Typst generator. If a frame contains no measurable
/// text, retain the inset-only placement rather than inventing a text height
/// for an image or decoration.
fn bottom_seated_frame_baseline_offset_pt(
    paragraphs: &[crate::ir::HeaderFooterParagraph],
    bottom_inset_pt: f64,
) -> f64 {
    for paragraph in paragraphs.iter().rev() {
        let runs = hf_paragraph_metric_runs(paragraph);
        if runs.is_empty() {
            continue;
        }
        let font_size_pt = runs
            .iter()
            .map(|run| run.style.font_size.unwrap_or(11.0))
            .reduce(f64::max)
            .unwrap_or(11.0)
            .max(0.0);
        return bottom_inset_pt + font_size_pt;
    }
    bottom_inset_pt
}

/// Write the full page setup for a SheetPage, including optional header/footer.
/// Write a sheet page's `#set page(…)`. `drawing_foreground`, when present, is
/// the markup that floats the sheet's drawings above its grid — the page
/// foreground is the only layer that paints after a body preceding it in the
/// source (issue #1168).
fn write_table_page_setup(
    out: &mut String,
    page: &SheetPage,
    size: &PageSize,
    ctx: &mut GenCtx,
    drawing_foreground: Option<&str>,
) {
    if page.header.is_none() && page.footer.is_none() && drawing_foreground.is_none() {
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
        generate_sheet_hf_content(out, header, size, &page.margins, ctx);
        out.push(']');
    }

    if let Some(footer) = &page.footer {
        if let Some(band) = sheet_footer_band_pt(footer, page.margins.bottom) {
            out.push_str(", footer-descent: 0pt, footer: ");
            if hf_needs_context(footer) {
                out.push_str("context ");
            }
            let _ = write!(
                out,
                "block(width: 100%, height: {}pt)[#set text(bottom-edge: {}); #place(bottom, block(width: 100%)[",
                format_f64(band),
                sheet_footer_bottom_edge(footer),
            );
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx);
            out.push_str("])]");
        } else if hf_needs_stack_offset(footer) {
            out.push_str(", footer: context { let footer_content = block(width: 100%)[");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx);
            out.push_str("]; move(dy: -measure(footer_content).height / 2)[#footer_content] }");
        } else if hf_needs_context(footer) {
            out.push_str(", footer: context [");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx);
            out.push(']');
        } else {
            out.push_str(", footer: [");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx);
            out.push(']');
        }
    }

    if let Some(foreground) = drawing_foreground {
        out.push_str(", foreground: ");
        out.push_str(foreground);
    }

    out.push_str(")\n");
}

/// Generate a worksheet header/footer inside Excel's fitted horizontal box.
///
/// Excel keeps the paper geometry fixed, lays a `scaleWithDoc` story out in
/// sheet coordinates, and applies the fit scale afterwards. Its sheet-space
/// clip edges sit on whole points and round outwards. On the issue #1510 A3
/// probe, the 50pt paper-space margins at 0.82 therefore become
/// `floor(50 / 0.82) * 0.82 = 49.2pt` on the left and
/// `ceil((1191 - 50) / 0.82) * 0.82 = 1141.44pt` on the right. Typst lays a
/// page header/footer between the paper-space margins, so move that story to
/// the scaled left edge and give it the scaled width. The run sizes were
/// already multiplied once by pagination; this wrapper does not scale them a
/// second time.
fn generate_sheet_hf_content(
    out: &mut String,
    hf: &HeaderFooter,
    size: &PageSize,
    margins: &Margins,
    ctx: &mut GenCtx,
) {
    let scaled_box: Option<(f64, f64)> = hf
        .sheet_print_scale
        .filter(|scale| *scale > 0.0 && *scale < 1.0)
        .map(|scale| {
            let hundredth = |points: f64| -> f64 { (points * 100.0).round() / 100.0 };
            let left_pt: f64 = hundredth((margins.left / scale).floor() * scale);
            let right_pt: f64 = hundredth(((size.width - margins.right) / scale).ceil() * scale);
            (
                hundredth(left_pt - margins.left),
                hundredth(right_pt - left_pt),
            )
        });

    if let Some((dx_pt, width_pt)) = scaled_box {
        let _ = write!(
            out,
            "#move(dx: {}pt)[#block(width: {}pt)[",
            format_f64(dx_pt),
            format_f64(width_pt),
        );
    }
    generate_hf_content(out, hf, ctx);
    if scaled_box.is_some() {
        out.push_str("]]");
    }
}

/// The band a seated sheet footer spans, from the bottom margin line down to
/// the bottom of its text's line box (issue #1142).
///
/// Excel measures a printed footer up from the paper through
/// `<pageMargins>/@footer`, so `footer-descent: 0pt` pins Typst's footer origin
/// on the bottom margin line and a block spanning the remainder ends where
/// Excel's footer text ends. `None` leaves the story on Typst's own descent:
/// either the seat is unknown, or the footer margin reaches past the bottom
/// margin and there is no band to span.
fn sheet_footer_band_pt(footer: &HeaderFooter, bottom_margin_pt: f64) -> Option<f64> {
    footer
        .distance_from_edge
        // Keep float noise (54 - 23.000000000000004) out of the emitted source.
        .map(|distance| ((bottom_margin_pt - distance) * 100.0).round() / 100.0)
        .filter(|band| *band > 0.0)
}

/// The `bottom-edge` a seated sheet footer's text takes, as a Typst value.
///
/// The band's bottom is where Excel's footer text bottoms out, so what sits
/// between it and the last baseline is that line's own sub-baseline share —
/// the face's bare `hhea` descent, which is what the native exports measure
/// (issue #1142). Typst's `"descender"` is its *normalised* one, a different
/// quantity, and is only the fallback for a line whose face cannot be read.
///
/// The deepest of the footer's paragraphs wins because Excel's left, centre
/// and right sections share one line — [`generate_hf_content`] lays them out
/// as one grid row — and a line bottoms out on whichever of its runs reaches
/// furthest below the baseline. `bottom-edge` is one `set text` value for the
/// whole story, so a footer that really does stack lines takes the deepest
/// face's ratio rather than its last line's; the two differ by well under a
/// point on any face pair measured here.
fn sheet_footer_bottom_edge(footer: &HeaderFooter) -> String {
    footer
        .paragraphs
        .iter()
        .filter_map(|paragraph| {
            text::sheet_line_box_descent_em(&hf_paragraph_metric_runs(paragraph))
        })
        .max_by(|a, b| a.total_cmp(b))
        .map(|descent_em| format!("-{}em", format_f64(descent_em)))
        .unwrap_or_else(|| "\"descender\"".to_string())
}

/// Check if a header/footer contains any context-dependent fields (page number or total pages).
fn hf_needs_context(hf: &HeaderFooter) -> bool {
    hf.paragraphs.iter().any(|p| {
        p.elements
            .iter()
            .any(|e| matches!(e, HFInline::PageNumber(_) | HFInline::TotalPages(_)))
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
    // Excel's left/center/right header sections share one line; stacking
    // them as separate lines pushed sections onto extra rows.
    let alignments: Vec<Option<Alignment>> =
        hf.paragraphs.iter().map(|p| p.style.alignment).collect();
    let is_single_line_sections = hf.paragraphs.len() > 1
        && hf.paragraphs.len() <= 3
        && alignments.iter().all(|a| {
            matches!(
                a,
                Some(Alignment::Left) | Some(Alignment::Center) | Some(Alignment::Right)
            )
        })
        && {
            let mut seen = alignments.clone();
            seen.dedup();
            seen.len() == alignments.len()
        };
    if is_single_line_sections {
        out.push_str("#grid(columns: (1fr, 1fr, 1fr), ");
        for slot in [Alignment::Left, Alignment::Center, Alignment::Right] {
            let _ = write!(out, "[");
            if let Some(para) = hf
                .paragraphs
                .iter()
                .find(|p| p.style.alignment == Some(slot))
            {
                generate_hf_styled_paragraph(out, para, ctx);
            }
            out.push_str("], ");
        }
        out.push(')');
        return;
    }
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

/// Where a header or footer paragraph's `<w:tab/>` runs place their segments.
///
/// Word's running-head idiom declares a right-aligned tab stop at the text
/// edge, or a centre stop and a right stop, and separates the segments with
/// tabs. Those two shapes are what `w:tabs` is used for in a header; anything
/// else keeps the plain advance below.
enum HeaderFooterTabLayout {
    /// `left`, tab, `right`.
    LeftRight(usize),
    /// `left`, tab, `centre`, tab, `right`.
    LeftCenterRight(usize, usize),
}

/// Resolve a header or footer paragraph's tabs against its own tab stops.
///
/// `generate_hf_elements` passed every `<w:tab/>` straight to `generate_run`,
/// which writes the tab into the Typst source as a literal tab character.
/// Typst's markup lexer treats that exactly as it treats a space, so the two
/// segments ended up one space apart and the one a right stop should have
/// pushed to the right margin sat beside the left one — on every page of a
/// document that uses the idiom (issue #579).
///
/// The `#h(1em)` advance below is a different element: `w:ptab`, which states
/// its own alignment rather than referring to a stop.
fn header_footer_tab_layout(
    paragraph: &crate::ir::HeaderFooterParagraph,
) -> Option<HeaderFooterTabLayout> {
    let tabs: Vec<usize> = paragraph
        .elements
        .iter()
        .enumerate()
        .filter(|(_, element)| matches!(element, HFInline::Run(run) if run.text == "\t"))
        .map(|(index, _)| index)
        .collect();
    let stops = paragraph.style.tab_stops.as_deref()?;
    let alignments: Vec<TabAlignment> = stops.iter().map(|stop| stop.alignment).collect();

    match (tabs.as_slice(), alignments.as_slice()) {
        ([tab], [.., TabAlignment::Right]) => Some(HeaderFooterTabLayout::LeftRight(*tab)),
        ([first, second], [TabAlignment::Center, .., TabAlignment::Right]) => {
            Some(HeaderFooterTabLayout::LeftCenterRight(*first, *second))
        }
        _ => None,
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
    let top_border = paragraph
        .border
        .as_ref()
        .and_then(|border| border.top.as_ref());
    // Word letterheads put the rule under the header text (`w:pBdr/w:bottom`)
    // just as often as above it, so both sides stack around the content.
    let bottom_border = paragraph
        .border
        .as_ref()
        .and_then(|border| border.bottom.as_ref());
    let stacks_rules: bool = top_border.is_some() || bottom_border.is_some();
    // `w:pBdr` sides declare their own `w:space` gap in points. Without one,
    // Word still leaves a hairline of clearance, which the 0.5 pt fallback
    // reproduces. Word measures the gap from the text's descender line, so the
    // stack pins the text bottom edge there.
    let space = |declared: Option<f64>| -> f64 { declared.filter(|gap| *gap > 0.0).unwrap_or(0.5) };
    let top_space: f64 = space(paragraph.border_space.map(|insets| insets.top));
    let bottom_space: f64 = space(paragraph.border_space.map(|insets| insets.bottom));

    if stacks_rules {
        out.push_str("#stack(dir: ttb, spacing: 0pt, ");
    }
    if let Some(border) = top_border {
        write_hf_border_rules(out, border);
        let _ = write!(out, ", block(height: {}pt)[], ", format_f64(top_space));
    }
    // Word measures `w:pBdr w:space` from the line's bottom, which for an East
    // Asian line is the 1.3x line box's lower edge rather than the face's
    // descender. Typst's `"descender"` answers with its normalised one —
    // 0.2002em for Malgun Gothic (OS/2 410/2048) against the 0.4412em its line
    // box carries — which left a Korean header's rule 1.98pt high, 51.30pt
    // against Word's 53.28pt (issue #737). Only a story that stacks rules pins
    // it; the rest keep Typst's baseline bottom edge.
    let metric_runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
    let bottom_edge_em: Option<f64> = stacks_rules
        .then(|| text::word_line_box_descent_em(&metric_runs))
        .flatten();
    if stacks_rules {
        let bottom_edge: String = bottom_edge_em
            .map(|descent_em| format!("-{}em", format_f64(descent_em)))
            .unwrap_or_else(|| "\"descender\"".to_string());
        let _ = write!(out, "[#set text(bottom-edge: {bottom_edge});");
    }

    if let Some(index) = right_tab {
        out.push_str("#grid(columns: (1fr, auto), [");
        generate_hf_elements(out, &paragraph.elements[..index], ctx);
        out.push_str("], [");
        generate_hf_elements(out, &paragraph.elements[index + 1..], ctx);
        out.push_str("])");
    } else {
        match header_footer_tab_layout(paragraph) {
            Some(HeaderFooterTabLayout::LeftRight(index)) => {
                out.push_str("#grid(columns: (1fr, auto), [");
                generate_hf_elements(out, &paragraph.elements[..index], ctx);
                out.push_str("], [");
                generate_hf_elements(out, &paragraph.elements[index + 1..], ctx);
                out.push_str("])");
            }
            Some(HeaderFooterTabLayout::LeftCenterRight(first, second)) => {
                out.push_str("#grid(columns: (1fr, auto, 1fr), align: (left, center, right), [");
                generate_hf_elements(out, &paragraph.elements[..first], ctx);
                out.push_str("], [");
                generate_hf_elements(out, &paragraph.elements[first + 1..second], ctx);
                out.push_str("], [");
                generate_hf_elements(out, &paragraph.elements[second + 1..], ctx);
                out.push_str("])");
            }
            None => generate_hf_elements(out, &paragraph.elements, ctx),
        }
    }

    if stacks_rules {
        out.push(']');
    }
    if let Some(border) = bottom_border {
        let _ = write!(out, ", block(height: {}pt)[], ", format_f64(bottom_space));
        write_hf_border_rules(out, border);
    }
    if stacks_rules {
        out.push(')');
    }
}

/// Emit the one or two `line()` blocks a single paragraph-border side needs.
/// Word draws a `double` side as two thin rules, so it takes two blocks.
fn write_hf_border_rules(out: &mut String, border: &BorderSide) {
    write_hf_border_line(out, border, border.style == BorderLineStyle::Double);
    if border.style == BorderLineStyle::Double {
        out.push_str(", ");
        write_hf_border_line(out, border, false);
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
    // Word paints a header rule past both edges of the text column, exactly as
    // it does a body paragraph's. `#move` shifts the line without disturbing
    // the band's layout (issue #644).
    //
    // The rule is wider than the text column by design, and a bare block
    // inherits the enclosing alignment — which for a header paragraph is its
    // own `w:jc`. A right-aligned header therefore pinned the *line's* right
    // edge to the column edge, so the whole overhang fell on the left and the
    // `#move` compounded it: on `03_meeting_minutes_ko` the rule started at
    // 66.53 against 69.41 once it is aligned left, exactly the 2.88pt of two
    // outsets. Word does not align a border by the paragraph's `w:jc`, so the
    // rule states `left` for itself and both ends overhang again (#840).
    //
    // The double-rule path in `typst_gen_text.rs` needs no such statement:
    // `#place(bottom, …)` leaves the horizontal component unset, which Typst
    // resolves to a fixed `start` rather than to the enclosing alignment.
    let _ = write!(
        out,
        "block(height: {}pt)[#align(left)[#move(dx: -{}pt)[#line(length: 100% + {}pt, stroke: (paint: {}, thickness: {}pt, dash: \"{}\"))]]]",
        format_f64(width),
        format_f64(TEXT_COLUMN_DECORATION_OVERHANG_PT),
        format_f64(2.0 * TEXT_COLUMN_DECORATION_OVERHANG_PT),
        rgb(&border.color),
        format_f64(width),
        if border.style == BorderLineStyle::Double {
            "solid"
        } else {
            dash
        }
    );
}

/// Emit a header/footer field result under its run's text properties.
///
/// The field's text is computed by the engine — a page number in whatever
/// numbering format the section states — so the emitter cannot name it and
/// `write_text_params` takes the script-safe kerning answer (issue #628).
fn write_hf_field(out: &mut String, style: &TextStyle, field: &str) {
    if has_text_properties(style) {
        out.push_str("#text(");
        write_text_params(out, style);
        let _ = write!(out, ")[{field}]");
    } else {
        out.push_str(field);
    }
}

fn generate_hf_elements(out: &mut String, elements: &[HFInline], ctx: &mut GenCtx) {
    for element in elements {
        match element {
            HFInline::Run(run) => generate_run(out, run),
            HFInline::Image(image) => generate_image(out, image, ctx),
            // Word applies the containing run's properties to the field
            // result, so the number matches the literals around it.
            // `context` is explicit because a header or footer paragraph
            // carrying a frame is emitted as a `#place` at document level,
            // where the page counter has no context of its own and Typst
            // refuses it. Inside a real header the value is the same.
            HFInline::PageNumber(style) => {
                write_hf_field(
                    out,
                    style,
                    &format!(
                        "#context counter(page).display(\"{}\")",
                        ctx.page_number_format.typst_pattern()
                    ),
                );
            }
            HFInline::TotalPages(style) => {
                write_hf_field(out, style, "#context counter(page).final().first()");
            }
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
            Block::FloatingShape(shape) => generate_floating_shape_overlay(out, shape, ctx),
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

/// The `#set text(...)` a contents entry is laid out in.
///
/// Word lays every entry out in body text — the document default family and
/// size — whatever the heading it points at looks like. Without this the
/// entries fall back to Typst's own 11pt Libertinus Serif, because the entry
/// text is a bare string carrying no formatting of its own (issue #610).
fn toc_entry_text_settings(default_text: Option<&crate::ir::TextStyle>) -> String {
    let Some(default_text) = default_text else {
        return String::new();
    };
    let mut settings = String::new();
    if let Some(size) = default_text.font_size {
        let _ = write!(settings, "set text(size: {}pt); ", format_f64(size));
    }
    settings
}

/// Declare the page-format state and teach the outline to read it back.
///
/// Typst numbers an outline entry from the page counter alone, which renders
/// the arabic count whatever the section it points into declared. The show
/// rule rebuilds the entry with its number run through the format recorded at
/// the entry's own location, so an entry pointing into roman-numbered front
/// matter reads `i` rather than the arabic `1` the page counter alone would
/// print (issue #605). Entry indentation and the leader stay Typst's, because the
/// entry's *styling* is a separate defect (issue #610).
fn write_page_format_state(out: &mut String) {
    // No Office application hangs punctuation into the margin; Typst does, and
    // it leaks two ways. In justified text a line ending in a comma pushed the
    // comma past the right margin (#640), and because a hung glyph leaves the
    // line's layout box, a centred line opening with a hyphen was measured
    // narrow and drawn off centre (#645). Set once for the document so header
    // and footer bands, which do not go through the paragraph settings, are
    // covered too.
    let _ = writeln!(out, "#set text(overhang: false)");
    // How far a justified line may squeeze its spaces. Typst's default floor
    // is two thirds of the space's natural width, which is far looser than
    // Word: on the audited official letter it squeezed 14 spaces to 90.1% each
    // to pull one more syllable onto a line, splitting 예정이오니 across the
    // break where Word wraps the word whole (issue #639).
    //
    // Word does shrink — the issue's "Word only ever stretches" is wrong.
    // Measuring every justified line of the ten native DOCX exports in
    // `tests/golden_mocks/business/expected/docx/`, as the rendered space gap
    // over the face's own `hmtx` advance, Word's spaces run from 0.9332 to
    // 1.3400 of natural. So the real floor sits in (0.9014, 0.9332]: below the
    // tightest line Word does set, above the one it refused.
    //
    // The number here is not that floor, because Typst's line breaker also
    // *prices* a shrink against the allowance it is given — halve the
    // allowance and a 6.6% squeeze it used to prefer loses to the alternative
    // break. So the constant is calibrated on the corpus instead: the letter's
    // over-shrink disappears at every value tested up to 85%, and at 88% the
    // research report's legitimate 0.9344 line starts wrapping early. 80% is
    // the middle of that window. At it, the only line that moves anywhere in
    // the 30-file corpus is the one this issue reports, and the three files
    // whose lines Word genuinely compresses keep their 0.9344/0.9639/0.9653
    // spaces against Word's 0.9332/0.9634/0.9636.
    //
    // Only the floor moves; the 150% ceiling is Typst's own default. A
    // paragraph carrying the East Asian/Latin auto space states its own
    // ceiling over this one, which is Word's rather than Typst's; see
    // `write_east_asian_justification_limits` (issue #1193).
    //
    // Every one of those ten exports declares `compatibilityMode 15`, so this
    // floor is calibrated on — and governs — Word's post-2013 justification.
    // The pre-2013 engine has no East Asian compression phase at all, and a
    // squeeze allowance is the wrong lever to model that with: the breaker
    // takes any allowance it is given. Such a paragraph gets Typst's first-fit
    // breaker instead, which never squeezes to seat a token; see
    // `justified_lines_take_natural_width_only` (issue #1130).
    //
    // Set for the document, next to the overhang rule above, so every
    // justified paragraph is covered whether it comes through the body, a
    // list or a table cell. It is inert wherever `justify` is false.
    let _ = writeln!(
        out,
        "#set par(justification-limits: (spacing: (min: {JUSTIFIED_SPACING_FLOOR}, max: 150%)))"
    );
    let _ = writeln!(
        out,
        "#let {PAGE_FORMAT_STATE} = state(\"{PAGE_FORMAT_STATE}\", \"1\")"
    );
    let _ = writeln!(
        out,
        "#show outline.entry: it => context {{ \
         let target = it.element.location(); \
         let shown = numbering({PAGE_FORMAT_STATE}.at(target), ..counter(page).at(target)); \
         link(target, it.indented(it.prefix(), \
         it.body() + box(width: 1fr, repeat[.]) + shown)) }}"
    );
}

/// Emit a `TOC` field's result.
///
/// Word recomputes the field when the document is opened, which is why a
/// generated document ships it empty; the entries, their page numbers, and the
/// dot leaders between them all come from where the headings actually land.
/// Typst's own outline resolves exactly that, against the `#heading` elements
/// the body already emits, so the entries stay correct as the layout moves
/// (issue #576).
///
/// The heading paragraph above the field carries the document's own title, so
/// the outline contributes none.
///
/// Both kinds of list number an entry through the format the entry's own
/// section declared: the heading outline through the `show` rule
/// [`write_page_format_state`] installs, the caption list through the same
/// lookup written into the row it builds here (issue #605).
fn generate_table_of_contents(out: &mut String, contents: &TableOfContents, ctx: &GenCtx) {
    match contents {
        TableOfContents::Headings { depth } => {
            // Word lays a contents entry out as body text indented by its
            // level, not as a copy of the heading. Typst's own outline builds
            // each entry from the heading's content, which carries the
            // heading's size and weight as inline markup, so the entries came
            // out large and bold (issue #610). Build the list from the plain
            // text each heading drops instead, and style it here.
            let _ = writeln!(
                out,
                "#context {{ {entry_style}for entry in query(<{TOC_ENTRY_LABEL}>) {{                  if entry.value.level <= {depth} {{                  let target = entry.location();                  let shown = numbering({PAGE_FORMAT_STATE}.at(target),                  ..counter(page).at(target));                  block(below: {}pt)[#h({}pt * (entry.value.level - 1))                 #text(font: entry.value.font)[#link(target, entry.value.text)]                  #box(width: 1fr, repeat[.]) #shown] }} }} }}",
                format_f64(TOC_ENTRY_SPACING_PT),
                format_f64(TOC_LEVEL_INDENT_PT),
                entry_style = toc_entry_text_settings(ctx.document_default_text.as_ref()),
            );
        }
        // A caption is not a heading, so Typst's outline cannot reach it. Each
        // one drops an invisible `#metadata` under a per-identifier label as it
        // is laid out; the list queries those, and asks the page counter where
        // each landed. Both halves of the answer therefore come from the
        // layout, as they do for the heading outline.
        TableOfContents::Captions { identifier } => {
            let label = caption_label(identifier);
            let _ = writeln!(
                out,
                "#context {{ {entry_style}for entry in query(<{label}>) {{ \
                 let target = entry.location(); \
                 let shown = numbering({PAGE_FORMAT_STATE}.at(target), \
                 ..counter(page).at(target)); \
                 block(below: {}pt)[#h({}pt)#entry.value                  #box(width: 1fr, repeat[.]) #shown] }} }}",
                format_f64(CAPTION_ENTRY_SPACING_PT),
                format_f64(CAPTION_LIST_INDENT_PT),
                entry_style = toc_entry_text_settings(ctx.document_default_text.as_ref()),
            );
        }
    }
}

/// The Typst label a `SEQ` identifier's captions carry.
///
/// A label is an identifier, so the `SEQ` name — which a document may write in
/// any script — is reduced to the characters a label allows. Distinct
/// identifiers that reduce to the same label would share a list, which is why
/// the prefix keeps them away from any label the document's own content might
/// produce.
fn caption_label(identifier: &str) -> String {
    let mut label = String::from("o2p-seq-");
    for character in identifier.chars() {
        if character.is_ascii_alphanumeric() {
            label.push(character);
        } else {
            let _ = write!(label, "-{:x}", character as u32);
        }
    }
    label
}

fn generate_block(out: &mut String, block: &Block, ctx: &mut GenCtx) -> Result<(), ConvertError> {
    match block {
        Block::Paragraph(para) => generate_paragraph(
            out,
            para,
            ctx.line_grid_pitch,
            ctx.default_tab_width_pt,
            ctx.breaks_hangul_at_eojeol,
            ctx.available_measure_pt,
        ),
        Block::TableOfContents(contents) => {
            generate_table_of_contents(out, contents, ctx);
            Ok(())
        }
        Block::Caption(caption) => {
            let _ = writeln!(
                out,
                "#metadata[{}]<{}>",
                escape_typst(&caption.entry_text),
                caption_label(&caption.identifier)
            );
            generate_paragraph(
                out,
                &caption.paragraph,
                ctx.line_grid_pitch,
                ctx.default_tab_width_pt,
                ctx.breaks_hangul_at_eojeol,
                ctx.available_measure_pt,
            )
        }
        Block::PageBreak => {
            out.push_str("#pagebreak()\n");
            Ok(())
        }
        Block::Table(table) => generate_table(out, table, ctx),
        Block::Image(img) => {
            // Word advances a picture paragraph by the picture plus its own
            // `w:spacing`. Leaving the element bare let Typst's 1.2em default
            // block spacing apply above and below instead, opening ~24pt
            // around an inline figure (issues #463, #491) - so both gaps stay
            // pinned, but to the declared spacing rather than to zero, which
            // dropped the gap Word actually draws (issue #499).
            write_image_block_open(out, img.paragraph_spacing);
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
            out.push_str("]\n");
            Ok(())
        }
        Block::InlineImages(images) => {
            // One paragraph, so one gap above and below the whole group: the
            // first picture carries the `before`, the last the `after`.
            let spacing: Option<ImageParagraphSpacing> = images
                .first()
                .and_then(|first| first.paragraph_spacing)
                .map(|first| ImageParagraphSpacing {
                    before: first.before,
                    after: images
                        .last()
                        .and_then(|last| last.paragraph_spacing)
                        .and_then(|last| last.after),
                });
            write_image_block_open(out, spacing);
            out.push('\n');
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
            generate_floating_shape(out, fs, ctx);
            Ok(())
        }
        Block::List(list) => {
            // Grid-snapped line height applies to list items too (Word's
            // document grid covers all body text).
            let first_paragraph = list.items.first().and_then(|item| item.content.first());
            let settings: Option<String> = first_paragraph.and_then(|paragraph| {
                word_line_height_settings(&paragraph.runs, &paragraph.style, ctx.line_grid_pitch)
            });
            // `generate_list` emits the wrapper itself, so the line box and
            // the list's own `w:spacing` gaps share one block (issue #463).
            let line_box_em: Option<(f64, f64)> = first_paragraph.and_then(|paragraph| {
                word_line_box_em(&paragraph.runs, &paragraph.style, ctx.line_grid_pitch)
            });
            generate_list(
                out,
                list,
                settings.as_deref(),
                ListEojeolWrap {
                    breaks_hangul_at_eojeol: ctx.breaks_hangul_at_eojeol,
                    line_box_em,
                    available_measure_pt: ctx.available_measure_pt,
                    baseline_snap: None,
                },
            )
        }
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

/// Collapse a style onto one of Typst's three named dash patterns.
///
/// This is the coarse fallback. DrawingML strokes go through
/// `drawingml_dash_array_pt`, which emits each preset's own rhythm; they reach
/// this only when the width is unusable, and then the nearest named pattern is
/// all that is left. Word and Excel borders use it as their normal path.
fn border_line_style_to_typst(style: BorderLineStyle) -> &'static str {
    match style {
        BorderLineStyle::Solid => "solid",
        BorderLineStyle::Dashed | BorderLineStyle::SystemDash | BorderLineStyle::LargeDash => {
            "dashed"
        }
        BorderLineStyle::Dotted | BorderLineStyle::SystemDot => "dotted",
        BorderLineStyle::DashDot
        | BorderLineStyle::DashDotDot
        | BorderLineStyle::SystemDashDot
        | BorderLineStyle::LargeDashDot
        | BorderLineStyle::SystemDashDotDot
        | BorderLineStyle::LargeDashDotDot => "dash-dotted",
        BorderLineStyle::Double => "solid",
        BorderLineStyle::None => "solid",
    }
}

/// Opens the block that wraps a flow picture, pinning both vertical gaps.
///
/// Both are always written. An absent `w:spacing` means Word draws no gap, not
/// that Typst should fall back to its 1.2em default, so `None` maps to `0pt`.
fn write_image_block_open(out: &mut String, spacing: Option<ImageParagraphSpacing>) {
    let above: f64 = spacing.and_then(|gap| gap.before).unwrap_or(0.0);
    let below: f64 = spacing.and_then(|gap| gap.after).unwrap_or(0.0);
    let _ = write!(
        out,
        "#block(width: 100%, above: {}pt, below: {}pt)[",
        format_f64(above),
        format_f64(below)
    );
}

fn generate_image(out: &mut String, img: &ImageData, ctx: &mut GenCtx) {
    // "Crop to shape": clip the image box to the picture's preset geometry.
    if let Some(clip) = img.clip_shape
        && let (Some(width), Some(height)) = (img.width, img.height)
    {
        let radius: String = match clip {
            crate::ir::ImageClipShape::Ellipse => "50%".to_string(),
            crate::ir::ImageClipShape::RoundedRect(fraction) => {
                format!("{}pt", format_f64(width.min(height) * fraction))
            }
        };
        let _ = write!(
            out,
            "#box(width: {}pt, height: {}pt, clip: true, radius: {radius})[",
            format_f64(width),
            format_f64(height)
        );
        let mut inner: ImageData = img.clone();
        inner.clip_shape = None;
        generate_image(out, &inner, ctx);
        out.pop();
        out.push_str("]\n");
        return;
    }

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
            generate_floating_image_content(out, &fi.image, &path);
            out.push_str("]\n");
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
            generate_floating_image_content(out, &fi.image, &path);
            out.push_str("]\n");
        }
        WrapMode::Square | WrapMode::Tight => {
            // Best-effort text wrapping: use #place with float: true
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, float: true)[",
                format_f64(fi.offset_x),
                format_f64(fi.offset_y)
            );
            generate_floating_image_content(out, &fi.image, &path);
            out.push_str("]\n");
        }
    }
}

/// Emit a floating picture while preserving its unrotated frame dimensions
/// and centre.
///
/// Word turns `a:xfrm/@rot` clockwise around the picture centre. The image is
/// already inside an absolute `#place`, so rotation changes only its painted
/// extent and does not affect document flow. As with oversized fixed elements,
/// Typst can clamp the body frame before resolving `origin: center`; pivot on
/// the unclamped top-left corner and translate it back to Word's centre instead
/// (issues #1032, #1366).
fn generate_floating_image_content(out: &mut String, image: &ImageData, path: &str) {
    let rotation = image.rotation_deg.filter(|degrees| *degrees != 0.0);
    let pivot_shift = rotation.and_then(|degrees| {
        image
            .width
            .zip(image.height)
            .map(|(width, height)| centre_pivot_shift(width, height, degrees, false, false))
    });
    if let Some(degrees) = rotation {
        if let Some((dx, dy)) = pivot_shift {
            let _ = write!(
                out,
                "#move(dx: {}pt, dy: {}pt)[#rotate({}deg, origin: top + left)[",
                format_f64(dx),
                format_f64(dy),
                format_f64(degrees)
            );
        } else {
            let _ = write!(out, "#rotate({}deg, origin: center)[", format_f64(degrees));
        }
    }

    out.push_str("#image(\"");
    out.push_str(path);
    out.push('"');
    if let Some(width) = image.width {
        let _ = write!(out, ", width: {}pt", format_f64(width));
    }
    if let Some(height) = image.height {
        let _ = write!(out, ", height: {}pt", format_f64(height));
    }
    out.push(')');

    if rotation.is_some() {
        out.push(']');
    }
    if pivot_shift.is_some() {
        out.push(']');
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
fn generate_floating_shape(out: &mut String, fs: &FloatingShape, ctx: &mut GenCtx) {
    out.push_str("#box(width: 0pt, height: 0pt)[\n");
    generate_floating_shape_overlay(out, fs, ctx);
    out.push_str("]\n");
}

fn generate_floating_shape_overlay(out: &mut String, fs: &FloatingShape, ctx: &mut GenCtx) {
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(fs.offset_x),
        format_f64(fs.offset_y)
    );
    shapes::generate_shape(out, &fs.shape, fs.width, fs.height, ctx);
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
    if let Some(rotation) = ftb.shape_rotation_deg.filter(|degrees| *degrees != 0.0) {
        let mut inner: FloatingTextBox = ftb.clone();
        inner.shape_rotation_deg = None;
        let (dx, dy): (f64, f64) = centre_pivot_shift(
            ftb.width.max(0.0),
            ftb.height.max(0.0),
            rotation,
            false,
            false,
        );
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[#rotate({}deg, origin: top + left, reflow: false)[",
            format_f64(dx),
            format_f64(dy),
            format_f64(rotation)
        );
        generate_floating_text_box_content(out, &inner, ctx)?;
        out.push_str("]]\n");
        return Ok(());
    }

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
            generate_fixed_text_box_block(out, block, ctx, Some(inner_width), false, false)?;
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
        generate_fixed_text_box_block(out, block, ctx, Some(inner_width), false, false)?;
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

    // Only a file that asked for it gets its text scaled. A box barely one
    // line tall used to qualify on its own, which overrode the declared size:
    // an 8pt label in a 9.6pt box came out at 4.9pt where the reference keeps
    // 8pt and lets the text overflow (issue #898). `auto_fit` represents the
    // dynamic fallback for `<a:normAutofit/>` without saved `fontScale` or
    // `lnSpcReduction`; the parser has already applied saved results.
    text_box.auto_fit.then_some(paragraph)
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

fn powerpoint_centered_single_line_text_box(
    text_box: &TextBoxData,
    inner_width_pt: f64,
) -> Option<TextBoxData> {
    if !matches!(text_box.vertical_align, TextBoxVerticalAlign::Center) {
        return None;
    }
    let [Block::Paragraph(paragraph)] = text_box.content.as_slice() else {
        return None;
    };
    if paragraph.runs.iter().any(|run| run.footnote.is_some()) {
        return None;
    }
    if powerpoint_hard_breaks_use_line_stack(
        &paragraph.runs,
        &paragraph.style,
        Some(inner_width_pt),
    ) {
        return None;
    }

    let mut has_hard_break: bool = false;
    let mut current_line_is_visible: bool = false;
    let mut visible_line_count: usize = 0;
    for character in paragraph.runs.iter().flat_map(|run| run.text.chars()) {
        if matches!(character, '\n' | '\r' | '\u{000B}') {
            has_hard_break = true;
            visible_line_count += usize::from(current_line_is_visible);
            current_line_is_visible = false;
        } else if !character.is_whitespace() {
            current_line_is_visible = true;
        }
    }
    visible_line_count += usize::from(current_line_is_visible);
    if !has_hard_break || visible_line_count != 1 {
        return None;
    }

    let mut normalized: TextBoxData = text_box.clone();
    let [Block::Paragraph(paragraph)] = normalized.content.as_mut_slice() else {
        unreachable!("the cloned text box has the same single paragraph")
    };
    for run in &mut paragraph.runs {
        run.text
            .retain(|character| !matches!(character, '\n' | '\r' | '\u{000B}'));
    }
    paragraph
        .runs
        .retain(|run| !run.text.is_empty() || run.footnote.is_some());
    Some(normalized)
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

fn common_powerpoint_list_baseline_snap(list: &List) -> Option<PowerPointBaselineSnap> {
    let root_level: u32 = list.items.first()?.level;
    let mut snaps = list.items.iter().map(|item| {
        let [paragraph] = item.content.as_slice() else {
            return None;
        };
        (item.level == root_level)
            .then(|| powerpoint_absolute_baseline_snap(&paragraph.runs, &paragraph.style))
            .flatten()
    });
    let first: PowerPointBaselineSnap = snaps.next()??;
    snaps.all(|snap| snap == Some(first)).then_some(first)
}

fn generate_fixed_text_box_block(
    out: &mut String,
    block: &Block,
    ctx: &mut GenCtx,
    available_width_pt: Option<f64>,
    no_wrap: bool,
    snap_absolute_baselines: bool,
) -> Result<(), ConvertError> {
    match block {
        Block::List(list) if can_render_fixed_text_list_inline(list) => generate_fixed_text_list(
            out,
            list,
            true,
            available_width_pt,
            true,
            snap_absolute_baselines,
        ),
        Block::Paragraph(para) => generate_fixed_text_paragraph(
            out,
            para,
            available_width_pt,
            no_wrap,
            snap_absolute_baselines,
        ),
        // A slide's bullet list paces on PowerPoint's line, not Word's. Routing
        // it through `generate_block` gave it the font's hhea pitch, which is up
        // to 4% short per line and accumulates down the list (issue #513).
        Block::List(list) => {
            let settings: Option<String> = list
                .items
                .first()
                .and_then(|item| item.content.first())
                .and_then(|paragraph| {
                    powerpoint_line_height_settings(&paragraph.runs, &paragraph.style)
                });
            // A slide's own breaking; PowerPoint splits Korean mid-word.
            let baseline_snap = snap_absolute_baselines
                .then(|| common_powerpoint_list_baseline_snap(list))
                .flatten();
            generate_list_with_spacing_model(
                out,
                list,
                settings.as_deref(),
                true,
                ListEojeolWrap {
                    baseline_snap,
                    ..ListEojeolWrap::default()
                },
            )
        }
        _ => generate_block(out, block, ctx),
    }
}

fn generate_fixed_text_paragraph(
    out: &mut String,
    para: &Paragraph,
    available_width_pt: Option<f64>,
    no_wrap: bool,
    snap_absolute_baseline: bool,
) -> Result<(), ConvertError> {
    let style: &ParagraphStyle = &para.style;
    let inset: Insets = fixed_text_paragraph_inset(style);
    let has_inset: bool = inset.left > 0.0 || inset.right > 0.0;
    let hanging_indent_pt: Option<f64> = fixed_text_paragraph_hanging_indent_pt(style);
    let first_line_indent_pt: Option<f64> =
        style.indent_first_line.filter(|value| value.abs() > 0.0001);
    // PowerPoint's own line, which supersedes the `size * 0.65` leading this
    // path used to guess with (issue #513).
    let line_height_settings: Option<String> = powerpoint_line_height_settings(&para.runs, style);
    let baseline_snap: Option<PowerPointBaselineSnap> = snap_absolute_baseline
        .then(|| powerpoint_absolute_baseline_snap(&para.runs, style))
        .flatten();
    let needs_text_scope: bool = common_text_style(&para.runs).is_some();
    let has_para_style: bool = needs_block_wrapper(style)
        || needs_text_scope
        || line_height_settings.is_some()
        || has_inset;

    if has_para_style {
        out.push_str("#block(");
        // A slide paragraph's gaps are its own `a:spcBef`/`a:spcAft` and
        // nothing else. Leaving them unset let Typst's 1.2em `block.spacing`
        // default in, which put 13pt between the lines of a code block that
        // declares no spacing at all (issue #513).
        let _ = write!(
            out,
            "above: {}pt, below: {}pt",
            format_f64(style.space_before.unwrap_or(0.0)),
            format_f64(style.space_after.unwrap_or(0.0)),
        );
        if has_inset {
            let _ = write!(out, ", width: 100%, inset: {}", format_insets(&inset));
        }
        out.push_str(")[\n");
        match line_height_settings {
            // The line box carries the whole advance and pins leading to zero,
            // so the leading `write_par_settings` derives from the same spacing
            // would only be a contradictory rule the box then overrides.
            Some(ref settings) => {
                write_par_settings(
                    out,
                    &ParagraphStyle {
                        line_spacing: None,
                        ..style.clone()
                    },
                    &para.runs,
                );
                write_common_text_settings(out, &para.runs, "  ");
                out.push_str(settings);
            }
            None => {
                write_par_settings(out, style, &para.runs);
                write_common_text_settings(out, &para.runs, "  ");
                write_fixed_text_default_par_settings(out, style, &para.runs, "  ");
            }
        }
    }

    if let Some(snap) = baseline_snap {
        snap.write_open(out);
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

    if let Some(hanging_indent_pt) = hanging_indent_pt {
        let _ = write!(
            out,
            "#par(hanging-indent: {}pt)[",
            format_f64(hanging_indent_pt)
        );
    } else if let Some(first_line_indent_pt) = first_line_indent_pt {
        let _ = write!(
            out,
            "#par(first-line-indent: (amount: {}pt, all: true))[",
            format_f64(first_line_indent_pt)
        );
    }

    if no_wrap {
        out.push_str("#box[");
        generate_runs_with_tabs_no_wrap(
            out,
            &para.runs,
            style.tab_stops.as_deref(),
            paragraph_default_tab_width_pt(style, DEFAULT_TAB_WIDTH_PT),
        );
    } else {
        generate_powerpoint_runs_with_tabs(
            out,
            &para.runs,
            style,
            style.tab_stops.as_deref(),
            paragraph_default_tab_width_pt(style, DEFAULT_TAB_WIDTH_PT),
            // PowerPoint splits Korean mid-word, so a slide's text box already
            // sits on the engine default.
            //
            // A Word *floating* text box also lands here, and Word does break
            // its Hangul at eojeol — but this path resolves neither input a
            // frame needs: the fixed text edges it would have to restore, and
            // the box's inner measure that bounds how wide a framed token may
            // be. Framing without them would shift baselines and overflow
            // narrow boxes, so a DOCX floating text box is a known gap in
            // #626 rather than a silently wrong emission. `GenCtx`'s
            // `breaks_hangul_at_eojeol` documents the same exclusion.
            EojeolWrap::Syllable,
            available_width_pt,
            true,
        );
    }
    // The line is placed by the width PowerPoint measured, which carries a
    // letter-space after its last glyph that Typst's shaping drops (#1075).
    // It goes inside the `no_wrap` box so that box's own width carries it too.
    //
    // This is the paragraph's *last* line only. Every line before a hard break
    // needs the same reserve, and `generate_powerpoint_runs_with_tabs` writes
    // those inside the paragraph markup, where each `#linebreak()` is (#1174).
    if let Some(spacing) = powerpoint_trailing_letter_space_pt(style, &para.runs) {
        let _ = write!(out, "#h({}pt)", format_f64(spacing));
    }
    if no_wrap {
        out.push(']');
    }

    if hanging_indent_pt.is_some() || first_line_indent_pt.is_some() {
        out.push(']');
    }

    if use_align {
        out.push(']');
    }

    if baseline_snap.is_some() {
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
