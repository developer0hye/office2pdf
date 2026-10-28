use std::fmt::Write;
use std::io::Cursor;

use image::{GenericImageView, ImageFormat as RasterImageFormat};

use crate::config::ConvertOptions;
use crate::error::ConvertError;
use crate::ir::{
    Alignment, ArrowHead, AxisTickMark, BarBandLayout, BaselineShiftEm, Block, BorderLineStyle,
    BorderSide, BreakClear, CellBorder, CellVerticalAlign, Chart, ChartGrouping, ChartType, Color,
    ColumnLayout, Document, FixedElement, FixedElementKind, FixedPage, FloatingImage,
    FloatingImageVerticalAnchor, FloatingShape, FloatingTable, FloatingTextBox, FlowPage,
    FrameAnchor, GradientFill, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterShapeContent,
    HorizontalRule, HorizontalRulePaint, HorizontalRuleWidth, IconShading, ImageCrop, ImageData,
    ImageFormat, ImageParagraphIndent, ImageParagraphSpacing, InlineTextBox, Insets,
    LegendPosition, LineBox, LineCap, LineJoin, LineSpacing, List, ListKind, Margins, MathEquation,
    Metadata, PAGE_NUMBER_FIELD_CHAR, Page, PageNumberFormat, PageSize, PairKerning, Paragraph,
    ParagraphStyle, PatternFill, PatternPreset, PositionedTabAlignment, PositionedTabRelativeTo,
    Run, Shadow, Shape, ShapeKind, SheetPage, SmartArt, SparklineInfo, TOTAL_PAGES_FIELD_CHAR,
    TabAlignment, TabLeader, TabStop, Table, TableBorderPaintModel, TableCell, TableOfContents,
    TableRow, TextBoxData, TextBoxVerticalAlign, TextDirection, TextFrame, TextStyle,
    VerticalTextAlign, WrapMode,
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
pub(crate) use self::diagrams::chart_uses_data_table;
pub(crate) use self::tables::header_row_count_covering_rowspans;
pub(crate) use self::text::is_cjk_like;
pub(crate) use self::text::{
    COMPACTED_SHEET_CELL_MIN_DESCENT_SEAT_PT, SHEET_CELL_MIN_DESCENT_SEAT_PT,
};
pub(crate) use self::text::{
    EAST_ASIAN_AUTO_SPACE_EM, EAST_ASIAN_AUTO_SPACE_GLYPH, EAST_ASIAN_JUSTIFIED_GAP_CEILING_EM,
    EAST_ASIAN_JUSTIFIED_GAP_RATIO_TERM_PERCENT,
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
/// Word's 0.96pt measured rule rounds to Typst's 1pt stroke for this separator.
const DOCX_COLUMN_SEPARATOR_STROKE_PT: f64 = 1.0;
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

/// The label of the `#metadata` that records a section's restarted page
/// number on the page where its counter restarts.
///
/// Typst lays a page's background out before the page's body, so a page
/// counter drawn there misses a `counter(page).update` on that same page; the
/// header and footer, laid out after the body, see it. A page field in a
/// header or footer table is drawn in the background, so it reads the restart
/// from this record instead (issue #2338).
const PAGE_RESTART_LABEL: &str = "o2p-page-restart";

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
    /// The continuous `<w:titlePg/>` sections that share the page setup the
    /// current section writes ([`title_page_run_members`]).
    flow_run_title_page_members: Vec<TitlePageRunMember>,
    /// Layers containing the previous section's borders must be reset explicitly.
    flow_border_layers: [bool; 2],
    has_flow_line_numbering: bool,
    in_flow_body: bool,
    /// The section body is written straight into the page flow, with no
    /// single-frame wrapper around it, so a table row there can continue on
    /// the next page (issue #2217).
    flow_body_is_breakable: bool,
    flow_page_alignments: Vec<crate::ir::PageVerticalAlignment>,
    flow_alignment_index: usize,
    flow_current_alignment: crate::ir::PageVerticalAlignment,
    preserve_page_after_spacing: bool,
    flow_previous_after_spacing: f64,
    block_depth: usize,
    images: Vec<ImageAsset>,
    next_image_id: usize,
    next_text_box_id: usize,
    /// Whether generated table cells belong to their enclosing top-aligned floating text box.
    in_top_aligned_floating_text_box: bool,
    /// Header/footer story tables preserve their sibling-derived blank
    /// paragraph metrics so paragraph-anchored artwork keeps its native seat.
    prefer_sibling_metrics_for_empty_cell_paragraphs: bool,
    /// Whether the header story being written lays each line out as Word's
    /// own line box, which is the only form that honours an exact line.
    hf_lines_are_word_boxes: bool,
    /// How many sheet drawing layers have been written. Each needs its own
    /// label, because the layer is a page foreground that has to recognise
    /// its own sheet's first printed page (issue #1168).
    next_sheet_drawing_layer_id: usize,
    next_excel_fill_id: usize,
    table_depth: usize,
    /// The table being written is a block of a breakable section body, so
    /// its rows may continue on the next page (issue #2217).
    table_rows_may_split: bool,
    /// Active section's Word document-grid line pitch, in points.
    line_grid_pitch: Option<f64>,
    /// `w:defaultTabStop` from the document settings, in points.
    document_default_tab_stop_pt: Option<f64>,
    /// Effective default tab stop interval, in points, for the active page.
    default_tab_width_pt: f64,
    /// True until the document's first page has been generated. Word keeps
    /// `w:spacing w:before` on the very first body paragraph, but suppresses it
    /// for a separate paragraph after a break; continuation fragments retain it.
    at_document_start: bool,
    /// The table row's East Asian line metrics, decided once per row so
    /// differences in cell text cannot shift the answer (issue #498).
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
    /// Whether the cell being generated seats its line box on the descender:
    /// the enclosing table is a spreadsheet and the cell's effective vertical
    /// alignment is bottom (issue #618).
    cell_seats_text_on_descender: bool,
    /// The current cell's effective vertical alignment. Word's compressed
    /// line box needs the resolved table default as well as an explicit cell
    /// value to reproduce the anchor-specific baseline seat (issue #1479).
    cell_vertical_align: Option<CellVerticalAlign>,
    /// Shared metrics for tight-row centered cells and horizontal bottom
    /// merges (#839). Unmerged top/bottom cells keep their own seats (#1721).
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
    /// Page geometry currently active in Typst. A continuous section can
    /// change its body frame without changing the page that already began.
    active_page_size: Option<PageSize>,
    active_page_margins: Option<Margins>,
    /// The top margin the active page setup writes, where a line that opens
    /// a page lands. A header taller than the section's margin grows it.
    active_body_top_pt: Option<f64>,
    /// Whether the flow section before the one being written ends at a page
    /// or column break, so the new section opens at a page or column top.
    previous_flow_ends_at_layout_break: bool,
    /// How many paragraphs have had a page-top gap label written, which
    /// keeps each label unique.
    next_page_top_gap_id: usize,
    /// The spreadsheet cell box the current paragraph is laid out in, so a
    /// centred line can be seated on Excel's whole-point grid (issue #1600).
    /// `None` outside a sheet table cell.
    sheet_cell_box: Option<SheetCellBox>,
    /// Whether the table being generated is a slide's, and so paces its cell
    /// text on PowerPoint's line box rather than Word's hhea one. Plain slide
    /// table lines use 1.2em; explicit proportional lines use the table's
    /// measured descent-gap rule (issue #1914).
    ///
    /// A slide's own text boxes already route through
    /// [`powerpoint_line_height_settings`], but a `<a:tbl>` reaches the shared
    /// table codegen, which gave its cells Word's model: 1.587em measured on
    /// `office2pdf_introduction_ko` slide 16, so multi-line cells grew and the
    /// table's bottom border moved down with them (issue #663).
    ///
    /// The plain-line target is PowerPoint's documented 1.2em, the same factor
    /// the slide's text boxes already use. No native PowerPoint export of that
    /// fixture is committed; a **LibreOffice** render of it advances 1.235em,
    /// which corroborates the direction and magnitude without being ground
    /// truth.
    table_uses_powerpoint_line_box: bool,
    /// The floating tables and text frames placed so far in the current
    /// flow, which the root paragraphs after them wrap around until a page,
    /// column or section break, or until those paragraphs have advanced a
    /// page (issues #2229, #2231). A float outlasts the start of a
    /// continuous section laid out in the same text area (issue #2369).
    body_floats: Vec<float_wrap::BodyFloat>,
    next_body_float_id: usize,
    /// The width and height of the single text column floating tables and
    /// text frames wrap in, or `None` where they do not wrap yet: outside a
    /// flow body and in multi-column sections.
    float_wrap_text_area_pt: Option<(f64, f64)>,
    /// Set for the paragraph `generate_block` is about to write at the flow
    /// root while `body_floats` is non-empty; `generate_paragraph` takes it,
    /// so the paragraphs nested inside that one never see it.
    float_wrap_pending: bool,
    /// Set with `float_wrap_pending` for the part of a paragraph after a
    /// clearing break, which starts below the floats it clears
    /// (issue #2230).
    float_wrap_clear: Option<BreakClear>,
    /// Set for the root text-anchored text frame or floating table
    /// `generate_block` is about to write: the gap after the flow block before
    /// it, when there is one (issues #2231, #2236). A top-margin frame box
    /// hangs from no flow block, so it never has one.
    flow_anchor_gap_pt: Option<f64>,
    /// The text area the previous flow section's floats wrapped in, which a
    /// continuous section in the same area keeps in force.
    carried_float_text_area_pt: Option<(f64, f64)>,
    /// Whether the current flow section keeps an empty frame's box at the top
    /// margin, which the section's full-width tables at the flow root clear
    /// (issue #2231). docxide-pdf's `empty_frame_table_bands` keeps the boxes
    /// to their own section too, and the check is a layout-time query that
    /// every such table would otherwise pay for.
    clears_top_margin_frame_boxes: bool,
    /// Whether the next block in the block list being written is a top-margin
    /// frame box, so a full-width table before it clears the box with the
    /// rows it carries onto the box's page (issue #2231).
    top_margin_frame_box_follows: bool,
    /// Set by such a table for the box after it: the bottom of the text area,
    /// from the top of the page.
    frame_box_carry_text_bottom_pt: Option<f64>,
}

impl GenCtx {
    /// The scale a spreadsheet cell's declared-space metrics must be read
    /// through: `Some(1.0)` on an unscaled sheet, `Some(the fit-to-page
    /// factor)` on a fitted one, and `None` off a sheet. This covers the
    /// measured wrapped-line advance (#1163), whole-point seats (#1238), and
    /// the declared-width border bands Excel scales with the page (#1564).
    fn sheet_print_scale(&self) -> Option<f64> {
        self.table_seats_bottom_aligned_text_on_descender
            .then(|| self.table_print_scale.unwrap_or(1.0))
    }

    fn new() -> Self {
        Self {
            flow_section_index: 0,
            flow_run_title_page_members: Vec::new(),
            flow_border_layers: [false; 2],
            has_flow_line_numbering: false,
            in_flow_body: false,
            flow_body_is_breakable: false,
            flow_page_alignments: Vec::new(),
            flow_alignment_index: 0,
            flow_current_alignment: crate::ir::PageVerticalAlignment::Top,
            preserve_page_after_spacing: false,
            flow_previous_after_spacing: 0.0,
            block_depth: 0,
            images: Vec::new(),
            next_image_id: 0,
            next_text_box_id: 0,
            in_top_aligned_floating_text_box: false,
            prefer_sibling_metrics_for_empty_cell_paragraphs: false,
            hf_lines_are_word_boxes: false,
            next_sheet_drawing_layer_id: 0,
            next_excel_fill_id: 0,
            table_depth: 0,
            table_rows_may_split: false,
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
            active_page_size: None,
            active_page_margins: None,
            active_body_top_pt: None,
            previous_flow_ends_at_layout_break: false,
            next_page_top_gap_id: 0,
            sheet_cell_box: None,
            body_floats: Vec::new(),
            next_body_float_id: 0,
            float_wrap_text_area_pt: None,
            float_wrap_pending: false,
            float_wrap_clear: None,
            flow_anchor_gap_pt: None,
            carried_float_text_area_pt: None,
            clears_top_margin_frame_boxes: false,
            top_margin_frame_box_follows: false,
            frame_box_carry_text_bottom_pt: None,
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

    fn add_generated_png(&mut self, data: Vec<u8>) -> String {
        let id = self.next_image_id;
        self.next_image_id += 1;
        let path = format!("img-{id}.png");
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
    write_document_language(&mut out, doc.styles.default_language.as_ref());
    write_hyphenation_rule(&mut out, doc.styles.hyphenates_automatically);
    if let Some(ligatures) = doc
        .styles
        .default_text
        .as_ref()
        .and_then(|text| text.ligatures)
    {
        write_document_ligature_state(&mut out, ligatures);
    }
    if doc.pages.iter().any(|page| matches!(page, Page::Fixed(_))) {
        write_powerpoint_ligature_state(&mut out);
        write_powerpoint_advance_grid_helpers(&mut out);
        out.push('\n');
    }

    if float_wrap::document_has_body_float(doc) {
        float_wrap::write_float_wrap_helper(&mut out);
    }
    if float_wrap::document_has_top_margin_frame_box(doc) {
        float_wrap::write_top_margin_frame_box_clearance_helper(&mut out);
    }

    let mut ctx = GenCtx::new();
    ctx.document_default_tab_stop_pt = doc.styles.default_tab_stop_pt;
    ctx.document_default_text = doc.styles.default_text.clone();
    let alignment_plan = section_alignment::alignment_plan(&doc.pages);
    for (index, page) in doc.pages.iter().enumerate() {
        let begins_continuously: bool = matches!(page, Page::FlowContinuous(_));
        let ends_continuously: bool = doc
            .pages
            .get(index + 1)
            .is_some_and(|next_page| matches!(next_page, Page::FlowContinuous(_)));
        if index > 0 && !begins_continuously {
            out.push_str("\n#pagebreak()\n");
        }
        match page {
            Page::Flow(flow) | Page::FlowContinuous(flow) => {
                ctx.flow_run_title_page_members =
                    title_page_run_members(&doc.pages, index, ctx.flow_section_index);
                ctx.previous_flow_ends_at_layout_break = index
                    .checked_sub(1)
                    .and_then(|previous| doc.pages.get(previous))
                    .is_some_and(|previous| match previous {
                        Page::Flow(previous) | Page::FlowContinuous(previous) => {
                            previous.content.last().is_some_and(is_layout_break)
                        }
                        _ => false,
                    });
                generate_flow_page(
                    &mut out,
                    flow,
                    begins_continuously && index > 0,
                    ends_continuously,
                    &alignment_plan[index],
                    &mut ctx,
                    options,
                )?;
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
    begins_continuously: bool,
    ends_continuously: bool,
    vertical_alignments: &[crate::ir::PageVerticalAlignment],
    ctx: &mut GenCtx,
    options: &ConvertOptions,
) -> Result<(), ConvertError> {
    let size = resolve_page_size(&page.size, options);
    ctx.in_flow_body = false;
    ctx.flow_page_alignments = vertical_alignments.to_vec();
    ctx.flow_alignment_index = 0;
    ctx.preserve_page_after_spacing = false;
    if !begins_continuously {
        ctx.flow_previous_after_spacing = 0.0;
    }
    write_line_numbering(out, page, ctx);
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
    // Absent w:defaultTabStop: East Asian Word editions (signalled by the
    // section's w:docGrid) default to 800 twips = 40pt where Western
    // editions use the ECMA 720 twips = 36pt (issue #393). Set before the
    // page setup, whose header and footer tabs use the same default stops
    // (issue #1821).
    ctx.default_tab_width_pt =
        ctx.document_default_tab_stop_pt
            .unwrap_or(if page.line_grid_pitch.is_some() {
                EAST_ASIAN_DEFAULT_TAB_WIDTH_PT
            } else {
                DEFAULT_TAB_WIDTH_PT
            });
    let mut section_content: Vec<Block> = page.content.clone();
    let mut section_mark_space: Option<(usize, f64)> = None;
    if begins_continuously {
        suppress_typst_spacing_at_continuous_section_start(&mut section_content);
        section_mark_space = take_section_mark_space(&mut section_content);
        if page.columns.is_none() && !ctx.previous_flow_ends_at_layout_break {
            keep_opening_space_before_at_page_top(&mut section_content);
        }
    }
    if !begins_continuously {
        write_flow_page_setup(out, page, &size, false, ctx)?;
        out.push('\n');
    }
    // The marker sits at the section's first page, so a first-page header can
    // resolve which page that is without assuming the section starts the
    // document (issue #846).
    if page.first_header.is_some()
        || page.first_footer.is_some()
        || page
            .page_borders
            .as_ref()
            .is_some_and(|border| border.display != crate::ir::PageBorderDisplay::AllPages)
    {
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
        let _ = writeln!(out, "#metadata({start}) <{PAGE_RESTART_LABEL}>");
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

    // Apply body alignment after page decorations have captured their styles.
    // Typst distributes each page's flow, preserving page breaks and horizontal alignment.
    ctx.in_flow_body = true;
    write_flow_body_alignment(out, ctx);
    let text_height_pt: f64 = size.height - page.margins.top - page.margins.bottom;
    // TODO(floats wrap in one column only): a floating table or text frame
    // placed in a multi-column section still excludes no text.
    let text_area_pt: Option<(f64, f64)> = ctx
        .available_measure_pt
        .filter(|_| page.columns.is_none() && text_height_pt > 0.0)
        .map(|width_pt| (width_pt, text_height_pt));
    // A float still excludes text in a continuous section laid out in the
    // same text area. In Word's exports of docxide-pdf's
    // `paragraph-offset-frame-immediate-long` (a text frame, issue #2231) and
    // `section-mark-after-table-float-compat15` (a floating table, issue
    // #2369), the float ends the first section, and the text it keeps off
    // opens the next.
    if !begins_continuously || ctx.carried_float_text_area_pt != text_area_pt {
        ctx.body_floats.clear();
    }
    ctx.float_wrap_text_area_pt = text_area_pt;
    ctx.clears_top_margin_frame_boxes = float_wrap::has_top_margin_frame_box(&section_content);

    // Word keeps `w:spacing w:before` on the document's first body paragraph,
    // but Typst collapses leading block spacing at a page boundary, pulling the
    // first heading up to the top margin. Emit that gap as explicit vertical
    // space instead, and drop the block's own `above` so it is not counted
    // twice. Word also suppresses spacing for a separate paragraph after a
    // break; continuation fragments preserve the source paragraph's before-gap.
    let leading_gap: Option<f64> = if ctx.at_document_start && page.columns.is_none() {
        match section_content.first() {
            Some(Block::Paragraph(paragraph)) => {
                paragraph.style.space_before.filter(|gap| *gap > 0.0)
            }
            _ => None,
        }
    } else {
        None
    };

    if let Some(gap) = leading_gap {
        let Some(Block::Paragraph(first)) = section_content.first() else {
            unreachable!("leading_gap is only set for a leading paragraph")
        };
        let mut adjusted = first.clone();
        adjusted.style.space_before = None;
        let _ = writeln!(out, "#v({}pt, weak: false)", format_f64(gap));
        ctx.flow_body_is_breakable = true;
        generate_block(out, &Block::Paragraph(adjusted), ctx)?;
        if section_content.len() > 1 {
            out.push('\n');
            generate_blocks(out, &section_content[1..], ctx)?;
        }
    } else if let Some(ref cols) = page.columns {
        generate_flow_page_columns(
            out,
            page,
            &size,
            &section_content,
            cols,
            FlowColumnRenderOptions {
                begins_continuously,
                ends_continuously,
                balance_final_segment: ends_continuously,
                section_mark_space,
            },
            ctx,
        )?;
    } else if begins_continuously
        && section_content
            .iter()
            .any(|block| matches!(block, Block::PageBreak))
    {
        generate_continuous_flow_blocks(
            out,
            page,
            &size,
            &section_content,
            ends_continuously,
            section_mark_space,
            ctx,
        )?;
    } else {
        let has_continuous_frame: bool =
            begins_continuously && write_continuous_flow_frame_open(out, page, ctx);
        // The continuous-section frame is a `#move`, which lays its body out
        // as one frame that cannot continue on the next page.
        ctx.flow_body_is_breakable = !has_continuous_frame;
        generate_blocks_with_layout_space_before(
            out,
            &section_content,
            ctx,
            section_mark_space,
            false,
            false,
            false,
        )?;
        if has_continuous_frame {
            write_continuous_flow_frame_close(out);
        }
    }
    ctx.flow_body_is_breakable = false;
    ctx.in_flow_body = false;
    ctx.preserve_page_after_spacing = false;
    ctx.carried_float_text_area_pt = ctx.float_wrap_text_area_pt.take();
    ctx.clears_top_margin_frame_boxes = false;
    reset_flow_body_alignment(out, ctx);
    Ok(())
}

fn reset_flow_body_alignment(out: &mut String, ctx: &mut GenCtx) {
    if ctx.flow_current_alignment != crate::ir::PageVerticalAlignment::Top {
        out.push_str("#set align(top)\n");
        ctx.flow_current_alignment = crate::ir::PageVerticalAlignment::Top;
    }
}

fn write_flow_body_alignment(out: &mut String, ctx: &mut GenCtx) {
    let vertical_alignment = ctx
        .flow_page_alignments
        .get(ctx.flow_alignment_index)
        .copied()
        .unwrap_or(crate::ir::PageVerticalAlignment::Top);
    let alignment = match vertical_alignment {
        crate::ir::PageVerticalAlignment::Center => "horizon",
        crate::ir::PageVerticalAlignment::Bottom => "bottom",
        _ => "top",
    };
    if ctx.flow_current_alignment != vertical_alignment {
        let _ = writeln!(out, "#set align({alignment})");
        ctx.flow_current_alignment = vertical_alignment;
    }
    ctx.preserve_page_after_spacing = matches!(
        vertical_alignment,
        crate::ir::PageVerticalAlignment::Center | crate::ir::PageVerticalAlignment::Bottom
    );
}

fn advance_flow_body_alignment(out: &mut String, ctx: &mut GenCtx) {
    ctx.flow_alignment_index += 1;
    ctx.flow_previous_after_spacing = 0.0;
    write_flow_body_alignment(out, ctx);
}

fn write_continuation_flow_setup(
    out: &mut String,
    page: &FlowPage,
    size: &PageSize,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    ctx.in_flow_body = false;
    ctx.preserve_page_after_spacing = false;
    reset_flow_body_alignment(out, ctx);
    write_flow_page_setup(out, page, size, true, ctx)?;
    ctx.in_flow_body = true;
    Ok(())
}

/// Remove Typst's fallback paragraph gap at the first body block of a
/// continuous section. The boundary is a Word section transition, not a
/// paragraph boundary, so implicit block spacing must not separate the flows.
/// A stated `w:before` stays: the section continues on the same page, which is
/// no page top, and Word draws the gap there (issue #2161).
fn suppress_typst_spacing_at_continuous_section_start(content: &mut [Block]) {
    let Some(first_content_index) = content.iter().position(block_has_paragraph_content) else {
        return;
    };
    if content[..first_content_index].iter().any(is_layout_break)
        || block_starts_after_layout_break(&content[first_content_index])
    {
        return;
    }
    let Some(style) = first_paragraph_style_mut(&mut content[first_content_index]) else {
        return;
    };
    if !style.space_before.is_some_and(|gap| gap > 0.0) {
        style.space_before = Some(0.0);
    }
}

/// Mark a continuous section's first block, a paragraph with text, to keep
/// its stated `w:before` at the top of the next page should it not fit on
/// this one: Word's export of docxide-pdf's `cont-first-text-overflow` keeps
/// the gap there (issue #2255), where Typst drops a block's `above` at a page
/// top. Run after [`take_section_mark_space`], which writes its gap as fixed
/// space and zeroes the paragraph's own. The caller skips a section that
/// opens after a page or column break, where Word suppresses the spacing of a
/// paragraph that follows a break.
///
/// TODO(unprobed page-top gaps): a multi-column section's opener, the
/// section-mark gap and a paragraph after the section's first block lose
/// their gap at a page top as well; no Word reference shows what they keep.
fn keep_opening_space_before_at_page_top(content: &mut [Block]) {
    let Some(first) = content
        .first_mut()
        .filter(|block| block_has_paragraph_content(block))
    else {
        return;
    };
    let style: &mut ParagraphStyle = match first {
        Block::Paragraph(paragraph) => &mut paragraph.style,
        Block::Caption(caption) => &mut caption.paragraph.style,
        _ => return,
    };
    style.keeps_space_before_at_page_top =
        !style.starts_after_layout_break && style.space_before.is_some_and(|gap| gap > 0.0);
}

/// The style of a block's first paragraph; a list's is its first item's.
fn first_paragraph_style_mut(block: &mut Block) -> Option<&mut ParagraphStyle> {
    match block {
        Block::Paragraph(paragraph) => Some(&mut paragraph.style),
        Block::Caption(caption) => Some(&mut caption.paragraph.style),
        Block::List(list) => list
            .items
            .first_mut()?
            .content
            .first_mut()
            .map(|paragraph| &mut paragraph.style),
        _ => None,
    }
}

/// The gap a continuous section's first paragraph takes after an empty
/// section mark that took no line ([`ParagraphStyle::section_mark_after`]),
/// with the index of the block it goes before. Word adds the part of the
/// paragraph's `w:before` beyond the mark's `w:after` to the previous
/// paragraph's whole after-gap (issue #2232), where Typst's block spacing
/// would keep only the larger of the two. So the gap is written as fixed
/// space, which Typst's weak block spacing looks past, and the paragraph's
/// own before-gap becomes zero. In a multi-column section the space goes in
/// the first column, where the paragraph is: Word's export of docxide-pdf's
/// `scraped/covid_insomnia_mental_health` starts its second column above the
/// first column's numbered heading. After a layout break the paragraph sits
/// at a page or column top, so it keeps its gaps as they were.
fn take_section_mark_space(content: &mut [Block]) -> Option<(usize, f64)> {
    let index: usize = content.iter().position(|block| {
        matches!(
            block,
            Block::Paragraph(_) | Block::Caption(_) | Block::List(_)
        )
    })?;
    if content[..index].iter().any(is_layout_break)
        || block_starts_after_layout_break(&content[index])
    {
        return None;
    }
    let style: &mut ParagraphStyle = first_paragraph_style_mut(&mut content[index])?;
    let mark_after: f64 = style.section_mark_after?;
    let space: f64 = (style.space_before.unwrap_or(0.0) - mark_after).max(0.0);
    style.space_before = Some(0.0);
    (space > 0.0).then_some((index, space))
}

/// Open the new section's horizontal text frame inside the page that is
/// already using the previous section's page margins.
fn write_continuous_flow_frame_open(out: &mut String, page: &FlowPage, ctx: &GenCtx) -> bool {
    let Some(active_margins) = ctx.active_page_margins else {
        return false;
    };
    let active_page_width: f64 = ctx
        .active_page_size
        .map(|size| size.width)
        .unwrap_or(page.size.width);
    let active_content_width: f64 = active_page_width - active_margins.left - active_margins.right;
    let section_content_width: f64 = page.size.width - page.margins.left - page.margins.right;
    let horizontal_offset: f64 = page.margins.left - active_margins.left;
    if horizontal_offset.abs() < 0.0001
        && (active_content_width - section_content_width).abs() < 0.0001
    {
        return false;
    }
    if section_content_width <= 0.0 {
        return false;
    }

    let _ = writeln!(
        out,
        "#move(dx: {}pt)[\n#block(width: {}pt)[",
        format_f64(horizontal_offset),
        format_f64(section_content_width)
    );
    true
}

fn write_continuous_flow_frame_close(out: &mut String) {
    out.push_str("]\n]\n");
}

/// Keep page breaks at the top level while applying continuous-section page
/// setup to each continuation page.
///
/// A break that ends the section leaves an empty last slice. The next section
/// starts a page of its own unless it continues on this one
/// (`ends_continuously`), so only then is that break written (issue #2370).
fn generate_continuous_flow_blocks(
    out: &mut String,
    page: &FlowPage,
    size: &PageSize,
    content: &[Block],
    ends_continuously: bool,
    section_mark_space: Option<(usize, f64)>,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let segments: Vec<(usize, &[Block])> = split_at_page_breaks(content);
    let last_segment_index: usize = segments.len() - 1;
    for (index, (segment_start, segment)) in segments.iter().enumerate() {
        let is_trailing_empty_segment: bool =
            index == last_segment_index && segment.is_empty() && !ends_continuously;
        if index > 0 && !is_trailing_empty_segment {
            out.push_str("#pagebreak()\n");
            write_continuation_flow_setup(out, page, size, ctx)?;
            advance_flow_body_alignment(out, ctx);
        }
        if segment.is_empty() {
            continue;
        }

        let has_continuous_frame: bool = write_continuous_flow_frame_open(out, page, ctx);
        let segment_end: usize = *segment_start + segment.len();
        let joins_previous_at_page_boundary: bool =
            paragraph_borders_join_across_layout_boundary(&content[..*segment_start], segment);
        let joins_next_at_page_boundary: bool =
            paragraph_borders_join_across_layout_boundary(segment, &content[segment_end..]);
        generate_blocks_with_layout_space_before(
            out,
            segment,
            ctx,
            section_mark_space.filter(|_| index == 0),
            joins_previous_at_page_boundary,
            joins_next_at_page_boundary,
            segment_end < content.len(),
        )?;
        if has_continuous_frame {
            write_continuous_flow_frame_close(out);
        }
    }
    Ok(())
}

/// Generate Typst markup for multi-column content, one page at a time.
///
/// Typst refuses `#pagebreak()` anywhere below the top level, so a section
/// whose content carries one cannot be emitted as a single column wrapper: the
/// engine rejects the whole document with "pagebreaks are not allowed inside of
/// containers" and no PDF is produced at all (issue #1695). Word ends the page
/// at such a break and restarts the section's columns in column 1 on the next
/// page, which is one wrapper per page's slice of the content with a top-level
/// break between the slices. A break at either end of the section leaves an
/// empty slice, and an empty wrapper would paint nothing while still claiming
/// vertical space, so only the break itself is emitted for those. A break
/// that ends the section is emitted only when the next section continues on
/// the same page; one that does not starts a page of its own (issue #2370).
struct FlowColumnRenderOptions {
    begins_continuously: bool,
    ends_continuously: bool,
    balance_final_segment: bool,
    /// [`take_section_mark_space`]'s gap, written in the first column.
    section_mark_space: Option<(usize, f64)>,
}

fn generate_flow_page_columns(
    out: &mut String,
    page: &FlowPage,
    size: &PageSize,
    content: &[Block],
    cols: &ColumnLayout,
    options: FlowColumnRenderOptions,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let has_leading_page_break: bool =
        ctx.at_document_start && starts_with_page_break_before_content(content);
    let segments: Vec<(usize, &[Block])> = split_at_page_breaks(content);
    let last_segment_index: usize = segments.len() - 1;
    for (index, (segment_start, segment)) in segments.iter().enumerate() {
        let is_trailing_empty_segment: bool =
            index == last_segment_index && segment.is_empty() && !options.ends_continuously;
        if index > 0 && !is_trailing_empty_segment {
            out.push_str("#pagebreak()\n");
            if options.begins_continuously {
                write_continuation_flow_setup(out, page, size, ctx)?;
            }
            advance_flow_body_alignment(out, ctx);
        }
        if segment.is_empty() {
            continue;
        }
        let mut page_content: Vec<Block> = segment.to_vec();
        // A page-top paragraph consumes space only in the column where it
        // starts; writing this gap before `#columns` would shift every column.
        let mut layout_break_space_before: Option<(usize, f64)> = None;
        if index == 0 {
            layout_break_space_before = options.section_mark_space;
        } else {
            let has_content_before_page_break: bool = segments[..index]
                .iter()
                .flat_map(|(_, previous_segment)| previous_segment.iter())
                .any(block_has_paragraph_content);
            let first_content_index: Option<usize> = page_content.iter().position(|block| {
                block_has_paragraph_content(block) || is_empty_paragraph_block(block)
            });
            if has_leading_page_break && index == 1 {
                if let Some(first_content_index) = first_content_index {
                    let space_before: f64 = take_layout_break_space_before(
                        &mut page_content[first_content_index],
                        has_content_before_page_break,
                    )
                    .unwrap_or_default()
                        - leading_empty_paragraph_after_spacing_pt(content);
                    if space_before != 0.0 {
                        layout_break_space_before = Some((first_content_index, space_before));
                    }
                }
            } else if let Some(first_content_index) = first_content_index {
                if block_starts_after_layout_break(&page_content[first_content_index]) {
                    layout_break_space_before = take_layout_break_space_before(
                        &mut page_content[first_content_index],
                        has_content_before_page_break,
                    )
                    .map(|space_before| (first_content_index, space_before));
                } else {
                    // Word collapses the first paragraph's before-spacing at
                    // a later page top unless the break continued that same
                    // source paragraph. A marker on a later block belongs to
                    // its own column break and must not affect this page top.
                    suppress_block_space_before(&mut page_content[first_content_index]);
                }
            }
        }
        let joins_previous_at_page_boundary: bool =
            paragraph_borders_join_across_layout_boundary(&content[..*segment_start], segment);
        let segment_end: usize = *segment_start + segment.len();
        let joins_next_at_page_boundary: bool =
            paragraph_borders_join_across_layout_boundary(segment, &content[segment_end..]);
        let page_boundaries: PageSegmentBoundaries = PageSegmentBoundaries {
            joins_previous_at_page_boundary,
            joins_next_at_page_boundary,
            ends_with_page_break: segment_end < content.len(),
            balance_at_section_end: options.balance_final_segment && index == last_segment_index,
        };
        let suppress_section_boundary_spacing: bool = options.begins_continuously
            && index == 0
            && !segment.is_empty()
            && !page_boundaries.balance_at_section_end;
        if suppress_section_boundary_spacing {
            out.push_str("#block(width: 100%, above: 0pt, below: 0pt)[\n");
        }
        let has_continuous_frame: bool =
            options.begins_continuously && write_continuous_flow_frame_open(out, page, ctx);
        generate_column_section_page(
            out,
            &page_content,
            cols,
            layout_break_space_before,
            page_boundaries,
            ctx,
        )?;
        if suppress_section_boundary_spacing {
            out.push_str("]\n");
        }
        if has_continuous_frame {
            write_continuous_flow_frame_close(out);
        }
    }
    Ok(())
}

fn starts_with_page_break_before_content(content: &[Block]) -> bool {
    content
        .iter()
        .find(|block| !is_empty_paragraph_block(block))
        .is_some_and(|block| matches!(block, Block::PageBreak))
}

/// Word subtracts collapsed after-spacing before an initial hard page break
/// from the first continuation paragraph's before-spacing.
fn leading_empty_paragraph_after_spacing_pt(content: &[Block]) -> f64 {
    content
        .iter()
        .take_while(|block| is_empty_paragraph_block(block))
        .filter_map(|block| paragraph_border_style(block)?.space_after)
        .fold(0.0, f64::max)
}

fn is_empty_paragraph_block(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => paragraph.runs.is_empty(),
        Block::Caption(caption) => caption.paragraph.runs.is_empty(),
        _ => false,
    }
}

fn block_has_paragraph_content(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => !paragraph.runs.is_empty(),
        Block::Caption(caption) => !caption.paragraph.runs.is_empty(),
        Block::List(list) => list
            .items
            .first()
            .and_then(|item| item.content.first())
            .is_some_and(|paragraph| !paragraph.runs.is_empty()),
        _ => false,
    }
}

fn block_starts_after_layout_break(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => paragraph.style.starts_after_layout_break,
        Block::Caption(caption) => caption.paragraph.style.starts_after_layout_break,
        Block::List(list) => list
            .items
            .first()
            .and_then(|item| item.content.first())
            .is_some_and(|paragraph| paragraph.style.starts_after_layout_break),
        _ => false,
    }
}

fn is_layout_break(block: &Block) -> bool {
    matches!(block, Block::PageBreak | Block::ColumnBreak)
}

fn paragraph_borders_join_across_layout_boundary(
    previous_blocks: &[Block],
    next_blocks: &[Block],
) -> bool {
    let previous_block: Option<&Block> = previous_blocks
        .iter()
        .rev()
        .find(|block| !is_layout_break(block));
    let next_block: Option<&Block> = next_blocks.iter().find(|block| !is_layout_break(block));
    paragraph_border_neighbors_match(previous_block, next_block)
}

fn paragraph_border_neighbors_match(
    previous_block: Option<&Block>,
    next_block: Option<&Block>,
) -> bool {
    previous_block
        .zip(next_block)
        .is_some_and(|(previous, next)| paragraph_borders_match(previous, next))
}

/// Generate one page's worth of a multi-column section.
///
/// Equal columns use `#columns(n, gutter: Xpt)[content]`.
/// Unequal columns use `#grid(columns: (W1pt, W2pt, ...), gutter: Xpt)` with
/// content split by `ColumnBreak` blocks into separate grid cells.
#[derive(Clone, Copy, Debug, Default)]
struct PageSegmentBoundaries {
    joins_previous_at_page_boundary: bool,
    joins_next_at_page_boundary: bool,
    ends_with_page_break: bool,
    balance_at_section_end: bool,
}

fn generate_column_section_page(
    out: &mut String,
    content: &[Block],
    cols: &ColumnLayout,
    layout_break_space_before: Option<(usize, f64)>,
    page_boundaries: PageSegmentBoundaries,
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
        let layout_space_location: Option<(usize, usize)> = layout_break_space_before
            .map(|(block_index, _)| column_segment_location(content, block_index));
        for (segment_index, segment) in segments.iter().enumerate() {
            out.push('[');
            let mut cell_content: Vec<Block> =
                segment.iter().map(|block| (*block).clone()).collect();
            if segment_index > 0
                && let Some(first_block) = cell_content.first_mut()
            {
                let has_content_before_column_break: bool = segments[..segment_index]
                    .iter()
                    .flatten()
                    .any(|block| block_has_paragraph_content(block));
                emit_layout_break_space_before(out, first_block, has_content_before_column_break);
            }
            let previous_column_block: Option<&Block> = segments[..segment_index]
                .iter()
                .rev()
                .flat_map(|previous_segment| previous_segment.iter().rev())
                .copied()
                .find(|block| !is_layout_break(block));
            let next_column_block: Option<&Block> = segments[segment_index + 1..]
                .iter()
                .flat_map(|next_segment| next_segment.iter())
                .copied()
                .find(|block| !is_layout_break(block));
            let first_column_block: Option<&Block> = segment
                .iter()
                .copied()
                .find(|block| !is_layout_break(block));
            let last_column_block: Option<&Block> = segment
                .iter()
                .rev()
                .copied()
                .find(|block| !is_layout_break(block));
            let joins_previous_at_boundary: bool = if segment_index == 0 {
                page_boundaries.joins_previous_at_page_boundary
            } else {
                paragraph_border_neighbors_match(previous_column_block, first_column_block)
            };
            let joins_next_at_boundary: bool = if segment_index + 1 == segments.len() {
                page_boundaries.joins_next_at_page_boundary
            } else {
                paragraph_border_neighbors_match(last_column_block, next_column_block)
            };
            let first_block_index: Option<usize> = cell_content
                .iter()
                .position(|block| !is_layout_break(block));
            let last_block_index: Option<usize> = cell_content
                .iter()
                .rposition(|block| !is_layout_break(block));
            for (block_index, original_block) in cell_content.iter().enumerate() {
                if block_index > 0 {
                    out.push('\n');
                }
                if layout_space_location == Some((segment_index, block_index))
                    && let Some((_, space_before)) = layout_break_space_before
                {
                    write_explicit_vertical_space(out, space_before);
                }
                let mut block: Block = original_block.clone();
                let joins_previous_across_page: bool = joins_previous_at_boundary
                    && first_block_index == Some(block_index)
                    && paragraph_border_group_bounds(&cell_content, block_index)
                        .is_some_and(|(group_start, _)| group_start == block_index);
                let joins_next_across_page: bool = joins_next_at_boundary
                    && last_block_index == Some(block_index)
                    && paragraph_border_group_bounds(&cell_content, block_index)
                        .is_some_and(|(_, group_end)| group_end == block_index + 1);
                let joins_previous: bool =
                    previous_non_layout_block_index(&cell_content, block_index).is_some_and(
                        |previous_index| {
                            paragraph_borders_match(&cell_content[previous_index], original_block)
                        },
                    ) || joins_previous_across_page;
                let joins_next: bool = next_non_layout_block_index(&cell_content, block_index)
                    .is_some_and(|next_index| {
                        paragraph_borders_match(original_block, &cell_content[next_index])
                    })
                    || joins_next_across_page;
                suppress_joined_paragraph_border_edges(&mut block, joins_previous, joins_next);
                generate_block(out, &block, ctx)?;
            }
            out.push(']');
        }
        out.push('\n');
    } else {
        if page_boundaries.balance_at_section_end
            && !content
                .iter()
                .any(|block| matches!(block, Block::ColumnBreak))
            && let Some(content_width) = ctx.available_measure_pt
        {
            generate_balanced_flow_columns(
                out,
                content,
                cols,
                content_width,
                layout_break_space_before,
                page_boundaries,
                ctx,
            )?;
        } else {
            // Equal columns: use Typst columns()
            let _ = writeln!(
                out,
                "#columns({}, gutter: {}pt)[",
                cols.num_columns,
                format_f64(cols.spacing)
            );
            generate_blocks_with_layout_space_before(
                out,
                content,
                ctx,
                layout_break_space_before,
                page_boundaries.joins_previous_at_page_boundary,
                page_boundaries.joins_next_at_page_boundary,
                page_boundaries.ends_with_page_break,
            )?;
            out.push_str("\n]\n");
        }
    }
    Ok(())
}

fn generate_balanced_flow_columns(
    out: &mut String,
    content: &[Block],
    cols: &ColumnLayout,
    content_width: f64,
    layout_break_space_before: Option<(usize, f64)>,
    page_boundaries: PageSegmentBoundaries,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let column_width: f64 = (content_width - cols.spacing * f64::from(cols.num_columns - 1))
        / f64::from(cols.num_columns);
    out.push_str("#context {\n  let o2p_section_content = [\n");
    generate_blocks_with_layout_space_before(
        out,
        content,
        ctx,
        layout_break_space_before,
        page_boundaries.joins_previous_at_page_boundary,
        page_boundaries.joins_next_at_page_boundary,
        page_boundaries.ends_with_page_break,
    )?;
    let _ = writeln!(
        out,
        "\n  ]\n  let o2p_column_width = {}pt\n  let o2p_column_height = o2p_section_content.children.fold(\n    measure(o2p_section_content, width: o2p_column_width).height / {},\n    (height, child) => calc.max(height, measure(child, width: o2p_column_width).height),\n  )\n  block(width: {}pt, height: o2p_column_height, above: 0pt, below: 0pt)[\n{}    #columns({}, gutter: {}pt)[#o2p_section_content]\n  ]\n}}",
        format_f64(column_width),
        cols.num_columns,
        format_f64(content_width),
        column_separator_markup(cols, column_width),
        cols.num_columns,
        format_f64(cols.spacing),
    );
    Ok(())
}

/// Emit Word's divider rules over the measured content block. Typst's columns
/// element has no separator parameter, and placing the rules keeps them out of
/// text flow so they do not change the column height.
fn column_separator_markup(cols: &ColumnLayout, column_width: f64) -> String {
    if !cols.has_separator {
        return String::new();
    }

    let mut markup: String = String::new();
    let separator_stroke: String = format_f64(DOCX_COLUMN_SEPARATOR_STROKE_PT);
    for divider_index in 1..cols.num_columns {
        let divider_offset: f64 = column_width * f64::from(divider_index)
            + cols.spacing * (f64::from(divider_index) - 0.5);
        let _ = writeln!(
            markup,
            "    #place(top + left, dx: {}pt, dy: 0pt, line(length: o2p_column_height, angle: 90deg, stroke: {separator_stroke}pt))",
            format_f64(divider_offset),
        );
    }
    markup
}

/// Split content blocks at `PageBreak` boundaries into contiguous slices.
///
/// One slice per page the section spans. A leading, trailing or repeated break
/// yields an empty slice, which the caller emits as the bare break.
fn split_at_page_breaks(content: &[Block]) -> Vec<(usize, &[Block])> {
    let mut segments: Vec<(usize, &[Block])> = Vec::new();
    let mut start: usize = 0;
    for (index, block) in content.iter().enumerate() {
        if matches!(block, Block::PageBreak) {
            segments.push((start, &content[start..index]));
            start = index + 1;
        }
    }
    segments.push((start, &content[start..]));
    segments
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
/// transformed, clipped, or destination-inset images may leave text visible
/// (issue #1432).
fn is_full_page_opaque_image(element: &FixedElement, page_size: PageSize) -> bool {
    const EDGE_TOLERANCE_PT: f64 = 0.5;

    let FixedElementKind::Image(image) = &element.kind else {
        return false;
    };
    if image.rotation_deg.is_some_and(|degrees| degrees != 0.0)
        || image.clip_shape.is_some()
        || image.fill_rect.is_some()
    {
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

    // Excel paints a fitted sheet against a paper box constructed in declared
    // sheet space and then scaled. The move applies to its grid and cell text;
    // floating drawing paint uses the same sheet origin (#1538, #1719).
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
        sheet_drawing_layer(page, centering_inset_pt, paint_offset_pt, ctx)
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
    Ok(())
}

/// Translation from the converter's physical table origin to Excel's fitted
/// sheet-space paint origin.
///
/// On issue #1538's 1,191pt-wide A3 page, the converter centres the 1,078.30pt
/// grid at `floor(50 + (1091 - 1078.3) / 2) - 1 = 55pt`. Excel first snaps the
/// paper box in sheet space: its left edge is `floor(50 / .82) * .82 = 49.20pt`,
/// and its printable width rounds up from there to
/// `ceil((1191 - 100) / .82) * .82 = 1091.42pt`, a right edge of 1140.62pt. Its
/// scaled 1pt centring bias leaves a 5.74pt inset, so the fitted paint origin
/// is 54.94pt, or -0.06pt from the converter's. The top origin similarly moves from 54pt to
/// `floor(54 / .82) * .82 = 53.30pt`.
///
/// Snapping the printable width rather than the right edge is a two-page fit:
/// it is the rule that reproduces both this native origin and the #1719 A3
/// portrait page at 0.78 on Excel's whole-point paper (issue #1730), where
/// `ceil((1191 - 50) / .82)` overshoots #1538 by one sheet point.
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
    let printable_width_pt: f64 = size.width - page.margins.left - page.margins.right;
    let right_pt: f64 = left_pt + (printable_width_pt / scale).ceil() * scale;
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

/// Points Excel prints a horizontally centred sheet left of the whole point
/// the exact centre of its printable width floors to.
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
/// is what makes the result an integer. With the margins (issue #1127) and the
/// ISO page box (issue #1730) both on that whole point, the converter floors
/// the exact centre too; on the fractional 1,190.55pt A3 page the gift budget's
/// 100% control printed 0.275pt right of the native chart frame.
const HORIZONTAL_CENTERING_BIAS_PT: f64 = 1.0;

/// Points to inset the printed sheet from its left margin, or `None` when the
/// sheet is not centred or its grid already fills the printable width.
fn horizontal_centering_inset_pt(page: &SheetPage, size: &PageSize) -> Option<f64> {
    if !page.table.centers_between_print_margins {
        return None;
    }
    let printable_width_pt: f64 = size.width - page.margins.left - page.margins.right;
    let grid_width_pt: f64 = page.table.column_widths.iter().sum();
    let centre_pt: f64 = page.margins.left + (printable_width_pt - grid_width_pt) / 2.0;
    // The epsilon keeps float noise on an exact whole-point centre from
    // flooring it a whole point left.
    let grid_left_pt: f64 = (centre_pt + 1e-9).floor() - HORIZONTAL_CENTERING_BIAS_PT;
    let inset_pt: f64 = grid_left_pt - page.margins.left;
    (inset_pt > 0.0).then_some(inset_pt)
}

/// A drawing overlaid on a sheet at its anchor's absolute coordinates.
enum SheetAnchor<'a> {
    Chart {
        sheet_chart: &'a crate::ir::SheetChart,
        /// The page height, which a page-column clip box reaches down to so
        /// that only its horizontal edges bound the chart (issue #1598).
        page_height_pt: f64,
    },
    Image(&'a crate::ir::SheetImage),
    TextBox(&'a crate::ir::SheetTextBox),
    Shape(&'a crate::ir::SheetShape),
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
    paint_offset_pt: Option<(f64, f64)>,
    ctx: &mut GenCtx,
) -> Option<SheetDrawingLayer> {
    let placed_charts: Vec<&crate::ir::SheetChart> = page
        .charts
        .iter()
        .filter(|chart| chart.placement.is_some())
        .collect();
    if placed_charts.is_empty()
        && page.images.is_empty()
        && page.text_boxes.is_empty()
        && page.shapes.is_empty()
    {
        return None;
    }

    let label: String = format!("o2p-sheet-drawings-{}", ctx.next_sheet_drawing_layer_id);
    ctx.next_sheet_drawing_layer_id += 1;
    // Block-level, not a `box`: an inline box still makes its paragraph lay out
    // a line, which dropped the whole grid by 13.2pt — Typst's default 11pt
    // text at 1.2 leading, independent of the sheet's own font (issue #1101).
    let marker: String =
        format!("#block(width: 100%, height: 0pt, spacing: 0pt)[#metadata(none)<{label}>]");

    // The page foreground does not inherit the table body's move. Apply its
    // fitted sheet-space origin explicitly to every drawing (#1542). A chart's
    // content shares that origin with its frame: Excel lays the plot out in
    // sheet points inside the frame it paints, so seating the content on the
    // converter's physical origin instead left it 0.854 sheet points low at
    // 0.82 and needed a calibration that broke at 100% (issue #1607).
    let (paint_dx_pt, paint_dy_pt) = paint_offset_pt.unwrap_or((0.0, 0.0));
    let left_pt: f64 = page.margins.left + centering_inset_pt.unwrap_or(0.0) + paint_dx_pt;
    let top_pt: f64 = page.margins.top + paint_dy_pt;

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
            &SheetAnchor::Chart {
                sheet_chart,
                page_height_pt: page.size.height,
            },
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
    for sheet_shape in &page.shapes {
        write_placed_sheet_drawing(
            &mut foreground,
            &SheetAnchor::Shape(sheet_shape),
            left_pt,
            top_pt + sheet_shape.y_offset_pt,
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
        SheetAnchor::Chart {
            sheet_chart,
            page_height_pt,
        } => {
            let Some(placement) = sheet_chart.placement else {
                return;
            };
            // A page-column window from width pagination: the chart clips at
            // the tile's printed edge and a continued copy starts at a
            // negative offset, as Excel prints it (issue #1598). The box
            // needs explicit sides because an inner `place` occupies
            // nothing; it reaches the page bottom so it bounds the chart
            // horizontally only.
            if let Some(window) = placement.clip_window {
                let _ = write!(
                    out,
                    "#place(top + left, dx: {}pt)[#box(width: {}pt, height: {}pt, clip: true)[",
                    format_f64(left_pt + window.left_pt),
                    format_f64(window.width_pt),
                    format_f64(*page_height_pt),
                );
            }
            // The anchor sizes the chart, the way a slide's graphicFrame
            // extent does (issue #548); rendering at the intrinsic size
            // instead left the anchored band empty beneath it (issue #982).
            let chart_dx_pt: f64 = match placement.clip_window {
                Some(window) => placement.x_offset_pt - window.left_pt,
                None => left_pt + placement.x_offset_pt,
            };
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt)[",
                format_f64(chart_dx_pt),
            );
            // A fitted sheet prints its drawings shrunk whole, so the chart is
            // laid out at its full frame and the transform brings its text down
            // with its geometry (issue #1069). The corner it grows from is the
            // anchor's, which is where the enclosing `place` put it.
            let fitted: bool = placement.print_scale != 1.0;
            let coordinate_scale: f64 = if placement.print_scale > 0.0 {
                placement.print_scale
            } else {
                1.0
            };
            if fitted {
                // Excel's fit scale is a whole percent, so rounding here keeps
                // 0.82 from printing as 82.00000000000001%.
                let percent: String = format_f64((placement.print_scale * 1e6).round() / 1e4);
                let _ = write!(
                    out,
                    "#scale(x: {percent}%, y: {percent}%, origin: top + left)[",
                );
            }
            let sheet_frame_origin_pt: (f64, f64) = (
                (left_pt + placement.x_offset_pt) / coordinate_scale,
                dy_pt / coordinate_scale,
            );
            generate_sheet_chart_in(
                out,
                &sheet_chart.chart,
                (placement.width, placement.height),
                sheet_frame_origin_pt,
            );
            if fitted {
                out.push(']');
            }
            out.push(']');
            if placement.clip_window.is_some() {
                out.push_str("]]");
            }
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
        SheetAnchor::Shape(sheet_shape) => {
            // The shape's own geometry is relative to its anchor corner, so
            // it is drawn in a `place` at that corner, as a slide shape is
            // drawn at its fixed element's (issue #1566).
            let _ = write!(
                out,
                "#place(top + left, dx: {}pt)[",
                format_f64(left_pt + sheet_shape.x_offset_pt),
            );
            generate_shape(
                out,
                &sheet_shape.shape,
                sheet_shape.width,
                sheet_shape.height,
                ctx,
            );
            out.push(']');
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
        FixedElementKind::TextBox(text_box) => {
            text::with_drawingml_highlight(|| generate_fixed_text_box(out, elem, text_box, ctx))?
        }
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
            // Keep the border in the image's frame transform; source crops and
            // destination stretch affect its pixels, not the authored outline.
            if img.stroke.is_some() {
                out.push_str("#place(top + left)[");
                shapes::write_image_outline(out, img, elem.width, elem.height);
                out.push_str("]\n");
            }
            if flipped {
                out.push(']');
            }
            if rotation.is_some() {
                out.push(']');
            }
            if rotation.is_some() || flipped {
                out.push(']');
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
/// 32.4pt. Typst emits 32.399994pt through krilla, 0.14 and 0.15 alike, putting the source's bottom
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
        || image.fill_rect.is_some()
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
    let aptos_width_scale: Option<f64> = unembedded_aptos_width_scale(text_box);
    let text_layout_width_pt: f64 = aptos_width_scale
        .map(|scale| inner_width_pt / scale)
        .unwrap_or(inner_width_pt);
    // The frame's own content region, which PowerPoint allows to be negative:
    // nothing stops `<a:bodyPr>` from declaring insets deeper than the shape's
    // `<a:ext cy>`. Keep the signed value — the seat below is measured from it.
    let content_region_height_pt: f64 =
        outer_height_pt - text_box.padding.top - text_box.padding.bottom;
    let inner_height_pt: f64 = content_region_height_pt.max(0.0);
    // Typst cannot lay a block out in a region that short: it truncates the
    // content or breaks it into regions stacked a whole inset sum apart, and
    // slide 1 of `tests/fixtures/pptx/poi/with_japanese.pptx` ended up with
    // every paragraph thousands of points above the slide (issue #1706).
    // PowerPoint instead lays the text out at its natural height and seats it
    // on that negative region, overflowing the frame and the slide edge. So
    // the frame stops containing the text: it paints its own fill, stroke and
    // shape, and the block is placed beside it at the seat computed below.
    let insets_exceed_frame: bool = content_region_height_pt <= 0.0;
    // With the text out of the frame, the frame's inset has nothing left to
    // offset — and a padded `#place` would put the shape background back where
    // the inset used to compensate for it.
    let frame_padding: Insets = if insets_exceed_frame {
        Insets::default()
    } else {
        text_box.padding
    };
    let text_box_id: usize = ctx.next_text_box_id();

    let has_custom_shape: bool = text_box.shape_kind.is_some();

    let _ = write!(
        out,
        "#block(width: {}pt, height: {}pt, inset: {}",
        format_f64(outer_width_pt),
        format_f64(outer_height_pt),
        format_insets(&frame_padding),
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
            &frame_padding,
            text_box.fill.as_ref(),
            text_box.opacity,
            &text_box.stroke,
        );
    }
    if insets_exceed_frame {
        // Close the paint-only frame. What follows is a sibling of it, so the
        // seat below is measured from the frame's own top-left corner and the
        // frame's height no longer caps the block's layout region.
        out.push_str("]\n");
    }
    if let Some(paragraph) = single_line_fit_paragraph(text_box, inner_height_pt) {
        let mut raw_paragraph: Paragraph = paragraph.clone();
        raw_paragraph.style.alignment = None;
        let estimated_line_height_pt: f64 = estimate_single_line_height_pt(paragraph);

        let _ = writeln!(out, "  #let text_box_raw_{text_box_id} = [");
        write_powerpoint_relative_baseline_scope_open(out);
        out.push_str("  ");
        // The measured raw paragraph must stay unbreakable through Typst layout,
        // otherwise mixed-font headers can reflow again inside the scaled box.
        generate_fixed_text_paragraph(
            out,
            &raw_paragraph,
            Some(inner_width_pt),
            true,
            PowerPointBaselineMode::Relative,
        )?;
        out.push_str("] }\n  ]\n");

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
    } else if let Some(paragraph) = wrapped_fit_paragraph(text_box, content_region_height_pt) {
        let _ = writeln!(
            out,
            "  #let text_box_raw_{text_box_id} = block(width: {}pt)[",
            format_f64(inner_width_pt),
        );
        write_powerpoint_relative_baseline_scope_open(out);
        out.push_str("  ");
        generate_fixed_text_paragraph(
            out,
            paragraph,
            Some(inner_width_pt),
            false,
            PowerPointBaselineMode::Relative,
        )?;
        out.push_str("] }\n  ]\n");

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
        if let Some(scale) = aptos_width_scale {
            let _ = writeln!(
                out,
                "  #let text_box_content_{text_box_id} = scale(x: {}%, y: 100%, origin: top + left, reflow: true)[",
                format_f64(scale * 100.0),
            );
            let _ = writeln!(
                out,
                "    #block(width: {}pt)[",
                format_f64(text_layout_width_pt),
            );
        } else {
            let _ = writeln!(
                out,
                "  #let text_box_content_{text_box_id} = block(width: {}pt)[",
                format_f64(inner_width_pt),
            );
        }
        write_powerpoint_relative_baseline_scope_open(out);
        for (index, block) in text_box.content.iter().enumerate() {
            if index > 0 {
                out.push('\n');
            }
            out.push_str("  ");
            generate_fixed_text_box_block(
                out,
                block,
                ctx,
                Some(text_layout_width_pt),
                text_box.no_wrap,
                if matches!(text_box.vertical_align, TextBoxVerticalAlign::Top) {
                    PowerPointBaselineMode::Top
                } else {
                    PowerPointBaselineMode::Relative
                },
            )?;
        }
        out.push_str("] }\n");
        out.push_str("  ]\n");
        if aptos_width_scale.is_some() {
            out.push_str("  ]\n");
        }
    }

    if insets_exceed_frame {
        write_overflowing_inset_seat(
            out,
            text_box_id,
            text_box.vertical_align,
            content_region_height_pt,
            &text_box.padding,
        );
        return Ok(());
    }

    match text_box.vertical_align {
        TextBoxVerticalAlign::Top => {
            let _ = writeln!(out, "  #text_box_content_{text_box_id}");
        }
        TextBoxVerticalAlign::Center | TextBoxVerticalAlign::Bottom => {
            out.push_str("  #context {\n");
            if !text_box.auto_fit && matches!(text_box.vertical_align, TextBoxVerticalAlign::Center)
            {
                // DrawingML's center anchor stays centered when no autofit
                // reduces content that overflows the text frame (#1963).
                let _ = writeln!(
                    out,
                    "    let text_box_slack_{text_box_id} = {}pt - measure(text_box_content_{text_box_id}).height",
                    format_f64(inner_height_pt),
                );
            } else {
                let _ = writeln!(
                    out,
                    "    let text_box_slack_{text_box_id} = calc.max({}pt - measure(text_box_content_{text_box_id}).height, 0pt)",
                    format_f64(inner_height_pt),
                );
            }
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

/// Place a block whose frame is shallower than its own vertical insets.
///
/// PowerPoint collapses a content region of zero or negative height to the
/// single line at its own middle, `tIns + (height - tIns - bIns) / 2` below the
/// frame's top, and seats the block's natural height on that line as the
/// anchor asks: `t` puts the block's top there, `ctr` its middle, `b` its
/// bottom. The block then overflows the frame — and the slide — in both
/// directions.
///
/// Measured, not inferred. Eight native PowerPoint for Mac exports of
/// `tests/fixtures/pptx/poi/with_japanese.pptx`, whose frame is 0pt tall with
/// 71.98pt insets: two with the frame grown to 540pt fix the block's own
/// height at 440.44pt, and the six below keep the degenerate frame and place
/// its top against that line. Each one's predicted position is within 0.21pt
/// of the export, inside PowerPoint's 0.24pt quantisation. The `b` row's frame
/// is also translated 360pt down the slide, which is a pure offset, because
/// its block is otherwise drawn entirely above the slide's top edge.
///
/// | anchor | tIns | bIns | predicted | exported |
/// | --- | ---: | ---: | ---: | ---: |
/// | `ctr` | 71.98 | 71.98 | -220.22 | -220.10 |
/// | `ctr` | 71.98 | 0 | -184.23 | -184.10 |
/// | `t` | 71.98 | 71.98 | 0.00 | -0.02 |
/// | `t` | 0 | 71.98 | -35.99 | -35.78 |
/// | `t` | 71.98 | 0 | +35.99 | +35.98 |
/// | `b` | 71.98 | 71.98 | -440.44 | -440.42 |
///
/// The block's top against the frame's top, in points. Neither seating the
/// block on the top inset nor centring it on the frame reproduces the `t`
/// rows: both are independent of `bIns`, which moves the exported block half
/// its own change (issue #1706).
fn write_overflowing_inset_seat(
    out: &mut String,
    text_box_id: usize,
    vertical_align: TextBoxVerticalAlign,
    content_region_height_pt: f64,
    padding: &Insets,
) {
    // The collapsed region: what is left of it once its own negative height is
    // spent equally against both insets.
    let seat_line_pt: f64 = padding.top + content_region_height_pt / 2.0;
    match vertical_align {
        // The block's top is the line itself, so this seat needs no
        // measurement.
        TextBoxVerticalAlign::Top => {
            let _ = writeln!(
                out,
                "  #place(top + left, dx: {}pt, dy: {}pt, text_box_content_{text_box_id})",
                format_f64(padding.left),
                format_f64(seat_line_pt),
            );
        }
        TextBoxVerticalAlign::Center | TextBoxVerticalAlign::Bottom => {
            let measured_height: String = format!("measure(text_box_content_{text_box_id}).height");
            let seat_expr: String = match vertical_align {
                TextBoxVerticalAlign::Center => {
                    format!("{}pt - {measured_height} / 2", format_f64(seat_line_pt))
                }
                TextBoxVerticalAlign::Bottom => {
                    format!("{}pt - {measured_height}", format_f64(seat_line_pt))
                }
                TextBoxVerticalAlign::Top => unreachable!(),
            };
            out.push_str("  #context {\n");
            let _ = writeln!(out, "    let text_box_seat_{text_box_id} = {seat_expr}");
            let _ = writeln!(
                out,
                "    place(top + left, dx: {}pt, dy: text_box_seat_{text_box_id}, text_box_content_{text_box_id})",
                format_f64(padding.left),
            );
            out.push_str("  }\n");
        }
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

/// The page-space `(top, bottom)` bands of the footer pictures its lines
/// cannot sit beside, lowest first, each widened by its `distT`/`distB`.
///
/// A picture wrapped top and bottom keeps text off both of its sides; a
/// square or tight one only where it leaves no room beside it in the text
/// column. TODO(side wrap): one that leaves room lets Word set the lines
/// beside it, narrowed, at their own height, so they stay where they are.
fn footer_line_obstacles(footer: &HeaderFooter, page: &FlowPage) -> Vec<(f64, f64)> {
    let column_left: f64 = page.margins.left;
    let column_right: f64 = page.size.width - page.margins.right;
    let mut obstacles: Vec<(f64, f64)> = footer
        .shapes
        .iter()
        .filter_map(|shape| {
            let HeaderFooterShapeContent::Picture(picture) = &shape.content else {
                return None;
            };
            let distance = picture.wrap_distance;
            let covers_column = || {
                let left: f64 = shape.frame.x? - distance.left;
                let right: f64 = shape.frame.x? + shape.width + distance.right;
                Some(left <= column_left && right >= column_right)
            };
            let keeps_lines_off: bool = match picture.wrap_mode {
                WrapMode::TopAndBottom => true,
                WrapMode::Square | WrapMode::Tight => covers_column().unwrap_or(false),
                WrapMode::Behind | WrapMode::InFront | WrapMode::None => false,
            };
            if !keeps_lines_off || shape.frame.vertical_anchor != FrameAnchor::Page {
                return None;
            }
            let top: f64 = shape.frame.y?;
            Some((top - distance.top, top + shape.height + distance.bottom))
        })
        .collect();
    obstacles.sort_by(|left, right| right.0.total_cmp(&left.0));
    obstacles
}

/// How far Word's line box for a flowed story's first line reaches above the
/// cap height Typst starts that line at, in points; zero where the face is
/// unknown, the line holds a picture, or the story is laid out as stacked
/// blocks of Word's own line heights.
fn hf_story_first_line_ascent_above_cap_pt(hf: &HeaderFooter) -> f64 {
    if hf_story_rules_a_paragraph(hf) {
        return 0.0;
    }
    hf.paragraphs
        .iter()
        .find(|paragraph| hf_paragraph_is_emitted(paragraph))
        .filter(|paragraph| {
            !paragraph
                .elements
                .iter()
                .any(|element| matches!(element, HFInline::Image(_)))
        })
        .and_then(|paragraph| text::word_header_band_shift_pt(&hf_paragraph_metric_runs(paragraph)))
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
        let descent_em: Option<f64> = footer
            .paragraphs
            .last()
            .map(hf_paragraph_metric_runs)
            .and_then(|runs| text::word_line_box_descent_em(&runs));
        let bottom_edge: String = descent_em
            .map(|descent_em| format!("-{}em", format_f64(descent_em)))
            .unwrap_or_else(|| "\"descender\"".to_string());
        // Word reserves the last paragraph's own `w:spacing w:after` between
        // its last line and that anchor, so the drawn band stops that far above
        // `w:pgMar/@w:footer` rather than on it. A gap deeper than the band
        // itself would invert the block, so the band is as far as the story can
        // be lifted (issue #1195).
        let space_after: f64 = hf_trailing_space_after_pt(footer);
        let drawn_band: f64 = round_to_hundredths((band - space_after).max(0.0));
        let _ = write!(
            value,
            "block(width: 100%, height: {}pt)[#set text(bottom-edge: {bottom_edge}); ",
            format_f64(drawn_band)
        );
        let obstacles: Vec<(f64, f64)> = footer_line_obstacles(footer, page);
        if obstacles.is_empty() {
            value.push_str("#place(bottom, block(width: 100%)[");
        } else {
            value.push_str("#context { let story = block(width: 100%)[");
        }
        // That edge sits under every line of the story, so the leading has to
        // leave room for it. Where it is the renderer's own `"descender"`, the
        // last paragraph's face did not resolve and no descent is known to
        // leave room for.
        generate_flow_hf_content(&mut value, footer, descent_em.unwrap_or(0.0), ctx);
        if obstacles.is_empty() {
            value.push_str("])]");
        } else {
            // Word sets the story's lines above a picture they cannot sit
            // beside: its box, `w:spacing w:after` included, ends at the top
            // of the first picture its bottom-anchored lines would cross, and
            // of each one that box then crosses (issue #2327).
            let page_height: f64 = page.size.height;
            let story_bottom: f64 = page_height - page.margins.bottom + drawn_band + space_after;
            let text_width: f64 = page.size.width - page.margins.left - page.margins.right;
            // The measured story starts at its first line's cap height;
            // Word's line box reaches that face's ascent.
            let word_box_extra: f64 = hf_story_first_line_ascent_above_cap_pt(footer) + space_after;
            let tuples: String = obstacles
                .iter()
                .map(|(above, below)| {
                    format!("({}pt, {}pt), ", format_f64(*above), format_f64(*below))
                })
                .collect();
            let _ = write!(
                value,
                "]; let height = measure(story, width: {}pt).height + {}pt; let edge = {}pt; for (above, below) in ({tuples}) {{ if edge - height < below and above < edge {{ edge = above }} }}; place(bottom, dy: edge - {}pt, story) }}]",
                format_f64(text_width),
                format_f64(word_box_extra),
                format_f64(story_bottom),
                format_f64(story_bottom),
            );
        }
        return FlowFooterValue {
            markup: value,
            wants_zero_descent: true,
        };
    }
    if hf_needs_stack_offset(footer) {
        value.push_str("context { let footer_content = block(width: 100%)[");
        generate_flow_hf_content(&mut value, footer, 0.0, ctx);
        value.push_str("]; move(dy: -measure(footer_content).height / 2)[#footer_content] }");
    } else {
        if hf_needs_context(footer) {
            value.push_str("context ");
        }
        value.push('[');
        generate_flow_hf_content(&mut value, footer, 0.0, ctx);
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

/// A continuous section that shows first-page stories (`<w:titlePg/>`) and
/// shares the page setup an earlier section writes. Typst cannot set a page
/// up again mid-page, so that setup chooses the section's first stories for
/// it, on the page its first line opens (issue #2254).
#[derive(Clone)]
struct TitlePageRunMember {
    section_index: usize,
    header: Option<HeaderFooter>,
    footer: Option<HeaderFooter>,
    first_header: Option<HeaderFooter>,
    first_footer: Option<HeaderFooter>,
}

/// The continuous title-page sections after `pages[index]` that share the
/// page setup `pages[index]` writes. `section_index` is that section's flow
/// index. A section holding a page break sets its page up again there
/// ([`generate_continuous_flow_blocks`]), so the sections after it join that
/// setup instead.
fn title_page_run_members(
    pages: &[Page],
    index: usize,
    section_index: usize,
) -> Vec<TitlePageRunMember> {
    let mut members: Vec<TitlePageRunMember> = Vec::new();
    for (offset, page) in pages.iter().skip(index + 1).enumerate() {
        let Page::FlowContinuous(flow) = page else {
            break;
        };
        // A section stating a first-page story always carries the label
        // `generate_flow_page` writes, which the member's branch reads.
        if flow.first_header.is_some() || flow.first_footer.is_some() {
            members.push(TitlePageRunMember {
                section_index: section_index + offset + 1,
                header: flow.header.clone(),
                footer: flow.footer.clone(),
                first_header: flow.first_header.clone(),
                first_footer: flow.first_footer.clone(),
            });
        }
        if flow
            .content
            .iter()
            .any(|block| matches!(block, Block::PageBreak))
        {
            break;
        }
    }
    members
}

/// Typst condition for "this page is a continuous section's first page".
/// The section's label is laid out with its first line, because Typst holds
/// a tag back until the next frame lands. Word gives a page the two sections
/// share to the earlier one, as its references for docxide-pdf
/// `cont-first-first-fits` show, so the label also has to sit at the body
/// top, `body_top_pt`, where a line that opens the page lands once Typst has
/// trimmed the weak gap above it.
fn continuous_section_first_page_condition(section_index: usize, body_top_pt: f64) -> String {
    format!(
        "{{ let start = locate(<{}>).position(); start.page == here().page() and start.y <= {}pt }}",
        section_first_page_label(section_index),
        format_f64(body_top_pt + 0.01)
    )
}

/// The first-page stories a page setup chooses between, each with the Typst
/// condition for the page that shows it: the section's own, then each
/// member's. A story without flow content shows nothing there.
fn first_page_story_choices<'a>(
    (own_first_page, own): (&str, Option<&'a HeaderFooter>),
    members: &'a [TitlePageRunMember],
    member_story: fn(&TitlePageRunMember) -> Option<&HeaderFooter>,
    body_top_pt: f64,
) -> Vec<(String, Option<&'a HeaderFooter>)> {
    own.map(|story| (own_first_page.to_string(), story))
        .into_iter()
        .chain(members.iter().filter_map(|member| {
            member_story(member).map(|story| {
                (
                    continuous_section_first_page_condition(member.section_index, body_top_pt),
                    story,
                )
            })
        }))
        .map(|(condition, story)| {
            (
                condition,
                Some(story).filter(|story| hf_has_flow_content(story)),
            )
        })
        .collect()
}

/// A page setup value that picks, per page, the first branch whose
/// condition holds, or `default`. `open` and `close` wrap each choice: `{`/`}`
/// around code, `[`/`]` around markup.
fn per_page_story_choice(
    branches: &[(String, String)],
    default: &str,
    open: &str,
    close: &str,
) -> String {
    let mut out = String::from("context { ");
    for (condition, value) in branches {
        let _ = write!(out, "if {condition} {open}{value}{close} else ");
    }
    let _ = write!(out, "{open}{default}{close} }}");
    out
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
                generate_flow_header_content(&mut out, header, ctx);
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
    generate_flow_header_content(&mut out, header, ctx);
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
/// On continuation pages, a page-anchored header shape that crosses the body
/// top margin also reserves its bottom and any following empty header line
/// boxes, as Word does (issue #1995). A table in the story's flow counts with
/// the story's lines (issue #2250), except in a section whose first page has
/// its own story (`TODO(#2216)` below).
///
/// A continuation page is one a continuous section sets up again after a page
/// break it holds. It is never the section's first page, so it reserves only
/// the default story (issue #2224).
///
/// Otherwise `w:top` is left alone, so the common case emits exactly what it
/// always did. When a header face cannot be measured, its nonempty text cannot
/// contribute a guessed height. Empty flow lines after a known overlay still
/// reserve their declared/default font size and paragraph spacing.
fn flow_page_top_margin_pt(
    page: &FlowPage,
    size: &PageSize,
    default_text: Option<&TextStyle>,
    continuation_page: bool,
) -> f64 {
    // One margin serves the whole section, so where `w:titlePg` gives the first
    // page its own story both have to fit it — taking the default story alone
    // would leave a taller first-page header overprinting page one, which is
    // the same defect one page in (#846 added that second story).
    let first_header: Option<&HeaderFooter> =
        page.first_header.as_ref().filter(|_| !continuation_page);
    [page.header.as_ref(), first_header]
        .into_iter()
        .flatten()
        .filter_map(|header| {
            let overlay_bottom: Option<f64> = continuation_page
                .then(|| page_anchored_header_shape_bottom_pt(header, size, page.margins.top))
                .flatten();
            // TODO(#2216): one margin serves both stories of a `w:titlePg`
            // section, so a default story's table would also lower a first
            // page that does not show it. Microsoft's Initials resume template
            // has a 147.6pt exact-row table in its default header and no
            // first-page story, and that one-page resume ran to two. Reserve
            // tables there once the margin is per page (#2300).
            let story_tables_pt: f64 = if first_header.is_none() {
                hf_story_tables_height_pt(
                    header,
                    page.line_grid_pitch.filter(|_| page.line_grid_snaps_lines),
                )
            } else {
                0.0
            };
            // A story of nothing but empty paragraphs still has their lines
            // (issue #2056), and one of nothing but a table its rows.
            if !hf_has_flow_content(header)
                && !header
                    .paragraphs
                    .iter()
                    .any(hf_paragraph_is_empty_flow_line)
                && overlay_bottom.is_none()
                && story_tables_pt <= 0.0
            {
                return None;
            }

            // The band runs from the `w:header` line down, so a story that
            // overflows needs the margin to reach the bottom of its content.
            // On continuation pages, also retain the empty flow paragraphs
            // that follow a page-anchored header overlay; Word reserves their
            // line boxes even though they paint no text.
            let text_width_pt: f64 = size.width - page.margins.left - page.margins.right;
            let distance_pt: f64 = header.distance_from_edge.unwrap_or(0.0);
            let content_reach = |line_heights: HfLineHeights| -> Option<f64> {
                hf_content_height_pt(
                    header,
                    default_text,
                    overlay_bottom.is_some(),
                    text_width_pt,
                    line_heights,
                )
                .map(|height| distance_pt + height)
            };
            // Where the drawn story's lines end, and where Word's own story
            // ends once each line takes its stated height and each paragraph
            // its spacing.
            let drawn_lines: HfLineHeights = if header_story_draws_line_boxes(header) {
                HfLineHeights::Boxed
            } else {
                HfLineHeights::Natural
            };
            let drawn_reach: Option<f64> = content_reach(drawn_lines);
            // Word stacks the story's flow tables with its paragraphs. The
            // drawn story leaves them out: they are placed at `w:header` in the
            // page background, so its band shift does not ride on them.
            let word_reach: Option<f64> =
                content_reach(HfLineHeights::Declared).map(|reach| reach + story_tables_pt);
            if drawn_reach.is_none() && word_reach.is_none() && overlay_bottom.is_none() {
                return None;
            }
            let overlay_reach: f64 = overlay_bottom
                .map(|bottom| {
                    bottom + hf_last_flow_paragraph_height_pt(header, default_text).unwrap_or(0.0)
                })
                .unwrap_or(0.0);
            let reach: f64 = drawn_reach.unwrap_or(0.0).max(overlay_reach);
            let word_reach: f64 = word_reach.unwrap_or(0.0);
            // A story that fits its band is left exactly as it renders today,
            // clamp and all. Growing those too would move the body a fraction
            // of a point on ordinary documents — measured at 0.63pt for a
            // single 12pt Malgun line well inside its band — which is a
            // different change from stopping an overflow, and one no reference
            // export here justifies.
            if reach.max(word_reach) <= page.margins.top {
                return None;
            }
            // An overflowing story needs room for the ascent correction as
            // well, or `write_shifted_header_band` clamps the shift straight
            // back off: the two-line Malgun letterhead measured 41.4984pt of
            // content against a 41.5pt band, 0.0016pt of room for a 6.84pt
            // shift (#629 seats that baseline). That clamp measures the drawn
            // story, so the room is kept against its reach. Word starts the
            // body below its own exact lines without that correction: a 160pt
            // line 6pt down starts the body at 166pt (issue #2056).
            Some((reach + header_band_shift_pt(header).unwrap_or(0.0).max(0.0)).max(word_reach))
        })
        .fold(page.margins.top, f64::max)
}

/// Which layout [`hf_content_height_pt`] measures a header story in.
#[derive(Clone, Copy)]
enum HfLineHeights {
    /// The face's natural line, which is what the flowed drawn story lays out.
    /// Neither drawn layout keeps paragraph spacing (issue #2062).
    Natural,
    /// An exact line at its stated height and any other at the face's natural
    /// line, which is what a story drawn as Word's line boxes lays out.
    Boxed,
    /// Word's own: the paragraph's exact, at-least or proportional
    /// `w:spacing w:line` where it states one, and its `w:before`/`w:after`
    /// gaps (see [`hf_content_height_pt`]).
    Declared,
}

/// Height a header or footer story's lines take, in points.
///
/// Reserve each measurable wrapped line plus its paragraph border rules. An
/// empty paragraph is a line of its paragraph mark in Word's measurement and
/// nothing in the drawn story's, which leaves it out (issue #2062). A
/// horizontal rule's paragraph is a line of its mark in both, because the
/// drawn story draws it (issue #2251).
///
/// Word's measure also keeps paragraph spacing and proportional lines, which
/// the drawn story drops (issue #2062). Native Word for Mac probes of a
/// three-paragraph Arial 12pt header (issue #2056) moved the body by exactly
/// the first paragraph's `w:before`, the larger of each adjacent
/// `w:after`/`w:before` pair (12pt after against 6, 12 and 18pt before gave
/// 12, 12 and 18pt), and the last paragraph's `w:after`. A `w:lineRule="auto"`
/// multiple of 1.5 or 0.5 scaled each 13.8pt line by it, within Word's 0.24pt
/// export grid. An empty paragraph with a `w:pBdr` rule retains its existing
/// behavior.
fn hf_content_height_pt(
    hf: &HeaderFooter,
    default_text: Option<&TextStyle>,
    include_empty_flow_paragraphs: bool,
    text_width_pt: f64,
    line_heights: HfLineHeights,
) -> Option<f64> {
    let mut total: f64 = 0.0;
    // The `w:after` below the paragraph last measured, which the next one's
    // `w:before` collapses against.
    let mut pending_after_pt: Option<f64> = None;
    for paragraph in &hf.paragraphs {
        if include_empty_flow_paragraphs
            && paragraph.frame.as_ref().is_some_and(is_page_anchored_frame)
        {
            continue;
        }
        if let HfLineHeights::Declared = line_heights {
            let before_pt: f64 = paragraph.style.space_before.unwrap_or(0.0).max(0.0);
            total += pending_after_pt.map_or(before_pt, |after_pt| after_pt.max(before_pt));
            pending_after_pt = Some(paragraph.style.space_after.unwrap_or(0.0).max(0.0));
        }
        // A rule's paragraph is a line of its mark in Word's measure and in
        // the drawn story alike, which draws it (issue #2251), and so is one
        // that anchors a drawing (issue #2327).
        if hf_paragraph_is_mark_line(paragraph) {
            total += hf_empty_paragraph_line_pt(paragraph, default_text, line_heights)?;
            continue;
        }
        if !include_empty_flow_paragraphs && hf_paragraph_is_empty_flow_line(paragraph) {
            total += match line_heights {
                HfLineHeights::Natural | HfLineHeights::Boxed => 0.0,
                HfLineHeights::Declared => {
                    hf_empty_paragraph_line_pt(paragraph, default_text, line_heights)?
                }
            };
            continue;
        }
        let line_count: usize = hf_paragraph_line_count(paragraph, text_width_pt);
        let stated_line_pt: Option<f64> = match line_heights {
            HfLineHeights::Natural => None,
            HfLineHeights::Boxed => hf_paragraph_exact_line_pt(paragraph),
            HfLineHeights::Declared => hf_paragraph_declared_line_pt(paragraph),
        };
        if let Some(line_pt) = stated_line_pt {
            total += line_count as f64 * line_pt + hf_paragraph_border_extent_pt(paragraph);
            continue;
        }
        if include_empty_flow_paragraphs {
            total += hf_paragraph_height_pt(paragraph, default_text, true)?;
        } else {
            total += hf_paragraph_height_pt(paragraph, None, false)?;
        }
        let extra_lines: usize = line_count - 1;
        if extra_lines > 0 {
            total += extra_lines as f64
                * text::word_line_advance_pt(&hf_paragraph_metric_runs(paragraph))?;
        }
    }
    Some(total + pending_after_pt.unwrap_or(0.0))
}

/// The line height a header or footer paragraph with text states for each of
/// its lines: exact, at least, or proportional.
///
/// Word lays every line of an exact paragraph out at exactly that height,
/// whatever its face, so the height needs no face metrics. A proportional
/// (`w:lineRule="auto"`) one scales the face's natural line, as a body line
/// does (issue #2056). An empty flow paragraph is measured by
/// [`hf_empty_paragraph_line_pt`] instead, which honours these as well. One
/// holding only a picture or a positioned tab still takes the natural-line
/// path, which cannot measure it (issue #2164).
///
/// TODO(at-least lines): an at-least height is only a floor under the face's
/// natural line, but the top-margin measure still takes it as the line itself.
fn hf_paragraph_declared_line_pt(paragraph: &crate::ir::HeaderFooterParagraph) -> Option<f64> {
    let runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
    if runs.is_empty() {
        return None;
    }
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Exact(line_pt) | crate::ir::LineSpacing::AtLeast(line_pt))
            if line_pt > 0.0 =>
        {
            Some(line_pt)
        }
        Some(crate::ir::LineSpacing::Proportional(factor)) if factor > 0.0 => {
            Some(text::word_line_advance_pt(&runs)? * factor)
        }
        _ => None,
    }
}

/// The height a header or footer paragraph with text states as an exact
/// line, as [`hf_paragraph_declared_line_pt`] reads it.
fn hf_paragraph_exact_line_pt(paragraph: &crate::ir::HeaderFooterParagraph) -> Option<f64> {
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Exact(_)) => hf_paragraph_declared_line_pt(paragraph),
        _ => None,
    }
}

/// Whether a header or footer paragraph is an empty line in its story's flow:
/// no content, no rule, and no `w:framePr` frame, which Word positions out of
/// the flow. A scraped docxide-pdf decree's first-page header is an empty
/// frame over an empty paragraph; counting the frame as a line ran the decree
/// to three pages where Word lays it out on two.
fn hf_paragraph_is_empty_flow_line(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    !hf_paragraph_has_content(paragraph) && paragraph.frame.is_none()
}

/// The line Word gives an empty header or footer paragraph, in points.
///
/// It is the line of the paragraph mark's face: two scraped docxide-pdf
/// specifications whose headers end in an empty paragraph, with an 11pt and a
/// 12pt Times New Roman mark, have Word start the body 1.25pt apart. An exact
/// `w:spacing w:line` sets it whatever the face, and in Word's measure a
/// proportional one scales it, as they do a line of text.
fn hf_empty_paragraph_line_pt(
    paragraph: &crate::ir::HeaderFooterParagraph,
    default_text: Option<&TextStyle>,
    line_heights: HfLineHeights,
) -> Option<f64> {
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Exact(line_pt) | crate::ir::LineSpacing::AtLeast(line_pt))
            if line_pt > 0.0 =>
        {
            Some(line_pt)
        }
        // The drawn story sets a rule's mark line at the face's natural line,
        // so only Word's measure scales it.
        Some(crate::ir::LineSpacing::Proportional(factor))
            if factor > 0.0 && matches!(line_heights, HfLineHeights::Declared) =>
        {
            match text::word_line_advance_pt(&[hf_empty_paragraph_mark_run(
                paragraph,
                default_text,
            )]) {
                Some(line_pt) => Some(line_pt * factor),
                // Without the face's metrics, the mark's size stands in for its
                // line, and that fallback already scales it.
                None => hf_paragraph_height_pt(paragraph, default_text, true),
            }
        }
        _ => hf_paragraph_height_pt(paragraph, default_text, true),
    }
}

/// How many lines a header or footer paragraph takes in `width_pt`, as
/// [`runs_line_count`] lays its runs out inside its indents.
fn hf_paragraph_line_count(paragraph: &crate::ir::HeaderFooterParagraph, width_pt: f64) -> usize {
    let available_pt: f64 = width_pt
        - paragraph.style.indent_left.unwrap_or(0.0)
        - paragraph.style.indent_right.unwrap_or(0.0);
    runs_line_count(&hf_paragraph_metric_runs(paragraph), available_pt)
}

/// How many lines a paragraph's runs take in `available_pt`.
///
/// Its words are laid greedily on the faces' own advances, as a line breaker
/// fills them, and a hard break starts a line. Kerning is ignored, which the
/// one-line-per-paragraph estimate this replaced ignored as well. A paragraph
/// that cannot be measured — a run without a face, a tab, a face that does
/// not resolve — counts one line, so the margin never grows on a guess.
fn runs_line_count(runs: &[Run], available_pt: f64) -> usize {
    if available_pt <= 0.0 || runs.iter().any(|run| run.text.contains('\t')) {
        return 1;
    }
    let mut lines: usize = 1;
    let mut line_width_pt: f64 = 0.0;
    let mut pending_word_pt: f64 = 0.0;
    let place_word = |word_pt: f64, line_width: &mut f64, line_count: &mut usize| {
        if word_pt > 0.0 && *line_width > 0.0 && *line_width + word_pt > available_pt {
            *line_count += 1;
            *line_width = word_pt;
        } else {
            *line_width += word_pt;
        }
    };
    for run in runs {
        let Some(family) = run.style.font_family.as_deref() else {
            return 1;
        };
        let is_bold: bool = matches!(run.style.bold, Some(true));
        let font_size_pt: f64 = run
            .style
            .font_size
            .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);
        let advance_pt = |text: &str| -> Option<f64> {
            // A table cell's page field shapes as digits, as a story
            // paragraph's [`hf_paragraph_metric_runs`] stands it in; its
            // marker has no glyph to measure (issue #2338).
            const FIELDS: [char; 2] = [PAGE_NUMBER_FIELD_CHAR, TOTAL_PAGES_FIELD_CHAR];
            let text: std::borrow::Cow<'_, str> = if text.contains(FIELDS) {
                text.replace(FIELDS, "1").into()
            } else {
                text.into()
            };
            Some(crate::render::pdf::text_advance_em(family, is_bold, &text)? * font_size_pt)
        };
        // OOXML may split a word into runs even when its formatting is
        // unchanged. Only an actual space or hard break ends the word.
        for piece in run.text.split_inclusive([' ', '\n', '\r', '\u{000B}']) {
            let word: &str = piece.trim_end_matches([' ', '\n', '\r', '\u{000B}']);
            let Some(fragment_pt) = advance_pt(word) else {
                return 1;
            };
            pending_word_pt += fragment_pt;
            if piece.ends_with(' ') {
                place_word(pending_word_pt, &mut line_width_pt, &mut lines);
                let Some(space_pt) = advance_pt(" ") else {
                    return 1;
                };
                line_width_pt += space_pt;
                pending_word_pt = 0.0;
            } else if piece.ends_with(['\n', '\r', '\u{000B}']) {
                place_word(pending_word_pt, &mut line_width_pt, &mut lines);
                lines += 1;
                line_width_pt = 0.0;
                pending_word_pt = 0.0;
            }
        }
    }
    place_word(pending_word_pt, &mut line_width_pt, &mut lines);
    lines
}

/// The bottom of a page-anchored header shape that crosses the body's top edge.
///
/// A table in the story's flow is not an overlay: its rows are measured with
/// the story's paragraphs on every page (issue #2250).
fn page_anchored_header_shape_bottom_pt(
    header: &HeaderFooter,
    size: &PageSize,
    top_margin_pt: f64,
) -> Option<f64> {
    header
        .shapes
        .iter()
        .filter(|shape| !shape.flows_in_story)
        .filter(|shape| shape.frame.vertical_anchor == crate::ir::FrameAnchor::Page)
        .filter_map(|shape| {
            if shape.height <= 0.0 {
                return None;
            }
            let top: f64 = shape.frame.y.unwrap_or_else(|| {
                aligned_offset(shape.frame.vertical_align, size.height, Some(shape.height))
            });
            let bottom: f64 = top.max(0.0) + shape.height;
            (top < top_margin_pt && bottom > top_margin_pt).then_some(bottom)
        })
        .max_by(f64::total_cmp)
}

/// The height of the tables in a header story's flow, in points, which Word
/// stacks with the story's paragraphs (issue #2250). A table that cannot be
/// measured adds nothing, so the margin never grows on a guess.
fn hf_story_tables_height_pt(header: &HeaderFooter, line_grid_pitch: Option<f64>) -> f64 {
    header
        .shapes
        .iter()
        .filter(|shape| shape.flows_in_story)
        .filter_map(|shape| match &shape.content {
            HeaderFooterShapeContent::Table(table) => {
                let height_pt: Option<f64> =
                    tables::word_story_table_height_pt(table, line_grid_pitch);
                if height_pt.is_none() {
                    tracing::debug!(
                        rows = table.rows.len(),
                        "header table in the story flow is not measurable; not reserved"
                    );
                }
                height_pt
            }
            HeaderFooterShapeContent::Shape(_) | HeaderFooterShapeContent::Picture(_) => None,
        })
        .sum()
}

/// Reserve the final flowing header line below its page-anchored overlay.
fn hf_last_flow_paragraph_height_pt(
    header: &HeaderFooter,
    default_text: Option<&TextStyle>,
) -> Option<f64> {
    header
        .paragraphs
        .iter()
        .rev()
        .filter(|paragraph| !paragraph.frame.as_ref().is_some_and(is_page_anchored_frame))
        .find_map(|paragraph| hf_paragraph_height_pt(paragraph, default_text, true))
}

/// The height one header or footer paragraph takes: Word's line for its face,
/// plus whatever its `w:pBdr` rules and their `w:space` gaps reserve. For an
/// empty flow paragraph retained after a page-anchored overlay, missing face
/// metrics fall back to its paragraph-mark/default font size and line spacing.
///
/// Shared with [`generate_stacked_hf_paragraphs`], which states it on the
/// paragraph's block, so the height a ruled story reserves in the band is the
/// height it actually lays out (issue #1824).
fn hf_paragraph_height_pt(
    paragraph: &crate::ir::HeaderFooterParagraph,
    default_text: Option<&TextStyle>,
    use_default_for_empty: bool,
) -> Option<f64> {
    let mut runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
    // A paragraph holding only a rule or anchored drawings is measured as the
    // empty line of its mark it is.
    let is_empty_flow_paragraph: bool =
        (runs.is_empty() || hf_paragraph_is_mark_line(paragraph)) && use_default_for_empty;
    if runs.is_empty() && use_default_for_empty {
        runs.push(hf_empty_paragraph_mark_run(paragraph, default_text));
    }
    let line_pt: f64 = text::word_line_advance_pt(&runs).or_else(|| {
        if !is_empty_flow_paragraph {
            return None;
        }
        // The overlay's final empty paragraph still owns a line box when
        // this platform cannot provide the face metrics.
        let font_size_pt: f64 = runs
            .iter()
            .filter_map(|run| run.style.font_size)
            .reduce(f64::max)
            .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);
        Some(match paragraph.style.line_spacing {
            Some(
                crate::ir::LineSpacing::Exact(height) | crate::ir::LineSpacing::AtLeast(height),
            ) => height.max(0.0),
            Some(crate::ir::LineSpacing::Proportional(multiplier)) => {
                (font_size_pt * multiplier).max(0.0)
            }
            None => font_size_pt,
        })
    })?;
    Some(line_pt + hf_paragraph_border_extent_pt(paragraph))
}

/// A run standing in for an empty paragraph's mark: its resolved style, or the
/// document default where the parser recorded none.
fn hf_empty_paragraph_mark_run(
    paragraph: &crate::ir::HeaderFooterParagraph,
    default_text: Option<&TextStyle>,
) -> Run {
    Run {
        text: "M".to_string(),
        style: paragraph
            .style
            .paragraph_mark_text_style
            .as_deref()
            .or(default_text)
            .cloned()
            .unwrap_or_default(),
        href: None,
        footnote: None,
        inline_box: None,
    }
}

/// What a header or footer paragraph's top and bottom `w:pBdr` rules and
/// their `w:space` gaps add to its lines, in points.
fn hf_paragraph_border_extent_pt(paragraph: &crate::ir::HeaderFooterParagraph) -> f64 {
    let Some(border) = paragraph.border.as_ref() else {
        return 0.0;
    };
    [
        (border.top.as_ref(), paragraph.border_space.map(|i| i.top)),
        (
            border.bottom.as_ref(),
            paragraph.border_space.map(|i| i.bottom),
        ),
    ]
    .into_iter()
    .filter_map(|(side, space)| {
        // An absent `w:space` is the schema's zero, which is what the rules
        // themselves are drawn with (issue #1824).
        side.map(|side| side.width + space.unwrap_or(0.0))
    })
    .sum()
}

#[path = "typst_gen_page_borders.rs"]
mod page_borders;
use page_borders::page_border_layer_markup;

#[path = "typst_gen_line_numbering.rs"]
mod line_numbering;
use line_numbering::write_line_numbering;

fn write_flow_page_setup(
    out: &mut String,
    page: &FlowPage,
    size: &PageSize,
    continuation_page: bool,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    ctx.active_page_size = Some(*size);
    ctx.active_page_margins = Some(page.margins);
    // A section may declare only a first-page story — `w:titlePg` with just a
    // `first` reference means pages after the first carry none — so the
    // shortcut has to ask about those too (issue #846), and about the first
    // stories of the continuous sections sharing this setup (issue #2254).
    let members: Vec<TitlePageRunMember> = ctx.flow_run_title_page_members.clone();
    if page.header.is_none()
        && page.footer.is_none()
        && page.first_header.is_none()
        && page.first_footer.is_none()
        && page.page_borders.is_none()
        && members.is_empty()
    {
        write_page_setup(out, size, &page.margins);
        ctx.active_body_top_pt = Some(page.margins.top);
        for (index, layer) in ["background", "foreground"].iter().enumerate() {
            if ctx.flow_border_layers[index] {
                let _ = writeln!(out, "#set page({layer}: none)");
            }
        }
        ctx.flow_border_layers = [false; 2];
        return Ok(());
    }

    // An overflowing header or a page-anchored overlay on a continuation page
    // grows the margin instead of overprinting the body (issues #736, #1995).
    let top_margin_pt: f64 = flow_page_top_margin_pt(
        page,
        size,
        ctx.document_default_text.as_ref(),
        continuation_page,
    );
    ctx.active_body_top_pt = Some(top_margin_pt);
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
    // the section's own margins, so they are built the same way. The choice
    // follows the declared first story even when it has no text: page one
    // then shows nothing, not the default story (issue #2224). A continuous
    // title-page section sharing this setup adds its own first page to the
    // choice (issue #2254).
    // TODO(#2373): its default stories and margins still come from this one.
    let own_first_page: String = format!(
        "here().page() == {}",
        section_first_page_expression(ctx.flow_section_index)
    );
    let default_header = page
        .header
        .as_ref()
        .filter(|header| hf_has_flow_content(header));
    let first_headers = first_page_story_choices(
        (&own_first_page, page.first_header.as_ref()),
        &members,
        |member| member.first_header.as_ref(),
        top_margin_pt,
    );
    if default_header.is_some() || first_headers.iter().any(|(_, story)| story.is_some()) {
        let default_value = default_header
            .map(|header| flow_header_value(header, page, size, top_margin_pt, ctx))
            .unwrap_or_default();
        let first_values: Vec<(String, FlowHeaderValue)> = first_headers
            .iter()
            .map(|(condition, header)| {
                let value = header
                    .map(|header| flow_header_value(header, page, size, top_margin_pt, ctx))
                    .unwrap_or_default();
                (condition.clone(), value)
            })
            .collect();
        // The ascent is a page property, so it is set when any story wants
        // it; a story that does not is unaffected by it.
        if default_value.wants_zero_ascent
            || first_values
                .iter()
                .any(|(_, value)| value.wants_zero_ascent)
        {
            out.push_str(", header-ascent: 0pt");
        }
        if first_values.is_empty() {
            let _ = write!(out, ", header: {}", default_value.markup);
        } else {
            let branches: Vec<(String, String)> = first_values
                .into_iter()
                .map(|(condition, value)| (condition, value.markup))
                .collect();
            let _ = write!(
                out,
                ", header: {}",
                per_page_story_choice(&branches, &default_value.markup, "{ ", " }")
            );
        }
    }

    // The footer takes the same first-page choice the header does (issue #846).
    let default_footer = page
        .footer
        .as_ref()
        .filter(|footer| hf_has_flow_content(footer));
    let first_footers = first_page_story_choices(
        (&own_first_page, page.first_footer.as_ref()),
        &members,
        |member| member.first_footer.as_ref(),
        top_margin_pt,
    );
    if default_footer.is_some() || first_footers.iter().any(|(_, story)| story.is_some()) {
        let default_value = default_footer
            .map(|footer| flow_footer_value(footer, page, ctx))
            .unwrap_or_default();
        let first_values: Vec<(String, FlowFooterValue)> = first_footers
            .iter()
            .map(|(condition, footer)| {
                let value = footer
                    .map(|footer| flow_footer_value(footer, page, ctx))
                    .unwrap_or_default();
                (condition.clone(), value)
            })
            .collect();
        if default_value.wants_zero_descent
            || first_values
                .iter()
                .any(|(_, value)| value.wants_zero_descent)
        {
            out.push_str(", footer-descent: 0pt");
        }
        if first_values.is_empty() {
            let _ = write!(out, ", footer: {}", default_value.markup);
        } else {
            let branches: Vec<(String, String)> = first_values
                .into_iter()
                .map(|(condition, value)| (condition, value.markup))
                .collect();
            let _ = write!(
                out,
                ", footer: {}",
                per_page_story_choice(&branches, &default_value.markup, "{ ", " }")
            );
        }
    }

    // `<wp:anchor behindDoc="1">` puts the shape under the page's own content,
    // which is where a decorative banner belongs — over it, the banner would
    // hide the body text it sits behind (issue #961).
    for (layer_index, (layer, behind_text)) in [("background", true), ("foreground", false)]
        .into_iter()
        .enumerate()
    {
        // `<w:titlePg/>` gives page one its own stories, and on this invoice
        // they are the only ones carrying a banner, so the layers are gated
        // per page exactly as the flow header and footer are (issue #961).
        let mut default_markup: String = page_anchored_layer_markup(
            page.header.as_ref(),
            page.footer.as_ref(),
            page,
            size,
            behind_text,
            ctx,
        )?;
        // The continuous title-page sections sharing this setup draw their
        // first stories' artwork on their own first pages (issue #2254).
        let first_stated: bool = page.first_header.is_some() || page.first_footer.is_some();
        let mut first_branches: Vec<(String, String)> = Vec::new();
        if first_stated {
            first_branches.push((
                own_first_page.clone(),
                page_anchored_layer_markup(
                    page.first_header.as_ref().or(page.header.as_ref()),
                    page.first_footer.as_ref().or(page.footer.as_ref()),
                    page,
                    size,
                    behind_text,
                    ctx,
                )?,
            ));
        }
        for member in &members {
            first_branches.push((
                continuous_section_first_page_condition(member.section_index, top_margin_pt),
                page_anchored_layer_markup(
                    member.first_header.as_ref().or(member.header.as_ref()),
                    member.first_footer.as_ref().or(member.footer.as_ref()),
                    page,
                    size,
                    behind_text,
                    ctx,
                )?,
            ));
        }
        let border_markup =
            page_border_layer_markup(page, size, behind_text, ctx.flow_section_index);
        for markup in std::iter::once(&mut default_markup)
            .chain(first_branches.iter_mut().map(|(_, markup)| markup))
        {
            if behind_text {
                // Back borders also sit below page-anchored header/footer artwork.
                markup.insert_str(0, &border_markup);
            } else {
                markup.push_str(&border_markup);
            }
        }
        let had_border: bool = ctx.flow_border_layers[layer_index];
        ctx.flow_border_layers[layer_index] = !border_markup.is_empty();
        if default_markup.is_empty() && first_branches.iter().all(|(_, markup)| markup.is_empty()) {
            if had_border {
                let _ = write!(out, ", {layer}: none");
            }
            continue;
        }
        if first_branches.is_empty() {
            let _ = write!(out, ", {layer}: [{default_markup}]");
        } else {
            let _ = write!(
                out,
                ", {layer}: {}",
                per_page_story_choice(&first_branches, &default_markup, "[", "]")
            );
        }
    }

    out.push_str(")\n");
    Ok(())
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
    !paragraph.elements.is_empty()
        || paragraph.border.is_some()
        || hf_paragraph_is_mark_line(paragraph)
}

/// Whether a header or footer paragraph with no runs is still a line of its
/// paragraph mark: one whose only content is a horizontal rule (issue #2251)
/// or drawings the story draws at their anchors (issue #2327).
fn hf_paragraph_is_mark_line(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    paragraph.elements.is_empty()
        && (paragraph.style.horizontal_rule.is_some() || paragraph.style.holds_anchored_drawing)
}

/// The horizontal rule a header or footer paragraph holds as its only
/// content, which it draws on a line of its paragraph mark (issue #2251).
fn hf_paragraph_rule(
    paragraph: &crate::ir::HeaderFooterParagraph,
) -> Option<&crate::ir::HorizontalRule> {
    paragraph
        .style
        .horizontal_rule
        .as_deref()
        .filter(|_| paragraph.elements.is_empty())
}

fn hf_paragraph_has_flow_content(paragraph: &crate::ir::HeaderFooterParagraph) -> bool {
    hf_paragraph_has_content(paragraph)
        && paragraph
            .frame
            .as_ref()
            .is_none_or(|frame| !is_page_anchored_frame(frame))
}

/// One page layer's markup: the page-anchored elements both stories put on it,
/// plus their framed paragraphs below the body in the background.
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
) -> Result<String, ConvertError> {
    // The stories' tables carry their page fields as markers in runs
    // (issue #2338).
    with_page_number_format(ctx.page_number_format, || {
        let mut markup = String::new();
        for (hf, is_footer) in [(header, false), (footer, true)] {
            let Some(hf) = hf else {
                continue;
            };
            generate_page_anchored_hf_elements(
                &mut markup,
                hf,
                page,
                size,
                behind_text,
                is_footer,
                ctx,
            )?;
            // Word paints the main story after header/footer framed paragraphs, so
            // body drawings can cover those frames. Typst's page background is the
            // layer that preserves that cross-story order (issue #1408).
            if behind_text {
                generate_page_anchored_hf_frames(&mut markup, hf, size, page.margins.right, ctx);
            }
        }
        Ok(markup)
    })
}

/// Place header/footer elements in the page overlay. Paragraph-relative
/// drawing shapes use their anchor paragraph's position; page-relative shapes
/// use their stored extent, tables use converted dimensions and frame anchors,
/// and pictures sit at their resolved frame position.
fn generate_page_anchored_hf_elements(
    out: &mut String,
    hf: &HeaderFooter,
    page: &FlowPage,
    page_size: &PageSize,
    behind_text: bool,
    is_footer: bool,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    for element in &hf.shapes {
        match &element.content {
            HeaderFooterShapeContent::Shape(shape) => {
                if element.behind_text != behind_text {
                    continue;
                }
                if element.anchor_paragraph_index.is_some() {
                    let Some((x, y)) = page_relative_hf_paragraph_shape_position(
                        hf, element, page, page_size, is_footer,
                    ) else {
                        continue;
                    };
                    generate_page_anchored_hf_shape(out, element, shape, x, y, ctx);
                    continue;
                }
                if !is_page_anchored_frame(&element.frame) {
                    continue;
                }
                let x: f64 = element.frame.x.unwrap_or_else(|| {
                    aligned_offset(
                        element.frame.horizontal_align,
                        page_size.width,
                        Some(element.width),
                    )
                });
                let y: f64 = element.frame.y.unwrap_or_else(|| {
                    aligned_offset(
                        element.frame.vertical_align,
                        page_size.height,
                        Some(element.height),
                    )
                });
                // A box keeps transforms centred on the shape's own extent,
                // including artwork that hangs past the page edge.
                generate_page_anchored_hf_shape(out, element, shape, x, y, ctx);
            }
            HeaderFooterShapeContent::Table(table) => {
                if element.behind_text != behind_text {
                    continue;
                }
                let x: f64 = page_relative_hf_table_x(element, page, page_size);
                let y: f64 = page_relative_hf_table_y(element, hf, page, page_size, is_footer);
                let _ = write!(
                    out,
                    "#place(top + left, dx: {}pt, dy: {}pt)[",
                    format_f64(x),
                    format_f64(y)
                );
                let previous_preference: bool =
                    ctx.prefer_sibling_metrics_for_empty_cell_paragraphs;
                ctx.prefer_sibling_metrics_for_empty_cell_paragraphs = true;
                let result: Result<(), ConvertError> = generate_table(out, table, ctx);
                ctx.prefer_sibling_metrics_for_empty_cell_paragraphs = previous_preference;
                result?;
                out.push(']');
            }
            HeaderFooterShapeContent::Picture(picture) => {
                if element.behind_text != behind_text {
                    continue;
                }
                let x: f64 = page_relative_hf_table_x(element, page, page_size);
                let y: f64 = page_relative_hf_table_y(element, hf, page, page_size, is_footer);
                let _ = write!(
                    out,
                    "#place(top + left, dx: {}pt, dy: {}pt)[",
                    format_f64(x),
                    format_f64(y)
                );
                generate_image(out, &picture.image, ctx);
                out.push_str("]\n");
            }
        }
    }
    Ok(())
}

fn page_relative_hf_table_x(
    element: &crate::ir::HeaderFooterShape,
    page: &FlowPage,
    page_size: &PageSize,
) -> f64 {
    let frame: &HeaderFooterFrame = &element.frame;
    let (origin, available): (f64, f64) = match frame.horizontal_anchor {
        FrameAnchor::Page => (0.0, page_size.width),
        FrameAnchor::Margin | FrameAnchor::Text => (
            page.margins.left,
            (page_size.width - page.margins.left - page.margins.right).max(0.0),
        ),
    };
    origin
        + frame.x.unwrap_or_else(|| {
            aligned_offset(
                frame.horizontal_align,
                available,
                frame.width.or(Some(element.width)),
            )
        })
}

fn page_relative_hf_table_y(
    element: &crate::ir::HeaderFooterShape,
    hf: &HeaderFooter,
    page: &FlowPage,
    page_size: &PageSize,
    is_footer: bool,
) -> f64 {
    let frame: &HeaderFooterFrame = &element.frame;
    let (origin, available): (f64, f64) = match frame.vertical_anchor {
        FrameAnchor::Page => (0.0, page_size.height),
        FrameAnchor::Margin | FrameAnchor::Text => (
            page.margins.top,
            (page_size.height - page.margins.top - page.margins.bottom).max(0.0),
        ),
    };
    if let Some(offset) = frame.y {
        return origin + offset;
    }
    if let Some(alignment) = frame.vertical_align {
        return origin + aligned_offset(Some(alignment), available, frame.height);
    }
    match frame.vertical_anchor {
        FrameAnchor::Page => 0.0,
        FrameAnchor::Margin | FrameAnchor::Text if is_footer => {
            page_size.height - hf.distance_from_edge.unwrap_or(0.0) - element.height
        }
        FrameAnchor::Margin | FrameAnchor::Text => hf.distance_from_edge.unwrap_or(0.0),
    }
}

/// `bottom_edge_em` is how far below the baseline the enclosing band set the
/// text's bottom edge, in em; zero where it kept the baseline.
fn generate_flow_hf_content(
    out: &mut String,
    hf: &HeaderFooter,
    bottom_edge_em: f64,
    ctx: &mut GenCtx,
) {
    // `par(leading:)` is what separates a paragraph's lines. Typst advances a
    // line by `top-edge + bottom-edge + leading`, and the story set no leading,
    // so it took the 0.65em default: an 8pt Arial header advanced 10.9305pt
    // where Word advances 9.1992pt. Stating the remainder after the edges
    // leaves the first baseline — seated against the top edge by #629 —
    // exactly where it was, and corrects only the advance between lines (#735).
    // A pinned footer band's descender edge is one of those edges: a leading
    // that ignored it advanced two 11pt Arial footer lines 14.98pt against
    // Word's 12.65pt (issue #2326).
    //
    // Word advances to the next paragraph's line by the same pitch, so the
    // story's paragraphs are separate Typst paragraphs whose `par(spacing:)`
    // is that leading too. Joining them with `\` line breaks into one Typst
    // paragraph did not survive a stated `w:jc`, a stop-placed tab, an indent
    // or a right `w:ptab`: each wraps its paragraph in a block-level element,
    // which split the story around it, and the line break became an empty
    // paragraph with the default paragraph gap on both sides. Two `jc="both"`
    // tabbed footer lines Word draws 13.52pt apart came out 46.6pt apart, the
    // upper one over the last body line (issue #2222). A block that states no
    // spacing of its own takes `par(spacing:)`, so the leading separates every
    // paragraph whatever wraps it.
    //
    // A ruled paragraph also reserves its rules and their `w:space` gaps,
    // which no leading expresses. When the story was still joined by line
    // breaks, the second line of a ruled header landed 20pt low, over the
    // first body line, and the taller measured story clamped the band shift
    // away and lifted the first line 2.49pt (issue #1824). Such a story is
    // laid out as the stack of blocks it already is, each carrying Word's own
    // line box, so the flow advances by Word's line and needs no shift to
    // seat it.
    if hf_story_rules_a_paragraph(hf) {
        generate_stacked_hf_paragraphs(out, hf, ctx);
        return;
    }
    let mut emitted = hf
        .paragraphs
        .iter()
        .filter(|paragraph| hf_paragraph_is_emitted(paragraph));
    let has_several_paragraphs: bool = emitted.clone().nth(1).is_some();
    let leading: Option<f64> = emitted.next().and_then(|paragraph| {
        let runs: Vec<Run> = hf_paragraph_metric_runs(paragraph);
        text::word_hf_line_leading_pt(&runs, bottom_edge_em)
    });
    if let Some(leading) = leading {
        let _ = writeln!(out, "#set par(leading: {}pt)", format_f64(leading));
    }
    if has_several_paragraphs {
        // Typst's own 0.65em leading where the face's metrics are unknown.
        let spacing: String = leading.map_or_else(
            || "0.65em".to_string(),
            |leading| format!("{}pt", format_f64(leading)),
        );
        let _ = writeln!(out, "#set par(spacing: {spacing})");
        // A story taller than its band overflows it, as Word lets it, but a
        // breakable block would rather shrink the gap above it to fit: two
        // indented 11pt Arial footer lines in a 21.6pt band advanced only
        // 11.70pt (issue #2222).
        out.push_str("#set block(breakable: false)\n");
    }
    let mut is_first: bool = true;
    for paragraph in &hf.paragraphs {
        if !hf_paragraph_is_emitted(paragraph) {
            continue;
        }
        if !is_first {
            out.push_str("\n#parbreak()\n");
        }
        generate_hf_styled_paragraph(out, paragraph, ctx);
        is_first = false;
    }
}

fn generate_page_anchored_hf_shape(
    out: &mut String,
    element: &crate::ir::HeaderFooterShape,
    shape: &crate::ir::Shape,
    x: f64,
    y: f64,
    ctx: &mut GenCtx,
) {
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[#box(width: {}pt, height: {}pt)[",
        format_f64(x),
        format_f64(y),
        format_f64(element.width),
        format_f64(element.height)
    );
    generate_shape(out, shape, element.width, element.height, ctx);
    out.push_str("]]\n");
}

fn page_relative_hf_paragraph_shape_position(
    hf: &HeaderFooter,
    element: &crate::ir::HeaderFooterShape,
    page: &FlowPage,
    page_size: &PageSize,
    is_footer: bool,
) -> Option<(f64, f64)> {
    let anchor_paragraph_index: usize = element.anchor_paragraph_index?;
    let frame: &HeaderFooterFrame = &element.frame;
    let horizontal_available: f64 =
        (page_size.width - page.margins.left - page.margins.right).max(0.0);
    let horizontal_origin: f64 = match frame.horizontal_anchor {
        FrameAnchor::Page => 0.0,
        FrameAnchor::Margin | FrameAnchor::Text => page.margins.left,
    };
    let x: f64 = horizontal_origin
        + frame.x.unwrap_or_else(|| {
            aligned_offset(
                frame.horizontal_align,
                match frame.horizontal_anchor {
                    FrameAnchor::Page => page_size.width,
                    FrameAnchor::Margin | FrameAnchor::Text => horizontal_available,
                },
                frame.width.or(Some(element.width)),
            )
        });

    let mut story_height: f64 = 0.0;
    let mut preceding_height: f64 = 0.0;
    for (paragraph_index, paragraph) in hf.paragraphs.iter().enumerate() {
        if paragraph.frame.as_ref().is_some_and(is_page_anchored_frame) {
            continue;
        }
        let paragraph_height: f64 = hf_paragraph_story_height_pt(paragraph);
        let paragraph_before: f64 = paragraph.style.space_before.unwrap_or(0.0).max(0.0);
        let paragraph_after: f64 = paragraph.style.space_after.unwrap_or(0.0).max(0.0);
        story_height += paragraph_before + paragraph_height + paragraph_after;
        if paragraph_index < anchor_paragraph_index {
            preceding_height += paragraph_before + paragraph_height + paragraph_after;
        }
    }
    let paragraph: &crate::ir::HeaderFooterParagraph = hf.paragraphs.get(anchor_paragraph_index)?;
    let paragraph_height: f64 = hf_paragraph_story_height_pt(paragraph);
    let vertical_available: f64 = paragraph_height;
    let vertical_offset: f64 = frame.y.unwrap_or_else(|| {
        aligned_offset(
            frame.vertical_align,
            vertical_available,
            frame.height.or(Some(element.height)),
        )
    });
    let story_origin_y: f64 = if is_footer {
        page_size.height - hf.distance_from_edge.unwrap_or(0.0) - story_height
    } else {
        hf.distance_from_edge.unwrap_or(0.0)
    };
    let y: f64 = story_origin_y + preceding_height + vertical_offset;
    Some((x, y))
}

fn hf_paragraph_story_height_pt(paragraph: &crate::ir::HeaderFooterParagraph) -> f64 {
    if let Some(height) = hf_paragraph_height_pt(paragraph, None, false) {
        return height;
    }
    match paragraph.style.line_spacing {
        Some(crate::ir::LineSpacing::Exact(height) | crate::ir::LineSpacing::AtLeast(height)) => {
            height.max(0.0)
        }
        Some(crate::ir::LineSpacing::Proportional(multiplier)) => (11.0 * multiplier).max(0.0),
        None => 11.0,
    }
}

/// Whether any paragraph this story emits rules itself off with a `w:pBdr`.
fn hf_story_rules_a_paragraph(hf: &HeaderFooter) -> bool {
    hf.paragraphs
        .iter()
        .filter(|paragraph| hf_paragraph_is_emitted(paragraph))
        .any(|paragraph| paragraph.border.is_some())
}

/// Emit the story as one block per paragraph, each as tall as Word's line.
///
/// Word seats a ruled paragraph's own line exactly where an unruled one sits,
/// hangs the rule `w:pBdr w:space` below the line's bottom edge, and starts
/// the next paragraph under the rule, so the story advances by the line plus
/// whatever the rules and their gaps reserve. Measured on native Word 16.113.1
/// at `w:space` 0, 1 and 8, and with a `w:bottom` stating none (issue #1824).
///
/// The height is stated rather than left to the content because the text cell
/// carries Typst's cap-height top edge, which is shorter than Word's line by
/// the seat the band shift already corrects for the story's first baseline.
/// Stating it keeps every baseline exactly where the flowed story form puts
/// it — [`hf_content_height_pt`] measures the band against these same terms —
/// and only moves what the rule adds.
fn generate_stacked_hf_paragraphs(out: &mut String, hf: &HeaderFooter, ctx: &mut GenCtx) {
    for paragraph in &hf.paragraphs {
        if !hf_paragraph_is_emitted(paragraph) {
            continue;
        }
        out.push_str("#block(width: 100%, above: 0pt, below: 0pt");
        if let Some(height) = hf_paragraph_height_pt(paragraph, None, false) {
            let _ = write!(out, ", height: {}pt", format_f64(height));
        }
        out.push_str(")[");
        generate_hf_styled_paragraph(out, paragraph, ctx);
        out.push_str("]\n");
    }
}

/// Write a header story's content, as Word's own line boxes where it has an
/// exact line.
fn generate_flow_header_content(out: &mut String, header: &HeaderFooter, ctx: &mut GenCtx) {
    if header_story_draws_line_boxes(header) {
        generate_line_boxed_header_paragraphs(out, header, ctx);
    } else {
        generate_flow_hf_content(out, header, 0.0, ctx);
    }
}

/// Whether a header story is drawn as Word's own line boxes.
///
/// Only a story with an exact line needs them, and only when every paragraph
/// it emits has a box to give: a natural line whose face does not resolve
/// keeps the flowed story's estimates. A ruled story keeps its stacked form,
/// which seats each rule against the face's natural line (issue #1824).
fn header_story_draws_line_boxes(header: &HeaderFooter) -> bool {
    let mut emitted = header
        .paragraphs
        .iter()
        .filter(|paragraph| hf_paragraph_is_emitted(paragraph));
    !hf_story_rules_a_paragraph(header)
        && emitted
            .clone()
            .any(|paragraph| hf_paragraph_exact_line_pt(paragraph).is_some())
        && emitted.all(|paragraph| hf_paragraph_line_box_pt(paragraph).is_some())
}

/// The ascent and descent, in points, Word gives each line of a header
/// paragraph: its exact line where it states one, otherwise its face's
/// natural line.
fn hf_paragraph_line_box_pt(paragraph: &crate::ir::HeaderFooterParagraph) -> Option<(f64, f64)> {
    match hf_paragraph_exact_line_pt(paragraph) {
        Some(line_pt) => Some(text::word_exact_line_box_pt(line_pt)),
        None => text::word_header_line_box_pt(&hf_paragraph_metric_runs(paragraph)),
    }
}

/// Emit a header story as one block per paragraph, each line in Word's box.
///
/// The flowed story cannot draw an exact line: its one leading and the band
/// shift seat every line one face ascent down and advance it by the face's
/// natural line, so a 160pt exact line's text sat one face ascent below the
/// band top where Word seats it 128pt down (issue #2214). Stating each line's
/// ascent and descent as the text edges, with no leading, has Typst lay Word's
/// boxes out directly. The first baseline then sits its own ascent below the
/// band top, so the story takes no band shift, and each line advances by its
/// own height.
fn generate_line_boxed_header_paragraphs(
    out: &mut String,
    header: &HeaderFooter,
    ctx: &mut GenCtx,
) {
    for paragraph in &header.paragraphs {
        if !hf_paragraph_is_emitted(paragraph) {
            continue;
        }
        out.push_str("#block(width: 100%, above: 0pt, below: 0pt)[");
        if let Some((ascent_pt, descent_pt)) = hf_paragraph_line_box_pt(paragraph) {
            let _ = write!(
                out,
                "#set text(top-edge: {}pt, bottom-edge: -{}pt)\n#set par(leading: 0pt)\n",
                format_f64(ascent_pt),
                format_f64(descent_pt)
            );
        }
        ctx.hf_lines_are_word_boxes = true;
        generate_hf_styled_paragraph(out, paragraph, ctx);
        ctx.hf_lines_are_word_boxes = false;
        out.push_str("]\n");
    }
}

/// Whether [`generate_flow_hf_content`] writes this paragraph into the story.
///
/// Page-anchored frames are drawn separately in the page background, and an
/// empty paragraph carrying neither content, a border, a horizontal rule nor
/// an anchored drawing produces nothing.
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
/// line ascent. Also `None` for a story drawn as Word's line boxes, whose
/// first line seats its own baseline (issue #2214).
fn header_band_shift_pt(hf: &HeaderFooter) -> Option<f64> {
    if header_story_draws_line_boxes(hf) {
        return None;
    }
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
/// Positioned tabs contribute no glyphs. A paragraph whose only content is a
/// horizontal rule or anchored drawings is a line of its paragraph mark, so it
/// measures as an empty run in the mark's style (issues #2251 and #2327).
fn hf_paragraph_metric_runs(paragraph: &crate::ir::HeaderFooterParagraph) -> Vec<Run> {
    if hf_paragraph_is_mark_line(paragraph) {
        return vec![Run {
            text: String::new(),
            style: paragraph
                .style
                .paragraph_mark_text_style
                .as_deref()
                .cloned()
                .unwrap_or_default(),
            href: None,
            footnote: None,
            inline_box: None,
        }];
    }
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
                inline_box: None,
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
    generate_flow_header_content(out, header, ctx);
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
        generate_sheet_hf_content(out, header, size, &page.margins, ctx, true, None);
        out.push(']');
    }

    if let Some(footer) = &page.footer {
        if let Some(seat) = sheet_footer_seat(footer, page.margins.bottom) {
            out.push_str(", footer-descent: 0pt, footer: ");
            if hf_needs_context(footer) {
                out.push_str("context ");
            }
            let _ = write!(
                out,
                "block(width: 100%, height: {}pt)[#set text(bottom-edge: {}); #place(bottom, block(width: 100%)[",
                format_f64(seat.band_pt),
                seat.story_bottom_edge(),
            );
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx, false, Some(&seat));
            out.push_str("])]");
        } else if hf_needs_stack_offset(footer) {
            out.push_str(", footer: context { let footer_content = block(width: 100%)[");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx, false, None);
            out.push_str("]; move(dy: -measure(footer_content).height / 2)[#footer_content] }");
        } else if hf_needs_context(footer) {
            out.push_str(", footer: context [");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx, false, None);
            out.push(']');
        } else {
            out.push_str(", footer: [");
            generate_sheet_hf_content(out, footer, size, &page.margins, ctx, false, None);
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
    is_header: bool,
    seat: Option<&SheetFooterSeat>,
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
    if !is_header || !generate_multiline_sheet_text(out, hf, ctx, None) {
        generate_hf_content(out, hf, ctx, seat);
    }
    if scaled_box.is_some() {
        out.push_str("]]");
    }
}

/// What Excel leaves between `<pageMargins>/@footer` and the bottom of the
/// footer text's line box, in sheet points.
///
/// Measured on Excel-for-Mac exports of one-factor variants of
/// `tests/fixtures/xlsx/headerFooterTest.xlsx`. A 12pt Calibri footer over a
/// 0.5in (36pt) footer margin puts its baseline 41pt above the page's bottom
/// edge; Calibri's `hhea` descender is 0.26855em, so 36 + 2 + 3.22 = 41.22
/// lands on the whole point Excel prints it at. The same 2pt holds across the
/// series 6, 8, 12, 14, 20, 40 and 80pt, across Arial, Verdana, Times New
/// Roman, Aptos and Segoe UI, and at footer margins of 0.3, 0.5, 0.75 and
/// 1.0in (issue #1142).
const SHEET_FOOTER_BAND_INSET_PT: f64 = 2.0;

/// How much lower Excel's rich-text path seats a footer section than its
/// plain path, in sheet points (issue #1552).
///
/// On native exports of `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
/// every section that splits its text — the 1pt `#` marker run ahead of the
/// Sensitivity label, a face switch mid-sentence, or a `_x000D_` ahead of
/// otherwise uniform text — prints one whole point below, and one point
/// further left than, the same text as one uniform run: Aptos 8pt at 24pt
/// against 25pt, across 10, 16, 20, 24 and 40pt and Arial and Calibri alike,
/// and per section, a rich left label beside a plain right one landing one
/// point apart on the same page. The horizontal point is
/// [`SHEET_PLAIN_LEFT_SECTION_INSET_PT`].
const SHEET_RICH_SECTION_DROP_PT: f64 = 1.0;

/// How far right of the floored left margin Excel's plain path starts a left
/// header/footer section, in sheet points; the rich-text path starts on the
/// margin itself (issue #1728).
///
/// Native Excel-for-Mac exports put `&L&"Calibri,Regular"&K000000top left`
/// and `bottom left` of `tests/fixtures/xlsx/headerFooterTest.xlsx` at x 55
/// over its 54pt margin, and the budget workbook's Aptos 8pt Sensitivity
/// label at 51 over a 50.4pt margin — with a colour code ahead of the text as
/// well — but at 50 behind the 1pt `#` marker run or a `_x000D_`. Centred and
/// right sections are not shifted: the centre reads 271.000 natively against
/// our 270.923, and the right section's start is the glyph advances of #1659.
/// Archived native `1552-footer-batch6-scaled` probes also confirm the inset
/// at an effective 78% print scale: plain Aptos 8/24pt sections start at
/// 18.72pt, versus the rich section's 17.939999pt. The 8pt face prints at
/// 6.24pt; the no-patch control matches geometry and all decoded page pixels.
const SHEET_PLAIN_LEFT_SECTION_INSET_PT: f64 = 1.0;

/// Write one paragraph of a worksheet header/footer section, starting a
/// plain left section [`SHEET_PLAIN_LEFT_SECTION_INSET_PT`] sheet points
/// right of the margin. `#move` leaves the section's layout alone, so a long
/// section wraps against the same width it always had.
fn generate_sheet_hf_section(
    out: &mut String,
    paragraph: &crate::ir::HeaderFooterParagraph,
    ctx: &mut GenCtx,
    scale: f64,
) {
    let starts_inset: bool =
        paragraph.style.alignment == Some(Alignment::Left) && !paragraph.sheet_section_is_rich;
    if starts_inset {
        let _ = write!(
            out,
            "#move(dx: {}pt)[",
            format_f64(round_to_hundredths(
                SHEET_PLAIN_LEFT_SECTION_INSET_PT * scale
            ))
        );
    }
    generate_hf_styled_paragraph(out, paragraph, ctx);
    if starts_inset {
        out.push(']');
    }
}

/// Where a seated sheet footer's block sits and where each of its sections'
/// baselines land (issues #1142, #1552).
struct SheetFooterSeat {
    /// Height of the band block, from the bottom margin line down to the
    /// band Excel seats the footer on, in paper points.
    band_pt: f64,
    /// Per paragraph, how far above the band's bottom its baseline sits, in
    /// paper points; `None` where no run's face has metrics.
    section_lift_pt: Vec<Option<f64>>,
}

impl SheetFooterSeat {
    /// The `bottom-edge` Typst value that puts a section's baseline `lift`
    /// points above the band: a negative edge reaches that far below the
    /// baseline, and `#place(bottom)` rests it on the band.
    fn bottom_edge_value(lift_pt: f64) -> String {
        format!("-{}pt", format_f64(lift_pt))
    }

    /// The story-wide `bottom-edge`: the first section that resolved, so a
    /// section whose face has no metrics still sits near its neighbours, and
    /// Typst's normalised descender only when none did.
    fn story_bottom_edge(&self) -> String {
        self.section_lift_pt
            .iter()
            .flatten()
            .next()
            .map(|lift_pt| Self::bottom_edge_value(*lift_pt))
            .unwrap_or_else(|| "\"descender\"".to_string())
    }

    /// The `bottom-edge` the paragraph at `index` states for itself.
    fn section_bottom_edge(&self, index: usize) -> Option<String> {
        self.section_lift_pt
            .get(index)
            .copied()
            .flatten()
            .map(Self::bottom_edge_value)
    }
}

/// Seat a sheet footer where Excel prints it (issues #1142, #1552).
///
/// Excel lays the footer out in the sheet's own coordinate space — the fit
/// scale multiplies every position afterwards, as it does the horizontal box
/// of issue #1510 — and works in whole points there:
///
/// 1. the footer margin floors to a whole sheet point, plus the 2pt band;
/// 2. the section's deepest run adds its `hhea` descent at the sheet size;
/// 3. the sum rounds to a whole point;
/// 4. a section on the rich-text path drops one more point.
///
/// Native exports of the reported workbook's 8pt label: 23 + 2.25 = 25.25
/// rounds to 25, and the rich label prints at 24pt above the Letter page's
/// bottom edge; the same sheet fitted at 0.78 floors 21.6 / 0.78 = 27.69 to
/// 27, seats 29 + 2.25 on 31, drops to 30 sheet points and prints at
/// 30 x 0.78 = 23.40pt above the A3 paper. Both pages had been drawn from a
/// paper-floored 23pt band with the *first* run's face — the 1pt `#`
/// marker's Trebuchet MS — scaled to the 8pt Aptos run, 0.78pt and 1.50pt
/// above where Excel prints them.
///
/// `footer-descent: 0pt` pins Typst's footer origin on the bottom margin line
/// and a block spanning the band's remainder ends where Excel's band does;
/// each section then states its own `bottom-edge`. `None` leaves the story on
/// Typst's own descent: either the margin is unknown, or it reaches past the
/// bottom margin and there is no band to span.
fn sheet_footer_seat(footer: &HeaderFooter, bottom_margin_pt: f64) -> Option<SheetFooterSeat> {
    let footer_margin_pt: f64 = footer.distance_from_edge?;
    let scale: f64 = footer
        .sheet_print_scale
        .filter(|scale| *scale > 0.0)
        .unwrap_or(1.0);
    // The quotient of two decimals is not exact in binary; a margin that is
    // a whole number of sheet points must not floor to the one below it.
    let band_sheet_pt: f64 = (footer_margin_pt / scale + 1e-9).floor() + SHEET_FOOTER_BAND_INSET_PT;
    // Keep float noise (54 - 23.000000000000004) out of the emitted source.
    let band_pt: f64 = round_to_hundredths(bottom_margin_pt - band_sheet_pt * scale);
    if band_pt <= 0.0 {
        return None;
    }
    let section_lift_pt: Vec<Option<f64>> = footer
        .paragraphs
        .iter()
        .map(|paragraph| {
            let deepest_descent_pt: f64 =
                text::sheet_line_deepest_descent_pt(&hf_paragraph_metric_runs(paragraph), scale)?;
            let drop_pt: f64 = if paragraph.sheet_section_is_rich {
                SHEET_RICH_SECTION_DROP_PT
            } else {
                0.0
            };
            let seat_sheet_pt: f64 = (band_sheet_pt + deepest_descent_pt).round() - drop_pt;
            Some(round_to_hundredths((seat_sheet_pt - band_sheet_pt) * scale))
        })
        .collect();
    Some(SheetFooterSeat {
        band_pt,
        section_lift_pt,
    })
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

/// Text-only sheet stories use native line boxes. Footers grow upward from
/// their last-line seats; headers grow downward from their first-line anchors.
/// Each line retains its own descent, including the footer rich-section adjustment.
fn generate_multiline_sheet_text(
    out: &mut String,
    hf: &HeaderFooter,
    ctx: &mut GenCtx,
    seat: Option<&SheetFooterSeat>,
) -> bool {
    // A picture can determine the line height; preserve the existing flow layout
    // rather than stacking an image-bearing line with text-only font metrics.
    if hf.paragraphs.iter().any(|paragraph| {
        paragraph
            .elements
            .iter()
            .any(|element| matches!(element, HFInline::Image(_)))
    }) {
        return false;
    }
    let slots = [Alignment::Left, Alignment::Center, Alignment::Right];
    if hf.paragraphs.iter().any(|paragraph| {
        !slots
            .iter()
            .any(|slot| paragraph.style.alignment == Some(*slot))
    }) {
        return false;
    }
    let groups: Vec<Vec<usize>> = slots
        .iter()
        .map(|slot| {
            hf.paragraphs
                .iter()
                .enumerate()
                .filter_map(|(index, paragraph)| {
                    (paragraph.style.alignment == Some(*slot)).then_some(index)
                })
                .collect()
        })
        .collect();
    if !groups.iter().any(|group| group.len() > 1) {
        return false;
    }
    let scale: f64 = hf
        .sheet_print_scale
        .filter(|scale| *scale > 0.0)
        .unwrap_or(1.0);
    let advances: Option<Vec<f64>> = hf
        .paragraphs
        .iter()
        .map(|paragraph| {
            hf_paragraph_metric_runs(paragraph)
                .iter()
                .map(|run| {
                    let family = run.style.font_family.as_deref()?;
                    let size: f64 = run.style.font_size.unwrap_or(11.0) / scale;
                    text::sheet_wrapped_line_advance_pt(family, size)
                        .or_else(|| text::sheet_row_line_advance_pt(family, size, false))
                        .map(|advance| advance * scale)
                })
                .collect::<Option<Vec<f64>>>()?
                .into_iter()
                .reduce(f64::max)
        })
        .collect();
    let Some(advances) = advances else {
        return false;
    };
    let section_lifts: Vec<Option<f64>> = match seat {
        Some(seat) => seat.section_lift_pt.clone(),
        // Native mixed-size header probes round descents in sheet points
        // before applying the print scale.
        None => hf
            .paragraphs
            .iter()
            .map(|paragraph| {
                text::sheet_line_deepest_descent_pt(&hf_paragraph_metric_runs(paragraph), scale)
                    .map(|descent| descent.round() * scale)
            })
            .collect(),
    };
    let multiple_sections: bool = groups.iter().filter(|group| !group.is_empty()).count() > 1;
    if multiple_sections {
        let edge = if seat.is_some() { "bottom" } else { "top" };
        let _ = write!(
            out,
            "#grid(columns: (1fr, 1fr, 1fr), align: (left + {edge}, center + {edge}, right + {edge}), "
        );
    }
    for group in &groups {
        if multiple_sections {
            out.push_str("[#box(width: 300%)[");
        }
        if !group.is_empty() {
            out.push_str("#layout(size => { ");
            for &index in group {
                let _ = write!(out, "let line_{index} = block(width: size.width)[");
                if let Some(lift_pt) = section_lifts[index] {
                    // Native single-line header sections sit one sheet point
                    // below multiline sections with the same font metrics.
                    let lift_pt = if seat.is_none() && group.len() == 1 {
                        lift_pt - scale
                    } else {
                        lift_pt
                    };
                    let edge = SheetFooterSeat::bottom_edge_value(lift_pt);
                    // The signed bottom edge contributes lift below the baseline.
                    let _ = write!(
                        out,
                        "#set text(top-edge: {}pt, bottom-edge: {edge}); #set par(leading: 0pt); ",
                        format_f64(advances[index] - lift_pt)
                    );
                }
                generate_sheet_hf_section(out, &hf.paragraphs[index], ctx, scale);
                out.push_str("]; ");
            }
            if seat.is_none() {
                let first_index = group[0];
                let _ = write!(out, "let offset_{first_index} = 0pt; ");
                for pair in group.windows(2) {
                    let previous = pair[0];
                    let index = pair[1];
                    let _ = write!(
                        out,
                        "let offset_{index} = offset_{previous} + calc.max({}pt, measure(line_{previous}).height); ",
                        format_f64(advances[previous])
                    );
                }
                // Reserve the first line's box so adding lines does not move
                // the first header upward through the page's top boundary.
                let _ = write!(
                    out,
                    "block(width: size.width, height: {}pt)[",
                    format_f64(advances[first_index])
                );
                for &index in group {
                    let _ = write!(out, "#place(top, dy: offset_{index}, line_{index})");
                }
            } else {
                let last_index = group[group.len() - 1];
                let _ = write!(out, "let offset_{last_index} = 0pt; ");
                for position in (0..group.len() - 1).rev() {
                    let index = group[position];
                    let next = group[position + 1];
                    // A wrapped paragraph contributes every visual line; even an
                    // empty explicit paragraph still reserves one native advance.
                    let _ = write!(
                        out,
                        "let offset_{index} = offset_{next} + calc.max({}pt, measure(line_{next}).height); ",
                        format_f64(advances[next])
                    );
                }
                out.push_str("block(width: size.width, height: calc.max(");
                for &index in group {
                    let _ = write!(out, "measure(line_{index}).height + offset_{index}, ");
                }
                out.push_str("))[");
                for &index in group {
                    let _ = write!(out, "#place(bottom, dy: -offset_{index}, line_{index})");
                }
            }
            out.push_str("] })");
        }
        if multiple_sections {
            out.push_str("]], ");
        }
    }
    if multiple_sections {
        out.push(')');
    }
    true
}

/// Generate inline content for a header or footer.
///
/// `seat` is a seated sheet footer's per-section `bottom-edge`: Excel decides
/// the plain or rich-text seat per section, so each grid cell states its own
/// edge and the row aligns its cells' bottoms on the band (issue #1552).
fn generate_hf_content(
    out: &mut String,
    hf: &HeaderFooter,
    ctx: &mut GenCtx,
    seat: Option<&SheetFooterSeat>,
) {
    if let Some(seat) = seat
        && generate_multiline_sheet_text(out, hf, ctx, Some(seat))
    {
        return;
    }
    let section_bottom_edge =
        |index: usize| -> Option<String> { seat.and_then(|seat| seat.section_bottom_edge(index)) };
    let scale: f64 = hf
        .sheet_print_scale
        .filter(|scale| *scale > 0.0)
        .unwrap_or(1.0);
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
        if seat.is_some() {
            out.push_str("align: (left + bottom, center + bottom, right + bottom), ");
        } else {
            out.push_str("align: (left, center, right), ");
        }
        for slot in [Alignment::Left, Alignment::Center, Alignment::Right] {
            // Excel sections share the page width; the grid only anchors them.
            out.push_str("[#box(width: 300%)[");
            if let Some((index, para)) = hf
                .paragraphs
                .iter()
                .enumerate()
                .find(|(_, p)| p.style.alignment == Some(slot))
            {
                if let Some(bottom_edge) = section_bottom_edge(index) {
                    let _ = write!(out, "#set text(bottom-edge: {bottom_edge}); ");
                }
                generate_sheet_hf_section(out, para, ctx, scale);
            }
            out.push_str("]], ");
        }
        out.push(')');
        return;
    }
    for (i, para) in hf.paragraphs.iter().enumerate() {
        if i > 0 {
            out.push_str("\\\n");
        }
        if let Some(bottom_edge) = section_bottom_edge(i) {
            let _ = write!(out, "#set text(bottom-edge: {bottom_edge}); ");
        }
        generate_sheet_hf_section(out, para, ctx, scale);
    }
}

fn generate_hf_styled_paragraph(
    out: &mut String,
    paragraph: &crate::ir::HeaderFooterParagraph,
    ctx: &mut GenCtx,
) {
    let indent_left: f64 = paragraph.style.indent_left.unwrap_or(0.0);
    let indent_right: f64 = paragraph.style.indent_right.unwrap_or(0.0);
    let has_indents: bool = indent_left.abs() > 0.0001 || indent_right.abs() > 0.0001;
    if has_indents {
        // Header/footer alignment uses the paragraph's own line measure, just
        // like body paragraphs. Keeping signed indents lets a negative Word
        // indent expand that measure into the corresponding margin.
        //
        // The block states no spacing, so in a story of several paragraphs it
        // takes the story's `par(spacing:)` and advances by Word's pitch like
        // the paragraphs around it (issue #2222). The other callers wrap it
        // alone, where Typst trims that spacing; worksheet sections state no
        // indent.
        let _ = write!(
            out,
            "#block(width: 100%, inset: (left: {}pt, right: {}pt))[",
            format_f64(indent_left),
            format_f64(indent_right)
        );
    }
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
    if has_indents {
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
    let border_space: Insets = paragraph.border_space.unwrap_or_default();
    let (left_border_outset_pt, right_border_outset_pt): (f64, f64) =
        text::paragraph_border_horizontal_outsets(paragraph.border.as_ref(), border_space);
    // `w:pBdr` sides declare their own `w:space` gap in points, measured from
    // the text's bottom edge, which is why the stack pins it there. An absent
    // `w:space` is the schema's zero, not a hairline: native Word 16.113.1
    // seats the following header line at 65.28pt both for `w:space="0"` and
    // for a `w:bottom` stating no `w:space` at all, where the 0.5pt fallback
    // this used to apply put it 0.44pt low (issue #1824).
    let space = |declared: Option<f64>| -> f64 { declared.unwrap_or(0.0) };
    let top_space: f64 = space(paragraph.border_space.map(|insets| insets.top));
    let bottom_space: f64 = space(paragraph.border_space.map(|insets| insets.bottom));
    let border_rule_extension = |border: &BorderSide, border_space: f64, is_top: bool| -> f64 {
        let (first_rule_height, second_rule_height): (f64, f64) =
            if border.style == BorderLineStyle::Double {
                (border.width * 0.67, border.width * 0.17)
            } else {
                (border.width, 0.0)
            };
        let extension_to_rule_center: f64 = if is_top {
            second_rule_height + first_rule_height / 2.0
        } else if second_rule_height > 0.0 {
            first_rule_height + second_rule_height / 2.0
        } else {
            first_rule_height / 2.0
        };
        border_space + extension_to_rule_center
    };
    let top_border_extension_pt: f64 = top_border
        .map(|border| border_rule_extension(border, top_space, true))
        .unwrap_or(0.0);
    let bottom_border_extension_pt: f64 = bottom_border
        .map(|border| border_rule_extension(border, bottom_space, false))
        .unwrap_or(0.0);

    if stacks_rules {
        out.push_str("#stack(dir: ttb, spacing: 0pt, ");
    }
    if let Some(border) = top_border {
        write_hf_border_rules(out, border, left_border_outset_pt, right_border_outset_pt);
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
    let has_vertical_borders: bool = text::write_hf_vertical_border_block_open(
        out,
        paragraph.border.as_ref(),
        left_border_outset_pt,
        right_border_outset_pt,
        top_border_extension_pt,
        bottom_border_extension_pt,
    );

    if hf_paragraph_is_mark_line(paragraph) {
        write_hf_mark_line(
            out,
            paragraph,
            hf_paragraph_rule(paragraph),
            ctx.hf_lines_are_word_boxes,
        );
    } else if let Some(index) = right_tab {
        out.push_str("#grid(columns: (1fr, auto), [");
        generate_hf_elements(out, &paragraph.elements[..index], ctx);
        out.push_str("], [");
        generate_hf_elements(out, &paragraph.elements[index + 1..], ctx);
        out.push_str("])");
    } else if paragraph.elements.iter().any(is_hf_tab) {
        // `generate_run` would write a `<w:tab/>` into the source as a literal
        // tab, which Typst's markup lexer treats as a space, so the segment a
        // stop should place sat beside the one before it (issue #579). The
        // `#h(1em)` of `generate_hf_elements` is a different element: `w:ptab`,
        // which states its own alignment rather than referring to a stop.
        let segments: Vec<&[HFInline]> = paragraph.elements.split(is_hf_tab).collect();
        let segment_runs: Vec<Vec<Run>> = segments
            .iter()
            .map(|segment| {
                segment
                    .iter()
                    .filter_map(|element| match element {
                        HFInline::Run(run) => Some(run.clone()),
                        _ => None,
                    })
                    .collect()
            })
            .collect();
        let tab_styles: Vec<&TextStyle> = paragraph
            .elements
            .iter()
            .filter_map(|element| match element {
                HFInline::Run(run) if is_hf_tab(element) => Some(&run.style),
                _ => None,
            })
            .collect();
        let default_tab_width_pt: f64 =
            paragraph_default_tab_width_pt(&paragraph.style, ctx.default_tab_width_pt);
        write_measured_tab_segments(
            out,
            &segment_runs,
            &tab_styles,
            paragraph.style.tab_stops.as_deref(),
            default_tab_width_pt,
            |out, index| generate_hf_elements(out, segments[index], ctx),
            text::generate_run,
        );
    } else {
        generate_hf_elements(out, &paragraph.elements, ctx);
    }

    if has_vertical_borders {
        out.push(']');
    }
    if stacks_rules {
        out.push(']');
    }
    if let Some(border) = bottom_border {
        let _ = write!(out, ", block(height: {}pt)[], ", format_f64(bottom_space));
        write_hf_border_rules(out, border, left_border_outset_pt, right_border_outset_pt);
    }
    if stacks_rules {
        out.push(')');
    }
}

/// Draw a header or footer paragraph that is a line of its mark, and its
/// horizontal rule on the bottom of that line (issues #2251 and #2327).
///
/// A zero-width strut in the mark's style lays that line out wherever the
/// story's form puts a line of text, so the lines around it keep their places.
/// A zero-height box as wide as the line stands on its baseline and carries
/// the rule, which Word stands on the line's bottom, one Word descent below
/// the baseline. That descent is the exact line's share where the story draws
/// Word's line boxes, and the face's natural one in the flowed and stacked
/// forms, which lay every line out on its face's natural line.
///
/// TODO(unknown mark face): with no metrics for the mark's face, that descent
/// is unknown and the rule stands on the baseline.
fn write_hf_mark_line(
    out: &mut String,
    paragraph: &crate::ir::HeaderFooterParagraph,
    rule: Option<&crate::ir::HorizontalRule>,
    lines_are_word_boxes: bool,
) {
    let mark: TextStyle = paragraph
        .style
        .paragraph_mark_text_style
        .as_deref()
        .cloned()
        .unwrap_or_default();
    if has_text_properties(&mark) {
        out.push_str("#text(");
        write_text_params(out, &mark);
        out.push_str(")[\u{2060}]");
    } else {
        out.push('\u{2060}');
    }
    let Some(rule) = rule else {
        return;
    };
    let line_box: Option<(f64, f64)> = if lines_are_word_boxes {
        hf_paragraph_line_box_pt(paragraph)
    } else {
        text::word_header_line_box_pt(&hf_paragraph_metric_runs(paragraph))
    };
    let descent_pt: f64 = match line_box {
        Some((_, descent_pt)) => descent_pt,
        None => {
            tracing::debug!(
                "horizontal rule's line has no face metrics; standing it on the baseline"
            );
            0.0
        }
    };
    out.push_str("#box(width: 100%, height: 0pt)[");
    write_horizontal_rule(out, rule, descent_pt);
    out.push(']');
}

/// Emit the one or two `line()` blocks a single paragraph-border side needs.
/// Word draws a `double` side as two thin rules, so it takes two blocks.
fn write_hf_border_rules(
    out: &mut String,
    border: &BorderSide,
    left_outset_pt: f64,
    right_outset_pt: f64,
) {
    write_hf_border_line(
        out,
        border,
        border.style == BorderLineStyle::Double,
        left_outset_pt,
        right_outset_pt,
    );
    if border.style == BorderLineStyle::Double {
        out.push_str(", ");
        write_hf_border_line(out, border, false, left_outset_pt, right_outset_pt);
    }
}

fn write_hf_border_line(
    out: &mut String,
    border: &BorderSide,
    is_primary_double: bool,
    left_outset_pt: f64,
    right_outset_pt: f64,
) {
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
        format_f64(left_outset_pt),
        format_f64(left_outset_pt + right_outset_pt),
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

/// A `<w:tab/>`: the parser gives each one a run of its own.
fn is_hf_tab(element: &HFInline) -> bool {
    matches!(element, HFInline::Run(run) if run.text == "\t")
}

fn generate_hf_elements(out: &mut String, elements: &[HFInline], ctx: &mut GenCtx) {
    for element in elements {
        match element {
            HFInline::Run(run) => {
                let mut run_markup: String = String::new();
                match text::generate_run_with_context(&mut run_markup, run, ctx) {
                    Ok(()) => out.push_str(&run_markup),
                    Err(error) => tracing::error!(
                        %error,
                        "failed to generate a header or footer run"
                    ),
                }
            }
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
    generate_blocks_with_layout_space_before(out, blocks, ctx, None, false, false, false)
}

fn generate_blocks_with_layout_space_before(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
    layout_space_before: Option<(usize, f64)>,
    joins_previous_at_boundary: bool,
    joins_next_at_boundary: bool,
    segment_ends_with_page_break: bool,
) -> Result<(), ConvertError> {
    let first_block_index: Option<usize> = blocks.iter().position(|block| !is_layout_break(block));
    let last_block_index: Option<usize> = blocks.iter().rposition(|block| !is_layout_break(block));
    let mut index: usize = 0;
    let mut shared_border_frame_end: Option<usize> = None;
    let mut pending_after_shading_space_pt: f64 = 0.0;
    while index < blocks.len() {
        if index > 0 {
            out.push('\n');
        }

        if let Some((space_index, space_before)) = layout_space_before
            && index == space_index
        {
            write_explicit_vertical_space(out, space_before);
        }

        if is_zero_size_floating_anchor(&blocks[index]) {
            let consumed = generate_floating_anchor_group(out, &blocks[index..], ctx)?;
            index += consumed;
            continue;
        }

        if shared_border_frame_end.is_none()
            && let Some((group_start, group_end)) = paragraph_border_group_bounds(blocks, index)
            && group_start == index
            && let Some(style) = paragraph_border_style(&blocks[index])
            && write_joined_paragraph_border_frame_open(out, style)
        {
            shared_border_frame_end = Some(group_end);
        }

        let mut block: Block = blocks[index].clone();
        let joins_previous_across_page: bool = joins_previous_at_boundary
            && first_block_index == Some(index)
            && paragraph_border_group_bounds(blocks, index)
                .is_some_and(|(group_start, _)| group_start == index);
        let joins_next_across_page: bool = joins_next_at_boundary
            && last_block_index == Some(index)
            && paragraph_border_group_bounds(blocks, index)
                .is_some_and(|(_, group_end)| group_end == index + 1);
        let joins_previous: bool =
            previous_non_layout_block_index(blocks, index).is_some_and(|previous_index| {
                paragraph_borders_match(&blocks[previous_index], &blocks[index])
            }) || joins_previous_across_page;
        let joins_next: bool = next_non_layout_block_index(blocks, index)
            .is_some_and(|next_index| paragraph_borders_match(&blocks[index], &blocks[next_index]))
            || joins_next_across_page;
        suppress_joined_paragraph_border_edges(&mut block, joins_previous, joins_next);
        if pending_after_shading_space_pt > 0.0 {
            reduce_paragraph_space_before(&mut block, pending_after_shading_space_pt);
            pending_after_shading_space_pt = 0.0;
        }
        let after_shading_space_pt: Option<f64> = joins_next
            .then(|| take_joined_paragraph_after_shading_space(&mut block))
            .flatten();
        ctx.flow_anchor_gap_pt = flow_anchor_gap(blocks, index, ctx);
        ctx.top_margin_frame_box_follows = matches!(
            blocks.get(index + 1),
            Some(Block::TextFrame(frame)) if frame.vertical_anchor != FrameAnchor::Text
                && !frame.text_below
        );
        if !matches!(blocks[index], Block::TextFrame(_)) {
            ctx.frame_box_carry_text_bottom_pt = None;
        }
        let follows_column_break: bool =
            index > 0 && matches!(blocks[index - 1], Block::ColumnBreak);
        let follows_source_paragraph_page_break: bool = index > 0
            && matches!(blocks[index - 1], Block::PageBreak)
            && block_starts_after_layout_break(&blocks[index]);
        let follows_initial_page_break: bool =
            ctx.at_document_start && index == 1 && matches!(blocks.first(), Some(Block::PageBreak));
        let empty_paragraph_page_break_shading_space_pt: Option<f64> =
            if is_empty_paragraph_block(&block)
                && !block_starts_after_layout_break(&block)
                && blocks.get(index + 1).is_none()
                && segment_ends_with_page_break
                && paragraph_border_style(&block).is_some_and(|style| style.background.is_some())
            {
                let style: &mut ParagraphStyle = match &mut block {
                    Block::Paragraph(paragraph) => &mut paragraph.style,
                    Block::Caption(caption) => &mut caption.paragraph.style,
                    _ => unreachable!("empty paragraph block must have a paragraph style"),
                };
                style
                    .space_after
                    .take()
                    .filter(|space_after_pt| *space_after_pt > 0.0)
            } else {
                None
            };

        if follows_column_break || follows_source_paragraph_page_break || follows_initial_page_break
        {
            let has_content_before_layout_break: bool = (follows_column_break
                || follows_source_paragraph_page_break)
                && blocks[..index - 1].iter().any(block_has_paragraph_content);
            generate_block_after_layout_break(out, &block, has_content_before_layout_break, ctx)?;
        } else {
            generate_block(out, &block, ctx)?;
        }
        if let Some(space_after_pt) = after_shading_space_pt {
            if let Some(style) = paragraph_border_style(&block) {
                write_paragraph_after_spacing_shading_block(out, style, space_after_pt);
            }
            pending_after_shading_space_pt = space_after_pt;
        }
        if let Some(space_after_pt) = empty_paragraph_page_break_shading_space_pt
            && let Some(style) = paragraph_border_style(&block)
        {
            write_paragraph_after_spacing_shading_block(out, style, space_after_pt);
        }
        if shared_border_frame_end == Some(index + 1) {
            out.push_str("\n]\n");
            shared_border_frame_end = None;
        }
        index += 1;
    }

    Ok(())
}

/// The gap Typst lays after the flow block before the text-anchored text
/// frame or placed floating table at `index`, or `None` when no block before
/// it on its page or column takes flow space (issues #2231, #2236). A
/// top-margin frame box and a margin- or page-anchored table are placed from
/// the top of the text area, not from the flow, so they get `None` too, and so
/// does a table laid out in the flow, whose own gaps are zero.
fn flow_anchor_gap(blocks: &[Block], index: usize, ctx: &GenCtx) -> Option<f64> {
    let at_flow_root: bool = ctx.in_flow_body && ctx.block_depth == 0;
    let hangs_from_the_flow: bool = match &blocks[index] {
        Block::TextFrame(frame) => frame.vertical_anchor == FrameAnchor::Text,
        Block::FloatingTable(table) => {
            table.frame.anchored_to_text
                && !floating_table_lays_out_in_flow(table, at_flow_root, ctx)
        }
        _ => false,
    };
    if !hangs_from_the_flow {
        return None;
    }
    let previous: &Block = blocks[..index]
        .iter()
        .rev()
        .find(|block| !leaves_flow_gap_unchanged(block, at_flow_root, ctx))
        .filter(|block| !is_layout_break(block))?;
    Some(section_alignment::paragraph_edge_spacing(previous).map_or(0.0, |(_, after)| after))
}

fn reduce_paragraph_space_before(block: &mut Block, preceding_after_space_pt: f64) {
    let style: Option<&mut ParagraphStyle> = match block {
        Block::Paragraph(paragraph) => Some(&mut paragraph.style),
        Block::Caption(caption) => Some(&mut caption.paragraph.style),
        _ => None,
    };
    if let Some(space_before_pt) = style.and_then(|style| style.space_before.as_mut()) {
        *space_before_pt = (*space_before_pt - preceding_after_space_pt).max(0.0);
    }
}

fn take_joined_paragraph_after_shading_space(block: &mut Block) -> Option<f64> {
    let style: &mut ParagraphStyle = match block {
        Block::Paragraph(paragraph) => &mut paragraph.style,
        Block::Caption(caption) => &mut caption.paragraph.style,
        _ => return None,
    };
    style.background?;
    style
        .space_after
        .take()
        .filter(|space_after_pt| *space_after_pt > 0.0)
}

fn previous_non_layout_block_index(blocks: &[Block], index: usize) -> Option<usize> {
    (0..index)
        .rev()
        .find(|previous_index| !is_layout_break(&blocks[*previous_index]))
}

fn next_non_layout_block_index(blocks: &[Block], index: usize) -> Option<usize> {
    (index + 1..blocks.len()).find(|next_index| !is_layout_break(&blocks[*next_index]))
}

/// Return the contiguous run of paragraph blocks with identical paragraph
/// border settings. Layout-break blocks end a frame segment, while paragraph
/// spacing inside the run remains within the shared outer frame.
fn paragraph_border_group_bounds(blocks: &[Block], index: usize) -> Option<(usize, usize)> {
    paragraph_border_style(blocks.get(index)?)?
        .border
        .as_ref()?;
    let mut start: usize = index;
    while start > 0 && paragraph_borders_match(&blocks[start - 1], &blocks[start]) {
        start -= 1;
    }
    let mut end: usize = index + 1;
    while end < blocks.len() && paragraph_borders_match(&blocks[end - 1], &blocks[end]) {
        end += 1;
    }
    (end > start + 1 && blocks[start..end].iter().any(block_has_paragraph_content))
        .then_some((start, end))
}

fn paragraph_border_style(block: &Block) -> Option<&ParagraphStyle> {
    match block {
        Block::Paragraph(paragraph) => Some(&paragraph.style),
        Block::Caption(caption) => Some(&caption.paragraph.style),
        _ => None,
    }
}

fn paragraph_borders_match(first: &Block, second: &Block) -> bool {
    if !block_has_paragraph_content(first) || !block_has_paragraph_content(second) {
        return false;
    }
    let Some(first_style) = paragraph_border_style(first) else {
        return false;
    };
    let Some(second_style) = paragraph_border_style(second) else {
        return false;
    };
    let Some(first_border) = first_style.border.as_deref() else {
        return false;
    };
    let Some(second_border) = second_style.border.as_deref() else {
        return false;
    };
    let first_space: Insets = first_style
        .border_space
        .as_deref()
        .copied()
        .unwrap_or_default();
    let second_space: Insets = second_style
        .border_space
        .as_deref()
        .copied()
        .unwrap_or_default();
    first_border == second_border && first_space == second_space
}

fn suppress_joined_paragraph_border_edges(
    block: &mut Block,
    joins_previous: bool,
    joins_next: bool,
) {
    let Some(style) = (match block {
        Block::Paragraph(paragraph) => Some(&mut paragraph.style),
        Block::Caption(caption) => Some(&mut caption.paragraph.style),
        _ => None,
    }) else {
        return;
    };
    let Some(border) = style.border.as_mut() else {
        return;
    };
    let mut space: Insets = style.border_space.as_deref().copied().unwrap_or_default();
    if joins_previous {
        let top_border_inset_pt: f64 = border.top.as_ref().map_or(0.0, |side| {
            space.top
                + if side.style == BorderLineStyle::Double {
                    side.width * 3.0
                } else {
                    side.width / 2.0
                }
        });
        if top_border_inset_pt > 0.0
            && let Some(before_spacing_pt) = style.decoration_before_spacing.or(style.space_before)
            && before_spacing_pt > 0.0
        {
            style.decoration_before_spacing = Some(before_spacing_pt + top_border_inset_pt);
        }
        border.top = None;
        space.top = 0.0;
    }
    if joins_next {
        border.bottom = None;
        space.bottom = 0.0;
    }
    if border.top.is_none()
        && border.bottom.is_none()
        && border.left.is_none()
        && border.right.is_none()
    {
        style.border = None;
        style.border_space = None;
    } else if style.border_space.is_some() {
        style.border_space = Some(Box::new(space));
    }
}

/// Preserve before-spacing when a layout break splits the same source paragraph.
/// A leading break keeps the full gap; a later break after content first absorbs
/// that paragraph's after-spacing.
fn generate_block_after_layout_break(
    out: &mut String,
    block: &Block,
    should_collapse_with_after_spacing: bool,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let mut continuation: Block = block.clone();
    emit_layout_break_space_before(out, &mut continuation, should_collapse_with_after_spacing);
    generate_block(out, &continuation, ctx)
}

fn emit_layout_break_space_before(
    out: &mut String,
    block: &mut Block,
    should_collapse_with_after_spacing: bool,
) {
    if let Some(space_before) =
        take_layout_break_space_before(block, should_collapse_with_after_spacing)
    {
        write_explicit_vertical_space(out, space_before);
    }
}

fn take_layout_break_space_before(
    block: &mut Block,
    should_collapse_with_after_spacing: bool,
) -> Option<f64> {
    if !block_has_paragraph_content(block)
        && !paragraph_border_style(block).is_some_and(|style| style.background.is_some())
    {
        return None;
    }
    let (space_before, space_after): (&mut Option<f64>, Option<f64>) = match block {
        Block::Paragraph(paragraph) => (
            &mut paragraph.style.space_before,
            paragraph.style.space_after,
        ),
        Block::Caption(caption) => (
            &mut caption.paragraph.style.space_before,
            caption.paragraph.style.space_after,
        ),
        Block::List(list) => {
            let paragraph = list.items.first_mut()?.content.first_mut()?;
            (
                &mut paragraph.style.space_before,
                paragraph.style.space_after,
            )
        }
        _ => return None,
    };
    let raw_space_before: f64 = space_before.take()?;
    if raw_space_before <= 0.0 {
        *space_before = Some(raw_space_before);
        return None;
    }
    let collapsed_space: f64 = if should_collapse_with_after_spacing {
        (raw_space_before - space_after.unwrap_or(0.0)).max(0.0)
    } else {
        raw_space_before
    };
    (collapsed_space > 0.0).then_some(collapsed_space)
}

fn write_explicit_vertical_space(out: &mut String, space_before: f64) {
    let _ = writeln!(out, "#v({}pt, weak: false)", format_f64(space_before));
}

fn column_segment_location(content: &[Block], block_index: usize) -> (usize, usize) {
    let mut segment_index: usize = 0;
    let mut block_in_segment_index: usize = 0;
    for block in &content[..block_index] {
        if matches!(block, Block::ColumnBreak) {
            segment_index += 1;
            block_in_segment_index = 0;
        } else {
            block_in_segment_index += 1;
        }
    }
    (segment_index, block_in_segment_index)
}

fn suppress_block_space_before(block: &mut Block) {
    let (space_before, space_after, decoration_before_spacing): (
        &mut Option<f64>,
        Option<f64>,
        &mut Option<f64>,
    ) = match block {
        Block::Paragraph(paragraph) => (
            &mut paragraph.style.space_before,
            paragraph.style.space_after,
            &mut paragraph.style.decoration_before_spacing,
        ),
        Block::Caption(caption) => (
            &mut caption.paragraph.style.space_before,
            caption.paragraph.style.space_after,
            &mut caption.paragraph.style.decoration_before_spacing,
        ),
        Block::List(list) => {
            let paragraph = list
                .items
                .first_mut()
                .and_then(|item| item.content.first_mut())
                .unwrap_or_else(|| unreachable!("a list content block has its first paragraph"));
            (
                &mut paragraph.style.space_before,
                paragraph.style.space_after,
                &mut paragraph.style.decoration_before_spacing,
            )
        }
        _ => return,
    };
    if let Some(raw_space_before) = *space_before
        && raw_space_before > 0.0
    {
        let decoration_spacing: f64 = (raw_space_before - space_after.unwrap_or(0.0)).max(0.0);
        if decoration_spacing > 0.0 {
            *decoration_before_spacing = Some(decoration_spacing);
        }
    }
    *space_before = None;
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
                generate_floating_text_box_overlay(out, text_box, ctx, false, text_box.offset_y)?;
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
    // breaker instead, which never squeezes to seat a token. The gate is
    // `justified_lines_use_first_fit_breaker`: it covers legacy East Asian
    // paragraphs and legacy table-cell paragraphs while leaving Latin body
    // paragraphs on the optimized breaker (issues #1130, #2020). A justified
    // list item always breaks first-fit (issue #2336).
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

/// State the language the document's text is written in.
///
/// Typst otherwise lays every document out as English and declares
/// `/Lang(en)` in the PDF, whatever language the source names. The language
/// also selects Typst's line-breaking and hyphenation rules.
///
/// `dir: ltr` keeps the base direction Typst used before the language was
/// stated: its `dir: auto` follows the language, so an `ar` or `he` default
/// would otherwise turn every left-to-right paragraph around. A right-to-left
/// paragraph still states `dir: rtl` itself.
fn write_document_language(out: &mut String, language: Option<&crate::ir::DocumentLanguage>) {
    let Some(language) = language else {
        return;
    };
    let _ = write!(
        out,
        "#set text(lang: \"{}\"",
        escape_typst_string(&language.language)
    );
    if let Some(region) = &language.region {
        let _ = write!(out, ", region: \"{}\"", escape_typst_string(region));
    }
    let _ = writeln!(out, ", dir: ltr)");
}

/// State whether words may be hyphenated at a line end.
///
/// An explicit rule prevents Typst's automatic justification policy from
/// introducing hyphens when the source enables no hyphenation. The native
/// Word fixture in issue #2052 demonstrates that default mismatch. The DOCX
/// parser can enable the rule through `w:autoHyphenation`; the other parsers
/// leave it disabled.
///
/// Set once for the document so header and footer bands, table cells and
/// slide text boxes are all covered.
fn write_hyphenation_rule(out: &mut String, hyphenates_automatically: bool) {
    let _ = writeln!(out, "#set text(hyphenate: {hyphenates_automatically})");
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
    let is_root_flow_block = ctx.in_flow_body && ctx.block_depth == 0;
    let at_page_root: bool = ctx.preserve_page_after_spacing && is_root_flow_block;
    let gaps: Option<(f64, f64)> = at_page_root
        .then(|| section_alignment::paragraph_edge_spacing(block))
        .flatten();
    if let Some((before, _)) = gaps {
        let residual_before: f64 = (before - ctx.flow_previous_after_spacing).max(0.0);
        let _ = writeln!(out, "#v({}pt, weak: true)", format_f64(residual_before));
    }
    ctx.float_wrap_pending = is_root_flow_block
        && !ctx.body_floats.is_empty()
        && matches!(block, Block::Paragraph(_) | Block::Caption(_));
    ctx.block_depth += 1;
    let result = generate_block_inner(out, block, ctx);
    ctx.block_depth -= 1;
    ctx.float_wrap_pending = false;
    // A float only reaches the text on its own page and in its own column.
    if is_root_flow_block && matches!(block, Block::PageBreak | Block::ColumnBreak) {
        ctx.body_floats.clear();
    }
    if is_root_flow_block && matches!(block, Block::PageBreak) {
        advance_flow_body_alignment(out, ctx);
    }
    if at_page_root {
        if let Some((_, after)) = gaps {
            out.push_str("#v(0pt, weak: true)\n");
            let _ = writeln!(out, "#v({}pt, weak: false)", format_f64(after));
            ctx.flow_previous_after_spacing = after;
        } else if !leaves_flow_gap_unchanged(block, is_root_flow_block, ctx) {
            ctx.flow_previous_after_spacing = 0.0;
        }
    }
    result
}

fn generate_block_inner(
    out: &mut String,
    block: &Block,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    match block {
        Block::Paragraph(para) => generate_paragraph(out, para, ctx),
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
            generate_paragraph(out, &caption.paragraph, ctx)
        }
        Block::PageBreak => {
            out.push_str("#pagebreak()\n");
            Ok(())
        }
        Block::FloatingTable(floating_table) => generate_floating_table(out, floating_table, ctx),
        Block::TextFrame(frame) => generate_text_frame(out, frame, ctx),
        Block::Table(table) => {
            // Word gives a table no vertical spacing of its own: the gap above
            // it is the preceding paragraph's `w:after` and the gap below it is
            // the following paragraph's `w:before`. Typst resolves the gap
            // between two blocks by weakness rather than by plain maximum.
            // `typst-layout`'s flow collector tags an `auto` gap with the
            // `par.spacing` fallback — 1.2em — at weakness 4 and a stated one
            // at weakness 3, and `keep_weak_rel_spacing` replaces the standing
            // gap only when the arriving one is strictly stronger or equally
            // weak and larger. A stated gap therefore always wins, and the
            // fallback survives only when neither neighbour states anything. A
            // bare `#table` beside a paragraph that states no `w:before` hit
            // exactly that case and opened 1.2em of engine whitespace: 13.2pt
            // at 11pt, under every table in a document written from Word's
            // default template (issue #1688). Stating both gaps leaves the
            // neighbour's own `w:spacing` as the only thing between them, the
            // way `write_image_block_open` already states both of a flow
            // picture's gaps instead of letting the fallback apply (issues
            // #463, #491, #499); a table differs only in having no `w:spacing`
            // of its own to state, so both are zero. The wrapper spans the text
            // column so a `w:tblPr/w:jc` table still centres in it, and the
            // table's own `auto` gaps vanish against the wrapper's edges, where
            // the same routine drops weak spacing that no frame precedes.
            let column_width_pt: Option<f64> = ctx
                .float_wrap_text_area_pt
                .filter(|_| ctx.clears_top_margin_frame_boxes && ctx.block_depth == 1)
                .map(|(width_pt, _)| width_pt);
            let clears_frame_boxes: bool = column_width_pt
                .is_some_and(|width_pt| table.column_widths.iter().sum::<f64>() >= width_pt - 0.05);
            if clears_frame_boxes {
                float_wrap::write_top_margin_frame_box_clearance(out);
            }
            // Read before the cells' own block lists overwrite the flag.
            let text_bottom_pt: Option<f64> = ctx
                .active_page_size
                .zip(ctx.active_page_margins)
                .map(|(size, margins)| size.height - margins.bottom)
                .filter(|_| clears_frame_boxes && ctx.top_margin_frame_box_follows);
            out.push_str("#block(width: 100%, above: 0pt, below: 0pt)[\n");
            if text_bottom_pt.is_some() {
                float_wrap::write_frame_box_table_label(out, "start");
            }
            // Only a table written straight into a breakable section body
            // meets a page break itself. A table inside a placed, boxed or
            // otherwise single-frame container keeps its rows whole: Typst
            // stacks a split row's remainder on its first part there.
            let enclosing_rows_may_split: bool = ctx.table_rows_may_split;
            ctx.table_rows_may_split = ctx.flow_body_is_breakable && ctx.block_depth == 1;
            let result = generate_table(out, table, ctx);
            ctx.table_rows_may_split = enclosing_rows_may_split;
            if text_bottom_pt.is_some() {
                float_wrap::write_frame_box_table_label(out, "end");
            }
            out.push_str("]\n");
            ctx.frame_box_carry_text_bottom_pt = text_bottom_pt;
            result
        }
        Block::Image(img) => {
            // Word advances a picture paragraph by the picture plus its own
            // `w:spacing`. Leaving the element bare let Typst's 1.2em default
            // block spacing apply above and below instead, opening ~24pt
            // around an inline figure (issues #463, #491) - so both gaps stay
            // pinned, but to the declared spacing rather than to zero, which
            // dropped the gap Word actually draws (issue #499).
            // The picture is its paragraph's only line, so it spans that
            // line's measure: from `left + first_line` to `right`, inside
            // which the alignment works, as for a line of text (issue #2201).
            let inset: Option<(f64, f64)> = img
                .paragraph_indent
                .map(|indent| (indent.left + indent.first_line, indent.right));
            write_image_block_open(out, img.paragraph_spacing, inset);
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
            // The pictures wrap like words: every line spans the paragraph's
            // `left` to `right`, and only the first also takes `first_line`.
            let indent: Option<ImageParagraphIndent> =
                images.first().and_then(|first| first.paragraph_indent);
            write_image_block_open(
                out,
                spacing,
                indent.map(|indent| (indent.left, indent.right)),
            );
            out.push('\n');
            let first_line: Option<f64> = indent
                .map(|indent| indent.first_line)
                .filter(|first_line| first_line.abs() > 0.0001);
            if let Some(first_line) = first_line {
                let _ = write!(
                    out,
                    "#par(first-line-indent: (amount: {}pt, all: true))[",
                    format_f64(first_line)
                );
            }
            for (index, image) in images.iter().enumerate() {
                if index > 0 {
                    out.push(' ');
                }
                out.push_str("#box[");
                generate_image(out, image, ctx);
                out.pop();
                out.push(']');
            }
            if first_line.is_some() {
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
        Block::DrawingGroup(group) => generate_drawing_group(out, group, ctx),
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
                    // A Word list hangs into the margin like its paragraphs.
                    hangs_into_margin: ctx.breaks_hangul_at_eojeol,
                },
                ctx,
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
/// A `(left, right)` inset narrows the picture's measure to its paragraph's
/// indents. A negative one widens it into the margin, as Word does, the same
/// way a body paragraph's negative indent widens its text lines (issue #2233).
fn write_image_block_open(
    out: &mut String,
    spacing: Option<ImageParagraphSpacing>,
    inset: Option<(f64, f64)>,
) {
    let above: f64 = spacing.and_then(|gap| gap.before).unwrap_or(0.0);
    let below: f64 = spacing.and_then(|gap| gap.after).unwrap_or(0.0);
    let _ = write!(
        out,
        "#block(width: 100%, above: {}pt, below: {}pt",
        format_f64(above),
        format_f64(below)
    );
    if let Some((left, right)) =
        inset.filter(|(left, right)| left.abs() > 0.0001 || right.abs() > 0.0001)
    {
        let _ = write!(
            out,
            ", inset: (left: {}pt, right: {}pt)",
            format_f64(left),
            format_f64(right)
        );
    }
    out.push_str(")[");
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

    if let Some(fill) = img.fill_rect
        && let (Some(width), Some(height)) = (img.width, img.height)
    {
        let fill_width = width * (1.0 - fill.left - fill.right);
        let fill_height = height * (1.0 - fill.top - fill.bottom);
        if !(fill_width.is_finite()
            && fill_height.is_finite()
            && fill_width > 0.0
            && fill_height > 0.0)
        {
            return;
        }
        // Destination outsets enlarge the source inside the fixed shape frame;
        // insets expose its background. The frame itself never changes size.
        let _ = write!(
            out,
            "#box(width: {}pt, height: {}pt, clip: true)[#place(top + left, dx: {}pt, dy: {}pt)[",
            format_f64(width),
            format_f64(height),
            format_f64(width * fill.left),
            format_f64(height * fill.top),
        );
        let mut inner = img.clone();
        inner.fill_rect = None;
        inner.width = Some(fill_width);
        inner.height = Some(fill_height);
        generate_image(out, &inner, ctx);
        out.push_str("]]\n");
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
/// - Square/Tight: `#place()` with `float: true` for best-effort text flow
/// - TopAndBottom: a block reserves the image height
fn generate_floating_image(out: &mut String, fi: &FloatingImage, ctx: &mut GenCtx) {
    generate_floating_image_with_vertical_offset(out, fi, ctx, fi.offset_y);
}

fn generate_floating_image_with_vertical_offset(
    out: &mut String,
    fi: &FloatingImage,
    ctx: &mut GenCtx,
    offset_y: f64,
) {
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
                format_f64(offset_y)
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
                format_f64(offset_y)
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
    // DrawingML's frame can have a different aspect ratio from the asset.
    // Stretch the retained source into that frame instead of cropping it.
    if image.width.is_some() && image.height.is_some() {
        out.push_str(", fit: \"stretch\"");
    }
    out.push(')');

    if rotation.is_some() {
        out.push(']');
    }
    if pivot_shift.is_some() {
        out.push(']');
    }
}

/// Whether a root text-anchored floating table leaves no room beside it in
/// the one text column, so that Word starts every line of its anchor
/// paragraph and after it that meets it below it. Such a table is laid out in
/// the flow, where Typst continues it on the next page as Word does (issue
/// #2236). A table hung above its anchor is placed, as the flow cannot go back
/// up to it.
fn floating_table_lays_out_in_flow(
    floating_table: &FloatingTable,
    at_flow_root: bool,
    ctx: &GenCtx,
) -> bool {
    floating_table.frame.anchored_to_text
        && at_flow_root
        && floating_table.frame.y.unwrap_or_default() >= 0.0
        && ctx.float_wrap_text_area_pt.is_some_and(|(width_pt, _)| {
            float_wrap::leaves_no_room_beside(
                float_wrap::FloatExtent::of_table(floating_table),
                width_pt,
            )
        })
}

/// Whether `block` leaves the gap between the flow blocks around it as it
/// was, taking no flow space of its own. A floating table laid out in the
/// flow takes some.
fn leaves_flow_gap_unchanged(block: &Block, at_flow_root: bool, ctx: &GenCtx) -> bool {
    section_alignment::leaves_flow_gap_unchanged(block)
        && !matches!(block, Block::FloatingTable(table)
            if floating_table_lays_out_in_flow(table, at_flow_root, ctx))
}

/// Draw a `w:tblpPr` floating table where Word does. A margin- or
/// page-anchored table's offsets are measured from the text area's top left.
/// A text-anchored one hangs `y` below the top edge of its anchor paragraph,
/// the one after it (issue #2236). That edge is the end of the gap after the
/// block before the table: Word's exports of docxide-pdf's `case32` and
/// `empty-float-float-indent-*` add that block's `w:after` to the offset and
/// not the anchor paragraph's own `w:before`.
///
/// A text-anchored table that leaves no room beside it is laid out in the
/// flow, after fixed space as tall as its offset, which adds to the gap before
/// it rather than merging with it. Word continues such a table on the next
/// page when it runs past the text area, as its exports of docxide-pdf's
/// `scraped/bulgarian_road_safety_program` and
/// `scraped/croatian_grant_guidelines` show, which a `#place` cannot, and
/// starts the anchor paragraph below its bottom edge. Any other text-anchored
/// table is placed by a `#place` without a vertical alignment, which finds the
/// anchor in the flow. As for a text-anchored text frame, the gap is collapsed
/// to nothing above the table and added to its offset, and a zero-height block
/// after it lays the gap again. At the flow root, the paragraphs after a
/// placed table wrap around it (issue #2229).
///
/// TODO(anchor lines above a full-width table's offset): Word keeps a line of
/// the anchor paragraph that ends above the table's top edge above it, as the
/// wrap does beside a placed table; here the whole paragraph starts below the
/// table. No corpus document hangs a full-width table more than a line below
/// a non-empty anchor.
fn generate_floating_table(
    out: &mut String,
    floating_table: &FloatingTable,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let offset_x: f64 = floating_table.frame.x.unwrap_or_default();
    let offset_y: f64 = floating_table.frame.y.unwrap_or_default();
    if floating_table_lays_out_in_flow(
        floating_table,
        ctx.in_flow_body && ctx.block_depth == 1,
        ctx,
    ) {
        return generate_floating_table_in_flow(out, floating_table, offset_x, offset_y, ctx);
    }
    // A vertically centred section writes its own explicit gaps between root
    // blocks; there the table lands below them.
    let anchor_gap: Option<f64> = std::mem::take(&mut ctx.flow_anchor_gap_pt)
        .filter(|_| ctx.block_depth == 1 && !ctx.preserve_page_after_spacing);
    // Only a table placed at the flow root has root paragraphs after it to
    // wrap around it.
    let wrap: Option<float_wrap::BodyFloat> = ctx
        .float_wrap_text_area_pt
        .filter(|_| ctx.in_flow_body && ctx.block_depth == 1)
        .map(|(width_pt, height_pt)| {
            let id: usize = ctx.next_body_float_id;
            ctx.next_body_float_id += 1;
            float_wrap::BodyFloat::new(
                id,
                float_wrap::FloatExtent::of_table(floating_table),
                width_pt,
                height_pt,
                false,
            )
        });
    // The labels mark the band the paragraphs after the table keep clear
    // of: the table widened by `w:topFromText` above and `w:bottomFromText`
    // below, as `w:vSpace` widens a text frame's. Word's exports of
    // docxide-pdf's `case164` and `case165`, one factor apart, lay a line
    // that ends 6pt above the table at the margin with no `topFromText` and
    // beside the table with 9pt of it.
    // FIXME(unprobed bottomFromText): no Word reference in the corpus
    // varies it on its own; it follows ECMA-376's wording, as `topFromText`
    // does.
    let (band_top_pt, band_bottom_pt): (f64, f64) = if wrap.is_some() {
        (
            floating_table.frame.top_from_text,
            floating_table.frame.bottom_from_text,
        )
    } else {
        (0.0, 0.0)
    };
    if floating_table.frame.anchored_to_text {
        if anchor_gap.is_some() {
            out.push_str("#v(0pt, weak: true)\n");
        }
        let _ = writeln!(
            out,
            "#place(left, dx: {}pt, dy: {}pt)[",
            format_f64(offset_x),
            format_f64(offset_y + anchor_gap.unwrap_or(0.0) - band_top_pt)
        );
    } else {
        let _ = writeln!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt)[",
            format_f64(offset_x),
            format_f64(offset_y - band_top_pt)
        );
    }
    let write_band_space = |out: &mut String, height_pt: f64| {
        if height_pt > 0.0 {
            let _ = writeln!(
                out,
                "#block(height: {}pt, above: 0pt, below: 0pt)",
                format_f64(height_pt)
            );
        }
    };
    if let Some(wrap) = wrap {
        float_wrap::write_float_top_label(out, wrap.id());
    }
    write_band_space(out, band_top_pt);
    let result: Result<(), ConvertError> = generate_table(out, &floating_table.table, ctx);
    write_band_space(out, band_bottom_pt);
    if let Some(wrap) = wrap {
        float_wrap::write_float_bottom_label(out, wrap.id());
        ctx.body_floats.push(wrap);
    }
    out.push_str("]\n");
    if let Some(gap) = anchor_gap {
        let _ = writeln!(
            out,
            "#block(height: 0pt, above: 0pt, below: {}pt)",
            format_f64(gap)
        );
    }
    result
}

/// Lay a text-anchored floating table that leaves no room beside it out in
/// the flow, `offset_y` below the end of the gap before it and `offset_x`
/// from the text area's left edge. Its rows may continue on the next page, as
/// a flow table's do. `offset_x` already places the table where its
/// `w:tblPr/w:jc` puts it, so the table is not aligned again.
fn generate_floating_table_in_flow(
    out: &mut String,
    floating_table: &FloatingTable,
    offset_x: f64,
    offset_y: f64,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    tracing::debug!(
        offset_x_pt = offset_x,
        offset_y_pt = offset_y,
        rows = floating_table.table.rows.len(),
        "laying a text-anchored floating table out in the flow"
    );
    if offset_y > 0.0 {
        write_explicit_vertical_space(out, offset_y);
    }
    let _ = writeln!(
        out,
        "#block(width: 100%, above: 0pt, below: 0pt, inset: (left: {}pt))[",
        format_f64(offset_x)
    );
    let mut table: Table = floating_table.table.clone();
    table.alignment = None;
    let enclosing_rows_may_split: bool = ctx.table_rows_may_split;
    ctx.table_rows_may_split = ctx.flow_body_is_breakable;
    let result: Result<(), ConvertError> = generate_table(out, &table, ctx);
    ctx.table_rows_may_split = enclosing_rows_may_split;
    out.push_str("]\n");
    // The text after the table starts `w:bottomFromText` below it, as a line
    // that meets a placed table's band does.
    // FIXME(unprobed bottomFromText): no Word reference in the corpus
    // states it on a table that leaves no room beside it.
    // TODO(bottomFromText across a page break): the space is fixed, so it
    // still opens the next page when the table ends at the bottom of one.
    if floating_table.frame.bottom_from_text > 0.0 {
        write_explicit_vertical_space(out, floating_table.frame.bottom_from_text);
    }
    result
}

/// Draw a lifted `w:framePr` frame where Word does (issue #2231): `y` below
/// its anchor, which a `#place` without a vertical alignment finds in the
/// flow, and `x` from the text area's left edge. At the flow root, the
/// paragraphs after the frame wrap around it as they do around a floating
/// table, with `w:vSpace` widening the band they keep clear of.
///
/// The frame takes no flow space. A bare placed item does not either, but
/// Typst merges the gaps on both sides of it and lands it below the merged
/// gap, where Word hangs the frame from the end of the gap after the block
/// before it: docxide-pdf's `paragraph-offset-frame-exact` puts its frame
/// where the immediate case does, though the paragraph after it states a
/// 12pt `w:before`. So the frame is placed with the previous block's gap
/// collapsed to nothing above it and that gap added to its offset, and a
/// zero-height block after it lays the gap again, where it still merges with
/// the next block's gap as before.
fn generate_text_frame(
    out: &mut String,
    frame: &TextFrame,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    if frame.text_below {
        return generate_fixed_text_frame(out, frame, ctx);
    }
    if frame.vertical_anchor != FrameAnchor::Text {
        return generate_top_margin_frame_box(out, frame, ctx);
    }
    // A vertically centred section writes its own explicit gaps between root
    // blocks; there the frame lands below them.
    let anchor_gap: Option<f64> = std::mem::take(&mut ctx.flow_anchor_gap_pt)
        .filter(|_| ctx.block_depth == 1 && !ctx.preserve_page_after_spacing);
    if anchor_gap.is_some() {
        // The strongest weak spacing replaces the gap the previous block
        // left, and no gap a later block brings outweighs it.
        out.push_str("#v(0pt, weak: true)\n");
    }
    let vertical_space: f64 = frame.vertical_space;
    let _ = writeln!(
        out,
        "#place(left, dx: {}pt, dy: {}pt)[",
        format_f64(frame.x),
        format_f64(frame.y + anchor_gap.unwrap_or(0.0) - vertical_space)
    );
    // The paragraphs after an exact frame start at its anchor, so once they
    // have advanced past its bottom edge none of them can meet it; an
    // at-least frame's content can reach further, to the page's end.
    let frame_bottom_pt: f64 = frame.y + anchor_gap.unwrap_or(0.0) + frame.height + vertical_space;
    let wrap: Option<float_wrap::BodyFloat> = ctx
        .float_wrap_text_area_pt
        .filter(|_| ctx.in_flow_body && ctx.block_depth == 1)
        .map(|(width_pt, height_pt)| {
            let id: usize = ctx.next_body_float_id;
            ctx.next_body_float_id += 1;
            let reach_pt: f64 = if frame.exact_height {
                frame_bottom_pt.clamp(0.0, height_pt)
            } else {
                height_pt
            };
            float_wrap::BodyFloat::new(
                id,
                float_wrap::FloatExtent::of_text_frame(frame),
                width_pt,
                reach_pt,
                true,
            )
        });
    if let Some(wrap) = wrap {
        float_wrap::write_float_top_label(out, wrap.id());
    }
    let write_vertical_space = |out: &mut String| {
        if vertical_space > 0.0 {
            let _ = writeln!(
                out,
                "#block(height: {}pt, above: 0pt, below: 0pt)",
                format_f64(vertical_space)
            );
        }
    };
    write_vertical_space(out);
    let width: String = format_f64(frame.width);
    let height: String = format_f64(frame.height);
    if frame.exact_height {
        let _ = writeln!(
            out,
            "#block(width: {width}pt, height: {height}pt, above: 0pt, below: 0pt)["
        );
    } else {
        // A least height: the frame is as tall as the taller of its stated
        // height and its content.
        let _ = writeln!(
            out,
            "#block(above: 0pt, below: 0pt, stack(dir: ltr, box(height: {height}pt), \
             block(width: {width}pt)["
        );
    }
    // A frame's content starts at its top, whatever the vertical alignment of
    // the section around it.
    out.push_str("#set align(top)\n");
    let enclosing_measure_pt: Option<f64> = ctx.available_measure_pt.replace(frame.width);
    let result: Result<(), ConvertError> = frame
        .content
        .iter()
        .try_for_each(|block| generate_block(out, block, ctx));
    ctx.available_measure_pt = enclosing_measure_pt;
    out.push_str(if frame.exact_height { "]\n" } else { "]))\n" });
    write_vertical_space(out);
    if let Some(wrap) = wrap {
        float_wrap::write_float_bottom_label(out, wrap.id());
        ctx.body_floats.push(wrap);
    }
    out.push_str("]\n");
    if let Some(gap) = anchor_gap {
        let _ = writeln!(
            out,
            "#block(height: 0pt, above: 0pt, below: {}pt)",
            format_f64(gap)
        );
    }
    result
}

/// Draw a frame Word keeps at a fixed place on its page and keeps body text
/// off the sides of (issue #2231): its content at `x` and `y` from the text
/// area's top left, `w:w` wide. At the root of a one-column flow, labels mark
/// its band, `w:vSpace` above and below it, and a root paragraph after it on
/// its page whose first line meets the band starts below it.
///
/// FIXME(auto frame width): Word makes a frame without `w:w` as wide as its
/// widest line or table. Its paragraphs here span the text area's width from
/// the frame's left edge, the width `measure` cannot find for blocks
/// `100%` wide, so a line centred or aligned right in it lands further right.
///
/// TODO(text before a fixed frame on its page): Word keeps body text that
/// comes before the frame in the document out of its band too; here only the
/// paragraphs after it move, and a table after it does not.
///
/// TODO(band reaching the page's end): a paragraph whose first line would
/// start below the band but end past the text area belongs on the next page;
/// the wrap's top pad does not move it there, so it overhangs the bottom
/// margin or, when the move would change its page, stays beside the band
/// (bulk `tdf104394_lostTextbox`).
fn generate_fixed_text_frame(
    out: &mut String,
    frame: &TextFrame,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let vertical_space: f64 = frame.vertical_space;
    let _ = writeln!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(frame.x),
        format_f64(frame.y - vertical_space)
    );
    let wrap: Option<float_wrap::BodyFloat> = ctx
        .float_wrap_text_area_pt
        .filter(|_| ctx.in_flow_body && ctx.block_depth == 1)
        .map(|(_, height_pt)| {
            let id: usize = ctx.next_body_float_id;
            ctx.next_body_float_id += 1;
            float_wrap::BodyFloat::below(id, height_pt)
        });
    if let Some(wrap) = wrap {
        float_wrap::write_float_top_label(out, wrap.id());
    }
    let write_vertical_space = |out: &mut String| {
        if vertical_space > 0.0 {
            let _ = writeln!(
                out,
                "#block(height: {}pt, above: 0pt, below: 0pt)",
                format_f64(vertical_space)
            );
        }
    };
    write_vertical_space(out);
    let width: String = if frame.width > 0.0 {
        format!("width: {}pt, ", format_f64(frame.width))
    } else {
        String::new()
    };
    let height: String = format_f64(frame.height);
    if frame.exact_height {
        let _ = writeln!(
            out,
            "#block({width}height: {height}pt, above: 0pt, below: 0pt)["
        );
    } else {
        let _ = writeln!(
            out,
            "#block(above: 0pt, below: 0pt, stack(dir: ltr, box(height: {height}pt), \
             block({width}above: 0pt, below: 0pt)["
        );
    }
    out.push_str("#set align(top)\n");
    let enclosing_measure_pt: Option<f64> = ctx.available_measure_pt;
    if frame.width > 0.0 {
        ctx.available_measure_pt = Some(frame.width);
    }
    let result: Result<(), ConvertError> = frame
        .content
        .iter()
        .try_for_each(|block| generate_block(out, block, ctx));
    ctx.available_measure_pt = enclosing_measure_pt;
    out.push_str(if frame.exact_height { "]\n" } else { "]))\n" });
    write_vertical_space(out);
    if let Some(wrap) = wrap {
        float_wrap::write_float_bottom_label(out, wrap.id());
        ctx.body_floats.push(wrap);
    }
    out.push_str("]\n");
    result
}

/// Draw the box Word keeps at the top margin for a group of empty framed
/// paragraphs (issue #2231): their lines at the frame's width, placed from the
/// top of the text area on the page the group lands on. At the root of a
/// one-column flow that pages, labels mark the box's top edge and its bottom
/// edge, `w:vSpace` below it, which a full-width table that starts between
/// them clears. The box takes no flow space, except above the rows a
/// full-width table it directly follows carries onto its page.
///
/// TODO(body text beside an empty frame's box, #2355): Word lays a body line
/// the box meets out on both sides of it, as docxide-pdf's
/// `empty-frame-control` shows; here the line runs as though the box were
/// absent.
fn generate_top_margin_frame_box(
    out: &mut String,
    frame: &TextFrame,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // A continuous section's frame is a `#move` whose top is not the page's
    // text area top.
    let labelled: bool = ctx.in_flow_body
        && ctx.block_depth == 1
        && ctx.flow_body_is_breakable
        && ctx.float_wrap_text_area_pt.is_some();
    let carry_text_bottom_pt: Option<f64> = ctx
        .frame_box_carry_text_bottom_pt
        .take()
        .filter(|_| labelled);
    if let Some(text_bottom_pt) = carry_text_bottom_pt {
        float_wrap::write_top_margin_frame_box_carry_open(out, text_bottom_pt, frame.x, frame.y);
    } else {
        let _ = writeln!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt)[",
            format_f64(frame.x),
            format_f64(frame.y)
        );
    }
    if labelled {
        float_wrap::write_top_margin_frame_box_label(out, "top");
    }
    let _ = writeln!(
        out,
        "#block(width: {}pt, above: 0pt, below: 0pt)[",
        format_f64(frame.width)
    );
    let enclosing_measure_pt: Option<f64> = ctx.available_measure_pt.replace(frame.width);
    let result: Result<(), ConvertError> = frame
        .content
        .iter()
        .try_for_each(|block| generate_block(out, block, ctx));
    ctx.available_measure_pt = enclosing_measure_pt;
    out.push_str("]\n");
    if frame.vertical_space > 0.0 {
        let _ = writeln!(
            out,
            "#block(height: {}pt, above: 0pt, below: 0pt)",
            format_f64(frame.vertical_space)
        );
    }
    if labelled {
        float_wrap::write_top_margin_frame_box_label(out, "bottom");
    }
    out.push_str("]\n");
    result
}

fn generate_floating_text_box(
    out: &mut String,
    ftb: &FloatingTextBox,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    match ftb.wrap_mode {
        WrapMode::TopAndBottom => {
            out.push_str("#block(width: 100%)[\n");
            generate_floating_text_box_overlay(out, ftb, ctx, false, 0.0)?;
            if ftb.height > 0.0 {
                let _ = writeln!(out, "  #v({}pt)", format_f64(ftb.height));
            }
            out.push_str("]\n");
        }
        WrapMode::Behind | WrapMode::InFront | WrapMode::None => {
            // Text-anchored boxes use this zero-size flow box to keep `#place`
            // tied to their paragraph. Page- and margin-anchored boxes resolve
            // an absolute page coordinate in `generate_floating_text_box_overlay`.
            out.push_str("#box(width: 0pt, height: 0pt)[\n");
            generate_floating_text_box_overlay(out, ftb, ctx, false, ftb.offset_y)?;
            out.push_str("]\n");
        }
        WrapMode::Square | WrapMode::Tight => {
            generate_floating_text_box_overlay(out, ftb, ctx, true, ftb.offset_y)?;
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

fn generate_drawing_group(
    out: &mut String,
    group: &crate::ir::DrawingGroup,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    if let Some(anchor) = &group.anchor {
        out.push_str("#box(width: 0pt, height: 0pt)[#context [\n");
        let margins = ctx.active_page_margins.unwrap_or_default();
        let x_origin = if anchor.horizontal_anchor == FrameAnchor::Margin {
            margins.left
        } else {
            0.0
        };
        let y_origin = if anchor.vertical_anchor == FrameAnchor::Margin {
            margins.top
        } else {
            0.0
        };
        let x = if anchor.horizontal_anchor == FrameAnchor::Text {
            format!("{}pt", format_f64(anchor.offset_x))
        } else {
            format!(
                "{}pt - here().position().x",
                format_f64(x_origin + anchor.offset_x)
            )
        };
        let y = if anchor.vertical_anchor == FrameAnchor::Text {
            format!("{}pt", format_f64(anchor.offset_y))
        } else {
            format!(
                "{}pt - here().position().y",
                format_f64(y_origin + anchor.offset_y)
            )
        };
        let _ = write!(out, "#place(top + left, dx: {x}, dy: {y}");
        if matches!(
            anchor.wrap_mode,
            crate::ir::WrapMode::TopAndBottom
                | crate::ir::WrapMode::Square
                | crate::ir::WrapMode::Tight
        ) {
            out.push_str(", float: true");
        }
        out.push_str(")[\n");
    }
    let _ = writeln!(
        out,
        "#block(width: {}pt, height: {}pt, breakable: false, above: 0pt, below: 0pt)[",
        format_f64(group.width),
        format_f64(group.height)
    );
    // Fixed DrawingML text calls these helpers even inside a flowing document.
    // Keep their definitions local so ordinary Word paragraphs are unaffected.
    write_powerpoint_ligature_state(out);
    write_powerpoint_advance_grid_helpers(out);
    for element in &group.elements {
        generate_fixed_element(out, element, group.height, ctx)?;
    }
    out.push_str("]\n");
    if group.anchor.is_some() {
        out.push_str("]\n]]\n");
    }
    Ok(())
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
    should_float: bool,
    flow_offset_y: f64,
) -> Result<(), ConvertError> {
    let aligned_page_x: Option<f64> = floating_text_box_aligned_page_x(ftb, ctx);
    let page_y: Option<f64> = floating_text_box_page_y(ftb, ctx);
    let needs_page_context: bool = aligned_page_x.is_some() || page_y.is_some();
    if needs_page_context {
        out.push_str("#context [\n");
        if aligned_page_x.is_some() {
            out.push_str("#let anchor_x = here().position().x\n");
        }
        if page_y.is_some() {
            out.push_str("#let anchor_y = here().position().y\n");
        }
    }
    let dx: String = match aligned_page_x {
        Some(target_x) => format!("{}pt - anchor_x", format_f64(target_x + ftb.offset_x)),
        None => format!("{}pt", format_f64(ftb.offset_x)),
    };
    let dy: String = match page_y {
        Some(target_y) => format!("{}pt - anchor_y", format_f64(target_y)),
        None => format!("{}pt", format_f64(flow_offset_y)),
    };
    let _ = write!(out, "#place(top + left, dx: {dx}, dy: {dy}");
    if should_float {
        out.push_str(", float: true");
    }
    out.push_str(")[\n");
    generate_floating_text_box_content(out, ftb, ctx)?;
    if needs_page_context {
        out.push_str("]\n]\n");
    } else {
        out.push_str("]\n");
    }
    Ok(())
}

fn floating_text_box_aligned_page_x(ftb: &FloatingTextBox, ctx: &GenCtx) -> Option<f64> {
    // Word aligns to a page frame, but Typst places at the anchor paragraph.
    // The generated context translates that flow location to this page-space x.
    let alignment: crate::ir::FrameAlign = ftb.horizontal_align?;
    let page_size: PageSize = ctx.active_page_size?;
    let margins: Margins = ctx.active_page_margins?;
    let (origin, available): (f64, f64) = match ftb.horizontal_anchor {
        FrameAnchor::Page => (0.0, page_size.width),
        FrameAnchor::Margin | FrameAnchor::Text => (
            margins.left,
            (page_size.width - margins.left - margins.right).max(0.0),
        ),
    };
    Some(origin + aligned_offset(Some(alignment), available, Some(ftb.width)))
}

fn floating_text_box_page_y(ftb: &FloatingTextBox, ctx: &GenCtx) -> Option<f64> {
    let page_size: PageSize = ctx.active_page_size?;
    let margins: Margins = ctx.active_page_margins?;
    let alignment: Option<crate::ir::FrameAlign> = ftb.vertical_position_align;
    let (origin, available): (f64, f64) = match ftb.vertical_anchor {
        FrameAnchor::Page => (0.0, page_size.height),
        FrameAnchor::Margin => (
            margins.top,
            (page_size.height - margins.top - margins.bottom).max(0.0),
        ),
        FrameAnchor::Text => return None,
    };
    let aligned_y: f64 = aligned_offset(alignment, available, Some(ftb.height));
    Some(
        origin
            + aligned_y
            + if alignment.is_none() {
                ftb.offset_y
            } else {
                0.0
            },
    )
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
    let _ = write!(
        out,
        "#box(width: {}pt, height: {}pt, inset: {}",
        format_f64(ftb.width),
        format_f64(ftb.height),
        inset,
    );
    if let Some(fill) = &ftb.fill {
        write_fill_color(out, fill, None);
    }
    write_shape_stroke(out, &ftb.stroke);
    out.push_str(")[\n");

    if matches!(ftb.vertical_align, TextBoxVerticalAlign::Top) {
        // Paragraph-first boxes need the correction for Typst's extra line-box
        // leading. A table carries its own top geometry, so its border stays
        // fixed while its expanded exact cell line is seated in the line box.
        let needs_top_leading_compensation: bool =
            matches!(ftb.content.first(), Some(Block::Paragraph(_)));
        if needs_top_leading_compensation {
            let _ = writeln!(
                out,
                "#place(top + left, dy: -{}pt)[",
                format_f64(FLOATING_TEXT_BOX_TOP_LEADING_COMPENSATION_PT),
            );
        }
        let was_in_top_aligned_floating_text_box: bool =
            std::mem::replace(&mut ctx.in_top_aligned_floating_text_box, true);
        let generation_result: Result<(), ConvertError> = (|| {
            let _ = writeln!(out, "#block(width: {}pt)[", format_f64(inner_width));
            for (index, block) in ftb.content.iter().enumerate() {
                if index > 0 {
                    out.push('\n');
                }
                generate_fixed_text_box_block(
                    out,
                    block,
                    ctx,
                    Some(inner_width),
                    false,
                    PowerPointBaselineMode::Disabled,
                )?;
            }
            out.push_str("]\n");
            Ok(())
        })();
        ctx.in_top_aligned_floating_text_box = was_in_top_aligned_floating_text_box;
        generation_result?;
        if needs_top_leading_compensation {
            out.push_str("]\n");
        }
        out.push_str("]\n");
        return Ok(());
    }

    let text_box_id: usize = ctx.next_text_box_id();
    let _ = writeln!(
        out,
        "#let floating_text_box_content_{text_box_id} = block(width: {}pt)[",
        format_f64(inner_width)
    );
    let was_in_top_aligned_floating_text_box: bool =
        std::mem::replace(&mut ctx.in_top_aligned_floating_text_box, false);
    let generation_result: Result<(), ConvertError> = (|| {
        for (index, block) in ftb.content.iter().enumerate() {
            if index > 0 {
                out.push('\n');
            }
            generate_fixed_text_box_block(
                out,
                block,
                ctx,
                Some(inner_width),
                false,
                PowerPointBaselineMode::Disabled,
            )?;
        }
        Ok(())
    })();
    ctx.in_top_aligned_floating_text_box = was_in_top_aligned_floating_text_box;
    generation_result?;
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

/// Width correction for the deterministic Noto Sans fallback used when Aptos
/// is unavailable. The expanded layout width and x-only paint scale keep the
/// fallback from adding a wrap or changing the line height.
const UNEMBEDDED_APTOS_WIDTH_SCALE: f64 = 0.916;

fn unembedded_aptos_width_scale(text_box: &TextBoxData) -> Option<f64> {
    if text_box.no_wrap
        || text_box.auto_fit
        || !matches!(text_box.vertical_align, TextBoxVerticalAlign::Top)
    {
        return None;
    }

    let [Block::Paragraph(paragraph)] = text_box.content.as_slice() else {
        return None;
    };
    if paragraph.runs.is_empty() || paragraph_has_forced_breaks(paragraph) {
        return None;
    }

    let has_only_plain_aptos_runs: bool = paragraph.runs.iter().all(|run| {
        run.footnote.is_none()
            && run.inline_box.is_none()
            && run
                .style
                .font_family
                .as_deref()
                .is_some_and(|family| family.trim().eq_ignore_ascii_case("Aptos"))
            && !matches!(run.style.bold, Some(true))
            && !matches!(run.style.italic, Some(true))
            && !matches!(run.style.small_caps, Some(true))
            && !matches!(run.style.all_caps, Some(true))
            && run
                .style
                .letter_spacing
                .is_none_or(|letter_spacing| letter_spacing == 0.0)
    });
    if !has_only_plain_aptos_runs
        || crate::render::font_subst::is_primary_font_available("Aptos")
        || !crate::render::font_subst::is_primary_font_available("Noto Sans")
    {
        return None;
    }

    Some(UNEMBEDDED_APTOS_WIDTH_SCALE)
}

fn wrapped_fit_paragraph(
    text_box: &TextBoxData,
    content_region_height_pt: f64,
) -> Option<&Paragraph> {
    // Wrapping and a short centered frame do not opt a file into smaller text
    // when it omitted `<a:normAutofit/>` (#1963).
    if !text_box.auto_fit
        || text_box.no_wrap
        || matches!(text_box.vertical_align, TextBoxVerticalAlign::Top)
    {
        return None;
    }
    // The scale this path emits is the content region over the block's own
    // height. A region that is zero or negative leaves nothing to fit and
    // scales the text to a point, so every word lands on one origin and none
    // of it is drawn. PowerPoint does not shrink a block it cannot fit either
    // way — it lets it overflow at its natural size (issue #1706).
    if content_region_height_pt <= 0.0 {
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
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) => {
            default_line_height_pt.max(points)
        }
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

fn common_powerpoint_list_baseline_snap(
    list: &List,
    mode: PowerPointBaselineMode,
) -> Option<PowerPointBaselineSnap> {
    let root_level: u32 = list.items.first()?.level;
    let mut snaps = list.items.iter().map(|item| {
        let [paragraph] = item.content.as_slice() else {
            return None;
        };
        (item.level == root_level)
            .then(|| mode.for_paragraph(&paragraph.runs, &paragraph.style))
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
    baseline_mode: PowerPointBaselineMode,
) -> Result<(), ConvertError> {
    match block {
        Block::List(list) if can_render_fixed_text_list_inline(list) => {
            generate_fixed_text_list(out, list, true, available_width_pt, true, baseline_mode)
        }
        Block::Paragraph(para) => {
            generate_fixed_text_paragraph(out, para, available_width_pt, no_wrap, baseline_mode)
        }
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
            let baseline_snap = common_powerpoint_list_baseline_snap(list, baseline_mode);
            generate_list_with_spacing_model(
                out,
                list,
                settings.as_deref(),
                true,
                ListEojeolWrap {
                    baseline_snap,
                    ..ListEojeolWrap::default()
                },
                ctx,
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
    baseline_mode: PowerPointBaselineMode,
) -> Result<(), ConvertError> {
    let style: &ParagraphStyle = &para.style;
    if let Some(line_height_pt) = powerpoint_blank_paragraph_height_pt(para) {
        let _ = writeln!(
            out,
            "#block(above: {}pt, below: {}pt, height: {}pt)[]",
            format_f64(style.space_before.unwrap_or(0.0)),
            format_f64(style.space_after.unwrap_or(0.0)),
            format_f64(line_height_pt),
        );
        return Ok(());
    }
    let inset: Insets = fixed_text_paragraph_inset(style);
    let has_inset: bool = inset.left > 0.0 || inset.right > 0.0;
    let hanging_indent_pt: Option<f64> = fixed_text_paragraph_hanging_indent_pt(style);
    let first_line_indent_pt: Option<f64> =
        style.indent_first_line.filter(|value| value.abs() > 0.0001);
    // PowerPoint's own line, which supersedes the `size * 0.65` leading this
    // path used to guess with (issue #513).
    let line_height_settings: Option<String> = powerpoint_line_height_settings(&para.runs, style);
    let baseline_snap: Option<PowerPointBaselineSnap> =
        baseline_mode.for_paragraph(&para.runs, style);
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
                    false,
                );
                write_common_text_settings(out, &para.runs, "  ");
                out.push_str(settings);
            }
            None => {
                write_par_settings(out, style, &para.runs, false);
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

    let needs_cjk_line_start_guard: bool =
        starts_with_cjk_line_start_shrinkable_punctuation(&para.runs);

    if no_wrap {
        out.push_str("#box[");
        if needs_cjk_line_start_guard {
            // Typst shortens some CJK punctuation at a line start. PowerPoint
            // keeps these glyphs on the text origin, so an empty inline box
            // keeps Typst from treating the punctuation as the first glyph
            // without adding anything to the searchable text layer (#1916).
            out.push_str("#box[]");
        }
        generate_runs_with_tabs_no_wrap(
            out,
            &para.runs,
            style.tab_stops.as_deref(),
            paragraph_default_tab_width_pt(style, DEFAULT_TAB_WIDTH_PT),
        );
    } else {
        if needs_cjk_line_start_guard {
            out.push_str("#box[]");
        }
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

fn starts_with_cjk_line_start_shrinkable_punctuation(runs: &[Run]) -> bool {
    for run in runs {
        if run.inline_box.is_some() {
            return false;
        }
        let Some(character) = run.text.chars().next() else {
            continue;
        };
        return matches!(
            character,
            '‘' | '“' | '《' | '（' | '『' | '「' | '【' | '〖' | '〔' | '〈' | '［' | '｛'
        );
    }
    false
}

/// The vertical advance of a DrawingML paragraph with no visible runs.
///
/// Each trailing `<a:br>` reserves the empty line it ends, and the paragraph
/// mark reserves the final empty line. Emitting both line heights explicitly
/// avoids Typst collapsing one of the blank physical lines in a text box. A
/// paragraph with no hard break uses only the mark's line height (#1917).
fn powerpoint_blank_paragraph_height_pt(paragraph: &Paragraph) -> Option<f64> {
    let mark_style: &TextStyle = paragraph.style.paragraph_mark_text_style.as_deref()?;
    let has_visible_content: bool = paragraph.runs.iter().any(|run| {
        run.footnote.is_some()
            || run.inline_box.is_some()
            || run
                .text
                .chars()
                .any(|character| character != '\u{000B}' && !character.is_whitespace())
    });
    if has_visible_content {
        return None;
    }

    let mut height_pt: f64 = 0.0;
    for run in &paragraph.runs {
        for character in run.text.chars() {
            if character == '\u{000B}' {
                height_pt += powerpoint_paragraph_line_advance_pt(&run.style, &paragraph.style);
            }
        }
    }
    height_pt += powerpoint_paragraph_line_advance_pt(mark_style, &paragraph.style);
    Some(height_pt)
}

fn powerpoint_paragraph_line_advance_pt(
    text_style: &TextStyle,
    paragraph_style: &ParagraphStyle,
) -> f64 {
    let metric_run: Run = Run {
        text: String::new(),
        style: text_style.clone(),
        href: None,
        footnote: None,
        inline_box: None,
    };
    let fallback_font_size_pt: f64 = text_style
        .font_size
        .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);
    let plain_line_height_pt: f64 = powerpoint_line_box_pt(std::slice::from_ref(&metric_run))
        .unwrap_or(crate::render::pdf::POWERPOINT_LINE_HEIGHT_FACTOR * fallback_font_size_pt);
    match paragraph_style.line_spacing {
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => plain_line_height_pt * factor,
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) if points > 0.0 => points,
        _ => plain_line_height_pt,
    }
}

#[cfg(test)]
#[path = "typst_gen_tests.rs"]
mod tests;

#[path = "typst_gen_section_alignment.rs"]
mod section_alignment;

#[path = "typst_gen_float_wrap.rs"]
mod float_wrap;
