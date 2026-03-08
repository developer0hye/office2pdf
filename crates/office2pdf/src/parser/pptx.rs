use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read};

use quick_xml::Reader;
use quick_xml::escape::unescape as unescape_xml_text;
use quick_xml::events::{BytesStart, Event};
use zip::ZipArchive;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Chart, Color,
    Document, FixedElement, FixedElementKind, FixedPage, GradientFill, ImageCrop, ImageData,
    ImageFormat, Insets, LineSpacing, List, ListItem, ListKind, ListLevelStyle, Page, PageSize,
    Paragraph, ParagraphStyle, Run, Shadow, Shape, ShapeKind, SmartArt, SmartArtNode, StyleSheet,
    Table, TableCell, TableRow, TextBoxData, TextBoxVerticalAlign, TextDirection, TextStyle,
};
use crate::parser::Parser;
use crate::parser::smartart;

use self::package::{
    load_chart_data, load_slide_images, load_smartart_data, load_theme, parse_presentation_xml,
    parse_rels_xml, read_zip_entry, resolve_layout_master_paths, resolve_relative_path,
    scan_chart_refs,
};
use self::shapes::{
    parse_group_shape, parse_src_rect, pptx_dash_to_border_style, prst_to_shape_kind,
};
use self::text::*;
use self::theme::{
    ColorMapData, ThemeData, default_color_map, parse_background_color, parse_background_gradient,
    parse_color_from_empty, parse_color_from_start, parse_effect_list, parse_master_color_map,
    parse_master_other_style, parse_shape_gradient_fill, parse_theme_xml,
    resolve_effective_color_map, resolve_theme_font,
};

#[path = "pptx_package.rs"]
mod package;
#[path = "pptx_shapes.rs"]
mod shapes;
#[path = "pptx_text.rs"]
mod text;
#[path = "pptx_theme.rs"]
mod theme;

/// Relationship metadata from a `.rels` file.
#[derive(Debug, Clone)]
struct Relationship {
    target: String,
    rel_type: Option<String>,
}

/// Image asset referenced by a slide relationship.
#[derive(Debug, Clone)]
struct SlideImageAsset {
    path: String,
    data: Vec<u8>,
    source: SlideImageSource,
}

impl SlideImageAsset {
    fn format(&self) -> Option<ImageFormat> {
        match self.source {
            SlideImageSource::Supported(format) => Some(format),
            SlideImageSource::Unsupported => None,
        }
    }

    fn is_supported(&self) -> bool {
        matches!(self.source, SlideImageSource::Supported(_))
    }

    fn file_name(&self) -> &str {
        self.path.rsplit('/').next().unwrap_or(self.path.as_str())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SlideImageSource {
    Supported(ImageFormat),
    Unsupported,
}

/// Map from relationship ID → slide image asset.
type SlideImageMap = HashMap<String, SlideImageAsset>;

/// Context for which element a `<a:solidFill>` belongs to.
#[derive(Clone, Copy, PartialEq, Eq)]
enum SolidFillCtx {
    None,
    /// Fill color of the shape itself (inside `<p:spPr>`, not `<a:ln>`).
    ShapeFill,
    /// Stroke/border color (inside `<a:ln>`).
    LineFill,
    /// Text run color (inside `<a:rPr>`).
    RunFill,
    /// Paragraph end-run color (inside `<a:endParaRPr>`).
    EndParaFill,
    /// Bullet marker color (inside `<a:buClr>`).
    BulletFill,
}

#[derive(Debug, Clone)]
struct PptxParagraphEntry {
    paragraph: Paragraph,
    list_marker: Option<PptxListMarker>,
}

const PPTX_DEFAULT_TEXT_BOX_LEFT_RIGHT_INSET_PT: f64 = 7.2;
const PPTX_DEFAULT_TEXT_BOX_TOP_BOTTOM_INSET_PT: f64 = 3.6;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';

fn default_pptx_text_box_padding() -> Insets {
    Insets {
        top: PPTX_DEFAULT_TEXT_BOX_TOP_BOTTOM_INSET_PT,
        right: PPTX_DEFAULT_TEXT_BOX_LEFT_RIGHT_INSET_PT,
        bottom: PPTX_DEFAULT_TEXT_BOX_TOP_BOTTOM_INSET_PT,
        left: PPTX_DEFAULT_TEXT_BOX_LEFT_RIGHT_INSET_PT,
    }
}

fn default_pptx_table_cell_padding() -> Insets {
    default_pptx_text_box_padding()
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PptxAutoNumbering {
    level: u32,
    numbering_pattern: Option<String>,
    start_at: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum PptxBulletKind {
    None,
    Character(String),
    AutoNumber(PptxAutoNumbering),
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum PptxBulletFontSource {
    FollowText,
    Explicit(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum PptxBulletColorSource {
    FollowText,
    Explicit(Color),
}

#[derive(Debug, Clone, PartialEq)]
enum PptxBulletSizeSource {
    FollowText,
    Percent(f64),
    Points(f64),
}

#[derive(Debug, Clone, Default)]
struct PptxBulletDefinition {
    kind: Option<PptxBulletKind>,
    font: Option<PptxBulletFontSource>,
    color: Option<PptxBulletColorSource>,
    size: Option<PptxBulletSizeSource>,
}

#[derive(Debug, Clone)]
enum PptxListMarker {
    Ordered {
        auto_numbering: PptxAutoNumbering,
        marker_style: Option<TextStyle>,
    },
    Unordered {
        level: u32,
        marker_text: String,
        marker_style: Option<TextStyle>,
    },
}

impl PptxListMarker {
    fn kind(&self) -> ListKind {
        match self {
            Self::Ordered { .. } => ListKind::Ordered,
            Self::Unordered { .. } => ListKind::Unordered,
        }
    }

    fn level(&self) -> u32 {
        match self {
            Self::Ordered { auto_numbering, .. } => auto_numbering.level,
            Self::Unordered { level, .. } => *level,
        }
    }

    fn numbering_pattern(&self) -> Option<&str> {
        match self {
            Self::Ordered { auto_numbering, .. } => auto_numbering.numbering_pattern.as_deref(),
            Self::Unordered { .. } => None,
        }
    }

    fn start_at(&self) -> Option<u32> {
        match self {
            Self::Ordered { auto_numbering, .. } => auto_numbering.start_at,
            Self::Unordered { .. } => None,
        }
    }

    fn marker_text(&self) -> Option<&str> {
        match self {
            Self::Ordered { .. } => None,
            Self::Unordered { marker_text, .. } => Some(marker_text),
        }
    }

    fn marker_style(&self) -> Option<&TextStyle> {
        match self {
            Self::Ordered { marker_style, .. } | Self::Unordered { marker_style, .. } => {
                marker_style.as_ref()
            }
        }
    }
}

#[derive(Debug, Clone)]
struct PendingPptxList {
    kind: ListKind,
    items: Vec<ListItem>,
    level_styles: BTreeMap<u32, ListLevelStyle>,
    last_level: u32,
}

impl PendingPptxList {
    fn new(marker: &PptxListMarker) -> Self {
        Self {
            kind: marker.kind(),
            items: Vec::new(),
            level_styles: BTreeMap::new(),
            last_level: 0,
        }
    }

    fn can_extend(&self, marker: &PptxListMarker) -> bool {
        if self.kind != marker.kind() {
            return false;
        }

        if self.items.is_empty() {
            return true;
        }

        if let PptxListMarker::Ordered { auto_numbering, .. } = marker {
            if auto_numbering.start_at.is_some() && auto_numbering.level <= self.last_level {
                return false;
            }

            return self
                .level_styles
                .get(&auto_numbering.level)
                .is_none_or(|style| {
                    style.numbering_pattern == auto_numbering.numbering_pattern
                        && style.marker_style.as_ref() == marker.marker_style()
                });
        }

        self.level_styles.get(&marker.level()).is_none_or(|style| {
            style.marker_text.as_deref() == marker.marker_text()
                && style.marker_style.as_ref() == marker.marker_style()
        })
    }

    fn push(&mut self, paragraph: Paragraph, marker: PptxListMarker) {
        let level: u32 = marker.level();
        let numbering_pattern: Option<String> = marker.numbering_pattern().map(str::to_string);
        let marker_text: Option<String> = marker.marker_text().map(str::to_string);
        let marker_style: Option<TextStyle> = marker.marker_style().cloned();
        self.level_styles
            .entry(level)
            .or_insert_with(|| ListLevelStyle {
                kind: self.kind,
                numbering_pattern,
                full_numbering: false,
                marker_text,
                marker_style,
            });
        self.items.push(ListItem {
            content: vec![paragraph],
            level,
            start_at: if self.items.is_empty() {
                marker.start_at()
            } else {
                None
            },
        });
        self.last_level = level;
    }

    fn into_block(self) -> Block {
        Block::List(List {
            kind: self.kind,
            items: self.items,
            level_styles: self.level_styles,
        })
    }
}

#[derive(Debug, Clone, Default)]
struct PptxTextLevelStyle {
    paragraph: ParagraphStyle,
    run: TextStyle,
    bullet: PptxBulletDefinition,
}

#[derive(Debug, Clone, Default)]
struct PptxTextBodyStyleDefaults {
    default_paragraph: ParagraphStyle,
    default_run: TextStyle,
    default_bullet: PptxBulletDefinition,
    levels: BTreeMap<u32, PptxTextLevelStyle>,
}

impl PptxTextBodyStyleDefaults {
    fn paragraph_style_for_level(&self, level: u32) -> ParagraphStyle {
        let mut style = self.default_paragraph.clone();
        if let Some(level_style) = self.levels.get(&level) {
            merge_paragraph_style(&mut style, &level_style.paragraph);
        }
        style
    }

    fn run_style_for_level(&self, level: u32) -> TextStyle {
        let mut style = self.default_run.clone();
        if let Some(level_style) = self.levels.get(&level) {
            merge_text_style(&mut style, &level_style.run);
        }
        style
    }

    fn bullet_for_level(&self, level: u32) -> PptxBulletDefinition {
        let mut bullet = self.default_bullet.clone();
        if let Some(level_style) = self.levels.get(&level) {
            merge_pptx_bullet_definition(&mut bullet, &level_style.bullet);
        }
        bullet
    }

    fn merge_from(&mut self, overlay: &PptxTextBodyStyleDefaults) {
        merge_paragraph_style(&mut self.default_paragraph, &overlay.default_paragraph);
        merge_text_style(&mut self.default_run, &overlay.default_run);
        merge_pptx_bullet_definition(&mut self.default_bullet, &overlay.default_bullet);

        for (level, overlay_style) in &overlay.levels {
            let target = self.levels.entry(*level).or_default();
            merge_paragraph_style(&mut target.paragraph, &overlay_style.paragraph);
            merge_text_style(&mut target.run, &overlay_style.run);
            merge_pptx_bullet_definition(&mut target.bullet, &overlay_style.bullet);
        }
    }
}

/// Parser for PPTX (Office Open XML PowerPoint) presentations.
pub struct PptxParser;

/// Convert EMU (English Metric Units) to points.
/// 1 inch = 914400 EMU, 1 inch = 72 points, so 1 pt = 12700 EMU.
fn emu_to_pt(emu: i64) -> f64 {
    emu as f64 / 12700.0
}

impl Parser for PptxParser {
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)
            .map_err(|e| ConvertError::Parse(format!("Failed to read PPTX: {e}")))?;

        // Extract metadata from docProps/core.xml
        let metadata = crate::parser::metadata::extract_metadata_from_zip(&mut archive);

        // Read and parse presentation.xml for slide size and slide references
        let pres_xml = read_zip_entry(&mut archive, "ppt/presentation.xml")?;
        let (slide_size, slide_rids) = parse_presentation_xml(&pres_xml)?;

        // Read and parse presentation.xml.rels for rId → slide path mapping
        let rels_xml = read_zip_entry(&mut archive, "ppt/_rels/presentation.xml.rels")?;
        let rel_map = parse_rels_xml(&rels_xml);

        // Load theme data (if available)
        let theme = load_theme(&rel_map, &mut archive);

        let mut warnings = Vec::new();

        // Parse each slide in order, skipping broken slides with warnings
        let mut pages = Vec::with_capacity(slide_rids.len());
        for (slide_idx, rid) in slide_rids.iter().enumerate() {
            // Filter by slide range if specified (1-indexed)
            let slide_number = (slide_idx as u32) + 1;
            if let Some(ref range) = options.slide_range
                && !range.contains(slide_number)
            {
                continue;
            }

            if let Some(target) = rel_map.get(rid) {
                let slide_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("ppt/{target}")
                };

                let slide_label = format!("slide {slide_number}");
                match parse_single_slide(
                    &slide_path,
                    &slide_label,
                    slide_size,
                    &theme,
                    &mut archive,
                ) {
                    Ok((page, slide_warnings)) => {
                        warnings.extend(slide_warnings);
                        // Emit structured warnings for fallback-rendered elements
                        if let Page::Fixed(ref fp) = page {
                            for elem in &fp.elements {
                                match &elem.kind {
                                    FixedElementKind::Chart(chart) => {
                                        let title = chart
                                            .title
                                            .as_deref()
                                            .unwrap_or("untitled")
                                            .to_string();
                                        warnings.push(ConvertWarning::FallbackUsed {
                                            format: "PPTX".to_string(),
                                            from: format!("chart ({title})"),
                                            to: "data table".to_string(),
                                        });
                                    }
                                    FixedElementKind::SmartArt(_) => {
                                        warnings.push(ConvertWarning::FallbackUsed {
                                            format: "PPTX".to_string(),
                                            from: "SmartArt diagram".to_string(),
                                            to: "text list".to_string(),
                                        });
                                    }
                                    _ => {}
                                }
                            }
                        }
                        pages.push(page);
                    }
                    Err(e) => {
                        warnings.push(ConvertWarning::ParseSkipped {
                            format: "PPTX".to_string(),
                            reason: format!(
                                "slide {} ({}) failed to parse: {e}",
                                slide_idx + 1,
                                slide_path
                            ),
                        });
                    }
                }
            }
        }

        Ok((
            Document {
                metadata,
                pages,
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

/// Parse a single slide from the archive, returning a Page or an error.
///
/// Resolves the inheritance chain (slide → layout → master) and
/// prepends master/layout elements behind slide elements.
fn parse_single_slide<R: Read + std::io::Seek>(
    slide_path: &str,
    slide_label: &str,
    slide_size: PageSize,
    theme: &ThemeData,
    archive: &mut ZipArchive<R>,
) -> Result<(Page, Vec<ConvertWarning>), ConvertError> {
    let slide_xml = read_zip_entry(archive, slide_path)?;
    let (layout_path, master_path) = resolve_layout_master_paths(slide_path, archive);
    let master_xml = master_path
        .as_ref()
        .and_then(|path| read_zip_entry(archive, path).ok());
    let layout_xml = layout_path
        .as_ref()
        .and_then(|path| read_zip_entry(archive, path).ok());
    let master_color_map = master_xml
        .as_deref()
        .map(parse_master_color_map)
        .unwrap_or_else(default_color_map);
    let master_text_style_defaults = master_xml
        .as_deref()
        .map(|xml| parse_master_other_style(xml, theme, &master_color_map))
        .unwrap_or_default();
    let slide_color_map = resolve_effective_color_map(&slide_xml, &master_color_map);
    let layout_color_map = layout_xml
        .as_deref()
        .map(|xml| resolve_effective_color_map(xml, &master_color_map));

    let slide_images = load_slide_images(slide_path, archive);
    let mut warnings = Vec::new();
    let (slide_elements, slide_warnings) = parse_slide_xml(
        &slide_xml,
        &slide_images,
        theme,
        &slide_color_map,
        slide_label,
        &master_text_style_defaults,
    )?;
    warnings.extend(slide_warnings);

    // Build element list: master (behind) → layout → slide (on top)
    let mut elements = Vec::new();

    // Master elements (furthest back)
    if let Some(ref path) = master_path
        && let Some(xml) = master_xml.as_deref()
    {
        let master_images = load_slide_images(path, archive);
        let master_label = format!("{slide_label} master");
        if let Ok((master_elements, master_warnings)) = parse_slide_xml(
            xml,
            &master_images,
            theme,
            &master_color_map,
            &master_label,
            &master_text_style_defaults,
        ) {
            elements.extend(master_elements);
            warnings.extend(master_warnings);
        }
    }

    // Layout elements (middle layer)
    if let Some(ref path) = layout_path
        && let Some(xml) = layout_xml.as_deref()
        && let Some(color_map) = layout_color_map.as_ref()
    {
        let layout_images = load_slide_images(path, archive);
        let layout_label = format!("{slide_label} layout");
        if let Ok((layout_elements, layout_warnings)) = parse_slide_xml(
            xml,
            &layout_images,
            theme,
            color_map,
            &layout_label,
            &master_text_style_defaults,
        ) {
            elements.extend(layout_elements);
            warnings.extend(layout_warnings);
        }
    }

    // Slide elements (on top)
    elements.extend(slide_elements);

    // SmartArt diagrams: scan for dgm:relIds in graphicFrames,
    // resolve data files from .rels, and add as SmartArt elements.
    let smartart_refs = smartart::scan_smartart_refs(&slide_xml);
    if !smartart_refs.is_empty() {
        let smartart_data = load_smartart_data(slide_path, archive);
        for sa_ref in &smartart_refs {
            if let Some(items) = smartart_data.get(&sa_ref.data_rid) {
                elements.push(FixedElement {
                    x: emu_to_pt(sa_ref.x),
                    y: emu_to_pt(sa_ref.y),
                    width: emu_to_pt(sa_ref.cx),
                    height: emu_to_pt(sa_ref.cy),
                    kind: FixedElementKind::SmartArt(SmartArt {
                        items: items.clone(),
                    }),
                });
            }
        }
    }

    // Embedded charts: scan for c:chart references in graphicFrames,
    // resolve chart XML files from .rels, and add as Chart elements.
    let chart_refs = scan_chart_refs(&slide_xml);
    if !chart_refs.is_empty() {
        let chart_data = load_chart_data(slide_path, archive);
        for c_ref in &chart_refs {
            if let Some(chart) = chart_data.get(&c_ref.chart_rid) {
                elements.push(FixedElement {
                    x: emu_to_pt(c_ref.x),
                    y: emu_to_pt(c_ref.y),
                    width: emu_to_pt(c_ref.cx),
                    height: emu_to_pt(c_ref.cy),
                    kind: FixedElementKind::Chart(chart.clone()),
                });
            }
        }
    }

    // Resolve background: try gradient first, then solid color.
    // Reuse already-resolved layout/master paths to avoid re-reading .rels files.
    let background_gradient = parse_background_gradient(&slide_xml, theme, &slide_color_map);
    let background_color = if background_gradient.is_some() {
        // When gradient is present, also extract first stop as fallback color
        background_gradient
            .as_ref()
            .and_then(|g| g.stops.first().map(|s| s.color))
    } else {
        parse_background_color(&slide_xml, theme, &slide_color_map)
            .or_else(|| {
                layout_xml.as_deref().and_then(|xml| {
                    layout_color_map
                        .as_ref()
                        .and_then(|map| parse_background_color(xml, theme, map))
                })
            })
            .or_else(|| {
                master_xml
                    .as_deref()
                    .and_then(|xml| parse_background_color(xml, theme, &master_color_map))
            })
    };

    Ok((
        Page::Fixed(FixedPage {
            size: slide_size,
            elements,
            background_color,
            background_gradient,
        }),
        warnings,
    ))
}

/// Map from relationship ID → list of SmartArt nodes with hierarchy depth.
type SmartArtMap = HashMap<String, Vec<SmartArtNode>>;

/// Reference to a chart found in a slide's graphicFrame.
struct ChartRef {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    chart_rid: String,
}

/// Map from relationship ID → parsed Chart data.
type ChartMap = HashMap<String, Chart>;

/// Parse a `<a:tbl>` element from the reader into a Table IR.
///
/// The reader should be positioned right after the `<a:tbl>` Start event.
/// Reads until the matching `</a:tbl>` End event.
fn parse_pptx_table(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Result<Table, ConvertError> {
    let mut column_widths = Vec::new();
    let mut rows: Vec<TableRow> = Vec::new();

    // Current row state
    let mut in_row = false;
    let mut row_height_emu: i64 = 0;
    let mut cells: Vec<TableCell> = Vec::new();

    // Current cell state
    let mut in_cell = false;
    let mut cell_col_span: u32 = 1;
    let mut cell_row_span: u32 = 1;
    let mut is_h_merge = false;
    let mut is_v_merge = false;
    let mut cell_text_entries: Vec<PptxParagraphEntry> = Vec::new();
    let mut cell_background: Option<Color> = None;
    let mut cell_vertical_align: Option<CellVerticalAlign> = None;
    let mut cell_padding: Option<Insets> = None;

    // Text parsing state (reused per cell)
    let mut in_txbody = false;
    let mut text_body_style_defaults = PptxTextBodyStyleDefaults::default();
    let mut in_para = false;
    let mut para_style = ParagraphStyle::default();
    let mut para_level: u32 = 0;
    let mut para_default_run_style = TextStyle::default();
    let mut para_end_run_style = TextStyle::default();
    let mut para_bullet_definition = PptxBulletDefinition::default();
    let mut in_ln_spc = false;
    let mut runs: Vec<Run> = Vec::new();
    let mut in_run = false;
    let mut run_style = TextStyle::default();
    let mut run_text = String::new();
    let mut in_text = false;
    let mut in_rpr = false;
    let mut in_end_para_rpr = false;
    let mut solid_fill_ctx = SolidFillCtx::None;

    // Cell property state
    let mut in_tc_pr = false;
    let mut border_left: Option<BorderSide> = None;
    let mut border_right: Option<BorderSide> = None;
    let mut border_top: Option<BorderSide> = None;
    let mut border_bottom: Option<BorderSide> = None;
    let mut in_border_ln = false;
    let mut border_ln_width_emu: i64 = 0;
    let mut border_ln_color: Option<Color> = None;
    let mut border_ln_dash_style: BorderLineStyle = BorderLineStyle::Solid;
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum BorderDir {
        None,
        Left,
        Right,
        Top,
        Bottom,
    }
    let mut current_border_dir = BorderDir::None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"gridCol" => {
                        // PowerPoint may attach extLst children to gridCol, so width
                        // must be captured on the start tag as well as self-closing form.
                        if let Some(w) = get_attr_i64(e, b"w") {
                            column_widths.push(emu_to_pt(w));
                        }
                    }
                    b"tr" => {
                        in_row = true;
                        row_height_emu = get_attr_i64(e, b"h").unwrap_or(0);
                        cells.clear();
                    }
                    b"tc" if in_row => {
                        in_cell = true;
                        cell_col_span = get_attr_i64(e, b"gridSpan").map(|v| v as u32).unwrap_or(1);
                        cell_row_span = get_attr_i64(e, b"rowSpan").map(|v| v as u32).unwrap_or(1);
                        is_h_merge = get_attr_str(e, b"hMerge").is_some();
                        is_v_merge = get_attr_str(e, b"vMerge").is_some();
                        cell_text_entries.clear();
                        cell_background = None;
                        cell_vertical_align = None;
                        cell_padding = None;
                        in_tc_pr = false;
                        border_left = None;
                        border_right = None;
                        border_top = None;
                        border_bottom = None;
                    }
                    b"txBody" if in_cell => {
                        in_txbody = true;
                        text_body_style_defaults = PptxTextBodyStyleDefaults::default();
                    }
                    b"lstStyle" if in_txbody => {
                        let local_defaults = parse_pptx_list_style(reader, theme, color_map);
                        text_body_style_defaults.merge_from(&local_defaults);
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_level = 0;
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        in_ln_spc = false;
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"r" if in_para => {
                        in_run = true;
                        run_style = para_default_run_style.clone();
                        run_text.clear();
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        in_end_para_rpr = true;
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"solidFill" if in_rpr => {
                        solid_fill_ctx = SolidFillCtx::RunFill;
                    }
                    b"solidFill" if in_end_para_rpr => {
                        solid_fill_ctx = SolidFillCtx::EndParaFill;
                    }
                    b"solidFill" if in_tc_pr && !in_border_ln => {
                        solid_fill_ctx = SolidFillCtx::ShapeFill;
                    }
                    b"solidFill" if in_border_ln => {
                        solid_fill_ctx = SolidFillCtx::LineFill;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let color = parse_color_from_start(reader, e, theme, color_map).color;
                        match solid_fill_ctx {
                            SolidFillCtx::ShapeFill => cell_background = color,
                            SolidFillCtx::LineFill => border_ln_color = color,
                            SolidFillCtx::RunFill => run_style.color = color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }
                    b"tcPr" if in_cell => {
                        in_tc_pr = true;
                        extract_pptx_table_cell_props(
                            e,
                            &mut cell_vertical_align,
                            &mut cell_padding,
                        );
                    }
                    b"lnL" if in_tc_pr => {
                        in_border_ln = true;
                        current_border_dir = BorderDir::Left;
                        border_ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_ln_color = None;
                        border_ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnR" if in_tc_pr => {
                        in_border_ln = true;
                        current_border_dir = BorderDir::Right;
                        border_ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_ln_color = None;
                        border_ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnT" if in_tc_pr => {
                        in_border_ln = true;
                        current_border_dir = BorderDir::Top;
                        border_ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_ln_color = None;
                        border_ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnB" if in_tc_pr => {
                        in_border_ln = true;
                        current_border_dir = BorderDir::Bottom;
                        border_ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_ln_color = None;
                        border_ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"prstDash" if in_border_ln => {
                        border_ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"gridCol" => {
                        if let Some(w) = get_attr_i64(e, b"w") {
                            column_widths.push(emu_to_pt(w));
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let color = parse_color_from_empty(e, theme, color_map).color;
                        match solid_fill_ctx {
                            SolidFillCtx::ShapeFill => cell_background = color,
                            SolidFillCtx::LineFill => border_ln_color = color,
                            SolidFillCtx::RunFill => run_style.color = color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }
                    b"prstDash" if in_border_ln => {
                        border_ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"tcPr" if in_cell => {
                        extract_pptx_table_cell_props(
                            e,
                            &mut cell_vertical_align,
                            &mut cell_padding,
                        );
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"latin" | b"ea" | b"cs" if in_rpr => {
                        apply_typeface_to_style(e, &mut run_style, theme);
                    }
                    b"latin" | b"ea" | b"cs" if in_end_para_rpr => {
                        apply_typeface_to_style(e, &mut para_end_run_style, theme);
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_text && let Some(text) = decode_pptx_text_event(t) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(ref reference)) => {
                if in_text && let Some(text) = decode_pptx_general_ref(reference) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"tbl" => break,
                    b"tr" if in_row => {
                        let height = if row_height_emu > 0 {
                            Some(emu_to_pt(row_height_emu))
                        } else {
                            None
                        };
                        rows.push(TableRow {
                            cells: std::mem::take(&mut cells),
                            height,
                        });
                        in_row = false;
                    }
                    b"tc" if in_cell => {
                        let has_border = border_left.is_some()
                            || border_right.is_some()
                            || border_top.is_some()
                            || border_bottom.is_some();

                        let (col_span, row_span) = if is_h_merge {
                            (0, 1)
                        } else if is_v_merge {
                            (1, 0)
                        } else {
                            (cell_col_span, cell_row_span)
                        };

                        cells.push(TableCell {
                            content: group_pptx_text_blocks(std::mem::take(&mut cell_text_entries)),
                            col_span,
                            row_span,
                            border: if has_border {
                                Some(CellBorder {
                                    left: border_left.take(),
                                    right: border_right.take(),
                                    top: border_top.take(),
                                    bottom: border_bottom.take(),
                                })
                            } else {
                                None
                            },
                            background: cell_background.take(),
                            data_bar: None,
                            icon_text: None,
                            vertical_align: cell_vertical_align.take(),
                            padding: cell_padding.take(),
                        });
                        in_cell = false;
                        in_tc_pr = false;
                    }
                    b"txBody" if in_txbody => {
                        in_txbody = false;
                    }
                    b"p" if in_para => {
                        let resolved_list_marker = resolve_pptx_list_marker(
                            &para_bullet_definition,
                            para_level,
                            &runs,
                            &para_end_run_style,
                            &para_default_run_style,
                        );
                        let paragraph_runs = std::mem::take(&mut runs);
                        cell_text_entries.push(PptxParagraphEntry {
                            paragraph: Paragraph {
                                style: para_style.clone(),
                                runs: paragraph_runs,
                            },
                            list_marker: resolved_list_marker,
                        });
                        in_para = false;
                    }
                    b"r" if in_run => {
                        if !run_text.is_empty() {
                            push_pptx_run(
                                &mut runs,
                                Run {
                                    text: std::mem::take(&mut run_text),
                                    style: run_style.clone(),
                                    href: None,
                                    footnote: None,
                                },
                            );
                        }
                        in_run = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                    }
                    b"endParaRPr" if in_end_para_rpr => {
                        in_end_para_rpr = false;
                    }
                    b"lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    b"solidFill" if solid_fill_ctx != SolidFillCtx::None => {
                        solid_fill_ctx = SolidFillCtx::None;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }
                    b"tcPr" if in_tc_pr => {
                        in_tc_pr = false;
                    }
                    b"lnL" | b"lnR" | b"lnT" | b"lnB" if in_border_ln => {
                        if let Some(color) = border_ln_color.take() {
                            let side = BorderSide {
                                width: border_ln_width_emu as f64 / 12700.0,
                                color,
                                style: border_ln_dash_style,
                            };
                            match current_border_dir {
                                BorderDir::Left => border_left = Some(side),
                                BorderDir::Right => border_right = Some(side),
                                BorderDir::Top => border_top = Some(side),
                                BorderDir::Bottom => border_bottom = Some(side),
                                BorderDir::None => {}
                            }
                        }
                        in_border_ln = false;
                        current_border_dir = BorderDir::None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ConvertError::Parse(format!("XML error in table: {e}"))),
            _ => {}
        }
    }

    Ok(Table {
        rows,
        column_widths,
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(default_pptx_table_cell_padding()),
        use_content_driven_row_heights: true,
    })
}

fn scale_pptx_table_geometry_to_frame(
    table: &mut Table,
    frame_width_pt: f64,
    frame_height_pt: f64,
) {
    let intrinsic_width_pt: f64 = table.column_widths.iter().sum();
    if intrinsic_width_pt > 0.0 && frame_width_pt > 0.0 {
        let x_scale: f64 = frame_width_pt / intrinsic_width_pt;
        for width in &mut table.column_widths {
            *width *= x_scale;
        }
    }

    let intrinsic_height_pt: f64 = table.rows.iter().filter_map(|row| row.height).sum();
    if intrinsic_height_pt > 0.0 && frame_height_pt > 0.0 {
        let y_scale: f64 = frame_height_pt / intrinsic_height_pt;
        for row in &mut table.rows {
            if let Some(height) = row.height.as_mut() {
                *height *= y_scale;
            }
        }
    }
}

fn describe_assets(assets: impl IntoIterator<Item = String>) -> String {
    assets.into_iter().collect::<Vec<_>>().join(", ")
}

fn pick_supported_asset(rid: &str, images: &SlideImageMap) -> Option<SlideImageAsset> {
    images
        .get(rid)
        .filter(|asset| asset.is_supported())
        .cloned()
}

fn select_picture_asset(
    images: &SlideImageMap,
    warning_context: &str,
    base_rid: Option<&str>,
    svg_rid: Option<&str>,
    img_layer_rids: &[String],
) -> (Option<SlideImageAsset>, Vec<ConvertWarning>) {
    let mut warnings = Vec::new();

    let unsupported_layers: Vec<String> = img_layer_rids
        .iter()
        .filter_map(|rid| images.get(rid))
        .filter(|asset| !asset.is_supported())
        .map(|asset| asset.file_name().to_string())
        .collect();
    if !unsupported_layers.is_empty() {
        warnings.push(ConvertWarning::PartialElement {
            format: "PPTX".to_string(),
            element: format!("{warning_context} picture"),
            detail: format!(
                "unsupported image layer omitted: {}",
                describe_assets(unsupported_layers)
            ),
        });
    }

    let selected = svg_rid
        .and_then(|rid| pick_supported_asset(rid, images))
        .or_else(|| base_rid.and_then(|rid| pick_supported_asset(rid, images)))
        .or_else(|| {
            img_layer_rids
                .iter()
                .find_map(|rid| pick_supported_asset(rid, images))
        });
    if selected.is_some() {
        return (selected, warnings);
    }

    let omitted_assets = svg_rid
        .into_iter()
        .chain(base_rid)
        .chain(img_layer_rids.iter().map(String::as_str))
        .filter_map(|rid| images.get(rid))
        .map(|asset| asset.file_name().to_string())
        .collect::<Vec<_>>();
    if !omitted_assets.is_empty() {
        warnings.push(ConvertWarning::UnsupportedElement {
            format: "PPTX".to_string(),
            element: format!(
                "{warning_context} image omitted: {}",
                describe_assets(omitted_assets)
            ),
        });
    }

    (None, warnings)
}

/// Parse a slide XML to extract positioned elements (text boxes, shapes, images).
fn parse_slide_xml(
    xml: &str,
    images: &SlideImageMap,
    theme: &ThemeData,
    color_map: &ColorMapData,
    warning_context: &str,
    inherited_text_body_defaults: &PptxTextBodyStyleDefaults,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut elements = Vec::new();
    let mut warnings = Vec::new();

    // ── Shape-level state ────────────────────────────────────────────────
    let mut in_shape = false;
    let mut shape_depth: usize = 0;
    let mut shape_x: i64 = 0;
    let mut shape_y: i64 = 0;
    let mut shape_cx: i64 = 0;
    let mut shape_cy: i64 = 0;
    let mut shape_has_placeholder = false;

    // Shape property state (geometry, fill, border)
    let mut in_sp_pr = false;
    let mut prst_geom: Option<String> = None;
    let mut shape_fill: Option<Color> = None;
    let mut shape_gradient_fill: Option<GradientFill> = None;
    let mut in_ln = false;
    let mut ln_width_emu: i64 = 0;
    let mut ln_color: Option<Color> = None;
    let mut ln_dash_style: BorderLineStyle = BorderLineStyle::Solid;
    let mut shape_rotation_deg: Option<f64> = None;
    let mut shape_opacity: Option<f64> = None;
    let mut shape_shadow: Option<Shadow> = None;

    // Transform state (for shapes)
    let mut in_xfrm = false;

    // Text body state
    let mut in_txbody = false;
    let mut paragraphs: Vec<PptxParagraphEntry> = Vec::new();
    let mut text_box_padding: Insets = default_pptx_text_box_padding();
    let mut text_box_vertical_align: TextBoxVerticalAlign = TextBoxVerticalAlign::Top;
    let mut text_body_style_defaults = PptxTextBodyStyleDefaults::default();

    // Paragraph state
    let mut in_para = false;
    let mut para_style = ParagraphStyle::default();
    let mut para_level: u32 = 0;
    let mut para_default_run_style = TextStyle::default();
    let mut para_end_run_style = TextStyle::default();
    let mut para_bullet_definition = PptxBulletDefinition::default();
    let mut in_ln_spc = false;
    let mut runs: Vec<Run> = Vec::new();

    // Run state
    let mut in_run = false;
    let mut run_style = TextStyle::default();
    let mut run_text = String::new();

    // Sub-element state
    let mut in_text = false;
    let mut in_rpr = false;
    let mut in_end_para_rpr = false;
    let mut solid_fill_ctx = SolidFillCtx::None;

    // ── Picture-level state ──────────────────────────────────────────────
    let mut in_pic = false;
    let mut pic_x: i64 = 0;
    let mut pic_y: i64 = 0;
    let mut pic_cx: i64 = 0;
    let mut pic_cy: i64 = 0;
    let mut blip_embed: Option<String> = None;
    let mut svg_blip_embed: Option<String> = None;
    let mut img_layer_embeds: Vec<String> = Vec::new();
    let mut pic_crop: Option<ImageCrop> = None;
    let mut in_pic_xfrm = false;

    // ── GraphicFrame-level state (for tables and SmartArt) ─────────────
    let mut in_graphic_frame = false;
    let mut gf_x: i64 = 0;
    let mut gf_y: i64 = 0;
    let mut gf_cx: i64 = 0;
    let mut gf_cy: i64 = 0;
    let mut in_gf_xfrm = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── GraphicFrame start ───────────────────────────
                    b"graphicFrame" if !in_shape && !in_pic && !in_graphic_frame => {
                        in_graphic_frame = true;
                        gf_x = 0;
                        gf_y = 0;
                        gf_cx = 0;
                        gf_cy = 0;
                        in_gf_xfrm = false;
                    }
                    b"xfrm" if in_graphic_frame && !in_shape => {
                        in_gf_xfrm = true;
                    }
                    b"tbl" if in_graphic_frame => {
                        // Parse the table and emit as a FixedElement
                        if let Ok(mut table) = parse_pptx_table(&mut reader, theme, color_map) {
                            scale_pptx_table_geometry_to_frame(
                                &mut table,
                                emu_to_pt(gf_cx),
                                emu_to_pt(gf_cy),
                            );
                            elements.push(FixedElement {
                                x: emu_to_pt(gf_x),
                                y: emu_to_pt(gf_y),
                                width: emu_to_pt(gf_cx),
                                height: emu_to_pt(gf_cy),
                                kind: FixedElementKind::Table(table),
                            });
                        }
                    }
                    // ── Group shape start ────────────────────────────
                    b"grpSp" if !in_shape && !in_pic && !in_graphic_frame => {
                        if let Ok((group_elems, group_warnings)) = parse_group_shape(
                            &mut reader,
                            xml,
                            images,
                            theme,
                            color_map,
                            warning_context,
                            inherited_text_body_defaults,
                        ) {
                            elements.extend(group_elems);
                            warnings.extend(group_warnings);
                        }
                    }

                    // ── Shape start ──────────────────────────────────
                    b"sp" if !in_shape && !in_pic => {
                        in_shape = true;
                        shape_depth = 1;
                        shape_x = 0;
                        shape_y = 0;
                        shape_cx = 0;
                        shape_cy = 0;
                        shape_has_placeholder = false;
                        in_sp_pr = false;
                        prst_geom = None;
                        shape_fill = None;
                        shape_gradient_fill = None;
                        in_ln = false;
                        ln_width_emu = 0;
                        ln_color = None;
                        shape_rotation_deg = None;
                        shape_opacity = None;
                        shape_shadow = None;
                        in_txbody = false;
                        paragraphs.clear();
                        text_box_padding = default_pptx_text_box_padding();
                        text_box_vertical_align = TextBoxVerticalAlign::Top;
                    }
                    b"sp" if in_shape => {
                        shape_depth += 1;
                    }

                    // ── Shape properties ─────────────────────────────
                    b"spPr" if in_shape && !in_txbody => {
                        in_sp_pr = true;
                    }
                    b"xfrm" if in_shape && in_sp_pr => {
                        in_xfrm = true;
                        // rot attribute: rotation in 60000ths of a degree
                        if let Some(rot) = get_attr_i64(e, b"rot") {
                            shape_rotation_deg = Some(rot as f64 / 60_000.0);
                        }
                    }
                    b"prstGeom" if in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            prst_geom = Some(prst);
                        }
                    }
                    b"solidFill" if in_sp_pr && !in_ln && !in_rpr => {
                        solid_fill_ctx = SolidFillCtx::ShapeFill;
                    }
                    b"gradFill" if in_sp_pr && !in_ln && !in_rpr => {
                        // Parse gradient fill using sub-parser; reader is consumed up to </gradFill>
                        shape_gradient_fill =
                            parse_shape_gradient_fill(&mut reader, theme, color_map);
                        // Also set solid fill as fallback (first stop color)
                        if let Some(ref gf) = shape_gradient_fill
                            && shape_fill.is_none()
                        {
                            shape_fill = gf.stops.first().map(|s| s.color);
                        }
                    }
                    b"effectLst" if in_sp_pr && !in_ln => {
                        shape_shadow = parse_effect_list(&mut reader, theme, color_map);
                    }
                    b"ln" if in_sp_pr => {
                        in_ln = true;
                        ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"prstDash" if in_ln => {
                        ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    b"solidFill" if in_ln => {
                        solid_fill_ctx = SolidFillCtx::LineFill;
                    }
                    b"ph" if in_shape => {
                        shape_has_placeholder = true;
                    }

                    // ── Text body ────────────────────────────────────
                    b"txBody" if in_shape => {
                        in_txbody = true;
                        text_body_style_defaults = if shape_has_placeholder {
                            PptxTextBodyStyleDefaults::default()
                        } else {
                            inherited_text_body_defaults.clone()
                        };
                    }
                    b"bodyPr" if in_shape && in_txbody => {
                        extract_pptx_text_box_body_props(
                            e,
                            &mut text_box_padding,
                            &mut text_box_vertical_align,
                        );
                    }
                    b"lstStyle" if in_shape && in_txbody => {
                        let local_defaults = parse_pptx_list_style(&mut reader, theme, color_map);
                        text_body_style_defaults.merge_from(&local_defaults);
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_level = 0;
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        in_ln_spc = false;
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"r" if in_para => {
                        in_run = true;
                        run_style = para_default_run_style.clone();
                        run_text.clear();
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        in_end_para_rpr = true;
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"solidFill" if in_rpr => {
                        solid_fill_ctx = SolidFillCtx::RunFill;
                    }
                    b"solidFill" if in_end_para_rpr => {
                        solid_fill_ctx = SolidFillCtx::EndParaFill;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let parsed = parse_color_from_start(&mut reader, e, theme, color_map);
                        match solid_fill_ctx {
                            SolidFillCtx::ShapeFill => {
                                shape_fill = parsed.color;
                                if let Some(alpha) = parsed.alpha {
                                    shape_opacity = Some(alpha);
                                }
                            }
                            SolidFillCtx::LineFill => ln_color = parsed.color,
                            SolidFillCtx::RunFill => run_style.color = parsed.color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = parsed.color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    parsed.color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }

                    // ── Picture start ────────────────────────────────
                    b"pic" if !in_shape && !in_pic => {
                        in_pic = true;
                        pic_x = 0;
                        pic_y = 0;
                        pic_cx = 0;
                        pic_cy = 0;
                        blip_embed = None;
                        svg_blip_embed = None;
                        img_layer_embeds.clear();
                        pic_crop = None;
                        in_pic_xfrm = false;
                    }
                    b"spPr" if in_pic => {
                        // Re-use nothing — just mark for xfrm detection below
                    }
                    b"xfrm" if in_pic => {
                        in_pic_xfrm = true;
                    }
                    b"blipFill" if in_pic => {}
                    b"blip" if in_pic => {
                        blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"svgBlip" if in_pic => {
                        svg_blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"imgLayer" if in_pic => {
                        if let Some(rid) = get_attr_str(e, b"r:embed") {
                            img_layer_embeds.push(rid);
                        }
                    }
                    b"srcRect" if in_pic => {
                        pic_crop = parse_src_rect(e);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── Shape xfrm offset/extent ─────────────────────
                    b"off" if in_xfrm => {
                        shape_x = get_attr_i64(e, b"x").unwrap_or(0);
                        shape_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_xfrm => {
                        shape_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        shape_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }

                    // ── Picture xfrm offset/extent ───────────────────
                    b"off" if in_pic_xfrm => {
                        pic_x = get_attr_i64(e, b"x").unwrap_or(0);
                        pic_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_pic_xfrm => {
                        pic_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        pic_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }

                    // ── GraphicFrame xfrm offset/extent ─────────────
                    b"off" if in_gf_xfrm => {
                        gf_x = get_attr_i64(e, b"x").unwrap_or(0);
                        gf_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_gf_xfrm => {
                        gf_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        gf_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }

                    // ── Blip (image reference) ───────────────────────
                    b"blip" if in_pic => {
                        blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"svgBlip" if in_pic => {
                        svg_blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"imgLayer" if in_pic => {
                        if let Some(rid) = get_attr_str(e, b"r:embed") {
                            img_layer_embeds.push(rid);
                        }
                    }
                    b"srcRect" if in_pic => {
                        pic_crop = parse_src_rect(e);
                    }

                    // ── Preset geometry (empty element) ──────────────
                    b"prstGeom" if in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            prst_geom = Some(prst);
                        }
                    }

                    // ── Line element (empty, no fill children) ───────
                    b"ln" if in_sp_pr => {
                        ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                    }

                    // ── Dash pattern (often self-closing) ───────────
                    b"prstDash" if in_ln => {
                        ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }

                    // ── Color value ──────────────────────────────────
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let parsed = parse_color_from_empty(e, theme, color_map);
                        match solid_fill_ctx {
                            SolidFillCtx::ShapeFill => {
                                shape_fill = parsed.color;
                                if let Some(alpha) = parsed.alpha {
                                    shape_opacity = Some(alpha);
                                }
                            }
                            SolidFillCtx::LineFill => ln_color = parsed.color,
                            SolidFillCtx::RunFill => run_style.color = parsed.color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = parsed.color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    parsed.color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }

                    // ── Run properties (empty element) ───────────────
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"latin" | b"ea" | b"cs" if in_rpr => {
                        apply_typeface_to_style(e, &mut run_style, theme);
                    }
                    b"latin" | b"ea" | b"cs" if in_end_para_rpr => {
                        apply_typeface_to_style(e, &mut para_end_run_style, theme);
                    }

                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_text && let Some(text) = decode_pptx_text_event(t) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(ref reference)) => {
                if in_text && let Some(text) = decode_pptx_general_ref(reference) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── Shape end ────────────────────────────────────
                    b"sp" if in_shape => {
                        shape_depth -= 1;
                        if shape_depth == 0 {
                            let has_text = paragraphs
                                .iter()
                                .any(|entry| !entry.paragraph.runs.is_empty());

                            if has_text {
                                // TextBox — has visible text content
                                let blocks: Vec<Block> =
                                    group_pptx_text_blocks(std::mem::take(&mut paragraphs));
                                elements.push(FixedElement {
                                    x: emu_to_pt(shape_x),
                                    y: emu_to_pt(shape_y),
                                    width: emu_to_pt(shape_cx),
                                    height: emu_to_pt(shape_cy),
                                    kind: FixedElementKind::TextBox(TextBoxData {
                                        content: blocks,
                                        padding: text_box_padding,
                                        vertical_align: text_box_vertical_align,
                                    }),
                                });
                            } else if let Some(ref geom) = prst_geom {
                                // Shape — no text, but has geometry
                                let kind = prst_to_shape_kind(
                                    geom,
                                    emu_to_pt(shape_cx),
                                    emu_to_pt(shape_cy),
                                );
                                let stroke = ln_color.map(|color| BorderSide {
                                    width: ln_width_emu as f64 / 12700.0,
                                    color,
                                    style: ln_dash_style,
                                });
                                elements.push(FixedElement {
                                    x: emu_to_pt(shape_x),
                                    y: emu_to_pt(shape_y),
                                    width: emu_to_pt(shape_cx),
                                    height: emu_to_pt(shape_cy),
                                    kind: FixedElementKind::Shape(Shape {
                                        kind,
                                        fill: shape_fill,
                                        gradient_fill: shape_gradient_fill.take(),
                                        stroke,
                                        rotation_deg: shape_rotation_deg,
                                        opacity: shape_opacity,
                                        shadow: shape_shadow.take(),
                                    }),
                                });
                            }
                            in_shape = false;
                        }
                    }

                    // ── Shape sub-elements end ───────────────────────
                    b"spPr" if in_sp_pr => {
                        in_sp_pr = false;
                    }
                    b"xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    b"ln" if in_ln => {
                        in_ln = false;
                    }
                    b"txBody" if in_txbody => {
                        in_txbody = false;
                    }
                    b"p" if in_para => {
                        let resolved_list_marker = resolve_pptx_list_marker(
                            &para_bullet_definition,
                            para_level,
                            &runs,
                            &para_end_run_style,
                            &para_default_run_style,
                        );
                        let paragraph_runs = std::mem::take(&mut runs);
                        paragraphs.push(PptxParagraphEntry {
                            paragraph: Paragraph {
                                style: para_style.clone(),
                                runs: paragraph_runs,
                            },
                            list_marker: resolved_list_marker,
                        });
                        in_para = false;
                    }
                    b"r" if in_run => {
                        if !run_text.is_empty() {
                            push_pptx_run(
                                &mut runs,
                                Run {
                                    text: std::mem::take(&mut run_text),
                                    style: run_style.clone(),
                                    href: None,
                                    footnote: None,
                                },
                            );
                        }
                        in_run = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                    }
                    b"endParaRPr" if in_end_para_rpr => {
                        in_end_para_rpr = false;
                    }
                    b"lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    b"solidFill" if solid_fill_ctx != SolidFillCtx::None => {
                        solid_fill_ctx = SolidFillCtx::None;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }

                    // ── Picture end ──────────────────────────────────
                    b"pic" if in_pic => {
                        let (selected_asset, picture_warnings) = select_picture_asset(
                            images,
                            warning_context,
                            blip_embed.as_deref(),
                            svg_blip_embed.as_deref(),
                            &img_layer_embeds,
                        );
                        warnings.extend(picture_warnings);
                        if let Some(asset) = selected_asset
                            && let Some(format) = asset.format()
                        {
                            elements.push(FixedElement {
                                x: emu_to_pt(pic_x),
                                y: emu_to_pt(pic_y),
                                width: emu_to_pt(pic_cx),
                                height: emu_to_pt(pic_cy),
                                kind: FixedElementKind::Image(ImageData {
                                    data: asset.data.clone(),
                                    format,
                                    width: Some(emu_to_pt(pic_cx)),
                                    height: Some(emu_to_pt(pic_cy)),
                                    crop: pic_crop,
                                }),
                            });
                        }
                        in_pic = false;
                    }
                    b"xfrm" if in_pic_xfrm => {
                        in_pic_xfrm = false;
                    }

                    // ── GraphicFrame end ─────────────────────────────
                    b"graphicFrame" if in_graphic_frame => {
                        in_graphic_frame = false;
                    }
                    b"xfrm" if in_gf_xfrm => {
                        in_gf_xfrm = false;
                    }

                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ConvertError::Parse(format!("XML error in slide: {e}"))),
            _ => {}
        }
    }

    Ok((elements, warnings))
}

/// Get a string attribute value from an XML element.
/// Matches on full qualified name first (e.g. `r:id`), then local name.
fn get_attr_str(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key || attr.key.local_name().as_ref() == key {
            return attr.unescape_value().ok().map(|v| v.to_string());
        }
    }
    None
}

/// Get an i64 attribute value from an XML element.
fn get_attr_i64(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<i64> {
    get_attr_str(e, key).and_then(|v| v.parse().ok())
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

#[cfg(test)]
#[path = "pptx_tests.rs"]
mod tests;
