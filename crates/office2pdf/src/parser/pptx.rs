use std::collections::{BTreeMap, HashMap};
#[cfg(test)]
use std::io::Cursor;
use std::io::Read;

use quick_xml::Reader;
use quick_xml::escape::unescape as unescape_xml_text;
use quick_xml::events::{BytesStart, Event};
use zip::ZipArchive;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, ArrowHead, BaselineShiftEm, Block, BorderLineStyle, BorderSide, CellBorder,
    CellVerticalAlign, Chart, Color, Document, FixedElement, FixedElementKind, FixedPage,
    GradientFill, ImageClipShape, ImageCrop, ImageData, ImageFormat, Insets, LineJoin, LineSpacing,
    List, ListItem, ListKind, ListLevelStyle, Page, PageSize, PairKerning, Paragraph,
    ParagraphStyle, PatternFill, PatternPreset, Run, Shadow, Shape, ShapeKind, SmartArt,
    SmartArtNode, StyleSheet, TabAlignment, TabLeader, TabStop, Table, TableBorderPaintModel,
    TableCell, TableRow, TextBoxData, TextBoxVerticalAlign, TextDirection, TextStyle, TopBevel,
};
use crate::parser::Parser;
use crate::parser::smartart;
use crate::parser::units::emu_to_pt;

use self::package::{
    load_master_color_map, load_table_styles, load_theme, parse_default_text_size_pt,
    parse_presentation_xml, parse_rels_xml, read_zip_entry,
};
#[cfg(test)]
use self::package::{resolve_relative_path, scan_chart_refs};
use self::shapes::{
    parse_arrow_head, parse_group_shape, parse_src_rect, pptx_dash_to_border_style,
    preset_text_rect_insets, prst_to_shape_kind,
};
use self::slides::{PresentationResources, SlideParseContext, parse_single_slide, parse_slide_xml};
use self::tables::{parse_pptx_table, scale_pptx_table_geometry_to_frame};
use self::text::*;
use self::theme::{
    ColorMapData, ParsedColor, PptxMasterTextStyles, ThemeData, ThemeLineStyle, default_color_map,
    drawingml_line_join, parse_background_color, parse_background_gradient,
    parse_background_image_rid, parse_background_ref, parse_color_from_empty,
    parse_color_from_start, parse_effect_list, parse_master_color_map, parse_master_text_styles,
    parse_shape_gradient_fill, parse_shape_pattern_fill, parse_theme_xml, resolve_effect_ref,
    resolve_effect_ref_top_bevel, resolve_effective_color_map, resolve_fill_ref,
    resolve_scheme_color, resolve_theme_font,
};

#[path = "pptx_custom_geometry.rs"]
pub(crate) mod custom_geometry;
#[path = "pptx_geometry_guides.rs"]
pub(crate) mod geometry_guides;
#[path = "pptx_package.rs"]
mod package;
#[path = "pptx_placeholders.rs"]
mod placeholders;
#[path = "pptx_shapes.rs"]
mod shapes;
#[path = "pptx_slides.rs"]
mod slides;
#[path = "pptx_table_styles.rs"]
mod table_styles;
#[path = "pptx_tables.rs"]
mod tables;
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
    /// Picture border color (inside `<p:pic>/<p:spPr>/<a:ln>`).
    PicLineFill,
}

#[derive(Debug, Clone)]
struct PptxParagraphEntry {
    paragraph: Paragraph,
    list_marker: Option<PptxListMarker>,
    /// Size of the invisible paragraph mark after inheritance. Visible runs
    /// determine percentage paragraph spacing when present; this is the
    /// fallback that lets an empty paragraph resolve the same rule.
    paragraph_mark_font_size_pt: Option<f64>,
}

const PPTX_DEFAULT_TEXT_BOX_LEFT_RIGHT_INSET_PT: f64 = 7.2;
const PPTX_DEFAULT_TEXT_BOX_TOP_BOTTOM_INSET_PT: f64 = 3.6;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';
/// Internal run marker for a DrawingML paragraph whose inherited list style
/// must not paint a bullet or number. The renderer turns this into a line-box
/// strut and never emits the character into Typst or the PDF text layer.
const PPTX_BLANK_LIST_ITEM_SENTINEL: &str = "\u{E000}\u{E001}";

fn pptx_paragraph_is_blank_list_item(paragraph: &Paragraph) -> bool {
    paragraph.runs.len() == 1
        && paragraph.runs[0].footnote.is_none()
        && paragraph.runs[0].text == PPTX_BLANK_LIST_ITEM_SENTINEL
}

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
    last_explicit_start_at: Option<u32>,
}

impl PendingPptxList {
    fn new(marker: &PptxListMarker) -> Self {
        Self {
            kind: marker.kind(),
            items: Vec::new(),
            level_styles: BTreeMap::new(),
            last_level: 0,
            last_explicit_start_at: None,
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
            if let Some(start_at) = auto_numbering.start_at
                && auto_numbering.level <= self.last_level
            {
                // PowerPoint renders consecutive same-level paragraphs with
                // the same explicit `startAt` as one sequence. Extend only
                // that verified form and preserve the existing restart
                // behavior for other explicit starts (issue #1347).
                let repeats_previous_marker: bool = auto_numbering.level == self.last_level
                    && self.last_explicit_start_at == Some(start_at);
                if !repeats_previous_marker {
                    return false;
                }
            }

            return self
                .level_styles
                .get(&auto_numbering.level)
                .is_none_or(|style| style.numbering_pattern == auto_numbering.numbering_pattern);
        }

        self.level_styles.get(&marker.level()).is_none_or(|style| {
            style.marker_text.as_deref() == marker.marker_text()
                && style.marker_style.as_ref() == marker.marker_style()
        })
    }

    fn push(&mut self, paragraph: Paragraph, marker: PptxListMarker) {
        let level: u32 = marker.level();
        let start_at: Option<u32> = marker.start_at();
        let numbering_pattern: Option<String> = marker.numbering_pattern().map(str::to_string);
        let marker_text: Option<String> = marker.marker_text().map(str::to_string);
        let marker_style: Option<TextStyle> = marker.marker_style().cloned();
        if !pptx_paragraph_is_blank_list_item(&paragraph) {
            self.level_styles
                .entry(level)
                .or_insert_with(|| ListLevelStyle {
                    kind: self.kind,
                    numbering_pattern,
                    full_numbering: false,
                    marker_text,
                    marker_style,
                });
        }
        self.items.push(ListItem {
            content: vec![paragraph],
            level,
            start_at: if self.items.is_empty() {
                start_at
            } else {
                None
            },
        });
        self.last_level = level;
        self.last_explicit_start_at = start_at;
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
        let mut style: ParagraphStyle = self.default_paragraph.clone();
        if let Some(level_style) = self.levels.get(&level) {
            style.merge_from(&level_style.paragraph);
        }
        style
    }

    fn run_style_for_level(&self, level: u32) -> TextStyle {
        let mut style: TextStyle = self.default_run.clone();
        if let Some(level_style) = self.levels.get(&level) {
            style.merge_from(&level_style.run);
        }
        style
    }

    fn bullet_for_level(&self, level: u32) -> PptxBulletDefinition {
        let mut bullet: PptxBulletDefinition = self.default_bullet.clone();
        if let Some(level_style) = self.levels.get(&level) {
            merge_pptx_bullet_definition(&mut bullet, &level_style.bullet);
        }
        bullet
    }

    /// Apply a default text color from `<p:style><a:fontRef>`.
    /// This overrides inherited layout/master defaults because `fontRef` is
    /// a shape-level style with higher precedence.
    fn apply_default_color(&mut self, color: Color) {
        self.default_run.color = Some(color);
        for level_style in self.levels.values_mut() {
            level_style.run.color = Some(color);
        }
    }

    /// Apply the theme face selected by a shape-level `<a:fontRef>`.
    fn apply_default_font_family(&mut self, font_family: &str) {
        self.default_run.font_family = Some(font_family.to_string());
        for level_style in self.levels.values_mut() {
            level_style.run.font_family = Some(font_family.to_string());
        }
    }

    fn merge_from(&mut self, overlay: &PptxTextBodyStyleDefaults) {
        self.default_paragraph
            .merge_from(&overlay.default_paragraph);
        self.default_run.merge_from(&overlay.default_run);
        merge_pptx_bullet_definition(&mut self.default_bullet, &overlay.default_bullet);

        for (level, overlay_style) in &overlay.levels {
            let target: &mut PptxTextLevelStyle = self.levels.entry(*level).or_default();
            target.paragraph.merge_from(&overlay_style.paragraph);
            target.run.merge_from(&overlay_style.run);
            merge_pptx_bullet_definition(&mut target.bullet, &overlay_style.bullet);
        }
    }
}

/// Parser for PPTX (Office Open XML PowerPoint) presentations.
pub struct PptxParser;

impl Parser for PptxParser {
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let mut archive = crate::parser::open_zip(data)?;
        // A face's declared family class does not vary by where it is named,
        // so one sweep of the package answers for the whole deck (issue #891).
        let declared_font_classes: std::collections::HashMap<String, crate::ir::DeclaredFontClass> =
            scan_package_font_classes(&mut archive);

        // Extract metadata from docProps/core.xml
        let metadata = crate::parser::metadata::extract_metadata_from_zip(&mut archive);

        // Read and parse presentation.xml for slide size and slide references
        let pres_xml = read_zip_entry(&mut archive, "ppt/presentation.xml")?;
        let (slide_size, slide_rids) = parse_presentation_xml(&pres_xml)?;
        // The size a text body falls back to when its own chain declares none
        // (issue #675). Read from the same XML rather than re-opening it.
        let default_text_size_pt: Option<f64> = parse_default_text_size_pt(&pres_xml);

        // Read and parse presentation.xml.rels for rId → slide path mapping
        let rels_xml = read_zip_entry(&mut archive, "ppt/_rels/presentation.xml.rels")?;
        let rel_map = parse_rels_xml(&rels_xml);

        // Load theme data (if available)
        let theme = load_theme(&rel_map, &mut archive);

        // Load table styles (uses theme colors for scheme color resolution)
        let master_color_map: ColorMapData = load_master_color_map(&rel_map, &mut archive);
        let table_styles: table_styles::TableStyleMap =
            load_table_styles(&mut archive, &theme, &master_color_map);

        let presentation_resources = PresentationResources {
            theme: &theme,
            table_styles: &table_styles,
            default_text_size_pt,
        };

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
                    slide_number,
                    slide_size,
                    &presentation_resources,
                    &mut archive,
                    options.include_hidden_slides,
                ) {
                    // Hidden slide (show="0" or show="false") omitted because
                    // include_hidden_slides is false, matching PowerPoint's
                    // default PDF export behavior.
                    Ok(None) => {}
                    Ok(Some((page, slide_warnings))) => {
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
                styles: StyleSheet {
                    declared_font_classes,
                    ..StyleSheet::default()
                },
            },
            warnings,
        ))
    }
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

// Re-export shared XML utilities so submodules can use `super::get_attr_str` etc.
use super::xml_util::get_attr_i64;
use super::xml_util::get_attr_str;
use super::xml_util::parse_hex_color;

#[cfg(test)]
#[path = "pptx_tests.rs"]
mod tests;

/// Sweep every DrawingML part in a package for faces that declare their family
/// class, so the substitution chooser can prefer the declaration over the name
/// (issue #891).
fn scan_package_font_classes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> std::collections::HashMap<String, crate::ir::DeclaredFontClass> {
    let mut classes = std::collections::HashMap::new();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|file| file.name().to_string())
        })
        .filter(|name| {
            name.ends_with(".xml")
                && (name.starts_with("ppt/slides/")
                    || name.starts_with("ppt/slideLayouts/")
                    || name.starts_with("ppt/slideMasters/")
                    || name.starts_with("ppt/theme/")
                    || name.starts_with("ppt/charts/")
                    || name.starts_with("ppt/notesSlides/"))
        })
        .collect();
    for name in names {
        if let Ok(xml) = self::package::read_zip_entry(archive, &name) {
            crate::parser::drawingml::scan_declared_font_classes(&xml, &mut classes);
        }
    }
    classes
}
