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
    Document, FixedElement, FixedElementKind, FixedPage, GradientFill, GradientStop, ImageCrop,
    ImageData, ImageFormat, Insets, LineSpacing, List, ListItem, ListKind, ListLevelStyle, Page,
    PageSize, Paragraph, ParagraphStyle, Run, Shadow, Shape, ShapeKind, SmartArt, SmartArtNode,
    StyleSheet, Table, TableCell, TableRow, TextBoxData, TextBoxVerticalAlign, TextDirection,
    TextStyle,
};
use crate::parser::Parser;
use crate::parser::chart as chart_parser;
use crate::parser::smartart;

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

/// Parsed theme data from ppt/theme/theme1.xml.
#[derive(Debug, Clone, Default)]
struct ThemeData {
    /// Color scheme: scheme name (e.g., "dk1", "accent1") → Color.
    colors: HashMap<String, Color>,
    /// Major (heading) font family name.
    major_font: Option<String>,
    /// Minor (body) font family name.
    minor_font: Option<String>,
}

/// Effective scheme-color aliases for a slide part.
#[derive(Debug, Clone, Default)]
struct ColorMapData {
    aliases: HashMap<String, String>,
}

#[derive(Debug, Clone, Copy, Default)]
struct ParsedColor {
    color: Option<Color>,
    alpha: Option<f64>,
}

#[derive(Debug, Clone, Copy)]
enum ColorTransform {
    LumMod(f64),
    LumOff(f64),
}

const COLOR_MAP_KEYS: &[&str] = &[
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// Convert EMU (English Metric Units) to points.
/// 1 inch = 914400 EMU, 1 inch = 72 points, so 1 pt = 12700 EMU.
fn emu_to_pt(emu: i64) -> f64 {
    emu as f64 / 12700.0
}

fn default_color_map() -> ColorMapData {
    let aliases = COLOR_MAP_KEYS
        .iter()
        .map(|name| ((*name).to_string(), (*name).to_string()))
        .collect();
    ColorMapData { aliases }
}

fn parse_color_map_attrs(element: &BytesStart<'_>) -> ColorMapData {
    let mut aliases = HashMap::new();
    for key in COLOR_MAP_KEYS {
        if let Some(target) = get_attr_str(element, key.as_bytes()) {
            aliases.insert((*key).to_string(), target);
        }
    }

    if aliases.is_empty() {
        default_color_map()
    } else {
        ColorMapData { aliases }
    }
}

fn parse_master_color_map(xml: &str) -> ColorMapData {
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"clrMap" =>
            {
                return parse_color_map_attrs(e);
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    default_color_map()
}

fn parse_master_other_style(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PptxTextBodyStyleDefaults {
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"otherStyle" => {
                return parse_pptx_list_style(&mut reader, theme, color_map);
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    PptxTextBodyStyleDefaults::default()
}

fn parse_color_map_override(xml: &str) -> Option<ColorMapData> {
    let mut reader = Reader::from_str(xml);
    let mut in_override = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"clrMapOvr" =>
            {
                in_override = true;
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if in_override && e.local_name().as_ref() == b"masterClrMapping" =>
            {
                return None;
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if in_override
                    && (e.local_name().as_ref() == b"overrideClrMapping"
                        || e.local_name().as_ref() == b"clrMap") =>
            {
                return Some(parse_color_map_attrs(e));
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"clrMapOvr" => {
                in_override = false;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    None
}

fn resolve_effective_color_map(xml: &str, master_color_map: &ColorMapData) -> ColorMapData {
    parse_color_map_override(xml).unwrap_or_else(|| master_color_map.clone())
}

fn resolve_scheme_color(
    theme: &ThemeData,
    color_map: &ColorMapData,
    scheme_name: &str,
) -> Option<Color> {
    let mapped_name = color_map
        .aliases
        .get(scheme_name)
        .map(String::as_str)
        .unwrap_or(scheme_name);

    theme
        .colors
        .get(mapped_name)
        .copied()
        .or_else(|| theme.colors.get(scheme_name).copied())
}

fn parse_base_color(
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<Color> {
    match element.local_name().as_ref() {
        b"srgbClr" => get_attr_str(element, b"val").and_then(|hex| parse_hex_color(&hex)),
        b"schemeClr" => get_attr_str(element, b"val")
            .and_then(|name| resolve_scheme_color(theme, color_map, &name)),
        b"sysClr" => get_attr_str(element, b"lastClr").and_then(|hex| parse_hex_color(&hex)),
        _ => None,
    }
}

fn parse_color_transform(element: &BytesStart<'_>) -> Option<ColorTransform> {
    let val = get_attr_i64(element, b"val")? as f64 / 100_000.0;
    match element.local_name().as_ref() {
        b"lumMod" => Some(ColorTransform::LumMod(val)),
        b"lumOff" => Some(ColorTransform::LumOff(val)),
        _ => None,
    }
}

fn apply_color_transforms(color: Color, transforms: &[ColorTransform]) -> Color {
    let (mut hue, mut saturation, mut lightness) = rgb_to_hsl(color);

    for transform in transforms {
        match transform {
            ColorTransform::LumMod(value) => {
                lightness = (lightness * value).clamp(0.0, 1.0);
            }
            ColorTransform::LumOff(value) => {
                lightness = (lightness + value).clamp(0.0, 1.0);
            }
        }
    }

    saturation = saturation.clamp(0.0, 1.0);
    hue = hue.rem_euclid(360.0);
    hsl_to_rgb(hue, saturation, lightness)
}

fn parse_color_from_empty(
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> ParsedColor {
    ParsedColor {
        color: parse_base_color(element, theme, color_map),
        alpha: None,
    }
}

fn parse_color_from_start(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> ParsedColor {
    let base_color = parse_base_color(element, theme, color_map);
    let mut transforms: Vec<ColorTransform> = Vec::new();
    let mut alpha: Option<f64> = None;
    let mut depth: usize = 1;

    while depth > 0 {
        match reader.read_event() {
            Ok(Event::Start(ref child)) => {
                depth += 1;
                if let Some(transform) = parse_color_transform(child) {
                    transforms.push(transform);
                } else if child.local_name().as_ref() == b"alpha" {
                    alpha = get_attr_i64(child, b"val").map(|v| v as f64 / 100_000.0);
                }
            }
            Ok(Event::Empty(ref child)) => {
                if let Some(transform) = parse_color_transform(child) {
                    transforms.push(transform);
                } else if child.local_name().as_ref() == b"alpha" {
                    alpha = get_attr_i64(child, b"val").map(|v| v as f64 / 100_000.0);
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let color = base_color.map(|base| apply_color_transforms(base, &transforms));

    ParsedColor { color, alpha }
}

fn rgb_to_hsl(color: Color) -> (f64, f64, f64) {
    let red = color.r as f64 / 255.0;
    let green = color.g as f64 / 255.0;
    let blue = color.b as f64 / 255.0;

    let max = red.max(green.max(blue));
    let min = red.min(green.min(blue));
    let delta = max - min;
    let lightness = (max + min) / 2.0;

    if delta == 0.0 {
        return (0.0, 0.0, lightness);
    }

    let saturation = delta / (1.0 - (2.0 * lightness - 1.0).abs());
    let hue_sector = if max == red {
        ((green - blue) / delta).rem_euclid(6.0)
    } else if max == green {
        ((blue - red) / delta) + 2.0
    } else {
        ((red - green) / delta) + 4.0
    };

    (60.0 * hue_sector, saturation, lightness)
}

fn hsl_to_rgb(hue: f64, saturation: f64, lightness: f64) -> Color {
    if saturation == 0.0 {
        let channel = (lightness * 255.0).round() as u8;
        return Color::new(channel, channel, channel);
    }

    let chroma = (1.0 - (2.0 * lightness - 1.0).abs()) * saturation;
    let hue_prime = hue / 60.0;
    let secondary = chroma * (1.0 - ((hue_prime.rem_euclid(2.0)) - 1.0).abs());
    let match_lightness = lightness - chroma / 2.0;

    let (red, green, blue) = match hue_prime {
        h if (0.0..1.0).contains(&h) => (chroma, secondary, 0.0),
        h if (1.0..2.0).contains(&h) => (secondary, chroma, 0.0),
        h if (2.0..3.0).contains(&h) => (0.0, chroma, secondary),
        h if (3.0..4.0).contains(&h) => (0.0, secondary, chroma),
        h if (4.0..5.0).contains(&h) => (secondary, 0.0, chroma),
        _ => (chroma, 0.0, secondary),
    };

    let to_u8 = |value: f64| ((value + match_lightness).clamp(0.0, 1.0) * 255.0).round() as u8;

    Color::new(to_u8(red), to_u8(green), to_u8(blue))
}

/// Map OOXML preset dash values to `BorderLineStyle`.
fn pptx_dash_to_border_style(val: &str) -> BorderLineStyle {
    match val {
        "dash" | "lgDash" | "sysDash" => BorderLineStyle::Dashed,
        "dot" | "sysDot" | "lgDashDot" => BorderLineStyle::Dotted,
        "dashDot" | "sysDashDot" => BorderLineStyle::DashDot,
        "lgDashDotDot" | "sysDashDotDot" => BorderLineStyle::DashDotDot,
        "solid" => BorderLineStyle::Solid,
        _ => BorderLineStyle::Solid,
    }
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

/// Build the .rels path for a given file path.
///
/// e.g., `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`
fn rels_path_for(path: &str) -> String {
    if let Some((dir, filename)) = path.rsplit_once('/') {
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{path}.rels")
    }
}

/// Resolve the layout and master file paths from a slide's .rels.
///
/// Returns `(Option<layout_path>, Option<master_path>)`.
fn resolve_layout_master_paths<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> (Option<String>, Option<String>) {
    let slide_dir = slide_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");

    // Read slide .rels to find layout
    let Ok(rels_xml) = read_zip_entry(archive, &rels_path_for(slide_path)) else {
        return (None, None);
    };
    let rels = parse_rels_xml(&rels_xml);

    let layout_path = rels
        .values()
        .find(|t| t.contains("slideLayout") || t.contains("slideLayouts"))
        .map(|target| resolve_relative_path(slide_dir, target));

    let Some(ref layout_path) = layout_path else {
        return (None, None);
    };

    // Read layout .rels to find master
    let layout_dir = layout_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
    let master_path = read_zip_entry(archive, &rels_path_for(layout_path))
        .ok()
        .and_then(|layout_rels_xml| {
            let layout_rels = parse_rels_xml(&layout_rels_xml);
            layout_rels
                .values()
                .find(|t| t.contains("slideMaster") || t.contains("slideMasters"))
                .map(|target| resolve_relative_path(layout_dir, target))
        });

    (Some(layout_path.clone()), master_path)
}

/// Resolve inherited background color from layout or master.
///
/// Reads the slide's .rels to find the layout, then the layout's .rels to find the master.
/// Returns the first background color found in the inheritance chain.
/// Resolve inherited background color from pre-resolved layout/master paths.
/// This avoids re-reading .rels files that were already parsed in `parse_single_slide`.
/// Read a file from the ZIP archive as a UTF-8 string.
fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String, ConvertError> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::Parse(format!("Missing {path} in PPTX: {e}")))?;
    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::Parse(format!("Failed to read {path}: {e}")))?;
    Ok(content)
}

/// Load images referenced by a slide from its .rels file and the ZIP archive.
///
/// Reads `<slide-dir>/_rels/<slide-filename>.rels`, finds image relationships,
/// and loads the corresponding image bytes from the ZIP.
fn load_slide_images<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> SlideImageMap {
    let mut images = SlideImageMap::new();

    // Build .rels path: ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
    let slide_rels_path = if let Some((dir, filename)) = slide_path.rsplit_once('/') {
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{slide_path}.rels")
    };

    let rels_xml = match read_zip_entry(archive, &slide_rels_path) {
        Ok(xml) => xml,
        Err(_) => return images, // No .rels file → no images
    };

    let slide_dir = slide_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
    let rels = parse_relationships_xml(&rels_xml);

    for (id, rel) in rels {
        if !is_image_relationship(rel.rel_type.as_deref(), &rel.target) {
            continue;
        }

        // Resolve relative path (e.g., "../media/image1.png" → "ppt/media/image1.png")
        let image_path = if let Some(stripped) = rel.target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, &rel.target)
        };

        if let Ok(mut file) = archive.by_name(&image_path) {
            let mut data = Vec::new();
            if file.read_to_end(&mut data).is_ok() {
                let source = image_format_from_ext(&rel.target)
                    .map(SlideImageSource::Supported)
                    .unwrap_or(SlideImageSource::Unsupported);
                images.insert(
                    id,
                    SlideImageAsset {
                        path: image_path,
                        data,
                        source,
                    },
                );
            }
        }
    }

    images
}

/// Map from relationship ID → list of SmartArt nodes with hierarchy depth.
type SmartArtMap = HashMap<String, Vec<SmartArtNode>>;

/// Pre-load SmartArt diagram data for a slide by scanning its .rels file
/// for diagram/data relationships and parsing the data XML files.
fn load_smartart_data<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> SmartArtMap {
    let mut map = SmartArtMap::new();

    let rels_xml = match read_zip_entry(archive, &rels_path_for(slide_path)) {
        Ok(xml) => xml,
        Err(_) => return map,
    };

    let slide_dir = slide_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");

    // parse_rels_xml gives Id→Target; scan for targets pointing to diagrams/data
    let rels = parse_rels_xml(&rels_xml);
    for (id, target) in &rels {
        if !target.contains("diagrams/data") && !target.contains("diagram/data") {
            continue;
        }
        let data_path = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, target)
        };
        if let Ok(data_xml) = read_zip_entry(archive, &data_path) {
            let texts = smartart::parse_smartart_data_xml(&data_xml);
            if !texts.is_empty() {
                map.insert(id.clone(), texts);
            }
        }
    }

    map
}

/// Reference to a chart found in a slide's graphicFrame.
struct ChartRef {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    chart_rid: String,
}

/// Scan slide XML for chart references within graphicFrame elements.
fn scan_chart_refs(slide_xml: &str) -> Vec<ChartRef> {
    let mut refs = Vec::new();
    let mut reader = Reader::from_str(slide_xml);

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
                    b"graphicFrame" if !in_graphic_frame => {
                        in_graphic_frame = true;
                        gf_x = 0;
                        gf_y = 0;
                        gf_cx = 0;
                        gf_cy = 0;
                        in_gf_xfrm = false;
                    }
                    b"xfrm" if in_graphic_frame => {
                        in_gf_xfrm = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"off" if in_gf_xfrm => {
                        gf_x = get_attr_i64(e, b"x").unwrap_or(0);
                        gf_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_gf_xfrm => {
                        gf_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        gf_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }
                    b"chart" if in_graphic_frame => {
                        if let Some(rid) = get_attr_str(e, b"r:id") {
                            refs.push(ChartRef {
                                x: gf_x,
                                y: gf_y,
                                cx: gf_cx,
                                cy: gf_cy,
                                chart_rid: rid,
                            });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
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
            Err(_) => break,
            _ => {}
        }
    }

    refs
}

/// Map from relationship ID → parsed Chart data.
type ChartMap = HashMap<String, Chart>;

/// Load chart data referenced by a slide from its .rels file and the ZIP archive.
fn load_chart_data<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> ChartMap {
    let mut charts = ChartMap::new();

    let rels_xml = match read_zip_entry(archive, &rels_path_for(slide_path)) {
        Ok(xml) => xml,
        Err(_) => return charts,
    };

    let slide_dir = slide_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
    let rel_map = parse_rels_xml(&rels_xml);

    for (id, target) in &rel_map {
        let lower = target.to_lowercase();
        if !lower.contains("chart") {
            continue;
        }
        if lower.contains("chartstyle") || lower.contains("chartcolor") {
            continue;
        }

        let chart_path = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, target)
        };

        if let Ok(chart_xml) = read_zip_entry(archive, &chart_path)
            && let Some(chart) = chart_parser::parse_chart_xml(&chart_xml)
        {
            charts.insert(id.clone(), chart);
        }
    }

    charts
}

/// Resolve a relative path against a base directory.
/// e.g., base="ppt/slides", relative="../media/image1.png" → "ppt/media/image1.png"
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    let mut parts: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for component in relative.split('/') {
        match component {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Determine image format from file extension, or None if not a recognized image.
fn image_format_from_ext(path: &str) -> Option<ImageFormat> {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        Some(ImageFormat::Png)
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some(ImageFormat::Jpeg)
    } else if lower.ends_with(".gif") {
        Some(ImageFormat::Gif)
    } else if lower.ends_with(".bmp") {
        Some(ImageFormat::Bmp)
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        Some(ImageFormat::Tiff)
    } else if lower.ends_with(".svg") {
        Some(ImageFormat::Svg)
    } else {
        None
    }
}

fn is_image_relationship(rel_type: Option<&str>, target: &str) -> bool {
    image_format_from_ext(target).is_some()
        || rel_type.is_some_and(|rel_type| {
            let lower = rel_type.to_ascii_lowercase();
            lower.contains("/image") || lower.contains("hdphoto")
        })
}

/// Parse presentation.xml to extract slide size and ordered slide relationship IDs.
fn parse_presentation_xml(xml: &str) -> Result<(PageSize, Vec<String>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    // Default slide dimensions: 10" x 7.5" (standard 4:3)
    let mut slide_size = PageSize {
        width: 720.0,
        height: 540.0,
    };
    let mut slide_rids = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) => {
                handle_presentation_element(e, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Start(ref e)) => {
                handle_presentation_element(e, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ConvertError::Parse(format!(
                    "XML error in presentation.xml: {e}"
                )));
            }
            _ => {}
        }
    }

    Ok((slide_size, slide_rids))
}

/// Handle a single element from presentation.xml (both Start and Empty events).
fn handle_presentation_element(
    e: &quick_xml::events::BytesStart,
    slide_size: &mut PageSize,
    slide_rids: &mut Vec<String>,
) {
    match e.local_name().as_ref() {
        b"sldSz" => {
            let cx = get_attr_i64(e, b"cx").unwrap_or(9_144_000);
            let cy = get_attr_i64(e, b"cy").unwrap_or(6_858_000);
            *slide_size = PageSize {
                width: emu_to_pt(cx),
                height: emu_to_pt(cy),
            };
        }
        b"sldId" => {
            // r:id attribute contains the relationship ID
            if let Some(rid) = get_attr_str(e, b"r:id") {
                slide_rids.push(rid);
            }
        }
        _ => {}
    }
}

/// Parse a .rels file to build Id → full relationship metadata mapping.
fn parse_relationships_xml(xml: &str) -> HashMap<String, Relationship> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                {
                    map.insert(
                        id,
                        Relationship {
                            target,
                            rel_type: get_attr_str(e, b"Type"),
                        },
                    );
                }
            }
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                {
                    map.insert(
                        id,
                        Relationship {
                            target,
                            rel_type: get_attr_str(e, b"Type"),
                        },
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Parse a .rels file to build Id → Target mapping.
fn parse_rels_xml(xml: &str) -> HashMap<String, String> {
    parse_relationships_xml(xml)
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect()
}

/// Find and load theme data from the PPTX archive.
///
/// Looks for a theme relationship in the presentation rels, reads the
/// theme XML, and parses the color scheme and font scheme.
fn load_theme<R: Read + std::io::Seek>(
    rel_map: &HashMap<String, String>,
    archive: &mut ZipArchive<R>,
) -> ThemeData {
    // Find theme target from rels (Type contains "theme")
    let theme_target = rel_map.values().find(|t| t.contains("theme"));
    let Some(target) = theme_target else {
        return ThemeData::default();
    };

    let theme_path = if let Some(stripped) = target.strip_prefix('/') {
        stripped.to_string()
    } else {
        format!("ppt/{target}")
    };

    let Ok(theme_xml) = read_zip_entry(archive, &theme_path) else {
        return ThemeData::default();
    };

    parse_theme_xml(&theme_xml)
}

/// Parse a theme XML string to extract the color scheme and font scheme.
fn parse_theme_xml(xml: &str) -> ThemeData {
    let mut theme = ThemeData::default();
    let mut reader = Reader::from_str(xml);

    // Color scheme element names in order
    const COLOR_NAMES: &[&str] = &[
        "dk1", "dk2", "lt1", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    let mut current_color_name: Option<String> = None;
    let mut in_major_font = false;
    let mut in_minor_font = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                // Check if this is a color scheme element
                if COLOR_NAMES.contains(&name) {
                    current_color_name = Some(name.to_string());
                }
                if name == "majorFont" {
                    in_major_font = true;
                }
                if name == "minorFont" {
                    in_minor_font = true;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                // Color scheme: <a:srgbClr val="RRGGBB"/> or <a:sysClr lastClr="RRGGBB"/>
                if let Some(ref cn) = current_color_name {
                    if name == "srgbClr"
                        && let Some(hex) = get_attr_str(e, b"val")
                        && let Some(color) = parse_hex_color(&hex)
                    {
                        theme.colors.insert(cn.clone(), color);
                    } else if name == "sysClr"
                        && let Some(hex) = get_attr_str(e, b"lastClr")
                        && let Some(color) = parse_hex_color(&hex)
                    {
                        theme.colors.insert(cn.clone(), color);
                    }
                }

                // Font scheme: <a:latin typeface="..."/>
                if name == "latin"
                    && let Some(typeface) = get_attr_str(e, b"typeface")
                {
                    if in_major_font {
                        theme.major_font = Some(typeface);
                    } else if in_minor_font {
                        theme.minor_font = Some(typeface);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                if current_color_name.as_deref() == Some(name) {
                    current_color_name = None;
                }
                if name == "majorFont" {
                    in_major_font = false;
                }
                if name == "minorFont" {
                    in_minor_font = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    theme
}

/// Parse background color from a `<p:bg>` element within a slide/layout/master XML.
///
/// Looks for `<p:bg><p:bgPr><a:solidFill>` and extracts the color
/// (either `<a:srgbClr>` or `<a:schemeClr>` resolved via theme).
fn parse_background_color(xml: &str, theme: &ThemeData, color_map: &ColorMapData) -> Option<Color> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_solid_fill = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => in_bg = true,
                    b"bgPr" if in_bg => in_bg_pr = true,
                    b"solidFill" if in_bg_pr => in_solid_fill = true,
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_solid_fill => {
                        return parse_color_from_start(&mut reader, e, theme, color_map).color;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_solid_fill => {
                        return parse_color_from_empty(e, theme, color_map).color;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => return None, // Found bg but no solid fill color
                    b"bgPr" => in_bg_pr = false,
                    b"solidFill" => in_solid_fill = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    None
}

/// Parse gradient fill from a `<p:bg>` element within a slide/layout/master XML.
///
/// Looks for `<p:bg><p:bgPr><a:gradFill>` and extracts gradient stops and angle.
/// Returns `None` if no gradient background is found.
fn parse_background_gradient(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<GradientFill> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_grad_fill = false;
    let mut in_gs_lst = false;
    let mut in_gs = false;
    let mut current_pos: f64 = 0.0;

    let mut stops: Vec<GradientStop> = Vec::new();
    let mut angle: f64 = 0.0;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => in_bg = true,
                    b"bgPr" if in_bg => in_bg_pr = true,
                    b"gradFill" if in_bg_pr => in_grad_fill = true,
                    b"gsLst" if in_grad_fill => in_gs_lst = true,
                    b"gs" if in_gs_lst => {
                        in_gs = true;
                        current_pos = get_attr_i64(e, b"pos").unwrap_or(0) as f64 / 100_000.0;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_gs => {
                        if let Some(color) =
                            parse_color_from_start(&mut reader, e, theme, color_map).color
                        {
                            stops.push(GradientStop {
                                position: current_pos,
                                color,
                            });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_gs => {
                        if let Some(color) = parse_color_from_empty(e, theme, color_map).color {
                            stops.push(GradientStop {
                                position: current_pos,
                                color,
                            });
                        }
                    }
                    b"lin" if in_grad_fill => {
                        // ang is in 60000ths of a degree
                        if let Some(ang) = get_attr_i64(e, b"ang") {
                            angle = ang as f64 / 60_000.0;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => {
                        if !stops.is_empty() {
                            return Some(GradientFill { stops, angle });
                        }
                        return None;
                    }
                    b"bgPr" => in_bg_pr = false,
                    b"gradFill" => in_grad_fill = false,
                    b"gsLst" => in_gs_lst = false,
                    b"gs" => in_gs = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    None
}

/// Parse gradient fill from shape properties XML.
///
/// Looks for `<a:gradFill>` within shape properties and extracts gradient stops and angle.
fn parse_shape_gradient_fill(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<GradientFill> {
    let mut in_gs_lst = false;
    let mut in_gs = false;
    let mut current_pos: f64 = 0.0;
    let mut stops: Vec<GradientStop> = Vec::new();
    let mut angle: f64 = 0.0;
    let mut depth: usize = 1; // we're already inside <a:gradFill>

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                let local = e.local_name();
                match local.as_ref() {
                    b"gsLst" => in_gs_lst = true,
                    b"gs" if in_gs_lst => {
                        in_gs = true;
                        current_pos = get_attr_i64(e, b"pos").unwrap_or(0) as f64 / 100_000.0;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_gs => {
                        if let Some(color) =
                            parse_color_from_start(reader, e, theme, color_map).color
                        {
                            stops.push(GradientStop {
                                position: current_pos,
                                color,
                            });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_gs => {
                        if let Some(color) = parse_color_from_empty(e, theme, color_map).color {
                            stops.push(GradientStop {
                                position: current_pos,
                                color,
                            });
                        }
                    }
                    b"lin" => {
                        if let Some(ang) = get_attr_i64(e, b"ang") {
                            angle = ang as f64 / 60_000.0;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                depth -= 1;
                if depth == 0 {
                    // End of gradFill
                    if stops.is_empty() {
                        return None;
                    }
                    return Some(GradientFill { stops, angle });
                }
                let local = e.local_name();
                match local.as_ref() {
                    b"gsLst" => in_gs_lst = false,
                    b"gs" => in_gs = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    None
}

/// Parse `<a:effectLst>` and extract outer shadow if present.
///
/// The reader should be positioned right after the `<a:effectLst>` Start event.
/// Reads until the matching `</a:effectLst>` End event.
fn parse_effect_list(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<Shadow> {
    let mut shadow: Option<Shadow> = None;
    let mut in_outer_shdw = false;
    let mut shdw_blur: f64 = 0.0;
    let mut shdw_dist: f64 = 0.0;
    let mut shdw_dir: f64 = 0.0;
    let mut shdw_color: Option<Color> = None;
    let mut shdw_opacity: f64 = 1.0;
    let mut depth: usize = 1; // already inside <a:effectLst>

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                let local = e.local_name();
                match local.as_ref() {
                    b"outerShdw" => {
                        in_outer_shdw = true;
                        // EMU values: 1 pt = 12700 EMU
                        shdw_blur = get_attr_i64(e, b"blurRad").unwrap_or(0) as f64 / 12_700.0;
                        shdw_dist = get_attr_i64(e, b"dist").unwrap_or(0) as f64 / 12_700.0;
                        // Direction in 60000ths of a degree
                        shdw_dir = get_attr_i64(e, b"dir").unwrap_or(0) as f64 / 60_000.0;
                        shdw_color = None;
                        shdw_opacity = 1.0;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_outer_shdw => {
                        let parsed = parse_color_from_start(reader, e, theme, color_map);
                        shdw_color = parsed.color;
                        if let Some(alpha) = parsed.alpha {
                            shdw_opacity = alpha;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"outerShdw" => {
                        // Self-closing <a:outerShdw/> with attributes
                        let blur = get_attr_i64(e, b"blurRad").unwrap_or(0) as f64 / 12_700.0;
                        let dist = get_attr_i64(e, b"dist").unwrap_or(0) as f64 / 12_700.0;
                        let dir = get_attr_i64(e, b"dir").unwrap_or(0) as f64 / 60_000.0;
                        shadow = Some(Shadow {
                            blur_radius: blur,
                            distance: dist,
                            direction: dir,
                            color: Color::new(0, 0, 0),
                            opacity: 1.0,
                        });
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_outer_shdw => {
                        let parsed = parse_color_from_empty(e, theme, color_map);
                        shdw_color = parsed.color;
                        if let Some(alpha) = parsed.alpha {
                            shdw_opacity = alpha;
                        }
                    }
                    b"alpha" if in_outer_shdw => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            shdw_opacity = val as f64 / 100_000.0;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                depth -= 1;
                if depth == 0 {
                    break; // End of effectLst
                }
                let local = e.local_name();
                if local.as_ref() == b"outerShdw" && in_outer_shdw {
                    in_outer_shdw = false;
                    if let Some(color) = shdw_color {
                        shadow = Some(Shadow {
                            blur_radius: shdw_blur,
                            distance: shdw_dist,
                            direction: shdw_dir,
                            color,
                            opacity: shdw_opacity,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    shadow
}

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

/// Group shape coordinate transform.
///
/// Maps child coordinates from the group's internal coordinate space
/// to the parent (slide or outer group) coordinate space.
#[derive(Debug, Default)]
struct GroupTransform {
    /// Group position on parent, in EMU.
    off_x: i64,
    off_y: i64,
    /// Group extent (size) on parent, in EMU.
    ext_cx: i64,
    ext_cy: i64,
    /// Child coordinate space origin, in EMU.
    ch_off_x: i64,
    ch_off_y: i64,
    /// Child coordinate space extent, in EMU.
    ch_ext_cx: i64,
    ch_ext_cy: i64,
}

impl GroupTransform {
    /// Apply the transform to a `FixedElement` whose coordinates are already in points.
    fn apply(&self, elem: &mut FixedElement) {
        let scale_x = if self.ch_ext_cx != 0 {
            self.ext_cx as f64 / self.ch_ext_cx as f64
        } else {
            1.0
        };
        let scale_y = if self.ch_ext_cy != 0 {
            self.ext_cy as f64 / self.ch_ext_cy as f64
        } else {
            1.0
        };

        let off_x_pt = emu_to_pt(self.off_x);
        let off_y_pt = emu_to_pt(self.off_y);
        let ch_off_x_pt = emu_to_pt(self.ch_off_x);
        let ch_off_y_pt = emu_to_pt(self.ch_off_y);

        elem.x = off_x_pt + (elem.x - ch_off_x_pt) * scale_x;
        elem.y = off_y_pt + (elem.y - ch_off_y_pt) * scale_y;
        elem.width *= scale_x;
        elem.height *= scale_y;
    }
}

/// Parse a `<p:grpSp>` group shape from the reader.
///
/// Called right after the `<p:grpSp>` start tag has been consumed.
/// Reads through the group's header sections (`nvGrpSpPr`, `grpSpPr`),
/// extracts the coordinate transform, then slices the original XML to
/// get the child shapes, and recursively parses them via `parse_slide_xml`.
fn parse_group_shape(
    reader: &mut Reader<&[u8]>,
    xml: &str,
    images: &SlideImageMap,
    theme: &ThemeData,
    color_map: &ColorMapData,
    warning_context: &str,
    inherited_text_body_defaults: &PptxTextBodyStyleDefaults,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    let mut transform = GroupTransform::default();
    let mut in_xfrm = false;
    let mut header_depth: usize = 0;
    let mut children_start = reader.buffer_position() as usize;

    // Phase 1: Read nvGrpSpPr and grpSpPr sections, extracting the transform.
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"nvGrpSpPr" if header_depth == 0 => header_depth = 1,
                b"grpSpPr" if header_depth == 0 => header_depth = 1,
                b"xfrm" if header_depth == 1 => in_xfrm = true,
                _ if header_depth > 0 => header_depth += 1,
                _ => break, // unexpected top-level child — treat as children start
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"grpSpPr" if header_depth == 0 => {
                    children_start = reader.buffer_position() as usize;
                    break;
                }
                b"off" if in_xfrm => {
                    transform.off_x = get_attr_i64(e, b"x").unwrap_or(0);
                    transform.off_y = get_attr_i64(e, b"y").unwrap_or(0);
                }
                b"ext" if in_xfrm => {
                    transform.ext_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                    transform.ext_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                }
                b"chOff" if in_xfrm => {
                    transform.ch_off_x = get_attr_i64(e, b"x").unwrap_or(0);
                    transform.ch_off_y = get_attr_i64(e, b"y").unwrap_or(0);
                }
                b"chExt" if in_xfrm => {
                    transform.ch_ext_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                    transform.ch_ext_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"xfrm" if in_xfrm => in_xfrm = false,
                b"grpSpPr" if header_depth == 1 => {
                    children_start = reader.buffer_position() as usize;
                    break;
                }
                b"nvGrpSpPr" if header_depth == 1 => header_depth = 0,
                _ if header_depth > 1 => header_depth -= 1,
                b"grpSp" => return Ok((Vec::new(), Vec::new())), // empty group
                _ => {}
            },
            Ok(Event::Eof) => return Ok((Vec::new(), Vec::new())),
            Err(e) => {
                return Err(ConvertError::Parse(format!(
                    "XML error in group shape: {e}"
                )));
            }
            _ => {}
        }
    }

    // Phase 2: Skip to </p:grpSp>, recording where the children end.
    let mut grp_depth: usize = 1;
    loop {
        let pos = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                grp_depth += 1;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                grp_depth -= 1;
                if grp_depth == 0 {
                    let children_xml = &xml[children_start..pos];
                    if children_xml.trim().is_empty() {
                        return Ok((Vec::new(), Vec::new()));
                    }

                    let wrapped = format!(
                        r#"<r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{children_xml}</r>"#
                    );

                    let (mut child_elements, warnings) = parse_slide_xml(
                        &wrapped,
                        images,
                        theme,
                        color_map,
                        warning_context,
                        inherited_text_body_defaults,
                    )?;
                    for elem in &mut child_elements {
                        transform.apply(elem);
                    }
                    return Ok((child_elements, warnings));
                }
            }
            Ok(Event::Eof) => return Ok((Vec::new(), Vec::new())),
            Err(e) => {
                return Err(ConvertError::Parse(format!(
                    "XML error in group shape: {e}"
                )));
            }
            _ => {}
        }
    }
}

fn parse_crop_fraction(e: &quick_xml::events::BytesStart, key: &[u8]) -> f64 {
    get_attr_i64(e, key)
        .map(|value| (value as f64 / 100_000.0).clamp(0.0, 1.0))
        .unwrap_or(0.0)
}

fn parse_src_rect(e: &quick_xml::events::BytesStart) -> Option<ImageCrop> {
    let crop = ImageCrop {
        left: parse_crop_fraction(e, b"l"),
        top: parse_crop_fraction(e, b"t"),
        right: parse_crop_fraction(e, b"r"),
        bottom: parse_crop_fraction(e, b"b"),
    };
    (!crop.is_empty()).then_some(crop)
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

/// Resolve a font typeface, substituting theme font references.
///
/// `+mj-lt` → major latin font from theme.
/// `+mn-lt` → minor latin font from theme.
/// Everything else is returned as-is.
fn resolve_theme_font(typeface: &str, theme: &ThemeData) -> String {
    match typeface {
        "+mj-lt" => theme
            .major_font
            .clone()
            .unwrap_or_else(|| typeface.to_string()),
        "+mn-lt" => theme
            .minor_font
            .clone()
            .unwrap_or_else(|| typeface.to_string()),
        other => other.to_string(),
    }
}

/// Map a PPTX preset geometry name to an IR ShapeKind.
fn prst_to_shape_kind(prst: &str, width: f64, height: f64) -> ShapeKind {
    match prst {
        "ellipse" => ShapeKind::Ellipse,
        "line" | "straightConnector1" => ShapeKind::Line {
            x2: width,
            y2: height,
        },
        "roundRect" => ShapeKind::RoundedRectangle {
            radius_fraction: 0.1,
        },
        "triangle" => ShapeKind::Polygon {
            vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
        },
        "rtTriangle" => ShapeKind::Polygon {
            vertices: vec![(0.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
        },
        "diamond" => ShapeKind::Polygon {
            vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
        },
        "pentagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(5),
        },
        "hexagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(6),
        },
        "octagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(8),
        },
        "rightArrow" | "arrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Right),
        },
        "leftArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Left),
        },
        "upArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Up),
        },
        "downArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Down),
        },
        "star4" => ShapeKind::Polygon {
            vertices: star_vertices(4),
        },
        "star5" => ShapeKind::Polygon {
            vertices: star_vertices(5),
        },
        "star6" => ShapeKind::Polygon {
            vertices: star_vertices(6),
        },
        // Rectangular fallback for unsupported presets
        _ => ShapeKind::Rectangle,
    }
}

enum ArrowDir {
    Right,
    Left,
    Up,
    Down,
}

/// Generate vertices for a regular polygon inscribed in the unit square (0–1).
fn regular_polygon_vertices(n: usize) -> Vec<(f64, f64)> {
    let mut vertices = Vec::with_capacity(n);
    for i in 0..n {
        // Start from top (−π/2) and go clockwise
        let angle = -std::f64::consts::FRAC_PI_2 + 2.0 * std::f64::consts::PI * i as f64 / n as f64;
        let x = 0.5 + 0.5 * angle.cos();
        let y = 0.5 + 0.5 * angle.sin();
        vertices.push((x, y));
    }
    vertices
}

/// Generate arrow polygon vertices (7-point arrow) in normalized coordinates.
fn arrow_vertices(dir: ArrowDir) -> Vec<(f64, f64)> {
    // Right-pointing arrow template
    let right: Vec<(f64, f64)> = vec![
        (0.0, 0.25),
        (0.6, 0.25),
        (0.6, 0.0),
        (1.0, 0.5),
        (0.6, 1.0),
        (0.6, 0.75),
        (0.0, 0.75),
    ];
    match dir {
        ArrowDir::Right => right,
        ArrowDir::Left => right.into_iter().map(|(x, y)| (1.0 - x, y)).collect(),
        ArrowDir::Up => right.into_iter().map(|(x, y)| (y, 1.0 - x)).collect(),
        ArrowDir::Down => right.into_iter().map(|(x, y)| (1.0 - y, x)).collect(),
    }
}

/// Generate star polygon vertices with `n` points inscribed in the unit square.
fn star_vertices(n: usize) -> Vec<(f64, f64)> {
    let mut vertices = Vec::with_capacity(n * 2);
    let inner_radius = 0.4; // ratio of inner to outer radius
    for i in 0..(n * 2) {
        let angle = -std::f64::consts::FRAC_PI_2 + std::f64::consts::PI * i as f64 / n as f64;
        let r = if i % 2 == 0 { 0.5 } else { 0.5 * inner_radius };
        let x = 0.5 + r * angle.cos();
        let y = 0.5 + r * angle.sin();
        vertices.push((x, y));
    }
    vertices
}

fn merge_paragraph_style(target: &mut ParagraphStyle, source: &ParagraphStyle) {
    if source.alignment.is_some() {
        target.alignment = source.alignment;
    }
    if source.indent_left.is_some() {
        target.indent_left = source.indent_left;
    }
    if source.indent_right.is_some() {
        target.indent_right = source.indent_right;
    }
    if source.indent_first_line.is_some() {
        target.indent_first_line = source.indent_first_line;
    }
    if source.line_spacing.is_some() {
        target.line_spacing = source.line_spacing;
    }
    if source.space_before.is_some() {
        target.space_before = source.space_before;
    }
    if source.space_after.is_some() {
        target.space_after = source.space_after;
    }
    if source.heading_level.is_some() {
        target.heading_level = source.heading_level;
    }
    if source.direction.is_some() {
        target.direction = source.direction;
    }
    if source.tab_stops.is_some() {
        target.tab_stops = source.tab_stops.clone();
    }
}

fn merge_text_style(target: &mut TextStyle, source: &TextStyle) {
    if source.font_family.is_some() {
        target.font_family = source.font_family.clone();
    }
    if source.font_size.is_some() {
        target.font_size = source.font_size;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline = source.underline;
    }
    if source.strikethrough.is_some() {
        target.strikethrough = source.strikethrough;
    }
    if source.color.is_some() {
        target.color = source.color;
    }
    if source.highlight.is_some() {
        target.highlight = source.highlight;
    }
    if source.vertical_align.is_some() {
        target.vertical_align = source.vertical_align;
    }
    if source.all_caps.is_some() {
        target.all_caps = source.all_caps;
    }
    if source.small_caps.is_some() {
        target.small_caps = source.small_caps;
    }
    if source.letter_spacing.is_some() {
        target.letter_spacing = source.letter_spacing;
    }
}

fn merge_pptx_bullet_definition(target: &mut PptxBulletDefinition, source: &PptxBulletDefinition) {
    if source.kind.is_some() {
        target.kind = source.kind.clone();
    }
    if source.font.is_some() {
        target.font = source.font.clone();
    }
    if source.color.is_some() {
        target.color = source.color.clone();
    }
    if source.size.is_some() {
        target.size = source.size.clone();
    }
}

fn parse_pptx_list_style_level(name: &[u8]) -> Option<u32> {
    if name.len() != 7 || !name.starts_with(b"lvl") || !name.ends_with(b"pPr") {
        return None;
    }
    let digit = name[3];
    if !(b'1'..=b'9').contains(&digit) {
        return None;
    }
    Some(u32::from(digit - b'1'))
}

fn apply_typeface_to_style(
    element: &quick_xml::events::BytesStart,
    style: &mut TextStyle,
    theme: &ThemeData,
) {
    let Some(typeface) = get_attr_str(element, b"typeface") else {
        return;
    };
    if typeface.trim().is_empty() || style.font_family.is_some() {
        return;
    }
    style.font_family = Some(resolve_theme_font(&typeface, theme));
}

fn parse_pptx_list_style(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PptxTextBodyStyleDefaults {
    #[derive(Clone, Copy)]
    enum ParagraphTarget {
        Default,
        Level(u32),
    }

    let mut defaults = PptxTextBodyStyleDefaults::default();
    let mut active_paragraph_target: Option<ParagraphTarget> = None;
    let mut active_run_target: Option<ParagraphTarget> = None;
    let mut in_ln_spc = false;
    let mut in_run_fill = false;
    let mut in_bullet_fill = false;

    fn paragraph_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut ParagraphStyle {
        match target {
            ParagraphTarget::Default => &mut defaults.default_paragraph,
            ParagraphTarget::Level(level) => {
                &mut defaults.levels.entry(level).or_default().paragraph
            }
        }
    }

    fn run_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut TextStyle {
        match target {
            ParagraphTarget::Default => &mut defaults.default_run,
            ParagraphTarget::Level(level) => &mut defaults.levels.entry(level).or_default().run,
        }
    }

    fn bullet_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut PptxBulletDefinition {
        match target {
            ParagraphTarget::Default => &mut defaults.default_bullet,
            ParagraphTarget::Level(level) => &mut defaults.levels.entry(level).or_default().bullet,
        }
    }

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"defPPr" => {
                        active_paragraph_target = Some(ParagraphTarget::Default);
                        extract_paragraph_props(
                            e,
                            paragraph_style_mut(&mut defaults, ParagraphTarget::Default),
                        );
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level = parse_pptx_list_style_level(name).unwrap();
                        active_paragraph_target = Some(ParagraphTarget::Level(level));
                        extract_paragraph_props(
                            e,
                            paragraph_style_mut(&mut defaults, ParagraphTarget::Level(level)),
                        );
                    }
                    b"lnSpc" if active_paragraph_target.is_some() => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pct(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"spcPts" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pts(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"buAutoNum" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind = Some(
                                PptxBulletKind::AutoNumber(parse_pptx_auto_numbering(e, level)),
                            );
                        }
                    }
                    b"buChar" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind =
                                parse_pptx_bullet_marker(e, level);
                        }
                    }
                    b"buNone" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).kind =
                                Some(PptxBulletKind::None);
                        }
                    }
                    b"buFontTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::FollowText);
                        }
                    }
                    b"buFont" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(typeface) = get_attr_str(e, b"typeface")
                        {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::Explicit(resolve_theme_font(
                                    &typeface, theme,
                                )));
                        }
                    }
                    b"buClrTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).color =
                                Some(PptxBulletColorSource::FollowText);
                        }
                    }
                    b"buClr" if active_paragraph_target.is_some() => {
                        in_bullet_fill = true;
                    }
                    b"buSzTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::FollowText);
                        }
                    }
                    b"buSzPct" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"defRPr" if active_paragraph_target.is_some() => {
                        active_run_target = active_paragraph_target;
                        if let Some(target) = active_run_target {
                            extract_rpr_attributes(e, run_style_mut(&mut defaults, target));
                        }
                    }
                    b"solidFill" if active_run_target.is_some() => {
                        in_run_fill = true;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_run_fill => {
                        let parsed = parse_color_from_start(reader, e, theme, color_map);
                        if let Some(target) = active_run_target {
                            run_style_mut(&mut defaults, target).color = parsed.color;
                        }
                    }
                    b"latin" | b"ea" | b"cs" if active_run_target.is_some() => {
                        if let Some(target) = active_run_target {
                            apply_typeface_to_style(e, run_style_mut(&mut defaults, target), theme);
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_bullet_fill => {
                        if let Some(target) = active_paragraph_target {
                            let parsed = parse_color_from_start(reader, e, theme, color_map);
                            bullet_style_mut(&mut defaults, target).color =
                                parsed.color.map(PptxBulletColorSource::Explicit);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"defPPr" => {
                        extract_paragraph_props(e, &mut defaults.default_paragraph);
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level = parse_pptx_list_style_level(name).unwrap();
                        extract_paragraph_props(
                            e,
                            &mut defaults.levels.entry(level).or_default().paragraph,
                        );
                    }
                    b"spcPct" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pct(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"spcPts" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pts(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"buAutoNum" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind = Some(
                                PptxBulletKind::AutoNumber(parse_pptx_auto_numbering(e, level)),
                            );
                        }
                    }
                    b"buChar" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind =
                                parse_pptx_bullet_marker(e, level);
                        }
                    }
                    b"buNone" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).kind =
                                Some(PptxBulletKind::None);
                        }
                    }
                    b"buFontTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::FollowText);
                        }
                    }
                    b"buFont" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(typeface) = get_attr_str(e, b"typeface")
                        {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::Explicit(resolve_theme_font(
                                    &typeface, theme,
                                )));
                        }
                    }
                    b"buClrTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).color =
                                Some(PptxBulletColorSource::FollowText);
                        }
                    }
                    b"buClr" if active_paragraph_target.is_some() => {}
                    b"buSzTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::FollowText);
                        }
                    }
                    b"buSzPct" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"defRPr" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            extract_rpr_attributes(e, run_style_mut(&mut defaults, target));
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_run_fill => {
                        let parsed = parse_color_from_empty(e, theme, color_map);
                        if let Some(target) = active_run_target {
                            run_style_mut(&mut defaults, target).color = parsed.color;
                        }
                    }
                    b"latin" | b"ea" | b"cs" if active_run_target.is_some() => {
                        if let Some(target) = active_run_target {
                            apply_typeface_to_style(e, run_style_mut(&mut defaults, target), theme);
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_bullet_fill => {
                        if let Some(target) = active_paragraph_target {
                            let parsed = parse_color_from_empty(e, theme, color_map);
                            bullet_style_mut(&mut defaults, target).color =
                                parsed.color.map(PptxBulletColorSource::Explicit);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"lstStyle" | b"otherStyle" => break,
                    b"defPPr" => {
                        active_paragraph_target = None;
                        in_ln_spc = false;
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        active_paragraph_target = None;
                        in_ln_spc = false;
                    }
                    b"defRPr" => {
                        active_run_target = None;
                        in_run_fill = false;
                    }
                    b"solidFill" if in_run_fill => {
                        in_run_fill = false;
                    }
                    b"buClr" if in_bullet_fill => {
                        in_bullet_fill = false;
                    }
                    b"lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    defaults
}

/// Extract paragraph alignment and direction from `<a:pPr>` attributes.
fn extract_paragraph_props(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    if let Some(algn) = get_attr_str(e, b"algn") {
        style.alignment = match algn.as_str() {
            "l" => Some(Alignment::Left),
            "ctr" => Some(Alignment::Center),
            "r" => Some(Alignment::Right),
            "just" => Some(Alignment::Justify),
            _ => None,
        };
    }
    if let Some(val) = get_attr_str(e, b"rtl")
        && (val == "1" || val == "true")
    {
        style.direction = Some(TextDirection::Rtl);
    }
    if let Some(value) = get_attr_i64(e, b"marL") {
        style.indent_left = Some(emu_to_pt(value));
    }
    if let Some(value) = get_attr_i64(e, b"marR") {
        style.indent_right = Some(emu_to_pt(value));
    }
    if let Some(value) = get_attr_i64(e, b"indent") {
        style.indent_first_line = Some(emu_to_pt(value));
    }
}

fn extract_pptx_line_spacing_pct(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    if let Some(value) = get_attr_i64(e, b"val") {
        style.line_spacing = Some(LineSpacing::Proportional(value as f64 / 100_000.0));
    }
}

fn extract_pptx_line_spacing_pts(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    if let Some(value) = get_attr_i64(e, b"val") {
        style.line_spacing = Some(LineSpacing::Exact(value as f64 / 100.0));
    }
}

fn extract_pptx_text_box_body_props(
    e: &quick_xml::events::BytesStart,
    padding: &mut Insets,
    vertical_align: &mut TextBoxVerticalAlign,
) {
    if let Some(value) = get_attr_i64(e, b"lIns") {
        padding.left = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"rIns") {
        padding.right = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"tIns") {
        padding.top = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"bIns") {
        padding.bottom = emu_to_pt(value);
    }
    if let Some(anchor) = get_attr_str(e, b"anchor") {
        *vertical_align = match anchor.as_str() {
            "ctr" => TextBoxVerticalAlign::Center,
            "b" => TextBoxVerticalAlign::Bottom,
            _ => TextBoxVerticalAlign::Top,
        };
    }
}

fn extract_pptx_table_cell_props(
    e: &quick_xml::events::BytesStart,
    vertical_align: &mut Option<CellVerticalAlign>,
    padding: &mut Option<Insets>,
) {
    if let Some(anchor) = get_attr_str(e, b"anchor") {
        *vertical_align = Some(match anchor.as_str() {
            "ctr" => CellVerticalAlign::Center,
            "b" => CellVerticalAlign::Bottom,
            _ => CellVerticalAlign::Top,
        });
    }

    let mut cell_padding = (*padding).unwrap_or_default();
    let mut has_padding = false;
    if let Some(value) = get_attr_i64(e, b"marL") {
        cell_padding.left = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marR") {
        cell_padding.right = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marT") {
        cell_padding.top = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marB") {
        cell_padding.bottom = emu_to_pt(value);
        has_padding = true;
    }
    if has_padding {
        *padding = Some(cell_padding);
    }
}

fn push_pptx_run(runs: &mut Vec<Run>, run: Run) {
    if let Some(previous) = runs.last_mut()
        && previous.style == run.style
        && previous.href == run.href
        && previous.footnote == run.footnote
    {
        previous.text.push_str(&run.text);
        return;
    }

    let mut run = run;
    normalize_pptx_run_boundary_spacing(runs.last(), &mut run);
    runs.push(run);
}

fn push_pptx_soft_line_break(runs: &mut Vec<Run>, style: &TextStyle) {
    push_pptx_run(
        runs,
        Run {
            text: PPTX_SOFT_LINE_BREAK_CHAR.to_string(),
            style: style.clone(),
            href: None,
            footnote: None,
        },
    );
}

fn decode_pptx_text_event(text: &quick_xml::events::BytesText<'_>) -> Option<String> {
    let decoded = text.decode().ok()?;
    let unescaped = unescape_xml_text(decoded.as_ref()).ok()?;
    Some(unescaped.into_owned())
}

fn decode_pptx_general_ref(reference: &quick_xml::events::BytesRef<'_>) -> Option<String> {
    let decoded = reference.decode().ok()?;
    let wrapped = format!("&{};", decoded.as_ref());
    let unescaped = unescape_xml_text(&wrapped).ok()?;
    Some(unescaped.into_owned())
}

fn normalize_pptx_run_boundary_spacing(previous: Option<&Run>, run: &mut Run) {
    let Some(previous) = previous else {
        return;
    };

    if previous.href != run.href
        || previous.footnote.is_some()
        || run.footnote.is_some()
        || previous
            .text
            .chars()
            .last()
            .is_some_and(char::is_whitespace)
    {
        return;
    }

    let mut chars = run.text.chars();
    let Some(first_char) = chars.next() else {
        return;
    };
    let Some(next_char) = chars.next() else {
        return;
    };

    if first_char == ' ' && should_preserve_pptx_run_boundary_space(next_char) {
        // PowerPoint often splits styled phrases into adjacent runs such as
        // `K` + ` = 100)`. Preserve that boundary space as non-breaking so
        // Typst does not wrap at the style change and spill punctuation.
        run.text.replace_range(0..1, "\u{00A0}");
    }
}

fn should_preserve_pptx_run_boundary_space(next_char: char) -> bool {
    matches!(
        next_char,
        '=' | '+' | '-' | '/' | '%' | ')' | ']' | '}' | ':' | ';' | ',' | '.'
    )
}

fn first_pptx_visible_run_style(runs: &[Run]) -> Option<TextStyle> {
    runs.iter()
        .find(|run| !run.text.is_empty() && run.footnote.is_none())
        .map(|run| run.style.clone())
}

fn resolve_pptx_marker_base_style(
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> TextStyle {
    first_pptx_visible_run_style(runs)
        .or_else(|| {
            (end_para_run_style != &TextStyle::default()).then(|| end_para_run_style.clone())
        })
        .unwrap_or_else(|| default_run_style.clone())
}

fn finalize_pptx_marker_style(style: TextStyle) -> Option<TextStyle> {
    (style != TextStyle::default()).then_some(style)
}

fn resolve_pptx_marker_style(
    bullet: &PptxBulletDefinition,
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<TextStyle> {
    let mut style = resolve_pptx_marker_base_style(runs, end_para_run_style, default_run_style);

    match bullet.font.as_ref() {
        Some(PptxBulletFontSource::FollowText) | None => {}
        Some(PptxBulletFontSource::Explicit(font_family)) => {
            style.font_family = Some(font_family.clone());
        }
    }

    match bullet.color.as_ref() {
        Some(PptxBulletColorSource::FollowText) | None => {}
        Some(PptxBulletColorSource::Explicit(color)) => {
            style.color = Some(*color);
        }
    }

    match bullet.size.as_ref() {
        Some(PptxBulletSizeSource::FollowText) | None => {}
        Some(PptxBulletSizeSource::Points(points)) => {
            style.font_size = Some(*points);
        }
        Some(PptxBulletSizeSource::Percent(percent)) => {
            style.font_size = style.font_size.map(|size| size * percent);
        }
    }

    finalize_pptx_marker_style(style)
}

fn resolve_pptx_list_marker(
    bullet: &PptxBulletDefinition,
    level: u32,
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<PptxListMarker> {
    let marker_style =
        resolve_pptx_marker_style(bullet, runs, end_para_run_style, default_run_style);
    match bullet.kind.as_ref()? {
        PptxBulletKind::None => None,
        PptxBulletKind::Character(character) => Some(PptxListMarker::Unordered {
            level,
            marker_text: character.clone(),
            marker_style,
        }),
        PptxBulletKind::AutoNumber(auto_numbering) => Some(PptxListMarker::Ordered {
            auto_numbering: auto_numbering.clone(),
            marker_style,
        }),
    }
}

fn extract_paragraph_level(e: &quick_xml::events::BytesStart) -> u32 {
    get_attr_i64(e, b"lvl")
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0)
}

fn parse_pptx_auto_numbering(e: &quick_xml::events::BytesStart, level: u32) -> PptxAutoNumbering {
    let numbering_pattern: Option<String> = get_attr_str(e, b"type")
        .as_deref()
        .and_then(pptx_auto_numbering_pattern)
        .map(str::to_string);
    let start_at: Option<u32> = get_attr_i64(e, b"startAt").and_then(|value| value.try_into().ok());

    PptxAutoNumbering {
        level,
        numbering_pattern,
        start_at,
    }
}

fn parse_pptx_bullet_marker(
    e: &quick_xml::events::BytesStart,
    level: u32,
) -> Option<PptxBulletKind> {
    get_attr_str(e, b"char")
        .map(PptxBulletKind::Character)
        .or_else(|| (level == 0).then(|| PptxBulletKind::Character("•".to_string())))
}

fn pptx_auto_numbering_pattern(numbering_type: &str) -> Option<&'static str> {
    match numbering_type {
        "arabicPeriod" => Some("1."),
        "arabicParenR" => Some("1)"),
        "arabicParenBoth" => Some("(1)"),
        "alphaLcPeriod" => Some("a."),
        "alphaUcPeriod" => Some("A."),
        "alphaLcParenR" => Some("a)"),
        "alphaUcParenR" => Some("A)"),
        "romanLcPeriod" => Some("i."),
        "romanUcPeriod" => Some("I."),
        "romanLcParenR" => Some("i)"),
        "romanUcParenR" => Some("I)"),
        _ => None,
    }
}

fn group_pptx_text_blocks(entries: Vec<PptxParagraphEntry>) -> Vec<Block> {
    let mut entries = entries;
    trim_trailing_empty_pptx_list_entries(&mut entries);

    let mut blocks: Vec<Block> = Vec::new();
    let mut pending_list: Option<PendingPptxList> = None;

    for entry in entries {
        match entry.list_marker {
            Some(list_marker) => {
                if pending_list
                    .as_ref()
                    .is_some_and(|list| !list.can_extend(&list_marker))
                {
                    blocks.push(pending_list.take().unwrap().into_block());
                }

                let paragraph: Paragraph = entry.paragraph;
                pending_list
                    .get_or_insert_with(|| PendingPptxList::new(&list_marker))
                    .push(paragraph, list_marker);
            }
            None => {
                if let Some(list) = pending_list.take() {
                    blocks.push(list.into_block());
                }
                blocks.push(Block::Paragraph(entry.paragraph));
            }
        }
    }

    if let Some(list) = pending_list {
        blocks.push(list.into_block());
    }

    blocks
}

fn trim_trailing_empty_pptx_list_entries(entries: &mut Vec<PptxParagraphEntry>) {
    while entries.len() > 1 {
        let Some(last_entry) = entries.last() else {
            break;
        };
        if last_entry.list_marker.is_none()
            || pptx_paragraph_has_visible_content(&last_entry.paragraph)
        {
            break;
        }
        entries.pop();
    }
}

fn pptx_paragraph_has_visible_content(paragraph: &Paragraph) -> bool {
    paragraph.runs.iter().any(|run| {
        run.footnote.is_some()
            || run.text.chars().any(|character| {
                character != PPTX_SOFT_LINE_BREAK_CHAR && !character.is_whitespace()
            })
    })
}

/// Extract text formatting attributes from `<a:rPr>` element.
fn extract_rpr_attributes(e: &quick_xml::events::BytesStart, style: &mut TextStyle) {
    if let Some(val) = get_attr_str(e, b"b") {
        style.bold = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"i") {
        style.italic = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"u") {
        style.underline = Some(val != "none");
    }
    if let Some(val) = get_attr_str(e, b"strike") {
        style.strikethrough = Some(val != "noStrike");
    }
    if let Some(sz) = get_attr_i64(e, b"sz") {
        // Font size in hundredths of a point (e.g. 1200 = 12pt)
        style.font_size = Some(sz as f64 / 100.0);
    }
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
mod tests {
    use super::*;
    use crate::ir::ImageCrop;
    use std::io::Write;
    use zip::write::FileOptions;

    // ── Test helpers ─────────────────────────────────────────────────────

    /// Build a minimal PPTX file as bytes from slide XML strings.
    fn build_test_pptx(slide_cx_emu: i64, slide_cy_emu: i64, slide_xmls: &[String]) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        for i in 0..slide_xmls.len() {
            ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
        }
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/presentation.xml
        let mut pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
            slide_cx_emu, slide_cy_emu
        );
        for i in 0..slide_xmls.len() {
            pres.push_str(&format!(
                r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                256 + i,
                2 + i
            ));
        }
        pres.push_str("</p:sldIdLst></p:presentation>");
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        let mut pres_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for i in 0..slide_xmls.len() {
            pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
        }
        pres_rels.push_str("</Relationships>");
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // Slides
        for (i, slide_xml) in slide_xmls.iter().enumerate() {
            zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();
        }

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Create a slide XML with the given shape elements.
    fn make_slide_xml(shapes: &[String]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
        );
        for shape in shapes {
            xml.push_str(shape);
        }
        xml.push_str("</p:spTree></p:cSld></p:sld>");
        xml
    }

    /// Create an empty slide XML (no shapes).
    fn make_empty_slide_xml() -> String {
        make_slide_xml(&[])
    }

    /// Create a simple text box shape XML.
    fn make_text_box(x: i64, y: i64, cx: i64, cy: i64, text: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    }

    fn make_text_box_with_body_pr(
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
        body_pr_xml: &str,
        text: &str,
    ) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody>{body_pr_xml}<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    }

    /// Create a text box with formatted text runs.
    fn make_formatted_text_box(x: i64, y: i64, cx: i64, cy: i64, runs_xml: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>{runs_xml}</a:p></p:txBody></p:sp>"#
        )
    }

    /// Create a text box with multiple paragraphs.
    fn make_multi_para_text_box(x: i64, y: i64, cx: i64, cy: i64, paragraphs_xml: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>{paragraphs_xml}</p:txBody></p:sp>"#
        )
    }

    /// Create a slide XML with a background and optional shape elements.
    fn make_slide_xml_with_bg(bg_xml: &str, shapes: &[String]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>"#,
        );
        xml.push_str(bg_xml);
        xml.push_str(r#"<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#);
        for shape in shapes {
            xml.push_str(shape);
        }
        xml.push_str("</p:spTree></p:cSld></p:sld>");
        xml
    }

    /// Standard 4:3 slide size in EMU (10" x 7.5").
    const SLIDE_CX: i64 = 9_144_000;
    const SLIDE_CY: i64 = 6_858_000;

    /// Helper: get the first FixedPage from a Document.
    fn first_fixed_page(doc: &Document) -> &FixedPage {
        match &doc.pages[0] {
            Page::Fixed(p) => p,
            _ => panic!("Expected FixedPage"),
        }
    }

    fn text_box_data(elem: &FixedElement) -> &TextBoxData {
        match &elem.kind {
            FixedElementKind::TextBox(text_box) => text_box,
            _ => panic!("Expected TextBox"),
        }
    }

    /// Helper: get the TextBox blocks from a FixedElement.
    fn text_box_blocks(elem: &FixedElement) -> &[Block] {
        &text_box_data(elem).content
    }

    // ── Tests ────────────────────────────────────────────────────────────

    #[test]
    fn test_parse_empty_presentation() {
        // PPTX with zero slides → document with no pages
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        assert!(doc.pages.is_empty(), "Expected no pages");
    }

    #[test]
    fn test_parse_single_slide() {
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        assert_eq!(doc.pages.len(), 1, "Expected 1 page");
        assert!(matches!(&doc.pages[0], Page::Fixed(_)));
    }

    #[test]
    fn test_slide_dimensions() {
        // 16:9 widescreen: 12192000 × 6858000 EMU = 960pt × 540pt
        let cx = 12_192_000i64;
        let cy = 6_858_000i64;
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(cx, cy, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let expected_w = cx as f64 / 12700.0;
        let expected_h = cy as f64 / 12700.0;
        assert!(
            (page.size.width - expected_w).abs() < 0.1,
            "Expected width ~{expected_w}pt, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - expected_h).abs() < 0.1,
            "Expected height ~{expected_h}pt, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_text_box_extraction() {
        let shape = make_text_box(0, 0, 1_000_000, 500_000, "Hello World");
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 element");

        let blocks = text_box_blocks(&page.elements[0]);
        assert!(!blocks.is_empty(), "Expected at least one block");

        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "Hello World");
    }

    #[test]
    fn test_text_box_auto_numbered_paragraphs_group_into_list() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        assert_eq!(list.kind, crate::ir::ListKind::Ordered);
        assert_eq!(list.items.len(), 2);
        assert_eq!(
            list.level_styles
                .get(&0)
                .and_then(|style| style.numbering_pattern.as_deref()),
            Some("1.")
        );
        assert_eq!(list.items[0].content[0].runs[0].text, "First");
        assert_eq!(list.items[1].content[0].runs[0].text, "Second");
    }

    #[test]
    fn test_text_box_bulleted_paragraphs_group_into_list() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>First bullet</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:t>Second bullet</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        assert_eq!(blocks.len(), 1, "Expected a single grouped list block");

        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        assert_eq!(list.kind, crate::ir::ListKind::Unordered);
        assert_eq!(list.items.len(), 2);
        assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
        assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
    }

    #[test]
    fn test_text_box_bulleted_paragraph_preserves_char_marker_and_uses_run_style() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr indent="-216000"><a:buFontTx/><a:buChar char="-"/></a:pPr>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Pretendard"/></a:rPr><a:t>First bullet</a:t></a:r>"#,
            r#"</a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        let style = list.level_styles.get(&0).expect("Expected level 0 style");
        assert_eq!(style.marker_text.as_deref(), Some("-"));
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_family.as_deref()),
            Some("Pretendard")
        );
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_size),
            Some(14.0)
        );
        assert_eq!(
            style.marker_style.as_ref().and_then(|style| style.color),
            Some(Color::new(0x11, 0x22, 0x33))
        );
    }

    #[test]
    fn test_text_box_bulleted_paragraph_preserves_explicit_marker_font() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr indent="-216000"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Symbol bullet</a:t></a:r>"#,
            r#"</a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        let style = list.level_styles.get(&0).expect("Expected level 0 style");
        assert_eq!(style.marker_text.as_deref(), Some("è"));
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_family.as_deref()),
            Some("Wingdings")
        );
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_size),
            Some(14.0)
        );
    }

    #[test]
    fn test_text_box_paragraph_line_spacing_pct_extracted() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
            r#"<a:p><a:r><a:t>Second</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected Paragraph block, got {other:?}"),
        };
        match paragraph.style.line_spacing {
            Some(crate::ir::LineSpacing::Proportional(factor)) => {
                assert!((factor - 1.5).abs() < f64::EPSILON);
            }
            other => panic!("Expected proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_text_box_body_pr_defaults_and_center_anchor_extracted() {
        let shape = make_text_box_with_body_pr(
            0,
            0,
            1_000_000,
            500_000,
            r#"<a:bodyPr anchor="ctr"><a:spAutoFit/></a:bodyPr>"#,
            "Centered",
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let text_box = match &page.elements[0].kind {
            FixedElementKind::TextBox(text_box) => text_box,
            other => panic!("Expected TextBox, got {other:?}"),
        };
        assert!((text_box.padding.left - 7.2).abs() < 0.001);
        assert!((text_box.padding.right - 7.2).abs() < 0.001);
        assert!((text_box.padding.top - 3.6).abs() < 0.001);
        assert!((text_box.padding.bottom - 3.6).abs() < 0.001);
        assert_eq!(
            text_box.vertical_align,
            crate::ir::TextBoxVerticalAlign::Center
        );
    }

    #[test]
    fn test_text_box_auto_numbered_paragraph_start_override_sets_list_start() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod" startAt="3"/></a:pPr><a:r><a:t>Gamma</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr indent="-216000"><a:buAutoNum type="alphaUcPeriod"/></a:pPr><a:r><a:t>Delta</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        assert_eq!(list.kind, crate::ir::ListKind::Ordered);
        assert_eq!(list.items[0].start_at, Some(3));
        assert_eq!(
            list.level_styles
                .get(&0)
                .and_then(|style| style.numbering_pattern.as_deref()),
            Some("A.")
        );
    }

    #[test]
    fn test_text_box_auto_numbered_paragraph_extracts_hanging_indent() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr marL="457200" indent="-457200"><a:buAutoNum type="arabicParenR"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };

        let paragraph = &list.items[0].content[0];
        assert_eq!(paragraph.style.indent_left, Some(36.0));
        assert_eq!(paragraph.style.indent_first_line, Some(-36.0));
        assert_eq!(
            list.level_styles
                .get(&0)
                .and_then(|style| style.numbering_pattern.as_deref()),
            Some("1)")
        );
    }

    #[test]
    fn test_text_box_auto_numbered_paragraph_resolves_marker_style_from_text() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr marL="457200" indent="-457200">"#,
            r#"<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buAutoNum type="arabicParenR"/>"#,
            r#"</a:pPr>"#,
            r#"<a:r><a:rPr lang="ko-KR" sz="2000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Pretendard Medium"/></a:rPr><a:t>First</a:t></a:r>"#,
            r#"</a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        let style = list.level_styles.get(&0).expect("Expected level 0 style");
        assert_eq!(style.numbering_pattern.as_deref(), Some("1)"));
        assert_eq!(style.marker_text, None);
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_family.as_deref()),
            Some("Pretendard Medium")
        );
        assert_eq!(
            style
                .marker_style
                .as_ref()
                .and_then(|style| style.font_size),
            Some(20.0)
        );
        assert_eq!(
            style.marker_style.as_ref().and_then(|style| style.color),
            Some(Color::black())
        );
    }

    #[test]
    fn test_text_box_paragraph_preserves_soft_line_breaks() {
        let paragraphs_xml = concat!(
            r#"<a:p>"#,
            r#"<a:r><a:t>Line 1</a:t></a:r>"#,
            r#"<a:br/>"#,
            r#"<a:r><a:t>Line 2</a:t></a:r>"#,
            r#"</a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected Paragraph block, got {other:?}"),
        };
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "Line 1\u{000B}Line 2");
    }

    #[test]
    fn test_text_box_plain_paragraph_between_bullets_breaks_list_sequence() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>1) First bullet</a:t></a:r></a:p>"#,
            r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr><a:r><a:t>2) Second bullet</a:t></a:r></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        assert_eq!(blocks.len(), 3, "Expected list / paragraph / list split");
        match &blocks[0] {
            Block::List(list) => {
                assert_eq!(list.items.len(), 1);
                assert_eq!(
                    list.level_styles
                        .get(&1)
                        .and_then(|style| style.marker_text.as_deref()),
                    Some("-")
                );
            }
            other => panic!("Expected first block to be a list, got {other:?}"),
        }
        match &blocks[1] {
            Block::Paragraph(paragraph) => {
                let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
                assert_eq!(text, "-> Continuation paragraph");
            }
            other => panic!("Expected middle block to be a paragraph, got {other:?}"),
        }
        match &blocks[2] {
            Block::List(list) => {
                assert_eq!(list.items.len(), 1);
                assert_eq!(
                    list.level_styles
                        .get(&1)
                        .and_then(|style| style.marker_text.as_deref()),
                    Some("-")
                );
            }
            other => panic!("Expected last block to be a list, got {other:?}"),
        }
    }

    #[test]
    fn test_text_box_plain_paragraph_preserves_leading_arrow_text() {
        let paragraphs_xml = r#"<a:p><a:r><a:t>-> Continuation paragraph</a:t></a:r></a:p>"#;
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected paragraph block, got {other:?}"),
        };
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "-> Continuation paragraph");
    }

    #[test]
    fn test_text_box_plain_paragraph_preserves_escaped_gt_entity() {
        let paragraphs_xml = r#"<a:p><a:r><a:t>-&gt; Continuation paragraph</a:t></a:r></a:p>"#;
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected paragraph block, got {other:?}"),
        };
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "-> Continuation paragraph");
    }

    #[test]
    fn test_text_box_trailing_empty_bullets_do_not_override_nested_marker_style() {
        let paragraphs_xml = concat!(
            r#"<a:p><a:pPr marL="742950" lvl="1" indent="-285750"><a:buFont typeface="Wingdings"/><a:buChar char="è"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400"><a:latin typeface="Pretendard"/></a:rPr><a:t>Arrow bullet</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
            r#"<a:p><a:pPr marL="285750" indent="-285750"><a:buFontTx/><a:buChar char="-"/></a:pPr></a:p>"#,
        );
        let shape = make_multi_para_text_box(0, 0, 1_000_000, 500_000, paragraphs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let list = match &blocks[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        assert_eq!(list.items.len(), 1);
        assert_eq!(list.items[0].level, 1);
        assert_eq!(
            list.level_styles
                .get(&1)
                .and_then(|style| style.marker_text.as_deref()),
            Some("è")
        );
        assert!(
            list.level_styles.get(&0).is_none(),
            "Trailing empty dash bullets should not create a level-0 marker style"
        );
    }

    #[test]
    fn test_text_box_lst_style_default_run_props_are_applied_to_runs() {
        let shape = String::from(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="032543"/></a:solidFill><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p></p:txBody></p:sp>"#,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected Paragraph block, got {other:?}"),
        };
        let run = &paragraph.runs[0];
        assert_eq!(
            run.style.font_family.as_deref(),
            Some("Pretendard SemiBold")
        );
        assert_eq!(run.style.font_size, Some(14.0));
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.color, Some(Color::new(0x03, 0x25, 0x43)));
    }

    #[test]
    fn test_non_placeholder_shape_inherits_master_other_style_run_defaults() {
        let slide_shape = concat!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Caption"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
            r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#,
            r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>신</a:t></a:r><a:r><a:rPr lang="ko-KR" sz="1800"/><a:t>형</a:t></a:r></a:p>"#,
            r#"</p:txBody></p:sp>"#,
        );
        let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="224466"/></a:solidFill><a:latin typeface="Pretendard"/><a:ea typeface="Pretendard"/><a:cs typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
        let data = build_test_pptx_with_layout_master(
            SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected Paragraph block, got {other:?}"),
        };
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "신형");
        assert!(
            paragraph
                .runs
                .iter()
                .all(|run| run.style.font_size == Some(18.0))
        );
        assert!(
            paragraph
                .runs
                .iter()
                .all(|run| run.style.font_family.as_deref() == Some("Pretendard"))
        );
        assert!(
            paragraph
                .runs
                .iter()
                .all(|run| run.style.color == Some(Color::new(0x22, 0x44, 0x66)))
        );
    }

    #[test]
    fn test_text_box_lst_style_overrides_master_other_style_run_defaults() {
        let slide_shape = concat!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>"#,
            r#"<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="Pretendard SemiBold"/><a:ea typeface="Pretendard SemiBold"/><a:cs typeface="Pretendard SemiBold"/></a:defRPr></a:lvl1pPr></a:lstStyle>"#,
            r#"<a:p><a:r><a:rPr lang="ko-KR"/><a:t>경력</a:t></a:r></a:p>"#,
            r#"</p:txBody></p:sp>"#,
        );
        let slide_xml = make_slide_xml(&[slide_shape.to_string()]);
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#;
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:otherStyle><a:lvl1pPr marL="0"><a:defRPr sz="1800"><a:latin typeface="Pretendard"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
        let data = build_test_pptx_with_layout_master(
            SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paragraph = match &blocks[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected Paragraph block, got {other:?}"),
        };
        assert_eq!(paragraph.runs[0].style.font_size, Some(24.0));
        assert_eq!(
            paragraph.runs[0].style.font_family.as_deref(),
            Some("Pretendard SemiBold")
        );
    }

    #[test]
    fn test_text_box_position_and_size() {
        // Position: 1000000 EMU x, 500000 EMU y → ~78.74pt, ~39.37pt
        // Size: 5000000 EMU cx, 2000000 EMU cy → ~393.70pt, ~157.48pt
        let x = 1_000_000i64;
        let y = 500_000i64;
        let cx = 5_000_000i64;
        let cy = 2_000_000i64;
        let shape = make_text_box(x, y, cx, cy, "Positioned");
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let elem = &page.elements[0];

        let expected_x = x as f64 / 12700.0;
        let expected_y = y as f64 / 12700.0;
        let expected_w = cx as f64 / 12700.0;
        let expected_h = cy as f64 / 12700.0;

        assert!(
            (elem.x - expected_x).abs() < 0.1,
            "Expected x ~{expected_x}, got {}",
            elem.x
        );
        assert!(
            (elem.y - expected_y).abs() < 0.1,
            "Expected y ~{expected_y}, got {}",
            elem.y
        );
        assert!(
            (elem.width - expected_w).abs() < 0.1,
            "Expected width ~{expected_w}, got {}",
            elem.width
        );
        assert!(
            (elem.height - expected_h).abs() < 0.1,
            "Expected height ~{expected_h}, got {}",
            elem.height
        );
    }

    #[test]
    fn test_text_box_bold_formatting() {
        let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Bold text");
        assert_eq!(para.runs[0].style.bold, Some(true));
    }

    #[test]
    fn test_text_box_italic_formatting() {
        let runs_xml = r#"<a:r><a:rPr i="1"/><a:t>Italic text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Italic text");
        assert_eq!(para.runs[0].style.italic, Some(true));
    }

    #[test]
    fn test_text_box_font_size() {
        // sz="2400" means 24pt (hundredths of a point)
        let runs_xml = r#"<a:r><a:rPr sz="2400"/><a:t>Large text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].style.font_size, Some(24.0));
    }

    #[test]
    fn test_text_box_combined_formatting() {
        let runs_xml = r#"<a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        let run = &para.runs[0];
        assert_eq!(run.text, "Styled text");
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.italic, Some(true));
        assert_eq!(run.style.underline, Some(true));
        assert_eq!(run.style.strikethrough, Some(true));
        assert_eq!(run.style.font_size, Some(18.0));
        assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
        assert_eq!(run.style.font_family, Some("Arial".to_string()));
    }

    #[test]
    fn test_multiple_text_boxes() {
        let shape1 = make_text_box(100_000, 100_000, 2_000_000, 500_000, "Box 1");
        let shape2 = make_text_box(100_000, 700_000, 2_000_000, 500_000, "Box 2");
        let slide = make_slide_xml(&[shape1, shape2]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 text boxes");

        // Check content of each box
        let get_text = |elem: &FixedElement| -> String {
            let blocks = text_box_blocks(elem);
            blocks
                .iter()
                .filter_map(|b| match b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
                    _ => None,
                })
                .collect()
        };
        assert_eq!(get_text(&page.elements[0]), "Box 1");
        assert_eq!(get_text(&page.elements[1]), "Box 2");
    }

    #[test]
    fn test_multiple_slides() {
        let slide1 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 1")]);
        let slide2 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 2")]);
        let slide3 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 3")]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 3, "Expected 3 pages");
        for page in &doc.pages {
            assert!(matches!(page, Page::Fixed(_)));
        }
    }

    #[test]
    fn test_text_box_multiple_paragraphs() {
        let paras_xml = r#"<a:p><a:r><a:rPr/><a:t>Paragraph 1</a:t></a:r></a:p><a:p><a:r><a:rPr/><a:t>Paragraph 2</a:t></a:r></a:p>"#;
        let shape = make_multi_para_text_box(0, 0, 3_000_000, 2_000_000, paras_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let paras: Vec<&Paragraph> = blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .collect();
        assert!(paras.len() >= 2, "Expected at least 2 paragraphs");
        assert_eq!(paras[0].runs[0].text, "Paragraph 1");
        assert_eq!(paras[1].runs[0].text, "Paragraph 2");
    }

    #[test]
    fn test_text_box_multiple_runs() {
        let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold </a:t></a:r><a:r><a:rPr i="1"/><a:t>Italic</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 2);
        assert_eq!(para.runs[0].text, "Bold ");
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert_eq!(para.runs[1].text, "Italic");
        assert_eq!(para.runs[1].style.italic, Some(true));
    }

    #[test]
    fn test_paragraph_alignment_center() {
        let paras_xml = r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr/><a:t>Centered</a:t></a:r></a:p>"#;
        let shape = make_multi_para_text_box(0, 0, 2_000_000, 500_000, paras_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_parse_invalid_data() {
        let parser = PptxParser;
        let result = parser.parse(b"not a valid pptx file", &ConvertOptions::default());
        assert!(result.is_err());
        match result.unwrap_err() {
            ConvertError::Parse(_) => {}
            other => panic!("Expected Parse error, got: {other:?}"),
        }
    }

    #[test]
    fn test_slide_default_dimensions_4x3() {
        // Standard 4:3: 9144000 × 6858000 EMU = 720pt × 540pt
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert!(
            (page.size.width - 720.0).abs() < 0.1,
            "Expected width ~720pt, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - 540.0).abs() < 0.1,
            "Expected height ~540pt, got {}",
            page.size.height
        );
    }

    // ── Shape test helpers ───────────────────────────────────────────────

    /// Create a shape XML element with preset geometry, optional fill and border.
    #[allow(clippy::too_many_arguments)]
    fn make_shape(
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
        prst: &str,
        fill_hex: Option<&str>,
        border_width_emu: Option<i64>,
        border_hex: Option<&str>,
    ) -> String {
        let fill_xml = fill_hex
            .map(|h| format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#))
            .unwrap_or_default();

        let ln_xml = match (border_width_emu, border_hex) {
            (Some(w), Some(h)) => {
                format!(r#"<a:ln w="{w}"><a:solidFill><a:srgbClr val="{h}"/></a:solidFill></a:ln>"#)
            }
            _ => String::new(),
        };

        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</p:spPr></p:sp>"#
        )
    }

    /// Helper: extract the Shape from a FixedElement or panic.
    fn get_shape(elem: &FixedElement) -> &Shape {
        match &elem.kind {
            FixedElementKind::Shape(s) => s,
            other => panic!("Expected Shape, got {other:?}"),
        }
    }

    // ── Image test helpers ───────────────────────────────────────────────

    /// Create a minimal valid BMP (1×1 pixel, red) for test images.
    fn make_test_bmp() -> Vec<u8> {
        let mut bmp = Vec::new();
        // BMP header (14 bytes)
        bmp.extend_from_slice(b"BM");
        bmp.extend_from_slice(&70u32.to_le_bytes()); // file size
        bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
        bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset
        // DIB header (40 bytes)
        bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
        bmp.extend_from_slice(&1i32.to_le_bytes()); // width
        bmp.extend_from_slice(&1i32.to_le_bytes()); // height
        bmp.extend_from_slice(&1u16.to_le_bytes()); // planes
        bmp.extend_from_slice(&24u16.to_le_bytes()); // bpp
        bmp.extend_from_slice(&0u32.to_le_bytes()); // compression
        bmp.extend_from_slice(&16u32.to_le_bytes()); // image size
        bmp.extend_from_slice(&2835u32.to_le_bytes()); // h resolution
        bmp.extend_from_slice(&2835u32.to_le_bytes()); // v resolution
        bmp.extend_from_slice(&0u32.to_le_bytes()); // colors
        bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors
        // Pixel data: 1 pixel (BGR) + 1 byte padding to align to 4 bytes
        bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);
        bmp
    }

    /// Create a minimal valid SVG image for test images.
    fn make_test_svg() -> Vec<u8> {
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1" viewBox="0 0 1 1"><rect width="1" height="1" fill="#ff0000"/></svg>"##.to_vec()
    }

    /// Create a picture XML element referencing an image via relationship ID.
    fn make_pic_xml(x: i64, y: i64, cx: i64, cy: i64, r_embed: &str) -> String {
        make_custom_pic_xml(
            x,
            y,
            cx,
            cy,
            &format!(r#"<a:blip r:embed="{r_embed}"/><a:stretch><a:fillRect/></a:stretch>"#),
        )
    }

    /// Create a picture XML element with custom `<p:blipFill>` contents.
    fn make_custom_pic_xml(x: i64, y: i64, cx: i64, cy: i64, blip_fill_xml: &str) -> String {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip_fill_xml}</p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr></p:pic>"#
        )
    }

    /// Slide image for the test PPTX builder.
    struct TestSlideImage {
        rid: String,
        path: String,
        data: Vec<u8>,
        relationship_type: Option<String>,
    }

    /// Build a PPTX file with slides that have image relationships.
    fn build_test_pptx_with_images(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slides: &[(String, Vec<TestSlideImage>)],
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        ct.push_str(r#"<Default Extension="png" ContentType="image/png"/>"#);
        ct.push_str(r#"<Default Extension="bmp" ContentType="image/bmp"/>"#);
        ct.push_str(r#"<Default Extension="jpeg" ContentType="image/jpeg"/>"#);
        ct.push_str(r#"<Default Extension="svg" ContentType="image/svg+xml"/>"#);
        for i in 0..slides.len() {
            ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
        }
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/presentation.xml
        let mut pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
            slide_cx_emu, slide_cy_emu
        );
        for i in 0..slides.len() {
            pres.push_str(&format!(
                r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                256 + i,
                2 + i
            ));
        }
        pres.push_str("</p:sldIdLst></p:presentation>");
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        let mut pres_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for i in 0..slides.len() {
            pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
        }
        pres_rels.push_str("</Relationships>");
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // Slides and their .rels files
        for (i, (slide_xml, slide_images)) in slides.iter().enumerate() {
            let slide_num = i + 1;

            // Write slide XML
            zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();

            // Write slide .rels if there are images
            if !slide_images.is_empty() {
                let mut rels = String::from(
                    r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                for img in slide_images {
                    rels.push_str(&format!(
                        r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                        img.rid,
                        img.relationship_type
                            .as_deref()
                            .unwrap_or("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                        img.path
                    ));
                }
                rels.push_str("</Relationships>");
                zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
                    .unwrap();
                zip.write_all(rels.as_bytes()).unwrap();

                // Write image media files
                for img in slide_images {
                    // Resolve the relative path (e.g., "../media/image1.png" → "ppt/media/image1.png")
                    let media_path = resolve_relative_path("ppt/slides", &img.path);
                    zip.start_file(media_path, opts).unwrap();
                    zip.write_all(&img.data).unwrap();
                }
            }
        }

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Helper: get the ImageData from a FixedElement or panic.
    fn get_image(elem: &FixedElement) -> &ImageData {
        match &elem.kind {
            FixedElementKind::Image(img) => img,
            other => panic!("Expected Image, got {other:?}"),
        }
    }

    // ── Shape tests ──────────────────────────────────────────────────────

    #[test]
    fn test_shape_rectangle_with_fill() {
        let shape = make_shape(
            1_000_000,
            500_000,
            3_000_000,
            2_000_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 shape element");

        let elem = &page.elements[0];
        assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
        assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
        assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
        assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

        let shape = get_shape(elem);
        assert!(matches!(shape.kind, ShapeKind::Rectangle));
        assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
        assert!(shape.stroke.is_none());
    }

    #[test]
    fn test_shape_ellipse() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "ellipse",
            Some("00FF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(matches!(s.kind, ShapeKind::Ellipse));
        assert_eq!(s.fill, Some(Color::new(0, 255, 0)));
    }

    #[test]
    fn test_shape_line() {
        let shape = make_shape(
            500_000,
            1_000_000,
            4_000_000,
            0,
            "line",
            None,
            Some(25400),
            Some("0000FF"),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Line { x2, y2 } => {
                assert!((*x2 - emu_to_pt(4_000_000)).abs() < 0.1);
                assert!((*y2 - 0.0).abs() < 0.1);
            }
            _ => panic!("Expected Line shape"),
        }
        assert!(s.fill.is_none());
        let stroke = s.stroke.as_ref().expect("Expected stroke on line");
        assert!((stroke.width - 2.0).abs() < 0.1); // 25400 EMU = 2pt
        assert_eq!(stroke.color, Color::new(0, 0, 255));
    }

    #[test]
    fn test_shape_with_fill_and_border() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            1_000_000,
            "rect",
            Some("FFFF00"),
            Some(12700),
            Some("000000"),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert_eq!(s.fill, Some(Color::new(255, 255, 0)));
        let stroke = s.stroke.as_ref().expect("Expected stroke");
        assert!((stroke.width - 1.0).abs() < 0.1); // 12700 EMU = 1pt
        assert_eq!(stroke.color, Color::black());
    }

    #[test]
    fn test_shape_no_fill_no_border() {
        let shape = make_shape(0, 0, 1_000_000, 1_000_000, "rect", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(s.fill.is_none());
        assert!(s.stroke.is_none());
    }

    #[test]
    fn test_multiple_shapes_on_slide() {
        let rect = make_shape(
            0,
            0,
            1_000_000,
            1_000_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let ellipse = make_shape(
            2_000_000,
            0,
            1_000_000,
            1_000_000,
            "ellipse",
            Some("00FF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[rect, ellipse]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 shape elements");
        assert!(matches!(
            get_shape(&page.elements[0]).kind,
            ShapeKind::Rectangle
        ));
        assert!(matches!(
            get_shape(&page.elements[1]).kind,
            ShapeKind::Ellipse
        ));
    }

    #[test]
    fn test_shapes_and_text_boxes_mixed() {
        let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Hello");
        let rect = make_shape(
            0,
            1_000_000,
            2_000_000,
            500_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let slide = make_slide_xml(&[text_box, rect]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 elements");
        assert!(matches!(
            &page.elements[0].kind,
            FixedElementKind::TextBox(_)
        ));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
    }

    // ── Image tests ──────────────────────────────────────────────────────

    #[test]
    fn test_image_basic_extraction() {
        let bmp_data = make_test_bmp();
        let pic = make_pic_xml(1_000_000, 500_000, 3_000_000, 2_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 image element");

        let elem = &page.elements[0];
        assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
        assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
        assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
        assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

        let img = get_image(elem);
        assert!(!img.data.is_empty(), "Image data should not be empty");
        assert_eq!(img.data, bmp_data);
    }

    #[test]
    fn test_image_format_detection() {
        let bmp_data = make_test_bmp();

        // Test BMP format
        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        assert_eq!(img.format, ImageFormat::Bmp);
    }

    #[test]
    fn test_svg_image_extraction() {
        let svg_data = make_test_svg();

        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.svg".to_string(),
            data: svg_data.clone(),
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 image element");

        let img = get_image(&page.elements[0]);
        assert_eq!(img.format, ImageFormat::Svg);
        assert_eq!(img.data, svg_data);
    }

    #[test]
    fn test_image_blip_start_tag_with_children_is_extracted() {
        let bmp_data = make_test_bmp();
        let pic = make_custom_pic_xml(
            0,
            0,
            1_000_000,
            1_000_000,
            r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
        );
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 image element");

        let img = get_image(&page.elements[0]);
        assert_eq!(img.data, bmp_data);
    }

    #[test]
    fn test_svg_blip_is_preferred_over_base_raster() {
        let bmp_data = make_test_bmp();
        let svg_data = make_test_svg();
        let pic = make_custom_pic_xml(
            0,
            0,
            1_000_000,
            1_000_000,
            r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId4"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
        );
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![
            TestSlideImage {
                rid: "rId3".to_string(),
                path: "../media/image1.bmp".to_string(),
                data: bmp_data,
                relationship_type: None,
            },
            TestSlideImage {
                rid: "rId4".to_string(),
                path: "../media/image2.svg".to_string(),
                data: svg_data.clone(),
                relationship_type: None,
            },
        ];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        assert_eq!(img.format, ImageFormat::Svg);
        assert_eq!(img.data, svg_data);
    }

    #[test]
    fn test_src_rect_crop_is_extracted() {
        let bmp_data = make_test_bmp();
        let pic = make_custom_pic_xml(
            0,
            0,
            2_000_000,
            1_000_000,
            r#"<a:blip r:embed="rId3"/><a:srcRect l="25000" t="10000" r="5000" b="20000"/><a:stretch><a:fillRect/></a:stretch>"#,
        );
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        assert_eq!(
            img.crop,
            Some(ImageCrop {
                left: 0.25,
                top: 0.10,
                right: 0.05,
                bottom: 0.20,
            })
        );
    }

    #[test]
    fn test_unsupported_img_layer_emits_partial_warning_but_keeps_base_image() {
        let bmp_data = make_test_bmp();
        let pic = make_custom_pic_xml(
            0,
            0,
            1_000_000,
            1_000_000,
            r#"<a:blip r:embed="rId3"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId4"/></a14:imgProps></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch>"#,
        );
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![
            TestSlideImage {
                rid: "rId3".to_string(),
                path: "../media/image1.bmp".to_string(),
                data: bmp_data.clone(),
                relationship_type: None,
            },
            TestSlideImage {
                rid: "rId4".to_string(),
                path: "../media/image2.wdp".to_string(),
                data: vec![0x00, 0x01, 0x02],
                relationship_type: Some(
                    "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
                ),
            },
        ];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Base image should still render");
        assert_eq!(get_image(&page.elements[0]).data, bmp_data);
        assert!(
            warnings.iter().any(|warning| matches!(
                warning,
                ConvertWarning::PartialElement { format, element, detail }
                    if format == "PPTX"
                        && element.contains("slide 1")
                        && detail.contains("image layer")
                        && detail.contains("image2.wdp")
            )),
            "Expected partial warning for unsupported image layer, got: {warnings:?}"
        );
    }

    #[test]
    fn test_wdp_only_picture_emits_unsupported_warning() {
        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.wdp".to_string(),
            data: vec![0x00, 0x01, 0x02],
            relationship_type: Some(
                "http://schemas.microsoft.com/office/2007/relationships/hdphoto".to_string(),
            ),
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(
            page.elements.len(),
            0,
            "Unsupported WDP image should be omitted"
        );
        assert!(
            warnings.iter().any(|warning| matches!(
                warning,
                ConvertWarning::UnsupportedElement { format, element }
                    if format == "PPTX"
                        && element.contains("slide 1")
                        && element.contains("image1.wdp")
            )),
            "Expected unsupported warning for WDP-only picture, got: {warnings:?}"
        );
    }

    #[test]
    fn test_image_dimensions_preserved() {
        let bmp_data = make_test_bmp();
        // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000 EMU
        let pic = make_pic_xml(0, 0, 2_540_000, 1_270_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        let w = img.width.expect("Expected width");
        let h = img.height.expect("Expected height");
        assert!((w - 200.0).abs() < 0.1, "Expected ~200pt, got {w}");
        assert!((h - 100.0).abs() < 0.1, "Expected ~100pt, got {h}");
    }

    #[test]
    fn test_image_with_shapes_and_text() {
        let bmp_data = make_test_bmp();
        let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Title");
        let rect = make_shape(
            0,
            600_000,
            1_000_000,
            500_000,
            "rect",
            Some("AABBCC"),
            None,
            None,
        );
        let pic = make_pic_xml(2_000_000, 600_000, 1_500_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[text_box, rect, pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
            relationship_type: None,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 3, "Expected 3 elements");
        assert!(matches!(
            &page.elements[0].kind,
            FixedElementKind::TextBox(_)
        ));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
        assert!(matches!(&page.elements[2].kind, FixedElementKind::Image(_)));
    }

    #[test]
    fn test_image_missing_rid_ignored() {
        // Picture references rId3 but no image data for that rId
        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId99");
        let slide_xml = make_slide_xml(&[pic]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(
            page.elements.len(),
            0,
            "Missing image ref should be skipped"
        );
    }

    #[test]
    fn test_multiple_images_on_slide() {
        let bmp_data = make_test_bmp();
        let pic1 = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let pic2 = make_pic_xml(2_000_000, 0, 1_500_000, 1_000_000, "rId4");
        let slide_xml = make_slide_xml(&[pic1, pic2]);
        let slide_images = vec![
            TestSlideImage {
                rid: "rId3".to_string(),
                path: "../media/image1.bmp".to_string(),
                data: bmp_data.clone(),
                relationship_type: None,
            },
            TestSlideImage {
                rid: "rId4".to_string(),
                path: "../media/image2.bmp".to_string(),
                data: bmp_data,
                relationship_type: None,
            },
        ];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 image elements");
        assert!(matches!(&page.elements[0].kind, FixedElementKind::Image(_)));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Image(_)));
    }

    // ── Theme test helpers ────────────────────────────────────────────

    /// Create a theme XML with the given color scheme and font scheme.
    fn make_theme_xml(colors: &[(&str, &str)], major_font: &str, minor_font: &str) -> String {
        let mut color_xml = String::new();
        for (name, hex) in colors {
            // dk1/lt1 use sysClr in real files; others use srgbClr
            if *name == "dk1" || *name == "lt1" {
                color_xml.push_str(&format!(
                    r#"<a:{name}><a:sysClr val="windowText" lastClr="{hex}"/></a:{name}>"#
                ));
            } else {
                color_xml.push_str(&format!(r#"<a:{name}><a:srgbClr val="{hex}"/></a:{name}>"#));
            }
        }
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Test">{color_xml}</a:clrScheme><a:fontScheme name="Test"><a:majorFont><a:latin typeface="{major_font}"/></a:majorFont><a:minorFont><a:latin typeface="{minor_font}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
        )
    }

    /// Standard theme color set used in tests.
    fn standard_theme_colors() -> Vec<(&'static str, &'static str)> {
        vec![
            ("dk1", "000000"),
            ("dk2", "1F4D78"),
            ("lt1", "FFFFFF"),
            ("lt2", "E7E6E6"),
            ("accent1", "4472C4"),
            ("accent2", "ED7D31"),
            ("accent3", "A5A5A5"),
            ("accent4", "FFC000"),
            ("accent5", "5B9BD5"),
            ("accent6", "70AD47"),
            ("hlink", "0563C1"),
            ("folHlink", "954F72"),
        ]
    }

    /// Build a test PPTX with a theme file included.
    fn build_test_pptx_with_theme(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xmls: &[String],
        theme_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        for i in 0..slide_xmls.len() {
            ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
        }
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/presentation.xml
        let mut pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
            slide_cx_emu, slide_cy_emu
        );
        for i in 0..slide_xmls.len() {
            pres.push_str(&format!(
                r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                256 + i,
                2 + i
            ));
        }
        pres.push_str("</p:sldIdLst></p:presentation>");
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels (includes theme relationship)
        let mut pres_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        // Theme relationship (rId1 in pres rels)
        pres_rels.push_str(
            r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        );
        for i in 0..slide_xmls.len() {
            pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
        }
        pres_rels.push_str("</Relationships>");
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // ppt/theme/theme1.xml
        zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
        zip.write_all(theme_xml.as_bytes()).unwrap();

        // Slides
        for (i, slide_xml) in slide_xmls.iter().enumerate() {
            zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();
        }

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Build a test PPTX with a single slide that has layout and master relationships.
    ///
    /// Creates: slide1 → slideLayout1 → slideMaster1
    fn build_test_pptx_with_layout_master(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xml: &str,
        layout_xml: &str,
        master_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let ct = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#;
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/presentation.xml
        let pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
        );
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#;
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // ppt/slides/slide1.xml
        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        // ppt/slides/_rels/slide1.xml.rels → points to layout
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        // ppt/slideLayouts/slideLayout1.xml
        zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
            .unwrap();
        zip.write_all(layout_xml.as_bytes()).unwrap();

        // ppt/slideLayouts/_rels/slideLayout1.xml.rels → points to master
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
        zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
            .unwrap();
        zip.write_all(layout_rels.as_bytes()).unwrap();

        // ppt/slideMasters/slideMaster1.xml
        zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
            .unwrap();
        zip.write_all(master_xml.as_bytes()).unwrap();

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Build a test PPTX with a single slide, layout/master chain, and theme.
    fn build_test_pptx_with_theme_layout_master(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xml: &str,
        layout_xml: &str,
        master_xml: &str,
        theme_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        let ct = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#;
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        let pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
        );
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
        zip.write_all(theme_xml.as_bytes()).unwrap();

        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
            .unwrap();
        zip.write_all(layout_xml.as_bytes()).unwrap();

        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
        zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
            .unwrap();
        zip.write_all(layout_rels.as_bytes()).unwrap();

        zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
            .unwrap();
        zip.write_all(master_xml.as_bytes()).unwrap();

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Build a test PPTX with multiple slides that all share the same layout and master.
    fn build_test_pptx_with_layout_master_multi_slide(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xmls: &[String],
        layout_xml: &str,
        master_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        for i in 0..slide_xmls.len() {
            ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
        }
        ct.push_str(r#"<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>"#);
        ct.push_str(r#"<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>"#);
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/presentation.xml
        let mut pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst>"#,
        );
        for i in 0..slide_xmls.len() {
            pres.push_str(&format!(
                r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                256 + i,
                2 + i
            ));
        }
        pres.push_str("</p:sldIdLst></p:presentation>");
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        let mut pres_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        pres_rels.push_str(
            r#"<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#,
        );
        for i in 0..slide_xmls.len() {
            pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
        }
        pres_rels.push_str("</Relationships>");
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // Slides and their .rels
        for (i, slide_xml) in slide_xmls.iter().enumerate() {
            let slide_num = i + 1;
            zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();

            // Each slide's .rels points to the shared layout
            let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
            zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
                .unwrap();
            zip.write_all(slide_rels.as_bytes()).unwrap();
        }

        // ppt/slideLayouts/slideLayout1.xml
        zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
            .unwrap();
        zip.write_all(layout_xml.as_bytes()).unwrap();

        // ppt/slideLayouts/_rels/slideLayout1.xml.rels → points to master
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
        zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
            .unwrap();
        zip.write_all(layout_rels.as_bytes()).unwrap();

        // ppt/slideMasters/slideMaster1.xml
        zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
            .unwrap();
        zip.write_all(master_xml.as_bytes()).unwrap();

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    // ── Theme unit tests ──────────────────────────────────────────────

    #[test]
    fn test_parse_theme_xml_colors() {
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let theme = parse_theme_xml(&theme_xml);

        assert_eq!(theme.colors.len(), 12);
        assert_eq!(theme.colors["dk1"], Color::new(0, 0, 0));
        assert_eq!(theme.colors["lt1"], Color::new(255, 255, 255));
        assert_eq!(theme.colors["accent1"], Color::new(0x44, 0x72, 0xC4));
        assert_eq!(theme.colors["accent2"], Color::new(0xED, 0x7D, 0x31));
        assert_eq!(theme.colors["hlink"], Color::new(0x05, 0x63, 0xC1));
        assert_eq!(theme.colors["folHlink"], Color::new(0x95, 0x4F, 0x72));
    }

    #[test]
    fn test_parse_theme_xml_fonts() {
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let theme = parse_theme_xml(&theme_xml);

        assert_eq!(theme.major_font, Some("Calibri Light".to_string()));
        assert_eq!(theme.minor_font, Some("Calibri".to_string()));
    }

    #[test]
    fn test_parse_theme_xml_sys_clr() {
        // dk1 and lt1 use sysClr with lastClr attribute
        let theme_xml = make_theme_xml(&[("dk1", "111111"), ("lt1", "EEEEEE")], "Arial", "Arial");
        let theme = parse_theme_xml(&theme_xml);

        assert_eq!(theme.colors["dk1"], Color::new(0x11, 0x11, 0x11));
        assert_eq!(theme.colors["lt1"], Color::new(0xEE, 0xEE, 0xEE));
    }

    #[test]
    fn test_parse_theme_xml_empty() {
        let theme = parse_theme_xml("");
        assert!(theme.colors.is_empty());
        assert!(theme.major_font.is_none());
        assert!(theme.minor_font.is_none());
    }

    #[test]
    fn test_resolve_theme_font_major() {
        let theme = ThemeData {
            major_font: Some("Calibri Light".to_string()),
            minor_font: Some("Calibri".to_string()),
            ..ThemeData::default()
        };
        assert_eq!(resolve_theme_font("+mj-lt", &theme), "Calibri Light");
    }

    #[test]
    fn test_resolve_theme_font_minor() {
        let theme = ThemeData {
            major_font: Some("Calibri Light".to_string()),
            minor_font: Some("Calibri".to_string()),
            ..ThemeData::default()
        };
        assert_eq!(resolve_theme_font("+mn-lt", &theme), "Calibri");
    }

    #[test]
    fn test_resolve_theme_font_explicit() {
        let theme = ThemeData::default();
        assert_eq!(resolve_theme_font("Arial", &theme), "Arial");
    }

    // ── Theme integration tests (full PPTX parsing) ───────────────────

    #[test]
    fn test_scheme_color_in_shape_fill() {
        // Shape with <a:schemeClr val="accent1"/> should resolve to accent1 color
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>"#;
        let slide = make_slide_xml(&[shape_xml.to_string()]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);
        let shape = get_shape(&page.elements[0]);
        assert_eq!(shape.fill, Some(Color::new(0x44, 0x72, 0xC4)));
    }

    #[test]
    fn test_scheme_color_in_line_stroke() {
        // Shape border using scheme color
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:ln></p:spPr></p:sp>"#;
        let slide = make_slide_xml(&[shape_xml.to_string()]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape = get_shape(&page.elements[0]);
        let stroke = shape.stroke.as_ref().expect("Expected stroke");
        assert_eq!(stroke.color, Color::new(0, 0, 0)); // dk1 = black
    }

    #[test]
    fn test_scheme_color_in_text_run() {
        // Text run using <a:schemeClr val="accent2"/>
        let runs_xml = r#"<a:r><a:rPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:rPr><a:t>Themed text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Themed text");
        assert_eq!(para.runs[0].style.color, Some(Color::new(0xED, 0x7D, 0x31)));
    }

    #[test]
    fn test_theme_major_font_in_text() {
        // Text with <a:latin typeface="+mj-lt"/> should resolve to major font
        let runs_xml =
            r#"<a:r><a:rPr><a:latin typeface="+mj-lt"/></a:rPr><a:t>Heading</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Heading");
        assert_eq!(
            para.runs[0].style.font_family,
            Some("Calibri Light".to_string())
        );
    }

    #[test]
    fn test_theme_minor_font_in_text() {
        // Text with <a:latin typeface="+mn-lt"/> should resolve to minor font
        let runs_xml =
            r#"<a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>Body text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Body text");
        assert_eq!(para.runs[0].style.font_family, Some("Calibri".to_string()));
    }

    #[test]
    fn test_pptx_with_theme_colors_and_fonts_combined() {
        // Full test: shape with scheme color + text with scheme color and theme font
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent5"/></a:solidFill></p:spPr></p:sp>"#;
        let runs_xml = r#"<a:r><a:rPr b="1" sz="2400"><a:solidFill><a:schemeClr val="dk2"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Theme styled</a:t></a:r>"#;
        let text_box = make_formatted_text_box(3_000_000, 0, 4_000_000, 1_000_000, runs_xml);
        let slide = make_slide_xml(&[shape_xml.to_string(), text_box]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2);

        // Shape fill = accent5
        let shape = get_shape(&page.elements[0]);
        assert_eq!(shape.fill, Some(Color::new(0x5B, 0x9B, 0xD5)));

        // Text run: color = dk2, font = major font, bold, 24pt
        let blocks = text_box_blocks(&page.elements[1]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        let run = &para.runs[0];
        assert_eq!(run.text, "Theme styled");
        assert_eq!(run.style.color, Some(Color::new(0x1F, 0x4D, 0x78)));
        assert_eq!(run.style.font_family, Some("Calibri Light".to_string()));
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.font_size, Some(24.0));
    }

    #[test]
    fn test_no_theme_scheme_color_ignored() {
        // When there's no theme, schemeClr references should produce None
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>"#;
        let slide = make_slide_xml(&[shape_xml.to_string()]);
        // Use regular build_test_pptx (no theme)
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape = get_shape(&page.elements[0]);
        // No theme → scheme color not resolved → fill is None
        assert!(shape.fill.is_none());
    }

    #[test]
    fn test_scheme_color_as_start_element() {
        // schemeClr can have children like <a:tint val="50000"/>, test it still works
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"><a:tint val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
        let slide = make_slide_xml(&[shape_xml.to_string()]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape = get_shape(&page.elements[0]);
        // Color is resolved from the scheme (tint is ignored for now but base color is read)
        assert_eq!(shape.fill, Some(Color::new(0xA5, 0xA5, 0xA5)));
    }

    #[test]
    fn test_scheme_color_lum_mod_applies_to_shape_fill() {
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>"#;
        let slide = make_slide_xml(&[shape_xml.to_string()]);
        let theme_xml = make_theme_xml(
            &[("dk1", "000000"), ("lt1", "FFFFFF"), ("accent1", "808080")],
            "Calibri Light",
            "Calibri",
        );
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape = get_shape(&page.elements[0]);
        assert_eq!(shape.fill, Some(Color::new(0x40, 0x40, 0x40)));
    }

    #[test]
    fn test_layout_shape_uses_master_color_map_with_luminance_offset() {
        let slide_xml = make_empty_slide_xml();
        let layout_shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="tx1"><a:lumOff val="50000"/></a:schemeClr></a:solidFill><a:ln w="6350"><a:noFill/></a:ln></p:spPr></p:sp>"#;
        let layout_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{layout_shape}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#
        );
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt1" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#;
        let theme_xml = make_theme_xml(
            &[
                ("dk1", "000000"),
                ("dk2", "222222"),
                ("lt1", "FFFFFF"),
                ("lt2", "EEEEEE"),
                ("accent1", "4472C4"),
                ("accent2", "ED7D31"),
                ("accent3", "A5A5A5"),
                ("accent4", "FFC000"),
                ("accent5", "5B9BD5"),
                ("accent6", "70AD47"),
                ("hlink", "0563C1"),
                ("folHlink", "954F72"),
            ],
            "Calibri Light",
            "Calibri",
        );
        let data = build_test_pptx_with_theme_layout_master(
            SLIDE_CX,
            SLIDE_CY,
            &slide_xml,
            &layout_xml,
            master_xml,
            &theme_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape = get_shape(&page.elements[0]);
        assert_eq!(shape.fill, Some(Color::new(0x80, 0x80, 0x80)));
    }

    // ── Slide background tests ───────────────────────────────────────────

    #[test]
    fn test_slide_solid_color_background() {
        // Slide with a solid red background via <p:bg>
        let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
        let slide = make_slide_xml_with_bg(bg_xml, &[]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_slide_no_background() {
        // Slide with no <p:bg> → background_color is None
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert!(page.background_color.is_none());
    }

    #[test]
    fn test_slide_background_with_scheme_color() {
        // Slide background using a theme scheme color reference
        let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
        let slide = make_slide_xml_with_bg(bg_xml, &[]);
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri Light", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide], &theme_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.background_color, Some(Color::new(0x44, 0x72, 0xC4)));
    }

    #[test]
    fn test_slide_background_with_text_content() {
        // Slide with both background and text shapes — both should be present
        let bg_xml = r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#;
        let text_box = make_text_box(100000, 100000, 5000000, 500000, "Hello");
        let slide = make_slide_xml_with_bg(bg_xml, &[text_box]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.background_color, Some(Color::new(0, 0, 255)));
        assert_eq!(page.elements.len(), 1);
    }

    #[test]
    fn test_slide_inherits_master_background() {
        // Slide has no background, but its master does → should inherit
        let slide_xml = make_empty_slide_xml();
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

        // Build PPTX with slide → layout → master chain
        let data = build_test_pptx_with_layout_master(
            SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        // Should inherit master's green background
        assert_eq!(page.background_color, Some(Color::new(0, 255, 0)));
    }

    /// Create a slide layout XML with the given shape elements.
    fn make_layout_xml(shapes: &[String]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
        );
        for shape in shapes {
            xml.push_str(shape);
        }
        xml.push_str("</p:spTree></p:cSld></p:sldLayout>");
        xml
    }

    /// Create a slide master XML with the given shape elements.
    fn make_master_xml(shapes: &[String]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
        );
        for shape in shapes {
            xml.push_str(shape);
        }
        xml.push_str("</p:spTree></p:cSld></p:sldMaster>");
        xml
    }

    // ── US-025: Slide master and layout inheritance tests ────────────────

    #[test]
    fn test_master_shape_appears_on_slide() {
        // Master has a rectangle shape → it should appear on the slide
        let slide_xml = make_empty_slide_xml();
        let layout_xml = make_layout_xml(&[]);
        let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Master Logo");
        let master_xml = make_master_xml(&[master_shape]);

        let data = build_test_pptx_with_layout_master(
            SLIDE_CX,
            SLIDE_CY,
            &slide_xml,
            &layout_xml,
            &master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        // Master element should be present
        assert_eq!(page.elements.len(), 1);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Master Logo");
    }

    #[test]
    fn test_layout_shape_appears_on_slide() {
        // Layout has a text box → it should appear on the slide
        let slide_xml = make_empty_slide_xml();
        let layout_shape = make_text_box(100_000, 100_000, 3_000_000, 500_000, "Layout Title");
        let layout_xml = make_layout_xml(&[layout_shape]);
        let master_xml = make_master_xml(&[]);

        let data = build_test_pptx_with_layout_master(
            SLIDE_CX,
            SLIDE_CY,
            &slide_xml,
            &layout_xml,
            &master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Layout Title");
    }

    #[test]
    fn test_inheritance_element_ordering() {
        // Master, layout, and slide all have elements → order: master, layout, slide
        let slide_shape = make_text_box(0, 0, 1_000_000, 500_000, "Slide Content");
        let slide_xml = make_slide_xml(&[slide_shape]);
        let layout_shape = make_text_box(0, 0, 1_000_000, 500_000, "Layout Content");
        let layout_xml = make_layout_xml(&[layout_shape]);
        let master_shape = make_text_box(0, 0, 1_000_000, 500_000, "Master Content");
        let master_xml = make_master_xml(&[master_shape]);

        let data = build_test_pptx_with_layout_master(
            SLIDE_CX,
            SLIDE_CY,
            &slide_xml,
            &layout_xml,
            &master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 3);

        // Master element is first (behind)
        let master_blocks = text_box_blocks(&page.elements[0]);
        let master_para = match &master_blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(master_para.runs[0].text, "Master Content");

        // Layout element is second
        let layout_blocks = text_box_blocks(&page.elements[1]);
        let layout_para = match &layout_blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(layout_para.runs[0].text, "Layout Content");

        // Slide element is last (on top)
        let slide_blocks = text_box_blocks(&page.elements[2]);
        let slide_para = match &slide_blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(slide_para.runs[0].text, "Slide Content");
    }

    #[test]
    fn test_master_elements_appear_on_all_slides() {
        // Build a PPTX with 2 slides and a master shape → both slides should have it
        let master_shape = make_text_box(0, 0, 2_000_000, 500_000, "Company Logo");
        let master_xml = make_master_xml(&[master_shape]);
        let layout_xml = make_layout_xml(&[]);

        let slide1_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 1");
        let slide1_xml = make_slide_xml(&[slide1_shape]);
        let slide2_shape = make_text_box(0, 1_000_000, 5_000_000, 2_000_000, "Slide 2");
        let slide2_xml = make_slide_xml(&[slide2_shape]);

        // Build PPTX with 2 slides, both pointing to same layout/master
        let data = build_test_pptx_with_layout_master_multi_slide(
            SLIDE_CX,
            SLIDE_CY,
            &[slide1_xml, slide2_xml],
            &layout_xml,
            &master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 2);

        // Both slides should have master element + their own element
        for (i, page) in doc.pages.iter().enumerate() {
            let fixed_page = match page {
                Page::Fixed(p) => p,
                _ => panic!("Expected FixedPage"),
            };
            assert_eq!(
                fixed_page.elements.len(),
                2,
                "Slide {} should have 2 elements (master + slide)",
                i + 1
            );

            // First element is the master shape
            let master_blocks = text_box_blocks(&fixed_page.elements[0]);
            let master_para = match &master_blocks[0] {
                Block::Paragraph(p) => p,
                _ => panic!("Expected Paragraph"),
            };
            assert_eq!(master_para.runs[0].text, "Company Logo");
        }
    }

    #[test]
    fn test_slide_without_layout_master_has_only_slide_elements() {
        // Standard PPTX without layout/master .rels → only slide elements
        let shape = make_text_box(0, 0, 1_000_000, 500_000, "Just Slide");
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].text, "Just Slide");
    }

    #[test]
    fn test_slide_inherits_layout_background_over_master() {
        // Layout has a background, master has a different one → layout wins
        let slide_xml = make_empty_slide_xml();
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

        let data = build_test_pptx_with_layout_master(
            SLIDE_CX, SLIDE_CY, &slide_xml, layout_xml, master_xml,
        );

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        // Should inherit layout's magenta background (not master's green)
        assert_eq!(page.background_color, Some(Color::new(255, 0, 255)));
    }

    // ── Table test helpers ──────────────────────────────────────────────

    /// Create a graphicFrame XML containing a table.
    /// `x`, `y`, `cx`, `cy` are in EMU.
    fn make_table_graphic_frame(
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
        col_widths_emu: &[i64],
        rows_xml: &str,
    ) -> String {
        let mut grid = String::new();
        for w in col_widths_emu {
            grid.push_str(&format!(r#"<a:gridCol w="{w}"/>"#));
        }
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{rows_xml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    /// Create a simple table row with text-only cells.
    fn make_table_row(cells: &[&str]) -> String {
        let mut xml = String::from(r#"<a:tr h="370840">"#);
        for text in cells {
            xml.push_str(&format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#
            ));
        }
        xml.push_str("</a:tr>");
        xml
    }

    /// Helper: get the Table from a FixedElement.
    fn table_element(elem: &FixedElement) -> &Table {
        match &elem.kind {
            FixedElementKind::Table(t) => t,
            _ => panic!("Expected Table, got {:?}", elem.kind),
        }
    }

    // ── Table tests ─────────────────────────────────────────────────────

    #[test]
    fn test_slide_with_basic_table() {
        // A slide with a 2×2 table
        let rows = format!(
            "{}{}",
            make_table_row(&["A1", "B1"]),
            make_table_row(&["A2", "B2"]),
        );
        let table_frame = make_table_graphic_frame(
            914400,              // x = 72pt
            914400,              // y = 72pt
            3657600,             // cx = 288pt
            1828800,             // cy = 144pt
            &[1828800, 1828800], // 2 columns, 144pt each
            &rows,
        );
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);

        let elem = &page.elements[0];
        assert!((elem.x - 72.0).abs() < 0.1);
        assert!((elem.y - 72.0).abs() < 0.1);

        let table = table_element(elem);
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.column_widths.len(), 2);
        assert!((table.column_widths[0] - 144.0).abs() < 0.1);

        // Check cell text
        let cell_00 = &table.rows[0].cells[0];
        assert_eq!(cell_00.content.len(), 1);
        if let Block::Paragraph(p) = &cell_00.content[0] {
            assert_eq!(p.runs[0].text, "A1");
        } else {
            panic!("Expected paragraph in cell");
        }

        let cell_11 = &table.rows[1].cells[1];
        if let Block::Paragraph(p) = &cell_11.content[0] {
            assert_eq!(p.runs[0].text, "B2");
        } else {
            panic!("Expected paragraph in cell");
        }
    }

    #[test]
    fn test_slide_table_scales_geometry_to_graphic_frame_extent() {
        let rows_xml = format!(
            "{}{}",
            make_table_row(&["A1", "B1"]),
            make_table_row(&["A2", "B2"]),
        );
        let table_frame = make_table_graphic_frame(
            914400,
            914400,
            3_657_600,
            1_483_360,
            &[914_400, 914_400],
            &rows_xml,
        );
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let elem = &page.elements[0];
        let table = table_element(elem);

        assert_eq!(table.column_widths.len(), 2);
        assert!((table.column_widths[0] - 144.0).abs() < 0.1);
        assert!((table.column_widths.iter().sum::<f64>() - elem.width).abs() < 0.1);

        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].height, Some(58.4));
        assert_eq!(table.rows[1].height, Some(58.4));
        assert!(
            (table
                .rows
                .iter()
                .map(|row| row.height.unwrap_or(0.0))
                .sum::<f64>()
                - elem.height)
                .abs()
                < 0.1
        );
    }

    #[test]
    fn test_slide_table_reads_column_widths_from_gridcol_with_extensions() {
        let rows_xml = make_table_row(&["A1", "B1"]);
        let table_frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="370840"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="1"/></a:ext></a:extLst></a:gridCol><a:gridCol w="914400"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"><a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="2"/></a:ext></a:extLst></a:gridCol></a:tblGrid>"#.to_string()
            + &rows_xml
            + r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        assert_eq!(table.column_widths.len(), 2);
        assert!((table.column_widths[0] - 72.0).abs() < 0.1);
        assert!((table.column_widths[1] - 72.0).abs() < 0.1);
    }

    #[test]
    fn test_slide_table_cell_anchor_maps_to_vertical_alignment() {
        let rows_xml = concat!(
            r#"<a:tr h="370840">"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Centered</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bottom</a:t></a:r></a:p></a:txBody><a:tcPr anchor="b"/></a:tc>"#,
            r#"</a:tr>"#,
        );
        let table_frame =
            make_table_graphic_frame(0, 0, 1_828_800, 370_840, &[914_400, 914_400], rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        assert_eq!(
            table.rows[0].cells[0].vertical_align,
            Some(crate::ir::CellVerticalAlign::Center)
        );
        assert_eq!(
            table.rows[0].cells[1].vertical_align,
            Some(crate::ir::CellVerticalAlign::Bottom)
        );
    }

    #[test]
    fn test_slide_table_cell_margins_map_to_padding() {
        let rows_xml = concat!(
            r#"<a:tr h="370840">"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Padded</a:t></a:r></a:p></a:txBody><a:tcPr marL="76200" marR="76200" marT="38100" marB="38100"/></a:tc>"#,
            r#"</a:tr>"#,
        );
        let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        assert_eq!(
            table.rows[0].cells[0].padding,
            Some(crate::ir::Insets {
                top: 3.0,
                right: 6.0,
                bottom: 3.0,
                left: 6.0,
            })
        );
    }

    #[test]
    fn test_slide_table_uses_powerpoint_default_cell_padding() {
        let rows_xml = concat!(
            r#"<a:tr h="370840">"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>DefaultPadding</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#,
            r#"</a:tr>"#,
        );
        let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        assert_eq!(
            table.default_cell_padding,
            Some(crate::ir::Insets {
                top: 3.6,
                right: 7.2,
                bottom: 3.6,
                left: 7.2,
            })
        );
        assert_eq!(table.rows[0].cells[0].padding, None);
        assert!(table.use_content_driven_row_heights);
    }

    #[test]
    fn test_slide_table_coalesces_adjacent_runs_with_same_style() {
        let rows_xml = concat!(
            r#"<a:tr h="370840">"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t>YOLOv8n + </a:t></a:r>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1100" err="1"><a:latin typeface="Arial"/></a:rPr><a:t>topk filtering on gpu(</a:t></a:r>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>K</a:t></a:r>"#,
            r#"<a:r><a:rPr lang="en-US" sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t> = 100)</a:t></a:r>"#,
            r#"</a:p></a:txBody><a:tcPr/></a:tc>"#,
            r#"</a:tr>"#,
        );
        let table_frame = make_table_graphic_frame(0, 0, 914_400, 370_840, &[914_400], rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);
        let paragraph = match &table.rows[0].cells[0].content[0] {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("Expected paragraph, got {other:?}"),
        };

        assert_eq!(paragraph.runs.len(), 3);
        assert_eq!(paragraph.runs[0].text, "YOLOv8n + topk filtering on gpu(");
        assert_eq!(paragraph.runs[1].text, "K");
        assert_eq!(paragraph.runs[2].text, "\u{00A0}= 100)");
        assert_eq!(paragraph.runs[1].style.italic, Some(true));
    }

    #[test]
    fn test_slide_table_cell_bulleted_paragraphs_group_into_list() {
        let rows_xml = concat!(
            r#"<a:tr h="740000">"#,
            r#"<a:tc><a:txBody><a:bodyPr/>"#,
            r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>First bullet</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr indent="-216000"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Second bullet</a:t></a:r></a:p>"#,
            r#"</a:txBody><a:tcPr/></a:tc>"#,
            r#"</a:tr>"#,
        );
        let table_frame = make_table_graphic_frame(0, 0, 914_400, 740_000, &[914_400], rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);
        assert_eq!(table.rows[0].cells[0].content.len(), 1);

        let list = match &table.rows[0].cells[0].content[0] {
            Block::List(list) => list,
            other => panic!("Expected List block, got {other:?}"),
        };
        assert_eq!(list.kind, crate::ir::ListKind::Unordered);
        assert_eq!(list.items.len(), 2);
        assert_eq!(list.items[0].content[0].runs[0].text, "First bullet");
        assert_eq!(list.items[1].content[0].runs[0].text, "Second bullet");
    }

    #[test]
    fn test_slide_table_with_merged_cells() {
        // Table with gridSpan (horizontal merge) and vMerge (vertical merge)
        let mut rows_xml = String::new();
        // Row 0: cell spanning 2 columns
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str(r#"<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str("</a:tr>");
        // Row 1: two normal cells
        rows_xml.push_str(&make_table_row(&["C1", "C2"]));

        let table_frame =
            make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        // Row 0: merged cell should have col_span=2
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(table.rows[0].cells[0].col_span, 2);
        // The hMerge cell should have col_span=0 (covered by merge)
        assert_eq!(table.rows[0].cells[1].col_span, 0);

        // Row 1: normal cells
        assert_eq!(table.rows[1].cells[0].col_span, 1);
        assert_eq!(table.rows[1].cells[1].col_span, 1);
    }

    #[test]
    fn test_slide_table_with_vertical_merge() {
        // Table with rowSpan (vertical merge)
        let mut rows_xml = String::new();
        // Row 0: first cell starts a rowSpan of 2
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>VMerged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str("</a:tr>");
        // Row 1: first cell is continuation of vMerge
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc vMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>B2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str("</a:tr>");

        let table_frame =
            make_table_graphic_frame(0, 0, 3657600, 1828800, &[1828800, 1828800], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        // Row 0: first cell rowSpan=2
        assert_eq!(table.rows[0].cells[0].row_span, 2);
        // Row 1: first cell vMerge continuation (row_span=0)
        assert_eq!(table.rows[1].cells[0].row_span, 0);
    }

    #[test]
    fn test_slide_table_with_formatted_text() {
        // Table cell with bold, colored text
        let mut rows_xml = String::new();
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" b="1" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Bold Red</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#);
        rows_xml.push_str("</a:tr>");

        let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        let cell = &table.rows[0].cells[0];
        if let Block::Paragraph(p) = &cell.content[0] {
            assert_eq!(p.runs[0].text, "Bold Red");
            assert_eq!(p.runs[0].style.bold, Some(true));
            assert_eq!(p.runs[0].style.font_size, Some(18.0));
            assert_eq!(p.runs[0].style.color, Some(Color::new(255, 0, 0)));
        } else {
            panic!("Expected paragraph in cell");
        }
    }

    #[test]
    fn test_slide_table_with_cell_background() {
        // Table cell with background fill
        let mut rows_xml = String::new();
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Filled</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr></a:tc>"#);
        rows_xml.push_str("</a:tr>");

        let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        let cell = &table.rows[0].cells[0];
        assert_eq!(cell.background, Some(Color::new(0, 255, 0)));
    }

    #[test]
    fn test_slide_table_with_cell_borders() {
        // Table cell with border specification
        let mut rows_xml = String::new();
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bordered</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB></a:tcPr></a:tc>"#);
        rows_xml.push_str("</a:tr>");

        let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);

        let cell = &table.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected border");
        assert!(border.left.is_some());
        assert!(border.right.is_some());
        assert!(border.top.is_some());
        assert!(border.bottom.is_some());
        let left = border.left.as_ref().unwrap();
        assert!((left.width - 1.0).abs() < 0.1); // 12700 EMU = 1pt
        assert_eq!(left.color, Color::new(0, 0, 0));
    }

    #[test]
    fn test_slide_table_cell_border_dash_styles() {
        // Table cell with dashed top and dotted bottom borders
        let mut rows_xml = String::new();
        rows_xml.push_str(r#"<a:tr h="370840">"#);
        rows_xml.push_str(r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Dashed</a:t></a:r></a:p></a:txBody><a:tcPr>"#);
        rows_xml.push_str(r#"<a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:lnT>"#);
        rows_xml.push_str(r#"<a:lnB w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dot"/></a:lnB>"#);
        rows_xml.push_str(r#"</a:tcPr></a:tc>"#);
        rows_xml.push_str("</a:tr>");

        let table_frame = make_table_graphic_frame(0, 0, 3657600, 370840, &[3657600], &rows_xml);
        let slide = make_slide_xml(&[table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let table = table_element(&page.elements[0]);
        let cell = &table.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected border");

        let top = border.top.as_ref().expect("Expected top border");
        assert_eq!(top.style, BorderLineStyle::Dashed, "Top should be dashed");

        let bottom = border.bottom.as_ref().expect("Expected bottom border");
        assert_eq!(
            bottom.style,
            BorderLineStyle::Dotted,
            "Bottom should be dotted"
        );
    }

    #[test]
    fn test_shape_outline_dash_style() {
        // Shape with dashed outline
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr></p:sp>"#.to_string();
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let shape_elem = &page.elements[0];
        if let FixedElementKind::Shape(ref s) = shape_elem.kind {
            let stroke = s.stroke.as_ref().expect("Expected stroke");
            assert_eq!(
                stroke.style,
                BorderLineStyle::Dashed,
                "Shape stroke should be dashed"
            );
        } else {
            panic!("Expected Shape element");
        }
    }

    #[test]
    fn test_slide_table_coexists_with_shapes() {
        // A slide with both a text box and a table
        let text_box = make_text_box(0, 0, 914400, 457200, "Header");
        let rows = make_table_row(&["Cell"]);
        let table_frame = make_table_graphic_frame(0, 914400, 914400, 370840, &[914400], &rows);
        let slide = make_slide_xml(&[text_box, table_frame]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2);

        // First element: TextBox
        assert!(matches!(
            &page.elements[0].kind,
            FixedElementKind::TextBox(_)
        ));
        // Second element: Table
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Table(_)));
    }

    // ----- US-029: Slide selection tests -----

    #[test]
    fn test_slide_filter_single_slide() {
        use crate::config::SlideRange;
        let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
        let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
        let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);

        let parser = PptxParser;
        let opts = ConvertOptions {
            slide_range: Some(SlideRange::new(2, 2)),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 1, "Should only include slide 2");
        // Verify slide 2 content
        let page = first_fixed_page(&doc);
        let text = match &page.elements[0].kind {
            FixedElementKind::TextBox(text_box) => match &text_box.content[0] {
                Block::Paragraph(p) => p.runs[0].text.clone(),
                _ => panic!("Expected Paragraph"),
            },
            _ => panic!("Expected TextBox"),
        };
        assert_eq!(text, "Slide 2");
    }

    #[test]
    fn test_slide_filter_range() {
        use crate::config::SlideRange;
        let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
        let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
        let slide3 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 3")]);
        let slide4 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 4")]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3, slide4]);

        let parser = PptxParser;
        let opts = ConvertOptions {
            slide_range: Some(SlideRange::new(2, 3)),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 2, "Should include slides 2 and 3");
    }

    #[test]
    fn test_slide_filter_none_includes_all() {
        let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
        let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

        let parser = PptxParser;
        let opts = ConvertOptions {
            slide_range: None,
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(doc.pages.len(), 2, "None should include all slides");
    }

    #[test]
    fn test_slide_filter_range_beyond_total() {
        use crate::config::SlideRange;
        let slide1 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 1")]);
        let slide2 = make_slide_xml(&[make_text_box(0, 0, 914400, 914400, "Slide 2")]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2]);

        let parser = PptxParser;
        let opts = ConvertOptions {
            slide_range: Some(SlideRange::new(5, 10)),
            ..Default::default()
        };
        let (doc, _warnings) = parser.parse(&data, &opts).unwrap();

        assert_eq!(
            doc.pages.len(),
            0,
            "Range beyond total slides should produce empty document"
        );
    }

    // ── Group shape helpers ─────────────────────────────────────────────

    /// Create a group shape XML with a coordinate transform and child shapes.
    #[allow(clippy::too_many_arguments)]
    fn make_group_shape(
        off_x: i64,
        off_y: i64,
        ext_cx: i64,
        ext_cy: i64,
        ch_off_x: i64,
        ch_off_y: i64,
        ch_ext_cx: i64,
        ch_ext_cy: i64,
        children: &[String],
    ) -> String {
        let mut xml = format!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/><a:chOff x="{ch_off_x}" y="{ch_off_y}"/><a:chExt cx="{ch_ext_cx}" cy="{ch_ext_cy}"/></a:xfrm></p:grpSpPr>"#
        );
        for child in children {
            xml.push_str(child);
        }
        xml.push_str("</p:grpSp>");
        xml
    }

    /// Create a rectangle shape XML (no text body) with a fill color.
    fn make_shape_rect(x: i64, y: i64, cx: i64, cy: i64, fill_hex: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill></p:spPr></p:sp>"#
        )
    }

    // ── Group shape tests ───────────────────────────────────────────────

    #[test]
    fn test_group_shape_two_text_boxes() {
        // Group at (1000000, 500000) with 1:1 mapping (ext == chExt)
        let child_a = make_text_box(0, 0, 2_000_000, 1_000_000, "Shape A");
        let child_b = make_text_box(2_000_000, 1_000_000, 2_000_000, 1_000_000, "Shape B");
        let group = make_group_shape(
            1_000_000,
            500_000, // off
            4_000_000,
            2_000_000, // ext
            0,
            0, // chOff
            4_000_000,
            2_000_000, // chExt
            &[child_a, child_b],
        );
        let slide = make_slide_xml(&[group]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 elements from group");

        // Shape A: (1000000, 500000) in EMU → ~78.74pt, ~39.37pt
        let a = &page.elements[0];
        assert!(
            (a.x - emu_to_pt(1_000_000)).abs() < 0.1,
            "Shape A x: got {}, expected {}",
            a.x,
            emu_to_pt(1_000_000)
        );
        assert!(
            (a.y - emu_to_pt(500_000)).abs() < 0.1,
            "Shape A y: got {}, expected {}",
            a.y,
            emu_to_pt(500_000)
        );

        // Shape B: (1000000+2000000, 500000+1000000) = (3000000, 1500000) EMU
        let b = &page.elements[1];
        assert!(
            (b.x - emu_to_pt(3_000_000)).abs() < 0.1,
            "Shape B x: got {}, expected {}",
            b.x,
            emu_to_pt(3_000_000)
        );
        assert!(
            (b.y - emu_to_pt(1_500_000)).abs() < 0.1,
            "Shape B y: got {}, expected {}",
            b.y,
            emu_to_pt(1_500_000)
        );

        // Verify text content
        let blocks_a = text_box_blocks(a);
        let para_a = match &blocks_a[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para_a.runs[0].text, "Shape A");
    }

    #[test]
    fn test_group_shape_with_scaling() {
        // Group: ext is half of chExt → children scaled down by 0.5
        let child = make_text_box(0, 0, 4_000_000, 2_000_000, "Scaled");
        let group = make_group_shape(
            0,
            0, // off
            2_000_000,
            1_000_000, // ext (half)
            0,
            0, // chOff
            4_000_000,
            2_000_000, // chExt (full)
            &[child],
        );
        let slide = make_slide_xml(&[group]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);

        let elem = &page.elements[0];
        // Width: 4000000 * 0.5 = 2000000 EMU → ~157.48pt
        let expected_w = emu_to_pt(2_000_000);
        assert!(
            (elem.width - expected_w).abs() < 0.1,
            "Scaled width: got {}, expected {}",
            elem.width,
            expected_w
        );
        let expected_h = emu_to_pt(1_000_000);
        assert!(
            (elem.height - expected_h).abs() < 0.1,
            "Scaled height: got {}, expected {}",
            elem.height,
            expected_h
        );
    }

    #[test]
    fn test_nested_group_shapes() {
        // Inner group at (0, 0) with 1:1 mapping containing a text box
        let inner_child = make_text_box(0, 0, 1_000_000, 1_000_000, "Nested");
        let inner_group = make_group_shape(
            0,
            0, // off
            2_000_000,
            2_000_000, // ext
            0,
            0, // chOff
            2_000_000,
            2_000_000, // chExt
            &[inner_child],
        );
        // Outer group at (1000000, 1000000) with 1:1 mapping
        let outer_group = make_group_shape(
            1_000_000,
            1_000_000, // off
            4_000_000,
            4_000_000, // ext
            0,
            0, // chOff
            4_000_000,
            4_000_000, // chExt
            &[inner_group],
        );
        let slide = make_slide_xml(&[outer_group]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(
            page.elements.len(),
            1,
            "Expected 1 element (nested text box)"
        );

        // The inner child at (0,0) → inner group maps to (0,0) → outer group maps to (1000000, 1000000)
        let elem = &page.elements[0];
        assert!(
            (elem.x - emu_to_pt(1_000_000)).abs() < 0.1,
            "Nested x: got {}, expected {}",
            elem.x,
            emu_to_pt(1_000_000)
        );
        assert!(
            (elem.y - emu_to_pt(1_000_000)).abs() < 0.1,
            "Nested y: got {}, expected {}",
            elem.y,
            emu_to_pt(1_000_000)
        );
        assert_eq!(elem.width, emu_to_pt(1_000_000));
        assert_eq!(elem.height, emu_to_pt(1_000_000));
    }

    #[test]
    fn test_group_shape_mixed_element_types() {
        // Group with a text box and a rectangle shape
        let text = make_text_box(0, 0, 2_000_000, 1_000_000, "Text");
        let rect = make_shape_rect(2_000_000, 0, 2_000_000, 1_000_000, "FF0000");
        let group = make_group_shape(
            0,
            0, // off
            4_000_000,
            2_000_000, // ext
            0,
            0, // chOff
            4_000_000,
            2_000_000, // chExt
            &[text, rect],
        );
        let slide = make_slide_xml(&[group]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected TextBox + Shape");

        // First element: TextBox
        assert!(
            matches!(&page.elements[0].kind, FixedElementKind::TextBox(_)),
            "First element should be TextBox"
        );
        // Second element: Shape
        assert!(
            matches!(&page.elements[1].kind, FixedElementKind::Shape(_)),
            "Second element should be Shape"
        );

        // Verify shape position: (2000000, 0) in child space → (2000000, 0) in slide space
        let shape_elem = &page.elements[1];
        assert!(
            (shape_elem.x - emu_to_pt(2_000_000)).abs() < 0.1,
            "Shape x: got {}, expected {}",
            shape_elem.x,
            emu_to_pt(2_000_000)
        );
    }

    #[test]
    fn test_group_shape_with_nonzero_child_offset() {
        // Group where chOff != (0,0) — children positioned relative to offset
        let child = make_text_box(1_000_000, 1_000_000, 2_000_000, 1_000_000, "Offset");
        let group = make_group_shape(
            500_000,
            500_000, // off (group position on slide)
            4_000_000,
            2_000_000, // ext
            1_000_000,
            1_000_000, // chOff (child space origin)
            4_000_000,
            2_000_000, // chExt
            &[child],
        );
        let slide = make_slide_xml(&[group]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1);

        // child_x=1000000, chOff_x=1000000 → (1000000-1000000)*1.0 + 500000 = 500000
        let elem = &page.elements[0];
        assert!(
            (elem.x - emu_to_pt(500_000)).abs() < 0.1,
            "Offset x: got {}, expected {}",
            elem.x,
            emu_to_pt(500_000)
        );
        assert!(
            (elem.y - emu_to_pt(500_000)).abs() < 0.1,
            "Offset y: got {}, expected {}",
            elem.y,
            emu_to_pt(500_000)
        );
    }

    // ── Shape style (rotation, transparency) test helpers ────────────────

    /// Create a shape XML with optional rotation and fill alpha.
    /// `rot` is in 60000ths of a degree (e.g. 5400000 = 90°).
    /// `alpha_thousandths` is in 1000ths of percent (e.g. 50000 = 50%).
    #[allow(clippy::too_many_arguments)]
    fn make_styled_shape(
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
        prst: &str,
        fill_hex: Option<&str>,
        rot: Option<i64>,
        alpha_thousandths: Option<i64>,
    ) -> String {
        let rot_attr = rot.map(|r| format!(r#" rot="{r}""#)).unwrap_or_default();

        let fill_xml = match (fill_hex, alpha_thousandths) {
            (Some(h), Some(a)) => format!(
                r#"<a:solidFill><a:srgbClr val="{h}"><a:alpha val="{a}"/></a:srgbClr></a:solidFill>"#
            ),
            (Some(h), None) => {
                format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#)
            }
            _ => String::new(),
        };

        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{rot_attr}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}</p:spPr></p:sp>"#
        )
    }

    // ── Shape style tests (US-034) ──────────────────────────────────────

    #[test]
    fn test_shape_rotation() {
        // 90° rotation = 5400000 (60000ths of a degree)
        let shape = make_styled_shape(
            0,
            0,
            2_000_000,
            1_000_000,
            "rect",
            Some("FF0000"),
            Some(5_400_000),
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(s.rotation_deg.is_some(), "Expected rotation_deg to be set");
        assert!(
            (s.rotation_deg.unwrap() - 90.0).abs() < 0.01,
            "Expected 90°, got {}",
            s.rotation_deg.unwrap()
        );
    }

    #[test]
    fn test_shape_transparency() {
        // 50% opacity = alpha val 50000 (in 1000ths of percent)
        let shape = make_styled_shape(
            0,
            0,
            2_000_000,
            1_000_000,
            "rect",
            Some("00FF00"),
            None,
            Some(50_000),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(s.opacity.is_some(), "Expected opacity to be set");
        assert!(
            (s.opacity.unwrap() - 0.5).abs() < 0.01,
            "Expected 0.5 opacity, got {}",
            s.opacity.unwrap()
        );
    }

    #[test]
    fn test_shape_rotation_and_transparency() {
        // 45° rotation (2700000) + 75% opacity (75000)
        let shape = make_styled_shape(
            1_000_000,
            500_000,
            3_000_000,
            2_000_000,
            "ellipse",
            Some("0000FF"),
            Some(2_700_000),
            Some(75_000),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(
            (s.rotation_deg.unwrap() - 45.0).abs() < 0.01,
            "Expected 45°, got {}",
            s.rotation_deg.unwrap()
        );
        assert!(
            (s.opacity.unwrap() - 0.75).abs() < 0.01,
            "Expected 0.75 opacity, got {}",
            s.opacity.unwrap()
        );
        assert!(matches!(s.kind, ShapeKind::Ellipse));
    }

    // ── SmartArt test helpers ───────────────────────────────────────────

    /// Create SmartArt data model XML with the given text items (flat, all depth 0).
    fn make_smartart_data_xml(items: &[&str]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>"#,
        );
        // Root doc node
        xml.push_str(
            r#"<dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
        );
        for (i, item) in items.iter().enumerate() {
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="node"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>{item}</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
                i + 1
            ));
        }
        xml.push_str("</dgm:ptLst>");
        // Connection list: doc→all nodes (flat)
        xml.push_str("<dgm:cxnLst>");
        for (i, _) in items.iter().enumerate() {
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="parOf" srcId="0" destId="{}"/>"#,
                100 + i,
                i + 1,
            ));
        }
        xml.push_str("</dgm:cxnLst>");
        xml.push_str("</dgm:dataModel>");
        xml
    }

    /// Create a SmartArt graphicFrame XML element for a slide.
    fn make_smartart_graphic_frame(x: i64, y: i64, cx: i64, cy: i64, dm_rid: &str) -> String {
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="{dm_rid}" r:lo="rId99" r:qs="rId98" r:cs="rId97"/></a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    /// Build a PPTX with SmartArt diagram data embedded.
    fn build_test_pptx_with_smartart(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xml: &str,
        data_rid: &str,
        data_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        )
        .unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/presentation.xml
        let pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#
        );
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/slides/slide1.xml
        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        // ppt/slides/_rels/slide1.xml.rels — links data_rid to diagram data
        let slide_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{data_rid}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#
        );
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        // ppt/diagrams/data1.xml — the SmartArt data model
        zip.start_file("ppt/diagrams/data1.xml", opts).unwrap();
        zip.write_all(data_xml.as_bytes()).unwrap();

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Extract SmartArt from a FixedElement.
    fn get_smartart(elem: &FixedElement) -> &SmartArt {
        match &elem.kind {
            FixedElementKind::SmartArt(sa) => sa,
            _ => panic!("Expected SmartArt, got {:?}", elem.kind),
        }
    }

    // ── SmartArt integration tests ──────────────────────────────────────

    #[test]
    fn test_slide_with_smartart_produces_items() {
        let smartart_frame = make_smartart_graphic_frame(
            914_400,   // x = 72pt
            1_828_800, // y = 144pt
            5_486_400, // cx = 432pt
            3_086_100, // cy ≈ 243pt
            "rId5",
        );
        let slide_xml = make_slide_xml(&[smartart_frame]);
        let data_xml = make_smartart_data_xml(&["Step 1", "Step 2", "Step 3"]);
        let data = build_test_pptx_with_smartart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &data_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        // Find the SmartArt element
        let sa_elems: Vec<_> = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::SmartArt(_)))
            .collect();
        assert_eq!(sa_elems.len(), 1, "Expected 1 SmartArt element");

        let sa = get_smartart(sa_elems[0]);
        let texts: Vec<&str> = sa.items.iter().map(|n| n.text.as_str()).collect();
        assert_eq!(texts, vec!["Step 1", "Step 2", "Step 3"]);
        // All should be depth 0 (flat)
        assert!(sa.items.iter().all(|n| n.depth == 0));

        // Check position
        assert!((sa_elems[0].x - 72.0).abs() < 0.1);
        assert!((sa_elems[0].y - 144.0).abs() < 0.1);
    }

    #[test]
    fn test_slide_with_smartart_and_text_box() {
        let text_box = make_text_box(100_000, 100_000, 500_000, 200_000, "Title");
        let smartart_frame =
            make_smartart_graphic_frame(500_000, 500_000, 3_000_000, 2_000_000, "rId5");
        let slide_xml = make_slide_xml(&[text_box, smartart_frame]);
        let data_xml = make_smartart_data_xml(&["Item A", "Item B"]);
        let data = build_test_pptx_with_smartart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &data_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        // Should have at least 2 elements: text box + SmartArt
        let sa_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::SmartArt(_)))
            .count();
        let tb_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::TextBox(_)))
            .count();
        assert_eq!(sa_count, 1);
        assert!(tb_count >= 1);

        // Verify SmartArt content
        let sa_elem = page
            .elements
            .iter()
            .find(|e| matches!(e.kind, FixedElementKind::SmartArt(_)))
            .unwrap();
        let sa = get_smartart(sa_elem);
        let texts: Vec<&str> = sa.items.iter().map(|n| n.text.as_str()).collect();
        assert_eq!(texts, vec!["Item A", "Item B"]);
    }

    #[test]
    fn test_slide_without_smartart_no_smartart_elements() {
        let text_box = make_text_box(0, 0, 500_000, 200_000, "No SmartArt");
        let slide_xml = make_slide_xml(&[text_box]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let sa_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::SmartArt(_)))
            .count();
        assert_eq!(sa_count, 0);
    }

    // ── Chart test helpers ────────────────────────────────────────────────

    fn make_chart_graphic_frame(x: i64, y: i64, cx: i64, cy: i64, chart_rid: &str) -> String {
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{chart_rid}"/></a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    fn make_bar_chart_xml(title: &str, categories: &[&str], values: &[f64]) -> String {
        let mut cat_xml = String::new();
        for (i, cat) in categories.iter().enumerate() {
            cat_xml.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{cat}</c:v></c:pt>"#));
        }
        let mut val_xml = String::new();
        for (i, val) in values.iter().enumerate() {
            val_xml.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{val}</c:v></c:pt>"#));
        }
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache>{cat_xml}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache>{val_xml}</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        )
    }

    fn build_test_pptx_with_chart(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slide_xml: &str,
        chart_rid: &str,
        chart_xml: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        let pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{slide_cx_emu}" cy="{slide_cy_emu}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#
        );
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        )
        .unwrap();

        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        let slide_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{chart_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#
        );
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        zip.start_file("ppt/charts/chart1.xml", opts).unwrap();
        zip.write_all(chart_xml.as_bytes()).unwrap();

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    fn get_chart(elem: &FixedElement) -> &Chart {
        match &elem.kind {
            FixedElementKind::Chart(c) => c,
            _ => panic!("Expected Chart, got {:?}", elem.kind),
        }
    }

    // ── Chart integration tests ───────────────────────────────────────────

    #[test]
    fn test_slide_with_chart_produces_chart_element() {
        let chart_frame =
            make_chart_graphic_frame(914_400, 1_828_800, 5_486_400, 3_086_100, "rId5");
        let slide_xml = make_slide_xml(&[chart_frame]);
        let chart_xml =
            make_bar_chart_xml("Sales Data", &["Q1", "Q2", "Q3"], &[100.0, 200.0, 150.0]);
        let data = build_test_pptx_with_chart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &chart_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let chart_elems: Vec<_> = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::Chart(_)))
            .collect();
        assert_eq!(chart_elems.len(), 1, "Expected 1 chart element");

        let chart = get_chart(chart_elems[0]);
        assert_eq!(chart.title.as_deref(), Some("Sales Data"));
        assert_eq!(chart.categories, vec!["Q1", "Q2", "Q3"]);
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].values, vec![100.0, 200.0, 150.0]);

        assert!((chart_elems[0].x - 72.0).abs() < 0.1);
        assert!((chart_elems[0].y - 144.0).abs() < 0.1);
    }

    #[test]
    fn test_slide_with_chart_and_text_box() {
        let text_box = make_text_box(100_000, 100_000, 500_000, 200_000, "Title");
        let chart_frame = make_chart_graphic_frame(500_000, 500_000, 3_000_000, 2_000_000, "rId5");
        let slide_xml = make_slide_xml(&[text_box, chart_frame]);
        let chart_xml = make_bar_chart_xml("Revenue", &["Jan", "Feb"], &[50.0, 75.0]);
        let data = build_test_pptx_with_chart(SLIDE_CX, SLIDE_CY, &slide_xml, "rId5", &chart_xml);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let chart_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::Chart(_)))
            .count();
        let tb_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::TextBox(_)))
            .count();
        assert_eq!(chart_count, 1);
        assert!(tb_count >= 1);
    }

    #[test]
    fn test_slide_without_chart_no_chart_elements() {
        let text_box = make_text_box(0, 0, 500_000, 200_000, "No Chart");
        let slide_xml = make_slide_xml(&[text_box]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let chart_count = page
            .elements
            .iter()
            .filter(|e| matches!(e.kind, FixedElementKind::Chart(_)))
            .count();
        assert_eq!(chart_count, 0);
    }

    #[test]
    fn test_scan_chart_refs_basic() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <p:cSld><p:spTree>
            <p:graphicFrame>
              <p:nvGraphicFramePr>
                <p:cNvPr id="5" name="Chart"/>
                <p:cNvGraphicFramePr/>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="914400" y="1828800"/>
                <a:ext cx="5486400" cy="3086100"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart r:id="rId5"/>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#;

        let refs = scan_chart_refs(slide_xml);
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].x, 914400);
        assert_eq!(refs[0].y, 1828800);
        assert_eq!(refs[0].cx, 5486400);
        assert_eq!(refs[0].cy, 3086100);
        assert_eq!(refs[0].chart_rid, "rId5");
    }

    #[test]
    fn test_scan_chart_refs_no_chart() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
              <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sld>"#;

        let refs = scan_chart_refs(slide_xml);
        assert!(refs.is_empty());
    }

    // ── Gradient background tests (US-050) ──────────────────────────────

    #[test]
    fn test_gradient_background_two_stops() {
        let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></p:bgPr></p:bg>"#;
        let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        // Should have a gradient background
        let gradient = page
            .background_gradient
            .as_ref()
            .expect("Expected gradient background");
        assert_eq!(gradient.stops.len(), 2);

        // First stop: red at 0%
        assert!((gradient.stops[0].position - 0.0).abs() < 0.001);
        assert_eq!(gradient.stops[0].color, Color::new(255, 0, 0));

        // Second stop: blue at 100%
        assert!((gradient.stops[1].position - 1.0).abs() < 0.001);
        assert_eq!(gradient.stops[1].color, Color::new(0, 0, 255));

        // Angle: 5400000 / 60000 = 90 degrees
        assert!((gradient.angle - 90.0).abs() < 0.001);

        // Fallback solid color should be first stop color
        assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_gradient_background_three_stops() {
        let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
        let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let gradient = page
            .background_gradient
            .as_ref()
            .expect("Expected gradient");
        assert_eq!(gradient.stops.len(), 3);
        assert!((gradient.stops[1].position - 0.5).abs() < 0.001);
        assert_eq!(gradient.stops[1].color, Color::new(0, 255, 0));
        assert!((gradient.angle - 0.0).abs() < 0.001);
    }

    #[test]
    fn test_gradient_background_with_scheme_colors() {
        // Use theme colors in gradient stops
        let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></p:bgPr></p:bg>"#;
        let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);

        // Build with theme
        let theme_xml = make_theme_xml(&standard_theme_colors(), "Calibri", "Calibri");
        let data = build_test_pptx_with_theme(SLIDE_CX, SLIDE_CY, &[slide_xml], &theme_xml);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let gradient = page
            .background_gradient
            .as_ref()
            .expect("Expected gradient");
        assert_eq!(gradient.stops.len(), 2);
        // angle = 2700000 / 60000 = 45 degrees
        assert!((gradient.angle - 45.0).abs() < 0.001);
    }

    #[test]
    fn test_solid_background_no_gradient() {
        // Solid fill background should NOT produce a gradient
        let bg_xml =
            r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></p:bgPr></p:bg>"#;
        let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        assert!(
            page.background_gradient.is_none(),
            "Solid fill should not produce gradient"
        );
        assert_eq!(page.background_color, Some(Color::new(255, 204, 0)));
    }

    #[test]
    fn test_gradient_shape_fill() {
        // Shape with gradient fill
        let shape_xml =
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></p:spPr></p:sp>"#
            .to_string();
        let slide_xml = make_slide_xml(&[shape_xml]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        assert_eq!(page.elements.len(), 1);
        let shape = get_shape(&page.elements[0]);

        // Should have gradient fill
        let gf = shape
            .gradient_fill
            .as_ref()
            .expect("Expected gradient fill on shape");
        assert_eq!(gf.stops.len(), 2);
        assert_eq!(gf.stops[0].color, Color::new(255, 0, 0));
        assert_eq!(gf.stops[1].color, Color::new(0, 255, 0));
        assert!((gf.angle - 90.0).abs() < 0.001);

        // Solid fill fallback should be first stop color
        assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_shape_solid_fill_no_gradient() {
        // Shape with only solid fill — gradient_fill should be None
        let shape_xml =
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#
            .to_string();
        let slide_xml = make_slide_xml(&[shape_xml]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let shape = get_shape(&page.elements[0]);
        assert!(
            shape.gradient_fill.is_none(),
            "Solid fill shape should have no gradient"
        );
        assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_gradient_background_no_angle() {
        // Gradient with no <a:lin> element → angle defaults to 0
        let bg_xml = r#"<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill></p:bgPr></p:bg>"#;
        let slide_xml = make_slide_xml_with_bg(bg_xml, &[]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let gradient = page
            .background_gradient
            .as_ref()
            .expect("Expected gradient");
        assert!(
            (gradient.angle - 0.0).abs() < 0.001,
            "Default angle should be 0"
        );
    }

    // ── Shadow / effects tests ─────────────────────────────────────────

    #[test]
    fn test_shape_outer_shadow_parsed() {
        // Shape with <a:effectLst><a:outerShdw> inside <p:spPr>
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
        let slide_xml = make_slide_xml(&[shape_xml]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let shape = get_shape(&page.elements[0]);
        let shadow = shape.shadow.as_ref().expect("Expected shadow");

        // blurRad=50800 EMU → 50800/12700 = 4.0 pt
        assert!(
            (shadow.blur_radius - 4.0).abs() < 0.01,
            "Expected blur_radius ~4.0, got {}",
            shadow.blur_radius
        );
        // dist=38100 EMU → 38100/12700 = 3.0 pt
        assert!(
            (shadow.distance - 3.0).abs() < 0.01,
            "Expected distance ~3.0, got {}",
            shadow.distance
        );
        // dir=2700000 → 2700000/60000 = 45.0 degrees
        assert!(
            (shadow.direction - 45.0).abs() < 0.01,
            "Expected direction ~45.0, got {}",
            shadow.direction
        );
        // color = black
        assert_eq!(shadow.color, Color::new(0, 0, 0));
        // alpha val=50000 → 50000/100000 = 0.5
        assert!(
            (shadow.opacity - 0.5).abs() < 0.01,
            "Expected opacity ~0.5, got {}",
            shadow.opacity
        );
    }

    #[test]
    fn test_shape_no_effects_no_shadow() {
        // Shape with no <a:effectLst> → shadow should be None
        let shape_xml = make_shape(
            100_000,
            200_000,
            500_000,
            300_000,
            "rect",
            Some("00FF00"),
            None,
            None,
        );
        let slide_xml = make_slide_xml(&[shape_xml]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let shape = get_shape(&page.elements[0]);
        assert!(
            shape.shadow.is_none(),
            "Shape without effectLst should have no shadow"
        );
    }

    #[test]
    fn test_shape_shadow_default_opacity() {
        // Shadow with no <a:alpha> element → opacity defaults to 1.0
        let shape_xml = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:outerShdw blurRad="25400" dist="12700" dir="5400000"><a:srgbClr val="333333"/></a:outerShdw></a:effectLst></p:spPr></p:sp>"#.to_string();
        let slide_xml = make_slide_xml(&[shape_xml]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);

        let shape = get_shape(&page.elements[0]);
        let shadow = shape.shadow.as_ref().expect("Expected shadow");

        // blurRad=25400 EMU → 2.0 pt
        assert!(
            (shadow.blur_radius - 2.0).abs() < 0.01,
            "Expected blur ~2.0, got {}",
            shadow.blur_radius
        );
        // dist=12700 EMU → 1.0 pt
        assert!(
            (shadow.distance - 1.0).abs() < 0.01,
            "Expected dist ~1.0, got {}",
            shadow.distance
        );
        // dir=5400000 → 90.0 degrees
        assert!(
            (shadow.direction - 90.0).abs() < 0.01,
            "Expected dir ~90.0, got {}",
            shadow.direction
        );
        // color = #333333
        assert_eq!(shadow.color, Color::new(0x33, 0x33, 0x33));
        // No alpha element → defaults to 1.0
        assert!(
            (shadow.opacity - 1.0).abs() < 0.01,
            "Expected opacity ~1.0 (default), got {}",
            shadow.opacity
        );
    }

    // ── Metadata extraction tests ──────────────────────────────────────

    /// Build a PPTX with metadata in docProps/core.xml.
    fn build_test_pptx_with_metadata(core_xml: &str) -> Vec<u8> {
        let slide = make_slide_xml(&[make_text_box(0, 0, 9144000, 6858000, "Hello")]);
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        ).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/presentation.xml
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
        ).unwrap();

        // ppt/_rels/presentation.xml.rels
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        ).unwrap();

        // ppt/slides/slide1.xml
        zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
        zip.write_all(slide.as_bytes()).unwrap();

        // docProps/core.xml
        zip.start_file("docProps/core.xml", opts).unwrap();
        zip.write_all(core_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_parse_pptx_extracts_metadata() {
        let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My PPTX Title</dc:title>
  <dc:creator>PPTX Author</dc:creator>
  <dc:subject>PPTX Subject</dc:subject>
  <dc:description>PPTX description</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-05-01T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-15T18:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

        let data = build_test_pptx_with_metadata(core_xml);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.metadata.title.as_deref(), Some("My PPTX Title"));
        assert_eq!(doc.metadata.author.as_deref(), Some("PPTX Author"));
        assert_eq!(doc.metadata.subject.as_deref(), Some("PPTX Subject"));
        assert_eq!(
            doc.metadata.description.as_deref(),
            Some("PPTX description")
        );
        assert_eq!(
            doc.metadata.created.as_deref(),
            Some("2024-05-01T09:00:00Z")
        );
        assert_eq!(
            doc.metadata.modified.as_deref(),
            Some("2024-06-15T18:30:00Z")
        );
    }

    #[test]
    fn test_parse_pptx_without_metadata_no_crash() {
        let slide = make_slide_xml(&[make_text_box(0, 0, 9144000, 6858000, "Hello")]);
        let data = build_test_pptx(9144000, 6858000, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        // Should not crash; fields are None or default
        assert!(doc.metadata.title.is_none());
        assert!(doc.metadata.author.is_none());
    }

    // ── Extended geometry tests (US-085) ──────────────────────────────────

    #[test]
    fn test_shape_triangle() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "triangle",
            Some("FF0000"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 3, "Triangle should have 3 vertices");
                // Top-center, bottom-right, bottom-left
                assert!((vertices[0].0 - 0.5).abs() < 0.01);
                assert!((vertices[0].1).abs() < 0.01);
            }
            other => panic!("Expected Polygon for triangle, got {other:?}"),
        }
        assert_eq!(s.fill, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_shape_right_triangle() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "rtTriangle",
            Some("00FF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 3, "Right triangle should have 3 vertices");
                // Top-left, bottom-right, bottom-left
                assert!((vertices[0].0).abs() < 0.01);
                assert!((vertices[0].1).abs() < 0.01);
            }
            other => panic!("Expected Polygon for rtTriangle, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_round_rect() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            1_000_000,
            "roundRect",
            Some("0000FF"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::RoundedRectangle { radius_fraction } => {
                assert!(*radius_fraction > 0.0, "Radius fraction should be positive");
            }
            other => panic!("Expected RoundedRectangle for roundRect, got {other:?}"),
        }
        assert_eq!(s.fill, Some(Color::new(0, 0, 255)));
    }

    #[test]
    fn test_shape_diamond() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "diamond",
            Some("FFFF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 4, "Diamond should have 4 vertices");
            }
            other => panic!("Expected Polygon for diamond, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_pentagon() {
        let shape = make_shape(0, 0, 2_000_000, 2_000_000, "pentagon", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 5, "Pentagon should have 5 vertices");
            }
            other => panic!("Expected Polygon for pentagon, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_hexagon() {
        let shape = make_shape(0, 0, 2_000_000, 2_000_000, "hexagon", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 6, "Hexagon should have 6 vertices");
            }
            other => panic!("Expected Polygon for hexagon, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_octagon() {
        let shape = make_shape(0, 0, 2_000_000, 2_000_000, "octagon", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 8, "Octagon should have 8 vertices");
            }
            other => panic!("Expected Polygon for octagon, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_right_arrow() {
        let shape = make_shape(
            0,
            0,
            3_000_000,
            1_500_000,
            "rightArrow",
            Some("FF8800"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 7, "Arrow should have 7 vertices");
                // Rightmost point should be at x=1.0
                let rightmost = vertices
                    .iter()
                    .map(|v| v.0)
                    .fold(f64::NEG_INFINITY, f64::max);
                assert!((rightmost - 1.0).abs() < 0.01);
            }
            other => panic!("Expected Polygon for rightArrow, got {other:?}"),
        }
        assert_eq!(s.fill, Some(Color::new(255, 136, 0)));
    }

    #[test]
    fn test_shape_left_arrow() {
        let shape = make_shape(0, 0, 3_000_000, 1_500_000, "leftArrow", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 7, "Left arrow should have 7 vertices");
                // Leftmost point should be at x=0.0
                let leftmost = vertices.iter().map(|v| v.0).fold(f64::INFINITY, f64::min);
                assert!(leftmost.abs() < 0.01);
            }
            other => panic!("Expected Polygon for leftArrow, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_up_arrow() {
        let shape = make_shape(0, 0, 1_500_000, 3_000_000, "upArrow", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 7, "Up arrow should have 7 vertices");
            }
            other => panic!("Expected Polygon for upArrow, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_down_arrow() {
        let shape = make_shape(0, 0, 1_500_000, 3_000_000, "downArrow", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(vertices.len(), 7, "Down arrow should have 7 vertices");
            }
            other => panic!("Expected Polygon for downArrow, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_star5() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "star5",
            Some("FFD700"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(
                    vertices.len(),
                    10,
                    "Star5 should have 10 vertices (5 outer + 5 inner)"
                );
            }
            other => panic!("Expected Polygon for star5, got {other:?}"),
        }
        assert_eq!(s.fill, Some(Color::new(255, 215, 0)));
    }

    #[test]
    fn test_shape_star4() {
        let shape = make_shape(0, 0, 2_000_000, 2_000_000, "star4", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(
                    vertices.len(),
                    8,
                    "Star4 should have 8 vertices (4 outer + 4 inner)"
                );
            }
            other => panic!("Expected Polygon for star4, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_star6() {
        let shape = make_shape(0, 0, 2_000_000, 2_000_000, "star6", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Polygon { vertices } => {
                assert_eq!(
                    vertices.len(),
                    12,
                    "Star6 should have 12 vertices (6 outer + 6 inner)"
                );
            }
            other => panic!("Expected Polygon for star6, got {other:?}"),
        }
    }

    #[test]
    fn test_unsupported_preset_falls_back_to_rectangle() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "cloudCallout",
            Some("AABBCC"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(
            matches!(s.kind, ShapeKind::Rectangle),
            "Unknown preset should fall back to Rectangle"
        );
    }
}
