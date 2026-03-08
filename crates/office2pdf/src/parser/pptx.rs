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
use crate::parser::smartart;

use self::package::{
    load_chart_data, load_slide_images, load_smartart_data, load_theme, parse_presentation_xml,
    parse_rels_xml, read_zip_entry, resolve_layout_master_paths, resolve_relative_path,
    scan_chart_refs,
};

#[path = "pptx_package.rs"]
mod package;

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
#[path = "pptx_tests.rs"]
mod tests;
