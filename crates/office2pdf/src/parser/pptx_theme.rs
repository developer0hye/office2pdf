use super::*;
use crate::ir::{GradientFill, GradientStop};
use crate::parser::units;
use crate::parser::xml_util::OOXML_XML_VERSION;

/// Parsed theme data from ppt/theme/theme1.xml.
#[derive(Debug, Clone, Default)]
pub(super) struct ThemeData {
    /// Color scheme: scheme name (e.g., "dk1", "accent1") → Color.
    pub(super) colors: HashMap<String, Color>,
    /// Major (heading) font family name.
    pub(super) major_font: Option<String>,
    /// Minor (body) font family name.
    pub(super) minor_font: Option<String>,
    /// Raw XML of each `<a:fmtScheme>/<a:fillStyleLst>` entry, for
    /// `<p:bgRef>` idx 1-999 resolution.
    pub(super) fill_styles: Vec<String>,
    /// Raw XML of each `<a:fmtScheme>/<a:bgFillStyleLst>` entry, for
    /// `<p:bgRef>` idx ≥ 1001 resolution.
    pub(super) bg_fill_styles: Vec<String>,
    /// Each `<a:fmtScheme>/<a:lnStyleLst>/<a:ln>` entry, for
    /// `<a:lnRef idx="N">` outline resolution.
    pub(super) line_styles: Vec<ThemeLineStyle>,
    /// Raw XML of each `<a:fmtScheme>/<a:effectStyleLst>/<a:effectStyle>`
    /// entry, for `<a:effectRef idx="N">` resolution.
    pub(super) effect_styles: Vec<String>,
}

/// Effective scheme-color aliases for a slide part.
#[derive(Debug, Clone, Default)]
pub(super) struct ColorMapData {
    pub(super) aliases: HashMap<String, String>,
}

pub(super) use crate::parser::drawingml::ParsedColor;
use crate::parser::drawingml::{self, SchemeColors};

const COLOR_MAP_KEYS: &[&str] = &[
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

pub(super) fn default_color_map() -> ColorMapData {
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

pub(super) fn parse_master_color_map(xml: &str) -> ColorMapData {
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

/// The three text-style buckets of a master's `<p:txStyles>`.
/// Placeholders resolve against them by type: title family → `title`,
/// body/content types → `body`, `dt`/`ftr`/`sldNum` → `other`.
#[derive(Debug, Clone, Default)]
pub(super) struct PptxMasterTextStyles {
    pub(super) title: PptxTextBodyStyleDefaults,
    pub(super) body: PptxTextBodyStyleDefaults,
    pub(super) other: PptxTextBodyStyleDefaults,
}

pub(super) fn parse_master_text_styles(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PptxMasterTextStyles {
    let mut reader = Reader::from_str(xml);
    let mut styles = PptxMasterTextStyles::default();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"titleStyle" => {
                    styles.title = parse_pptx_list_style(&mut reader, theme, color_map);
                }
                b"bodyStyle" => {
                    styles.body = parse_pptx_list_style(&mut reader, theme, color_map);
                }
                b"otherStyle" => {
                    styles.other = parse_pptx_list_style(&mut reader, theme, color_map);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    styles
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

pub(super) fn resolve_effective_color_map(
    xml: &str,
    master_color_map: &ColorMapData,
) -> ColorMapData {
    parse_color_map_override(xml).unwrap_or_else(|| master_color_map.clone())
}

pub(super) fn resolve_scheme_color(
    theme: &ThemeData,
    color_map: &ColorMapData,
    scheme_name: &str,
) -> Option<Color> {
    drawingml::resolve_scheme_color(&scheme_colors(theme, color_map), scheme_name)
}

/// Adapt pptx theme + clrMap to the shared DrawingML color scheme view.
fn scheme_colors<'a>(theme: &'a ThemeData, color_map: &'a ColorMapData) -> SchemeColors<'a> {
    SchemeColors {
        colors: &theme.colors,
        aliases: &color_map.aliases,
    }
}

pub(super) fn parse_color_from_empty(
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> ParsedColor {
    drawingml::parse_color_from_empty(element, &scheme_colors(theme, color_map))
}

pub(super) fn parse_color_from_start(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> ParsedColor {
    drawingml::parse_color_from_start(reader, element, &scheme_colors(theme, color_map))
}

/// Parse a theme XML string to extract the color scheme and font scheme.
pub(super) fn parse_theme_xml(xml: &str) -> ThemeData {
    let mut theme = ThemeData::default();
    let mut reader = Reader::from_str(xml);

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

                if let Some(ref color_name) = current_color_name {
                    if name == "srgbClr"
                        && let Some(hex) = get_attr_str(e, b"val")
                        && let Some(color) = parse_hex_color(&hex)
                    {
                        theme.colors.insert(color_name.clone(), color);
                    } else if name == "sysClr"
                        && let Some(hex) = get_attr_str(e, b"lastClr")
                        && let Some(color) = parse_hex_color(&hex)
                    {
                        theme.colors.insert(color_name.clone(), color);
                    }
                }

                if name == "latin"
                    && let Some(typeface) = get_attr_str(e, b"typeface")
                        .map(|typeface| typeface.trim().to_string())
                        .filter(|typeface| !typeface.is_empty())
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

    theme.fill_styles = extract_style_list_entries(xml, b"fillStyleLst");
    theme.bg_fill_styles = extract_style_list_entries(xml, b"bgFillStyleLst");
    theme.effect_styles = extract_style_list_entries(xml, b"effectStyleLst");
    theme.line_styles = extract_line_styles(xml);

    theme
}

/// One `<a:ln>` of the theme `<a:lnStyleLst>`, as far as `<a:lnRef>` needs it.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub(crate) struct ThemeLineStyle {
    /// The `w` attribute in EMU; 0 when absent.
    pub(crate) width_emu: i64,
    /// The stated corner join, or `None` when the entry names none of
    /// `a:round`, `a:bevel` and `a:miter` (issue #1090).
    pub(crate) join: Option<LineJoin>,
    /// The stated end geometry, or `None` when the entry declares no `cap`.
    ///
    /// PowerPoint's stock themes write `cap="flat"` here, but some state `rnd`,
    /// and a shape taking such an entry through `<a:lnRef idx>` rounds its ends
    /// without naming a cap of its own (issue #1682).
    pub(crate) cap: Option<LineCap>,
}

/// Extract each `<a:ln>` inside the theme `<a:lnStyleLst>`.
///
/// One walk builds width, join and cap together so the three stay index-aligned
/// with the `<a:lnRef idx="N">` that selects an entry.
fn extract_line_styles(xml: &str) -> Vec<ThemeLineStyle> {
    let mut reader = Reader::from_str(xml);
    let mut styles: Vec<ThemeLineStyle> = Vec::new();
    let mut in_list = false;
    let mut depth: usize = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"lnStyleLst" => in_list = true,
                b"ln" if in_list => {
                    depth += 1;
                    if depth == 1 {
                        styles.push(ThemeLineStyle {
                            width_emu: line_width_attr(e),
                            join: None,
                            cap: crate::parser::drawingml::line_cap(e),
                        });
                    }
                }
                name if in_list && depth == 1 => {
                    if let Some(join) = drawingml_line_join(name)
                        && let Some(style) = styles.last_mut()
                    {
                        style.join = Some(join);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) if in_list => match e.local_name().as_ref() {
                b"ln" if depth == 0 => styles.push(ThemeLineStyle {
                    width_emu: line_width_attr(e),
                    join: None,
                    cap: crate::parser::drawingml::line_cap(e),
                }),
                name if depth == 1 => {
                    if let Some(join) = drawingml_line_join(name)
                        && let Some(style) = styles.last_mut()
                    {
                        style.join = Some(join);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if in_list && e.local_name().as_ref() == b"ln" => {
                depth = depth.saturating_sub(1);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lnStyleLst" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    styles
}

/// Parse the theme line matrix for DrawingML shapes outside PowerPoint slides.
pub(crate) fn parse_theme_line_styles(xml: &str) -> Vec<ThemeLineStyle> {
    extract_line_styles(xml)
}

/// Map an `a:ln` child element name to the corner join it selects.
pub(crate) fn drawingml_line_join(local_name: &[u8]) -> Option<LineJoin> {
    match local_name {
        b"round" => Some(LineJoin::Round),
        b"bevel" => Some(LineJoin::Bevel),
        b"miter" => Some(LineJoin::Miter),
        _ => None,
    }
}

fn line_width_attr(e: &BytesStart<'_>) -> i64 {
    e.attributes()
        .flatten()
        .find(|attr| attr.key.local_name().as_ref() == b"w")
        .and_then(|attr| attr.normalized_value(OOXML_XML_VERSION).ok())
        .and_then(|value| value.parse::<i64>().ok())
        .unwrap_or(0)
}

/// Extract the raw XML of each top-level entry inside the named
/// `<a:fmtScheme>` list — a fill (`<a:solidFill>`, `<a:gradFill>`, ...) for the
/// fill lists, an `<a:effectStyle>` for the effect list.
///
/// The entries are kept as XML rather than parsed here so each reference site
/// can run its own parser over them, with `phClr` bound to whatever color the
/// reference carries.
fn extract_style_list_entries(xml: &str, list_tag: &[u8]) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut entries: Vec<String> = Vec::new();
    let mut in_list = false;
    let mut child_depth: usize = 0;
    let mut entry_start: usize = 0;

    loop {
        let position_before: usize = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == list_tag {
                    in_list = true;
                    child_depth = 0;
                } else if in_list {
                    if child_depth == 0 {
                        entry_start = position_before;
                    }
                    child_depth += 1;
                }
            }
            Ok(Event::Empty(_)) => {
                if in_list && child_depth == 0 {
                    let position_after: usize = reader.buffer_position() as usize;
                    entries.push(xml[position_before..position_after].to_string());
                }
            }
            Ok(Event::End(ref e)) => {
                if in_list {
                    if e.local_name().as_ref() == list_tag {
                        return entries;
                    }
                    if child_depth > 0 {
                        child_depth -= 1;
                        if child_depth == 0 {
                            let position_after: usize = reader.buffer_position() as usize;
                            entries.push(xml[entry_start..position_after].to_string());
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    entries
}

/// Parse the image relationship id of a `<p:bg><p:bgPr><a:blipFill>` picture
/// background from a slide/layout/master XML.
pub(super) fn parse_background_image_rid(xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_blip_fill = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"bg" => in_bg = true,
                b"bgPr" if in_bg => in_bg_pr = true,
                b"blipFill" if in_bg_pr => in_blip_fill = true,
                b"blip" if in_blip_fill => {
                    return get_attr_str(e, b"r:embed");
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"bg" => return None,
                b"bgPr" => in_bg_pr = false,
                b"blipFill" => in_blip_fill = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    None
}

/// Resolve a `<p:bg><p:bgRef idx="N">` background reference against the
/// theme's fill style lists (ECMA-376 §19.3.1.2: idx 1-999 → fillStyleLst,
/// idx ≥ 1001 → bgFillStyleLst; the entry's `phClr` takes the bgRef child
/// color). Returns `None` when the XML has no resolvable `bgRef`.
pub(super) fn parse_background_ref(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<(Option<Color>, Option<GradientFill>)> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_ref = false;
    let mut style_index: i64 = 0;
    let mut base_color: Option<Color> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"bg" => in_bg = true,
                b"bgRef" if in_bg => {
                    in_bg_ref = true;
                    style_index = get_attr_i64(e, b"idx").unwrap_or(0);
                }
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_bg_ref => {
                    base_color = parse_color_from_start(&mut reader, e, theme, color_map).color;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"bgRef" if in_bg => {
                    in_bg_ref = true;
                    style_index = get_attr_i64(e, b"idx").unwrap_or(0);
                }
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_bg_ref => {
                    base_color = parse_color_from_empty(e, theme, color_map).color;
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"bgRef" | b"bg" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if !in_bg_ref {
        return None;
    }
    let entry: &str = if style_index >= 1001 {
        theme.bg_fill_styles.get((style_index - 1001) as usize)?
    } else if style_index >= 1 {
        theme.fill_styles.get((style_index - 1) as usize)?
    } else {
        return None;
    };

    resolve_theme_fill_entry(entry, base_color, theme, color_map)
}

/// Resolve a shape's 1-based `<a:fillRef idx>` against the theme's
/// `fillStyleLst`, substituting the reference child color for `phClr`.
pub(super) fn resolve_fill_ref(
    style_index: usize,
    base_color: Option<Color>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<(Option<Color>, Option<GradientFill>)> {
    let entry: &str = theme.fill_styles.get(style_index.checked_sub(1)?)?;
    resolve_theme_fill_entry(entry, base_color, theme, color_map)
}

fn resolve_theme_fill_entry(
    entry: &str,
    base_color: Option<Color>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<(Option<Color>, Option<GradientFill>)> {
    // A style entry uses the same fill grammar as bgPr. Reusing that parser
    // keeps gradients and color transforms identical for bgRef and fillRef.
    let mut theme_with_placeholder: ThemeData = theme.clone();
    if let Some(color) = base_color {
        theme_with_placeholder
            .colors
            .insert("phClr".to_string(), color);
    }
    let synthetic_bg: String = format!("<p:bg><p:bgPr>{entry}<a:effectLst/></p:bgPr></p:bg>");
    let gradient: Option<GradientFill> =
        parse_background_gradient(&synthetic_bg, &theme_with_placeholder, color_map);
    let color: Option<Color> =
        parse_background_color(&synthetic_bg, &theme_with_placeholder, color_map).or_else(|| {
            gradient
                .as_ref()
                .and_then(|g| g.stops.first().map(|s| s.color))
        });
    if color.is_none() && gradient.is_none() {
        return None;
    }
    Some((color, gradient))
}

/// Parse background color from a `<p:bg>` element within a slide/layout/master XML.
pub(super) fn parse_background_color(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<Color> {
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
                    b"bg" => return None,
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
pub(super) fn parse_background_gradient(
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
pub(super) fn parse_shape_gradient_fill(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<GradientFill> {
    let mut in_gs_lst = false;
    let mut in_gs = false;
    let mut current_pos: f64 = 0.0;
    let mut stops: Vec<GradientStop> = Vec::new();
    let mut angle: f64 = 0.0;
    let mut depth: usize = 1;

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
                        // `parse_color_from_start` consumes the matching end tag too.
                        depth = depth.saturating_sub(1);
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

/// Parse a DrawingML `<a:pattFill>` after its start event has been read.
pub(super) fn parse_shape_pattern_fill(
    reader: &mut Reader<&[u8]>,
    preset: PatternPreset,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PatternFill {
    #[derive(Clone, Copy)]
    enum Target {
        Foreground,
        Background,
    }

    let mut target: Option<Target> = None;
    let mut foreground: Option<Color> = None;
    let mut background: Option<Color> = None;
    let mut depth: usize = 1;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                match e.local_name().as_ref() {
                    b"fgClr" => target = Some(Target::Foreground),
                    b"bgClr" => target = Some(Target::Background),
                    b"srgbClr" | b"schemeClr" | b"sysClr" if target.is_some() => {
                        let color = parse_color_from_start(reader, e, theme, color_map).color;
                        match target {
                            Some(Target::Foreground) => foreground = color,
                            Some(Target::Background) => background = color,
                            None => {}
                        }
                        // The color parser consumes the corresponding end tag.
                        depth = depth.saturating_sub(1);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if matches!(
                    e.local_name().as_ref(),
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                ) {
                    let color = parse_color_from_empty(e, theme, color_map).color;
                    match target {
                        Some(Target::Foreground) => foreground = color,
                        Some(Target::Background) => background = color,
                        None => {}
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
                if matches!(e.local_name().as_ref(), b"fgClr" | b"bgClr") {
                    target = None;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    PatternFill {
        preset,
        foreground: foreground.unwrap_or_else(Color::black),
        background: background.unwrap_or_else(Color::white),
    }
}

/// Resolve `<a:effectRef idx="N">` into the theme's effect style list.
///
/// A shape's `<p:style>` can carry its whole appearance by reference, and the
/// effect is the one part that used to be dropped: `a:lnRef` and `a:fillRef`
/// resolved, `a:effectRef` did not, so a shadow that exists only in the theme's
/// `<a:effectStyleLst>` never reached the renderer (issue #740).
///
/// `idx` is 1-based, and `idx="0"` means no effect. The reference's own color
/// child binds `phClr` for the entry, the same way `<p:bgRef>` binds it for a
/// fill entry — stock themes state the shadow color literally, so this matters
/// only for themes that reference the placeholder.
pub(super) fn resolve_effect_ref(
    style_index: i64,
    ref_color: Option<Color>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Option<Shadow> {
    if style_index < 1 {
        return None;
    }
    let entry: &str = theme.effect_styles.get((style_index - 1) as usize)?;

    let mut theme_with_placeholder: ThemeData = theme.clone();
    if let Some(color) = ref_color {
        theme_with_placeholder
            .colors
            .insert("phClr".to_string(), color);
    }

    // Reuse the direct-`a:effectLst` parser by seeking to the entry's own
    // effectLst, so a themed shadow and a literal one take the same path.
    let mut reader = Reader::from_str(entry);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                return parse_effect_list(&mut reader, &theme_with_placeholder, color_map);
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// Resolve the supported top bevel from a theme `<a:effectStyle>`.
///
/// DrawingML stores the front-face bevel in `a:sp3d`, beside (not inside)
/// `a:effectLst`, while its lighting comes from the sibling `a:scene3d`.
/// Keep the first implementation deliberately bounded to the combination
/// PowerPoint writes for `customGeo.pptx` page 44: an orthographic front
/// camera, a top-directed three-point light rig, and the default circular
/// bevel. Unknown cameras, rigs, directions, and presets remain unrendered
/// instead of receiving a misleading rim (issue #1298).
pub(super) fn resolve_effect_ref_top_bevel(
    style_index: i64,
    theme: &ThemeData,
) -> Option<TopBevel> {
    if style_index < 1 {
        return None;
    }
    let entry: &str = theme.effect_styles.get((style_index - 1) as usize)?;
    let mut reader = Reader::from_str(entry);
    let mut in_scene_3d = false;
    let mut in_shape_3d = false;
    let mut orthographic_front = false;
    let mut supported_light_rig = false;
    let mut light_rig_rotation_deg = 0.0;
    let mut bevel: Option<TopBevel> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"scene3d" => in_scene_3d = true,
                b"sp3d" => in_shape_3d = true,
                b"camera" if in_scene_3d => {
                    orthographic_front =
                        get_attr_str(e, b"prst").as_deref() == Some("orthographicFront");
                }
                b"lightRig" if in_scene_3d => {
                    supported_light_rig = get_attr_str(e, b"rig").as_deref() == Some("threePt")
                        && get_attr_str(e, b"dir").as_deref() == Some("t");
                }
                b"rot" if in_scene_3d && supported_light_rig => {
                    light_rig_rotation_deg = get_attr_i64(e, b"rev").unwrap_or(0) as f64 / 60_000.0;
                }
                b"bevelT" if in_shape_3d => {
                    let supported_preset =
                        get_attr_str(e, b"prst").is_none_or(|preset| preset == "circle");
                    if supported_preset {
                        let width = units::emu_to_pt(get_attr_i64(e, b"w").unwrap_or(0));
                        let height = units::emu_to_pt(get_attr_i64(e, b"h").unwrap_or(0));
                        if width > 0.0 && height > 0.0 {
                            bevel = Some(TopBevel {
                                width,
                                height,
                                light_rig_rotation_deg,
                            });
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"scene3d" => in_scene_3d = false,
                b"sp3d" => in_shape_3d = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    (orthographic_front && supported_light_rig)
        .then_some(bevel)
        .flatten()
}

/// Parse `<a:effectLst>` and extract outer shadow if present.
pub(super) fn parse_effect_list(
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
    let mut depth: usize = 1;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                let local = e.local_name();
                match local.as_ref() {
                    b"outerShdw" => {
                        in_outer_shdw = true;
                        shdw_blur = units::emu_to_pt(get_attr_i64(e, b"blurRad").unwrap_or(0));
                        shdw_dist = units::emu_to_pt(get_attr_i64(e, b"dist").unwrap_or(0));
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
                        // `parse_color_from_start` consumes the matching end tag too.
                        depth = depth.saturating_sub(1);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"outerShdw" => {
                        let blur = units::emu_to_pt(get_attr_i64(e, b"blurRad").unwrap_or(0));
                        let dist = units::emu_to_pt(get_attr_i64(e, b"dist").unwrap_or(0));
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
                    break;
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

/// Resolve a font typeface, substituting theme font references.
pub(super) fn resolve_theme_font(typeface: &str, theme: &ThemeData) -> String {
    let typeface = typeface.trim();
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
