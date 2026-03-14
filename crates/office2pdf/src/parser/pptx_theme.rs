use super::*;
use crate::ir::{GradientFill, GradientStop};
use crate::parser::units;

/// Parsed theme data from ppt/theme/theme1.xml.
#[derive(Debug, Clone, Default)]
pub(super) struct ThemeData {
    /// Color scheme: scheme name (e.g., "dk1", "accent1") → Color.
    pub(super) colors: HashMap<String, Color>,
    /// Major (heading) font family name.
    pub(super) major_font: Option<String>,
    /// Minor (body) font family name.
    pub(super) minor_font: Option<String>,
}

/// Effective scheme-color aliases for a slide part.
#[derive(Debug, Clone, Default)]
pub(super) struct ColorMapData {
    pub(super) aliases: HashMap<String, String>,
}

#[derive(Debug, Clone, Copy, Default)]
pub(super) struct ParsedColor {
    pub(super) color: Option<Color>,
    pub(super) alpha: Option<f64>,
}

#[derive(Debug, Clone, Copy)]
enum ColorTransform {
    Tint(f64),
    Shade(f64),
    LumMod(f64),
    LumOff(f64),
}

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

pub(super) fn parse_master_other_style(
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

pub(super) fn resolve_effective_color_map(
    xml: &str,
    master_color_map: &ColorMapData,
) -> ColorMapData {
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
        b"tint" => Some(ColorTransform::Tint(val)),
        b"shade" => Some(ColorTransform::Shade(val)),
        b"lumMod" => Some(ColorTransform::LumMod(val)),
        b"lumOff" => Some(ColorTransform::LumOff(val)),
        _ => None,
    }
}

fn apply_color_transforms(color: Color, transforms: &[ColorTransform]) -> Color {
    // Apply tint/shade in RGB space first (OOXML spec: blend toward white/black).
    let mut r: f64 = color.r as f64;
    let mut g: f64 = color.g as f64;
    let mut b: f64 = color.b as f64;

    for transform in transforms {
        match transform {
            ColorTransform::Tint(t) => {
                r = 255.0 - (255.0 - r) * t;
                g = 255.0 - (255.0 - g) * t;
                b = 255.0 - (255.0 - b) * t;
            }
            ColorTransform::Shade(s) => {
                r *= s;
                g *= s;
                b *= s;
            }
            _ => {}
        }
    }

    let tinted = Color::new(
        r.round().clamp(0.0, 255.0) as u8,
        g.round().clamp(0.0, 255.0) as u8,
        b.round().clamp(0.0, 255.0) as u8,
    );

    // Then apply luminance transforms in HSL space.
    let has_lum_transforms: bool = transforms
        .iter()
        .any(|t| matches!(t, ColorTransform::LumMod(_) | ColorTransform::LumOff(_)));

    if !has_lum_transforms {
        return tinted;
    }

    let (mut hue, mut saturation, mut lightness) = rgb_to_hsl(tinted);

    for transform in transforms {
        match transform {
            ColorTransform::LumMod(value) => {
                lightness = (lightness * value).clamp(0.0, 1.0);
            }
            ColorTransform::LumOff(value) => {
                lightness = (lightness + value).clamp(0.0, 1.0);
            }
            _ => {}
        }
    }

    saturation = saturation.clamp(0.0, 1.0);
    hue = hue.rem_euclid(360.0);
    hsl_to_rgb(hue, saturation, lightness)
}

pub(super) fn parse_color_from_empty(
    element: &BytesStart<'_>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> ParsedColor {
    ParsedColor {
        color: parse_base_color(element, theme, color_map),
        alpha: None,
    }
}

pub(super) fn parse_color_from_start(
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
