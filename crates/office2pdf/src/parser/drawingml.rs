//! Shared DrawingML color primitives: scheme-color resolution and OOXML
//! color transforms (tint/shade/lumMod/lumOff/alpha).
//!
//! DrawingML color markup (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>` with
//! nested transform children) is identical across pptx, docx, and xlsx parts.
//! This module holds the single implementation; format parsers supply their
//! own theme palette and alias map through [`SchemeColors`].

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::ir::Color;
use crate::parser::xml_util::{get_attr_i64, get_attr_str, parse_hex_color};

/// A format-agnostic view of a theme's color scheme.
///
/// `aliases` carries part-level scheme aliases (the pptx `<p:clrMap>`); pass
/// an empty map when the format has no alias layer (docx, xlsx).
pub(crate) struct SchemeColors<'a> {
    pub(crate) colors: &'a HashMap<String, Color>,
    pub(crate) aliases: &'a HashMap<String, String>,
}

/// Resolve a scheme color name (e.g. `accent1`, `bg1`) against the theme.
///
/// The alias map is consulted first; if the aliased entry is missing, the raw
/// name is tried so a partially populated theme still resolves.
pub(crate) fn resolve_scheme_color(scheme: &SchemeColors<'_>, scheme_name: &str) -> Option<Color> {
    let mapped_name = scheme
        .aliases
        .get(scheme_name)
        .map(String::as_str)
        .unwrap_or(scheme_name);

    scheme
        .colors
        .get(mapped_name)
        .copied()
        .or_else(|| scheme.colors.get(scheme_name).copied())
}

#[derive(Debug, Clone, Copy)]
pub(crate) enum ColorTransform {
    Tint(f64),
    Shade(f64),
    LumMod(f64),
    LumOff(f64),
}

/// A parsed DrawingML color: the resolved RGB value plus an optional alpha.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct ParsedColor {
    pub(crate) color: Option<Color>,
    pub(crate) alpha: Option<f64>,
}

fn parse_base_color(element: &BytesStart<'_>, scheme: &SchemeColors<'_>) -> Option<Color> {
    match element.local_name().as_ref() {
        b"srgbClr" => get_attr_str(element, b"val").and_then(|hex| parse_hex_color(&hex)),
        b"schemeClr" => {
            get_attr_str(element, b"val").and_then(|name| resolve_scheme_color(scheme, &name))
        }
        b"sysClr" => get_attr_str(element, b"lastClr").and_then(|hex| parse_hex_color(&hex)),
        _ => None,
    }
}

pub(crate) fn parse_color_transform(element: &BytesStart<'_>) -> Option<ColorTransform> {
    let val = get_attr_i64(element, b"val")? as f64 / 100_000.0;
    match element.local_name().as_ref() {
        b"tint" => Some(ColorTransform::Tint(val)),
        b"shade" => Some(ColorTransform::Shade(val)),
        b"lumMod" => Some(ColorTransform::LumMod(val)),
        b"lumOff" => Some(ColorTransform::LumOff(val)),
        _ => None,
    }
}

pub(crate) fn apply_color_transforms(color: Color, transforms: &[ColorTransform]) -> Color {
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

/// Parse a self-closing color element (no transform children possible).
pub(crate) fn parse_color_from_empty(
    element: &BytesStart<'_>,
    scheme: &SchemeColors<'_>,
) -> ParsedColor {
    ParsedColor {
        color: parse_base_color(element, scheme),
        alpha: None,
    }
}

/// Parse a color element with children, consuming events through its end tag
/// and applying any nested transforms.
pub(crate) fn parse_color_from_start(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
    scheme: &SchemeColors<'_>,
) -> ParsedColor {
    let base_color = parse_base_color(element, scheme);
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

/// Scheme slot names a `<a:clrScheme>` defines, in document order.
const CLR_SCHEME_SLOTS: &[&str] = &[
    "dk1", "dk2", "lt1", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// Parse just the `<a:clrScheme>` palette out of a theme part
/// (`theme1.xml`) into a scheme-name → color map.
///
/// `srgbClr` uses `val`; `sysClr` uses the application-resolved `lastClr`.
/// The pptx parser keeps its own combined single-pass reader because it also
/// collects fonts and fill styles from the same document.
pub(crate) fn parse_theme_color_scheme(xml: &str) -> HashMap<String, Color> {
    let mut colors: HashMap<String, Color> = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut current_slot: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                if CLR_SCHEME_SLOTS.contains(&name) {
                    current_slot = Some(name.to_string());
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(ref slot) = current_slot {
                    let local = e.local_name();
                    let color = match local.as_ref() {
                        b"srgbClr" => get_attr_str(e, b"val").and_then(|hex| parse_hex_color(&hex)),
                        b"sysClr" => {
                            get_attr_str(e, b"lastClr").and_then(|hex| parse_hex_color(&hex))
                        }
                        _ => None,
                    };
                    if let Some(color) = color {
                        colors.insert(slot.clone(), color);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                if current_slot.as_deref() == Some(name) {
                    current_slot = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    colors
}

#[cfg(test)]
#[path = "drawingml_tests.rs"]
mod tests;
