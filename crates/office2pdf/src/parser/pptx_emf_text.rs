//! Unicode EMF text follows MS-EMF EmrText/LogFont offsets and GDI alignment.
//! Like docxide-pdf's metafile translator, retain fonts and explicit advances.

use std::fmt::Write;

use super::{Point, RgbColor, read_i32, read_u32};

#[derive(Clone)]
pub(super) struct EmfFont {
    pub(super) height: i32,
    face: String,
    weight: i32,
    italic: bool,
    underline: bool,
    strikeout: bool,
    escapement: i32,
}

#[derive(Clone)]
pub(super) struct EmfTextState {
    pub(super) font: Option<EmfFont>,
    pub(super) align: u32,
    pub(super) color: RgbColor,
    pub(super) background: RgbColor,
    pub(super) is_opaque: bool,
}

impl Default for EmfTextState {
    fn default() -> Self {
        Self {
            font: None,
            align: 0,
            color: RgbColor::from_colorref(0),
            background: RgbColor::from_colorref(0xffffff),
            is_opaque: true,
        }
    }
}

pub(super) struct TextElement {
    pub(super) svg: String,
    pub(super) next_point: Point,
}

pub(super) fn read_font(data: &[u8]) -> Option<(u32, EmfFont)> {
    let face: Vec<u16> = data
        .get(32..96)?
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .take_while(|&c| c != 0)
        .collect();
    Some((
        read_u32(data, 0)?,
        EmfFont {
            height: read_i32(data, 4)?,
            escapement: read_i32(data, 12)?,
            weight: match read_i32(data, 20)? {
                0 => 400,
                weight => weight.clamp(100, 900),
            },
            italic: *data.get(24)? != 0,
            underline: *data.get(25)? != 0,
            strikeout: *data.get(26)? != 0,
            face: String::from_utf16_lossy(&face),
        },
    ))
}

pub(super) fn header_frame(body: &[u8]) -> Option<[f64; 4]> {
    let device = (read_i32(body, 64)?, read_i32(body, 68)?);
    let millimeters = (read_i32(body, 72)?, read_i32(body, 76)?);
    let frame = [
        read_i32(body, 16)?,
        read_i32(body, 20)?,
        read_i32(body, 24)?,
        read_i32(body, 28)?,
    ];
    if device.0 <= 0
        || device.1 <= 0
        || millimeters.0 <= 0
        || millimeters.1 <= 0
        || frame[2] <= frame[0]
        || frame[3] <= frame[1]
    {
        return None;
    }
    let sx = f64::from(device.0) / (f64::from(millimeters.0) * 100.0);
    let sy = f64::from(device.1) / (f64::from(millimeters.1) * 100.0);
    Some([
        f64::from(frame[0]) * sx,
        f64::from(frame[1]) * sy,
        f64::from(frame[2]) * sx,
        f64::from(frame[3]) * sy,
    ])
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

pub(super) fn write_text(
    body: &[u8],
    state: &EmfTextState,
    current: Option<Point>,
) -> Option<Option<TextElement>> {
    let font = state.font.as_ref()?;
    let count = read_u32(body, 36)? as usize;
    let options = read_u32(body, 44)?;
    if count == 0 || font.height == 0 {
        return Some(None);
    }
    if options & 0x10 != 0 {
        // Glyph indices are not Unicode. Preserve other paint without inventing characters.
        tracing::warn!("Unsupported glyph-index EMF text");
        return Some(None);
    }
    let offset = (read_u32(body, 40)? as usize).checked_sub(8)?;
    let bytes = body.get(offset..offset.checked_add(count.checked_mul(2)?)?)?;
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let text = String::from_utf16_lossy(&units);
    let em = if font.height < 0 {
        f64::from(font.height).abs()
    } else {
        let pitch =
            crate::render::pdf::font_line_metrics_em(&font.face).map_or(1.0, |(_, _, pitch)| pitch);
        f64::from(font.height) / pitch.max(0.01)
    };
    let point = if state.align & 1 != 0 {
        current.unwrap_or(Point { x: 0, y: 0 })
    } else {
        Point {
            x: read_i32(body, 28)?,
            y: read_i32(body, 32)?,
        }
    };
    let dx_offset = read_u32(body, 64)? as usize;
    let stride: usize = if options & 0x2000 != 0 { 8 } else { 4 };
    let advances = if dx_offset == 0 {
        None
    } else {
        let start = dx_offset.checked_sub(8)?;
        let array = body.get(start..start.checked_add(count.checked_mul(stride)?)?)?;
        Some(
            array
                .chunks_exact(stride)
                .map(|c| {
                    let x = i32::from_le_bytes(c[0..4].try_into().unwrap());
                    let y = if stride == 8 {
                        i32::from_le_bytes(c[4..8].try_into().unwrap())
                    } else {
                        0
                    };
                    (x, y)
                })
                .collect::<Vec<_>>(),
        )
    };
    let width: f64 = advances
        .as_ref()
        .map_or(0.0, |a| a.iter().map(|(x, _)| f64::from(*x)).sum());
    let shift = match state.align & 6 {
        2 => width,
        6 => width / 2.0,
        _ => 0.0,
    };
    let mut x = f64::from(point.x) - shift;
    let mut y = f64::from(point.y);
    let mut xs = Vec::new();
    let mut ys = Vec::new();
    let mut unit_index = 0;
    for c in text.chars() {
        xs.push(x.to_string());
        ys.push(y.to_string());
        for _ in 0..c.len_utf16() {
            if let Some(a) = advances.as_ref() {
                x += f64::from(a[unit_index].0);
                y += f64::from(a[unit_index].1);
            }
            unit_index += 1;
        }
    }
    let baseline = match state.align & 24 {
        24 => "alphabetic",
        8 => "text-after-edge",
        _ => "text-before-edge",
    };
    let mut svg = String::new();
    if options & 4 != 0 {
        let l = read_i32(body, 48)?;
        let t = read_i32(body, 52)?;
        let width = i64::from(read_i32(body, 56)?) - i64::from(l);
        let height = i64::from(read_i32(body, 60)?) - i64::from(t);
        let _ = write!(
            svg,
            "<svg x=\"{l}\" y=\"{t}\" width=\"{}\" height=\"{}\" viewBox=\"{l} {t} {} {}\" overflow=\"hidden\">",
            width.max(0),
            height.max(0),
            width.max(0),
            height.max(0)
        );
    }
    if options & 2 != 0 {
        let l = read_i32(body, 48)?;
        let t = read_i32(body, 52)?;
        let width = i64::from(read_i32(body, 56)?) - i64::from(l);
        let height = i64::from(read_i32(body, 60)?) - i64::from(t);
        let _ = write!(
            svg,
            "<rect x=\"{l}\" y=\"{t}\" width=\"{}\" height=\"{}\" fill=\"{}\"/>",
            width.max(0),
            height.max(0),
            state.background.as_svg_hex()
        );
    } else if state.is_opaque && advances.is_some() {
        let (ascent, descent, _) =
            crate::render::pdf::font_line_metrics_em(&font.face).unwrap_or((1.0, 0.0, 1.0));
        let top = match state.align & 24 {
            24 => f64::from(point.y) - ascent * em,
            8 => f64::from(point.y) - (ascent + descent) * em,
            _ => f64::from(point.y),
        };
        let _ = write!(
            svg,
            "<rect x=\"{}\" y=\"{top}\" width=\"{}\" height=\"{}\" fill=\"{}\"/>",
            f64::from(point.x) - shift,
            width.abs(),
            (ascent + descent) * em,
            state.background.as_svg_hex()
        );
    }
    let style = if font.italic { "italic" } else { "normal" };
    let decorations = match (font.underline, font.strikeout) {
        (true, true) => "underline line-through",
        (true, false) => "underline",
        (false, true) => "line-through",
        _ => "none",
    };
    let anchor = if advances.is_some() {
        "start"
    } else {
        match state.align & 6 {
            2 => "end",
            6 => "middle",
            _ => "start",
        }
    };
    let x_values = if advances.is_some() {
        xs.join(" ")
    } else {
        point.x.to_string()
    };
    let y_values = if advances.is_some() {
        ys.join(" ")
    } else {
        point.y.to_string()
    };
    let families = crate::render::font_subst::paint_family_candidates_for_text(&font.face, &text)
        .into_iter()
        .map(|family| format!("\"{}\"", family.replace('"', "\\\"")))
        .collect::<Vec<_>>()
        .join(", ");
    let _ = write!(
        svg,
        "<text xml:space=\"preserve\" x=\"{x_values}\" y=\"{y_values}\" font-family=\"{}\" font-size=\"{em}\" font-weight=\"{}\" font-style=\"{style}\" text-decoration=\"{decorations}\" dominant-baseline=\"{baseline}\" text-anchor=\"{anchor}\" fill=\"{}\"",
        escape(&families),
        font.weight,
        state.color.as_svg_hex()
    );
    if font.escapement != 0 {
        let _ = write!(
            svg,
            " transform=\"rotate({} {} {})\"",
            -f64::from(font.escapement) / 10.0,
            point.x,
            point.y
        );
    }
    svg.push('>');
    svg.push_str(&escape(&text));
    svg.push_str("</text>");
    if options & 4 != 0 {
        svg.push_str("</svg>");
    }
    Some(Some(TextElement {
        svg,
        next_point: Point {
            x: x.round() as i32,
            y: y.round() as i32,
        },
    }))
}
