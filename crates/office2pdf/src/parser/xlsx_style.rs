use crate::ir::{BorderLineStyle, BorderSide, CellBorder, Color, TextStyle};
use crate::parser::xml_util::parse_argb_color;

/// Map Excel border style name to width in points.
pub(super) fn border_style_to_width(style: &str) -> Option<f64> {
    match style {
        "hair" => Some(0.25),
        "thin" | "dashed" | "dotted" | "dashDot" | "dashDotDot" => Some(0.5),
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "double"
        | "slantDashDot" => Some(1.0),
        "thick" => Some(2.0),
        _ => None, // "none" or unknown
    }
}

/// Extract font styling from a cell's style into an IR TextStyle.
pub(super) fn extract_cell_text_style(cell: &umya_spreadsheet::Cell) -> TextStyle {
    let style = cell.get_style();
    let Some(font) = style.get_font() else {
        return TextStyle::default();
    };

    let bold = if *font.get_bold() { Some(true) } else { None };
    let italic = if *font.get_italic() { Some(true) } else { None };
    let underline = match font.get_underline() {
        "none" | "" => None,
        _ => Some(true),
    };
    let strikethrough = if *font.get_strikethrough() {
        Some(true)
    } else {
        None
    };

    // Font name: skip default "Calibri" (Excel default) — only set if explicitly customized
    let font_name = font.get_name();
    let font_family = if font_name.is_empty() || font_name == "Calibri" {
        None
    } else {
        Some(font_name.to_string())
    };

    // Font size: skip default 11.0 (Excel default)
    let raw_size = *font.get_size();
    let font_size = if (raw_size - 11.0).abs() < 0.01 {
        None
    } else {
        Some(raw_size)
    };

    // Font color
    let color_argb = font.get_color().get_argb();
    let color = if color_argb.is_empty() || color_argb == "FF000000" {
        // Default black — skip
        None
    } else {
        parse_argb_color(color_argb)
    };

    TextStyle {
        font_family,
        font_size,
        bold,
        italic,
        underline,
        strikethrough,
        color,
        highlight: None,
        vertical_align: None,
        all_caps: None,
        small_caps: None,
        letter_spacing: None,
    }
}

/// Extract background color from a cell's style.
pub(super) fn extract_cell_background(cell: &umya_spreadsheet::Cell) -> Option<Color> {
    let bg = cell.get_style().get_background_color()?;
    parse_argb_color(bg.get_argb())
}

/// Map Excel border style name to `BorderLineStyle`.
pub(super) fn border_style_to_line_style(style: &str) -> BorderLineStyle {
    match style {
        "dashed" | "mediumDashed" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "dashDot" | "mediumDashDot" | "slantDashDot" => BorderLineStyle::DashDot,
        "dashDotDot" | "mediumDashDotDot" => BorderLineStyle::DashDotDot,
        "double" => BorderLineStyle::Double,
        _ => BorderLineStyle::Solid,
    }
}

/// Extract a single border side from an umya Border object.
pub(super) fn extract_border_side(border: &umya_spreadsheet::Border) -> Option<BorderSide> {
    let border_style_str = border.get_border_style();
    let width = border_style_to_width(border_style_str)?;
    let color = parse_argb_color(border.get_color().get_argb()).unwrap_or(Color::black());
    let style = border_style_to_line_style(border_style_str);
    Some(BorderSide {
        width,
        color,
        style,
    })
}

/// Extract cell border properties.
pub(super) fn extract_cell_borders(cell: &umya_spreadsheet::Cell) -> Option<CellBorder> {
    let borders = cell.get_style().get_borders()?;
    let top = extract_border_side(borders.get_top());
    let bottom = extract_border_side(borders.get_bottom());
    let left = extract_border_side(borders.get_left());
    let right = extract_border_side(borders.get_right());
    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }
    Some(CellBorder {
        top,
        bottom,
        left,
        right,
    })
}
