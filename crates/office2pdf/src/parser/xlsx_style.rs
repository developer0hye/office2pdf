use crate::ir::{BorderLineStyle, BorderSide, CellBorder, Color, LineCap, LineJoin, TextStyle};
use crate::parser::xml_util::parse_argb_color;

use super::xlsx_cells::ThemeFontSlot;

/// Resolve a style colour, following `<color theme="N" tint="T"/>` into the
/// workbook's colour scheme.
///
/// A themed colour carries no `rgb`, so `get_argb()` returns the empty string
/// for it and every themed run, fill and border fell back to the renderer's
/// default — a workbook styled entirely from its theme came out black
/// (issue #853). `get_argb_with_theme` indexes umya's colour map, which is
/// `[lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink]` — the ECMA-376 order,
/// including the light/dark swap on slots 0 and 1 — and applies `tint` to the
/// resolved luminance.
///
/// `theme` is `None` when the workbook ships no readable theme part; an
/// explicit `rgb` still resolves, and a themed colour is left unset rather
/// than guessed.
pub(crate) fn resolve_style_color(
    color: &umya_spreadsheet::Color,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<Color> {
    match theme {
        Some(theme) => parse_argb_color(&color.get_argb_with_theme(theme)),
        None => parse_argb_color(color.get_argb()),
    }
}

/// Map an Excel border style name to its printed band width in points.
///
/// Measured directly on a native Excel 16.111 one-factor probe against
/// golden-mock GT traces (issue #619): every border prints as a filled band
/// anchored to the grid boundary — `thin` and `hair` 1pt, `medium` 2pt,
/// `thick` 3pt. The earlier 0.5/1.75/2.5 values (issue #487) were calibrated
/// for Typst strokes centred on the boundary; the direct probe supersedes
/// that indirect scaling.
pub(super) fn border_style_to_width(style: &str) -> Option<f64> {
    match style {
        // `hair` shares `thin`'s 1pt band; only its dotted texture differs.
        "hair" => Some(1.0),
        "thin" | "dashed" | "dotted" | "dashDot" | "dashDotDot" => Some(1.0),
        // A double rule is drawn as two 1pt bands with a 1pt gap between
        // them, so each band carries the `thin` weight rather than the
        // combined one.
        "double" => Some(1.0),
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot" => {
            Some(2.0)
        }
        "thick" => Some(3.0),
        _ => None, // "none" or unknown
    }
}

/// Extract font styling from a cell's style into an IR TextStyle.
///
/// A cell with no style of its own uses `cellXfs[0]`, whose font is the
/// workbook Normal font; umya reports no font for such a cell, so
/// `normal_font` carries it in from `styles.xml` directly. Without it the
/// renderer picks its own default family and size, and a spreadsheet in a
/// sans-serif Normal font rendered serif (issue #462).
pub(super) fn extract_cell_text_style(
    cell: &umya_spreadsheet::Cell,
    normal_font: Option<&super::xlsx_cells::NormalFont>,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> TextStyle {
    let style = cell.get_style();
    let Some(font) = style.get_font() else {
        return normal_font_text_style(normal_font);
    };

    let bold = if *font.get_bold() { Some(true) } else { None };
    let italic = if *font.get_italic() { Some(true) } else { None };
    // Font::get_underline() reads the raw enum value, whose library default is
    // "single" even when the style has no <u> element at all. get_val() checks
    // element presence, so only explicit underlines survive.
    let underline = match font.get_font_underline().get_val() {
        umya_spreadsheet::UnderlineValues::None => None,
        _ => Some(true),
    };
    let strikethrough = if *font.get_strikethrough() {
        Some(true)
    } else {
        None
    };

    // A font the cell style names wins; otherwise the cell inherits the
    // workbook Normal font. Calibri is not special-cased: it is the most
    // common Normal font, and dropping it left those workbooks to the
    // renderer's serif default (issue #462). Either way the face is the one
    // Excel resolves: a `<scheme>` defers it to the theme (issue #1380).
    let font_name: &str = font.get_name();
    let font_family: Option<String> = if font_name.is_empty() {
        normal_font.map(|normal| normal.resolved_family().to_string())
    } else {
        Some(scheme_resolved_family(font, normal_font).unwrap_or_else(|| font_name.to_string()))
    };

    let raw_size: f64 = *font.get_size();
    let font_size: Option<f64> = if raw_size > 0.0 {
        Some(raw_size)
    } else {
        normal_font.map(|normal| normal.size_pt)
    };

    // Font color. An explicit `FF000000` is still skipped so the renderer's
    // own default black stands; a themed colour carries no `rgb` at all and is
    // resolved through the workbook's colour scheme instead (issue #853).
    let color: Option<Color> = if font.get_color().get_argb() == "FF000000" {
        None
    } else {
        resolve_style_color(font.get_color(), theme)
    };

    TextStyle {
        font_family,
        // A workbook states one font per cell; XLSX has no per-script slot.
        east_asian_font_family: None,
        font_size,
        bold,
        italic,
        underline,
        strikethrough,
        color,
        // `xf/font/color` states no opacity, so cell ink is always opaque.
        color_alpha: None,
        highlight: None,
        vertical_align: None,
        baseline_shift: None,
        all_caps: None,
        small_caps: None,
        letter_spacing: None,
        // XLSX has no kerning concept at all, so the renderer's own default
        // stands (issue #628).
        pair_kerning: None,
    }
}

/// The text style a cell inherits when its own style names no font at all.
fn normal_font_text_style(normal_font: Option<&super::xlsx_cells::NormalFont>) -> TextStyle {
    let Some(normal) = normal_font else {
        return TextStyle::default();
    };
    TextStyle {
        font_family: Some(normal.resolved_family().to_string()),
        font_size: Some(normal.size_pt),
        ..TextStyle::default()
    }
}

/// The face a cell or rich-run font that carries its own `<scheme>` paints
/// in: the theme's UI-script face for the scheme slot — its Hang face, then
/// its nonempty East Asian face, under a Korean UI; its Latin face under a
/// Latin UI. `None` for a font naming its face outright (`scheme` absent or
/// `none`) or a theme naming neither face for the slot, both of which stay on
/// the declared name (issues #1380, #1711).
///
/// The theme faces ride on the Normal font because both are read from the
/// same stylesheet and theme parts; a workbook with no readable stylesheet
/// has no cell fonts to resolve either.
fn scheme_resolved_family(
    font: &umya_spreadsheet::Font,
    normal_font: Option<&super::xlsx_cells::NormalFont>,
) -> Option<String> {
    let slot: ThemeFontSlot = ThemeFontSlot::parse(font.get_scheme())?;
    normal_font?
        .theme_ui_script_faces
        .face(slot)
        .map(str::to_string)
}

/// Overlay a rich-text run's own font properties onto the cell-level style.
/// Excel writes only the properties a run changes into its `<rPr>`, so
/// unspecified properties (empty name, zero size, absent color) must keep the
/// cell style rather than reset to defaults.
pub(super) fn apply_rich_run_font(
    base: &TextStyle,
    font: &umya_spreadsheet::Font,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
    normal_font: Option<&super::xlsx_cells::NormalFont>,
) -> TextStyle {
    let mut style = base.clone();

    let font_name: &str = font.get_name();
    if let Some(resolved) = scheme_resolved_family(font, normal_font) {
        style.font_family = Some(resolved);
    } else if !font_name.is_empty() {
        style.font_family = Some(font_name.to_string());
    }

    let raw_size: f64 = *font.get_size();
    if raw_size > 0.0 {
        style.font_size = Some(raw_size);
    }

    if *font.get_bold() {
        style.bold = Some(true);
    }
    if *font.get_italic() {
        style.italic = Some(true);
    }
    if !matches!(
        font.get_font_underline().get_val(),
        umya_spreadsheet::UnderlineValues::None
    ) {
        style.underline = Some(true);
    }
    if *font.get_strikethrough() {
        style.strikethrough = Some(true);
    }

    if let Some(color) = resolve_style_color(font.get_color(), theme) {
        style.color = Some(color);
    }

    style
}

/// The fraction of a cell that a spreadsheetML pattern's foreground covers.
///
/// Excel tiles each `patternType` as an 8x8 hatch of `fgColor` over
/// `bgColor`. Rendering the hatch itself needs a fill value a cell background
/// cannot carry yet (it is a flat [`Color`]), so the two colours are composited
/// at the pattern's coverage — which is what the reference renderer does, and
/// what the #926 measurement pins down.
///
/// Measured on a one-factor probe: one cell per `patternType`, `fgColor`
/// `FF000000` over `bgColor` `FFFFFFFF`, exported by LibreOffice 24.2 and read
/// back as the mean grey of the swatch interior at 300 DPI. Every value came
/// back as an exact 8-bit quantisation of the fraction below (`0.2471` is
/// `1 - 192/255`, `0.4980` is `1 - 128/255`).
///
/// `lightUp`'s quarter is independently confirmed against the #841 Gantt
/// template, where `735773` over an omitted background prints `DCD5DC` — a
/// quarter of the way from white on all three channels, to within one level.
///
/// Two values do not follow from the names: `darkGrid` covers a half, not the
/// three quarters a union of two dark line families would, and `lightTrellis`
/// covers three eighths rather than `lightGrid`'s seven sixteenths.
pub(crate) fn pattern_ink_coverage(pattern: &umya_spreadsheet::PatternValues) -> f64 {
    use umya_spreadsheet::PatternValues::*;
    match pattern {
        Solid => 1.0,
        DarkGray | DarkTrellis => 0.75,
        MediumGray | DarkHorizontal | DarkVertical | DarkDown | DarkUp | DarkGrid => 0.5,
        LightGrid => 0.4375,
        LightTrellis => 0.375,
        LightGray | LightHorizontal | LightVertical | LightDown | LightUp => 0.25,
        Gray125 => 0.125,
        Gray0625 => 0.0625,
        None => 0.0,
    }
}

/// Composite `foreground` over `background` at `coverage`.
pub(crate) fn blend_color(background: Color, foreground: Color, coverage: f64) -> Color {
    let mix = |below: u8, above: u8| -> u8 {
        (f64::from(below) * (1.0 - coverage) + f64::from(above) * coverage).round() as u8
    };
    Color::new(
        mix(background.r, foreground.r),
        mix(background.g, foreground.g),
        mix(background.b, foreground.b),
    )
}

/// Extract background color from a cell's style.
///
/// `umya`'s `Style::get_background_color` hands back the `fgColor` whatever the
/// `patternType` is, so every hatch used to print as a solid foreground —
/// collapsing a legend's `lightUp`, `lightUp`-over-a-tint and `solid` swatches
/// onto one colour (issue #926). Read the pattern fill directly instead.
pub(super) fn extract_cell_background(
    cell: &umya_spreadsheet::Cell,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<Color> {
    extract_style_background(cell.get_style(), theme)
}

/// Resolve a style's pattern fill to the colour it paints. Shared by cell
/// styles and row-level `customFormat` styles, which fill a row's empty
/// cells without any `<c>` element to hang a cell style on (issue #718).
pub(super) fn extract_style_background(
    style: &umya_spreadsheet::Style,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<Color> {
    let pattern = style.get_fill()?.get_pattern_fill()?;
    let pattern_type: &umya_spreadsheet::PatternValues = pattern.get_pattern_type();
    if matches!(pattern_type, umya_spreadsheet::PatternValues::None) {
        return Option::None;
    }
    let foreground: Color =
        resolve_style_color(pattern.get_foreground_color()?, theme).unwrap_or(Color::black());
    let coverage: f64 = pattern_ink_coverage(pattern_type);
    if coverage >= 1.0 {
        return Some(foreground);
    }
    // An omitted `bgColor` is white — the first row of #926's table, where
    // `lightUp` over nothing prints as a quarter-strength foreground rather
    // than as a quarter of the way to black.
    let background: Color = pattern
        .get_background_color()
        .and_then(|color| resolve_style_color(color, theme))
        .unwrap_or(Color::white());
    Some(blend_color(background, foreground, coverage))
}

/// Map Excel border style name to `BorderLineStyle`.
pub(super) fn border_style_to_line_style(style: &str) -> BorderLineStyle {
    match style {
        "dashed" | "mediumDashed" => BorderLineStyle::Dashed,
        // Excel prints `hair` as a dot-textured 1pt band (it tiles an 8x8
        // dot pattern along the rule — issue #619 probe), not a solid line.
        "dotted" | "hair" => BorderLineStyle::Dotted,
        "dashDot" | "mediumDashDot" | "slantDashDot" => BorderLineStyle::DashDot,
        "dashDotDot" | "mediumDashDotDot" => BorderLineStyle::DashDotDot,
        "double" => BorderLineStyle::Double,
        _ => BorderLineStyle::Solid,
    }
}

/// Extract a single border side from an umya Border object.
pub(super) fn extract_border_side(
    border: &umya_spreadsheet::Border,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<BorderSide> {
    let border_style_str = border.get_border_style();
    let width = border_style_to_width(border_style_str)?;
    let color = resolve_style_color(border.get_color(), theme).unwrap_or(Color::black());
    let style = border_style_to_line_style(border_style_str);
    Some(BorderSide {
        width,
        color,
        style,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    })
}

/// Extract cell border properties.
pub(super) fn extract_cell_borders(
    cell: &umya_spreadsheet::Cell,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<CellBorder> {
    let borders = cell.get_style().get_borders()?;
    let top = extract_border_side(borders.get_top(), theme);
    let bottom = extract_border_side(borders.get_bottom(), theme);
    let left = extract_border_side(borders.get_left(), theme);
    let right = extract_border_side(borders.get_right(), theme);
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

/// Extract explicit cell alignment into IR values: (horizontal, vertical).
/// Excel's "general" horizontal default maps to None (renderer default).
pub(super) fn extract_cell_alignment(
    cell: &umya_spreadsheet::Cell,
) -> (
    Option<crate::ir::Alignment>,
    Option<crate::ir::CellVerticalAlign>,
) {
    let Some(alignment) = cell.get_style().get_alignment() else {
        return (None, None);
    };

    use umya_spreadsheet::{HorizontalAlignmentValues as H, VerticalAlignmentValues as V};
    let horizontal = match alignment.get_horizontal() {
        H::Center | H::CenterContinuous => Some(crate::ir::Alignment::Center),
        H::Right => Some(crate::ir::Alignment::Right),
        H::Left => Some(crate::ir::Alignment::Left),
        H::Justify => Some(crate::ir::Alignment::Justify),
        _ => None,
    };
    let vertical = match alignment.get_vertical() {
        V::Center => Some(crate::ir::CellVerticalAlign::Center),
        V::Top => Some(crate::ir::CellVerticalAlign::Top),
        // "bottom" is Excel's default; leave None so the renderer default applies.
        _ => None,
    };
    (horizontal, vertical)
}
