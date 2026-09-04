use std::collections::{HashMap, HashSet};

use crate::ir::{Block, Paragraph, ParagraphStyle, Run, TableRow};
use crate::parser::cond_fmt::build_cond_fmt_overrides;

use super::xlsx_style::{
    apply_rich_run_font, extract_cell_alignment, extract_cell_background, extract_cell_borders,
    extract_cell_text_style, extract_style_background, resolve_style_color,
};
use crate::ir::{BorderSide, CellBorder, Color, Insets, TableCell, TextStyle};

/// The last addressable spreadsheet column (XFD), bounding how far a text
/// overflow may extend the printed range.
const MAX_XLSX_COLUMNS: u32 = 16384;

/// Return a cell's displayed text, preserving whitespace from a literal-only
/// zero section that `umya-spreadsheet` currently trims.
///
/// The workspace patch can select `\-\ \ ` instead of falling back to the
/// positive section, but its width-independent string formatter returns `-`.
/// Excel keeps both escaped spaces; they matter when the text is right-aligned.
/// Keep this narrow compatibility layer until a released dependency carries
/// the complete behavior (issue #1262).
fn formatted_cell_value(cell: &umya_spreadsheet::Cell) -> String {
    if cell.get_value_number().is_some_and(|value| value == 0.0)
        && let Some(number_format) = cell.get_style().get_number_format()
        && let Some(literal) = literal_zero_section_text(number_format.get_format_code())
    {
        return literal;
    }
    cell.get_formatted_value()
}

/// Decode a conventional three- or four-section format's zero section when it
/// contains only quoted or escaped literal text (plus bracketed controls).
/// Value-dependent sections stay on the dependency's normal formatter path.
fn literal_zero_section_text(format: &str) -> Option<String> {
    let mut sections: Vec<&str> = Vec::with_capacity(4);
    let mut section_start: usize = 0;
    let mut in_quotes = false;
    let mut skips_next = false;

    for (index, ch) in format.char_indices() {
        if skips_next {
            skips_next = false;
            continue;
        }
        match ch {
            '\\' | '_' | '*' if !in_quotes => skips_next = true,
            '"' => in_quotes = !in_quotes,
            ';' if !in_quotes => {
                sections.push(&format[section_start..index]);
                section_start = index + ch.len_utf8();
            }
            _ => {}
        }
    }
    if in_quotes || skips_next {
        return None;
    }
    sections.push(&format[section_start..]);
    if !matches!(sections.len(), 3 | 4) {
        return None;
    }
    if sections
        .iter()
        .any(|section| section_has_condition(section))
    {
        return None;
    }

    let mut literal = String::with_capacity(sections[2].len());
    let mut chars = sections[2].chars();
    let mut in_quotes = false;
    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            } else {
                literal.push(ch);
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => literal.push(chars.next()?),
            '[' => {
                let mut control = String::new();
                let mut closed = false;
                for control_char in chars.by_ref() {
                    if control_char == ']' {
                        closed = true;
                        break;
                    }
                    control.push(control_char);
                }
                if !closed || !is_non_rendering_number_format_control(&control) {
                    return None;
                }
            }
            ch if ch.is_whitespace() => literal.push(ch),
            _ => return None,
        }
    }

    (!in_quotes).then_some(literal)
}

fn section_has_condition(section: &str) -> bool {
    let mut chars = section.chars();
    let mut in_quotes = false;
    while let Some(ch) = chars.next() {
        match ch {
            '"' => in_quotes = !in_quotes,
            '\\' | '_' | '*' if !in_quotes => {
                chars.next();
            }
            '[' if !in_quotes => {
                let control: String = chars.by_ref().take_while(|ch| *ch != ']').collect();
                if matches!(control.chars().next(), Some('<' | '>' | '=')) {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

fn is_non_rendering_number_format_control(control: &str) -> bool {
    const COLORS: &[&str] = &[
        "black", "blue", "cyan", "green", "magenta", "red", "white", "yellow",
    ];
    let lowercase = control.to_ascii_lowercase();
    COLORS.contains(&lowercase.as_str())
        || lowercase
            .strip_prefix("color")
            .is_some_and(|index| !index.is_empty() && index.chars().all(|ch| ch.is_ascii_digit()))
        || lowercase
            .strip_prefix("$-")
            .is_some_and(|locale| !locale.is_empty())
}

/// A cell range within a sheet (1-indexed, inclusive).
#[derive(Debug, Clone, Copy)]
pub(crate) struct CellRange {
    pub(crate) start_col: u32,
    pub(crate) start_row: u32,
    pub(crate) end_col: u32,
    pub(crate) end_row: u32,
}

/// A (column, row) coordinate pair (1-indexed).
pub(crate) type CellPos = (u32, u32);

/// Info about a merged cell region, keyed by its top-left coordinate.
pub(super) struct MergeInfo {
    pub(super) col_span: u32,
    pub(super) row_span: u32,
}

/// Convert Excel column width (character units) to points.
/// OOXML widths are expressed relative to the Normal font's column unit. The
/// stored width already incorporates Excel's cell padding adjustment, so
/// print geometry must not add padding again. Excel prints each declared
/// column at an integer point count: probe calibri11frac (issue #621) shows
/// width 10.6 at the 6pt Calibri-11 unit printing 64pt, not 63.6pt.
pub(super) fn column_width_to_pt(char_width: f64, column_unit_pt: f64) -> f64 {
    round_half_up_pt(char_width * column_unit_pt)
}

/// Round to the nearest integer point, halves upward. Excel's column metric
/// rounds half UP, not half-even: the Times New Roman 13 probe lands exactly
/// on 6.500pt and prints a 7pt unit (issue #621). Inputs are non-negative.
fn round_half_up_pt(value: f64) -> f64 {
    (value + 0.5).floor()
}

/// Read the workbook's Normal font (the first `<font>` in `xl/styles.xml`)
/// straight from the archive; umya does not expose the stylesheet. Excel
/// derives all column print metrics from this font, not from cell fonts.
///
/// The first `<font>` is the Normal font as Excel reads it, not the one the
/// `Normal` cell style's `cellStyleXfs` entry points at: repointing that
/// entry at another `fontId`, one factor with a byte-identical re-zip control,
/// left both the probe workbook of issue #1094 and `03_inventory_en.xlsx`
/// exporting identically, while editing the first `<font>` moved every track.
pub(super) fn extract_normal_font(
    data: &[u8],
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> Option<NormalFont> {
    use quick_xml::events::Event;
    use std::io::Read;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(data)).ok()?;
    let theme_declares_script_faces: bool = theme_minor_font_declares_script_faces(&mut archive);
    let mut file = archive.by_name("xl/styles.xml").ok()?;
    let mut xml = String::new();
    file.read_to_string(&mut xml).ok()?;

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut in_first_font = false;
    let mut name: Option<String> = None;
    let mut size: Option<f64> = None;
    let mut font_color: Option<umya_spreadsheet::Color> = None;
    let mut uses_theme_scheme = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                in_first_font = true;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Empty(ref e)) if in_first_font => {
                let val = e
                    .try_get_attribute("val")
                    .ok()
                    .flatten()
                    .and_then(|a| String::from_utf8(a.value.into_owned()).ok());
                match e.local_name().as_ref() {
                    b"name" => name = val,
                    b"sz" => size = val.and_then(|v| v.parse::<f64>().ok()),
                    b"color" => {
                        let mut color = umya_spreadsheet::Color::default();
                        let mut has_value = false;
                        for attribute in e.attributes().flatten() {
                            let Ok(value) = std::str::from_utf8(attribute.value.as_ref()) else {
                                continue;
                            };
                            match attribute.key.local_name().as_ref() {
                                b"rgb" => {
                                    color.set_argb(value);
                                    has_value = true;
                                }
                                b"theme" => {
                                    if let Ok(index) = value.parse::<u32>() {
                                        color.set_theme_index(index);
                                        has_value = true;
                                    }
                                }
                                b"indexed" => {
                                    if let Ok(index) = value.parse::<u32>() {
                                        color.set_indexed(index);
                                        has_value = true;
                                    }
                                }
                                b"tint" => {
                                    if let Ok(tint) = value.parse::<f64>() {
                                        color.set_tint(tint);
                                    }
                                }
                                _ => {}
                            }
                        }
                        if has_value {
                            font_color = Some(color);
                        }
                    }
                    b"scheme" => uses_theme_scheme = true,
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Some(NormalFont {
        family: name?,
        size_pt: size.unwrap_or(11.0),
        color: font_color
            .as_ref()
            .and_then(|color| resolve_style_color(color, theme)),
        uses_theme_scheme,
        theme_declares_script_faces,
    })
}

/// Whether the workbook's theme gives its minor font scheme per-script faces
/// — the `<a:font script="Hang" .../>` list every Office theme carries, and
/// which a theme written by LibreOffice or by hand leaves out entirely.
///
/// Excel resolves a `<scheme>` font's face through that list rather than
/// through the scheme's `<a:latin>` typeface, which is what makes the same
/// declared Calibri 11 lay out against two different faces in two workbooks.
fn theme_minor_font_declares_script_faces(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> bool {
    use quick_xml::events::Event;
    use std::io::Read;

    // Themes are numbered parts; a workbook carries one, and a workbook that
    // carries none has no scheme to resolve through in the first place.
    let Some(part): Option<String> = archive
        .file_names()
        .find(|name| name.starts_with("xl/theme/") && name.ends_with(".xml"))
        .map(str::to_string)
    else {
        return false;
    };
    let mut xml = String::new();
    if archive
        .by_name(&part)
        .ok()
        .and_then(|mut file| file.read_to_string(&mut xml).ok())
        .is_none()
    {
        return false;
    }

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut in_minor_font: bool = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"minorFont" => {
                in_minor_font = true;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"minorFont" => {
                in_minor_font = false;
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if in_minor_font && e.local_name().as_ref() == b"font" =>
            {
                if e.try_get_attribute("script").ok().flatten().is_some() {
                    return true;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    false
}

/// The workbook's Normal font: the `xl/styles.xml` font that cells and
/// header/footer prefixes with no style of their own inherit, and the font
/// Excel derives every column print metric from.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct NormalFont {
    pub(super) family: String,
    pub(super) size_pt: f64,
    pub(super) color: Option<Color>,
    /// Whether the font defers its face to the theme's font scheme
    /// (`<scheme val="minor"/>`), rather than naming it outright.
    ///
    /// Excel then lays rows out against whatever the scheme resolves to,
    /// which is not necessarily `family`: on the reference machine the minor
    /// scheme of the standard Office theme resolves to the locale UI face
    /// (Malgun Gothic), and the same declared Calibri 11 gives a 17pt
    /// default row against a scheme-less 15pt (issue #1047).
    pub(super) uses_theme_scheme: bool,
    /// Whether the workbook's theme gives its minor font scheme per-script
    /// faces. Only bears on a font that `uses_theme_scheme`: that is the
    /// list Excel resolves such a font through (issue #1094).
    pub(super) theme_declares_script_faces: bool,
}

/// Max digit advance of Calibri (and metrically identical Carlito), Excel's
/// default Normal font — the last-resort metric when a family is unknown to
/// the reference table and no real face resolves.
const CALIBRI_DIGIT_ADVANCE_EM: f64 = 0.506836;

/// Reference maximum digit advances (em over U+0030..=U+0039) of the faces
/// Excel itself ships, read from their `hmtx` tables by the issue #621 probe
/// tooling. These outrank live font resolution on purpose: the converting
/// machine may substitute a digit-incompatible face (Calibri → Liberation
/// Sans advances 0.556em against Calibri's 0.5068), which would shift column
/// geometry per machine, while Excel's own print metric always comes from the
/// face Excel resolves. The table also keeps wasm and font-less environments
/// on the exact native-Excel numbers.
pub(super) fn reference_digit_advance_em(family: &str) -> Option<f64> {
    match family.to_ascii_lowercase().as_str() {
        "calibri" | "carlito" => Some(CALIBRI_DIGIT_ADVANCE_EM),
        // Selawik is Microsoft's OFL metric-compatible replacement for Segoe
        // UI. Both advance every decimal digit by 1104/2048em (issue #1472).
        "segoe ui" | "selawik" => Some(0.5390625),
        "arial" | "helvetica" | "liberation sans" => Some(0.556152),
        "verdana" => Some(0.635742),
        "courier new" => Some(0.600098),
        "times new roman" => Some(0.500000),
        "malgun gothic" | "맑은 고딕" => Some(0.550781),
        _ => None,
    }
}

/// The maximum digit advance, in em, of the face `family` names: the
/// reference table first, then the live face, then Excel's default Normal
/// font. Shared by the column metric and the single-line width estimate so
/// both price a family from the same number.
pub(super) fn digit_advance_em(family: &str) -> f64 {
    reference_digit_advance_em(family)
        .or_else(|| crate::render::pdf::max_digit_advance_em(family))
        .unwrap_or(CALIBRI_DIGIT_ADVANCE_EM)
}

/// Points Excel allots to one column character unit for the given Normal
/// font: `round_half_up(max digit advance × size)` — an INTEGER point count.
/// Measured on 17 one-factor native Excel-for-Mac probes (issue #621); the
/// probe set discriminates this model from every integer-96dpi-pixel model
/// (Calibri 10 → 5pt, where pixel-ceiling gave 7px = 5.25pt) and from other
/// rounding modes (Times New Roman 13 = 6.500 → 7 kills half-even; Calibri 9
/// and Verdana 11 kill truncation; Calibri 10 and Verdana 10 kill ceiling).
pub(super) fn column_unit_pt(family: &str, size_pt: f64) -> f64 {
    round_half_up_pt(digit_advance_em(family) * size_pt)
}

/// Points Excel ends a cell's text left of the cell's own right gridline.
///
/// The inset is a quarter of the cell font's whole-point column unit, rounded
/// up. Issue #1232 measured both sides over 48 native Excel-for-Mac probe rows:
/// this right side is one point behind [`cell_left_inset_pt`] at every family
/// and size. [`aligned_cell_padding`] rebalances that asymmetric pair into
/// equal sides for a centred cell while preserving its total width.
pub(super) fn cell_right_inset_pt(family: &str, size_pt: f64) -> f64 {
    (column_unit_pt(family, size_pt) / 4.0).ceil()
}

/// Points Excel starts a cell's text right of the cell's own left gridline.
///
/// The inset is not the constant [`XLSX_CELL_PADDING`] states: it is a whole
/// point that steps with the *cell* font's own column unit, so a title starts
/// further in than the body line under it in the same column.
///
/// One native Excel for Mac export of a one-factor probe (issue #1165): 37
/// rows of one left-aligned `8` in a 65pt column, one (family, size) each,
/// boxed ruler cells printing the column gridlines as border fills so the
/// boundary is drawn rather than inferred. Pen origins from
/// `mutool draw -F trace`, every one a whole point off the ruled boundary:
///
/// | family | size | inset |
/// | --- | --- | ---: |
/// | Calibri | 6, 7, 8 | 2 |
/// | Calibri | 9 - 16 | 3 |
/// | Calibri | 17 - 24 | 4 |
/// | Calibri | 25 - 32 | 5 |
/// | Calibri | 33, 36 | 6 |
/// | Arial | 8 | 2 |
/// | Arial | 10 - 14 | 3 |
/// | Arial | 16 - 20 | 4 |
/// | Arial | 24 | 5 |
/// | Arial | 32 | 6 |
/// | Times New Roman | 11, 16 | 3 |
/// | Times New Roman | 18 | 4 |
/// | Verdana | 10, 11 | 3 |
/// | Century Gothic | 14 | 3 |
/// | Century Gothic | 24 | 5 |
/// | Century Gothic | 30 | 6 |
/// | Segoe UI | 11, 14 | 3 |
///
/// The step is not at the same size in each family — Calibri 16 takes 3 where
/// Arial 16 takes 4 — so the driver is not the point size. It is the
/// whole-point digit advance the column-width model already carries, a
/// quarter of it: `ceil(unit / 4) + 1`. All 37 rows fit exactly, and the
/// Calibri 11 controls that bracket the sweep both read 3.
///
/// The same rule accounts for the corpus residual that filed the issue: the
/// reported workbook's column B steps 3pt from a Segoe UI 14 body line
/// (unit 8, inset 3) to a Century Gothic 30 title (unit 17, inset 6), and its
/// column F steps 2pt from Segoe UI 11 (unit 6, inset 3) to Century Gothic 24
/// (unit 13, inset 5) — the 3.0pt and 2.0pt the native export was measured to
/// carry.
///
/// This probe read the left inset only. Issue #1232's own probe, which read
/// both sides of the same step, has the right inset one point behind it at
/// every size. [`aligned_cell_padding`] rebalances that asymmetric pair for a
/// centred cell, preserving its total width while centring on the column.
pub(super) fn cell_left_inset_pt(family: &str, size_pt: f64) -> f64 {
    cell_right_inset_pt(family, size_pt) + 1.0
}

/// The box of a cell laid out in `style`: the font the cell states, else the
/// workbook Normal font it inherits, else Excel's own Calibri 11 default — the
/// same fallback order the cell's runs resolve through. The vertical sides do
/// not vary with the font; only the horizontal pair does (issues #1165, #1232).
fn styled_cell_padding(style: &TextStyle, normal_font: Option<&NormalFont>) -> Insets {
    let family: &str = style
        .font_family
        .as_deref()
        .or_else(|| normal_font.map(|font| font.family.as_str()))
        .unwrap_or("Calibri");
    let size_pt: f64 = style
        .font_size
        .or_else(|| normal_font.map(|font| font.size_pt))
        .unwrap_or(11.0);
    Insets {
        left: cell_left_inset_pt(family, size_pt),
        right: cell_right_inset_pt(family, size_pt),
        ..XLSX_CELL_PADDING
    }
}

/// The table-level cell box inherited by cells in the workbook Normal font.
pub(super) fn default_cell_padding(normal_font: Option<&NormalFont>) -> Insets {
    styled_cell_padding(&TextStyle::default(), normal_font)
}

/// Space advance of Calibri (and metrically identical Carlito), Excel's
/// default Normal font — the last-resort metric when a family is unknown to
/// the reference table and no real face resolves.
const CALIBRI_SPACE_ADVANCE_EM: f64 = 0.226074;

/// Reference space advances (U+0020, in em) of the faces Excel itself ships,
/// read from their `hmtx` tables. Reference-first for the same reason the
/// digit table above is: the converting machine may substitute a metrically
/// different face, while Excel's own indent comes from the face Excel
/// resolves. The family set is kept in step with that table.
fn reference_space_advance_em(family: &str) -> Option<f64> {
    match family.to_ascii_lowercase().as_str() {
        "calibri" | "carlito" => Some(CALIBRI_SPACE_ADVANCE_EM),
        // The issue #982 workbook's two-level instruction-panel indent is
        // measured before document-scoped fallback faces are materialized.
        // Pin Segoe UI and its metric-compatible Selawik replacement to their
        // shared 561/2048em space so parsing cannot fall back to Calibri's
        // narrower 0.226074em space (issue #1472).
        "segoe ui" | "selawik" => Some(0.27392578125),
        "arial" | "helvetica" | "liberation sans" => Some(0.277832),
        "verdana" => Some(0.351562),
        "courier new" => Some(0.600098),
        "times new roman" => Some(0.250000),
        "malgun gothic" | "맑은 고딕" => Some(0.351562),
        _ => None,
    }
}

/// The space advance, in em, of the face `family` names: the reference table
/// first, then the live face, then Excel's default Normal font.
fn space_advance_em(family: &str) -> f64 {
    reference_space_advance_em(family)
        .or_else(|| crate::render::pdf::space_advance_em(family))
        .unwrap_or(CALIBRI_SPACE_ADVANCE_EM)
}

/// One indent level is three spaces (ECMA-376 §18.8.1, `indent`: "an
/// increment of 1 represents 3 spaces").
const SPACES_PER_INDENT_LEVEL: f64 = 3.0;

/// Points Excel insets an indented cell's text per `<alignment indent="N"/>`
/// level: three spaces of the workbook **Normal** font, each rounded to the
/// whole-point advance grid the column unit already sits on.
///
/// Measured over ten Normal fonts in eleven one-factor native Excel-for-Mac
/// exports, each workbook holding indent-0/indent-N pairs whose only
/// difference is the level (issue #1109):
///
/// | Normal font | space | rounded | unit |
/// | --- | ---: | ---: | ---: |
/// | Calibri 6 | 1.36 | 1 | 3 |
/// | Calibri 8 | 1.81 | 2 | 6 |
/// | Calibri 11 | 2.49 | 2 | 6 |
/// | Calibri 14 | 3.17 | 3 | 9 |
/// | Calibri 16 | 3.62 | 4 | 12 |
/// | Arial 9 | 2.50 | 3 | 9 |
/// | Arial 11 | 3.06 | 3 | 9 |
/// | Arial 20 | 5.56 | 6 | 18 |
/// | Times New Roman 11 | 2.75 | 3 | 9 |
/// | Courier New 11 | 6.60 | 7 | 21 |
///
/// The rounding is what makes this three *whole-point* spaces rather than
/// three fractional ones: Calibri 11 prints 6pt where `3 × 2.49` is 7.46,
/// and Courier New 11 prints 21pt where `3 × 6.60` is 19.8. Cell fonts do
/// not participate — 8pt, 11pt and 22pt cells in one workbook all moved by
/// the same unit, as they do for the column metric (issue #366).
pub(super) fn indent_unit_pt(family: &str, size_pt: f64) -> f64 {
    SPACES_PER_INDENT_LEVEL * round_half_up_pt(space_advance_em(family) * size_pt)
}

/// The indent unit for a sheet whose Normal font is `normal_font`.
///
/// The fallback is Excel's own default Normal font rather than the dominant
/// cell font `resolve_column_unit_pt` falls back to: a workbook with no
/// readable `xl/styles.xml` has no `cellXfs` to carry an indent either, so
/// nothing reaches this with a level to scale.
pub(super) fn resolve_indent_unit_pt(normal_font: Option<&NormalFont>) -> f64 {
    match normal_font {
        Some(font) => indent_unit_pt(&font.family, font.size_pt),
        None => indent_unit_pt("Calibri", 11.0),
    }
}

/// Width in points of a column with no `<col>` entry.
///
/// With no declared `defaultColWidth` either, Excel prints
/// `baseColWidth × unit + 5` points — not 8.43 character units — where
/// `baseColWidth` defaults to 8 when `sheetFormatPr` omits it too. Measured
/// by the issue #621 probes: no-baseColWidth workbooks print 45/53/61pt at
/// unit 5/6/7, and the round-3 probes calibri11base10/calibri11base12
/// (`<sheetFormatPr baseColWidth="10|12"/>`, no defaultColWidth, 6pt
/// Calibri-11 unit) print 65pt and 77pt default columns — killing the
/// ignore-baseColWidth model (53pt). When the sheet does declare
/// `defaultColWidth`, it outranks `baseColWidth` (ECMA-376 §18.3.1.81) and
/// is assumed to quantize like any declared width (the probes only covered
/// the absent case; declared widths quantize this way, so the declared
/// default is routed through the same rule).
pub(super) fn default_column_width_pt(
    declared_width_chars: Option<f64>,
    base_col_width_chars: Option<u32>,
    column_unit_pt: f64,
) -> f64 {
    match declared_width_chars {
        Some(width_chars) => round_half_up_pt(width_chars * column_unit_pt),
        None => f64::from(base_col_width_chars.unwrap_or(8)) * column_unit_pt + 5.0,
    }
}

/// The sheet's `defaultColWidth`, only when the file actually declares one.
/// umya reports 0.0 for an absent attribute, a width Excel never writes.
pub(super) fn declared_default_column_width(sheet: &umya_spreadsheet::Worksheet) -> Option<f64> {
    let width_chars: f64 = *sheet
        .get_sheet_format_properties()
        .get_default_column_width();
    (width_chars > 0.0).then_some(width_chars)
}

/// The sheet's `sheetFormatPr@baseColWidth`, only when the file declares
/// one. umya reports 0 for an absent attribute, a base width Excel never
/// writes.
pub(super) fn declared_base_column_width(sheet: &umya_spreadsheet::Worksheet) -> Option<u32> {
    let base_width_chars: u32 = *sheet.get_sheet_format_properties().get_base_column_width();
    (base_width_chars > 0).then_some(base_width_chars)
}

/// Fallback when `xl/styles.xml` is unreadable: infer the column unit from
/// the dominant cell font. umya resolves each cell's effective style while
/// reading, so the dominant family is a stable approximation. The Normal
/// size is unknown too, so Excel's default of 11pt is assumed.
pub(super) fn sheet_column_unit_pt(sheet: &umya_spreadsheet::Worksheet) -> f64 {
    let mut family_counts: HashMap<String, usize> = HashMap::new();
    for cell in sheet.get_cell_collection() {
        let Some(font) = cell.get_style().get_font() else {
            continue;
        };
        let family = font.get_name().trim();
        if !family.is_empty() {
            *family_counts
                .entry(family.to_ascii_lowercase())
                .or_default() += 1;
        }
    }

    let dominant_family: Option<String> = family_counts
        .into_iter()
        .max_by(|(family_a, count_a), (family_b, count_b)| {
            count_a.cmp(count_b).then_with(|| family_b.cmp(family_a))
        })
        .map(|(family, _)| family);

    match dominant_family {
        Some(family) => column_unit_pt(&family, 11.0),
        // No fonts at all: keep the legacy 7px × 0.75 = 5.25pt UNIT (issue
        // #716). Only the unit survives from the old model — the widths built
        // on it still change under #621: default columns move from 44.2575pt
        // (8.43 × 5.25) to 8 × 5.25 + 5 = 47pt, and declared widths now
        // quantize to integer points. Those surrounding changes are
        // extrapolated from the probed model, not measured: the #621 probes
        // never covered a workbook without a readable stylesheet.
        None => 5.25,
    }
}

/// Parse an Excel column letter string (e.g., "A", "B", "AA") into a 1-indexed column number.
pub(super) fn parse_column_letters(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }
    let mut col: u32 = 0;
    for c in s.chars() {
        if !c.is_ascii_uppercase() {
            return None;
        }
        col = col * 26 + (c as u32 - b'A' as u32 + 1);
    }
    Some(col)
}

/// Parse a cell reference like "$A$1", "A1", "$B$10" into (col, row), both 1-indexed.
pub(crate) fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    // Strip dollar signs
    let s = s.replace('$', "");
    // Split into letter part and number part
    let split_pos = s.find(|c: char| c.is_ascii_digit())?;
    let col_str = &s[..split_pos];
    let row_str = &s[split_pos..];
    let col = parse_column_letters(col_str)?;
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
}

/// Parse a print area address string (e.g., "Sheet1!$A$1:$C$10") into a CellRange.
pub(super) fn parse_print_area_range(address: &str) -> Option<CellRange> {
    // Strip optional sheet prefix (everything up to and including '!')
    let range_part = if let Some(pos) = address.rfind('!') {
        &address[pos + 1..]
    } else {
        address
    };

    let (start_str, end_str) = range_part.split_once(':')?;
    let (start_col, start_row) = parse_cell_ref(start_str)?;
    let (end_col, end_row) = parse_cell_ref(end_str)?;
    Some(CellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    })
}

/// Look up the print area for a given sheet from its defined names.
pub(super) fn find_print_area(sheet: &umya_spreadsheet::Worksheet) -> Option<CellRange> {
    for dn in sheet.get_defined_names() {
        if dn.get_name() == "_xlnm.Print_Area" {
            let addr = dn.get_address();
            if let Some(range) = parse_print_area_range(&addr) {
                return Some(range);
            }
        }
    }
    None
}

/// Print-title ranges from `_xlnm.Print_Titles`: rows and/or columns that
/// Excel repeats on every printed page (1-indexed, inclusive).
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct PrintTitles {
    pub(super) rows: Option<(u32, u32)>,
    pub(super) cols: Option<(u32, u32)>,
}

/// Look up the sheet's print titles. The defined name holds one or two
/// comma-separated parts like `Sheet4!$A:$B,Sheet4!$2:$3`. Sheet-scoped
/// names (localSheetId) land on the worksheet; names the reader could not
/// scope stay at the workbook level, so both are consulted.
pub(super) fn find_print_titles(
    book: &umya_spreadsheet::Spreadsheet,
    sheet: &umya_spreadsheet::Worksheet,
) -> PrintTitles {
    let mut titles = PrintTitles::default();
    for dn in sheet.get_defined_names() {
        if dn.get_name() == "_xlnm.Print_Titles" {
            parse_print_title_address(&dn.get_address(), &mut titles);
        }
    }
    if titles.rows.is_none() && titles.cols.is_none() {
        let plain_prefix: String = format!("{}!", sheet.get_name());
        let quoted_prefix: String = format!("'{}'!", sheet.get_name());
        for dn in book.get_defined_names() {
            let address: String = dn.get_address();
            if dn.get_name() == "_xlnm.Print_Titles"
                && (address.contains(&plain_prefix) || address.contains(&quoted_prefix))
            {
                parse_print_title_address(&address, &mut titles);
            }
        }
    }
    titles
}

fn parse_print_title_address(address: &str, titles: &mut PrintTitles) {
    for part in address.split(',') {
        let range_part: String = part
            .rsplit('!')
            .next()
            .unwrap_or(part)
            .replace('$', "")
            .trim()
            .to_string();
        let Some((start_str, end_str)) = range_part.split_once(':') else {
            continue;
        };
        if let (Ok(row_start), Ok(row_end)) = (start_str.parse::<u32>(), end_str.parse::<u32>()) {
            titles.rows = Some((row_start.min(row_end), row_start.max(row_end)));
        } else if let (Some(col_start), Some(col_end)) = (
            parse_column_letters(start_str),
            parse_column_letters(end_str),
        ) {
            titles.cols = Some((col_start.min(col_end), col_start.max(col_end)));
        }
    }
}

/// Collect sorted manual row page break positions from a sheet.
pub(super) fn collect_row_breaks(sheet: &umya_spreadsheet::Worksheet) -> Vec<u32> {
    let mut breaks: Vec<u32> = sheet
        .get_row_breaks()
        .get_break_list()
        .iter()
        .filter(|b| *b.get_manual_page_break())
        .map(|b| *b.get_id())
        .collect();
    breaks.sort_unstable();
    breaks.dedup();
    breaks
}

/// Build a lookup of merge info from the sheet's merged cell ranges.
///
/// Returns two structures:
/// - `top_left_map`: top-left coordinate → MergeInfo for each merge
/// - `skip_set`: set of coordinates that are inside a merge but NOT the top-left
pub(super) fn build_merge_maps(
    sheet: &umya_spreadsheet::Worksheet,
) -> (HashMap<CellPos, MergeInfo>, HashSet<CellPos>) {
    let mut top_left_map: HashMap<CellPos, MergeInfo> = HashMap::new();
    let mut skip_set: HashSet<CellPos> = HashSet::new();

    for range in sheet.get_merge_cells() {
        let start_col = range
            .get_coordinate_start_col()
            .map(|c| *c.get_num())
            .unwrap_or(1);
        let start_row = range
            .get_coordinate_start_row()
            .map(|r| *r.get_num())
            .unwrap_or(1);
        let end_col = range
            .get_coordinate_end_col()
            .map(|c| *c.get_num())
            .unwrap_or(start_col);
        let end_row = range
            .get_coordinate_end_row()
            .map(|r| *r.get_num())
            .unwrap_or(start_row);

        let col_span = end_col.saturating_sub(start_col) + 1;
        let row_span = end_row.saturating_sub(start_row) + 1;

        top_left_map.insert((start_col, start_row), MergeInfo { col_span, row_span });

        // Mark all other cells in the range as skip
        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if r != start_row || c != start_col {
                    skip_set.insert((c, r));
                }
            }
        }
    }

    (top_left_map, skip_set)
}

/// Shared context for processing a single XLSX sheet.
pub(super) struct SheetContext {
    pub(super) col_start: u32,
    pub(super) col_end: u32,
    pub(super) num_cols: usize,
    pub(super) column_widths: Vec<f64>,
    /// Printed width of a column with no `<col>` entry, honouring a declared
    /// `defaultColWidth` (issue #621).
    pub(super) default_column_width_pt: f64,
    pub(super) merge_tops: HashMap<(u32, u32), MergeInfo>,
    pub(super) merge_skips: HashSet<(u32, u32)>,
    pub(super) cond_fmt_overrides: HashMap<(u32, u32), crate::parser::cond_fmt::CondFmtOverride>,
    /// Render-ready x14 line sparklines keyed by their destination cell.
    pub(super) sparklines: HashMap<(u32, u32), crate::ir::SparklineInfo>,
    /// The workbook Normal font, which every cell without its own font
    /// inherits (issue #462). `None` when `styles.xml` is unreadable.
    pub(super) normal_font: Option<NormalFont>,
    /// The paint the sheet's tables take from their built-in style: banded-row
    /// shading (issue #532), header and foot rules, and a bold header row
    /// (issue #1080).
    pub(super) table_styles: Vec<crate::parser::xlsx::tables::TableStyleRange>,
    /// The workbook's colour scheme, which `<color theme="N"/>` indexes into
    /// (issue #853). Cloned rather than borrowed so the context stays free of
    /// the workbook's lifetime; it is twelve colours and a font scheme.
    pub(super) theme: Option<umya_spreadsheet::structs::drawing::Theme>,
    /// The alignment indent level of every cell that declares one, read from
    /// the raw package because umya drops the attribute (issue #1109).
    pub(super) cell_indents: super::indent::CellIndentLevels,
    /// Fixed printed points reserved by automatic-row `thickTop` and
    /// `thickBot` flags, read from the raw package because the crates.io v2
    /// umya release drops `thickTop` (issue #1228).
    pub(super) row_boundary_points: super::row_boundaries::RowBoundaryPoints,
    /// Points one indent level insets a cell's text by, from the workbook
    /// Normal font.
    pub(super) indent_unit_pt: f64,
}

impl SheetContext {
    /// Points the cell at `(col, row)` insets its text by for its alignment
    /// indent level: nothing for the cells that declare none.
    pub(super) fn cell_indent_pt(&self, col: u32, row: u32) -> f64 {
        f64::from(self.cell_indents.get(&(col, row)).copied().unwrap_or(0)) * self.indent_unit_pt
    }
}

/// The insets a cell is laid out with before any indent, for the way its text
/// aligns horizontally.
///
/// `box_insets` is Excel's text box for this cell's own font, and a left- or
/// right-aligned run sits against one of its two edges, so it takes the box as
/// it stands. A centred run does not centre in that box: Excel centres it on
/// the column itself. Over the ten business mocks the 570 centred runs match a
/// symmetric split of the same total to a mean of +0.012pt and the asymmetric
/// box to +0.512pt — exactly the half point the asymmetry would move them by
/// (issue #1157).
///
/// Splitting the total rather than dropping the inset keeps the width the cell
/// has for wrapping unchanged, so only the centre moves.
fn aligned_cell_padding(box_insets: Insets, alignment: Option<crate::ir::Alignment>) -> Insets {
    match alignment {
        Some(crate::ir::Alignment::Center) => {
            let half: f64 = (box_insets.left + box_insets.right) / 2.0;
            Insets {
                left: half,
                right: half,
                ..box_insets
            }
        }
        _ => box_insets,
    }
}

/// The extra inset an indented cell's text takes, on the side its alignment
/// anchors to.
///
/// Excel takes the indent off the left for a left- or general-aligned text
/// cell and off the right for a right-aligned one — a general-aligned number
/// right-aligns, so its indent comes off the right too (probe-measured for
/// issue #1109).
///
/// A centred cell is left alone: its native behaviour splits on whether the
/// text wraps — an unwrapped centred line moves a whole unit left per level
/// while a wrapped one stays centred — and Excel's own UI switches alignment
/// to left rather than let the two combine. Nothing in the reported workbook
/// carries a centred indent to measure the split on.
fn indented_cell_padding(
    base: Insets,
    indent_pt: f64,
    alignment: Option<crate::ir::Alignment>,
) -> Insets {
    if indent_pt <= 0.0 {
        return base;
    }
    match alignment {
        Some(crate::ir::Alignment::Center) => base,
        Some(crate::ir::Alignment::Right) => Insets {
            right: base.right + indent_pt,
            ..base
        },
        _ => Insets {
            left: base.left + indent_pt,
            ..base
        },
    }
}

/// First strong bidi direction of a character: Some(true) for right-to-left
/// scripts (Hebrew, Arabic and its supplements), Some(false) for Latin-like
/// letters, None for neutral characters (digits, punctuation, spaces).
fn strong_direction(c: char) -> Option<bool> {
    match c as u32 {
        // Hebrew, Arabic, Syriac, Thaana, and Arabic presentation forms.
        0x0590..=0x08FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF => Some(true),
        _ if c.is_alphabetic() => Some(false),
        _ => None,
    }
}

/// Map ASCII digits (and separators) to Arabic-Indic digits, as Excel does
/// for number formats carrying a native-digit locale prefix like
/// `[$-3000401]`.
fn to_arabic_indic_digits(value: &str) -> String {
    value
        .chars()
        .map(|c| match c {
            '0'..='9' => char::from_u32(0x0660 + (c as u32 - '0' as u32)).unwrap_or(c),
            '.' => '\u{066B}',
            ',' => '\u{066C}',
            _ => c,
        })
        .collect()
}

/// Excel number formats may carry a locale prefix `[$-XXXXXXXX]` whose high
/// byte selects digit shaping (>= 2 substitutes national digits). Arabic
/// primary language (low byte 0x01) then prints Arabic-Indic digits.
fn uses_native_arabic_digits(format_code: &str) -> bool {
    let Some(rest) = format_code.strip_prefix("[$-") else {
        return false;
    };
    let Some(end) = rest.find(']') else {
        return false;
    };
    let Ok(locale) = u64::from_str_radix(&rest[..end], 16) else {
        return false;
    };
    let digit_substitution: u64 = locale >> 24;
    let language_id: u64 = locale & 0xFF;
    digit_substitution >= 2 && language_id == 0x01
}

/// Advance of each printable ASCII character, in multiples of its face's own
/// maximum digit advance.
///
/// A proportional face prices its glyphs nearly proportionally to that digit
/// advance — the same number every column width already scales by, see
/// [`reference_digit_advance_em`] — so one table plus that per-family number
/// reproduces a line's real `hmtx` advance sum closely. The entries are the
/// mean over the `hmtx` tables of Arial, Calibri, Verdana and Tahoma; measured
/// against those four faces' own sums over realistic cell strings the mean
/// error is 0.9–1.6%, and 3.3% on Times New Roman.
///
/// It replaced a flat half-em per ASCII character, which cost a realistic
/// sentence a third more than its glyphs really advance (issue #1054) and
/// priced `iiiiiiiiii` like `WWWWWWWWWW`. Fixed-pitch families are the
/// remaining miss: Courier New advances every glyph by its digit advance, so
/// proportional ratios misprice its lines by 13.7% on average and up to 23%
/// — where the flat rule was a uniform 8.4% under-count.
///
/// The table is static rather than read from the resolved face for the reason
/// [`reference_digit_advance_em`] gives: a machine substituting a metrically
/// different face must not move the printed range, and wasm resolves no face
/// at all.
const ASCII_ADVANCE_RATIO: [f64; 95] = [
    0.5178, // U+0020 space
    0.5924, // U+0021 '!'
    0.7216, // U+0022 '"'
    1.1507, // U+0023 '#'
    1.0000, // U+0024 '$'
    1.6227, // U+0025 '%'
    1.2306, // U+0026 '&'
    0.3969, // U+0027 '\''
    0.6531, // U+0028 '('
    0.6531, // U+0029 ')'
    0.9206, // U+002A '*'
    1.1632, // U+002B '+'
    0.5297, // U+002C ','
    0.6456, // U+002D '-'
    0.5311, // U+002E '.'
    0.6691, // U+002F '/'
    1.0000, // U+0030 '0'
    1.0000, // U+0031 '1'
    1.0000, // U+0032 '2'
    1.0000, // U+0033 '3'
    1.0000, // U+0034 '4'
    1.0000, // U+0035 '5'
    1.0000, // U+0036 '6'
    1.0000, // U+0037 '7'
    1.0000, // U+0038 '8'
    1.0000, // U+0039 '9'
    0.5973, // U+003A ':'
    0.5973, // U+003B ';'
    1.1632, // U+003C '<'
    1.1632, // U+003D '='
    1.1632, // U+003E '>'
    0.9099, // U+003F '?'
    1.7069, // U+0040 '@'
    1.1286, // U+0041 'A'
    1.1076, // U+0042 'B'
    1.1373, // U+0043 'C'
    1.2417, // U+0044 'D'
    1.0463, // U+0045 'E'
    0.9660, // U+0046 'F'
    1.2714, // U+0047 'G'
    1.2367, // U+0048 'H'
    0.5855, // U+0049 'I'
    0.7515, // U+004A 'J'
    1.0978, // U+004B 'K'
    0.9041, // U+004C 'L'
    1.4805, // U+004D 'M'
    1.2429, // U+004E 'N'
    1.3098, // U+004F 'O'
    1.0442, // U+0050 'P'
    1.3151, // U+0051 'Q'
    1.1501, // U+0052 'R'
    1.0504, // U+0053 'S'
    1.0247, // U+0054 'T'
    1.2292, // U+0055 'U'
    1.1218, // U+0056 'V'
    1.6649, // U+0057 'W'
    1.0911, // U+0058 'X'
    1.0460, // U+0059 'Y'
    1.0310, // U+005A 'Z'
    0.6300, // U+005B '['
    0.6691, // U+005C '\\'
    0.6300, // U+005D ']'
    1.1116, // U+005E '^'
    0.9957, // U+005F '_'
    0.7932, // U+0060 '`'
    0.9628, // U+0061 'a'
    1.0073, // U+0062 'b'
    0.8495, // U+0063 'c'
    1.0073, // U+0064 'd'
    0.9707, // U+0065 'e'
    0.5595, // U+0066 'f'
    0.9803, // U+0067 'g'
    1.0134, // U+0068 'h'
    0.4256, // U+0069 'i'
    0.4823, // U+006A 'j'
    0.9098, // U+006B 'k'
    0.4256, // U+006C 'l'
    1.5356, // U+006D 'm'
    1.0134, // U+006E 'n'
    0.9974, // U+006F 'o'
    1.0073, // U+0070 'p'
    1.0073, // U+0071 'q'
    0.6545, // U+0072 'r'
    0.8269, // U+0073 's'
    0.5982, // U+0074 't'
    1.0134, // U+0075 'u'
    0.9083, // U+0076 'v'
    1.3389, // U+0077 'w'
    0.8979, // U+0078 'x'
    0.9088, // U+0079 'y'
    0.8297, // U+007A 'z'
    0.7749, // U+007B '{'
    0.6975, // U+007C '|'
    0.7749, // U+007D '}'
    1.1632, // U+007E '~'
];

/// Single-line text width estimate in points, summed over the runs' own
/// families and sizes.
fn estimate_text_width_pt(runs: &[Run]) -> f64 {
    runs.iter()
        .map(|run| {
            estimate_line_width_pt(
                &run.text,
                run.style.font_family.as_deref(),
                run.style.font_size.unwrap_or(11.0),
            )
        })
        .sum()
}

/// The single-run form of [`estimate_text_width_pt`], for callers that have a
/// bare string, family and font size rather than IR runs.
///
/// Printable ASCII costs its [`ASCII_ADVANCE_RATIO`] share of the family's
/// digit advance; CJK and other non-Latin glyphs are priced full-width, and
/// ASCII control characters draw nothing. A cell naming no family is measured
/// on Excel's default Normal font, the same last resort [`column_unit_pt`]
/// takes.
pub(super) fn estimate_line_width_pt(text: &str, family: Option<&str>, font_size: f64) -> f64 {
    let digit_advance_em: f64 = family.map_or(CALIBRI_DIGIT_ADVANCE_EM, digit_advance_em);
    text.chars()
        .map(|c| match c {
            ' '..='~' => {
                ASCII_ADVANCE_RATIO[c as usize - ' ' as usize] * digit_advance_em * font_size
            }
            _ if c.is_ascii() => 0.0,
            _ => 1.05 * font_size,
        })
        .sum::<f64>()
}

/// The width an unwrapped cell's single line may paint across, or `None` when
/// the text fits its own column and needs no special handling.
///
/// `wrapText="false"` means exactly that in Excel: the text never moves to a
/// second line. What varies is only how far it may paint before being clipped —
/// a general/left cell paints on across consecutive empty neighbours to its
/// right, and a cell with nowhere to go is clipped at its own edge. Probed
/// against Excel 16.0: a centred cell whose text runs well past its column,
/// with occupied cells on both sides, prints one clipped line; it does not
/// wrap.
///
/// Restricting this to left alignment made every overflowing centred or
/// right-aligned cell fall through to wrapping, which grew the row and, once
/// rows take the height Excel recorded, overflowed it (issue #615).
#[allow(clippy::too_many_arguments)]
fn compute_spill_width(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    col_idx: u32,
    row_idx: u32,
    runs: &[Run],
    cell_alignment: Option<crate::ir::Alignment>,
    col_span: u32,
    umya_cell: Option<&umya_spreadsheet::Cell>,
    cell_padding: Insets,
) -> Option<f64> {
    if runs.is_empty() {
        return None;
    }
    // Explicit wrapText wraps inside the cell instead.
    let has_wrap_text: bool = umya_cell
        .and_then(|cell| cell.get_style().get_alignment().cloned())
        .map(|alignment| *alignment.get_wrap_text())
        .unwrap_or(false);
    if has_wrap_text {
        return None;
    }
    // Embedded line breaks always wrap.
    if runs.iter().any(|run| run.text.contains('\n')) {
        return None;
    }

    // A merged cell never paints past the merge edge: Excel keeps unwrapped
    // text on one line and clips it at the merged width. Apply this even when
    // the text fits — column pagination may clamp the merge to fewer columns
    // on a page, and the line is still not wrapped there.
    //
    // Two caveats this width alone does not carry. The line is clipped at the
    // page-column edge and its remainder is redrawn on the next page-column;
    // we blank that continuation instead (#631). And the renderer lays this
    // width out as a wrapping box rather than a one-line clip, so the fragment
    // left visible is the tail of the text, not its head (#811).
    if col_span > 1 {
        let merged_width: f64 = (col_idx..col_idx + col_span)
            .map(|c| {
                ctx.column_widths
                    .get((c - ctx.col_start) as usize)
                    .copied()
                    .unwrap_or(0.0)
            })
            .sum();
        return Some(merged_width);
    }

    let own_width: f64 = *ctx.column_widths.get((col_idx - ctx.col_start) as usize)?;
    // Leave room for the horizontal cell inset. Passed in rather than written
    // out, so the threshold cannot drift from the padding the cell is actually
    // laid out with (it did when the sides moved to 3pt for issue #657), and
    // so a large-font cell prices the wider inset its own font takes
    // (issue #1165). An alignment indent comes out of the same width, so a
    // line that fits flush may still reach its column edge indented
    // (issue #1109).
    let horizontal_inset: f64 =
        cell_padding.left + cell_padding.right + ctx.cell_indent_pt(col_idx, row_idx);
    if estimate_text_width_pt(runs) <= own_width - horizontal_inset {
        return None;
    }

    // Only a general/left cell paints on into what lies to its right. A centred
    // or right-aligned one is clipped at its own edge — but still on one line.
    let spills_right: bool = matches!(cell_alignment, None | Some(crate::ir::Alignment::Left));
    if !spills_right {
        return Some(own_width);
    }

    let mut total_width: f64 = own_width;
    let mut has_empty_neighbor = false;
    let mut blocked = false;
    for neighbor_col in (col_idx + 1)..=ctx.col_end {
        // Merged regions block the spill like occupied cells do.
        if ctx.merge_skips.contains(&(neighbor_col, row_idx))
            || ctx.merge_tops.contains_key(&(neighbor_col, row_idx))
        {
            blocked = true;
            break;
        }
        let neighbor_is_empty: bool = sheet
            .get_cell((neighbor_col, row_idx))
            .map(|cell| formatted_cell_value(cell).is_empty())
            .unwrap_or(true);
        if !neighbor_is_empty {
            blocked = true;
            break;
        }
        total_width += *ctx
            .column_widths
            .get((neighbor_col - ctx.col_start) as usize)
            .unwrap_or(&0.0);
        has_empty_neighbor = true;
    }

    // Every used cell to the right is empty: Excel keeps painting across
    // the virtual empty cells beyond the used range toward the page edge.
    // Give the text the width it needs; the page boundary clips the rest.
    if !blocked {
        let needed_width: f64 = estimate_text_width_pt(runs) + 4.0;
        if needed_width > total_width {
            total_width = needed_width;
            has_empty_neighbor = true;
        }
    }

    // Nowhere to spill: the line is clipped at the cell's own edge rather than
    // wrapped onto a second line, which is what Excel prints.
    if !has_empty_neighbor {
        return Some(own_width);
    }
    Some(total_width)
}

/// The stored-value fallback when the sheet declares no `defaultRowHeight`.
const EXCEL_DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;

/// The height Excel recomputes for a row that records none, when the model
/// for this Normal font is measured. `None` leaves the caller on the
/// declared `defaultRowHeight`.
///
/// Excel does not honour that declared hint for such rows unless the sheet
/// marks it `customHeight`; it lays them out from the workbook's Normal font
/// instead. Probe-measured against native Excel-for-Mac exports of a
/// 20-row workbook whose rows carry no `ht` (issue #1047): declaring 9, 15,
/// 20 or 30 exports byte-identical 17pt rows, and only adding
/// `customHeight="1"` brings the declared value back (30 → 30pt rows). The
/// same probe read the heights back through AppleScript, which reports 15.0
/// for a declared 30 — the hint never reaches the row.
///
/// Keyed on size within a face, not on `family`: with
/// `<scheme val="minor"/>` Excel resolves the Normal font through the
/// theme's font scheme, which on the reference machine renders even ASCII
/// text in the locale UI face. Both an 11pt "Calibri" and a 10pt "Arial"
/// scheme font exported as Malgun Gothic there, at 17pt and 15pt rows —
/// while removing `scheme` from the same package, changing nothing else,
/// dropped the 11pt rows to 15pt. So the scheme fonts share one table, while
/// scheme-less Calibri and Aptos each use their own measured series; a size
/// none of those tables measures keeps the declared hint.
///
/// | scheme font size | recomputed row | printed track |
/// | ---: | ---: | ---: |
/// | 8 | 13 | 13 |
/// | 9 | 14 | 14 |
/// | 10 | 15 | 15 |
/// | 11 | 17 | 17 |
/// | 12 | 18 | 18 |
/// | 13 | 19 | 19 |
/// | 14 | 20 | 20 |
/// | 15 | 22 | 22 |
/// | 16 | 23 | 23 |
/// | 17 | 26 | 26 |
/// | 18 | 27 | 27 |
/// | 20 | 30 | 30 |
/// | 22 | 32 | 32 |
/// | 24 | 35 | 35 |
pub(super) fn recomputed_default_row_height_pt(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font: Option<&NormalFont>,
) -> Option<f64> {
    if *sheet.get_sheet_format_properties().get_custom_height() {
        return None;
    }
    let font: &NormalFont = normal_font?;
    measured_row_height_pt(font, font.size_pt)
}

/// The row this workbook's recompute gives a `size_pt` line, or `None` where
/// its Normal font's face has no measured series or the size sits between
/// the measured points.
///
/// The series belongs to the face, so a row sized by a font of its own is
/// looked up in the workbook's table at that font's size. On the workbook
/// this was measured against that is exactly right, because its cells name
/// the very face its theme scheme resolves to (issue #1140). A row whose
/// cells name some *other* face has no series of its own here.
fn measured_row_height_pt(font: &NormalFont, size_pt: f64) -> Option<f64> {
    let measured: &[(f64, f64)] = if font.uses_theme_scheme {
        &UI_SCRIPT_FACE_ROW_HEIGHTS
    } else if font.family.eq_ignore_ascii_case("Calibri") {
        &CALIBRI_RECOMPUTED_ROW_HEIGHTS
    } else if font.family.eq_ignore_ascii_case("Aptos") {
        &APTOS_RECOMPUTED_ROW_HEIGHTS
    } else {
        named_face_row_heights(&font.family)?
    };
    measured
        .iter()
        .find(|(measured_size_pt, _)| (size_pt - measured_size_pt).abs() < 0.01)
        .map(|(_, height_pt)| *height_pt)
}

/// The series measured for a Normal font that names `family` outright, or
/// `None` where no sweep has covered it.
///
/// Matched on the whole name, case-insensitively — never as a prefix. The
/// Calibri and Aptos recompute tables above follow the same rule; prefix
/// matching belongs only to the separately measured printed-grid remap below.
/// `Arial Narrow`, `Arial Black` and `Arial Unicode MS` are separate faces
/// with row heights of their own, and lending them Arial's would be a guess
/// wearing a measurement's clothes.
fn named_face_row_heights(family: &str) -> Option<&'static [(f64, f64)]> {
    NAMED_FACE_ROW_HEIGHTS
        .iter()
        .find(|face| {
            face.families
                .iter()
                .any(|candidate| candidate.eq_ignore_ascii_case(family))
        })
        .map(|face| face.heights)
}

/// The row a dimension-less row takes under a Normal font that resolves
/// through the theme's per-script face list — Malgun Gothic on the reference
/// machine, whose measurement the doc comment above records.
///
/// The original scheme sweep covered six points. Issue #1150 swept the same
/// resolved face, Malgun Gothic, across fourteen sizes: its six overlapping
/// readings agree exactly, which identifies one series, and issue #1226 adds
/// the remaining eight. The native customer-workbook exports captured before
/// that change exercise its 12pt entry over 101 and 1001 dimension-less rows.
const UI_SCRIPT_FACE_ROW_HEIGHTS: [(f64, f64); 14] = [
    (8.0, 13.0),
    (9.0, 14.0),
    (10.0, 15.0),
    (11.0, 17.0),
    (12.0, 18.0),
    (13.0, 19.0),
    (14.0, 20.0),
    (15.0, 22.0),
    (16.0, 23.0),
    (17.0, 26.0),
    (18.0, 27.0),
    (20.0, 30.0),
    (22.0, 32.0),
    (24.0, 35.0),
];

/// The recompute under a Normal font that names Calibri outright, a face the
/// reference machine substitutes.
///
/// Swept one size per export on `issue_1066_blip_effect_picture.xlsx`, which
/// declares `defaultRowHeight="18"`, reading `standard height` and every
/// row's `height` back through AppleScript. Not one size answered 18, so the
/// hint reaches none of them; re-running the sweep with the hint rewritten to
/// 30 returned the same column of numbers.
///
/// Issue #1225 extended the sweep to fourteen sizes and repeated both the
/// Calibri and Aptos columns. The no-patch re-zip control kept the same
/// reading, and both runs reproduced this column exactly.
///
/// | size | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 20 | 22 | 24 |
/// | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
/// | row | 11 | 12 | 14 | 15 | 16 | 17 | 19 | 20 | 21 | 23 | 24 | 26 | 29 | 31 |
///
/// This is the *worksheet* height, not the printed track: unlike the scheme
/// faces above, these compact on the way to the PDF, so 12pt Calibri prints a
/// `round(16 x 0.92) = 15pt` row. Callers map it through
/// `native_excel_pdf_row_height`.
///
/// Every other family the sweep reached has a column of its own, none of them
/// following this one; `NAMED_FACE_ROW_HEIGHTS` carries them (issue #1150).
const CALIBRI_RECOMPUTED_ROW_HEIGHTS: [(f64, f64); 14] = [
    (8.0, 11.0),
    (9.0, 12.0),
    (10.0, 14.0),
    (11.0, 15.0),
    (12.0, 16.0),
    (13.0, 17.0),
    (14.0, 19.0),
    (15.0, 20.0),
    (16.0, 21.0),
    (17.0, 23.0),
    (18.0, 24.0),
    (20.0, 26.0),
    (22.0, 29.0),
    (24.0, 31.0),
];

/// The corresponding Aptos recompute. One shared 11pt reading in issue #1102
/// originally paired it with Calibri, but the full issue #1225 sweep separates
/// the faces at 9, 13, 16, 20 and 24pt. This exact table therefore stays
/// independent of the broader `REMAPPED_NORMAL_FAMILIES` printed-grid rule.
///
/// | size | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 20 | 22 | 24 |
/// | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
/// | row | 11 | 13 | 14 | 15 | 16 | 18 | 19 | 20 | 22 | 23 | 24 | 27 | 29 | 32 |
const APTOS_RECOMPUTED_ROW_HEIGHTS: [(f64, f64); 14] = [
    (8.0, 11.0),
    (9.0, 13.0),
    (10.0, 14.0),
    (11.0, 15.0),
    (12.0, 16.0),
    (13.0, 18.0),
    (14.0, 19.0),
    (15.0, 20.0),
    (16.0, 22.0),
    (17.0, 23.0),
    (18.0, 24.0),
    (20.0, 27.0),
    (22.0, 29.0),
    (24.0, 32.0),
];

/// One face's measured recompute, and the Normal-font family names that
/// select it.
struct NamedFaceRowHeights {
    /// Matched whole and case-insensitively by `named_face_row_heights`.
    families: &'static [&'static str],
    /// `(Normal font size, recomputed worksheet row height)`, both in points.
    heights: &'static [(f64, f64)],
}

/// What Excel recomputes for a dimension-less row under a Normal font that
/// names its face outright, for every family swept in issue #1150.
///
/// Same method as the two tables above and the same base package
/// (`issue_1066_blip_effect_picture.xlsx`, which declares
/// `defaultRowHeight="18"`): one `xl/styles.xml` `<font>` per variant with
/// nothing else touched, reading `standard height of worksheet 1` back
/// through AppleScript. `scripts/measure_excel_row_height.py` drives it, and
/// re-running the whole matrix reproduced every column exactly.
///
/// The declaration reaches none of these faces, which is the same control
/// #1047 and #1102 ran and not an inference from theirs: rebuilding the base
/// with `defaultRowHeight="30"` and re-sweeping Arial, Courier New, Segoe UI
/// and `나눔명조` returned all four columns unchanged. Individual readings do
/// land on 18 — four of the six series below read exactly that at 14pt — so a
/// column agreeing with the declaration at one size says nothing either way.
///
/// | family | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 20 | 22 | 24 |
/// | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
/// | Arial, Times New Roman, Verdana | 11 | 12 | 13 | 14 | 16 | 17 | 18 | 19 | 20 | 22 | 23 | 25 | 28 | 30 |
/// | Tahoma | 11 | 12 | 13 | 14 | 15 | 17 | 18 | 19 | 20 | 22 | 23 | 25 | 28 | 30 |
/// | Georgia | 11 | 12 | 13 | 14 | 16 | 17 | 18 | 19 | 21 | 22 | 23 | 25 | 28 | 30 |
/// | Helvetica, 나눔명조 | 11 | 12 | 13 | 15 | 16 | 17 | 18 | 19 | 21 | 22 | 23 | 26 | 28 | 31 |
/// | Courier New | 11 | 13 | 14 | 15 | 17 | 18 | 19 | 21 | 22 | 23 | 24 | 27 | 30 | 32 |
/// | Segoe UI | 11 | 13 | 14 | 16 | 16 | 20 | 21 | 23 | 23 | 25 | 26 | 28 | 31 | 33 |
///
/// Segoe UI really does answer the same row at 11 and 12, and again at 15 and
/// 16; both sweeps read it that way.
///
/// Families share a series here only where their columns came back identical
/// at all fourteen sizes, which is a measurement and not a claim that they
/// resolve to one face. **Do not extrapolate between the sizes listed, and do
/// not derive a missing family from its metrics**: read out of the installed
/// `hhea` tables, Arial and Times New Roman do share a line of 2355 units per
/// 2048 em, but Verdana's 2489 answers their column all the same, Georgia's
/// 2327 sits *under* Arial's and prints a taller row at 16, and Courier New's
/// 2320 is the shortest line of those six faces and gives the tallest column
/// of the six. Whatever Excel measures here, it is not the face's line box.
///
/// `맑은 고딕` and `Malgun Gothic` name the face the theme's per-script list
/// resolves to on this machine, so a font naming it outright takes the
/// scheme's own `UI_SCRIPT_FACE_ROW_HEIGHTS`. Its six original scheme points
/// agree with the named-face sweep exactly, which identifies the series, and
/// issue #1226 folds in that sweep's other eight sizes after capturing native
/// Excel exports of the two customer workbooks that exercise 12pt. Their 18pt
/// row pitch now agrees; the remaining theme-face/column, final-column text
/// overflow and pristine-sheet paper differences are tracked in #1380, #1381
/// and #1382 respectively.
///
/// A spelling is only aliased once it has been swept as its own variant.
/// `NanumMyeongjo` is not `나눔명조` here: it answers a column of its own that
/// falls at 13pt (16 at 12, 15 at 13) and at 18pt, reproducibly, so it is left
/// to the declared hint rather than lent the Korean spelling's series.
///
/// A size a family's series skips keeps the declared hint: nothing here
/// interpolates.
const NAMED_FACE_ROW_HEIGHTS: [NamedFaceRowHeights; 7] = [
    NamedFaceRowHeights {
        families: &["Arial", "Times New Roman", "Verdana"],
        heights: &[
            (8.0, 11.0),
            (9.0, 12.0),
            (10.0, 13.0),
            (11.0, 14.0),
            (12.0, 16.0),
            (13.0, 17.0),
            (14.0, 18.0),
            (15.0, 19.0),
            (16.0, 20.0),
            (17.0, 22.0),
            (18.0, 23.0),
            (20.0, 25.0),
            (22.0, 28.0),
            (24.0, 30.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["Tahoma"],
        heights: &[
            (8.0, 11.0),
            (9.0, 12.0),
            (10.0, 13.0),
            (11.0, 14.0),
            (12.0, 15.0),
            (13.0, 17.0),
            (14.0, 18.0),
            (15.0, 19.0),
            (16.0, 20.0),
            (17.0, 22.0),
            (18.0, 23.0),
            (20.0, 25.0),
            (22.0, 28.0),
            (24.0, 30.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["Georgia"],
        heights: &[
            (8.0, 11.0),
            (9.0, 12.0),
            (10.0, 13.0),
            (11.0, 14.0),
            (12.0, 16.0),
            (13.0, 17.0),
            (14.0, 18.0),
            (15.0, 19.0),
            (16.0, 21.0),
            (17.0, 22.0),
            (18.0, 23.0),
            (20.0, 25.0),
            (22.0, 28.0),
            (24.0, 30.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["Helvetica", "나눔명조"],
        heights: &[
            (8.0, 11.0),
            (9.0, 12.0),
            (10.0, 13.0),
            (11.0, 15.0),
            (12.0, 16.0),
            (13.0, 17.0),
            (14.0, 18.0),
            (15.0, 19.0),
            (16.0, 21.0),
            (17.0, 22.0),
            (18.0, 23.0),
            (20.0, 26.0),
            (22.0, 28.0),
            (24.0, 31.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["Courier New"],
        heights: &[
            (8.0, 11.0),
            (9.0, 13.0),
            (10.0, 14.0),
            (11.0, 15.0),
            (12.0, 17.0),
            (13.0, 18.0),
            (14.0, 19.0),
            (15.0, 21.0),
            (16.0, 22.0),
            (17.0, 23.0),
            (18.0, 24.0),
            (20.0, 27.0),
            (22.0, 30.0),
            (24.0, 32.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["Segoe UI"],
        heights: &[
            (8.0, 11.0),
            (9.0, 13.0),
            (10.0, 14.0),
            (11.0, 16.0),
            (12.0, 16.0),
            (13.0, 20.0),
            (14.0, 21.0),
            (15.0, 23.0),
            (16.0, 23.0),
            (17.0, 25.0),
            (18.0, 26.0),
            (20.0, 28.0),
            (22.0, 31.0),
            (24.0, 33.0),
        ],
    },
    NamedFaceRowHeights {
        families: &["맑은 고딕", "Malgun Gothic"],
        heights: &UI_SCRIPT_FACE_ROW_HEIGHTS,
    },
];

/// Whether this Normal font names one of the families Excel substitutes a
/// face for on the reference machine.
fn names_a_substituted_family(font: &NormalFont) -> bool {
    let family: String = font.family.to_ascii_lowercase();
    REMAPPED_NORMAL_FAMILIES
        .iter()
        .any(|remapped| family.starts_with(remapped))
}

/// The sheet's `defaultRowHeight`, or Excel's stored default when it
/// declares none.
pub(super) fn declared_default_row_height_pt(sheet: &umya_spreadsheet::Worksheet) -> f64 {
    let declared: f64 = *sheet.get_sheet_format_properties().get_default_row_height();
    if declared > 0.0 {
        declared
    } else {
        EXCEL_DEFAULT_ROW_HEIGHT_PT
    }
}

/// Families the standard Office theme resolves to, and the ones the
/// reference machine has no face of its own for. This printed-grid rule is
/// matched as a prefix so the whole Aptos family — `Aptos`, `Aptos Narrow`,
/// `Aptos Display` — and `Calibri Light` come along. The recomputed worksheet
/// row tables above are narrower exact-face measurements.
const REMAPPED_NORMAL_FAMILIES: [&str; 2] = ["calibri", "aptos"];

/// Below this the remap leaves the printed grid alone: the same declared
/// `ht=40` exports a 40pt track under Calibri 9 and Calibri 10, and a 37pt
/// one from 11pt up.
const REMAPPED_NORMAL_MIN_SIZE_PT: f64 = 11.0;

/// The measured whole-point tracks Excel's printed grid gives an Arial 12
/// worksheet row. Inputs not listed here are deliberately not interpolated.
///
/// Issue #1224 first measured both the 16pt dimension-less row and a 36pt
/// fixed row. A follow-up native PDF sweep then used three equal fixed rows
/// per package and read their two baseline pitches. Its no-patch re-zip
/// control was layout-identical. The coarse sweep covered the eleven selected
/// points listed below; a 0.5pt sweep supplied every point from 10 through 24
/// and showed that the output is a staircase, not `round(height * factor)`.
const ARIAL_12_PRINTED_GRID_ROW_HEIGHTS: [(f64, f64); 34] = [
    (10.0, 9.0),
    (10.5, 9.0),
    (11.0, 10.0),
    (11.5, 10.0),
    (12.0, 11.0),
    (12.5, 11.0),
    (13.0, 12.0),
    (13.5, 12.0),
    (14.0, 13.0),
    (14.5, 13.0),
    (15.0, 14.0),
    (15.5, 14.0),
    (16.0, 15.0),
    (16.5, 15.0),
    (17.0, 16.0),
    (17.5, 16.0),
    (18.0, 17.0),
    (18.5, 17.0),
    (19.0, 17.0),
    (19.5, 18.0),
    (20.0, 18.0),
    (20.5, 19.0),
    (21.0, 19.0),
    (21.5, 20.0),
    (22.0, 20.0),
    (22.5, 21.0),
    (23.0, 21.0),
    (23.5, 22.0),
    (24.0, 22.0),
    (25.5, 24.0),
    (30.0, 28.0),
    (36.0, 33.0),
    (40.0, 37.0),
    (49.5, 46.0),
];

/// The corresponding measured points for Courier New 12. Its staircase is
/// different from Arial's: for example, the same 36pt row prints at 31pt
/// here and 33pt in Arial. That is why this is a face table rather than a
/// broader "compacts" boolean or a shared scale.
const COURIER_NEW_12_PRINTED_GRID_ROW_HEIGHTS: [(f64, f64); 11] = [
    (12.0, 10.0),
    (15.0, 13.0),
    (16.0, 14.0),
    (17.0, 15.0),
    (18.0, 16.0),
    (20.0, 17.0),
    (25.5, 22.0),
    (30.0, 26.0),
    (36.0, 31.0),
    (40.0, 35.0),
    (49.5, 43.0),
];

/// A measured printed track for a Normal font that names its face outright.
///
/// The issue #1224 sweep also measured Verdana 12 and Malgun Gothic 12 over
/// the same eleven heights. Both keep the worksheet height whole, so they
/// need no table. Only 12pt was covered for Arial and Courier New; other font
/// sizes likewise stay on the conservative whole-height path. Family names
/// match exactly, so Arial Narrow/Black/Unicode MS borrow nothing from Arial.
fn measured_named_face_printed_grid_row_height(
    height: f64,
    normal_font: &NormalFont,
) -> Option<f64> {
    if normal_font.uses_theme_scheme || (normal_font.size_pt - 12.0).abs() >= 0.01 {
        return None;
    }
    let measured: &[(f64, f64)] = if normal_font.family.eq_ignore_ascii_case("Arial") {
        &ARIAL_12_PRINTED_GRID_ROW_HEIGHTS
    } else if normal_font.family.eq_ignore_ascii_case("Courier New") {
        &COURIER_NEW_12_PRINTED_GRID_ROW_HEIGHTS
    } else {
        return None;
    };
    measured
        .iter()
        .find(|(worksheet_height, _)| (height - worksheet_height).abs() < 0.01)
        .map(|(_, printed_height)| *printed_height)
}

/// How Excel's printed grid maps this workbook's worksheet row height. This
/// belongs to the resolved Normal face rather than to the row or to the
/// Calibri/Aptos substitution set.
///
/// One factor per export, sweeping only the first `<font>` of `xl/styles.xml`
/// on `03_inventory_en.xlsx` with `ht=40 customHeight="true"` data rows and a
/// byte-identical re-zip as the control:
///
/// | Normal font | printed track |
/// | --- | ---: |
/// | Calibri 9, Calibri 10 | 40 |
/// | Calibri 11, 12, 13 | 37 |
/// | Calibri 14, 16 | 38 |
/// | Aptos 11, Aptos Narrow 11, Aptos Narrow 12 | 37 |
/// | Segoe UI 10, Segoe UI 11 | 40 |
/// | Arial 10, Arial 11 | 40 |
/// | Times New Roman 11, Century Gothic 10 | 40 |
///
/// Nothing else moved it: `defaultRowHeight`, `sheetFormatPr@customHeight`,
/// the frozen pane, the sheet zoom, the paper size and the fit-to-page scale
/// were each swapped on their own and all left the track where it was.
/// Declaring that workbook's same Calibri 11 through `<scheme val="minor"/>`
/// instead of by name also left it at 37 — but that is a property of *its*
/// theme, not of the scheme flag, as the sweep below shows.
///
/// Calibri and Aptos are the faces the standard theme's minor scheme
/// resolves to, and the two this reference machine has no face of its own
/// for — issue #1047 measures the same substitution on the dimension-less
/// path — but that is the reading, not the measurement.
///
/// What a `<scheme>` font resolves to is the *theme's* business, and the two
/// corpora write two different themes. Sweeping the probe workbook of issue
/// #1094 one factor at a time, byte-identical re-zip control each time:
///
/// | fonts\[0\] | theme minor scheme | printed track for `ht=36` |
/// | --- | --- | ---: |
/// | Calibri 11, `<scheme val="minor"/>` | Office, per-script faces | 36 |
/// | Calibri 11, no `scheme` | Office, per-script faces | 33 |
/// | Calibri 11, `<scheme val="minor"/>` | script faces stripped | 33 |
/// | Calibri 11, `<scheme val="minor"/>` | `script="Hang"` -> Calibri | 33 |
///
/// So the scheme flag alone decides nothing. An Office theme names a face per
/// script and Excel resolves the UI script's — Malgun Gothic on this machine,
/// the face issue #1047 measures the dimension-less path against — and that
/// face keeps the grid. A theme listing no script faces leaves the scheme on
/// its `<a:latin>` Calibri, which compacts exactly as a font naming Calibri
/// outright does. `03_inventory_en.xlsx` carries such a bare theme, which is
/// why adding `<scheme val="minor"/>` to *its* Normal font leaves the track at
/// 37 while the same edit on the probe leaves it at 36.
///
/// Two limits are known and unmodelled: the step to 0.95 from 14pt up (no
/// tracked workbook declares one, so those compact at 0.92, a point short at
/// `ht=40`), and a scheme font naming no face at all, which resolves to the UI
/// face whatever the theme lists — that workbook has no `<name>` for
/// `extract_normal_font` to return, so it lands on the `None` arm here and
/// compacts. No tracked workbook writes one.
///
fn measured_printed_grid_row_height(height: f64, normal_font: Option<&NormalFont>) -> Option<f64> {
    let font = match normal_font {
        // Excel's own Normal font is Calibri/Aptos 11; a workbook we cannot
        // read a stylesheet from is laid out against it.
        None => return Some((height * 0.92).round()),
        // Resolved by script to the UI face, which is not one Excel remaps.
        Some(font) if font.uses_theme_scheme && font.theme_declares_script_faces => return None,
        Some(font) => font,
    };

    if font.size_pt >= REMAPPED_NORMAL_MIN_SIZE_PT && names_a_substituted_family(font) {
        return Some((height * 0.92).round());
    }
    measured_named_face_printed_grid_row_height(height, font)
}

/// Convert an OOXML row height to the whole-point track emitted by native
/// Excel's macOS PDF path. Excel exposes the stored value in points in the
/// worksheet UI, and its PDF grid snaps that to whole PDF points — after
/// compacting it, for the Normal fonts that compact at all.
///
/// Where Calibri/Aptos compacts, the ten XLSX audit workbooks (Calibri 11
/// Normal) map
/// their two repeated fixed heights consistently: 15pt -> 14pt and
/// 25.5pt -> 23pt. "Consistently" is measured, not assumed: reading the
/// golden exports' horizontal rules with `mutool draw -F trace` gives 23.00pt
/// for every `ht=25.5 customHeight="true"` header in all ten, and 14.00pt for
/// every `ht=15 customHeight="true"` row. Issue #658 reports two of them at
/// 24.00pt ("50 px @150 DPI"); that reads the band's *outer* extent — a 23pt
/// track plus the 1pt rule bounding each end is 24pt, or 50px at 150 DPI —
/// where this maps rule centre to rule centre. Sweeping that same workbook
/// over eight declared heights pins the factor at 0.92 rather than the 0.925
/// a single sample allows: 20 -> 18 and 49.5 -> 46 both need it.
///
/// Where it does not, the height simply truncates: the reported fit-to-page
/// workbook of issue #1068 (Segoe UI 10 Normal) exports 12/15/18/25/30/40/49
/// for declared 12/15/18/25.5/30/40/49.5, and its nine row boundaries down
/// the page all land within 0.12pt of that model.
/// A fresh Excel 16.112.3 export found a narrower exception: the theme-scheme
/// Trebuchet MS 10 workbook of #1262 snaps 19.5pt custom rows up to 20pt
/// (#1514). That combination remains on the conservative truncating path until
/// its fractional-height and theme controls are measured.
///
/// Both declared and recomputed worksheet heights go through here. An
/// auto-sized row additionally prints at the taller of this track and the
/// height its own font needs: the same `ht=15` row measures 14.00pt in Arial
/// 10 and 15.00pt in Malgun Gothic 10 in the golden exports. That font term is
/// not applied yet (issue #709), so Korean auto rows print 1.00pt short.
///
/// Keep this conversion in the XLSX parser rather than the generic table
/// renderer so DOCX/PPTX table heights retain their native semantics.
pub(super) fn native_excel_pdf_row_height(height: f64, normal_font: Option<&NormalFont>) -> f64 {
    match measured_printed_grid_row_height(height, normal_font) {
        Some(measured) => measured.max(1.0),
        None => height.floor().max(1.0),
    }
}

/// Whether this Normal font takes the 3pt bottom-aligned descent floor.
///
/// Keep this separate from [`measured_printed_grid_row_height`]. The older
/// issues #1097/#1199 probe matrix measured Calibri/Aptos compaction and the
/// descent seat together, but the issue #1224 Arial/Courier face sweep
/// measured row tracks only. Borrowing its new face set for the text seat
/// would move baselines on evidence that says nothing about them.
fn uses_compacted_bottom_aligned_descent_floor(normal_font: Option<&NormalFont>) -> bool {
    match normal_font {
        None => true,
        Some(font) if font.uses_theme_scheme && font.theme_declares_script_faces => false,
        Some(font) => {
            font.size_pt >= REMAPPED_NORMAL_MIN_SIZE_PT && names_a_substituted_family(font)
        }
    }
}

/// How far above its row's bottom boundary Excel holds a bottom-aligned
/// cell's baseline in this workbook, however small the font, in points
/// (issues #1097, #1199).
///
/// The older floor probe matrix measured this together with its Calibri/Aptos
/// track split. Six purpose-built workbooks print their declared tracks whole
/// and floor the seat at 4pt; the compacting corpus workbooks floor it at 3pt,
/// read off a ruled re-export of `10_kpi_tracker_en`'s note cell over an
/// eleven-size sweep. Issue #1224 added row-only measurements for Arial and
/// Courier New, so those faces affect the track without being guessed into
/// this separately measured floor.
///
/// Issue #1097 read the compacting family as having no floor, from the one
/// sample then available — `09_expense_report_en`'s Arial Bold 14 title, whose
/// bare `round(0.211914 x 14)` is already 3 and so cannot tell a 3pt floor
/// from none. The sizes at or under 11pt can, and #1199 measured them.
///
/// The mechanism behind the split is the open part. A workbook cannot be
/// pushed from one family to the other by anything the probes varied — the
/// title's row height, its spill, the cell's border, or the workbook's Normal
/// font *family and size*, which is Calibri 11 on both sides of the divide.
/// Within that floor probe matrix, what differs is whether the Normal font
/// resolves through a theme that names per-script faces (issues #1068,
/// #1094). That statement does not extend to the row-only face sweep.
pub(super) fn bottom_aligned_descent_floor_pt(normal_font: Option<&NormalFont>) -> f64 {
    if uses_compacted_bottom_aligned_descent_floor(normal_font) {
        crate::render::typst_gen::COMPACTED_SHEET_CELL_MIN_DESCENT_SEAT_PT
    } else {
        crate::render::typst_gen::SHEET_CELL_MIN_DESCENT_SEAT_PT
    }
}

/// Horizontal space an icon-set icon takes before its cell's value.
///
/// Excel reserves the icon's advance and then aligns the value in what is left
/// to its right; the icon itself is drawn out of layout here, so without this
/// a centred value centres in the whole cell and lands left of Excel's (#652).
///
/// Fitted, not derived. `10_kpi_tracker_en` is the only workbook in the corpus
/// with icon sets *and* a ground-truth Excel export to measure against: every
/// value in its icon column sat 4.79-5.01pt left of Excel's, and this reserve
/// closes that to within 0.4pt.
///
/// Its icons are `3Arrows`. Other tracked fixtures carry `3TrafficLights1`,
/// `3Flags`, `3Symbols`, `4Rating`, `5ArrowsGray` and more, none of which has
/// a ground truth here, so none was used to derive or check this. A set whose
/// icons are a different width will want a different advance — the honest
/// shape of this is per-icon-set, once there is something to measure it on.
const ICON_SET_VALUE_RESERVE_PT: f64 = 9.6;

/// Cell insets for spreadsheet tables. Excel's native single-line track is
/// asymmetric around bottom-aligned text: 1pt above and 1.5pt below. Typst's
/// default 5pt vertical inset overflowed auto-height rows (issue #396), while
/// a 1pt bottom inset left them about 0.5pt short (issue #411).
///
/// The sides are asymmetric too — 3pt left against 2pt right — measured on the
/// ten native Excel for Mac exports under
/// `tests/golden_mocks/business/expected/xlsx/`. With this pair the 139
/// left-aligned runs sit on Excel's pen origin to a median of 0.000pt and the
/// 135 right-aligned ones on its pen end to a mean of +0.006pt; a 3pt right
/// inset leaves those same right-aligned runs a whole point short (-0.994pt).
///
/// Issue #657 set both sides to 3pt on the reading that an asymmetric pair
/// moves every *centred* run by half the difference. That much is right, and
/// it is why [`aligned_cell_padding`] splits this total evenly for a centred
/// cell; what it got wrong was carrying the symmetry over to the right-aligned
/// case, where Excel's own export puts the edge a whole point further out
/// (issue #1157).
///
/// The 5pt total is also the `+5` the default column-width formula carries at
/// the Calibri 11 workbook default — the column formula and the text box are
/// the same padding seen from two sides.
///
/// Neither side is a constant of Excel's: both step with the cell font's
/// whole-point digit advance, and this pair is what that step gives a Calibri
/// 11 cell. [`styled_cell_padding`] prices both sides per cell (issues #1165,
/// #1232); this constant remains the Calibri 11 default and the source of the
/// font-independent vertical insets.
pub(super) const XLSX_CELL_PADDING: crate::ir::Insets = crate::ir::Insets {
    top: 1.0,
    right: 2.0,
    bottom: 1.5,
    left: 3.0,
};

/// Whether this cell's wrapped text needs more than the single line its row's
/// mapped track allows.
///
/// `wrapText` only says the cell *may* wrap. What decides is whether the text
/// fits the width it has — its own column, or the whole merge when it spans
/// several — after the horizontal inset the cell is laid out with. An explicit
/// line break always needs a second line.
fn cell_wraps_past_one_line(
    ctx: &SheetContext,
    col_idx: u32,
    row_idx: u32,
    col_span: u32,
    runs: &[Run],
    umya_cell: Option<&umya_spreadsheet::Cell>,
    cell_padding: Insets,
) -> bool {
    if runs.is_empty() {
        return false;
    }
    let has_wrap_text: bool = umya_cell
        .and_then(|cell| cell.get_style().get_alignment().cloned())
        .map(|alignment| *alignment.get_wrap_text())
        .unwrap_or(false);
    if !has_wrap_text {
        return false;
    }
    if runs.iter().any(|run| run.text.contains('\n')) {
        return true;
    }
    let available_width: f64 = (col_idx..col_idx + col_span)
        .map(|col| {
            ctx.column_widths
                .get((col - ctx.col_start) as usize)
                .copied()
                .unwrap_or(0.0)
        })
        .sum::<f64>()
        - cell_padding.left
        - cell_padding.right
        - ctx.cell_indent_pt(col_idx, row_idx);
    estimate_text_width_pt(runs) > available_width
}

/// The height a row prints at: the fixed track of
/// `printed_grid_row_height_pt`, calibrated to native Excel's PDF grid.
/// Exception: auto-sized rows (customHeight=false) whose wrapped text needs a
/// second line stay content-driven — our text metrics differ slightly from
/// Excel's and a fixed height could clip a wrapped line.
///
/// The exception is deliberately about the *text*, not about the `wrapText`
/// flag. Keying it on the flag made it fire on rows that never wrap: the ten
/// business mocks set `wrapText` on every data cell, so every ht=15 auto row
/// was sized by its own content box instead of by Excel's track — 15.00pt
/// against Excel's 14.00pt on the six Latin workbooks (issue #710), and
/// 22.32pt against 15.00pt on the Korean ones, where the East Asian line
/// factor compounds it (issue #709).
fn printed_row_height(
    sheet: &umya_spreadsheet::Worksheet,
    row_idx: u32,
    row_wraps_past_one_line: bool,
    ctx: &SheetContext,
) -> Option<f64> {
    let is_custom_height: bool = sheet
        .get_row_dimension(&row_idx)
        .map(|row| *row.get_custom_height())
        .unwrap_or(false);
    if !is_custom_height && row_wraps_past_one_line {
        return None;
    }
    Some(printed_grid_row_height_pt(
        sheet,
        row_idx,
        ctx.normal_font.as_ref(),
        Some(&ctx.row_boundary_points),
    ))
}

/// The worksheet height of a row that records none: what Excel recomputes
/// from the Normal font, or the declared hint where that recompute is
/// unmeasured.
pub(super) fn worksheet_default_row_height_pt(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font: Option<&NormalFont>,
) -> f64 {
    recomputed_default_row_height_pt(sheet, normal_font)
        .unwrap_or_else(|| declared_default_row_height_pt(sheet))
}

/// The worksheet height of one particular auto-sized row — one recording no
/// `ht`, or recording one without `customHeight`. Excel sizes such a row from
/// the tallest font its *own cells* carry; the sheet's recomputed default only
/// covers a row whose cells hold nothing taller.
///
/// Native Excel-for-Mac export of `issue_1060_sheet_row_line_box_probe.xlsx`,
/// whose Normal font is a theme-scheme Calibri 11 that none of its cells use;
/// AppleScript `row height of row N` and the `mutool draw -F trace` baseline
/// pitch agree on every row (issue #1140):
///
/// | rows | cell font | Excel row |
/// | --- | --- | ---: |
/// | 1-6, 8-13 | Malgun Gothic 14 | 20.00pt |
/// | 15-17 | Malgun Gothic 24 | 35.00pt |
///
/// Sizing all fifteen from the Normal font's own 17pt track instead left the
/// last of them 88.00pt up the page.
///
/// The term only ever raises the track. What Excel gives a row whose cells
/// are *smaller* than the Normal font is not measured here, and neither is a
/// size the face's series skips, so neither is interpolated between measured
/// points that step irregularly (15pt at 10, 17pt at 11, 20pt at 14). A row
/// holding something taller than the Normal font at a size no series covers
/// falls back to its own cached `ht` where it has one: that is the height
/// Excel last measured for this very text, and the sheet default would print
/// a 24pt title into a 15pt row.
///
/// A sheet marking its default `customHeight` is left alone as well: that
/// declared default is honoured for `ht`-less rows (issue #1047), and whether
/// a tall cell still grows one of them was never exported.
///
/// `cached_height_pt` is the `ht` such a row records without `customHeight` —
/// Excel's own last recompute of it, kept as the base only where this
/// workbook's Normal font has no measured recompute of its own (issue #1151).
fn auto_row_height_pt(
    sheet: &umya_spreadsheet::Worksheet,
    row_idx: u32,
    normal_font: Option<&NormalFont>,
    cached_height_pt: Option<f64>,
) -> f64 {
    let base_height_pt: f64 = match cached_height_pt {
        Some(cached_height_pt) => {
            recomputed_default_row_height_pt(sheet, normal_font).unwrap_or(cached_height_pt)
        }
        None => worksheet_default_row_height_pt(sheet, normal_font),
    };
    if *sheet.get_sheet_format_properties().get_custom_height() {
        return base_height_pt;
    }
    let Some(font) = normal_font else {
        return base_height_pt;
    };
    let Some(tallest_cell_size_pt) = tallest_cell_font_size_pt(sheet, row_idx) else {
        return base_height_pt;
    };
    if tallest_cell_size_pt <= font.size_pt {
        return base_height_pt;
    }
    let unmeasured_height_pt: f64 = cached_height_pt.unwrap_or(base_height_pt);
    measured_row_height_pt(font, tallest_cell_size_pt)
        .unwrap_or(unmeasured_height_pt)
        .max(base_height_pt)
}

/// The largest font size any cell of this row states, ignoring cells that
/// state none — those inherit the Normal font, which the caller already has.
fn tallest_cell_font_size_pt(sheet: &umya_spreadsheet::Worksheet, row_idx: u32) -> Option<f64> {
    sheet
        .get_collection_by_row(&row_idx)
        .into_iter()
        .filter_map(|cell| cell.get_style().get_font())
        .map(|font| *font.get_size())
        .filter(|size_pt| *size_pt > 0.0)
        .max_by(f64::total_cmp)
}

/// The whole-point track a row occupies in Excel's printed grid, whatever its
/// own content would need.
///
/// This is the one row metric anything laid out on the page may use — cells
/// and drawing anchors alike. Excel prints a drawing against the same grid it
/// prints the cells against: on `issue_1066_blip_effect_picture.xlsx` the
/// worksheet seats the picture 96.00pt down and 112.00pt tall over 16pt rows,
/// and the export draws it 90.00pt down and 105.00pt tall over the 15pt track
/// those rows compact to — six and seven of each, the ratio exact
/// (issue #1102).
///
/// A recorded `ht` is that track only when the row marks it `customHeight`.
/// Without the flag it is a cached auto-height that Excel discards, sizing the
/// row from the Normal font on load as it does a row recording nothing at all:
/// swept one Normal font size per export of the same workbook, whose rows 1,
/// 3, 4 and 5 carry `ht="16"` and no `customHeight` while row 7 carries no
/// dimension, `height of row 1` and `height of row 7` answer the same number
/// at every size, and neither is 16 except where the recompute lands there
/// (issue #1151).
///
/// The returned table-row height governs this row's printed baseline, not
/// merely its worksheet `height` property. Excel reserves one PDF point when
/// this row begins with `thickTop`, plus one when the preceding row ends with
/// `thickBot`. Native PDF baseline pitches on
/// `SH107-9-x-9-Formatted-Table.xlsx` are 19/17/19/19/17/17/17/17pt against
/// its bare 17pt track, exactly the previous-bottom plus current-top pattern. A
/// separate 8-24pt Normal-font sweep kept each term at 1pt, so apply it after
/// the font-dependent grid mapping rather than scaling it. A `customHeight`
/// already declares the full boundary, so only flags belonging to automatic
/// rows contribute: adding the custom flags in `issue_1181_fit_to_height.xlsx`
/// again grows its native 161.87pt chart area to 163.43pt (issue #1228).
pub(super) fn printed_grid_row_height_pt(
    sheet: &umya_spreadsheet::Worksheet,
    row_idx: u32,
    normal_font: Option<&NormalFont>,
    row_boundary_points: Option<&super::row_boundaries::RowBoundaryPoints>,
) -> f64 {
    let dimension: Option<&umya_spreadsheet::structs::Row> = sheet.get_row_dimension(&row_idx);
    let declared_height: Option<f64> = dimension
        .map(|row| *row.get_height())
        .filter(|height| *height > 0.0);
    let is_custom_height: bool = dimension
        .map(|row| *row.get_custom_height())
        .unwrap_or(false);
    let worksheet_height: f64 = match (is_custom_height, declared_height) {
        (true, Some(height)) => height,
        (_, cached_height) => auto_row_height_pt(sheet, row_idx, normal_font, cached_height),
    };
    let boundary_points: f64 = f64::from(
        row_boundary_points
            .and_then(|points| points.get(&row_idx))
            .copied()
            .unwrap_or(0),
    );
    native_excel_pdf_row_height(worksheet_height, normal_font) + boundary_points
}

/// The outline a merged range prints: each side taken from the members that
/// sit on that edge, rather than from the top-left member alone.
///
/// Excel writes a range's border format onto its constituent cells, so a rule
/// under a two-row header lands on the *bottom* row's cells and a rule down the
/// right-hand side lands on the right column's — neither of which the top-left
/// member records. Collapsing the range to that one cell dropped both
/// (issue #939).
///
/// One IR border holds a single side each, so the first member along an edge
/// that declares that side wins. Excel lets the members disagree and paints
/// each segment from its own cell; a range whose edge is formatted as a unit —
/// which is what applying a border to a merged range produces — has them all
/// agreeing anyway.
fn merged_range_border(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    col: u32,
    row: u32,
    info: &MergeInfo,
) -> Option<CellBorder> {
    let last_col: u32 = col + info.col_span.saturating_sub(1);
    let last_row: u32 = row + info.row_span.saturating_sub(1);
    let side_of = |member_col: u32, member_row: u32| -> Option<CellBorder> {
        sheet
            .get_cell((member_col, member_row))
            .and_then(|cell| extract_cell_borders(cell, ctx.theme.as_ref()))
    };
    let first_along = |cells: &mut dyn Iterator<Item = (u32, u32)>,
                       pick: fn(CellBorder) -> Option<BorderSide>|
     -> Option<BorderSide> {
        cells
            .filter_map(|(c, r)| side_of(c, r).and_then(pick))
            .next()
    };

    let border = CellBorder {
        top: first_along(&mut (col..=last_col).map(|c| (c, row)), |b| b.top),
        bottom: first_along(&mut (col..=last_col).map(|c| (c, last_row)), |b| b.bottom),
        left: first_along(&mut (row..=last_row).map(|r| (col, r)), |b| b.left),
        right: first_along(&mut (row..=last_row).map(|r| (last_col, r)), |b| b.right),
    };
    let CellBorder {
        top,
        bottom,
        left,
        right,
    } = &border;
    (top.is_some() || bottom.is_some() || left.is_some() || right.is_some()).then_some(border)
}

/// Lay a table style's rules under whatever borders the cell declares itself.
///
/// Direct cell formatting beats the table style in Excel, so a side the cell
/// already states is left alone and the style only fills in the rest
/// (issue #1080) — the same precedence the banded fill follows.
fn add_table_style_rules(
    declared: Option<CellBorder>,
    ctx: &SheetContext,
    col: u32,
    row: u32,
) -> Option<CellBorder> {
    let Some(rules) = ctx
        .table_styles
        .iter()
        .find_map(|style| style.border_at(col, row))
    else {
        return declared;
    };
    let Some(declared) = declared else {
        return Some(rules);
    };
    Some(CellBorder {
        top: declared.top.or(rules.top),
        bottom: declared.bottom.or(rules.bottom),
        left: declared.left.or(rules.left),
        right: declared.right.or(rules.right),
    })
}

/// Whether the cell XF merely repeats the workbook's default dark text ink.
///
/// An un-tinted `theme="1"` font colour is not a direct colour veto. Excel
/// copies it into XFs that carry unrelated direct formatting: SH107's I1
/// names the regular theme-dark Normal font beside its pale-yellow fill, but
/// the `TableStyleMedium2` header still prints the run white (issue #1230).
/// A non-default RGB, indexed colour, or tinted theme colour remains a real
/// direct override and keeps precedence over the table style.
fn cell_font_repeats_default_dark_ink(cell: &umya_spreadsheet::Cell) -> bool {
    let Some(font) = cell.get_style().get_font() else {
        return false;
    };
    let color = font.get_color();
    color.get_argb().is_empty()
        && *color.get_theme_index() == 1
        && color.get_tint().abs() < f64::EPSILON
}

/// Build TableRows for a range of rows in a sheet.
pub(super) fn build_rows_for_range(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    row_start: u32,
    row_end: u32,
) -> Vec<TableRow> {
    let num_rows = (row_end - row_start + 1) as usize;
    let mut rows = Vec::with_capacity(num_rows);
    for row_idx in row_start..=row_end {
        let mut cells = Vec::with_capacity(ctx.num_cols);
        let mut row_wraps_past_one_line = false;
        // A row-level `customFormat` style paints the row's cell-less grid
        // positions too — Excel fills the whole printed band, including
        // spill-reached columns holding no `<c>` element (issue #718).
        let row_style_fill: Option<Color> = sheet
            .get_row_dimension(&row_idx)
            .and_then(|row| extract_style_background(row.get_style(), ctx.theme.as_ref()));
        for col_idx in ctx.col_start..=ctx.col_end {
            // Skip cells that are part of a merge but not the top-left
            if ctx.merge_skips.contains(&(col_idx, row_idx)) {
                continue;
            }

            // umya-spreadsheet tuple is (column, row), both 1-indexed
            let umya_cell = sheet.get_cell((col_idx, row_idx));
            let cell_indent_pt: f64 = ctx.cell_indent_pt(col_idx, row_idx);
            let mut value = umya_cell.map(formatted_cell_value).unwrap_or_default();
            if let Some(cell) = umya_cell
                && let Some(number_format) = cell.get_style().get_number_format()
                && uses_native_arabic_digits(number_format.get_format_code())
            {
                value = to_arabic_indic_digits(&value);
            }

            // Extract formatting from the cell
            let mut text_style = umya_cell
                .map(|cell| {
                    extract_cell_text_style(cell, ctx.normal_font.as_ref(), ctx.theme.as_ref())
                })
                .unwrap_or_default();
            // A table style prints its header row bold. The cell's own font
            // wins where it declares a weight, and conditional formatting
            // below overrides both (issue #1080).
            if ctx
                .table_styles
                .iter()
                .any(|style| style.bolds_header_at(col_idx, row_idx))
            {
                text_style.bold.get_or_insert(true);
            }
            // A Medium table fills its header row in the accent and prints the
            // runs on it white. A real direct font colour keeps precedence,
            // but the default theme-dark ink repeated in a cell XF does not
            // veto the table header's colour (issues #1125, #1230).
            if let Some(header_ink) = ctx
                .table_styles
                .iter()
                .find_map(|style| style.header_text_color_at(col_idx, row_idx))
                && (text_style.color.is_none()
                    || umya_cell.is_some_and(cell_font_repeats_default_dark_ink))
            {
                text_style.color = Some(header_ink);
            }
            // Excel prices both horizontal sides of the cell's text box from
            // the cell's own font. Read before the runs take the style, and
            // use the same box for the width the line has and where either
            // aligned edge sits (issues #1165, #1232).
            let cell_padding: Insets = styled_cell_padding(&text_style, ctx.normal_font.as_ref());
            let (cell_alignment, cell_vertical_align) = umya_cell
                .map(extract_cell_alignment)
                .unwrap_or((None, None));
            let mut background = umya_cell
                .and_then(|cell| extract_cell_background(cell, ctx.theme.as_ref()))
                .or(if umya_cell.is_none() {
                    row_style_fill
                } else {
                    None
                });
            // A merged range is one IR cell, but Excel composes its outline
            // from the members on each edge, so reading only the top-left one
            // loses every side the range declares elsewhere (issue #939).
            let border = match ctx.merge_tops.get(&(col_idx, row_idx)) {
                Some(info) => merged_range_border(sheet, ctx, col_idx, row_idx, info),
                None => umya_cell.and_then(|cell| extract_cell_borders(cell, ctx.theme.as_ref())),
            };
            let border: Option<CellBorder> = add_table_style_rules(border, ctx, col_idx, row_idx);

            // Apply conditional formatting overrides
            let mut data_bar = None;
            let mut icon_text = None;
            let mut icon_color = None;
            let mut icon_shading = None;
            if let Some(ovr) = ctx.cond_fmt_overrides.get(&(col_idx, row_idx)) {
                if ovr.background.is_some() {
                    background = ovr.background;
                }
                if ovr.font_color.is_some() {
                    text_style.color = ovr.font_color;
                }
                if let Some(bold) = ovr.bold {
                    text_style.bold = Some(bold);
                }
                data_bar = ovr.data_bar.clone();
                icon_text = ovr.icon_text.clone();
                icon_color = ovr.icon_color;
                icon_shading = ovr.icon_shading;
            }
            let sparkline = ctx.sparklines.get(&(col_idx, row_idx)).cloned();

            // Rich-text shared strings carry per-run formatting (bold labels,
            // per-run fonts/colors) that the cell's single xf style loses —
            // emit one IR run per rich run instead of flattening.
            let rich_text: Option<umya_spreadsheet::RichText> =
                umya_cell.and_then(|cell| cell.get_cell_value().get_raw_value().get_rich_text());
            let runs: Vec<Run> = if let Some(rich_text) = rich_text {
                rich_text
                    .get_rich_text_elements()
                    .iter()
                    .filter(|element| !element.get_text().is_empty())
                    .map(|element| Run {
                        text: element.get_text().to_string(),
                        style: element
                            .get_run_properties()
                            .map(|font| apply_rich_run_font(&text_style, font, ctx.theme.as_ref()))
                            .unwrap_or_else(|| text_style.clone()),
                        href: None,
                        footnote: None,
                    })
                    .collect()
            } else if value.is_empty() {
                Vec::new()
            } else {
                vec![Run {
                    text: value,
                    style: text_style,
                    href: None,
                    footnote: None,
                }]
            };

            // Excel's "general" horizontal alignment follows the text
            // direction: cells whose text starts with a right-to-left script
            // print right-aligned.
            let cell_alignment: Option<crate::ir::Alignment> = cell_alignment.or_else(|| {
                runs.iter()
                    .flat_map(|run| run.text.chars())
                    .find_map(strong_direction)
                    .filter(|is_rtl| *is_rtl)
                    .map(|_| crate::ir::Alignment::Right)
            });
            let paragraph_alignment = cell_alignment.or_else(|| {
                umya_cell
                    .and_then(|cell| cell.get_value_number())
                    .map(|_| crate::ir::Alignment::Right)
            });

            let (col_span, row_span) = if let Some(info) = ctx.merge_tops.get(&(col_idx, row_idx)) {
                (info.col_span, info.row_span)
            } else {
                (1, 1)
            };

            let spill_width: Option<f64> = compute_spill_width(
                sheet,
                ctx,
                col_idx,
                row_idx,
                &runs,
                paragraph_alignment,
                col_span,
                umya_cell,
                cell_padding,
            );

            row_wraps_past_one_line |= cell_wraps_past_one_line(
                ctx,
                col_idx,
                row_idx,
                col_span,
                &runs,
                umya_cell,
                cell_padding,
            );

            let content = if runs.is_empty() {
                Vec::new()
            } else {
                vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle {
                        alignment: paragraph_alignment,
                        ..ParagraphStyle::default()
                    },
                    runs,
                })]
            };

            // An explicit cell fill wins; the table's banding only shows
            // through where the cell declares none (issue #532).
            let background: Option<Color> = background.or_else(|| {
                ctx.table_styles
                    .iter()
                    .find_map(|stripes| stripes.fill_at(col_idx, row_idx))
            });

            // An icon is drawn out of layout at the cell's left edge, so
            // it consumes no width and a centred value centres in the
            // whole cell. Excel reserves the icon's advance first and
            // aligns the value in what remains to its right, which is
            // where the extra left inset comes from (issue #652).
            let padding: Option<Insets> = {
                let aligned: Insets = aligned_cell_padding(cell_padding, paragraph_alignment);
                let base: Insets = match icon_text {
                    Some(_) => Insets {
                        left: aligned.left + ICON_SET_VALUE_RESERVE_PT,
                        ..aligned
                    },
                    None => aligned,
                };
                let indented: Insets =
                    indented_cell_padding(base, cell_indent_pt, paragraph_alignment);
                (indented != default_cell_padding(ctx.normal_font.as_ref())).then_some(indented)
            };

            cells.push(TableCell {
                content,
                col_span,
                row_span,
                border,
                background,
                background_alpha: None,
                data_bar,
                sparkline,
                padding,
                icon_text,
                icon_color,
                icon_shading,
                spill_width,
                vertical_align: cell_vertical_align,
            });
        }

        let height: Option<f64> = printed_row_height(sheet, row_idx, row_wraps_past_one_line, ctx);

        rows.push(TableRow {
            cells,
            height,
            minimum_height: None,
        });
    }
    rows
}

/// The point metric every column width is scaled by. Excel derives it from
/// the workbook Normal font; cell fonts do not participate (issue #366).
/// When `xl/styles.xml` was unreadable, fall back to the dominant cell font
/// — which on a sheet with no cells lands on the legacy 5.25pt default.
/// Shared by populated and drawing-only sheets so both scale from the same
/// digit metric (issue #620); drawing-only sheets still price every column at
/// the default width because their context carries no `<cols>` overrides.
pub(super) fn resolve_column_unit_pt(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font: Option<&NormalFont>,
) -> f64 {
    normal_font
        .map(|font| column_unit_pt(&font.family, font.size_pt))
        .unwrap_or_else(|| sheet_column_unit_pt(sheet))
}

/// Last column that contributes visible sheet ink before text overflow.
///
/// `Worksheet::get_highest_column_and_row` follows the worksheet dimension,
/// which includes value-less cells carrying non-visual metadata such as
/// `quotePrefix`. Excel does not let those cells claim printed width: SH107
/// declares A:K, but its quote-prefix-only J/K cells produce no second page
/// in the native export (issue #1229).
///
/// A value-less cell that paints a fill or border remains part of the printed
/// grid. The same applies to visible conditional formatting and a table style
/// whose band or rules cover otherwise empty cells.
fn inferred_printed_max_col(
    sheet: &umya_spreadsheet::Worksheet,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
    table_styles: &[crate::parser::xlsx::tables::TableStyleRange],
    cond_fmt_overrides: &HashMap<(u32, u32), crate::parser::cond_fmt::CondFmtOverride>,
) -> u32 {
    let cell_max: u32 = sheet
        .get_cell_collection()
        .iter()
        .filter(|cell| {
            !cell.get_cell_value().get_raw_value().is_empty()
                || extract_cell_background(cell, theme).is_some()
                || extract_cell_borders(cell, theme).is_some()
        })
        .map(|cell| *cell.get_coordinate().get_col_num())
        .max()
        .unwrap_or(0);
    let table_max: u32 = table_styles
        .iter()
        .filter_map(|style| style.painted_end_col())
        .max()
        .unwrap_or(0);
    let conditional_max: u32 = cond_fmt_overrides
        .iter()
        .filter(|(_, ovr)| {
            ovr.background.is_some() || ovr.data_bar.is_some() || ovr.icon_text.is_some()
        })
        .map(|(&(col, _), _)| col)
        .max()
        .unwrap_or(0);

    cell_max.max(table_max).max(conditional_max)
}

/// The last printed column once unwrapped text overflow is honoured.
///
/// Excel extends a sheet's printed range past its used range to every column
/// that a cell's overflowing text visibly reaches (issue #718). Probe-measured
/// on `NumberFormatTests.xlsx` (Tests sheet): deleting the trailing styled
/// `<col>` run, the row-level `customFormat`, the frozen pane, the `A1:XFD1`
/// selection, and the `pageSetup` element each left the printed edge at
/// column E, while deleting the rows whose D-column text paints past the used
/// range pulled it back to column D — and lengthening that text made the
/// native export print further trailing columns, adding horizontal spill
/// pages. The bound is the overflow's reach: not one column, not the page
/// width, and not the extent of any formatting.
///
/// Mirrors `compute_spill_width`'s spill conditions: only an unwrapped,
/// general- or left-aligned text cell whose row holds nothing between it and
/// the used-range edge paints past that edge.
#[allow(clippy::too_many_arguments)]
fn spill_reach_max_col(
    sheet: &umya_spreadsheet::Worksheet,
    max_col: u32,
    normal_font: Option<&NormalFont>,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
    merge_tops: &HashMap<(u32, u32), MergeInfo>,
    merge_skips: &HashSet<(u32, u32)>,
    unit_pt: f64,
    default_width_pt: f64,
) -> u32 {
    let column_width_pt = |col: u32| -> f64 {
        sheet
            .get_column_dimension_by_number(&col)
            .map(|c| column_width_to_pt(*c.get_width(), unit_pt))
            .unwrap_or(default_width_pt)
    };

    // Only a row's last value-bearing cell can paint past the used range —
    // any populated cell to its right would block the spill first.
    let mut last_populated_by_row: HashMap<u32, (u32, &umya_spreadsheet::Cell)> = HashMap::new();
    for cell in sheet.get_cell_collection() {
        if cell.get_cell_value().get_raw_value().is_empty() {
            continue;
        }
        let col: u32 = *cell.get_coordinate().get_col_num();
        let row: u32 = *cell.get_coordinate().get_row_num();
        let entry = last_populated_by_row.entry(row).or_insert((col, cell));
        if col >= entry.0 {
            *entry = (col, cell);
        }
    }

    let mut extended_max_col: u32 = max_col;
    for (&row, &(col, cell)) in &last_populated_by_row {
        // A merged cell clips at the merge edge instead of spilling.
        if merge_tops.contains_key(&(col, row)) || merge_skips.contains(&(col, row)) {
            continue;
        }
        let text: String = formatted_cell_value(cell);
        if text.is_empty() || text.contains('\n') {
            continue;
        }
        let has_wrap_text: bool = cell
            .get_style()
            .get_alignment()
            .map(|alignment| *alignment.get_wrap_text())
            .unwrap_or(false);
        if has_wrap_text {
            continue;
        }
        // The same alignment resolution the cell loop applies: explicit
        // horizontal alignment, else RTL text prints right-aligned, else
        // numbers print right-aligned — and only general/left spills right.
        let alignment: Option<crate::ir::Alignment> = extract_cell_alignment(cell)
            .0
            .or_else(|| {
                text.chars()
                    .find_map(strong_direction)
                    .filter(|is_rtl| *is_rtl)
                    .map(|_| crate::ir::Alignment::Right)
            })
            .or_else(|| cell.get_value_number().map(|_| crate::ir::Alignment::Right));
        if !matches!(alignment, None | Some(crate::ir::Alignment::Left)) {
            continue;
        }
        // A merged range between the cell and the used-range edge blocks the
        // spill; value-bearing cells cannot occur there by construction.
        if ((col + 1)..=max_col)
            .any(|c| merge_tops.contains_key(&(c, row)) || merge_skips.contains(&(c, row)))
        {
            continue;
        }

        let style: TextStyle = extract_cell_text_style(cell, normal_font, theme);
        let estimate: f64 = estimate_line_width_pt(
            &text,
            style.font_family.as_deref(),
            style.font_size.unwrap_or(11.0),
        );
        let own_width: f64 = column_width_pt(col);
        let cell_padding: Insets = styled_cell_padding(&style, normal_font);
        let horizontal_inset: f64 = cell_padding.left + cell_padding.right;
        if estimate <= own_width - horizontal_inset {
            continue;
        }
        // The width the line needs, as `compute_spill_width` prices it; walk
        // whole trailing columns until the text's reach is covered.
        let needed_width: f64 = estimate + 4.0;
        let mut covered_width: f64 = (col..=max_col).map(column_width_pt).sum();
        let mut reach_col: u32 = max_col;
        while needed_width > covered_width && reach_col < MAX_XLSX_COLUMNS {
            reach_col += 1;
            covered_width += column_width_pt(reach_col);
        }
        extended_max_col = extended_max_col.max(reach_col);
    }
    extended_max_col
}

/// Prepare the shared context for processing a sheet (dimensions, merges, styles, etc.).
/// Returns (SheetContext, row_start, row_end) or None if the sheet is empty.
#[allow(clippy::too_many_arguments)]
pub(super) fn prepare_sheet_context(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font: Option<&NormalFont>,
    raw_cond_fmt_hints: Option<&super::cond_fmt_raw::RawCondFmtHints>,
    defined_names: &HashMap<String, String>,
    table_styles: Vec<crate::parser::xlsx::tables::TableStyleRange>,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
    cell_indents: Option<&super::indent::CellIndentLevels>,
    row_boundary_points: Option<&super::row_boundaries::RowBoundaryPoints>,
    sparklines: Option<&HashMap<(u32, u32), crate::ir::SparklineInfo>>,
) -> Option<(SheetContext, u32, u32)> {
    let (worksheet_max_col, mut max_row) = sheet.get_highest_column_and_row();
    let sparklines = sparklines.cloned().unwrap_or_default();
    let sparkline_max_col = sparklines.keys().map(|(col, _)| *col).max().unwrap_or(0);
    let sparkline_max_row = sparklines.keys().map(|(_, row)| *row).max().unwrap_or(0);
    max_row = max_row.max(sparkline_max_row);
    if (worksheet_max_col == 0 && sparkline_max_col == 0) || max_row == 0 {
        return None;
    }

    let print_area = find_print_area(sheet);
    let cond_fmt_overrides =
        build_cond_fmt_overrides(sheet, raw_cond_fmt_hints, defined_names, theme);
    let mut max_col: u32 = if print_area.is_some() {
        worksheet_max_col
    } else {
        inferred_printed_max_col(sheet, theme, &table_styles, &cond_fmt_overrides)
    }
    .max(sparkline_max_col);
    if max_col == 0 && print_area.is_none() {
        return None;
    }

    // Expand grid to include the extent of all merged ranges
    for range in sheet.get_merge_cells() {
        if let Some(c) = range.get_coordinate_end_col() {
            max_col = max_col.max(*c.get_num());
        }
        if let Some(r) = range.get_coordinate_end_row() {
            max_row = max_row.max(*r.get_num());
        }
    }

    let unit_pt: f64 = resolve_column_unit_pt(sheet, normal_font);
    let default_width_pt: f64 = default_column_width_pt(
        declared_default_column_width(sheet),
        declared_base_column_width(sheet),
        unit_pt,
    );
    let cell_indents: super::indent::CellIndentLevels = cell_indents.cloned().unwrap_or_default();
    let row_boundary_points: super::row_boundaries::RowBoundaryPoints =
        row_boundary_points.cloned().unwrap_or_default();
    let indent_unit_pt: f64 = resolve_indent_unit_pt(normal_font);
    let (merge_tops, merge_skips) = build_merge_maps(sheet);

    // Check for print area — limit to that range if defined. Without one,
    // the printed range grows past the used range to the columns that
    // unwrapped text overflow visibly reaches, as Excel prints them
    // (issue #718). An explicit print area is exact and never grows.
    let (col_start, col_end, row_start, row_end) = if let Some(pa) = print_area {
        (pa.start_col, pa.end_col, pa.start_row, pa.end_row)
    } else {
        let max_col: u32 = spill_reach_max_col(
            sheet,
            max_col,
            normal_font,
            theme,
            &merge_tops,
            &merge_skips,
            unit_pt,
            default_width_pt,
        );
        (1, max_col, 1, max_row)
    };

    let column_widths: Vec<f64> = (col_start..=col_end)
        .map(|col| {
            sheet
                .get_column_dimension_by_number(&col)
                .map(|c| column_width_to_pt(*c.get_width(), unit_pt))
                .unwrap_or(default_width_pt)
        })
        .collect();
    let num_cols = (col_end - col_start + 1) as usize;

    Some((
        SheetContext {
            col_start,
            col_end,
            num_cols,
            column_widths,
            default_column_width_pt: default_width_pt,
            merge_tops,
            merge_skips,
            cond_fmt_overrides,
            sparklines,
            normal_font: normal_font.cloned(),
            table_styles,
            theme: theme.cloned(),
            cell_indents,
            row_boundary_points,
            indent_unit_pt,
        },
        row_start,
        row_end,
    ))
}

#[cfg(test)]
mod literal_zero_section_tests {
    use super::literal_zero_section_text;

    #[test]
    fn decodes_all_escaped_whitespace_from_the_zero_section() {
        assert_eq!(
            literal_zero_section_text(r"#,##0_);[Red]\(#,##0\);\-\ \ "),
            Some("-  ".to_string())
        );
    }

    #[test]
    fn supports_quoted_and_empty_zero_sections() {
        assert_eq!(
            literal_zero_section_text(r#""TRUE";"TRUE";"FALSE""#),
            Some("FALSE".to_string())
        );
        assert_eq!(literal_zero_section_text("0;-0;"), Some(String::new()));
    }

    #[test]
    fn leaves_value_dependent_and_non_zero_specific_formats_to_umya() {
        assert_eq!(literal_zero_section_text("#,##0.00"), None);
        assert_eq!(literal_zero_section_text("0;-0;0"), None);
        assert_eq!(
            literal_zero_section_text(r#"0;[Red]-0;"unterminated"#),
            None
        );
        assert_eq!(
            literal_zero_section_text(r#"[=0]"zero";[>0]"positive";"negative""#),
            None
        );
        assert_eq!(literal_zero_section_text(r#"0;-0;[h]" hours""#), None);
    }
}
