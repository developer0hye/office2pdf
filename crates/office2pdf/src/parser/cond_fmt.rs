use std::collections::HashMap;

use crate::ir::{Color, DataBarInfo};
use crate::parser::xlsx::cond_fmt_raw::{RawCondFmtHint, RawCondFmtHints};
use crate::parser::xlsx::{CellPos, CellRange, parse_cell_ref};
use crate::parser::xml_util;

/// A conditional formatting override for a specific cell.
#[derive(Default)]
pub(crate) struct CondFmtOverride {
    pub background: Option<Color>,
    pub font_color: Option<Color>,
    pub bold: Option<bool>,
    pub data_bar: Option<DataBarInfo>,
    pub icon_text: Option<String>,
    pub icon_color: Option<Color>,
}

/// Parse an sqref string (e.g., "A1:C10" or "A1") into a list of CellRanges.
fn parse_sqref(sqref: &str) -> Vec<CellRange> {
    sqref
        .split_whitespace()
        .filter_map(|part| {
            if let Some((start_str, end_str)) = part.split_once(':') {
                let (sc, sr) = parse_cell_ref(start_str)?;
                let (ec, er) = parse_cell_ref(end_str)?;
                Some(CellRange {
                    start_col: sc,
                    start_row: sr,
                    end_col: ec,
                    end_row: er,
                })
            } else {
                let (c, r) = parse_cell_ref(part)?;
                Some(CellRange {
                    start_col: c,
                    start_row: r,
                    end_col: c,
                    end_row: r,
                })
            }
        })
        .collect()
}

use xml_util::parse_argb_color;

/// Try to get a numeric value from a cell.
fn cell_numeric_value(cell: &umya_spreadsheet::Cell) -> Option<f64> {
    let raw = cell.get_raw_value().to_string();
    if let Ok(v) = raw.parse::<f64>() {
        return Some(v);
    }
    cell.get_value().to_string().parse::<f64>().ok()
}

/// Evaluate a CellIs conditional formatting rule against a cell value.
fn evaluate_cell_is_rule(
    cell_val: f64,
    operator: &umya_spreadsheet::ConditionalFormattingOperatorValues,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
) -> bool {
    use umya_spreadsheet::ConditionalFormattingOperatorValues::*;

    let formula_val = rule.get_formula().and_then(|f| {
        let s = f.get_address_str();
        s.trim().parse::<f64>().ok()
    });

    let Some(threshold) = formula_val else {
        return false;
    };

    match operator {
        GreaterThan => cell_val > threshold,
        GreaterThanOrEqual => cell_val >= threshold,
        LessThan => cell_val < threshold,
        LessThanOrEqual => cell_val <= threshold,
        Equal => (cell_val - threshold).abs() < f64::EPSILON,
        NotEqual => (cell_val - threshold).abs() >= f64::EPSILON,
        Between => cell_val >= threshold,
        NotBetween => cell_val < threshold,
        _ => false,
    }
}

/// Extract formatting overrides from a conditional formatting rule's style.
fn extract_cond_fmt_style(rule: &umya_spreadsheet::ConditionalFormattingRule) -> CondFmtOverride {
    let mut result = CondFmtOverride::default();

    if let Some(style) = rule.get_style() {
        if let Some(bg) = style.get_background_color() {
            result.background = parse_argb_color(bg.get_argb());
        }
        // Differential formats (dxf) store solid CF fills in the pattern's
        // bgColor with no fgColor, which get_background_color misses.
        if result.background.is_none()
            && let Some(bg) = style
                .get_fill()
                .and_then(|fill| fill.get_pattern_fill()?.get_background_color())
        {
            result.background = parse_argb_color(bg.get_argb());
        }
        if let Some(font) = style.get_font() {
            if *font.get_bold() {
                result.bold = Some(true);
            }
            let color_argb = font.get_color().get_argb();
            if !color_argb.is_empty() && color_argb != "FF000000" {
                result.font_color = parse_argb_color(color_argb);
            }
        }
    }

    result
}

/// Parse an ARGB hex string from umya Color into an IR Color.
fn parse_umya_color_argb(color: &umya_spreadsheet::Color) -> Option<Color> {
    let argb = color.get_argb();
    if argb.is_empty() {
        return None;
    }
    parse_argb_color(argb)
}

/// Interpolate between two colors based on a ratio (0.0 = color_a, 1.0 = color_b).
fn interpolate_color(color_a: Color, color_b: Color, ratio: f64) -> Color {
    let ratio = ratio.clamp(0.0, 1.0);
    let r = (color_a.r as f64 + (color_b.r as f64 - color_a.r as f64) * ratio).round() as u8;
    let g = (color_a.g as f64 + (color_b.g as f64 - color_a.g as f64) * ratio).round() as u8;
    let b = (color_a.b as f64 + (color_b.b as f64 - color_a.b as f64) * ratio).round() as u8;
    Color::new(r, g, b)
}

/// Collect all numeric values in ranges from the sheet (for color scale min/max).
fn collect_numeric_values_in_ranges(
    sheet: &umya_spreadsheet::Worksheet,
    ranges: &[CellRange],
) -> Vec<f64> {
    let mut values = Vec::new();
    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if let Some(cell) = sheet.get_cell((col, row))
                    && let Some(val) = cell_numeric_value(cell)
                {
                    values.push(val);
                }
            }
        }
    }
    values
}

/// Compute the min, max, and range span of a set of values.
/// Returns `None` if the slice is empty.
fn compute_min_max(values: &[f64]) -> Option<(f64, f64, f64)> {
    if values.is_empty() {
        return None;
    }
    let min_val: f64 = values.iter().cloned().fold(f64::INFINITY, f64::min);
    let max_val: f64 = values.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
    let val_range: f64 = max_val - min_val;
    Some((min_val, max_val, val_range))
}

/// Apply a CellIs conditional formatting rule to matching cells in the given ranges.
/// Apply text-match conditional rules (containsText / notContainsText /
/// beginsWith / endsWith) using the rule's `text` attribute.
fn apply_text_rule(
    sheet: &umya_spreadsheet::Worksheet,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
    ranges: &[CellRange],
    overrides: &mut HashMap<CellPos, CondFmtOverride>,
) {
    use umya_spreadsheet::ConditionalFormatValues;
    let needle: &str = rule.get_text();
    if needle.is_empty() {
        return;
    }
    let fmt = extract_cond_fmt_style(rule);

    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                let Some(cell) = sheet.get_cell((col, row)) else {
                    continue;
                };
                let value = cell.get_formatted_value();
                let matched = match rule.get_type() {
                    ConditionalFormatValues::ContainsText => value.contains(needle),
                    ConditionalFormatValues::NotContainsText => !value.contains(needle),
                    ConditionalFormatValues::BeginsWith => value.starts_with(needle),
                    ConditionalFormatValues::EndsWith => value.ends_with(needle),
                    _ => false,
                };
                if matched {
                    let entry = overrides.entry((col, row)).or_default();
                    if fmt.background.is_some() {
                        entry.background = fmt.background;
                    }
                    if fmt.font_color.is_some() {
                        entry.font_color = fmt.font_color;
                    }
                    if fmt.bold.is_some() {
                        entry.bold = fmt.bold;
                    }
                }
            }
        }
    }
}

fn apply_cell_is_rule(
    sheet: &umya_spreadsheet::Worksheet,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
    ranges: &[CellRange],
    overrides: &mut HashMap<CellPos, CondFmtOverride>,
) {
    let operator = rule.get_operator();
    let fmt = extract_cond_fmt_style(rule);

    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if let Some(cell) = sheet.get_cell((col, row))
                    && let Some(val) = cell_numeric_value(cell)
                    && evaluate_cell_is_rule(val, operator, rule)
                {
                    let entry = overrides.entry((col, row)).or_default();
                    if fmt.background.is_some() {
                        entry.background = fmt.background;
                    }
                    if fmt.font_color.is_some() {
                        entry.font_color = fmt.font_color;
                    }
                    if fmt.bold.is_some() {
                        entry.bold = fmt.bold;
                    }
                }
            }
        }
    }
}

/// Apply a ColorScale conditional formatting rule to cells in the given ranges.
fn apply_color_scale_rule(
    sheet: &umya_spreadsheet::Worksheet,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
    ranges: &[CellRange],
    overrides: &mut HashMap<CellPos, CondFmtOverride>,
) {
    let Some(cs) = rule.get_color_scale() else {
        return;
    };

    let colors: Vec<Option<Color>> = cs
        .get_color_collection()
        .iter()
        .map(parse_umya_color_argb)
        .collect();

    if colors.len() < 2 {
        return;
    }

    let numeric_vals: Vec<f64> = collect_numeric_values_in_ranges(sheet, ranges);
    let Some((min_val, _max_val, val_range)) = compute_min_max(&numeric_vals) else {
        return;
    };

    let color_min: Color = colors[0].unwrap_or(Color::white());
    let color_max: Color = colors[colors.len() - 1].unwrap_or(Color::black());

    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if let Some(cell) = sheet.get_cell((col, row))
                    && let Some(val) = cell_numeric_value(cell)
                {
                    let ratio: f64 = if val_range.abs() < f64::EPSILON {
                        0.5
                    } else {
                        (val - min_val) / val_range
                    };

                    let color: Color = if colors.len() == 3 {
                        let color_mid: Color = colors[1].unwrap_or(Color::new(255, 255, 0));
                        if ratio <= 0.5 {
                            interpolate_color(color_min, color_mid, ratio * 2.0)
                        } else {
                            interpolate_color(color_mid, color_max, (ratio - 0.5) * 2.0)
                        }
                    } else {
                        interpolate_color(color_min, color_max, ratio)
                    };

                    let entry = overrides.entry((col, row)).or_default();
                    entry.background = Some(color);
                }
            }
        }
    }
}

/// Apply a DataBar conditional formatting rule to cells in the given ranges.
fn apply_data_bar_rule(
    sheet: &umya_spreadsheet::Worksheet,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
    ranges: &[CellRange],
    overrides: &mut HashMap<CellPos, CondFmtOverride>,
    raw_hint: Option<&RawCondFmtHint>,
) {
    let Some(db) = rule.get_data_bar() else {
        return;
    };

    let bar_color: Color = db
        .get_color_collection()
        .first()
        .and_then(parse_umya_color_argb)
        .unwrap_or(Color::new(0x63, 0x8E, 0xC6)); // default blue

    let numeric_vals: Vec<f64> = collect_numeric_values_in_ranges(sheet, ranges);
    let Some((range_min, range_max, _)) = compute_min_max(&numeric_vals) else {
        return;
    };

    // Resolve the bar axis from the cfvo pair; fixed axes (num/percent) are
    // independent of the observed values, exactly like Excel.
    let cfvos = db.get_cfvo_collection();
    let axis_min: f64 = cfvos
        .first()
        .and_then(|cfvo| resolve_data_bar_cfvo(cfvo, range_min, range_max))
        .unwrap_or(range_min);
    let axis_max: f64 = cfvos
        .get(1)
        .and_then(|cfvo| resolve_data_bar_cfvo(cfvo, range_min, range_max))
        .unwrap_or(range_max);
    let axis_range: f64 = axis_max - axis_min;

    // Excel maps the axis onto [minLength, maxLength] percent of the cell
    // width (spec defaults 10/90), so the minimum still shows a short bar.
    let min_length: f64 = f64::from(raw_hint.and_then(|hint| hint.min_length).unwrap_or(10));
    let max_length: f64 = f64::from(raw_hint.and_then(|hint| hint.max_length).unwrap_or(90));

    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if let Some(cell) = sheet.get_cell((col, row))
                    && let Some(val) = cell_numeric_value(cell)
                {
                    let fraction: f64 = if axis_range.abs() < f64::EPSILON {
                        0.5
                    } else {
                        ((val - axis_min) / axis_range).clamp(0.0, 1.0)
                    };
                    let pct: f64 = min_length + (max_length - min_length) * fraction;
                    let entry = overrides.entry((col, row)).or_default();
                    entry.data_bar = Some(DataBarInfo {
                        color: bar_color,
                        fill_pct: pct,
                    });
                }
            }
        }
    }
}

/// Resolve a dataBar cfvo to an absolute axis value. Returns None for types
/// that fall back to the observed range bounds (min/max/formula).
fn resolve_data_bar_cfvo(
    cfvo: &umya_spreadsheet::ConditionalFormatValueObject,
    range_min: f64,
    range_max: f64,
) -> Option<f64> {
    use umya_spreadsheet::ConditionalFormatValueObjectValues as CfvoType;
    match cfvo.get_type() {
        CfvoType::Number => cfvo.get_val().parse().ok(),
        CfvoType::Percent => {
            let pct: f64 = cfvo.get_val().parse().ok()?;
            Some(range_min + (range_max - range_min) * (pct / 100.0))
        }
        _ => None,
    }
}

/// Apply an IconSet conditional formatting rule to cells in the given ranges.
fn apply_icon_set_rule(
    sheet: &umya_spreadsheet::Worksheet,
    rule: &umya_spreadsheet::ConditionalFormattingRule,
    ranges: &[CellRange],
    overrides: &mut HashMap<CellPos, CondFmtOverride>,
    raw_hint: Option<&RawCondFmtHint>,
) {
    let numeric_vals: Vec<f64> = collect_numeric_values_in_ranges(sheet, ranges);
    let Some((min_val, _max_val, val_range)) = compute_min_max(&numeric_vals) else {
        return;
    };

    // Try to parse thresholds from IconSet cfvos
    let cfvo_thresholds: Vec<f64> = rule
        .get_icon_set()
        .map(|is| is.get_cfvo_collection())
        .unwrap_or(&[])
        .iter()
        .filter_map(|cfvo| {
            let pct: f64 = cfvo.get_val().parse().ok()?;
            Some(min_val + val_range * (pct / 100.0))
        })
        .collect();

    // Default to 3-icon equal-thirds if no thresholds available
    let thresholds: Vec<f64> = if cfvo_thresholds.len() >= 2 {
        cfvo_thresholds
    } else {
        vec![
            min_val,
            min_val + val_range / 3.0,
            min_val + val_range * 2.0 / 3.0,
        ]
    };

    let set_type: &str = raw_hint
        .and_then(|hint| hint.icon_set_type.as_deref())
        .unwrap_or("");
    let icons: Vec<(&'static str, Option<Color>)> =
        icon_set_glyphs(set_type, thresholds.len().max(3));

    for range in ranges {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if let Some(cell) = sheet.get_cell((col, row))
                    && let Some(val) = cell_numeric_value(cell)
                {
                    let icon_idx: usize = evaluate_icon_index(val, &thresholds, icons.len());
                    let (glyph, color) = &icons[icon_idx];
                    let entry = overrides.entry((col, row)).or_default();
                    entry.icon_text = Some((*glyph).to_string());
                    entry.icon_color = *color;
                }
            }
        }
    }
}

/// Excel's icon band colors (sampled from Excel's own PDF output).
const ICON_RED: Color = Color {
    r: 214,
    g: 85,
    b: 50,
};
const ICON_YELLOW: Color = Color {
    r: 234,
    g: 191,
    b: 87,
};
const ICON_GREEN: Color = Color {
    r: 104,
    g: 164,
    b: 144,
};
const ICON_GRAY: Color = Color {
    r: 128,
    g: 128,
    b: 128,
};
const ICON_BLACK: Color = Color { r: 0, g: 0, b: 0 };

/// Map an OOXML iconSet type to per-band (glyph, color) pairs, low band first.
/// An absent attribute means the spec default 3TrafficLights1. Unknown set
/// types fall back to colored arrows of the requested band count.
fn icon_set_glyphs(set_type: &str, band_count: usize) -> Vec<(&'static str, Option<Color>)> {
    let effective_type: &str = if set_type.is_empty() {
        "3TrafficLights1"
    } else {
        set_type
    };
    match effective_type {
        "3TrafficLights1" | "3TrafficLights2" | "3Signs" => vec![
            ("●", Some(ICON_RED)),
            ("●", Some(ICON_YELLOW)),
            ("●", Some(ICON_GREEN)),
        ],
        "4TrafficLights" => vec![
            ("●", Some(ICON_BLACK)),
            ("●", Some(ICON_RED)),
            ("●", Some(ICON_YELLOW)),
            ("●", Some(ICON_GREEN)),
        ],
        "3Symbols" | "3Symbols2" => vec![
            ("✗", Some(ICON_RED)),
            ("!", Some(ICON_YELLOW)),
            ("✓", Some(ICON_GREEN)),
        ],
        "3Flags" => vec![
            ("⚑", Some(ICON_RED)),
            ("⚑", Some(ICON_YELLOW)),
            ("⚑", Some(ICON_GREEN)),
        ],
        "3Arrows" => vec![
            ("↓", Some(ICON_RED)),
            ("→", Some(ICON_YELLOW)),
            ("↑", Some(ICON_GREEN)),
        ],
        "3ArrowsGray" => vec![
            ("↓", Some(ICON_GRAY)),
            ("→", Some(ICON_GRAY)),
            ("↑", Some(ICON_GRAY)),
        ],
        "4Arrows" => vec![
            ("↓", Some(ICON_RED)),
            ("↘", Some(ICON_YELLOW)),
            ("↗", Some(ICON_YELLOW)),
            ("↑", Some(ICON_GREEN)),
        ],
        "4ArrowsGray" => vec![
            ("↓", Some(ICON_GRAY)),
            ("↘", Some(ICON_GRAY)),
            ("↗", Some(ICON_GRAY)),
            ("↑", Some(ICON_GRAY)),
        ],
        "5Arrows" => vec![
            ("↓", Some(ICON_RED)),
            ("↘", Some(ICON_YELLOW)),
            ("→", Some(ICON_YELLOW)),
            ("↗", Some(ICON_YELLOW)),
            ("↑", Some(ICON_GREEN)),
        ],
        "5ArrowsGray" => vec![
            ("↓", Some(ICON_GRAY)),
            ("↘", Some(ICON_GRAY)),
            ("→", Some(ICON_GRAY)),
            ("↗", Some(ICON_GRAY)),
            ("↑", Some(ICON_GRAY)),
        ],
        _ => {
            if band_count >= 5 {
                vec![
                    ("⇊", None),
                    ("↓", None),
                    ("→", None),
                    ("↑", None),
                    ("⇈", None),
                ]
            } else {
                vec![("↓", None), ("→", None), ("↑", None)]
            }
        }
    }
}

/// Build a map of conditional formatting overrides for all cells in the sheet.
pub(crate) fn build_cond_fmt_overrides(
    sheet: &umya_spreadsheet::Worksheet,
    raw_hints: Option<&RawCondFmtHints>,
) -> HashMap<(u32, u32), CondFmtOverride> {
    let mut overrides: HashMap<CellPos, CondFmtOverride> = HashMap::new();

    for cf in sheet.get_conditional_formatting_collection() {
        let sqref = cf.get_sequence_of_references().get_sqref();
        let ranges: Vec<CellRange> = parse_sqref(&sqref);
        if ranges.is_empty() {
            continue;
        }

        for rule in cf.get_conditional_collection() {
            use umya_spreadsheet::ConditionalFormatValues;
            let raw_hint = raw_hints.and_then(|hints| hints.get(rule.get_priority()));

            match rule.get_type() {
                ConditionalFormatValues::CellIs => {
                    apply_cell_is_rule(sheet, rule, &ranges, &mut overrides);
                }
                ConditionalFormatValues::ContainsText
                | ConditionalFormatValues::NotContainsText
                | ConditionalFormatValues::BeginsWith
                | ConditionalFormatValues::EndsWith => {
                    apply_text_rule(sheet, rule, &ranges, &mut overrides);
                }
                ConditionalFormatValues::ColorScale => {
                    apply_color_scale_rule(sheet, rule, &ranges, &mut overrides);
                }
                ConditionalFormatValues::DataBar => {
                    apply_data_bar_rule(sheet, rule, &ranges, &mut overrides, raw_hint);
                }
                ConditionalFormatValues::IconSet => {
                    apply_icon_set_rule(sheet, rule, &ranges, &mut overrides, raw_hint);
                }
                _ => {}
            }
        }
    }

    overrides
}

/// Determine which icon index a value falls into based on thresholds.
fn evaluate_icon_index(val: f64, thresholds: &[f64], num_icons: usize) -> usize {
    if num_icons == 0 {
        return 0;
    }
    // Iterate thresholds from highest to lowest
    for i in (1..thresholds.len()).rev() {
        if val >= thresholds[i] {
            return (i).min(num_icons - 1);
        }
    }
    0
}

#[cfg(test)]
#[path = "cond_fmt_tests.rs"]
mod tests;
