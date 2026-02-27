use std::collections::HashMap;

use crate::ir::{Color, DataBarInfo};

/// A (column, row) coordinate pair (1-indexed).
type CellPos = (u32, u32);

/// A conditional formatting override for a specific cell.
pub(crate) struct CondFmtOverride {
    pub background: Option<Color>,
    pub font_color: Option<Color>,
    pub bold: Option<bool>,
    pub data_bar: Option<DataBarInfo>,
    pub icon_text: Option<String>,
}

/// A cell range within a sheet (1-indexed, inclusive).
#[derive(Debug, Clone, Copy)]
struct CellRange {
    start_col: u32,
    start_row: u32,
    end_col: u32,
    end_row: u32,
}

/// Parse an Excel column letter string (e.g., "A", "B", "AA") into a 1-indexed column number.
fn parse_column_letters(s: &str) -> Option<u32> {
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
fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let s = s.replace('$', "");
    let split_pos = s.find(|c: char| c.is_ascii_digit())?;
    let col_str = &s[..split_pos];
    let row_str = &s[split_pos..];
    let col = parse_column_letters(col_str)?;
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
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

/// Parse an ARGB hex string (e.g. "FFFF0000") into an IR Color.
fn parse_argb_color(argb: &str) -> Option<Color> {
    if argb.len() < 8 {
        return None;
    }
    let r = u8::from_str_radix(&argb[2..4], 16).ok()?;
    let g = u8::from_str_radix(&argb[4..6], 16).ok()?;
    let b = u8::from_str_radix(&argb[6..8], 16).ok()?;
    Some(Color::new(r, g, b))
}

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
    let mut result = CondFmtOverride {
        background: None,
        font_color: None,
        bold: None,
        data_bar: None,
        icon_text: None,
    };

    if let Some(style) = rule.get_style() {
        if let Some(bg) = style.get_background_color() {
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

/// Build a map of conditional formatting overrides for all cells in the sheet.
pub(crate) fn build_cond_fmt_overrides(
    sheet: &umya_spreadsheet::Worksheet,
) -> HashMap<(u32, u32), CondFmtOverride> {
    let mut overrides: HashMap<CellPos, CondFmtOverride> = HashMap::new();

    for cf in sheet.get_conditional_formatting_collection() {
        let sqref = cf.get_sequence_of_references().get_sqref();
        let ranges = parse_sqref(&sqref);
        if ranges.is_empty() {
            continue;
        }

        for rule in cf.get_conditional_collection() {
            let rule_type = rule.get_type();
            use umya_spreadsheet::ConditionalFormatValues;

            match rule_type {
                ConditionalFormatValues::CellIs => {
                    let operator = rule.get_operator();
                    let fmt = extract_cond_fmt_style(rule);

                    for range in &ranges {
                        for row in range.start_row..=range.end_row {
                            for col in range.start_col..=range.end_col {
                                if let Some(cell) = sheet.get_cell((col, row))
                                    && let Some(val) = cell_numeric_value(cell)
                                    && evaluate_cell_is_rule(val, operator, rule)
                                {
                                    let entry =
                                        overrides.entry((col, row)).or_insert(CondFmtOverride {
                                            background: None,
                                            font_color: None,
                                            bold: None,
                                            data_bar: None,
                                            icon_text: None,
                                        });
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
                ConditionalFormatValues::ColorScale => {
                    if let Some(cs) = rule.get_color_scale() {
                        let colors: Vec<Option<Color>> = cs
                            .get_color_collection()
                            .iter()
                            .map(parse_umya_color_argb)
                            .collect();

                        if colors.len() < 2 {
                            continue;
                        }

                        let numeric_vals = collect_numeric_values_in_ranges(sheet, &ranges);
                        if numeric_vals.is_empty() {
                            continue;
                        }

                        let min_val = numeric_vals.iter().cloned().fold(f64::INFINITY, f64::min);
                        let max_val = numeric_vals
                            .iter()
                            .cloned()
                            .fold(f64::NEG_INFINITY, f64::max);
                        let val_range = max_val - min_val;

                        let color_min = colors[0].unwrap_or(Color::white());
                        let color_max = colors[colors.len() - 1].unwrap_or(Color::black());

                        for range in &ranges {
                            for row in range.start_row..=range.end_row {
                                for col in range.start_col..=range.end_col {
                                    if let Some(cell) = sheet.get_cell((col, row))
                                        && let Some(val) = cell_numeric_value(cell)
                                    {
                                        let ratio = if val_range.abs() < f64::EPSILON {
                                            0.5
                                        } else {
                                            (val - min_val) / val_range
                                        };

                                        let color = if colors.len() == 3 {
                                            let color_mid =
                                                colors[1].unwrap_or(Color::new(255, 255, 0));
                                            if ratio <= 0.5 {
                                                interpolate_color(color_min, color_mid, ratio * 2.0)
                                            } else {
                                                interpolate_color(
                                                    color_mid,
                                                    color_max,
                                                    (ratio - 0.5) * 2.0,
                                                )
                                            }
                                        } else {
                                            interpolate_color(color_min, color_max, ratio)
                                        };

                                        let entry = overrides.entry((col, row)).or_insert(
                                            CondFmtOverride {
                                                background: None,
                                                font_color: None,
                                                bold: None,
                                                data_bar: None,
                                                icon_text: None,
                                            },
                                        );
                                        entry.background = Some(color);
                                    }
                                }
                            }
                        }
                    }
                }
                ConditionalFormatValues::DataBar => {
                    if let Some(db) = rule.get_data_bar() {
                        let bar_color = db
                            .get_color_collection()
                            .first()
                            .and_then(parse_umya_color_argb)
                            .unwrap_or(Color::new(0x63, 0x8E, 0xC6)); // default blue

                        let numeric_vals = collect_numeric_values_in_ranges(sheet, &ranges);
                        if numeric_vals.is_empty() {
                            continue;
                        }

                        let min_val = numeric_vals.iter().cloned().fold(f64::INFINITY, f64::min);
                        let max_val = numeric_vals
                            .iter()
                            .cloned()
                            .fold(f64::NEG_INFINITY, f64::max);
                        let val_range = max_val - min_val;

                        for range in &ranges {
                            for row in range.start_row..=range.end_row {
                                for col in range.start_col..=range.end_col {
                                    if let Some(cell) = sheet.get_cell((col, row))
                                        && let Some(val) = cell_numeric_value(cell)
                                    {
                                        let pct = if val_range.abs() < f64::EPSILON {
                                            50.0
                                        } else {
                                            ((val - min_val) / val_range) * 100.0
                                        };
                                        let entry = overrides.entry((col, row)).or_insert(
                                            CondFmtOverride {
                                                background: None,
                                                font_color: None,
                                                bold: None,
                                                data_bar: None,
                                                icon_text: None,
                                            },
                                        );
                                        entry.data_bar = Some(DataBarInfo {
                                            color: bar_color,
                                            fill_pct: pct,
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
                ConditionalFormatValues::IconSet => {
                    let numeric_vals = collect_numeric_values_in_ranges(sheet, &ranges);
                    if numeric_vals.is_empty() {
                        continue;
                    }

                    let min_val = numeric_vals.iter().cloned().fold(f64::INFINITY, f64::min);
                    let max_val = numeric_vals
                        .iter()
                        .cloned()
                        .fold(f64::NEG_INFINITY, f64::max);
                    let val_range = max_val - min_val;

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
                    let thresholds = if cfvo_thresholds.len() >= 2 {
                        cfvo_thresholds
                    } else {
                        vec![
                            min_val,
                            min_val + val_range / 3.0,
                            min_val + val_range * 2.0 / 3.0,
                        ]
                    };

                    // Default 3-icon arrows: ↓ (low), → (mid), ↑ (high)
                    let icons: &[&str] = if thresholds.len() >= 5 {
                        &["⇊", "↓", "→", "↑", "⇈"]
                    } else {
                        &["↓", "→", "↑"]
                    };

                    for range in &ranges {
                        for row in range.start_row..=range.end_row {
                            for col in range.start_col..=range.end_col {
                                if let Some(cell) = sheet.get_cell((col, row))
                                    && let Some(val) = cell_numeric_value(cell)
                                {
                                    let icon_idx =
                                        evaluate_icon_index(val, &thresholds, icons.len());
                                    let entry =
                                        overrides.entry((col, row)).or_insert(CondFmtOverride {
                                            background: None,
                                            font_color: None,
                                            bold: None,
                                            data_bar: None,
                                            icon_text: None,
                                        });
                                    entry.icon_text = Some(icons[icon_idx].to_string());
                                }
                            }
                        }
                    }
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
mod tests {
    use super::*;

    #[test]
    fn test_parse_sqref_single_range() {
        let ranges = parse_sqref("A1:C3");
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].start_col, 1);
        assert_eq!(ranges[0].start_row, 1);
        assert_eq!(ranges[0].end_col, 3);
        assert_eq!(ranges[0].end_row, 3);
    }

    #[test]
    fn test_parse_sqref_single_cell() {
        let ranges = parse_sqref("B5");
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].start_col, 2);
        assert_eq!(ranges[0].start_row, 5);
        assert_eq!(ranges[0].end_col, 2);
        assert_eq!(ranges[0].end_row, 5);
    }

    #[test]
    fn test_parse_sqref_multiple_ranges() {
        let ranges = parse_sqref("A1:B2 D4:E5");
        assert_eq!(ranges.len(), 2);
    }

    #[test]
    fn test_interpolate_color_extremes() {
        let white = Color::new(255, 255, 255);
        let red = Color::new(255, 0, 0);

        let at_min = interpolate_color(white, red, 0.0);
        assert_eq!(at_min, white);

        let at_max = interpolate_color(white, red, 1.0);
        assert_eq!(at_max, red);
    }

    #[test]
    fn test_interpolate_color_midpoint() {
        let white = Color::new(255, 255, 255);
        let red = Color::new(255, 0, 0);

        let mid = interpolate_color(white, red, 0.5);
        assert_eq!(mid.r, 255);
        assert_eq!(mid.g, 128);
        assert_eq!(mid.b, 128);
    }
}
