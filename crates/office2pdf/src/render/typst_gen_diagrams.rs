use super::*;

/// Generate Typst markup for a chart with improved visual representation.
///
/// Renders charts in a bordered box with title header and type-specific
/// visual representation:
/// - Bar/Column: proportional visual bars
/// - Pie: percentage legend table
/// - Line: data table with trend indicators (↑↓→)
/// - Others: standard data table
pub(super) fn generate_chart(out: &mut String, chart: &Chart) {
    // Bar/Column charts render as an axis-scaled plot that mirrors the native
    // PowerPoint/Excel composition (title, gridlines, tick labels, legend).
    if matches!(chart.chart_type, ChartType::Bar | ChartType::Column)
        && !chart.series.is_empty()
        && !chart.categories.is_empty()
    {
        generate_chart_axis(out, chart);
        return;
    }

    let _ = writeln!(
        out,
        "#block(stroke: 1pt + rgb(100, 100, 100), radius: 4pt, inset: 10pt, width: 100%)["
    );

    let type_label: &str = match &chart.chart_type {
        ChartType::Bar => "Bar Chart",
        ChartType::Column => "Column Chart",
        ChartType::Line => "Line Chart",
        ChartType::Pie => "Pie Chart",
        ChartType::Area => "Area Chart",
        ChartType::Scatter => "Scatter Chart",
        ChartType::Other(label) => label.as_str(),
    };

    if let Some(title) = chart.title.as_ref() {
        let escaped: String = escape_typst(title);
        let _ = writeln!(
            out,
            "#align(center)[#text(size: 14pt, weight: \"bold\")[{escaped}]]\n"
        );
    }
    let _ = writeln!(
        out,
        "#align(center)[#text(fill: rgb(100, 100, 100))[_{type_label}_]]\n"
    );

    if chart.series.is_empty() {
        out.push_str("]\n");
        return;
    }

    match &chart.chart_type {
        ChartType::Bar | ChartType::Column => generate_chart_bar(out, chart),
        ChartType::Pie => generate_chart_pie(out, chart),
        ChartType::Line => generate_chart_line(out, chart),
        _ => generate_chart_table(out, chart),
    }

    out.push_str("]\n");
}

/// Series palette matching Office's default accent colors.
const CHART_SERIES_COLORS: [&str; 6] = [
    "rgb(68, 114, 196)",
    "rgb(237, 125, 49)",
    "rgb(165, 165, 165)",
    "rgb(255, 192, 0)",
    "rgb(91, 155, 213)",
    "rgb(112, 173, 71)",
];

/// Format a chart value without floating-point noise (e.g. 8.2000001 → 8.2).
pub(super) fn chart_value_label(value: f64) -> String {
    if value.fract().abs() < 1e-9 {
        return format!("{}", value.round() as i64);
    }
    // Round to at most 4 significant fractional digits, then trim zeros.
    let rounded: f64 = (value * 10_000.0).round() / 10_000.0;
    let mut text: String = format!("{rounded}");
    if text.contains('.') {
        while text.ends_with('0') {
            text.pop();
        }
        if text.ends_with('.') {
            text.pop();
        }
    }
    text
}

/// Choose a "nice" axis maximum and tick step covering `[0, max]`
/// (e.g. max 8.2 → (10, 2), giving ticks 0,2,4,6,8,10).
fn nice_axis(max_value: f64) -> (f64, f64) {
    if max_value <= 0.0 {
        return (1.0, 1.0);
    }
    let magnitude: f64 = 10f64.powf(max_value.log10().floor());
    let normalized: f64 = max_value / magnitude;
    let nice_norm: f64 = if normalized <= 1.0 {
        1.0
    } else if normalized <= 2.0 {
        2.0
    } else if normalized <= 5.0 {
        5.0
    } else {
        10.0
    };
    let nice_max: f64 = nice_norm * magnitude;
    let step: f64 = nice_max / 5.0;
    (nice_max, step)
}

/// Render a bar (horizontal) or column (vertical) chart as an axis-scaled
/// plot with gridlines, tick labels, and a legend.
fn generate_chart_axis(out: &mut String, chart: &Chart) {
    const PLOT_MAIN: f64 = 300.0; // value-axis length in points
    const ROW: f64 = 34.0; // per-category thickness
    const LABEL_W: f64 = 62.0; // category label gutter
    const TICK_GAP: f64 = 22.0; // value tick label gutter
    const GAP: f64 = 6.0;

    let horizontal: bool = matches!(chart.chart_type, ChartType::Bar);
    let categories: usize = chart.categories.len();
    let series: &[crate::ir::ChartSeries] = &chart.series;
    let series_count: usize = series.len().max(1);

    let max_value: f64 = series
        .iter()
        .flat_map(|s| s.values.iter())
        .copied()
        .fold(0.0_f64, f64::max);
    let (nice_max, step) = nice_axis(max_value);
    let plot_cross: f64 = categories as f64 * ROW; // category-axis length

    // Chart-area title: the explicit chart title, else the single series
    // name (which is what the audited fixture carries).
    let area_title: Option<&str> = chart.title.as_deref().or_else(|| {
        if series.len() == 1 {
            series[0].name.as_deref()
        } else {
            None
        }
    });
    if let Some(title) = area_title {
        let _ = writeln!(
            out,
            "#align(center)[#text(size: 11pt, weight: \"bold\")[{}]]",
            escape_typst(title)
        );
        out.push_str("#v(4pt)\n");
    }

    let legend_w: f64 = 78.0;
    let (total_w, total_h) = if horizontal {
        (
            LABEL_W + GAP + PLOT_MAIN + GAP + legend_w,
            plot_cross + TICK_GAP,
        )
    } else {
        (
            TICK_GAP + GAP + plot_cross + GAP + legend_w,
            PLOT_MAIN + ROW,
        )
    };

    let _ = writeln!(
        out,
        "#box(width: {}pt, height: {}pt)[",
        format_f64(total_w),
        format_f64(total_h)
    );

    // Plot-area origin (top-left of the plotting rectangle).
    let (plot_x, plot_y, plot_w, plot_h) = if horizontal {
        (LABEL_W + GAP, 0.0, PLOT_MAIN, plot_cross)
    } else {
        (TICK_GAP + GAP, 0.0, plot_cross, PLOT_MAIN)
    };

    // Gridlines + value tick labels.
    let mut tick: f64 = 0.0;
    while tick <= nice_max + step * 1e-6 {
        let frac: f64 = tick / nice_max;
        if horizontal {
            let x: f64 = plot_x + frac * plot_w;
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, line(end: (0pt, {}pt), stroke: 0.6pt + rgb(200, 200, 200)))",
                format_f64(x),
                format_f64(plot_y),
                format_f64(plot_h)
            );
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, box(width: 24pt)[#align(center)[#text(size: 8pt)[{}]]])",
                format_f64(x - 12.0),
                format_f64(plot_y + plot_h + 4.0),
                chart_value_label(tick)
            );
        } else {
            let y: f64 = plot_y + (1.0 - frac) * plot_h;
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, line(end: ({}pt, 0pt), stroke: 0.6pt + rgb(200, 200, 200)))",
                format_f64(plot_x),
                format_f64(y),
                format_f64(plot_w)
            );
            let _ = writeln!(
                out,
                "#place(top + left, dx: 0pt, dy: {}pt, box(width: {}pt, height: 10pt)[#align(right + horizon)[#text(size: 8pt)[{}]]])",
                format_f64(y - 5.0),
                format_f64(TICK_GAP),
                chart_value_label(tick)
            );
        }
        tick += step;
    }

    // Bars, grouped per category when multiple series are present.
    for (cat_index, category) in chart.categories.iter().enumerate() {
        let group_start: f64 = cat_index as f64 * ROW;
        let sub: f64 = (ROW * 0.7) / series_count as f64;
        for (s_index, s) in series.iter().enumerate() {
            let value: f64 = s.values.get(cat_index).copied().unwrap_or(0.0);
            let frac: f64 = (value / nice_max).clamp(0.0, 1.0);
            let color: &str = CHART_SERIES_COLORS[s_index % CHART_SERIES_COLORS.len()];
            let offset: f64 = ROW * 0.15 + s_index as f64 * sub;
            if horizontal {
                // Bar charts stack categories bottom-up.
                let row_top: f64 = plot_h - (cat_index as f64 + 1.0) * ROW;
                let bar_w: f64 = frac * plot_w;
                let _ = writeln!(
                    out,
                    "#place(top + left, dx: {}pt, dy: {}pt, rect(width: {}pt, height: {}pt, fill: {}, stroke: none))",
                    format_f64(plot_x),
                    format_f64(row_top + offset),
                    format_f64(bar_w.max(0.0)),
                    format_f64(sub),
                    color
                );
            } else {
                let bar_h: f64 = frac * plot_h;
                let _ = writeln!(
                    out,
                    "#place(top + left, dx: {}pt, dy: {}pt, rect(width: {}pt, height: {}pt, fill: {}, stroke: none))",
                    format_f64(plot_x + group_start + offset),
                    format_f64(plot_y + plot_h - bar_h),
                    format_f64(sub),
                    format_f64(bar_h.max(0.0)),
                    color
                );
            }
        }
        // Category label.
        if horizontal {
            let row_top: f64 = plot_h - (cat_index as f64 + 1.0) * ROW;
            let _ = writeln!(
                out,
                "#place(top + left, dx: 0pt, dy: {}pt, box(width: {}pt, height: {}pt)[#align(right + horizon)[#text(size: 9pt)[{}]]])",
                format_f64(row_top),
                format_f64(LABEL_W),
                format_f64(ROW),
                escape_typst(category)
            );
        } else {
            let _ = writeln!(
                out,
                "#place(top + left, dx: {}pt, dy: {}pt, box(width: {}pt, height: {}pt)[#align(center + horizon)[#text(size: 9pt)[{}]]])",
                format_f64(plot_x + group_start),
                format_f64(plot_y + plot_h + 2.0),
                format_f64(ROW),
                format_f64(ROW),
                escape_typst(category)
            );
        }
    }

    // Axis baseline (value = 0 line).
    if horizontal {
        let _ = writeln!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt, line(end: (0pt, {}pt), stroke: 0.8pt + rgb(120, 120, 120)))",
            format_f64(plot_x),
            format_f64(plot_y),
            format_f64(plot_h)
        );
    } else {
        let _ = writeln!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt, line(end: ({}pt, 0pt), stroke: 0.8pt + rgb(120, 120, 120)))",
            format_f64(plot_x),
            format_f64(plot_y + plot_h),
            format_f64(plot_w)
        );
    }

    // Legend on the right, vertically centered.
    let legend_x: f64 = plot_x + plot_w + GAP;
    let legend_h: f64 = series_count as f64 * 14.0;
    let legend_y: f64 = (plot_h - legend_h).max(0.0) / 2.0;
    for (s_index, s) in series.iter().enumerate() {
        let color: &str = CHART_SERIES_COLORS[s_index % CHART_SERIES_COLORS.len()];
        let default_name: String = format!("Series {}", s_index + 1);
        let name: &str = s.name.as_deref().unwrap_or(&default_name);
        let _ = writeln!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt, box[#box(width: 9pt, height: 9pt, fill: {}) #text(size: 9pt)[{}]])",
            format_f64(legend_x),
            format_f64(legend_y + s_index as f64 * 14.0),
            color,
            escape_typst(name)
        );
    }

    out.push_str("]\n");
}

fn generate_chart_bar(out: &mut String, chart: &Chart) {
    let max_value: f64 = chart
        .series
        .iter()
        .flat_map(|series| series.values.iter())
        .copied()
        .fold(0.0_f64, f64::max);
    let max_value: f64 = if max_value == 0.0 { 1.0 } else { max_value };

    let colors: [&str; 4] = [
        "rgb(66, 133, 244)",
        "rgb(219, 68, 55)",
        "rgb(244, 180, 0)",
        "rgb(15, 157, 88)",
    ];

    for (row_index, category) in chart.categories.iter().enumerate() {
        let escaped_category: String = escape_typst(category);
        let _ = writeln!(out, "#text(weight: \"bold\")[{escaped_category}]");
        for (series_index, series) in chart.series.iter().enumerate() {
            let value: f64 = series.values.get(row_index).copied().unwrap_or(0.0);
            let percent: u32 = (value / max_value * 100.0).round().min(100.0) as u32;
            let color: &str = colors[series_index % colors.len()];
            let _ = writeln!(
                out,
                "#box(width: {percent}%, height: 14pt, fill: {color}, radius: 2pt)[#text(size: 8pt, fill: white)[ {}]]",
                format_f64(value)
            );
        }
        let _ = writeln!(out);
    }

    if chart.series.len() > 1 {
        let _ = writeln!(out);
        for (index, series) in chart.series.iter().enumerate() {
            let default_name: String = format!("Series {}", index + 1);
            let name: &str = series.name.as_deref().unwrap_or(&default_name);
            let color: &str = colors[index % colors.len()];
            let _ = writeln!(
                out,
                "#box(width: 10pt, height: 10pt, fill: {color}) #text(size: 9pt)[{name}] "
            );
        }
    }
}

fn generate_chart_pie(out: &mut String, chart: &Chart) {
    let Some(series) = chart.series.first() else {
        return;
    };

    let total: f64 = series.values.iter().sum();
    let total: f64 = if total == 0.0 { 1.0 } else { total };

    let colors: [&str; 6] = [
        "rgb(66, 133, 244)",
        "rgb(219, 68, 55)",
        "rgb(244, 180, 0)",
        "rgb(15, 157, 88)",
        "rgb(171, 71, 188)",
        "rgb(0, 172, 193)",
    ];

    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: 3,");
    let _ = writeln!(out, "  [*Slice*], [*Value*], [*%*],");

    for (index, category) in chart.categories.iter().enumerate() {
        let value: f64 = series.values.get(index).copied().unwrap_or(0.0);
        let percent: f64 = value / total * 100.0;
        let escaped_category: String = escape_typst(category);
        let color: &str = colors[index % colors.len()];
        let _ = writeln!(
            out,
            "  [#box(width: 8pt, height: 8pt, fill: {color}) {escaped_category}], [{}], [{:.1}%],",
            format_f64(value),
            percent
        );
    }

    let _ = writeln!(out, ")\n");
}

fn generate_chart_line(out: &mut String, chart: &Chart) {
    let column_count: usize = 1 + chart.series.len();
    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: {column_count},");

    out.push_str("  [*Category*], ");
    for (index, series) in chart.series.iter().enumerate() {
        let default_name: String = format!("Series {}", index + 1);
        let name: &str = series.name.as_deref().unwrap_or(&default_name);
        let _ = write!(out, "[*{name}*]");
        if index + 1 < chart.series.len() {
            out.push_str(", ");
        }
    }
    out.push_str(",\n");

    for (row_index, category) in chart.categories.iter().enumerate() {
        let escaped_category: String = escape_typst(category);
        let _ = write!(out, "  [{escaped_category}], ");
        for (series_index, series) in chart.series.iter().enumerate() {
            let value: f64 = series.values.get(row_index).copied().unwrap_or(0.0);
            let trend: &str = if row_index > 0 {
                let previous: f64 = series.values.get(row_index - 1).copied().unwrap_or(0.0);
                if value > previous {
                    " ↑"
                } else if value < previous {
                    " ↓"
                } else {
                    " →"
                }
            } else {
                ""
            };
            let _ = write!(out, "[{}{}]", format_f64(value), trend);
            if series_index + 1 < chart.series.len() {
                out.push_str(", ");
            }
        }
        out.push_str(",\n");
    }

    let _ = writeln!(out, ")\n");
}

fn generate_chart_table(out: &mut String, chart: &Chart) {
    let column_count: usize = 1 + chart.series.len();
    let _ = writeln!(out, "#table(");
    let _ = writeln!(out, "  columns: {column_count},");

    out.push_str("  [*Category*], ");
    for (index, series) in chart.series.iter().enumerate() {
        let default_name: String = format!("Series {}", index + 1);
        let name: &str = series.name.as_deref().unwrap_or(&default_name);
        let _ = write!(out, "[*{name}*]");
        if index + 1 < chart.series.len() {
            out.push_str(", ");
        }
    }
    out.push_str(",\n");

    for (row_index, category) in chart.categories.iter().enumerate() {
        let escaped_category: String = escape_typst(category);
        let _ = write!(out, "  [{escaped_category}], ");
        for (index, series) in chart.series.iter().enumerate() {
            let value: f64 = series.values.get(row_index).copied().unwrap_or(0.0);
            let _ = write!(out, "[{}]", format_f64(value));
            if index + 1 < chart.series.len() {
                out.push_str(", ");
            }
        }
        out.push_str(",\n");
    }

    let _ = writeln!(out, ")\n");
}

/// Generate Typst markup for a SmartArt diagram.
///
/// Renders SmartArt as a visually distinct bordered box with:
/// - Hierarchy items (varying depths): indented tree with depth-based padding
/// - Flat items (all same depth): numbered steps with arrows
pub(super) fn generate_smartart(out: &mut String, smartart: &SmartArt, width: f64, height: f64) {
    let _ = writeln!(
        out,
        "#block(width: {}pt, height: {}pt, stroke: 1pt + rgb(70, 130, 180), radius: 4pt, inset: 10pt, fill: rgb(245, 248, 255))[",
        format_f64(width),
        format_f64(height),
    );
    let _ = writeln!(
        out,
        "#align(center)[#text(size: 11pt, weight: \"bold\", fill: rgb(70, 130, 180))[SmartArt Diagram]]\n"
    );

    if smartart.items.is_empty() {
        out.push_str("]\n");
        return;
    }

    let has_hierarchy: bool = smartart.items.iter().any(|node| node.depth > 0);

    if has_hierarchy {
        generate_smartart_hierarchy(out, smartart);
    } else {
        generate_smartart_steps(out, smartart);
    }

    out.push_str("]\n");
}

fn generate_smartart_hierarchy(out: &mut String, smartart: &SmartArt) {
    for node in &smartart.items {
        let escaped: String = escape_typst(&node.text);
        if node.depth == 0 {
            let _ = writeln!(out, "#text(weight: \"bold\")[{escaped}]");
        } else {
            let indent: f64 = node.depth as f64 * 16.0;
            let branch: &str = if node.depth == 1 { "├" } else { "└" };
            let _ = writeln!(
                out,
                "#pad(left: {}pt)[{branch} {escaped}]",
                format_f64(indent),
            );
        }
    }
}

fn generate_smartart_steps(out: &mut String, smartart: &SmartArt) {
    for (index, node) in smartart.items.iter().enumerate() {
        let escaped: String = escape_typst(&node.text);
        let step_number: usize = index + 1;
        let _ = writeln!(
            out,
            "#box(stroke: 0.5pt + rgb(70, 130, 180), radius: 3pt, inset: 6pt)[#text(weight: \"bold\")[{}. ] {escaped}]",
            step_number,
        );
        if index + 1 < smartart.items.len() {
            let _ = writeln!(out, "#align(center)[#text(size: 14pt)[↓]]");
        }
    }
}

#[cfg(test)]
mod chart_value_label_tests {
    use super::{chart_value_label, nice_axis};

    #[test]
    fn formats_without_float_noise() {
        assert_eq!(chart_value_label(8.200000000000001), "8.2");
        assert_eq!(chart_value_label(3.0), "3");
        assert_eq!(chart_value_label(0.0), "0");
        assert_eq!(chart_value_label(1234.5), "1234.5");
        assert_eq!(chart_value_label(0.333333333), "0.3333");
    }

    #[test]
    fn nice_axis_rounds_up() {
        assert_eq!(nice_axis(8.2), (10.0, 2.0));
        assert_eq!(nice_axis(3.2), (5.0, 1.0));
        assert_eq!(nice_axis(45.0), (50.0, 10.0));
        assert_eq!(nice_axis(0.0), (1.0, 1.0));
    }
}
