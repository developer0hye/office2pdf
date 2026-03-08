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
