use super::*;

/// Generate Typst markup for a list (ordered or unordered).
///
/// Uses Typst's `#enum()` for ordered lists and `#list()` for unordered lists.
/// Nested items are wrapped in `list.item()` / `enum.item()` with a sub-list.
struct EffectiveListStyle<'a> {
    kind: ListKind,
    numbering_pattern: Option<&'a str>,
    full_numbering: bool,
    marker_text: Option<&'a str>,
    marker_style: Option<&'a TextStyle>,
}

fn list_style_for_level<'a>(list: &'a List, level: u32) -> EffectiveListStyle<'a> {
    if let Some(style) = list.level_styles.get(&level) {
        EffectiveListStyle {
            kind: style.kind,
            numbering_pattern: style.numbering_pattern.as_deref(),
            full_numbering: style.full_numbering,
            marker_text: style.marker_text.as_deref(),
            marker_style: style.marker_style.as_ref(),
        }
    } else {
        EffectiveListStyle {
            kind: list.kind,
            numbering_pattern: None,
            full_numbering: false,
            marker_text: None,
            marker_style: None,
        }
    }
}

fn list_funcs(kind: ListKind) -> (&'static str, &'static str) {
    match kind {
        ListKind::Ordered => ("enum", "enum.item"),
        ListKind::Unordered => ("list", "list.item"),
    }
}

fn write_list_open(
    out: &mut String,
    prefix: &str,
    style: &EffectiveListStyle<'_>,
    start_at: Option<u32>,
) {
    let (func, _) = list_funcs(style.kind);
    let _ = write!(out, "{prefix}{func}(");

    if style.kind == ListKind::Ordered {
        if style.marker_style.is_some_and(has_text_properties) {
            write_ordered_list_numbering_function(out, style);
            out.push_str(", ");
        } else if let Some(numbering_pattern) = style.numbering_pattern {
            let _ = write!(
                out,
                "numbering: \"{}\", ",
                escape_typst_string(numbering_pattern)
            );
        }
        if let Some(start_at) = start_at {
            let _ = write!(out, "start: {start_at}, ");
        }
        if style.full_numbering {
            out.push_str("full: true, ");
        }
    } else if style.marker_text.is_some() || style.marker_style.is_some() {
        out.push_str("marker: [");
        write_unordered_list_marker_content(out, style);
        out.push_str("], ");
    }

    out.push('\n');
}

fn write_ordered_list_numbering_function(out: &mut String, style: &EffectiveListStyle<'_>) {
    let pattern: &str = style.numbering_pattern.unwrap_or("1.");
    out.push_str("numbering: (..nums) => [");
    if let Some(marker_style) = style
        .marker_style
        .filter(|style| has_text_properties(style))
    {
        out.push_str("#text(");
        write_text_params(out, marker_style);
        out.push_str(")[");
    }
    let _ = write!(
        out,
        "#numbering(\"{}\", ..nums)",
        escape_typst_string(pattern)
    );
    if style.marker_style.is_some_and(has_text_properties) {
        out.push(']');
    }
    out.push(']');
}

fn write_unordered_list_marker_content(out: &mut String, style: &EffectiveListStyle<'_>) {
    let (marker_text, marker_style) =
        renderable_unordered_marker(style.marker_text.unwrap_or("•"), style.marker_style);
    if let Some(marker_style) = marker_style
        .as_ref()
        .filter(|style| has_text_properties(style))
    {
        out.push_str("#text(");
        write_text_params(out, marker_style);
        out.push_str(")[");
        out.push_str(&escape_typst(&marker_text));
        out.push(']');
        return;
    }

    out.push_str(&escape_typst(&marker_text));
}

fn list_root_level(list: &List) -> u32 {
    list.items.first().map(|item| item.level).unwrap_or(0)
}

pub(super) fn generate_list(out: &mut String, list: &List) -> Result<(), ConvertError> {
    let root_level: u32 = list_root_level(list);
    let style = list_style_for_level(list, root_level);
    let start_at = list.items.first().and_then(|item| item.start_at);
    write_list_open(out, "#", &style, start_at);
    generate_list_items(out, list, &list.items, root_level)?;
    out.push_str(")\n");
    Ok(())
}

pub(super) fn can_render_fixed_text_list_inline(list: &List) -> bool {
    let Some(first_item) = list.items.first() else {
        return false;
    };
    let root_level: u32 = first_item.level;
    let root_style: EffectiveListStyle<'_> = list_style_for_level(list, root_level);
    if list.kind == ListKind::Unordered && root_style.marker_text == Some("-") {
        return false;
    }
    if first_item.content.len() != 1 {
        return false;
    }

    let first_style: &ParagraphStyle = &first_item.content[0].style;
    list.items.iter().all(|item| {
        item.level == root_level
            && item.content.len() == 1
            && paragraph_styles_match(&item.content[0].style, first_style)
    })
}

fn paragraph_styles_match(left: &ParagraphStyle, right: &ParagraphStyle) -> bool {
    alignment_matches(left.alignment, right.alignment)
        && option_f64_matches(left.indent_left, right.indent_left)
        && option_f64_matches(left.indent_right, right.indent_right)
        && option_f64_matches(left.indent_first_line, right.indent_first_line)
        && line_spacing_matches(left.line_spacing, right.line_spacing)
        && option_f64_matches(left.space_before, right.space_before)
        && option_f64_matches(left.space_after, right.space_after)
        && left.heading_level == right.heading_level
        && left.direction == right.direction
        && tab_stops_match(left.tab_stops.as_deref(), right.tab_stops.as_deref())
}

fn alignment_matches(left: Option<Alignment>, right: Option<Alignment>) -> bool {
    match (left, right) {
        (Some(Alignment::Left), None) | (None, Some(Alignment::Left)) => true,
        _ => left == right,
    }
}

fn option_f64_matches(left: Option<f64>, right: Option<f64>) -> bool {
    match (left, right) {
        (Some(l), Some(r)) => (l - r).abs() < 0.0001,
        (None, None) => true,
        _ => false,
    }
}

fn line_spacing_matches(left: Option<LineSpacing>, right: Option<LineSpacing>) -> bool {
    match (left, right) {
        (Some(LineSpacing::Proportional(l)), Some(LineSpacing::Proportional(r))) => {
            (l - r).abs() < 0.0001
        }
        (Some(LineSpacing::Exact(l)), Some(LineSpacing::Exact(r))) => (l - r).abs() < 0.0001,
        (None, None) => true,
        _ => false,
    }
}

fn tab_stops_match(left: Option<&[TabStop]>, right: Option<&[TabStop]>) -> bool {
    match (left, right) {
        (Some(left_stops), Some(right_stops)) => left_stops == right_stops,
        (None, None) => true,
        _ => false,
    }
}

pub(super) fn generate_fixed_text_list(
    out: &mut String,
    list: &List,
    include_item_spacing: bool,
    available_width_pt: Option<f64>,
) -> Result<(), ConvertError> {
    let paragraph: &Paragraph = &list.items[0].content[0];
    let style: &ParagraphStyle = &paragraph.style;
    let root_level: u32 = list_root_level(list);
    let effective_style: EffectiveListStyle<'_> = list_style_for_level(list, root_level);
    let has_para_style: bool = needs_block_wrapper(style);
    let line_gap_pt: Option<f64> = fixed_text_list_line_gap_pt(style, list);

    if has_para_style {
        out.push_str("#block(");
        write_block_params(out, style);
        out.push_str(")[\n");
        write_fixed_text_list_par_settings(out, style, line_gap_pt);
    }

    let align_str: Option<&str> = fixed_text_list_alignment(style.alignment);
    let mut current_number: u32 = list
        .items
        .first()
        .and_then(|item| item.start_at)
        .unwrap_or(1);
    if available_width_pt.is_some() {
        for (index, item) in list.items.iter().enumerate() {
            if index > 0 {
                out.push('\n');
                if let Some(gap) = line_gap_pt.filter(|gap| *gap > 0.0 && include_item_spacing) {
                    let _ = writeln!(out, "#v({}pt)", format_f64(gap));
                }
                if let Some(start_at) = item.start_at {
                    current_number = start_at;
                }
            }

            let item_paragraph: &Paragraph = &item.content[0];
            let marker_text: String = fixed_text_list_marker(
                list.kind,
                &effective_style,
                current_number,
                &item_paragraph.runs,
            );
            let runs: Vec<Run> = prepend_fixed_text_list_marker_run(
                &item_paragraph.style,
                &effective_style,
                &item_paragraph.runs,
                marker_text,
            );
            write_fixed_text_list_item(out, item_paragraph, &runs, align_str, available_width_pt);
            out.push('\n');

            if list.kind == ListKind::Ordered {
                current_number += 1;
            }
        }
    } else {
        out.push_str("#stack(dir: ttb");
        if let Some(gap) = line_gap_pt.filter(|gap| *gap > 0.0 && include_item_spacing) {
            let _ = write!(out, ", spacing: {}pt", format_f64(gap));
        }
        out.push_str(",\n");

        for (index, item) in list.items.iter().enumerate() {
            if index > 0 {
                out.push_str(",\n");
                if let Some(start_at) = item.start_at {
                    current_number = start_at;
                }
            }

            let item_paragraph: &Paragraph = &item.content[0];
            let marker_text: String = fixed_text_list_marker(
                list.kind,
                &effective_style,
                current_number,
                &item_paragraph.runs,
            );
            let runs: Vec<Run> = prepend_fixed_text_list_marker_run(
                &item_paragraph.style,
                &effective_style,
                &item_paragraph.runs,
                marker_text,
            );
            out.push('[');
            write_fixed_text_list_item(out, item_paragraph, &runs, align_str, available_width_pt);
            out.push(']');

            if list.kind == ListKind::Ordered {
                current_number += 1;
            }
        }

        out.push_str("\n)");
    }
    if has_para_style {
        out.push_str("\n]");
    }
    out.push('\n');
    Ok(())
}

fn fixed_text_list_alignment(alignment: Option<Alignment>) -> Option<&'static str> {
    match alignment {
        Some(Alignment::Center) => Some("center"),
        Some(Alignment::Right) => Some("right"),
        _ => None,
    }
}

fn write_fixed_text_list_item(
    out: &mut String,
    paragraph: &Paragraph,
    runs: &[Run],
    align_str: Option<&str>,
    available_width_pt: Option<f64>,
) {
    let inset: Insets = fixed_text_list_item_inset(&paragraph.style);
    let has_inset: bool = inset.left > 0.0 || inset.right > 0.0;

    out.push_str("#block(width: ");
    if let Some(width_pt) = available_width_pt {
        let _ = write!(out, "{}pt", format_f64(width_pt));
    } else {
        out.push_str("100%");
    }
    if has_inset {
        let _ = write!(out, ", inset: {}", format_insets(&inset));
    }
    out.push_str(")[");

    if let Some(align) = align_str {
        let _ = write!(out, "#align({align})[");
    }

    write_fixed_text_list_item_paragraph(out, &paragraph.style, runs);

    if align_str.is_some() {
        out.push(']');
    }
    out.push(']');
}

fn fixed_text_list_item_inset(style: &ParagraphStyle) -> Insets {
    let left_inset: f64 = if fixed_text_list_hanging_indent_pt(style).is_some() {
        fixed_text_list_marker_origin_pt(style)
    } else {
        style.indent_left.unwrap_or(0.0).max(0.0)
    };
    Insets {
        top: 0.0,
        right: style.indent_right.unwrap_or(0.0).max(0.0),
        bottom: 0.0,
        left: left_inset,
    }
}

fn write_fixed_text_list_item_paragraph(out: &mut String, style: &ParagraphStyle, runs: &[Run]) {
    write_common_text_settings(out, runs, "");
    write_fixed_text_default_par_settings(out, style, runs, "");
    let hanging_indent_pt: Option<f64> = fixed_text_list_hanging_indent_pt(style);
    let tab_stops: Option<Vec<TabStop>> = fixed_text_list_tab_stops(style, hanging_indent_pt);
    if let Some(hanging_indent_pt) = hanging_indent_pt {
        let _ = write!(
            out,
            "#par(hanging-indent: {}pt)[",
            format_f64(hanging_indent_pt)
        );
    } else if let Some(indent) = style.indent_first_line.filter(|value| value.abs() > 0.0001) {
        let _ = write!(
            out,
            "#par(first-line-indent: (amount: {}pt, all: true))[",
            format_f64(indent)
        );
    } else {
        out.push_str("#par[");
    }

    generate_runs_with_tabs(out, runs, tab_stops.as_deref());
    out.push(']');
}

fn fixed_text_list_marker_origin_pt(style: &ParagraphStyle) -> f64 {
    let indent_left: f64 = style.indent_left.unwrap_or(0.0).max(0.0);
    let indent_first_line: f64 = style.indent_first_line.unwrap_or(0.0);

    if indent_first_line < 0.0 {
        (indent_left + indent_first_line).max(0.0)
    } else {
        indent_left
    }
}

fn fixed_text_list_hanging_indent_pt(style: &ParagraphStyle) -> Option<f64> {
    let indent_first_line: f64 = style.indent_first_line.unwrap_or(0.0);
    if indent_first_line >= -0.0001 {
        return None;
    }

    let indent_left: f64 = style.indent_left.unwrap_or(0.0).max(0.0);
    let hanging_indent_pt: f64 = (indent_left - fixed_text_list_marker_origin_pt(style)).max(0.0);
    (hanging_indent_pt > 0.0001).then_some(hanging_indent_pt)
}

fn fixed_text_list_tab_stops(
    style: &ParagraphStyle,
    hanging_indent_pt: Option<f64>,
) -> Option<Vec<TabStop>> {
    let mut tab_stops: Vec<TabStop> = style.tab_stops.clone().unwrap_or_default();

    if let Some(hanging_indent_pt) = hanging_indent_pt
        && !tab_stops
            .iter()
            .any(|stop| (stop.position - hanging_indent_pt).abs() < 0.0001)
    {
        tab_stops.push(TabStop {
            position: hanging_indent_pt,
            alignment: TabAlignment::Left,
            leader: TabLeader::None,
        });
        tab_stops.sort_by(|left, right| left.position.total_cmp(&right.position));
    }

    (!tab_stops.is_empty()).then_some(tab_stops)
}

pub(super) fn write_common_text_settings(out: &mut String, runs: &[Run], indent: &str) {
    let Some(style) = common_text_style(runs) else {
        return;
    };

    out.push_str(indent);
    out.push_str("#set text(");
    write_text_params(out, &style);
    out.push_str(")\n");
}

pub(super) fn write_fixed_text_default_par_settings(
    out: &mut String,
    style: &ParagraphStyle,
    runs: &[Run],
    indent: &str,
) {
    if style.line_spacing.is_some() {
        return;
    }

    let Some(leading_pt) = fixed_text_default_leading_pt(runs) else {
        return;
    };

    out.push_str(indent);
    let _ = writeln!(out, "#set par(leading: {}pt)", format_f64(leading_pt));
}

pub(super) fn common_text_style(runs: &[Run]) -> Option<TextStyle> {
    let mut visible_runs = runs
        .iter()
        .filter(|run| run.footnote.is_none() && !run.text.is_empty());
    let first_style: TextStyle = visible_runs.next()?.style.clone();
    let common_style: TextStyle = visible_runs.fold(first_style, |common, run| {
        intersect_text_style(&common, &run.style)
    });

    has_text_properties(&common_style).then_some(common_style)
}

fn fixed_text_default_leading_pt(runs: &[Run]) -> Option<f64> {
    let font_size_pt: Option<f64> = common_text_style(runs)
        .and_then(|style| style.font_size)
        .or_else(|| {
            runs.iter()
                .filter_map(|run| run.style.font_size)
                .max_by(f64::total_cmp)
        });
    font_size_pt.map(|size| size * 0.65)
}

fn intersect_text_style(left: &TextStyle, right: &TextStyle) -> TextStyle {
    TextStyle {
        font_family: (left.font_family == right.font_family)
            .then(|| left.font_family.clone())
            .flatten(),
        font_size: (left.font_size == right.font_size)
            .then_some(left.font_size)
            .flatten(),
        bold: (left.bold == right.bold).then_some(left.bold).flatten(),
        italic: (left.italic == right.italic)
            .then_some(left.italic)
            .flatten(),
        color: (left.color == right.color).then_some(left.color).flatten(),
        letter_spacing: (left.letter_spacing == right.letter_spacing)
            .then_some(left.letter_spacing)
            .flatten(),
        ..TextStyle::default()
    }
}

fn fixed_text_list_line_gap_pt(style: &ParagraphStyle, list: &List) -> Option<f64> {
    let font_size_pt: f64 = fixed_text_list_font_size_pt(list);
    match style.line_spacing {
        Some(LineSpacing::Proportional(factor)) if factor > 1.0 => {
            Some((font_size_pt * (factor - 1.0)).max(0.0))
        }
        Some(LineSpacing::Exact(points)) => Some((points - font_size_pt).max(0.0)),
        _ => None,
    }
}

fn fixed_text_list_font_size_pt(list: &List) -> f64 {
    let max_explicit_size: Option<f64> = list
        .items
        .iter()
        .flat_map(|item| item.content.iter())
        .flat_map(|paragraph| paragraph.runs.iter())
        .filter_map(|run| run.style.font_size)
        .max_by(f64::total_cmp);
    max_explicit_size.unwrap_or(12.0)
}

fn write_fixed_text_list_par_settings(
    out: &mut String,
    style: &ParagraphStyle,
    line_gap_pt: Option<f64>,
) {
    if let Some(gap) = line_gap_pt.filter(|gap| *gap > 0.0) {
        let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(gap));
    } else {
        write_par_settings(out, style);
        return;
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

fn fixed_text_list_marker(
    kind: ListKind,
    style: &EffectiveListStyle<'_>,
    number: u32,
    runs: &[Run],
) -> String {
    let marker: String = match kind {
        ListKind::Ordered => ordered_marker(style.numbering_pattern.unwrap_or("1."), number),
        ListKind::Unordered => {
            let (marker_text, _) =
                renderable_unordered_marker(style.marker_text.unwrap_or("•"), style.marker_style);
            marker_text
        }
    };
    if first_visible_char_is_whitespace(runs) {
        marker
    } else {
        format!("{marker} ")
    }
}

fn prepend_marker_run(
    runs: &[Run],
    marker_text: String,
    marker_style: Option<&TextStyle>,
) -> Vec<Run> {
    let marker_style: TextStyle = marker_style
        .cloned()
        .or_else(|| runs.first().map(|run| run.style.clone()))
        .unwrap_or_default();
    let mut combined_runs: Vec<Run> = Vec::with_capacity(runs.len() + 1);
    combined_runs.push(Run {
        text: marker_text,
        style: marker_style,
        href: None,
        footnote: None,
    });
    combined_runs.extend_from_slice(runs);
    combined_runs
}

fn prepend_fixed_text_list_marker_run(
    style: &ParagraphStyle,
    list_style: &EffectiveListStyle<'_>,
    runs: &[Run],
    marker_text: String,
) -> Vec<Run> {
    let normalized_marker_style: Option<TextStyle> = if list_style.kind == ListKind::Unordered {
        renderable_unordered_marker(
            list_style.marker_text.unwrap_or("•"),
            list_style.marker_style,
        )
        .1
    } else {
        list_style.marker_style.cloned()
    };
    if fixed_text_list_hanging_indent_pt(style).is_some() {
        return prepend_marker_run(
            runs,
            format!("{marker_text}\t"),
            normalized_marker_style.as_ref(),
        );
    }

    prepend_marker_run(runs, marker_text, normalized_marker_style.as_ref())
}

fn first_visible_char_is_whitespace(runs: &[Run]) -> bool {
    runs.iter()
        .find_map(|run| run.text.chars().next())
        .is_some_and(char::is_whitespace)
}

fn ordered_marker(pattern: &str, number: u32) -> String {
    if pattern.contains('1') {
        return pattern.replacen('1', &number.to_string(), 1);
    }
    if pattern.contains('a') {
        return pattern.replacen('a', &alpha_marker(number, false), 1);
    }
    if pattern.contains('A') {
        return pattern.replacen('A', &alpha_marker(number, true), 1);
    }
    if pattern.contains('i') {
        return pattern.replacen('i', &roman_marker(number, false), 1);
    }
    if pattern.contains('I') {
        return pattern.replacen('I', &roman_marker(number, true), 1);
    }
    format!("{number}.")
}

fn renderable_unordered_marker(
    marker_text: &str,
    marker_style: Option<&TextStyle>,
) -> (String, Option<TextStyle>) {
    let mut normalized_text: String = marker_text.to_string();
    let mut normalized_style: Option<TextStyle> = marker_style.cloned();

    if let Some(font_family) = marker_style.and_then(|style| style.font_family.as_deref())
        && let Some(mapped_text) = map_symbol_font_marker(font_family, marker_text)
    {
        normalized_text = mapped_text.to_string();
        if let Some(style) = normalized_style.as_mut() {
            style.font_family = None;
        }
        if normalized_style
            .as_ref()
            .is_some_and(|style| !has_text_properties(style))
        {
            normalized_style = None;
        }
    }

    (normalized_text, normalized_style)
}

fn map_symbol_font_marker(font_family: &str, marker_text: &str) -> Option<&'static str> {
    let mut chars = marker_text.chars();
    let marker_char = chars.next()?;
    if chars.next().is_some() {
        return None;
    }

    let normalized_family: String = font_family
        .chars()
        .filter(|character| !character.is_whitespace() && *character != '-')
        .flat_map(char::to_lowercase)
        .collect();

    match (normalized_family.as_str(), marker_char) {
        ("wingdings", '\u{00D8}') => Some("➢"),
        ("wingdings", '\u{00E8}') => Some("➔"),
        ("wingdings", '\u{00FB}') => Some("✖"),
        ("wingdings", '\u{00FC}') => Some("✔"),
        ("wingdings", '\u{00FD}') => Some("☒"),
        ("wingdings", '\u{00FE}') => Some("☑"),
        _ => None,
    }
}

fn alpha_marker(mut number: u32, uppercase: bool) -> String {
    let mut chars: Vec<char> = Vec::new();
    while number > 0 {
        let remainder: u8 = ((number - 1) % 26) as u8;
        let base: u8 = if uppercase { b'A' } else { b'a' };
        chars.push((base + remainder) as char);
        number = (number - 1) / 26;
    }
    chars.iter().rev().collect()
}

fn roman_marker(mut number: u32, uppercase: bool) -> String {
    const ROMAN_VALUES: &[(u32, &str)] = &[
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    let mut result: String = String::new();
    for (value, symbol) in ROMAN_VALUES {
        while number >= *value {
            number -= *value;
            result.push_str(symbol);
        }
    }
    if uppercase {
        result
    } else {
        result.to_lowercase()
    }
}

/// Recursively generate list items, grouping consecutive items at the same or deeper level.
fn generate_list_items(
    out: &mut String,
    list: &List,
    items: &[crate::ir::ListItem],
    base_level: u32,
) -> Result<(), ConvertError> {
    let style = list_style_for_level(list, base_level);
    let (_, item_func) = list_funcs(style.kind);
    let mut i = 0;
    while i < items.len() {
        let item = &items[i];
        if item.level == base_level {
            let _ = write!(out, "  {item_func}[");
            for para in &item.content {
                for run in &para.runs {
                    generate_run(out, run);
                }
            }

            let nested_start = i + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].level > base_level {
                nested_end += 1;
            }

            if nested_end > nested_start {
                let nested_style = list_style_for_level(list, base_level + 1);
                let nested_start_at = items[nested_start].start_at;
                write_list_open(out, " #", &nested_style, nested_start_at);
                generate_list_items(out, list, &items[nested_start..nested_end], base_level + 1)?;
                out.push(')');
                i = nested_end;
            } else {
                i += 1;
            }

            out.push_str("],\n");
        } else {
            let _ = write!(out, "  {item_func}[");
            for para in &item.content {
                for run in &para.runs {
                    generate_run(out, run);
                }
            }
            out.push_str("],\n");
            i += 1;
        }
    }
    Ok(())
}
