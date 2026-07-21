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

#[derive(Clone, Copy)]
struct ListIndentGeometry {
    marker_origin_pt: f64,
    marker_width_pt: f64,
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
    fallback_marker_style: Option<&TextStyle>,
    indent: Option<ListIndentGeometry>,
    spacing_pt: Option<f64>,
    start_at: Option<u32>,
) {
    let (func, _) = list_funcs(style.kind);
    let _ = write!(out, "{prefix}{func}(");

    if let Some(indent) = indent {
        let _ = write!(
            out,
            "indent: {}pt, body-indent: 0pt, ",
            format_f64(indent.marker_origin_pt)
        );
    }

    if let Some(spacing_pt) = spacing_pt {
        let _ = write!(out, "spacing: {}pt, ", format_f64(spacing_pt));
    }

    if style.kind == ListKind::Ordered {
        let marker_style = merge_marker_style(fallback_marker_style, style.marker_style);
        if marker_style.as_ref().is_some_and(has_text_properties) || indent.is_some() {
            write_ordered_list_numbering_function(out, style, marker_style.as_ref(), indent);
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
    } else if style.marker_text.is_some()
        || style.marker_style.is_some()
        || fallback_marker_style.is_some()
        || indent.is_some()
    {
        let (marker_text, explicit_marker_style) =
            renderable_unordered_marker(style.marker_text.unwrap_or("•"), style.marker_style);
        let marker_style =
            merge_marker_style(fallback_marker_style, explicit_marker_style.as_ref());
        out.push_str("marker: [");
        if let Some(indent) = indent {
            write_marker_box_open(out, indent.marker_width_pt);
        }
        write_unordered_list_marker_content(out, &marker_text, marker_style.as_ref());
        if indent.is_some() {
            out.push_str("])");
        }
        out.push_str("], ");
    }

    out.push('\n');
}

fn write_ordered_list_numbering_function(
    out: &mut String,
    style: &EffectiveListStyle<'_>,
    marker_style: Option<&TextStyle>,
    indent: Option<ListIndentGeometry>,
) {
    let pattern: &str = style.numbering_pattern.unwrap_or("1.");
    out.push_str("numbering: (..nums) => [");
    if let Some(indent) = indent {
        write_marker_box_open(out, indent.marker_width_pt);
    }
    if let Some(marker_style) = marker_style.filter(|style| has_text_properties(style)) {
        out.push_str("#text(");
        write_text_params(out, marker_style);
        out.push_str(")[");
    }
    let _ = write!(
        out,
        "#numbering(\"{}\", ..nums)",
        escape_typst_string(pattern)
    );
    if marker_style.is_some_and(has_text_properties) {
        out.push(']');
    }
    if indent.is_some() {
        out.push_str("])");
    }
    out.push(']');
}

fn write_marker_box_open(out: &mut String, width_pt: f64) {
    // Typst normally sizes the marker column from the glyph. Word instead
    // fixes the body at the numbering level's left indent, so reserve the
    // complete hanging-indent span even for narrow bullets and digits.
    let _ = write!(out, "#box(width: {}pt, align(left)[", format_f64(width_pt));
}

fn write_unordered_list_marker_content(
    out: &mut String,
    marker_text: &str,
    marker_style: Option<&TextStyle>,
) {
    if let Some(marker_style) = marker_style.filter(|style| has_text_properties(style)) {
        out.push_str("#text(");
        write_text_params(out, marker_style);
        out.push_str(")[");
        out.push_str(&escape_typst(marker_text));
        out.push(']');
        return;
    }

    out.push_str(&escape_typst(marker_text));
}

fn list_root_level(list: &List) -> u32 {
    list.items.first().map(|item| item.level).unwrap_or(0)
}

fn paragraph_list_indent(style: &ParagraphStyle) -> Option<ListIndentGeometry> {
    let indent_left = style.indent_left?.max(0.0);
    let first_line_indent = style.indent_first_line?;
    if first_line_indent >= -0.0001 {
        return None;
    }

    let marker_origin_pt = (indent_left + first_line_indent).max(0.0);
    let marker_width_pt = indent_left - marker_origin_pt;
    (marker_width_pt > 0.0001).then_some(ListIndentGeometry {
        marker_origin_pt,
        marker_width_pt,
    })
}

fn common_list_level_indent(
    items: &[crate::ir::ListItem],
    level: u32,
) -> Option<ListIndentGeometry> {
    let mut geometries = items
        .iter()
        .filter(|item| item.level == level)
        .filter_map(|item| item.content.first())
        .map(|paragraph| paragraph_list_indent(&paragraph.style));
    let first = geometries.next()??;

    geometries
        .all(|geometry| {
            geometry.is_some_and(|geometry| {
                f64_approx_eq(geometry.marker_origin_pt, first.marker_origin_pt)
                    && f64_approx_eq(geometry.marker_width_pt, first.marker_width_pt)
            })
        })
        .then_some(first)
}

fn paragraph_space_before(item: &crate::ir::ListItem) -> f64 {
    item.content
        .first()
        .and_then(|paragraph| paragraph.style.space_before)
        .unwrap_or(0.0)
        .max(0.0)
}

fn paragraph_space_after(item: &crate::ir::ListItem) -> f64 {
    item.content
        .last()
        .and_then(|paragraph| paragraph.style.space_after)
        .unwrap_or(0.0)
        .max(0.0)
}

fn paragraph_line_height(paragraph: &Paragraph) -> f64 {
    let font_size = paragraph
        .runs
        .iter()
        .filter_map(|run| run.style.font_size)
        .max_by(f64::total_cmp)
        .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);

    if let Some(line_box) = paragraph.style.line_box {
        return font_size * (line_box.ascent_em + line_box.descent_em);
    }

    match paragraph.style.line_spacing {
        Some(LineSpacing::Proportional(factor)) => font_size * factor.max(0.0),
        Some(LineSpacing::Exact(points)) => points.max(0.0),
        None => font_size,
    }
}

fn list_boundary_spacing(
    previous: &crate::ir::ListItem,
    next: &crate::ir::ListItem,
    metric_leading_pt: Option<f64>,
) -> Option<f64> {
    let previous_paragraph = previous.content.last()?;
    let next_paragraph = next.content.first()?;
    let has_paragraph_spacing = previous_paragraph.style.space_after.is_some()
        || next_paragraph.style.space_before.is_some();
    if !has_paragraph_spacing {
        return None;
    }

    let paragraph_gap = paragraph_space_after(previous).max(paragraph_space_before(next));
    if previous_paragraph.style.line_box.is_some() && next_paragraph.style.line_box.is_some() {
        return Some(paragraph_gap);
    }
    // Under Word metric text edges the enclosing wrapper's leading already
    // tops each line box up to the single-space (or grid) advance; Word adds
    // w:spacing before/after on top of that advance, so the item gap is the
    // leading plus the paragraph gap. Adding a whole line height instead
    // stretched every spaced list by roughly a line per item (issue #384).
    if let Some(leading) = metric_leading_pt
        && paragraph_uses_metric_edges(previous_paragraph)
        && paragraph_uses_metric_edges(next_paragraph)
    {
        return Some(leading + paragraph_gap);
    }

    // An explicit Typst list spacing replaces its automatic paragraph
    // leading. Carry the line box as well as the before/after gap so adding
    // paragraph spacing cannot accidentally make a tight list tighter.
    let line_height =
        paragraph_line_height(previous_paragraph).max(paragraph_line_height(next_paragraph));
    Some(line_height + paragraph_gap)
}

/// Whether the paragraph renders under the wrapper's metric text edges
/// (no explicit line spacing or fixed line box of its own).
fn paragraph_uses_metric_edges(paragraph: &Paragraph) -> bool {
    paragraph.style.line_spacing.is_none() && paragraph.style.line_box.is_none()
}

fn common_list_level_spacing(
    items: &[crate::ir::ListItem],
    level: u32,
    metric_leading_pt: Option<f64>,
) -> Option<f64> {
    let level_items = items
        .iter()
        .filter(|item| item.level == level)
        .collect::<Vec<_>>();
    let mut boundaries = level_items
        .windows(2)
        .map(|pair| list_boundary_spacing(pair[0], pair[1], metric_leading_pt));
    let first = boundaries.next()??;

    (first > 0.0001
        && boundaries.all(|spacing| spacing.is_some_and(|spacing| f64_approx_eq(spacing, first))))
    .then_some(first)
}

fn list_edge_spacing(
    list: &List,
    level: u32,
    metric_leading_pt: Option<f64>,
) -> (Option<f64>, Option<f64>) {
    let first = list.items.iter().find(|item| item.level == level);
    let last = list.items.iter().rev().find(|item| item.level == level);
    let edge = |paragraph: &Paragraph, spacing: Option<f64>| {
        if paragraph.style.line_box.is_some() {
            return Some(spacing.unwrap_or(0.0).max(0.0));
        }
        // Same Word semantics as the item boundaries: before/after extends
        // the line advance, so under metric edges the whitespace is the
        // wrapper leading plus the gap (issue #384).
        if let Some(leading) = metric_leading_pt
            && paragraph_uses_metric_edges(paragraph)
        {
            return spacing
                .map(|spacing| leading + spacing.max(0.0))
                .filter(|spacing| *spacing > 0.0001);
        }
        spacing
            .map(|spacing| paragraph_line_height(paragraph) + spacing.max(0.0))
            .filter(|spacing| *spacing > 0.0001)
    };
    let above = first
        .and_then(|item| item.content.first())
        .and_then(|paragraph| edge(paragraph, paragraph.style.space_before));
    let below = last
        .and_then(|item| item.content.last())
        .and_then(|paragraph| edge(paragraph, paragraph.style.space_after));
    (above, below)
}

fn common_list_line_box(list: &List) -> Option<LineBox> {
    let root_level = list_root_level(list);
    let mut line_boxes = list
        .items
        .iter()
        .filter(|item| item.level == root_level)
        .flat_map(|item| item.content.iter())
        .map(|paragraph| paragraph.style.line_box);
    let first = line_boxes.next()??;
    line_boxes
        .all(|line_box| line_box.is_some_and(|line_box| line_box == first))
        .then_some(first)
}

pub(super) fn generate_list(
    out: &mut String,
    list: &List,
    metric_leading_pt: Option<f64>,
) -> Result<(), ConvertError> {
    let root_level: u32 = list_root_level(list);
    let style = list_style_for_level(list, root_level);
    let fallback_marker_style = common_list_level_text_style(&list.items, root_level);
    let indent = common_list_level_indent(&list.items, root_level);
    let spacing_pt = common_list_level_spacing(&list.items, root_level, metric_leading_pt);
    let (space_before, space_after) = list_edge_spacing(list, root_level, metric_leading_pt);
    let line_box = common_list_line_box(list);
    let start_at = list.items.first().and_then(|item| item.start_at);
    if space_before.is_some() || space_after.is_some() {
        out.push_str("#block(");
        write_block_params(
            out,
            &ParagraphStyle {
                space_before,
                space_after,
                ..ParagraphStyle::default()
            },
        );
        out.push_str(")[\n");
        write_line_box_settings(out, line_box);
    }
    write_list_open(
        out,
        "#",
        &style,
        fallback_marker_style.as_ref(),
        indent,
        spacing_pt,
        start_at,
    );
    generate_list_items(out, list, &list.items, root_level, metric_leading_pt)?;
    out.push_str(")\n");
    if space_before.is_some() || space_after.is_some() {
        out.push_str("]\n");
    }
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
        && both_match(left.indent_left, right.indent_left, f64_approx_eq)
        && both_match(left.indent_right, right.indent_right, f64_approx_eq)
        && both_match(
            left.indent_first_line,
            right.indent_first_line,
            f64_approx_eq,
        )
        && both_match(left.line_spacing, right.line_spacing, line_spacing_eq)
        && left.line_box == right.line_box
        && both_match(left.space_before, right.space_before, f64_approx_eq)
        && both_match(left.space_after, right.space_after, f64_approx_eq)
        && left.heading_level == right.heading_level
        && left.direction == right.direction
        && both_match(
            left.tab_stops.as_deref(),
            right.tab_stops.as_deref(),
            |left_stops, right_stops| left_stops == right_stops,
        )
}

/// Compare two `Option` values: both `None` => true, both `Some` => delegate to `eq_fn`,
/// mismatched `Some`/`None` => false.
fn both_match<T>(left: Option<T>, right: Option<T>, eq_fn: impl FnOnce(T, T) -> bool) -> bool {
    match (left, right) {
        (Some(l), Some(r)) => eq_fn(l, r),
        (None, None) => true,
        _ => false,
    }
}

fn f64_approx_eq(left: f64, right: f64) -> bool {
    (left - right).abs() < 0.0001
}

fn alignment_matches(left: Option<Alignment>, right: Option<Alignment>) -> bool {
    match (left, right) {
        (Some(Alignment::Left), None) | (None, Some(Alignment::Left)) => true,
        _ => left == right,
    }
}

fn line_spacing_eq(left: LineSpacing, right: LineSpacing) -> bool {
    match (left, right) {
        (LineSpacing::Proportional(l), LineSpacing::Proportional(r)) => f64_approx_eq(l, r),
        (LineSpacing::Exact(l), LineSpacing::Exact(r)) => f64_approx_eq(l, r),
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
    let active_gap: Option<f64> = line_gap_pt.filter(|gap| *gap > 0.0 && include_item_spacing);
    let use_stack: bool = available_width_pt.is_none();

    if use_stack {
        out.push_str("#stack(dir: ttb");
        if let Some(gap) = active_gap {
            let _ = write!(out, ", spacing: {}pt", format_f64(gap));
        }
        out.push_str(",\n");
    }

    for (index, item) in list.items.iter().enumerate() {
        if index > 0 {
            if use_stack {
                out.push_str(",\n");
            } else {
                out.push('\n');
                if let Some(gap) = active_gap {
                    let _ = writeln!(out, "#v({}pt)", format_f64(gap));
                }
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

        if use_stack {
            out.push('[');
        }
        write_fixed_text_list_item(
            out,
            item_paragraph,
            &effective_style,
            &marker_text,
            align_str,
            available_width_pt,
        );
        if use_stack {
            out.push(']');
        } else {
            out.push('\n');
        }

        if list.kind == ListKind::Ordered {
            current_number += 1;
        }
    }

    if use_stack {
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
    list_style: &EffectiveListStyle<'_>,
    marker_text: &str,
    align_str: Option<&str>,
    available_width_pt: Option<f64>,
) {
    let inset: Insets = fixed_text_list_item_inset(&paragraph.style);
    let has_inset: bool = inset.left > 0.0 || inset.right > 0.0;
    let hanging_indent_pt: Option<f64> = fixed_text_list_hanging_indent_pt(&paragraph.style);
    let use_marker_grid: bool = list_style.kind == ListKind::Ordered && hanging_indent_pt.is_some();

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

    if use_marker_grid {
        write_fixed_text_ordered_marker_grid(
            out,
            paragraph,
            list_style,
            marker_text,
            hanging_indent_pt.unwrap_or(0.0),
        );
    } else {
        let runs: Vec<Run> = prepend_fixed_text_list_marker_run(
            &paragraph.style,
            list_style,
            &paragraph.runs,
            marker_text.to_string(),
        );
        write_fixed_text_list_item_paragraph(out, &paragraph.style, &runs);
    }

    if align_str.is_some() {
        out.push(']');
    }
    out.push(']');
}

fn write_fixed_text_ordered_marker_grid(
    out: &mut String,
    paragraph: &Paragraph,
    list_style: &EffectiveListStyle<'_>,
    marker_text: &str,
    hanging_indent_pt: f64,
) {
    let normalized_marker_text: String = normalize_fixed_text_ordered_grid_marker(marker_text);
    let marker_run: Run =
        fixed_text_list_marker_run(list_style, &paragraph.runs, normalized_marker_text);
    let mut body_style: ParagraphStyle = paragraph.style.clone();
    body_style.indent_left = None;
    body_style.indent_first_line = None;
    let trimmed_runs: Vec<Run> = trim_fixed_text_list_body_runs(&paragraph.runs);

    let _ = writeln!(
        out,
        "#grid(columns: ({}pt, 1fr), gutter: 0pt,",
        format_f64(hanging_indent_pt),
    );
    out.push('[');
    let _ = write!(
        out,
        "#box(width: {}pt)[#align(right)[",
        format_f64(hanging_indent_pt),
    );
    generate_run(out, &marker_run);
    out.push_str("]]");
    out.push_str("],\n");
    out.push('[');
    write_fixed_text_list_item_paragraph(out, &body_style, &trimmed_runs);
    out.push_str("],\n)");
}

fn normalize_fixed_text_ordered_grid_marker(marker_text: &str) -> String {
    format!("{} ", marker_text.trim_end())
}

fn trim_fixed_text_list_body_runs(runs: &[Run]) -> Vec<Run> {
    let mut trimmed_runs: Vec<Run> = Vec::with_capacity(runs.len());
    let mut is_trimming_leading_whitespace: bool = true;

    for run in runs {
        if run.footnote.is_some() {
            trimmed_runs.push(run.clone());
            continue;
        }

        if !is_trimming_leading_whitespace {
            trimmed_runs.push(run.clone());
            continue;
        }

        let trimmed_text: String = run.text.trim_start_matches(char::is_whitespace).to_string();
        if trimmed_text.is_empty() {
            continue;
        }

        let mut trimmed_run: Run = run.clone();
        trimmed_run.text = trimmed_text;
        trimmed_runs.push(trimmed_run);
        is_trimming_leading_whitespace = false;
    }

    if trimmed_runs.is_empty() {
        runs.to_vec()
    } else {
        trimmed_runs
    }
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
    if style.line_spacing.is_some() || style.line_box.is_some() {
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

fn common_list_level_text_style(items: &[crate::ir::ListItem], level: u32) -> Option<TextStyle> {
    let mut visible_styles = items
        .iter()
        .filter(|item| item.level == level)
        .flat_map(|item| item.content.iter())
        .flat_map(|paragraph| paragraph.runs.iter())
        .filter(|run| run.footnote.is_none() && !run.text.is_empty())
        .map(|run| &run.style);
    let first_style = visible_styles.next()?.clone();
    let common_style = visible_styles.fold(first_style, |common, style| {
        intersect_text_style(&common, style)
    });

    has_text_properties(&common_style).then_some(common_style)
}

fn merge_marker_style(
    fallback: Option<&TextStyle>,
    explicit: Option<&TextStyle>,
) -> Option<TextStyle> {
    let mut merged = fallback.cloned().unwrap_or_default();
    if let Some(explicit) = explicit {
        merged.merge_from(explicit);
    }
    has_text_properties(&merged).then_some(merged)
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
    write_line_box_settings(out, style.line_box);
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

    let marker_run: Run = fixed_text_list_marker_run(list_style, runs, marker_text);
    let mut combined_runs: Vec<Run> = Vec::with_capacity(runs.len() + 1);
    combined_runs.push(marker_run);
    combined_runs.extend_from_slice(runs);
    combined_runs
}

fn fixed_text_list_marker_run(
    list_style: &EffectiveListStyle<'_>,
    runs: &[Run],
    marker_text: String,
) -> Run {
    let normalized_marker_style: Option<TextStyle> = if list_style.kind == ListKind::Unordered {
        renderable_unordered_marker(
            list_style.marker_text.unwrap_or("•"),
            list_style.marker_style,
        )
        .1
    } else {
        list_style.marker_style.cloned()
    };
    let marker_style: TextStyle = normalized_marker_style
        .or_else(|| runs.first().map(|run| run.style.clone()))
        .unwrap_or_default();

    Run {
        text: marker_text,
        style: marker_style,
        href: None,
        footnote: None,
    }
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

    // Symbol-font bullets without their font metadata arrive as Unicode
    // private-use codepoints; map the common Word defaults and fall back to
    // a disc so no tofu box reaches the page.
    if normalized_text
        .chars()
        .any(|ch| ('\u{E000}'..='\u{F8FF}').contains(&ch))
    {
        normalized_text = match normalized_text.chars().next() {
            Some('\u{F0B7}') => "•".to_string(),
            Some('\u{F0A7}') => "▪".to_string(),
            Some('\u{F0D8}') => "➢".to_string(),
            Some('\u{F076}') => "❖".to_string(),
            _ => "•".to_string(),
        };
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
        ("symbol", '\u{F0B7}') => Some("•"),
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

fn write_list_item_content(out: &mut String, item: &crate::ir::ListItem) {
    for para in &item.content {
        for run in &para.runs {
            generate_run(out, run);
        }
    }
}

/// Recursively generate list items, grouping consecutive items at the same or deeper level.
fn generate_list_items(
    out: &mut String,
    list: &List,
    items: &[crate::ir::ListItem],
    base_level: u32,
    metric_leading_pt: Option<f64>,
) -> Result<(), ConvertError> {
    let style = list_style_for_level(list, base_level);
    let (_, item_func) = list_funcs(style.kind);
    let mut i = 0;
    while i < items.len() {
        let item = &items[i];
        let _ = write!(out, "  {item_func}");
        if style.kind == ListKind::Ordered
            && i > 0
            && let Some(start_at) = item.start_at
        {
            let _ = write!(out, "({start_at})");
        }
        out.push('[');
        write_list_item_content(out, item);

        if item.level == base_level {
            let nested_start = i + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].level > base_level {
                nested_end += 1;
            }

            if nested_end > nested_start {
                let nested_style = list_style_for_level(list, base_level + 1);
                let fallback_marker_style =
                    common_list_level_text_style(&items[nested_start..nested_end], base_level + 1);
                // Word indents are absolute from the margin, but a nested
                // Typst list is laid out inside the parent item's body:
                // subtract the parent's text origin so the child marker
                // lands at its absolute position (issue #356).
                let parent_indent = common_list_level_indent(&items[i..=i], base_level);
                let indent =
                    common_list_level_indent(&items[nested_start..nested_end], base_level + 1).map(
                        |mut child| {
                            if let Some(parent) = parent_indent {
                                child.marker_origin_pt = (child.marker_origin_pt
                                    - (parent.marker_origin_pt + parent.marker_width_pt))
                                    .max(0.0);
                            }
                            child
                        },
                    );
                let spacing_pt = common_list_level_spacing(
                    &items[nested_start..nested_end],
                    base_level + 1,
                    metric_leading_pt,
                );
                let nested_start_at = items[nested_start].start_at;
                write_list_open(
                    out,
                    " #",
                    &nested_style,
                    fallback_marker_style.as_ref(),
                    indent,
                    spacing_pt,
                    nested_start_at,
                );
                generate_list_items(
                    out,
                    list,
                    &items[nested_start..nested_end],
                    base_level + 1,
                    metric_leading_pt,
                )?;
                out.push(')');
                i = nested_end;
            } else {
                i += 1;
            }
        } else {
            i += 1;
        }

        out.push_str("],\n");
    }
    Ok(())
}
