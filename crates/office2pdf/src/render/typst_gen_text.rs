use std::borrow::Cow;
use std::fmt::Write;

use unicode_normalization::UnicodeNormalization;

use crate::ir::{ParagraphBorder, ParagraphBorderSide};
use crate::render::font_subst;

use super::*;

/// Word's default tab stop interval (0.5 inch = 36pt).
const DEFAULT_TAB_WIDTH_PT: f64 = 36.0;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';
// `eaLnBrk="0"` disables East Asian auto-breaking between adjacent glyphs.
const WORD_JOINER_CHAR: char = '\u{2060}';
// Empirical Word baseline-to-baseline height for single line spacing is ~1.2em.
const WORD_SINGLE_LINE_HEIGHT_EM: f64 = 1.2;

pub(super) fn generate_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
    let style = &para.style;
    let disable_east_asian_breaks: bool = matches!(style.east_asian_line_break, Some(false));
    let has_para_style = needs_block_wrapper(style);
    let has_outer_pad = write_container_indent_wrapper_start(out, style);
    let needs_inline_justify: bool =
        !has_para_style && matches!(style.alignment, Some(Alignment::Justify));
    let needs_inline_rtl: bool =
        !has_para_style && matches!(style.direction, Some(TextDirection::Rtl));

    if has_para_style {
        out.push_str("#block(");
        write_block_params(out, style, Some(&para.runs));
        out.push_str(")[\n");
        write_par_settings(out, style, Some(&para.runs));
    }

    if needs_inline_rtl {
        out.push_str("#text(dir: rtl)[");
    }
    if needs_inline_justify {
        out.push_str("#par(justify: true)[");
    }

    let alignment = style.alignment;
    let use_align = matches!(
        alignment,
        Some(Alignment::Center) | Some(Alignment::Right) | Some(Alignment::Left)
    );

    if use_align {
        let align_str = match alignment {
            Some(Alignment::Left) => "left",
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            _ => "left",
        };
        let _ = write!(out, "#align({align_str})[");
    }

    if let Some(level) = style.heading_level {
        let _ = write!(out, "#heading(level: {level})[");
        generate_runs_with_tabs(
            out,
            &para.runs,
            style.tab_stops.as_deref(),
            disable_east_asian_breaks,
        );
        out.push(']');
    } else {
        generate_runs_with_tabs(
            out,
            &para.runs,
            style.tab_stops.as_deref(),
            disable_east_asian_breaks,
        );
    }

    if use_align {
        out.push(']');
    }

    if needs_inline_justify {
        out.push(']');
    }
    if needs_inline_rtl {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }
    write_container_indent_wrapper_end(out, has_outer_pad);

    out.push('\n');
    Ok(())
}

pub(super) fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.line_spacing.is_some()
        || style.container.is_some()
}

pub(super) fn write_container_indent_wrapper_start(
    out: &mut String,
    style: &ParagraphStyle,
) -> bool {
    let left = style.indent_left.filter(|value| value.abs() > 0.0001);
    let right = style.indent_right.filter(|value| value.abs() > 0.0001);
    if style.container.is_none() || (left.is_none() && right.is_none()) {
        return false;
    }

    out.push_str("#pad(");
    let mut first = true;
    if let Some(left) = left {
        write_param(out, &mut first, &format!("left: {}pt", format_f64(left)));
    }
    if let Some(right) = right {
        write_param(out, &mut first, &format!("right: {}pt", format_f64(right)));
    }
    out.push_str(")[\n");
    true
}

pub(super) fn write_container_indent_wrapper_end(out: &mut String, has_wrapper: bool) {
    if has_wrapper {
        out.push_str("\n]");
    }
}

pub(super) fn write_block_params(out: &mut String, style: &ParagraphStyle, runs: Option<&[Run]>) {
    let mut first = true;
    let line_gap_pt = paragraph_leading_pt(style, runs).filter(|gap| *gap > 0.0);

    if let Some(above) = style.space_before {
        write_param(out, &mut first, &format!("above: {}pt", format_f64(above)));
    }
    if style.space_after.is_some() || line_gap_pt.is_some() {
        let below = style.space_after.unwrap_or(0.0) + line_gap_pt.unwrap_or(0.0);
        write_param(out, &mut first, &format!("below: {}pt", format_f64(below)));
    }
    if let Some(container) = style.container.as_ref() {
        write_param(out, &mut first, "width: 100%");
        if let Some(background) = container.background {
            write_param(out, &mut first, &format_color(&background));
        }
        if let Some(border) = container.border.as_ref()
            && let Some(stroke) = format_paragraph_border(border)
        {
            write_param(out, &mut first, &stroke);
        }
        if let Some(inset) = container.padding.as_ref() {
            write_param(out, &mut first, &format!("inset: {}", format_insets(inset)));
        }
    }
}

fn format_paragraph_border(border: &ParagraphBorder) -> Option<String> {
    let mut parts: Vec<String> = Vec::new();

    if let Some(ref side) = border.top {
        parts.push(format!("top: {}", format_paragraph_border_side(side)));
    }
    if let Some(ref side) = border.right {
        parts.push(format!("right: {}", format_paragraph_border_side(side)));
    }
    if let Some(ref side) = border.bottom {
        parts.push(format!("bottom: {}", format_paragraph_border_side(side)));
    }
    if let Some(ref side) = border.left {
        parts.push(format!("left: {}", format_paragraph_border_side(side)));
    }

    if parts.is_empty() {
        None
    } else {
        Some(format!("stroke: ({})", parts.join(", ")))
    }
}

fn format_paragraph_border_side(side: &ParagraphBorderSide) -> String {
    let base = format!(
        "{}pt + rgb({}, {}, {})",
        format_f64(side.width),
        side.color.r,
        side.color.g,
        side.color.b,
    );
    match side.style {
        BorderLineStyle::Solid | BorderLineStyle::None => base,
        _ => format!(
            "(paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\")",
            side.color.r,
            side.color.g,
            side.color.b,
            format_f64(side.width),
            border_line_style_to_typst(side.style),
        ),
    }
}

pub(super) fn write_par_settings(out: &mut String, style: &ParagraphStyle, runs: Option<&[Run]>) {
    if let Some(leading_pt) = paragraph_leading_pt(style, runs) {
        let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(leading_pt));
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

pub(super) fn paragraph_leading_pt(style: &ParagraphStyle, runs: Option<&[Run]>) -> Option<f64> {
    match style.line_spacing {
        Some(LineSpacing::Proportional(factor)) => {
            proportional_line_spacing_leading_pt(factor, runs)
        }
        Some(LineSpacing::Exact(line_height_pt)) => {
            exact_line_spacing_leading_pt(line_height_pt, runs)
        }
        None => None,
    }
}

fn proportional_line_spacing_leading_pt(factor: f64, runs: Option<&[Run]>) -> Option<f64> {
    if factor <= 0.0 {
        return None;
    }

    let font_size_pt = paragraph_effective_font_size_pt(runs).unwrap_or(12.0);
    let leading_pt = font_size_pt * ((factor * WORD_SINGLE_LINE_HEIGHT_EM) - 1.0);
    Some(round_two_decimals(leading_pt.max(0.0)))
}

fn exact_line_spacing_leading_pt(line_height_pt: f64, runs: Option<&[Run]>) -> Option<f64> {
    let font_size_pt = paragraph_effective_font_size_pt(runs).unwrap_or(12.0);
    Some(round_two_decimals((line_height_pt - font_size_pt).max(0.0)))
}

fn paragraph_effective_font_size_pt(runs: Option<&[Run]>) -> Option<f64> {
    runs.and_then(|runs| {
        runs.iter()
            .filter_map(|run| run.style.font_size)
            .max_by(f64::total_cmp)
    })
}

fn round_two_decimals(value: f64) -> f64 {
    (value * 100.0).round() / 100.0
}

pub(super) fn generate_runs_with_tabs(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
    disable_east_asian_breaks: bool,
) {
    if !paragraph_contains_tabs(runs) {
        generate_runs(out, runs, disable_east_asian_breaks);
        return;
    }

    let segments: Vec<Vec<Run>> = split_runs_on_tabs(runs);
    out.push_str("#context {\n");

    for (index, segment) in segments.iter().enumerate() {
        let _ = write!(out, "  let tab_segment_{index} = [");
        generate_runs(out, segment, disable_east_asian_breaks);
        out.push_str("]\n");

        if index == 0 {
            out.push_str("  let tab_prefix_0 = tab_segment_0\n");
            continue;
        }

        let _ = writeln!(
            out,
            "  let tab_prefix_width_{index} = measure(tab_prefix_{}).width",
            index - 1
        );
        let _ = writeln!(
            out,
            "  let tab_segment_width_{index} = measure(tab_segment_{index}).width"
        );

        if let Some(anchor_runs) = extract_decimal_anchor_runs(segment) {
            let _ = write!(out, "  let tab_decimal_anchor_{index} = [");
            generate_runs(out, &anchor_runs, disable_east_asian_breaks);
            out.push_str("]\n");
            let _ = writeln!(
                out,
                "  let tab_decimal_width_{index} = measure(tab_decimal_anchor_{index}).width"
            );
        }

        let _ = writeln!(
            out,
            "  let tab_default_remainder_{index} = calc.rem-euclid(tab_prefix_width_{index}.abs.pt(), {})",
            format_f64(DEFAULT_TAB_WIDTH_PT)
        );
        let _ = writeln!(
            out,
            "  let tab_advance_{index} = {}",
            build_tab_advance_expr(index, segment, tab_stops)
        );
        let _ = writeln!(
            out,
            "  let tab_fill_{index} = {}",
            build_tab_fill_expr(index, tab_stops)
        );
        let _ = writeln!(
            out,
            "  let tab_prefix_{index} = [#tab_prefix_{}#tab_fill_{index}#tab_segment_{index}]",
            index - 1
        );
    }

    let _ = writeln!(out, "  tab_prefix_{}", segments.len() - 1);
    out.push('}');
}

fn paragraph_contains_tabs(runs: &[Run]) -> bool {
    runs.iter().any(|run| run.text.contains('\t'))
}

pub(super) fn generate_runs(out: &mut String, runs: &[Run], disable_east_asian_breaks: bool) {
    let mut previous_tail: Option<char> = None;

    for run in runs {
        if disable_east_asian_breaks
            && run.footnote.is_none()
            && previous_tail.is_some_and(is_east_asian_word_char)
            && first_layout_char(&run.text).is_some_and(is_east_asian_word_char)
        {
            out.push(WORD_JOINER_CHAR);
        }

        generate_run(out, run, disable_east_asian_breaks);

        previous_tail = if disable_east_asian_breaks && run.footnote.is_none() {
            last_layout_char(&run.text)
        } else {
            None
        };
    }
}

fn first_layout_char(text: &str) -> Option<char> {
    text.chars().next()
}

fn last_layout_char(text: &str) -> Option<char> {
    text.chars().next_back()
}

fn is_east_asian_word_char(ch: char) -> bool {
    matches!(
        ch as u32,
        0x3040..=0x30FF
            | 0x3400..=0x4DBF
            | 0x4E00..=0x9FFF
            | 0xAC00..=0xD7AF
            | 0xF900..=0xFAFF
            | 0x20000..=0x2EBEF
    )
}

fn suppress_east_asian_breaks(text: &str) -> String {
    let mut output = String::with_capacity(text.len() + 8);
    let mut previous: Option<char> = None;

    for current in text.chars() {
        if previous.is_some_and(is_east_asian_word_char) && is_east_asian_word_char(current) {
            output.push(WORD_JOINER_CHAR);
        }
        output.push(current);
        previous = Some(current);
    }

    output
}

fn split_runs_on_tabs(runs: &[Run]) -> Vec<Vec<Run>> {
    let mut segments: Vec<Vec<Run>> = vec![Vec::new()];

    for run in runs {
        if run.footnote.is_some() || !run.text.contains('\t') {
            if run.footnote.is_some() || !run.text.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(run.clone());
            }
            continue;
        }

        for (index, part) in run.text.split('\t').enumerate() {
            if index > 0 {
                segments.push(Vec::new());
            }

            if !part.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(Run {
                        text: part.to_string(),
                        style: run.style.clone(),
                        href: run.href.clone(),
                        footnote: None,
                    });
            }
        }
    }

    segments
}

fn extract_decimal_anchor_runs(runs: &[Run]) -> Option<Vec<Run>> {
    let visible_text: String = runs
        .iter()
        .filter(|run| run.footnote.is_none())
        .map(|run| run.text.as_str())
        .collect();
    let separator_offset = find_decimal_separator_offset(&visible_text)?;

    let mut anchor_runs: Vec<Run> = Vec::new();
    let mut visible_offset: usize = 0;

    for run in runs {
        if let Some(content) = &run.footnote {
            anchor_runs.push(Run {
                text: String::new(),
                style: run.style.clone(),
                href: run.href.clone(),
                footnote: Some(content.clone()),
            });
            continue;
        }

        let run_end = visible_offset + run.text.len();
        if run_end <= separator_offset {
            if !run.text.is_empty() {
                anchor_runs.push(run.clone());
            }
            visible_offset = run_end;
            continue;
        }

        let offset = separator_offset.saturating_sub(visible_offset);
        if offset > 0 {
            anchor_runs.push(Run {
                text: run.text[..offset].to_string(),
                style: run.style.clone(),
                href: run.href.clone(),
                footnote: None,
            });
        }

        return Some(anchor_runs);
    }

    None
}

fn find_decimal_separator_offset(text: &str) -> Option<usize> {
    let separator = text.char_indices().rev().find(|(offset, ch)| {
        matches!(ch, '.' | ',')
            && has_ascii_digit_before(text, *offset)
            && has_ascii_digit_after(text, *offset + ch.len_utf8())
    })?;

    if is_grouped_integer(
        &text
            .chars()
            .filter(|ch| ch.is_ascii_digit() || matches!(ch, '.' | ','))
            .collect::<String>(),
        separator.1,
    ) {
        return None;
    }

    Some(separator.0)
}

fn has_ascii_digit_before(text: &str, offset: usize) -> bool {
    text[..offset].chars().rev().any(|ch| ch.is_ascii_digit())
}

fn has_ascii_digit_after(text: &str, offset: usize) -> bool {
    text[offset..].chars().any(|ch| ch.is_ascii_digit())
}

fn is_grouped_integer(text: &str, separator: char) -> bool {
    if text
        .chars()
        .any(|ch| matches!(ch, '.' | ',') && ch != separator)
    {
        return false;
    }

    let parts: Vec<&str> = text.split(separator).collect();
    parts.len() > 1
        && parts
            .iter()
            .all(|part| !part.is_empty() && part.chars().all(|ch| ch.is_ascii_digit()))
        && parts[1..].iter().all(|part| part.len() == 3)
}

fn build_tab_advance_expr(index: usize, segment: &[Run], tab_stops: Option<&[TabStop]>) -> String {
    let prefix_width_var = format!("tab_prefix_width_{index}");
    let segment_width_var = format!("tab_segment_width_{index}");
    let decimal_width_var =
        extract_decimal_anchor_runs(segment).map(|_| format!("tab_decimal_width_{index}"));
    let default_expr = build_default_tab_advance_expr(index);

    let Some(tab_stops) = tab_stops else {
        return default_expr;
    };

    if tab_stops.is_empty() {
        return default_expr;
    }

    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = format!(
            "calc.max(0pt, {}pt - {prefix_width_var} - {})",
            format_f64(stop.position),
            tab_alignment_offset_expr(stop, &segment_width_var, decimal_width_var.as_deref())
        );

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ {default_expr} }}");
    expr
}

fn build_tab_fill_expr(index: usize, tab_stops: Option<&[TabStop]>) -> String {
    let Some(tab_stops) = tab_stops else {
        return format!("h(tab_advance_{index})");
    };

    if tab_stops.is_empty() {
        return format!("h(tab_advance_{index})");
    }

    let prefix_width_var = format!("tab_prefix_width_{index}");
    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = tab_fill_content_expr(index, stop.leader);

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ h(tab_advance_{index}) }}");
    expr
}

fn tab_fill_content_expr(index: usize, leader: TabLeader) -> String {
    let leader_markup = match leader {
        TabLeader::None => return format!("h(tab_advance_{index})"),
        TabLeader::Dot => ".",
        TabLeader::Hyphen => "-",
        TabLeader::Underscore => "\\_",
    };

    format!("box(width: tab_advance_{index}, repeat[{leader_markup}])")
}

fn build_default_tab_advance_expr(index: usize) -> String {
    format!(
        "if tab_default_remainder_{index} == 0 {{ {}pt }} else {{ ({} - tab_default_remainder_{index}) * 1pt }}",
        format_f64(DEFAULT_TAB_WIDTH_PT),
        format_f64(DEFAULT_TAB_WIDTH_PT)
    )
}

fn tab_alignment_offset_expr(
    stop: &TabStop,
    segment_width_var: &str,
    decimal_width_var: Option<&str>,
) -> String {
    match stop.alignment {
        TabAlignment::Left => "0pt".to_string(),
        TabAlignment::Center => format!("{segment_width_var} / 2"),
        TabAlignment::Right => segment_width_var.to_string(),
        TabAlignment::Decimal => decimal_width_var.unwrap_or(segment_width_var).to_string(),
    }
}

pub(super) fn generate_run(out: &mut String, run: &Run, disable_east_asian_breaks: bool) {
    if let Some(ref content) = run.footnote {
        let escaped_content = escape_typst(content);
        let _ = write!(out, "#footnote[{escaped_content}]");
        return;
    }

    if run_contains_line_break(&run.text) {
        write_run_with_line_breaks(out, run, disable_east_asian_breaks);
        return;
    }

    write_run_segment(out, run, &run.text, true, disable_east_asian_breaks);
}

fn run_contains_line_break(text: &str) -> bool {
    text.chars()
        .any(|character| is_inline_line_break(character))
}

fn is_inline_line_break(character: char) -> bool {
    character == PPTX_SOFT_LINE_BREAK_CHAR || character == '\n' || character == '\r'
}

fn write_run_with_line_breaks(out: &mut String, run: &Run, disable_east_asian_breaks: bool) {
    let mut segment_start: usize = 0;
    let mut chars = run.text.char_indices().peekable();

    while let Some((offset, character)) = chars.next() {
        if !is_inline_line_break(character) {
            continue;
        }

        if segment_start < offset {
            write_run_segment(
                out,
                run,
                &run.text[segment_start..offset],
                true,
                disable_east_asian_breaks,
            );
        }
        out.push_str("#linebreak()");

        // Treat CRLF as one logical line break.
        if character == '\r'
            && let Some((next_offset, '\n')) = chars.peek().copied()
        {
            chars.next();
            segment_start = next_offset + '\n'.len_utf8();
            continue;
        }

        segment_start = offset + character.len_utf8();
    }

    if segment_start < run.text.len() {
        write_run_segment(
            out,
            run,
            &run.text[segment_start..],
            true,
            disable_east_asian_breaks,
        );
    }
}

fn preserve_line_segment_leading_spaces(text: &str) -> Cow<'_, str> {
    let leading_space_count: usize = text.bytes().take_while(|byte| *byte == b' ').count();
    // Keep a single boundary space collapsible, but preserve indentation width.
    if leading_space_count < 2 {
        return Cow::Borrowed(text);
    }

    let mut normalized: String = String::with_capacity(text.len() + leading_space_count);
    for _ in 0..leading_space_count {
        normalized.push('\u{00A0}');
    }
    normalized.push_str(&text[leading_space_count..]);
    Cow::Owned(normalized)
}

fn write_run_segment(
    out: &mut String,
    run: &Run,
    text: &str,
    preserve_leading_spaces: bool,
    disable_east_asian_breaks: bool,
) {
    let normalized_text: Cow<'_, str> = if preserve_leading_spaces {
        preserve_line_segment_leading_spaces(text)
    } else {
        Cow::Borrowed(text)
    };
    let style: TextStyle = effective_text_style_for_content(&run.style, normalized_text.as_ref());
    let layout_text: Cow<'_, str> = if disable_east_asian_breaks {
        Cow::Owned(suppress_east_asian_breaks(normalized_text.as_ref()))
    } else {
        Cow::Borrowed(normalized_text.as_ref())
    };
    let escaped = escape_typst(layout_text.as_ref());

    let has_text_props = has_text_properties(&style);
    let needs_underline = matches!(style.underline, Some(true));
    let needs_strike = matches!(style.strikethrough, Some(true));
    let has_link = run.href.is_some();
    let needs_highlight = style.highlight.is_some();
    let needs_super = matches!(style.vertical_align, Some(VerticalTextAlign::Superscript));
    let needs_sub = matches!(style.vertical_align, Some(VerticalTextAlign::Subscript));
    let needs_small_caps = matches!(style.small_caps, Some(true));
    let needs_all_caps = matches!(style.all_caps, Some(true));

    let escaped: String = if needs_all_caps {
        escape_typst(&layout_text.to_uppercase())
    } else {
        escaped
    };

    if let Some(ref href) = run.href {
        let _ = write!(out, "#link(\"{href}\")[");
    }

    if let Some(ref highlight) = style.highlight {
        let _ = write!(
            out,
            "#highlight(fill: rgb({}, {}, {}))[",
            highlight.r, highlight.g, highlight.b
        );
    }

    if needs_strike {
        out.push_str("#strike[");
    }
    if needs_underline {
        out.push_str("#underline[");
    }
    if needs_super {
        out.push_str("#super[");
    }
    if needs_sub {
        out.push_str("#sub[");
    }
    if needs_small_caps {
        out.push_str("#smallcaps[");
    }

    if has_text_props {
        out.push_str("#text(");
        write_text_params(out, &style);
        out.push_str(")[");
        out.push_str(&escaped);
        out.push(']');
    } else {
        let needs_wrap = !escaped.is_empty()
            && out.ends_with(']')
            && !out.ends_with("\\]")
            && matches!(escaped.as_bytes()[0], b'(' | b'.' | b'[');
        if needs_wrap {
            out.push_str("#[");
            out.push_str(&escaped);
            out.push(']');
        } else {
            out.push_str(&escaped);
        }
    }

    if needs_small_caps {
        out.push(']');
    }
    if needs_sub {
        out.push(']');
    }
    if needs_super {
        out.push(']');
    }
    if needs_underline {
        out.push(']');
    }
    if needs_strike {
        out.push(']');
    }
    if needs_highlight {
        out.push(']');
    }
    if has_link {
        out.push(']');
    }
}

fn effective_text_style_for_content(style: &TextStyle, text: &str) -> TextStyle {
    let mut resolved_style: TextStyle = style.clone();
    if let Some(font_family) = preferred_font_family_for_content(style, text) {
        resolved_style.font_family = Some(font_family);
    }
    resolved_style
}

fn preferred_font_family_for_content(style: &TextStyle, text: &str) -> Option<String> {
    if contains_east_asian_char(text) {
        return style
            .font_family_east_asia
            .clone()
            .or_else(|| style.font_family.clone())
            .or_else(|| style.font_family_ascii.clone())
            .or_else(|| style.font_family_hansi.clone())
            .or_else(|| style.font_family_cs.clone());
    }

    if contains_complex_script_char(text) {
        return style
            .font_family_cs
            .clone()
            .or_else(|| style.font_family.clone())
            .or_else(|| style.font_family_ascii.clone())
            .or_else(|| style.font_family_hansi.clone())
            .or_else(|| style.font_family_east_asia.clone());
    }

    style
        .font_family_ascii
        .clone()
        .or_else(|| style.font_family_hansi.clone())
        .or_else(|| style.font_family.clone())
        .or_else(|| style.font_family_east_asia.clone())
        .or_else(|| style.font_family_cs.clone())
}

fn contains_east_asian_char(text: &str) -> bool {
    text.chars().any(is_east_asian_char)
}

fn is_east_asian_char(character: char) -> bool {
    matches!(
        character,
        '\u{1100}'..='\u{11FF}'
            | '\u{2E80}'..='\u{2EFF}'
            | '\u{2F00}'..='\u{2FDF}'
            | '\u{3000}'..='\u{303F}'
            | '\u{3040}'..='\u{30FF}'
            | '\u{3130}'..='\u{318F}'
            | '\u{31F0}'..='\u{31FF}'
            | '\u{3400}'..='\u{4DBF}'
            | '\u{4E00}'..='\u{9FFF}'
            | '\u{AC00}'..='\u{D7AF}'
            | '\u{F900}'..='\u{FAFF}'
            | '\u{FF00}'..='\u{FFEF}'
            | '\u{20000}'..='\u{2A6DF}'
            | '\u{2A700}'..='\u{2B73F}'
            | '\u{2B740}'..='\u{2B81F}'
            | '\u{2B820}'..='\u{2CEAF}'
            | '\u{2CEB0}'..='\u{2EBEF}'
    )
}

fn contains_complex_script_char(text: &str) -> bool {
    text.chars().any(is_complex_script_char)
}

fn is_complex_script_char(character: char) -> bool {
    matches!(
        character,
        '\u{0590}'..='\u{05FF}'
            | '\u{0600}'..='\u{06FF}'
            | '\u{0750}'..='\u{077F}'
            | '\u{08A0}'..='\u{08FF}'
            | '\u{FB50}'..='\u{FDFF}'
            | '\u{FE70}'..='\u{FEFF}'
    )
}

pub(super) fn has_text_properties(style: &TextStyle) -> bool {
    matches!(style.bold, Some(true))
        || matches!(style.italic, Some(true))
        || style.font_size.is_some()
        || style.color.is_some()
        || style.font_family.is_some()
        || style.letter_spacing.is_some()
}

fn inferred_font_weight(font_family: &str) -> Option<&'static str> {
    let lower = font_family.trim().to_ascii_lowercase();
    if lower.contains("extrabold") || lower.contains("extra bold") {
        Some("extrabold")
    } else if lower.contains("semibold") || lower.contains("semi bold") {
        Some("semibold")
    } else if lower.contains("medium") {
        Some("medium")
    } else if lower.contains("light") {
        Some("light")
    } else {
        None
    }
}

fn font_weight_rank(weight: &str) -> u8 {
    match weight {
        "light" => 1,
        "medium" => 2,
        "semibold" => 3,
        "bold" => 4,
        "extrabold" => 5,
        "black" => 6,
        _ => 0,
    }
}

fn effective_font_weight(style: &TextStyle) -> Option<&'static str> {
    let inferred = style.font_family.as_deref().and_then(inferred_font_weight);
    let explicit = matches!(style.bold, Some(true)).then_some("bold");
    match (explicit, inferred) {
        (Some(explicit), Some(inferred)) => {
            if font_weight_rank(explicit) >= font_weight_rank(inferred) {
                Some(explicit)
            } else {
                Some(inferred)
            }
        }
        (Some(explicit), None) => Some(explicit),
        (None, Some(inferred)) => Some(inferred),
        (None, None) => None,
    }
}

pub(super) fn write_text_params(out: &mut String, style: &TextStyle) {
    let mut first = true;

    if let Some(ref family) = style.font_family {
        let font_value = font_subst::font_with_fallbacks(family);
        write_param(out, &mut first, &format!("font: {font_value}"));
    }
    if let Some(size) = style.font_size {
        write_param(out, &mut first, &format!("size: {}pt", format_f64(size)));
    }
    if let Some(weight) = effective_font_weight(style) {
        write_param(out, &mut first, &format!("weight: \"{weight}\""));
    }
    if matches!(style.italic, Some(true)) {
        write_param(out, &mut first, "style: \"italic\"");
    }
    if let Some(ref color) = style.color {
        write_param(out, &mut first, &format_color(color));
    }
    if let Some(spacing) = style.letter_spacing {
        write_param(
            out,
            &mut first,
            &format!("tracking: {}pt", format_f64(spacing)),
        );
    }
}

pub(super) fn write_param(out: &mut String, first: &mut bool, param: &str) {
    if !*first {
        out.push_str(", ");
    }
    out.push_str(param);
    *first = false;
}

pub(super) fn format_color(color: &Color) -> String {
    format!("fill: rgb({}, {}, {})", color.r, color.g, color.b)
}

pub(super) fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

pub(super) fn escape_typst(text: &str) -> String {
    let normalized_text: String = text.nfc().collect();
    let mut result = String::with_capacity(normalized_text.len());
    let mut chars = normalized_text.chars().peekable();
    let mut is_first_char = true;

    while let Some(ch) = chars.next() {
        let should_escape_list_prefix: bool = is_first_char
            && matches!(ch, '-' | '+')
            && chars.peek().is_some_and(|next| next.is_whitespace());

        match ch {
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' | '$' | '[' | ']' | '{'
            | '}'
                if !should_escape_list_prefix =>
            {
                result.push('\\');
                result.push(ch);
            }
            _ if should_escape_list_prefix => {
                result.push('\\');
                result.push(ch);
            }
            _ => result.push(ch),
        }

        is_first_char = false;
    }
    result
}
