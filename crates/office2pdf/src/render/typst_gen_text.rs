use std::fmt::Write;

use unicode_normalization::UnicodeNormalization;

use crate::render::font_subst;

use super::*;

/// Word's default tab stop interval (0.5 inch = 36pt).
const DEFAULT_TAB_WIDTH_PT: f64 = 36.0;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';

pub(super) fn generate_paragraph(
    out: &mut String,
    para: &Paragraph,
    line_grid_pitch: Option<f64>,
) -> Result<(), ConvertError> {
    let style = &para.style;

    if let Some(level) = style.heading_level {
        let _ = write!(out, "#heading(level: {level})[");
        generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());
        out.push_str("]\n");
        return Ok(());
    }

    let line_height_settings: Option<String> =
        word_line_height_settings(&para.runs, style, line_grid_pitch);
    let has_para_style = needs_block_wrapper(style) || line_height_settings.is_some();

    if has_para_style {
        // The wrapper must span the full line width: Typst blocks shrink to
        // their content by default, which would defeat the inner #align.
        out.push_str("#block(width: 100%");
        write_block_params_continuation(out, style);
        out.push_str(")[\n");
        write_paragraph_double_border_overlays(out, &style.border);
        write_line_box_settings(out, style.line_box);
        write_par_settings(out, style);
        if let Some(ref settings) = line_height_settings {
            out.push_str(settings);
        }
    }

    if para.runs.is_empty() {
        out.push_str("#v(12pt)");
        if has_para_style {
            out.push_str("\n]");
        }
        out.push('\n');
        return Ok(());
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

    generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

pub(super) fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.background.is_some()
        || style.border.is_some()
        || style.line_spacing.is_some()
        || style.line_box.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
}

/// Word snaps body lines to the section's document grid (`w:docGrid`
/// `w:linePitch`); Typst's glyph-tight default renders such documents
/// 20-30% shorter and shifts every page break. When the section carries a
/// grid pitch, the paragraph has no explicit line spacing, and the font's
/// metrics are known, emit metric line edges plus leading that tops the
/// line box up to the next grid multiple.
pub(super) fn word_line_height_settings(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<String> {
    let leading_pt: f64 = word_line_leading_pt(runs, style, line_grid_pitch)?;
    Some(format!(
        "#set text(top-edge: \"ascender\", bottom-edge: \"descender\")\n#set par(leading: {}pt)\n",
        format_f64(leading_pt)
    ))
}

/// The leading that accompanies Word metric text edges: the whitespace
/// Typst must insert between metric line boxes so consecutive lines land at
/// Word's single-space advance (or the document grid pitch). `None` when
/// the paragraph carries explicit line spacing or the font's metrics are
/// unknown — the metric-edge treatment does not apply then.
pub(super) fn word_line_leading_pt(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<f64> {
    if style.line_spacing.is_some() || style.line_box.is_some() {
        return None;
    }
    let family: &str = runs
        .iter()
        .find_map(|run| run.style.font_family.as_deref())?;
    let (ascender_em, descender_em, word_pitch_em) =
        crate::render::pdf::font_line_metrics_em(family)?;
    let font_size: f64 = runs
        .iter()
        .filter_map(|run| run.style.font_size)
        .fold(f64::NAN, f64::max);
    let font_size: f64 = if font_size.is_nan() { 11.0 } else { font_size };
    let line_box_pt: f64 = (ascender_em + descender_em) * font_size;
    if line_box_pt <= 0.0 {
        return None;
    }

    // Word only snaps East Asian text to the document grid: with the
    // default grid type, Latin-only paragraphs keep their hhea line height
    // (native Word GT: Arial 10.5 lines are 12pt while Korean lines in the
    // same document snap to the 18pt grid). Snapping Latin paragraphs
    // inflated every Western document by 30-50% (issue #354).
    let has_east_asian_text: bool = runs.iter().any(|run| run.text.chars().any(is_cjk_like));

    let leading_pt: f64 = match line_grid_pitch {
        Some(pitch) if pitch > 0.0 && has_east_asian_text => {
            let grid_lines: f64 = (line_box_pt / pitch).ceil().max(1.0);
            grid_lines * pitch - line_box_pt
        }
        // Word's single spacing is the font's full hhea line (ascender +
        // descender + line gap); Typst's metric edges resolve the typo
        // values, so the leading bridges the difference.
        _ => (word_pitch_em * font_size - line_box_pt).max(0.0),
    };
    Some(leading_pt)
}

pub(super) fn write_block_params(out: &mut String, style: &ParagraphStyle) {
    let mut first = true;

    if let Some(above) = style.space_before {
        write_param(out, &mut first, &format!("above: {}pt", format_f64(above)));
    }
    if let Some(below) = style.space_after {
        write_param(out, &mut first, &format!("below: {}pt", format_f64(below)));
    }
}

/// Like `write_block_params`, but for a parameter list that already has a
/// first entry (every parameter is prefixed with a comma).
fn write_block_params_continuation(out: &mut String, style: &ParagraphStyle) {
    if let Some(above) = style.space_before {
        let _ = write!(out, ", above: {}pt", format_f64(above));
    }
    if let Some(below) = style.space_after {
        let _ = write!(out, ", below: {}pt", format_f64(below));
    }
    if let Some(background) = style.background {
        // Word paints w:pPr/w:shd across the full paragraph width.
        let _ = write!(
            out,
            ", fill: rgb({}, {}, {})",
            background.r, background.g, background.b
        );
    }
    if let Some(border) = &style.border {
        write_paragraph_border_params(out, border);
    }
}

/// Word offsets paragraph border rules slightly from the text; `w:space` is
/// not carried per side, so a fixed 4pt gap approximates typical documents.
const PARAGRAPH_BORDER_GAP_PT: f64 = 4.0;

fn stroke_literal(side: &BorderSide) -> String {
    let paint = format!("rgb({}, {}, {})", side.color.r, side.color.g, side.color.b);
    let dash = match side.style {
        BorderLineStyle::Dotted => Some("dotted"),
        BorderLineStyle::Dashed => Some("dashed"),
        BorderLineStyle::DashDot | BorderLineStyle::DashDotDot => Some("dash-dotted"),
        _ => None,
    };
    match dash {
        Some(dash) => format!(
            "(paint: {paint}, thickness: {}pt, dash: \"{dash}\")",
            format_f64(side.width)
        ),
        None => format!("{}pt + {paint}", format_f64(side.width)),
    }
}

/// Emit `stroke:`/`inset:` block parameters for the paragraph's borders.
/// Double rules are drawn as overlays (Typst strokes have no double style),
/// so those sides only reserve inset space here.
fn write_paragraph_border_params(out: &mut String, border: &CellBorder) {
    let mut strokes: Vec<String> = Vec::new();
    let mut insets: Vec<String> = Vec::new();

    let mut push_side = |name: &str, side: &Option<BorderSide>| {
        let Some(side) = side else {
            return;
        };
        let reserved = if side.style == BorderLineStyle::Double {
            PARAGRAPH_BORDER_GAP_PT + side.width * 3.0
        } else {
            strokes.push(format!("{name}: {}", stroke_literal(side)));
            PARAGRAPH_BORDER_GAP_PT + side.width
        };
        insets.push(format!("{name}: {}pt", format_f64(reserved)));
    };
    push_side("top", &border.top);
    push_side("bottom", &border.bottom);
    push_side("left", &border.left);
    push_side("right", &border.right);

    if !strokes.is_empty() {
        let _ = write!(out, ", stroke: ({})", strokes.join(", "));
    }
    if !insets.is_empty() {
        let _ = write!(out, ", inset: ({})", insets.join(", "));
    }
}

/// Draw double-rule paragraph borders as two placed hairlines; Typst strokes
/// cannot render Word's double style. Only horizontal doubles occur in
/// practice (letterhead rules); vertical doubles fall back to a single
/// stroke drawn by `write_paragraph_border_params`.
fn write_paragraph_double_border_overlays(out: &mut String, border: &Option<Box<CellBorder>>) {
    let Some(border) = border else {
        return;
    };
    for (name, side) in [("top", &border.top), ("bottom", &border.bottom)] {
        let Some(side) = side else {
            continue;
        };
        if side.style != BorderLineStyle::Double {
            continue;
        }
        let w = side.width;
        let near_dy = PARAGRAPH_BORDER_GAP_PT + w;
        let far_dy = PARAGRAPH_BORDER_GAP_PT + w * 3.0;
        let (align, sign) = if name == "top" {
            ("top", -1.0)
        } else {
            ("bottom", 1.0)
        };
        for dy in [near_dy, far_dy] {
            let _ = write!(
                out,
                "#place({align}, dy: {}pt, line(length: 100%, stroke: {}pt + rgb({}, {}, {})))",
                format_f64(sign * dy),
                format_f64(w),
                side.color.r,
                side.color.g,
                side.color.b,
            );
        }
    }
}

pub(super) fn write_par_settings(out: &mut String, style: &ParagraphStyle) {
    if let Some(ref spacing) = style.line_spacing {
        match spacing {
            LineSpacing::Proportional(factor) => {
                let leading = factor * 0.65;
                let _ = writeln!(out, "  #set par(leading: {}em)", format_f64(leading));
            }
            LineSpacing::Exact(pts) => {
                let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(*pts));
            }
        }
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

pub(super) fn write_line_box_settings(out: &mut String, line_box: Option<LineBox>) {
    let Some(line_box) = line_box else {
        return;
    };
    let _ = writeln!(
        out,
        "#set text(top-edge: {}em, bottom-edge: -{}em)",
        format_f64(line_box.ascent_em),
        format_f64(line_box.descent_em),
    );
    out.push_str("#set par(leading: 0pt)\n");
}

pub(super) fn generate_runs_with_tabs(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
    if !paragraph_contains_tabs(runs) {
        generate_runs(out, runs);
        return;
    }

    let segments: Vec<Vec<Run>> = split_runs_on_tabs(runs);
    out.push_str("#context {\n");

    for (index, segment) in segments.iter().enumerate() {
        let _ = write!(out, "  let tab_segment_{index} = [");
        generate_runs(out, segment);
        out.push_str("]\n");

        if index == 0 {
            out.push_str("  let tab_prefix_0 = tab_segment_0\n");
            continue;
        }

        write_tab_segment_bindings(out, index, segment, tab_stops);
    }

    let _ = writeln!(out, "  tab_prefix_{}", segments.len() - 1);
    out.push('}');
}

pub(super) fn generate_runs_with_tabs_no_wrap(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
    let preserve_cjk_no_wrap: bool = runs
        .iter()
        .filter(|run| run.footnote.is_none())
        .any(|run| run.text.chars().any(is_cjk_like));
    let mut no_wrap_state: NoWrapState = NoWrapState::default();
    let transformed_runs: Vec<Run> = runs
        .iter()
        .map(|run| {
            let mut transformed_run: Run = run.clone();
            if transformed_run.footnote.is_none() {
                transformed_run.text = no_wrap_text(
                    &transformed_run.text,
                    preserve_cjk_no_wrap,
                    &mut no_wrap_state,
                );
            } else {
                no_wrap_state = NoWrapState::default();
            }
            transformed_run
        })
        .collect();

    generate_runs_with_tabs(out, &transformed_runs, tab_stops);
}

#[derive(Clone, Copy, Default)]
struct NoWrapState {
    previous_visible_char: Option<char>,
    previous_non_breaking_space: bool,
}

/// Emits Typst variable bindings for a non-first tab segment: measurement,
/// decimal anchor (if applicable), default remainder, advance, fill, and
/// the accumulated prefix content variable.
fn write_tab_segment_bindings(
    out: &mut String,
    index: usize,
    segment: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
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
        generate_runs(out, &anchor_runs);
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

fn paragraph_contains_tabs(runs: &[Run]) -> bool {
    runs.iter().any(|run| run.text.contains('\t'))
}

pub(super) fn generate_runs(out: &mut String, runs: &[Run]) {
    for run in runs {
        generate_run(out, run);
    }
}

fn no_wrap_text(text: &str, preserve_cjk_no_wrap: bool, state: &mut NoWrapState) -> String {
    if !preserve_cjk_no_wrap {
        return text.to_string();
    }

    let mut out: String = String::new();

    for ch in text.chars() {
        if matches!(ch, '\t' | PPTX_SOFT_LINE_BREAK_CHAR) {
            out.push(ch);
            *state = NoWrapState::default();
            continue;
        }

        if ch == ' ' {
            out.push('\u{00A0}');
            state.previous_visible_char = None;
            state.previous_non_breaking_space = true;
            continue;
        }

        if state.previous_non_breaking_space
            || state
                .previous_visible_char
                .is_some_and(|prev| needs_no_wrap_joiner(prev, ch))
        {
            out.push('\u{2060}');
        }
        out.push(ch);
        state.previous_visible_char = Some(ch);
        state.previous_non_breaking_space = false;
    }

    out
}

fn needs_no_wrap_joiner(previous: char, current: char) -> bool {
    !previous.is_whitespace() && !current.is_whitespace()
}

fn is_cjk_like(ch: char) -> bool {
    matches!(
        ch as u32,
        0x1100..=0x11FF
            | 0x2E80..=0x2FFF
            | 0x3000..=0x303F
            | 0x3040..=0x30FF
            | 0x3130..=0x318F
            | 0x31F0..=0x31FF
            | 0x3400..=0x4DBF
            | 0x4E00..=0x9FFF
            | 0xAC00..=0xD7AF
            | 0xF900..=0xFAFF
            | 0xFF00..=0xFFEF
    )
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
    let separator_offset: usize = find_decimal_separator_offset(&visible_text)?;

    let mut anchor_runs: Vec<Run> = Vec::new();
    let mut visible_offset: usize = 0;

    for run in runs {
        if run.footnote.is_some() {
            anchor_runs.push(run.clone());
            continue;
        }

        let run_end: usize = visible_offset + run.text.len();

        // Entire run falls before the separator — include it whole.
        if run_end <= separator_offset {
            if !run.text.is_empty() {
                anchor_runs.push(run.clone());
            }
            visible_offset = run_end;
            continue;
        }

        // This run spans the separator — include only the portion before it.
        let chars_before_separator: usize = separator_offset.saturating_sub(visible_offset);
        if chars_before_separator > 0 {
            anchor_runs.push(Run {
                text: run.text[..chars_before_separator].to_string(),
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

pub(super) fn generate_run(out: &mut String, run: &Run) {
    if let Some(ref content) = run.footnote {
        let escaped_content = escape_typst(content);
        let _ = write!(out, "#footnote[{escaped_content}]");
        return;
    }

    if run.text.contains(PPTX_SOFT_LINE_BREAK_CHAR) {
        write_run_with_soft_line_breaks(out, run);
        return;
    }

    write_run_segment(out, run, &run.text);
}

fn write_run_with_soft_line_breaks(out: &mut String, run: &Run) {
    let mut segment_start: usize = 0;

    for (offset, ch) in run.text.char_indices() {
        if ch != PPTX_SOFT_LINE_BREAK_CHAR {
            continue;
        }

        if segment_start < offset {
            write_run_segment(out, run, &run.text[segment_start..offset]);
        }
        out.push_str("#linebreak()");
        segment_start = offset + ch.len_utf8();
    }

    if segment_start < run.text.len() {
        write_run_segment(out, run, &run.text[segment_start..]);
    }
}

fn write_run_segment(out: &mut String, run: &Run, text: &str) {
    let style = &run.style;

    let needs_all_caps: bool = matches!(style.all_caps, Some(true));
    let escaped: String = if needs_all_caps {
        escape_typst(&text.to_uppercase())
    } else {
        escape_typst(text)
    };

    let wrappers: Vec<String> = collect_formatting_wrappers(run);

    for wrapper in &wrappers {
        out.push_str(wrapper);
    }

    write_run_content(out, &escaped, style);

    for _ in &wrappers {
        out.push(']');
    }
}

/// Builds the ordered list of `#command[` openers that wrap a run's content.
/// The order matches the original nesting: link > highlight > strike >
/// underline > super/sub > smallcaps.
fn collect_formatting_wrappers(run: &Run) -> Vec<String> {
    let style: &TextStyle = &run.style;
    let mut wrappers: Vec<String> = Vec::new();

    if let Some(ref href) = run.href {
        wrappers.push(format!("#link(\"{href}\")["));
    }
    if let Some(ref highlight) = style.highlight {
        wrappers.push(format!(
            "#highlight(fill: rgb({}, {}, {}))[",
            highlight.r, highlight.g, highlight.b
        ));
    }
    if matches!(style.strikethrough, Some(true)) {
        wrappers.push("#strike[".to_string());
    }
    if matches!(style.underline, Some(true)) {
        wrappers.push("#underline[".to_string());
    }
    if matches!(style.vertical_align, Some(VerticalTextAlign::Superscript)) {
        wrappers.push("#super[".to_string());
    }
    if matches!(style.vertical_align, Some(VerticalTextAlign::Subscript)) {
        wrappers.push("#sub[".to_string());
    }
    if matches!(style.small_caps, Some(true)) {
        wrappers.push("#smallcaps[".to_string());
    }

    wrappers
}

/// Writes the innermost content of a run: either `#text(params)[escaped]`
/// when text properties are present, or the escaped text directly (with a
/// `#[...]` safety wrapper when needed to prevent Typst syntax ambiguity).
fn write_run_content(out: &mut String, escaped: &str, style: &TextStyle) {
    if has_text_properties(style) {
        out.push_str("#text(");
        write_text_params(out, style);
        out.push_str(")[");
        out.push_str(escaped);
        out.push(']');
        return;
    }

    let needs_safety_wrap: bool = !escaped.is_empty()
        && out.ends_with(']')
        && !out.ends_with("\\]")
        && matches!(escaped.as_bytes()[0], b'(' | b'.' | b'[');

    if needs_safety_wrap {
        out.push_str("#[");
        out.push_str(escaped);
        out.push(']');
    } else {
        out.push_str(escaped);
    }
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
    // Only infer weight from font family name when the font (or its alias)
    // is actually available.  When using fallback fonts, uncommonly heavy
    // weights (e.g. "extrabold" = 800) may not exist in the substitute,
    // causing Typst to fall back to its built-in serif font instead.
    let inferred = style.font_family.as_deref().and_then(|family| {
        if font_subst::is_primary_font_available(family) {
            inferred_font_weight(family)
        } else {
            None
        }
    });
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

    // A leading "<digits>. " would be re-typeset as a Typst numbered-list
    // marker (e.g. "2026. 07. 17." became "2026. 7. 17."); escape its dot.
    let enum_marker_dot: Option<usize> = {
        let digit_count = normalized_text
            .chars()
            .take_while(|c| c.is_ascii_digit())
            .count();
        let rest = &normalized_text[digit_count..];
        if digit_count > 0 && (rest.starts_with(". ") || rest == ".") {
            Some(digit_count)
        } else {
            None
        }
    };

    let mut result = String::with_capacity(normalized_text.len());
    let mut chars = normalized_text.chars().peekable();
    let mut is_first_char = true;
    let mut char_index: usize = 0;

    let mut after_linebreak = false;
    while let Some(ch) = chars.next() {
        let should_escape_list_prefix: bool = is_first_char
            && matches!(ch, '-' | '+')
            && chars.peek().is_some_and(|next| next.is_whitespace());

        match ch {
            // A hard line break (`<w:br/>`, carried through the IR as '\n') must
            // force a new line. A bare newline in Typst markup collapses to a
            // space, which silently merged code lines like `echo` / `printf`
            // (issue #176).
            '\n' => result.push_str("#linebreak()"),
            '\r' => {}
            // Word preserves literal space runs (xml:space="preserve") that
            // documents use for manual alignment and code indentation; Typst
            // markup collapses consecutive and line-leading spaces to one.
            // Emit runs of two or more — and post-break indentation — as a
            // code-mode string, which markup cannot collapse (issue #352).
            // Single run-leading spaces stay literal: they sit between
            // sibling runs in the same markup line and survive as-is.
            ' ' if after_linebreak || chars.peek().is_some_and(|next| *next == ' ') => {
                let mut run_len: usize = 1;
                while chars.peek().is_some_and(|next| *next == ' ') {
                    chars.next();
                    run_len += 1;
                    char_index += 1;
                }
                result.push_str("#\"");
                result.push_str(&" ".repeat(run_len));
                // The semicolon ends the code expression: without it, a
                // following `(` or `[` in the text would chain onto the
                // string as a function call (`#"  "(SIB)`).
                result.push_str("\";");
            }
            // Quotes and hyphens are Typst markup shorthands: smartquote
            // curls straight quotes, `--` ligates to an en dash, and a
            // hyphen before digits becomes a Unicode minus. Word stores the
            // literal characters the author typed, so all of them must
            // render verbatim (issue #353).
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' | '$' | '[' | ']' | '{'
            | '}' | '"' | '\'' | '-'
                if !should_escape_list_prefix =>
            {
                result.push('\\');
                result.push(ch);
            }
            _ if should_escape_list_prefix => {
                result.push('\\');
                result.push(ch);
            }
            '.' if enum_marker_dot == Some(char_index) => {
                result.push('\\');
                result.push('.');
            }
            _ => result.push(ch),
        }

        after_linebreak = ch == '\n';
        is_first_char = false;
        char_index += 1;
    }
    result
}
