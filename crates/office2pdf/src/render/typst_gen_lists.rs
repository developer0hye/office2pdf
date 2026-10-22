use super::*;
use crate::ir::{ListMarkerSuffix, NumberingPattern};
use crate::numbering::format_roman_counter as roman_marker;

const PPTX_BLANK_LIST_ITEM_SENTINEL: &str = "\u{E000}\u{E001}";

fn is_pptx_blank_list_item(paragraph: &Paragraph) -> bool {
    paragraph.runs.len() == 1
        && paragraph.runs[0].footnote.is_none()
        && paragraph.runs[0].text == PPTX_BLANK_LIST_ITEM_SENTINEL
}

/// Generate Typst markup for a list (ordered or unordered).
///
/// Uses Typst's `#enum()` for ordered lists and `#list()` for unordered lists.
/// Nested items are wrapped in `list.item()` / `enum.item()` with a sub-list.
struct EffectiveListStyle<'a> {
    kind: ListKind,
    numbering_pattern: Option<&'a NumberingPattern>,
    full_numbering: bool,
    marker_text: Option<&'a str>,
    marker_style: Option<&'a TextStyle>,
}

#[derive(Clone, Copy)]
struct ListIndentGeometry {
    marker_origin_pt: f64,
    marker_width_pt: f64,
}

#[derive(Clone, Copy, Default)]
struct ListOpenOptions {
    indent: Option<ListIndentGeometry>,
    spacing_pt: Option<f64>,
    start_at: Option<u32>,
    baseline_snap: Option<PowerPointBaselineSnap>,
}

fn list_style_for_level<'a>(list: &'a List, level: u32) -> EffectiveListStyle<'a> {
    if let Some(style) = list.level_styles.get(&level) {
        EffectiveListStyle {
            kind: style.kind,
            numbering_pattern: style.numbering_pattern.as_ref(),
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
    labels: &LevelLabelStyles,
    options: ListOpenOptions,
) {
    let fallback_marker_style: Option<&TextStyle> = labels.common.as_ref();
    let ListOpenOptions {
        indent,
        spacing_pt,
        start_at,
        baseline_snap,
    } = options;
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
        let numbered_marker_styles: Vec<(u32, Option<TextStyle>)> = labels
            .by_number
            .iter()
            .map(|(number, label)| (*number, merge_marker_style(Some(label), style.marker_style)))
            .collect();
        if marker_style.as_ref().is_some_and(has_text_properties)
            || !numbered_marker_styles.is_empty()
            || indent.is_some()
            || baseline_snap.is_some()
        {
            write_ordered_list_numbering_function(
                out,
                style,
                marker_style.as_ref(),
                &numbered_marker_styles,
                indent,
                baseline_snap,
            );
            out.push_str(", ");
        } else if let Some(numbering_pattern) = style.numbering_pattern {
            out.push_str("numbering: ");
            if let Some(pattern) = typst_numbering_pattern(numbering_pattern) {
                let _ = write!(out, "\"{}\"", escape_typst_string(&pattern));
            } else {
                out.push_str("(..nums) => ");
                write_numbering_label(out, Some(numbering_pattern), MarkerNumbers::Arguments);
            }
            out.push_str(", ");
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
        if let Some(snap) = baseline_snap {
            snap.write_marker_open(out);
        }
        if let Some(indent) = indent {
            write_marker_box_open(out, indent.marker_width_pt);
        }
        write_unordered_list_marker_content(out, &marker_text, marker_style.as_ref());
        if indent.is_some() {
            out.push_str("])");
        }
        if baseline_snap.is_some() {
            out.push(']');
        }
        out.push_str("], ");
    }

    out.push('\n');
}

/// `numbered_marker_styles` gives single items their own label formatting by
/// the number Typst passes the function, the only thing it learns about the
/// item; every other label takes `marker_style`.
fn write_ordered_list_numbering_function(
    out: &mut String,
    style: &EffectiveListStyle<'_>,
    marker_style: Option<&TextStyle>,
    numbered_marker_styles: &[(u32, Option<TextStyle>)],
    indent: Option<ListIndentGeometry>,
    baseline_snap: Option<PowerPointBaselineSnap>,
) {
    out.push_str("numbering: (..nums) => [");
    if let Some(snap) = baseline_snap {
        snap.write_marker_open(out);
    }
    if let Some(indent) = indent {
        write_marker_box_open(out, indent.marker_width_pt);
    }
    if numbered_marker_styles.is_empty() {
        write_ordered_list_label(out, style, marker_style);
    } else {
        out.push_str("#{ let number = nums.pos().last(); ");
        for (number, numbered_style) in numbered_marker_styles {
            let _ = write!(out, "if number == {number} [");
            write_ordered_list_label(out, style, numbered_style.as_ref());
            out.push_str("] else ");
        }
        out.push('[');
        write_ordered_list_label(out, style, marker_style);
        out.push_str("] }");
    }
    if indent.is_some() {
        out.push_str("])");
    }
    if baseline_snap.is_some() {
        out.push(']');
    }
    out.push(']');
}

/// Markup for one label of a numbering function, which binds `nums`.
fn write_ordered_list_label(
    out: &mut String,
    style: &EffectiveListStyle<'_>,
    marker_style: Option<&TextStyle>,
) {
    if let Some(marker_style) = marker_style.filter(|style| has_text_properties(style)) {
        out.push_str("#text(");
        // The marker's text is `#numbering`'s result, which the engine
        // computes from the pattern and the item's position — unnameable
        // here, so the kerning answer stays on the safe side (issue #628).
        write_text_params(out, marker_style);
        out.push_str(")[");
    }
    out.push('#');
    write_numbering_label(out, style.numbering_pattern, MarkerNumbers::Arguments);
    if marker_style.is_some_and(has_text_properties) {
        out.push(']');
    }
}

/// The numbers a label formats: a numbering function's `..nums`, or values
/// known while the item is written.
#[derive(Clone, Copy)]
enum MarkerNumbers<'a> {
    Arguments,
    Values(&'a [u32]),
}

/// One Typst pattern string for `pattern`, if Typst reads exactly its
/// counters back: its literal text may hold a counting symbol (issue #2221).
fn typst_numbering_pattern(pattern: &NumberingPattern) -> Option<String> {
    let mut joined: String = String::new();
    for (prefix, symbol) in &pattern.pieces {
        joined.push_str(prefix);
        joined.push(*symbol);
    }
    joined.push_str(&pattern.suffix);
    let counters: usize = joined
        .parse::<typst::model::NumberingPattern>()
        .ok()?
        .pieces();
    (counters == pattern.pieces.len()).then_some(joined)
}

/// Writes a code expression for the label. When one pattern string would turn
/// literal text into counters, each counter is formatted alone between literal
/// strings, taking the number at its position; a counter without one is
/// dropped with its literal text, as Typst's patterns do.
fn write_numbering_label(
    out: &mut String,
    pattern: Option<&NumberingPattern>,
    numbers: MarkerNumbers<'_>,
) {
    let typst_pattern: Option<String> = match pattern {
        Some(pattern) => typst_numbering_pattern(pattern),
        None => Some("1.".to_string()),
    };
    if let Some(typst_pattern) = typst_pattern {
        let _ = write!(out, "numbering(\"{}\"", escape_typst_string(&typst_pattern));
        match numbers {
            MarkerNumbers::Arguments => out.push_str(", ..nums"),
            MarkerNumbers::Values(values) => {
                for value in values {
                    let _ = write!(out, ", {value}");
                }
            }
        }
        out.push(')');
        return;
    }
    let Some(pattern) = pattern else {
        return;
    };
    out.push_str("{ let n = ");
    match numbers {
        MarkerNumbers::Arguments => out.push_str("nums.pos()"),
        MarkerNumbers::Values(values) => {
            out.push('(');
            for value in values {
                let _ = write!(out, "{value},");
            }
            out.push(')');
        }
    }
    out.push_str("; \"\"");
    for (index, (prefix, symbol)) in pattern.pieces.iter().enumerate() {
        let _ = write!(
            out,
            " + if n.len() > {index} {{ \"{}\" + numbering(\"{symbol}\", n.at({index})) }} else {{ \"\" }}",
            escape_typst_string(prefix)
        );
    }
    let _ = write!(out, " + \"{}\" }}", escape_typst_string(&pattern.suffix));
}

fn write_marker_box_open(out: &mut String, width_pt: f64) {
    // Typst normally sizes the marker column from the glyph. Word instead
    // reserves the nominal hanging-indent span for narrow bullets and digits.
    // Wide Word numbers advance only the first body line to its next tab.
    let _ = write!(out, "#box(width: {}pt, align(left)[", format_f64(width_pt));
}

fn write_unordered_list_marker_content(
    out: &mut String,
    marker_text: &str,
    marker_style: Option<&TextStyle>,
) {
    if let Some(marker_style) = marker_style.filter(|style| has_text_properties(style)) {
        out.push_str("#text(");
        write_text_params_for_text(out, marker_style, marker_text);
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

/// A list paragraph's marker origin and hanging span from the list's origin,
/// or `None` without a hanging indent.
///
/// With `hangs_into_margin` both keep their sign, as in a Word body list: a
/// negative `w:ind w:left`, or a hanging span wider than it, starts the marker
/// left of the margin (issue #2233). Otherwise the left indent and the marker
/// clamp at the origin.
fn paragraph_list_indent(
    style: &ParagraphStyle,
    hangs_into_margin: bool,
) -> Option<ListIndentGeometry> {
    let clamp = |pt: f64| if hangs_into_margin { pt } else { pt.max(0.0) };
    let indent_left = clamp(style.indent_left?);
    let first_line_indent = style.indent_first_line?;
    if first_line_indent >= -0.0001 {
        return None;
    }

    let marker_origin_pt = clamp(indent_left + first_line_indent);
    let marker_width_pt = indent_left - marker_origin_pt;
    (marker_width_pt > 0.0001).then_some(ListIndentGeometry {
        marker_origin_pt,
        marker_width_pt,
    })
}

/// How far a level's items may stray from its first item's marker origin and
/// hanging span and still share that geometry.
///
/// Word stores indents in twips, and one level's items often disagree by a
/// rounding step: 992 vs 993 twips, or 357 (0.63cm) vs 360 (0.25in). Requiring
/// an exact match dropped the whole level to Typst's default indents, putting
/// the markers on the margin (issue #2332). In the docxide and bulk corpora
/// such spreads stay within 15 twips (0.75pt), well under what a reader sees,
/// while the next smallest is 3.8pt.
const LIST_LEVEL_INDENT_TOLERANCE_PT: f64 = 1.0;

/// The marker origin and hanging span one level's items share, taken from its
/// first item, or `None` if any item lacks one or strays further than
/// [`LIST_LEVEL_INDENT_TOLERANCE_PT`] from it.
fn common_list_level_indent(
    items: &[crate::ir::ListItem],
    level: u32,
    hangs_into_margin: bool,
) -> Option<ListIndentGeometry> {
    let mut geometries = items
        .iter()
        .filter(|item| item.level == level)
        .filter_map(|item| item.content.first())
        .map(|paragraph| paragraph_list_indent(&paragraph.style, hangs_into_margin));
    let first = geometries.next()??;
    let within_tolerance =
        |left: f64, right: f64| (left - right).abs() <= LIST_LEVEL_INDENT_TOLERANCE_PT;

    geometries
        .all(|geometry| {
            geometry.is_some_and(|geometry| {
                within_tolerance(geometry.marker_origin_pt, first.marker_origin_pt)
                    && within_tolerance(geometry.marker_width_pt, first.marker_width_pt)
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
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) => points.max(0.0),
        None => font_size,
    }
}

fn list_boundary_spacing(
    previous: &crate::ir::ListItem,
    next: &crate::ir::ListItem,
    wrapper_spans_full_line: bool,
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
    // The rendered line already carries its advance, through a fixed wrapper
    // or a natural fallback box. Word adds w:spacing before/after on top of
    // that advance, so the item gap is exactly the paragraph gap.
    // Adding a whole line height instead
    // stretched every spaced list by roughly a line per item (issues #384,
    // #452, #1685).
    if wrapper_spans_full_line
        && paragraph_uses_full_line_box(previous_paragraph)
        && paragraph_uses_full_line_box(next_paragraph)
    {
        return Some(paragraph_gap);
    }

    // An explicit Typst list spacing replaces its automatic paragraph
    // leading. Carry the line box as well as the before/after gap so adding
    // paragraph spacing cannot accidentally make a tight list tighter.
    let line_height =
        paragraph_line_height(previous_paragraph).max(paragraph_line_height(next_paragraph));
    Some(line_height + paragraph_gap)
}

/// Whether the paragraph renders under the wrapper's full-advance line box.
///
/// A stated line multiple is not an exception to that box; it is part of it.
/// `word_line_leading_pt` scales Word's advance by the multiple, and
/// `powerpoint_paragraph_line_box_em` does the same with `a:lnSpc`, before
/// either hands the product to the wrapper's fixed text edges. A paragraph
/// stating `w:spacing w:line="259" w:lineRule="auto"` therefore already renders
/// under a box spanning its own scaled line, and treating it as boxless counted
/// that line a second time — on every gap between items and on the gap below
/// the list. Word's default template writes `w:line="259"` into `w:pPrDefault`,
/// so the double count reached nearly every authored list (issue #1685).
///
/// A fixed `LineBox` and an exact line rule stay out. An exact rule leaves the
/// Word path with no wrapper settings at all, so the caller's
/// `wrapper_spans_full_line` gate already denies it this branch, and a declared
/// `LineBox` is answered by the earlier branch of each caller.
fn paragraph_uses_full_line_box(paragraph: &Paragraph) -> bool {
    let line_spacing_is_in_the_box: bool = match paragraph.style.line_spacing {
        None => true,
        Some(LineSpacing::Proportional(factor)) => factor > 0.0,
        Some(LineSpacing::Exact(_) | LineSpacing::AtLeast(_)) => false,
    };
    line_spacing_is_in_the_box && paragraph.style.line_box.is_none()
}

fn common_list_level_spacing(
    items: &[crate::ir::ListItem],
    level: u32,
    wrapper_spans_full_line: bool,
) -> Option<f64> {
    // A nested group is a real document boundary. Filtering the items first
    // paired root items on opposite sides of that group, so an accidental
    // match could hoist a synthetic gap and suppress every real per-item gap
    // in the outline (issue #659).
    let mut boundaries = items
        .windows(2)
        .filter(|pair| pair[0].level == level && pair[1].level == level)
        .map(|pair| list_boundary_spacing(&pair[0], &pair[1], wrapper_spans_full_line));
    let first = boundaries.next()??;

    (first > 0.0001
        && boundaries.all(|spacing| spacing.is_some_and(|spacing| f64_approx_eq(spacing, first))))
    .then_some(first)
}

fn list_edge_spacing(
    list: &List,
    level: u32,
    wrapper_spans_full_line: bool,
) -> (Option<f64>, Option<f64>) {
    let first = list.items.iter().find(|item| item.level == level);
    let last = list.items.iter().rev().find(|item| item.level == level);
    let edge = |paragraph: &Paragraph, spacing: Option<f64>| {
        if paragraph.style.line_box.is_some() {
            return Some(spacing.unwrap_or(0.0).max(0.0));
        }
        // Same Word semantics as the item boundaries: before/after extends
        // the line advance, which the rendered line already carries, so
        // the whitespace is the gap alone (issues #384, #452, #1685).
        if wrapper_spans_full_line && paragraph_uses_full_line_box(paragraph) {
            return spacing
                .map(|spacing| spacing.max(0.0))
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

/// Emit a list, wrapped in a single block that carries both the Word line
/// box (`line_height_settings`, when the paragraphs use one) and the list's
/// own `w:spacing` edge gaps.
///
/// The two used to be separate blocks, with the gaps on the inner one. Typst
/// does not apply block spacing at a container's edges, so the gaps never
/// reached the boundary and the outer wrapper fell back to Typst's own 1.2em
/// `block.spacing` — about 10pt too much after a one-item numbered list
/// (issue #463).
pub(super) fn generate_list(
    out: &mut String,
    list: &List,
    line_height_settings: Option<&str>,
    eojeol_wrap: ListEojeolWrap,
    ctx: &mut super::GenCtx,
) -> Result<(), ConvertError> {
    generate_list_with_spacing_model(out, list, line_height_settings, false, eojeol_wrap, ctx)
}

/// `per_item_gaps` selects PowerPoint's paragraph spacing model over Word's:
/// a slide item's `a:spcAft` belongs to that item and does not collapse
/// against its neighbour's `a:spcBef`, so items declaring different gaps each
/// keep their own (issue #524).
pub(super) fn generate_list_with_spacing_model(
    out: &mut String,
    list: &List,
    line_height_settings: Option<&str>,
    per_item_gaps: bool,
    mut eojeol_wrap: ListEojeolWrap,
    ctx: &mut super::GenCtx,
) -> Result<(), ConvertError> {
    if let Some((normalized, blank_height_pt)) = normalize_pptx_blank_list_items(list) {
        if normalized.items.is_empty() {
            let _ = writeln!(out, "#block(height: {}pt)[]", format_f64(blank_height_pt));
            return Ok(());
        }
        let normalized_line_settings: Option<String> = normalized
            .items
            .first()
            .and_then(|item| item.content.first())
            .and_then(|paragraph| {
                powerpoint_line_height_settings(&paragraph.runs, &paragraph.style)
            });
        return generate_list_with_spacing_model(
            out,
            &normalized,
            normalized_line_settings.as_deref(),
            per_item_gaps,
            eojeol_wrap,
            ctx,
        );
    }

    let wrapper_spans_full_line: bool = line_height_settings.is_some();
    // A Word paragraph with unavailable font metrics keeps a natural line
    // box. Its height carries the advance; paragraph gaps must not add it
    // again when a continuation becomes a separate block. Unspecified font
    // families retain the generic list spacing model.
    let spacing_spans_full_line: bool = wrapper_spans_full_line
        || (!per_item_gaps
            && list
                .items
                .iter()
                .flat_map(|item| &item.content)
                .all(|paragraph| {
                    paragraph_uses_full_line_box(paragraph)
                        && super::text::word_paragraph_metric_family(
                            &paragraph.runs,
                            &paragraph.style,
                        )
                        .is_some()
                }));
    let root_level: u32 = list_root_level(list);
    let style = list_style_for_level(list, root_level);
    let labels = level_label_styles(&list.items, root_level, style.kind);
    let indent = common_list_level_indent(&list.items, root_level, eojeol_wrap.hangs_into_margin);
    // Typst applies one root `spacing:` after every complete root item. In a
    // PowerPoint outline that item can end in a nested child, whose own
    // a:spcAft is the real boundary gap. Even one adjacent root pair is not
    // enough evidence to hoist a value across those different boundaries.
    let has_nested_items = list.items.iter().any(|item| item.level > root_level);
    // A Word level transition is a paragraph boundary, even when Typst wraps
    // the child level in a separate list. Carry each collapsed gap across
    // those containers instead of inferring it from same-level neighbours.
    let word_boundary_gaps: bool = !per_item_gaps
        && has_nested_items
        && list
            .items
            .iter()
            .flat_map(|item| &item.content)
            .any(|paragraph| {
                paragraph.style.space_before.is_some() || paragraph.style.space_after.is_some()
            })
        && list
            .items
            .iter()
            .flat_map(|item| &item.content)
            .all(|paragraph| {
                paragraph_uses_full_line_box(paragraph) || paragraph.style.line_box.is_some()
            });
    let spacing_pt = if word_boundary_gaps {
        Some(0.0)
    } else if per_item_gaps && has_nested_items {
        None
    } else {
        common_list_level_spacing(&list.items, root_level, spacing_spans_full_line)
    };
    let (space_before, mut space_after) =
        list_edge_spacing(list, root_level, spacing_spans_full_line);
    if word_boundary_gaps {
        // The final physical paragraph owns the gap, even when it is a child.
        space_after = list.items.last().map(paragraph_space_after);
    }
    let line_box = common_list_line_box(list);
    let start_at = list.items.first().and_then(|item| item.start_at);
    let needs_wrapper: bool =
        space_before.is_some() || space_after.is_some() || line_height_settings.is_some();
    if needs_wrapper {
        // The wrapper must span the full line width: Typst blocks shrink to
        // their content by default, which would strand a wide list item.
        out.push_str("#block(width: 100%");
        if let Some(above) = space_before {
            let _ = write!(out, ", above: {}pt", format_f64(above));
        }
        if let Some(below) = space_after {
            let _ = write!(out, ", below: {}pt", format_f64(below));
        }
        out.push_str(")[\n");
        if let Some(settings) = line_height_settings {
            out.push_str(settings);
        }
        write_line_box_settings(out, line_box);
    }
    // A framed eojeol restores exactly the fixed text edges the wrapper just
    // emitted, and nothing else: `write_line_box_settings` comes last, so a
    // declared `LineBox` wins over the caller's computed Word line, and
    // without the wrapper no fixed edges were emitted at all (issue #626).
    eojeol_wrap.line_box_em = needs_wrapper
        .then(|| {
            line_box
                .map(|line_box| (line_box.ascent_em, line_box.descent_em))
                .or_else(|| eojeol_wrap.line_box_em.filter(|_| wrapper_spans_full_line))
        })
        .flatten();
    // Typst lays out markers separately from their item bodies. Keep one
    // scope for this flat slide list so each marker follows its own body's
    // first physical line after wrapping has been decided.
    let adjusts_lines = eojeol_wrap
        .baseline_snap
        .is_some_and(PowerPointBaselineSnap::adjusts_physical_lines);
    if adjusts_lines {
        out.push_str("#context { metadata(\"office2pdf-pptx-list-start\") + [\n");
    }
    write_list_open(
        out,
        "#",
        &style,
        &labels,
        ListOpenOptions {
            indent,
            spacing_pt,
            start_at,
            baseline_snap: eojeol_wrap.baseline_snap,
        },
    );
    generate_list_items(
        out,
        list,
        &list.items,
        root_level,
        wrapper_spans_full_line,
        spacing_pt.is_some() || !per_item_gaps,
        per_item_gaps,
        word_boundary_gaps,
        None,
        &eojeol_wrap,
        &[],
        ctx,
    )?;
    out.push_str(")\n");
    if adjusts_lines {
        out.push_str("] + metadata(\"office2pdf-pptx-list-end\") }\n");
    }
    if needs_wrapper {
        out.push_str("]\n");
    }
    Ok(())
}

/// Remove marker-suppressed DrawingML paragraphs before the generic Typst
/// list path sees them. A Typst `list.item` always paints its marker, so each
/// blank line is carried as paragraph spacing on its preceding visible item
/// (or as leading spacing on the first visible item). Removing the item also
/// means an ordered list does not consume a number for PowerPoint's blank.
fn normalize_pptx_blank_list_items(list: &List) -> Option<(List, f64)> {
    if !list
        .items
        .iter()
        .any(|item| item.content.len() == 1 && is_pptx_blank_list_item(&item.content[0]))
    {
        return None;
    }

    let mut normalized: List = list.clone();
    normalized.items.clear();
    let mut leading_blank_height_pt: f64 = 0.0;
    let mut leading_start_at: Option<u32> = None;

    for item in &list.items {
        if item.content.len() == 1 && is_pptx_blank_list_item(&item.content[0]) {
            let paragraph: &Paragraph = &item.content[0];
            let blank_height_pt: f64 = pptx_blank_list_item_advance_pt(paragraph)
                + paragraph.style.space_before.unwrap_or(0.0)
                + paragraph.style.space_after.unwrap_or(0.0);
            if let Some(previous) = normalized.items.last_mut()
                && let Some(paragraph) = previous.content.last_mut()
            {
                paragraph.style.space_after =
                    Some(paragraph.style.space_after.unwrap_or(0.0) + blank_height_pt);
            } else {
                leading_blank_height_pt += blank_height_pt;
                leading_start_at = leading_start_at.or(item.start_at);
            }
            continue;
        }

        let mut visible_item = item.clone();
        if normalized.items.is_empty() {
            if let Some(paragraph) = visible_item.content.first_mut() {
                paragraph.style.space_before =
                    Some(paragraph.style.space_before.unwrap_or(0.0) + leading_blank_height_pt);
            }
            if visible_item.start_at.is_none() {
                visible_item.start_at = leading_start_at;
            }
        }
        normalized.items.push(visible_item);
    }

    Some((normalized, leading_blank_height_pt))
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
        (
            LineSpacing::Exact(l) | LineSpacing::AtLeast(l),
            LineSpacing::Exact(r) | LineSpacing::AtLeast(r),
        ) => f64_approx_eq(l, r),
        _ => false,
    }
}

pub(super) fn generate_fixed_text_list(
    out: &mut String,
    list: &List,
    include_item_spacing: bool,
    available_width_pt: Option<f64>,
    uses_powerpoint_line: bool,
    baseline_mode: PowerPointBaselineMode,
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
        write_fixed_text_list_par_settings(out, style, &paragraph.runs, line_gap_pt);
    }
    // Item advances and paragraph gaps are emitted explicitly below. Typst's
    // default paragraph spacing otherwise sits between the metadata-wrapped
    // item blocks as an unrequested extra line (issue #1913).
    out.push_str("#context { set par(spacing: 0pt); [\n");

    let align_str: Option<&str> = fixed_text_list_alignment(style.alignment);
    let mut current_number: u32 = list
        .items
        .first()
        .and_then(|item| item.start_at)
        .unwrap_or(1);
    // The line advance and the paragraph gap are separate quantities and both
    // land between items: the first is what `#set par(leading:)` puts between
    // the wrapped lines of one item, the second is PowerPoint's spcAft+spcBef
    // (issue #928). The advance has to be emitted rather than left to the item
    // blocks' own spacing, which is Typst's 1.2em of the *ambient* size and has
    // nothing to do with the list's leading (issue #934).
    // A slide's list paces on PowerPoint's line box, the same model its text
    // boxes and table cells already use, rather than on Typst's cap-height box
    // plus 0.65em (issue #934).
    let powerpoint_line: Option<(String, f64)> = uses_powerpoint_line
        .then(|| fixed_text_list_powerpoint_line(list))
        .flatten();
    // On the PowerPoint path the item's own block already stands one full line
    // box tall, so the boundary carries only what the paragraphs declare on top
    // of it. Adding the advance there as well counted the line twice — 67.6pt
    // against a reference's 38.8pt on the `AGENDA` list of #841.
    let leading_pt: f64 = match powerpoint_line {
        Some(_) => 0.0,
        None => fixed_text_list_leading_pt(style, list),
    };
    let paragraph_gap_pt: f64 = fixed_text_list_paragraph_gap_pt(list).unwrap_or(0.0);
    let item_gap_pt: f64 = leading_pt + paragraph_gap_pt;
    let active_gap: Option<f64> =
        Some(item_gap_pt).filter(|gap| *gap > 0.0 && include_item_spacing);
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
                if include_item_spacing {
                    // Emitted apart so each stays legible as the quantity it
                    // is: the line advance, then whatever spacing this
                    // boundary's own paragraphs declare on top of it.
                    if leading_pt > 0.0 {
                        let _ = writeln!(out, "#v({}pt)", format_f64(leading_pt));
                    }
                    let boundary_gap_pt: f64 =
                        fixed_text_list_boundary_gap_pt(&list.items[index - 1], item);
                    if boundary_gap_pt > 0.0 {
                        let _ = writeln!(out, "#v({}pt)", format_f64(boundary_gap_pt));
                    }
                }
            }
            if let Some(start_at) = item.start_at {
                current_number = start_at;
            }
        }

        let item_paragraph: &Paragraph = &item.content[0];
        let suppress_marker: bool = is_pptx_blank_list_item(item_paragraph);
        let marker_text: String = if suppress_marker {
            String::new()
        } else {
            fixed_text_list_marker(
                list.kind,
                &effective_style,
                current_number,
                &item_paragraph.runs,
            )
        };

        if use_stack {
            out.push('[');
        }
        let baseline_snap: Option<PowerPointBaselineSnap> = baseline_mode
            .for_paragraph(&item_paragraph.runs, &item_paragraph.style)
            .filter(|_| !suppress_marker);
        if let Some(snap) = baseline_snap {
            snap.write_open(out);
        }
        write_fixed_text_list_item(
            out,
            item_paragraph,
            &effective_style,
            &marker_text,
            align_str,
            available_width_pt,
            powerpoint_line
                .as_ref()
                .map(|(settings, _)| settings.as_str()),
        );
        if baseline_snap.is_some() {
            out.push(']');
        }
        if use_stack {
            out.push(']');
        } else {
            out.push('\n');
        }

        if list.kind == ListKind::Ordered && !suppress_marker {
            current_number += 1;
        }
    }

    if use_stack {
        out.push_str("\n)");
    }
    out.push_str("] }\n");
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
    powerpoint_line_settings: Option<&str>,
) {
    let inset: Insets = fixed_text_paragraph_inset(&paragraph.style);
    let has_inset: bool = inset.left > 0.0 || inset.right > 0.0;
    let hanging_indent_pt: Option<f64> = fixed_text_paragraph_hanging_indent_pt(&paragraph.style);
    let use_marker_grid: bool = list_style.kind == ListKind::Ordered && hanging_indent_pt.is_some();

    out.push_str("#block(width: ");
    if let Some(width_pt) = available_width_pt {
        let _ = write!(out, "{}pt", format_f64(width_pt));
    } else {
        out.push_str("100%");
    }
    // The distance between two items is emitted between them, so the item's own
    // block must contribute nothing: Typst's default is 1.2em of the ambient
    // size and would be added on top (issue #934).
    if powerpoint_line_settings.is_some() {
        out.push_str(", above: 0pt, below: 0pt");
    }
    if has_inset {
        let _ = write!(out, ", inset: {}", format_insets(&inset));
    }
    out.push_str(")[");
    if let Some(settings) = powerpoint_line_settings {
        out.push_str(settings);
    }

    if is_pptx_blank_list_item(paragraph) {
        let height_pt: f64 = pptx_blank_list_item_advance_pt(paragraph);
        let _ = write!(out, "#block(height: {}pt)[]", format_f64(height_pt));
        out.push(']');
        return;
    }

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
            powerpoint_line_settings.is_some(),
        );
    } else {
        let runs: Vec<Run> = prepend_fixed_text_list_marker_run(
            &paragraph.style,
            list_style,
            &paragraph.runs,
            marker_text.to_string(),
        );
        write_fixed_text_list_item_paragraph(
            out,
            &paragraph.style,
            &runs,
            powerpoint_line_settings.is_some(),
        );
    }

    if align_str.is_some() {
        out.push(']');
    }
    out.push(']');
}

fn pptx_blank_list_item_advance_pt(paragraph: &Paragraph) -> f64 {
    let plain_advance_pt: f64 = powerpoint_line_box_pt(&paragraph.runs)
        .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT * 1.2);
    match paragraph.style.line_spacing {
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => plain_advance_pt * factor,
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) if points > 0.0 => points,
        _ => plain_advance_pt,
    }
}

fn write_fixed_text_ordered_marker_grid(
    out: &mut String,
    paragraph: &Paragraph,
    list_style: &EffectiveListStyle<'_>,
    marker_text: &str,
    hanging_indent_pt: f64,
    uses_powerpoint_line: bool,
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
    // PowerPoint anchors the marker at marL + indent; the fixed-width cell
    // reserves the remaining hanging indent for the gap before the body.
    let _ = write!(out, "#box(width: {}pt)[", format_f64(hanging_indent_pt),);
    generate_run(out, &marker_run);
    out.push(']');
    out.push_str("],\n");
    out.push('[');
    write_fixed_text_list_item_paragraph(out, &body_style, &trimmed_runs, uses_powerpoint_line);
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

/// Seat the first line at its own origin; a hanging-indent wrapper moves every
/// continuation line back to the paragraph's declared left margin.
pub(super) fn fixed_text_paragraph_inset(style: &ParagraphStyle) -> Insets {
    let left_inset: f64 = if fixed_text_paragraph_hanging_indent_pt(style).is_some() {
        fixed_text_paragraph_first_line_origin_pt(style)
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

fn write_fixed_text_list_item_paragraph(
    out: &mut String,
    style: &ParagraphStyle,
    runs: &[Run],
    uses_powerpoint_line: bool,
) {
    write_common_text_settings(out, runs, "");
    // The enclosing item already pins PowerPoint's full line box and zero
    // leading. Reintroducing Typst's 0.65em default here adds that leading to
    // every natural wrap, while item boundaries still use the fixed box
    // alone (issue #1335).
    if !uses_powerpoint_line {
        write_fixed_text_default_par_settings(out, style, runs, "");
    }
    let hanging_indent_pt: Option<f64> = fixed_text_paragraph_hanging_indent_pt(style);
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

    // A slide's inline list; PowerPoint splits Korean mid-word.
    generate_runs_with_tabs(
        out,
        runs,
        tab_stops.as_deref(),
        paragraph_default_tab_width_pt(style, DEFAULT_TAB_WIDTH_PT),
        EojeolWrap::Syllable,
    );
    out.push(']');
}

fn fixed_text_paragraph_first_line_origin_pt(style: &ParagraphStyle) -> f64 {
    let indent_left: f64 = style.indent_left.unwrap_or(0.0).max(0.0);
    let indent_first_line: f64 = style.indent_first_line.unwrap_or(0.0);

    if indent_first_line < 0.0 {
        (indent_left + indent_first_line).max(0.0)
    } else {
        indent_left
    }
}

pub(super) fn fixed_text_paragraph_hanging_indent_pt(style: &ParagraphStyle) -> Option<f64> {
    let indent_first_line: f64 = style.indent_first_line.unwrap_or(0.0);
    if indent_first_line >= -0.0001 {
        return None;
    }

    let indent_left: f64 = style.indent_left.unwrap_or(0.0).max(0.0);
    let hanging_indent_pt: f64 =
        (indent_left - fixed_text_paragraph_first_line_origin_pt(style)).max(0.0);
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
    // The rule states the document's own kerning decision; the empty text is
    // that answer rather than a claim about the runs' content, whose font list
    // is chosen per run (issue #628).
    write_text_params_for_text(out, &style, "");
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
        pair_kerning: (left.pair_kerning == right.pair_kerning)
            .then_some(left.pair_kerning)
            .flatten(),
        ligatures: (left.ligatures == right.ligatures)
            .then_some(left.ligatures)
            .flatten(),
        ..TextStyle::default()
    }
}

/// The formatting one list level's labels take before the level's own run
/// properties.
struct LevelLabelStyles {
    /// What every label at the level shares.
    common: Option<TextStyle>,
    /// Each ordered item's own formatting by the number Typst gives it, when
    /// the level's Word paragraph marks disagree (issue #2337); empty
    /// otherwise.
    by_number: Vec<(u32, TextStyle)>,
}

impl LevelLabelStyles {
    /// The formatting the label numbered `number` takes.
    fn for_number(&self, number: u32) -> Option<&TextStyle> {
        self.by_number
            .iter()
            .find(|(label_number, _)| *label_number == number)
            .map(|(_, style)| style)
            .or(self.common.as_ref())
    }
}

/// The label formatting of `items` at `level`.
///
/// Typst gives a list level one numbering function, which learns only each
/// item's number, so an ordered level whose items' Word marks disagree keys
/// each mark's formatting by that number.
fn level_label_styles(
    items: &[crate::ir::ListItem],
    level: u32,
    kind: ListKind,
) -> LevelLabelStyles {
    let level_items: Vec<&crate::ir::ListItem> =
        items.iter().filter(|item| item.level == level).collect();
    // TODO(#2384): Word labels a bullet with its mark's face but not, on
    // docxide-pdf's `alfies_arc_adult_safeguarding_policy` and
    // `streamnet_steering_committee_meeting`, with the mark's bold. Until a
    // native probe settles which mark properties a bullet takes, bullets keep
    // the formatting their runs share.
    let marks: Option<Vec<&TextStyle>> = (kind == ListKind::Ordered)
        .then(|| {
            level_items
                .iter()
                .map(|item| {
                    item.content
                        .first()
                        .and_then(|paragraph| paragraph.style.numbering_label_style.as_deref())
                })
                .collect::<Option<Vec<&TextStyle>>>()
        })
        .flatten()
        .filter(|marks| !marks.is_empty());
    let Some(marks) = marks else {
        return LevelLabelStyles {
            common: common_list_level_text_style(level_items.iter().copied()),
            by_number: Vec::new(),
        };
    };
    let common: Option<TextStyle> = intersect_label_styles(marks.iter().copied());
    // Compared as `intersect_text_style` keeps them: properties no label
    // shares are not worth a branch per item.
    let agree = marks.windows(2).all(|pair| {
        intersect_text_style(pair[0], pair[0]) == intersect_text_style(pair[1], pair[1])
    });
    if agree {
        return LevelLabelStyles {
            common,
            by_number: Vec::new(),
        };
    }
    // Typst numbers the first item from the list's `start`, an item with its
    // own start from that, and every other item one past the previous.
    let mut by_number: Vec<(u32, TextStyle)> = Vec::new();
    let mut number: u32 = level_items
        .first()
        .and_then(|item| item.start_at)
        .unwrap_or(1);
    for (item, mark) in level_items.iter().zip(marks) {
        number = item.start_at.unwrap_or(number);
        match by_number
            .iter_mut()
            .find(|(label_number, _)| *label_number == number)
        {
            // A restart onto a number already used: the function cannot tell
            // the two labels apart, so they keep what they share.
            Some((_, shared)) => {
                tracing::debug!(number, level, "List label number repeats within one level");
                *shared = intersect_text_style(shared, mark);
            }
            None => by_number.push((number, mark.clone())),
        }
        number = number.saturating_add(1);
    }
    LevelLabelStyles { common, by_number }
}

/// What `styles` share, or `None` when that is nothing a label can show.
fn intersect_label_styles<'a>(
    mut styles: impl Iterator<Item = &'a TextStyle>,
) -> Option<TextStyle> {
    let first_style = styles.next()?.clone();
    let common_style = styles.fold(first_style, |common, style| {
        intersect_text_style(&common, style)
    });
    has_text_properties(&common_style).then_some(common_style)
}

/// The formatting the visible runs of a level's `items` agree on.
fn common_list_level_text_style<'a>(
    items: impl Iterator<Item = &'a crate::ir::ListItem>,
) -> Option<TextStyle> {
    intersect_label_styles(
        items
            .flat_map(|item| item.content.iter())
            .flat_map(|paragraph| paragraph.runs.iter())
            .filter(|run| {
                run.footnote.is_none()
                    && !run.text.is_empty()
                    && run.text != PPTX_BLANK_LIST_ITEM_SENTINEL
            })
            .map(|run| &run.style),
    )
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
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) => {
            Some((points - font_size_pt).max(0.0))
        }
        _ => None,
    }
}

/// The line advance a list's items take between them, which is the same
/// `par(leading:)` its wrapped lines take within one item.
///
/// `fixed_text_list_line_gap_pt` states it only when the paragraph declares a
/// line spacing that widens the line; otherwise the leading in force is
/// whatever `write_par_settings` emits, or — when it emits nothing — Typst's
/// own 0.65em default. All three have to be reproduced here, because the gap
/// between two items is emitted rather than inherited (issue #934).
fn fixed_text_list_leading_pt(style: &ParagraphStyle, list: &List) -> f64 {
    if let Some(gap) = fixed_text_list_line_gap_pt(style, list).filter(|gap| *gap > 0.0) {
        return gap;
    }
    let font_size_pt: f64 = fixed_text_list_font_size_pt(list);
    match style.line_spacing {
        // `write_par_settings` scales Typst's default by the declared factor.
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => {
            (font_size_pt * TYPST_DEFAULT_LEADING_EM * factor).max(0.0)
        }
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) => points.max(0.0),
        _ => font_size_pt * TYPST_DEFAULT_LEADING_EM,
    }
}

/// Typst's own `par(leading:)` default, in em.
const TYPST_DEFAULT_LEADING_EM: f64 = 0.65;

/// A slide list's PowerPoint line box: the `#set text(top-edge:/bottom-edge:)`
/// settings that give each item the same line PowerPoint gives a text box's
/// paragraphs, and the advance in points that box works out to.
///
/// The two travel together because the gap between items is emitted rather
/// than inherited, and it has to be the same quantity the box states
/// (issue #934).
fn fixed_text_list_powerpoint_line(list: &List) -> Option<(String, f64)> {
    let paragraph: &Paragraph = list.items.first()?.content.first()?;
    let settings: String = powerpoint_line_height_settings(&paragraph.runs, &paragraph.style)?;
    // `powerpoint_line_height_settings` resizes its edges by the declared
    // `a:lnSpc`: percentages scale the plain line, while absolute spacing
    // replaces it with the point advance. The value paired with those settings
    // must follow the same rule.
    let plain_advance_pt: f64 = powerpoint_line_box_pt(&paragraph.runs)?;
    let advance_pt: f64 = match paragraph.style.line_spacing {
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => plain_advance_pt * factor,
        Some(LineSpacing::Exact(points) | LineSpacing::AtLeast(points)) if points > 0.0 => points,
        _ => plain_advance_pt,
    };
    Some((settings, advance_pt))
}

/// What PowerPoint puts between two items on top of the line advance: the
/// `a:spcAft` of the one above plus the `a:spcBef` of the one below. The two
/// are ADDED — PowerPoint does not collapse them to the larger the way CSS
/// margins do.
///
/// Both values used to reach only the block wrapped around the whole list, so
/// a list whose every paragraph declared them got them once, at its outer
/// edges, and nothing between its items (issue #928).
///
/// One gap is emitted for the whole level, so a list whose boundaries disagree
/// returns `None` and keeps the unspaced advance rather than picking one of
/// them; `fixed_text_list_boundary_gap_pt` states each boundary's own.
/// The `a:spcAft` of the item above plus the `a:spcBef` of the one below —
/// what that one boundary asks for, whatever its neighbours ask for.
///
/// The shared value above is what a `#stack`'s uniform `spacing:` can carry;
/// the emitted path states each boundary separately, so a list whose items
/// disagree keeps every one of them instead of falling back to none
/// (issue #934).
fn fixed_text_list_boundary_gap_pt(
    above: &crate::ir::ListItem,
    below: &crate::ir::ListItem,
) -> f64 {
    let after: f64 = above
        .content
        .last()
        .and_then(|paragraph| paragraph.style.space_after)
        .unwrap_or(0.0);
    let before: f64 = below
        .content
        .first()
        .and_then(|paragraph| paragraph.style.space_before)
        .unwrap_or(0.0);
    (after + before).max(0.0)
}

fn fixed_text_list_paragraph_gap_pt(list: &List) -> Option<f64> {
    let boundary_gap = |above: &crate::ir::ListItem, below: &crate::ir::ListItem| -> Option<f64> {
        let after: f64 = above.content.last()?.style.space_after.unwrap_or(0.0);
        let before: f64 = below.content.first()?.style.space_before.unwrap_or(0.0);
        Some((after + before).max(0.0))
    };
    let mut boundaries = list
        .items
        .windows(2)
        .map(|pair| boundary_gap(&pair[0], &pair[1]));
    let first: f64 = boundaries.next()??;
    (first > 0.0001 && boundaries.all(|gap| gap.is_some_and(|gap| f64_approx_eq(gap, first))))
        .then_some(first)
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
    runs: &[Run],
    line_gap_pt: Option<f64>,
) {
    write_line_box_settings(out, style.line_box);
    if let Some(gap) = line_gap_pt.filter(|gap| *gap > 0.0) {
        let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(gap));
    } else {
        write_par_settings(out, style, runs, false);
        return;
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
        write_east_asian_justification_limits(out, runs);
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
        ListKind::Ordered => ordered_marker(style.numbering_pattern, number),
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
        inline_box: None,
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
    if fixed_text_paragraph_hanging_indent_pt(style).is_some() {
        // The tab carries the whole gap to the indent, so the space
        // `fixed_text_list_marker` puts after the glyph is a second separator.
        // It pushed the text 2.59pt past the indent on the audited deck, which
        // is enough to move a wrap point (issue #685).
        return prepend_marker_run(
            runs,
            format!("{}\t", marker_text.trim_end()),
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
        inline_box: None,
    }
}

fn first_visible_char_is_whitespace(runs: &[Run]) -> bool {
    runs.iter()
        .find_map(|run| run.text.chars().next())
        .is_some_and(char::is_whitespace)
}

/// The item's number formats the first counter. Like Typst's `numbering`
/// given one value, later counters are dropped with their literal text; a
/// pattern without counters is static text.
fn ordered_marker(pattern: Option<&NumberingPattern>, number: u32) -> String {
    let Some(pattern) = pattern else {
        return format!("{number}.");
    };
    let Some((prefix, symbol)) = pattern.pieces.first() else {
        return pattern.suffix.clone();
    };
    let counter: String = match symbol {
        'a' => alpha_marker(number, false),
        'A' => alpha_marker(number, true),
        'i' => roman_marker(number, false),
        'I' => roman_marker(number, true),
        _ => number.to_string(),
    };
    format!("{prefix}{counter}{}", pattern.suffix)
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

/// Carry a Word boundary's collapsed gap or a PowerPoint item's own gap.
///
/// `common_list_level_spacing` only returns a value when *every* boundary at
/// the level agrees, so a list whose items declare different gaps got none at
/// all — 04_training_deck_ko's outline alternates 6pt and 10pt `a:spcAft` and
/// lost every one of them, which cost up to 18.8pt by the last bullet
/// (issue #524). Emitting the gap inside the item that owns it also survives
/// nesting, where a shared `spacing:` cannot reach.
fn write_list_item_trailing_gap(
    out: &mut String,
    item: &crate::ir::ListItem,
    has_uniform_spacing: bool,
    wrapper_spans_full_line: bool,
    word_boundary_gap: Option<f64>,
) {
    if has_uniform_spacing && word_boundary_gap.is_none() {
        return;
    }
    // PowerPoint's boxless path carries line height through shared spacing.
    // Word's explicit boundary gap also survives unavailable font metrics.
    if !wrapper_spans_full_line && word_boundary_gap.is_none() {
        return;
    }
    let gap: f64 = word_boundary_gap.unwrap_or_else(|| paragraph_space_after(item));
    if gap <= 0.0001 {
        return;
    }
    let _ = write!(out, "#v({}pt)", format_f64(gap));
}

/// Emits one of an item's paragraphs.
///
/// Each paragraph decides its own Hangul breaking from its own style — its
/// effective `w:wordWrap` and its indents — but against the *list's* fixed
/// line box: whatever [`generate_list_with_spacing_model`] put in force on the
/// wrapper is what a framed eojeol has to restore inside itself (issue #626).
///
/// `tab_origin`, a Typst length expression, is where the paragraph's first
/// line starts from the margin. With it, a tab advances to the paragraph's
/// next stop, as in a Word body paragraph (issue #2223); without it, a tab
/// stays literal whitespace, which Typst collapses.
fn write_list_paragraph_runs(
    out: &mut String,
    para: &Paragraph,
    wrap: &ListEojeolWrap,
    tab_origin: Option<&str>,
    ctx: &mut super::GenCtx,
) -> Result<(), ConvertError> {
    let eojeol_wrap = paragraph_eojeol_wrap(
        wrap.breaks_hangul_at_eojeol,
        &para.style,
        wrap.line_box_em,
        wrap.available_measure_pt,
    );
    let Some(tab_origin) = tab_origin else {
        return super::text::generate_runs_with_context(out, &para.runs, eojeol_wrap, ctx);
    };
    let default_tab_width_pt =
        paragraph_default_tab_width_pt(&para.style, ctx.default_tab_width_pt);
    super::text::generate_runs_with_tabs_from_origin_context(
        out,
        &para.runs,
        para.style.tab_stops.as_deref(),
        default_tab_width_pt,
        tab_origin,
        eojeol_wrap,
        ctx,
    )
}

/// Opens a scope justifying one list paragraph whose own alignment asks for
/// it, as a body paragraph's does, and reports whether it did; the caller
/// closes it with `]` after the paragraph. The scope keeps the setting off the
/// item's nested list, whose paragraphs declare their own alignment (issue
/// #2336).
///
/// The lines break where an unjustified item's do, first-fit at natural
/// width. On docxide-pdf's scraped `k4_regulations_synthesis_document` that
/// keeps Word's break on every numbered line, where Typst's optimized breaker
/// rewrapped 25 of page 1's 48 lines away from Word's.
/// TODO(#2063): body paragraphs keep the optimized breaker; settle one rule
/// for both with a native probe.
fn write_list_paragraph_justification_open(out: &mut String, para: &Paragraph) -> bool {
    if para.style.alignment != Some(Alignment::Justify) {
        return false;
    }
    out.push_str("#[\n");
    write_justification_settings(out, &para.runs, true);
    true
}

/// The Typst binding [`write_word_numbering_suffix_advance`] leaves holding
/// where the first body line starts, measured from the margin.
const WORD_BODY_ORIGIN: &str = "word_body_origin";

/// The suffix and indent by which [`write_word_numbering_suffix_advance`]
/// places an item's first body line, or `None` when it leaves that line at the
/// list's nominal body origin.
fn word_numbering_suffix_placement(
    kind: ListKind,
    paragraph_style: &ParagraphStyle,
    hangs_into_margin: bool,
) -> Option<(ListMarkerSuffix, ListIndentGeometry)> {
    let suffix = paragraph_style.numbering_suffix?;
    if suffix == ListMarkerSuffix::Tab && kind != ListKind::Ordered {
        return None;
    }
    Some((
        suffix,
        paragraph_list_indent(paragraph_style, hangs_into_margin)?,
    ))
}

/// What a list's items need to decide their Hangul line breaking (issue #626).
#[derive(Clone, Copy, Default)]
pub(super) struct ListEojeolWrap {
    /// Whether the enclosing page is a Word flow page. False for a slide,
    /// whose list keeps PowerPoint's own mid-word breaking.
    pub(super) breaks_hangul_at_eojeol: bool,
    /// The fixed `(top-edge, bottom-edge)` the list wrapper emits, in em.
    pub(super) line_box_em: Option<(f64, f64)>,
    /// The width one line of the list has, in points, before the item's own
    /// indents.
    pub(super) available_measure_pt: Option<f64>,
    /// PowerPoint's raw and layout baseline seats for absolute slide snapping.
    pub(super) baseline_snap: Option<PowerPointBaselineSnap>,
    /// Whether a marker or body may start left of the list's origin, as a
    /// Word body list hanging into the margin does (issue #2233).
    pub(super) hangs_into_margin: bool,
}

/// Word's suffix positions only the first body line; continuation lines keep
/// the declared indent. Follow docxide-pdf's `list_label::text_hanging`: measure
/// the styled marker, then use a space, no gap, or the source/default tab.
///
/// A body containing tabs comes in as `tabbed_body`, emitted with
/// [`WORD_BODY_ORIGIN`] as its tab origin. It is written inside this scope,
/// after the binding, because a body tab starts from wherever the suffix left
/// the pen: past a wide marker that is the stop the suffix reached.
#[allow(clippy::too_many_arguments)]
fn write_word_numbering_suffix_advance(
    out: &mut String,
    style: &EffectiveListStyle<'_>,
    fallback_marker_style: Option<&TextStyle>,
    paragraph_style: &ParagraphStyle,
    numbers: &[u32],
    default_tab_width_pt: f64,
    tabbed_body: Option<&str>,
    hangs_into_margin: bool,
) {
    let Some((suffix, indent)) =
        word_numbering_suffix_placement(style.kind, paragraph_style, hangs_into_margin)
    else {
        return;
    };
    let (bullet_text, explicit_marker_style) = if style.kind == ListKind::Unordered {
        renderable_unordered_marker(style.marker_text.unwrap_or("•"), style.marker_style)
    } else {
        (String::new(), style.marker_style.cloned())
    };
    let marker_style = merge_marker_style(fallback_marker_style, explicit_marker_style.as_ref());
    let marker_style = marker_style
        .as_ref()
        .filter(|style| has_text_properties(style));
    let body_origin_pt = indent.marker_origin_pt + indent.marker_width_pt;
    let interval = paragraph_default_tab_width_pt(paragraph_style, default_tab_width_pt);
    out.push_str("#context { let word_list_marker = [");
    if style.kind == ListKind::Unordered {
        write_unordered_list_marker_content(out, &bullet_text, marker_style);
    } else {
        if let Some(marker_style) = marker_style {
            out.push_str("#text(");
            write_text_params(out, marker_style);
            out.push_str(")[");
        }
        let values = if style.full_numbering {
            numbers
        } else {
            &numbers[numbers.len() - 1..]
        };
        out.push('#');
        write_numbering_label(out, style.numbering_pattern, MarkerNumbers::Values(values));
        if marker_style.is_some() {
            out.push(']');
        }
    }
    if suffix != ListMarkerSuffix::Tab {
        let _ = write!(
            out,
            "]; let word_marker_end = {}pt + measure(word_list_marker).width; ",
            format_f64(indent.marker_origin_pt)
        );
        let body_start = if suffix == ListMarkerSuffix::Space {
            // Native Word uses Arial for this generated separator at the
            // marker's size, even when the marker declares another family.
            let separator_style = TextStyle {
                font_family: Some("Arial".to_string()),
                font_size: marker_style.and_then(|style| style.font_size),
                ..TextStyle::default()
            };
            out.push_str("let word_list_space = [#text(");
            write_text_params(out, &separator_style);
            out.push_str(")[");
            out.push_str(r#"#" ""#);
            out.push_str("]]; ");
            "word_marker_end + measure(word_list_space).width"
        } else {
            "word_marker_end"
        };
        if let Some(body) = tabbed_body {
            let _ = write!(
                out,
                "let {WORD_BODY_ORIGIN} = {body_start}; h({WORD_BODY_ORIGIN} - {}pt); [{body}] }}",
                format_f64(body_origin_pt)
            );
        } else {
            let _ = write!(out, "h({body_start} - {}pt) }}", format_f64(body_origin_pt));
        }
        return;
    }
    let _ = write!(
        out,
        "]; let word_marker_end = {}pt + measure(word_list_marker).width; ",
        format_f64(indent.marker_origin_pt)
    );
    if tabbed_body.is_some() {
        let _ = write!(
            out,
            "let {WORD_BODY_ORIGIN} = {}pt; ",
            format_f64(body_origin_pt)
        );
    }
    let _ = write!(
        out,
        "if word_marker_end >= {}pt - 0.01pt {{ let word_body_tab = ",
        format_f64(body_origin_pt)
    );
    let mut stops: Vec<f64> = paragraph_style
        .tab_stops
        .as_deref()
        .unwrap_or_default()
        .iter()
        .filter(|stop| stop.alignment == TabAlignment::Left && stop.position.is_finite())
        .map(|stop| stop.position)
        .collect();
    stops.sort_by(f64::total_cmp);
    stops.dedup();
    let stop_count = stops.len();
    for stop in stops {
        let _ = write!(
            out,
            "if word_marker_end + 0.01pt < {}pt {{ {}pt }} else {{ ",
            format_f64(stop),
            format_f64(stop)
        );
    }
    let _ = write!(
        out,
        "(calc.floor(word_marker_end / {}pt) + 1) * {}pt",
        format_f64(interval),
        format_f64(interval)
    );
    for _ in 0..stop_count {
        out.push_str(" }");
    }
    if let Some(body) = tabbed_body {
        let _ = write!(
            out,
            "; {WORD_BODY_ORIGIN} = calc.max({origin}pt, word_body_tab); h({WORD_BODY_ORIGIN} - {origin}pt) }}; [{body}] }}",
            origin = format_f64(body_origin_pt)
        );
    } else {
        let _ = write!(
            out,
            "; h(calc.max(0pt, word_body_tab - {}pt)) }} }}",
            format_f64(body_origin_pt)
        );
    }
}

/// Recursively generate list items, grouping consecutive items at the same or deeper level.
#[allow(clippy::too_many_arguments)]
fn generate_list_items(
    out: &mut String,
    list: &List,
    items: &[crate::ir::ListItem],
    base_level: u32,
    wrapper_spans_full_line: bool,
    has_uniform_spacing: bool,
    per_item_gaps: bool,
    word_boundary_gaps: bool,
    following_item: Option<&crate::ir::ListItem>,
    eojeol_wrap: &ListEojeolWrap,
    parent_numbers: &[u32],
    ctx: &mut super::GenCtx,
) -> Result<(), ConvertError> {
    let style = list_style_for_level(list, base_level);
    let (_, item_func) = list_funcs(style.kind);
    let labels = level_label_styles(items, base_level, style.kind);
    let uses_word_suffixes = eojeol_wrap.breaks_hangul_at_eojeol
        && common_list_level_indent(items, base_level, eojeol_wrap.hangs_into_margin).is_some()
        && items
            .iter()
            .filter(|item| item.level == base_level)
            .any(|item| {
                item.content.first().is_some_and(|paragraph| {
                    match paragraph.style.numbering_suffix {
                        Some(ListMarkerSuffix::Tab) => style.kind == ListKind::Ordered,
                        Some(ListMarkerSuffix::Space | ListMarkerSuffix::Nothing) => true,
                        None => false,
                    }
                })
            });
    if uses_word_suffixes {
        tracing::debug!(
            base_level,
            default_tab_width_pt = ctx.default_tab_width_pt,
            "Enabled Word numbering suffix placement"
        );
    }
    let mut number = items.first().and_then(|item| item.start_at).unwrap_or(1);
    let mut i = 0;
    while i < items.len() {
        let item = &items[i];
        number = item.start_at.unwrap_or(number);
        let mut item_numbers = parent_numbers.to_vec();
        if style.kind == ListKind::Ordered && eojeol_wrap.breaks_hangul_at_eojeol {
            item_numbers.push(number);
        }
        let word_boundary_gap: Option<f64> = word_boundary_gaps
            .then(|| items.get(i + 1).or(following_item))
            .flatten()
            .map(|next| paragraph_space_after(item).max(paragraph_space_before(next)));
        let _ = write!(out, "  {item_func}");
        if style.kind == ListKind::Ordered
            && i > 0
            && let Some(start_at) = item.start_at
        {
            let _ = write!(out, "({start_at})");
        }
        out.push('[');
        if let Some(snap) = eojeol_wrap.baseline_snap {
            snap.write_open(out);
        }
        let mut paragraphs = item.content.iter();
        if let Some(paragraph) = paragraphs.next() {
            let justified = write_list_paragraph_justification_open(out, paragraph);
            // Word counts a body tab's stops from the margin, as in any
            // paragraph, so the body's tabs measure from where its first line
            // starts. A slide list keeps its own tab handling.
            let body_origin_pt: Option<f64> = (eojeol_wrap.breaks_hangul_at_eojeol
                && super::text::paragraph_contains_tabs(&paragraph.runs))
            .then(|| paragraph_list_indent(&paragraph.style, eojeol_wrap.hangs_into_margin))
            .flatten()
            .map(|indent| indent.marker_origin_pt + indent.marker_width_pt);
            if uses_word_suffixes
                && word_numbering_suffix_placement(
                    style.kind,
                    &paragraph.style,
                    eojeol_wrap.hangs_into_margin,
                )
                .is_some()
            {
                let tabbed_body: Option<String> = match body_origin_pt {
                    Some(_) => {
                        let mut body = String::new();
                        write_list_paragraph_runs(
                            &mut body,
                            paragraph,
                            eojeol_wrap,
                            Some(WORD_BODY_ORIGIN),
                            ctx,
                        )?;
                        Some(body)
                    }
                    None => None,
                };
                write_word_numbering_suffix_advance(
                    out,
                    &style,
                    labels.for_number(number),
                    &paragraph.style,
                    &item_numbers,
                    ctx.default_tab_width_pt,
                    tabbed_body.as_deref(),
                    eojeol_wrap.hangs_into_margin,
                );
                if tabbed_body.is_none() {
                    write_list_paragraph_runs(out, paragraph, eojeol_wrap, None, ctx)?;
                }
            } else {
                let tab_origin: Option<String> =
                    body_origin_pt.map(|origin| format!("{}pt", format_f64(origin)));
                write_list_paragraph_runs(out, paragraph, eojeol_wrap, tab_origin.as_deref(), ctx)?;
            }
            if justified {
                out.push(']');
            }
        }
        for paragraph in paragraphs {
            let justified = write_list_paragraph_justification_open(out, paragraph);
            write_list_paragraph_runs(out, paragraph, eojeol_wrap, None, ctx)?;
            if justified {
                out.push(']');
            }
        }
        if eojeol_wrap.baseline_snap.is_some() {
            out.push(']');
        }

        if item.level == base_level {
            let nested_start = i + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].level > base_level {
                nested_end += 1;
            }

            if nested_end > nested_start {
                let mut child_start = nested_start;
                while child_start < nested_end {
                    // A skipped Word level has no physical item to provide its
                    // style or indent. Start at the actual level and split when
                    // a later child returns to a shallower level (#2087).
                    let child_level = items[child_start].level;
                    if child_level > base_level.saturating_add(1) {
                        tracing::trace!(base_level, child_level, "Rendering skipped list level");
                    }
                    let mut child_end = child_start + 1;
                    while child_end < nested_end && items[child_end].level >= child_level {
                        child_end += 1;
                    }
                    let nested_gap = if child_start != nested_start {
                        // A second child subtree must remain a block; placing
                        // its list inline after the first adds a parent strut.
                        Some(0.0)
                    } else if word_boundary_gaps {
                        word_boundary_gap
                    } else {
                        (per_item_gaps && wrapper_spans_full_line)
                            .then(|| paragraph_space_after(item))
                            .filter(|gap| *gap > 0.0001)
                    };
                    if child_start == nested_start && nested_gap.is_none() {
                        write_list_item_trailing_gap(
                            out,
                            item,
                            has_uniform_spacing,
                            wrapper_spans_full_line,
                            word_boundary_gap,
                        );
                    }
                    let nested_style = list_style_for_level(list, child_level);
                    let nested_labels = level_label_styles(
                        &items[child_start..child_end],
                        child_level,
                        nested_style.kind,
                    );
                    // Word indents are absolute from the margin, but a nested
                    // Typst list is laid out inside the parent item's body:
                    // subtract the parent's text origin so the child marker
                    // lands at its absolute position (issue #356).
                    let parent_indent = common_list_level_indent(
                        &items[i..=i],
                        base_level,
                        eojeol_wrap.hangs_into_margin,
                    );
                    let indent = common_list_level_indent(
                        &items[child_start..child_end],
                        child_level,
                        eojeol_wrap.hangs_into_margin,
                    )
                    .map(|mut child| {
                        if let Some(parent) = parent_indent {
                            child.marker_origin_pt = (child.marker_origin_pt
                                - (parent.marker_origin_pt + parent.marker_width_pt))
                                .max(0.0);
                        }
                        child
                    });
                    // PowerPoint's a:spcAft belongs to the item that declares it.
                    // A shared nested `spacing:` cannot carry the last child's
                    // gap across the return to its parent level, so retain every
                    // child gap on that path (issue #659).
                    let spacing_pt = if word_boundary_gaps {
                        Some(0.0)
                    } else if per_item_gaps {
                        None
                    } else {
                        common_list_level_spacing(
                            &items[child_start..child_end],
                            child_level,
                            wrapper_spans_full_line,
                        )
                    };
                    let nested_start_at = items[child_start].start_at;
                    if let Some(gap) = nested_gap {
                        // `#v` immediately after inline parent text creates a new
                        // line of its own before the nested list. Block spacing
                        // adds only the resolved paragraph gap to the normal
                        // parent-to-child line advance (issue #659).
                        let _ =
                            writeln!(out, " #block(width: 100%, above: {}pt)[", format_f64(gap));
                    }
                    write_list_open(
                        out,
                        if nested_gap.is_some() { "#" } else { " #" },
                        &nested_style,
                        &nested_labels,
                        ListOpenOptions {
                            indent,
                            spacing_pt,
                            start_at: nested_start_at,
                            baseline_snap: None,
                        },
                    );
                    generate_list_items(
                        out,
                        list,
                        &items[child_start..child_end],
                        child_level,
                        wrapper_spans_full_line,
                        spacing_pt.is_some() || !per_item_gaps,
                        per_item_gaps,
                        word_boundary_gaps,
                        items.get(child_end).or(following_item),
                        eojeol_wrap,
                        &item_numbers,
                        ctx,
                    )?;
                    out.push(')');
                    if nested_gap.is_some() {
                        out.push_str("\n]");
                    }
                    child_start = child_end;
                }
                i = nested_end;
            } else {
                write_list_item_trailing_gap(
                    out,
                    item,
                    has_uniform_spacing,
                    wrapper_spans_full_line,
                    word_boundary_gap,
                );
                i += 1;
            }
        } else {
            write_list_item_trailing_gap(
                out,
                item,
                has_uniform_spacing,
                wrapper_spans_full_line,
                word_boundary_gap,
            );
            i += 1;
        }

        out.push_str("],\n");
        number = number.saturating_add(1);
    }
    Ok(())
}
