use std::fmt::Write;

use unicode_normalization::UnicodeNormalization;

use crate::render::font_subst;

use super::*;

/// Word's default tab stop interval (0.5 inch = 36pt).
pub(super) const DEFAULT_TAB_WIDTH_PT: f64 = 36.0;
/// East Asian Word's default tab stop (800 twips) when settings.xml omits
/// `w:defaultTabStop`.
pub(super) const EAST_ASIAN_DEFAULT_TAB_WIDTH_PT: f64 = 40.0;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';
/// In-text marker the PPTX parser places between a Hangul syllable and
/// following terminal punctuation (issue #438); never emitted literally.
const HANGUL_KINSOKU_BREAK_CHAR: char = '\u{200B}';
/// In-text marker the DOCX parser places at an East Asian/Latin boundary that
/// carries no literal space (issue #521); never emitted literally.
const EAST_ASIAN_AUTO_SPACE_CHAR: char = '\u{E001}';
/// Word's automatic space at such a boundary, as a fraction of the run's font
/// size. Measured as exactly a quarter em on a native export at two sizes.
const EAST_ASIAN_AUTO_SPACE_EM: f64 = 0.25;

/// The character the auto space is drawn with: a no-break space, so the
/// boundary it marks stays as unbreakable as the eojeol frame kept it.
const EAST_ASIAN_AUTO_SPACE_GLYPH: &str = "\u{00A0}";

/// How wide Word lets an expandable gap of a justified line grow before it
/// levels every gap on that line to one common width, in ems of the gap's own
/// text (issue #1053).
const EAST_ASIAN_JUSTIFIED_GAP_CEILING_EM: f64 = 0.5;

/// PowerPoint snaps nominal glyph advances to this grid before accumulating
/// them into a line. Typst keeps the font's exact fractional advances, so the
/// sub-point error otherwise compounds across long slide lines (issue #661).
const POWERPOINT_ADVANCE_GRID_PT: f64 = 0.125;

/// Excel rounds every glyph advance to this grid before accumulating it into a
/// cell line.
///
/// Measured on the ten committed XLSX golden mocks: reconstructing each
/// line from `round_half_up(hmtx_advance x size)` reproduces all 3,656 glyph
/// origins of their first pages — Arial, Arial Bold, Malgun Gothic and Malgun
/// Gothic Bold at 9, 10 and 14pt — to within the 0.23pt the export's own
/// integer `TJ` offsets can express. Rounding the *running pen position*
/// instead, the other model that puts origins on whole points, misses: on
/// `Region` at Malgun Gothic 12 it predicts 53, 60, 67, 74, 77, 84 where the
/// export prints 53, 60, 66, 73, 76, 83 (issue #1088).
///
/// The face's own fractional advances are therefore ~5% narrow on a long line
/// (19.5pt over 99 glyphs of 9pt Arial), because the rounding is biased upward
/// for the many advances that land just below a half point.
///
/// The whole-point grid is the same quantum Excel's column widths take
/// (issue #621). A fitted sheet applies this quantum in its declared
/// coordinate system and scales the result afterwards (issue #1238). Word
/// does not share it — no `docx` golden mock's export puts its origins on
/// whole points.
const SHEET_ADVANCE_GRID_PT: f64 = 1.0;

/// Emit the ligature rule every PowerPoint slide follows.
///
/// PowerPoint does not apply the OpenType `liga`/`clig` features. DrawingML has
/// no run property that asks for them — `a:rPr` carries `@spc`, `@kern` and
/// `@baseline` but nothing for ligation — and a native macOS PowerPoint export
/// of `customGeo.pptx` page 46 draws `testing:` and `setting:` as discrete
/// glyphs with a dotted `i`. Typst ligates by default, so the same runs came
/// out with the face's fused `ti`/`tti` form and a dotless `i`, which also
/// pulled every glyph after the ligature left by the advance the merge saved —
/// 1.4pt by the trailing colon of `setting:` (issue #1058).
///
/// Stated once for the document rather than on each run: it holds for every
/// slide run including the emission sites that cannot name their own text, and
/// only PowerPoint produces fixed pages, so nothing else can inherit it.
/// Typst's `ligatures: false` switches off `liga` and `clig` alone, so the
/// required-ligature and contextual features that Arabic and Indic shaping
/// depend on (`rlig`, `ccmp`, `calt`) are untouched.
pub(super) fn write_powerpoint_ligature_state(out: &mut String) {
    let _ = writeln!(out, "#set text(ligatures: false)");
}

/// Emit the contextual Typst helpers used by fixed-page (PowerPoint) text.
///
/// A word stays one shaped item, preserving its kerning and PDF text mapping.
/// Only its horizontal scale and layout width change by the
/// difference between exact and independently grid-rounded nominal glyph
/// widths. Pair kerning therefore remains part of the shaped word instead of
/// being lost to one-box-per-glyph output. Spaces remain real text characters
/// and carry a weak correction, so they keep both extraction and line-break
/// behavior. The word box measures and restores the active text edge below
/// the baseline, so it occupies the same line box as unboxed text and cannot
/// disturb vertical centring.
pub(super) fn write_powerpoint_advance_grid_helpers(out: &mut String) {
    let _ = writeln!(
        out,
        r#"#let o2p-pptx-advance-grid = {POWERPOINT_ADVANCE_GRID_PT}pt
#let o2p-pptx-word(body, glyphs) = context {{
  let natural = measure(body).width
  let nominal = glyphs.map(glyph => measure(glyph).width).sum()
  let snapped = glyphs.map(glyph => calc.round(measure(glyph).width / o2p-pptx-advance-grid) * o2p-pptx-advance-grid).sum()
  let target = natural + snapped - nominal
  let baseline-body = text(bottom-edge: "baseline", body)
  let seat = measure(body).height - measure(baseline-body).height
  if natural == 0pt {{ body }} else {{
    box(inset: (bottom: seat), baseline: seat)[#text(bottom-edge: "baseline")[#scale(x: target / natural * 100%, origin: left, body)]] + h(target - natural)
  }}
}}
#let o2p-pptx-space() = context {{
  let natural = measure(" ").width
  let target = calc.round(natural / o2p-pptx-advance-grid) * o2p-pptx-advance-grid
  [#" "; #h(target - natural, weak: true)]
}}"#
    );
}

/// The auto space sized against the *run*, not the paragraph. The width is
/// emitted in points wherever the run states a size, because an `em` in the
/// wrapper would resolve against the paragraph's default size instead — 11pt
/// where the run is 10.5pt, which is 0.12pt too wide at every boundary.
///
/// It is a no-break space carrying its own `spacing`, not an `#h()` spacer,
/// because a justified line has to be able to stretch it. Typst's justifier
/// reaches a line's *glyphs* alone — `Line::stretchability` sums the text
/// items and skips every spacing item — so a spacer stays rigid and the whole
/// of a line's stretch demand lands in the word spaces: 8.70pt against Word's
/// 6.81pt on the line issue #1193 measured, with the auto spaces left at
/// their quarter em against Word's 6.80pt.
///
/// `text(spacing:)` restates the space's own advance, so the quarter em is
/// still the width the line breaks on and nothing re-wraps; and U+00A0 is
/// Unicode line-break class GL, which forbids a break on either side of it,
/// which is what keeps the mid-word break that #521's frame existed to
/// prevent from reopening. See [`write_east_asian_justification_limits`] for
/// the ceiling that makes the stretch match Word's.
fn east_asian_auto_space(run: &Run) -> String {
    let width: String = match run.style.font_size {
        Some(size) => format!("{}pt", format_f64(size * EAST_ASIAN_AUTO_SPACE_EM)),
        None => format!("{EAST_ASIAN_AUTO_SPACE_EM}em"),
    };
    let mut params: String = String::new();
    write_text_params_for_run(
        &mut params,
        &run.style,
        EAST_ASIAN_AUTO_SPACE_GLYPH,
        EAST_ASIAN_AUTO_SPACE_GLYPH,
    );
    if !params.is_empty() {
        params.push_str(", ");
    }
    let _ = write!(params, "spacing: {width}");
    format!("#text({params})[\\u{{00A0}}]")
}

pub(super) fn generate_paragraph(
    out: &mut String,
    para: &Paragraph,
    line_grid_pitch: Option<f64>,
    default_tab_width_pt: f64,
    breaks_hangul_at_eojeol: bool,
    available_measure_pt: Option<f64>,
) -> Result<(), ConvertError> {
    let style = &para.style;
    let paragraph_tab_width_pt: f64 = paragraph_default_tab_width_pt(style, default_tab_width_pt);

    if let Some(level) = style.heading_level {
        // A heading is still a paragraph: Word paints its `w:pBdr` and `w:shd`
        // around it exactly as it does around body copy, and a chapter-rule
        // heading style is the commonest place a `w:pBdr` appears at all.
        // Returning here before any decoration was emitted dropped every one
        // of them — 22 chapter rules in the technical-brief fixture, while the
        // header rule on the same page, declared directly rather than through
        // a style, survived (issue #581).
        //
        // Word spaces and measures it as one too. While the wrapper opened
        // only for decoration, both the block spacing and the line box were
        // Typst's own `#set heading` defaults — numbers no `w:spacing`, no
        // style definition and no Word rule produced — and since Typst
        // collapses adjacent block spacing to the larger of the two, that
        // default swallowed the neighbouring paragraph's declared gap as well
        // (issue #1132).
        //
        // A heading resolves `w:spacing` exactly as body copy does, so
        // `style` already holds the answer and no heading-specific fallback
        // exists to add. Measured on native Word exports of a package whose
        // `Heading1` paragraphs state no `w:spacing`: with
        // `w:docDefaults/w:pPrDefault` declared the export is layout-identical
        // to one stating `w:before="0" w:after="0"`, and without it identical
        // to `w:before="0" w:after="160"` — Word's built-in `Normal`, the same
        // fallback #1085 measured for body paragraphs. Word's built-in
        // `Heading N` spacing takes no part.
        let decorated = style.background.is_some() || style.border.is_some();
        let line_height_settings: Option<String> =
            word_line_height_settings(&para.runs, style, line_grid_pitch);
        // The gaps measure from the line box's edges, so the box has to come
        // with them: on Typst's glyph-tight default the block ends at the
        // baseline, and the heading's own descender goes missing from the gap
        // below it.
        let wrapped: bool = decorated
            || style.space_before.is_some()
            || style.space_after.is_some()
            || style.line_box.is_some()
            || line_height_settings.is_some();
        if wrapped {
            out.push_str("#block(width: 100%");
            write_block_spacing_params(out, style);
            write_block_decoration_params(out, style);
            out.push_str(")[\n");
            write_paragraph_double_border_overlays(
                out,
                &style.border,
                style.border_space.as_deref().copied().unwrap_or_default(),
            );
            write_line_box_settings(out, style.line_box);
            if let Some(ref settings) = line_height_settings {
                out.push_str(settings);
            }
        }
        // A contents entry is laid out as body text, not as a copy of the
        // heading, so it cannot be built from the heading's rendered content —
        // the size and weight are inline markup inside it and no enclosing set
        // rule beats them. Drop the heading's plain text under a label instead
        // and let the list style it (issue #610), the same shape the caption
        // lists already use.
        let plain: String = paragraph_plain_text(&para.runs);
        let _ = writeln!(
            out,
            "#metadata((level: {level}, text: \"{}\", font: {}))<{}>",
            escape_typst_string(&plain),
            crate::render::font_subst::font_with_fallbacks_for_text(
                first_run_family(&para.runs).unwrap_or("Calibri"),
                &plain,
            ),
            TOC_ENTRY_LABEL
        );
        // Whichever fixed line box the wrapper just put in force is what a
        // framed eojeol has to restore inside itself (issue #626); an
        // unwrapped heading still emits no fixed edges and needs no
        // correction.
        let line_box_em: Option<(f64, f64)> = wrapped
            .then(|| {
                word_line_box_em(&para.runs, style, line_grid_pitch).or_else(|| {
                    style
                        .line_box
                        .map(|line_box| (line_box.ascent_em, line_box.descent_em))
                })
            })
            .flatten();
        let _ = write!(out, "#heading(level: {level})[");
        generate_runs_with_tabs(
            out,
            &para.runs,
            style.tab_stops.as_deref(),
            paragraph_tab_width_pt,
            paragraph_eojeol_wrap(
                breaks_hangul_at_eojeol,
                style,
                line_box_em,
                available_measure_pt,
            ),
        );
        out.push_str("]\n");
        if wrapped {
            out.push_str("]\n");
        }
        return Ok(());
    }

    let line_height_settings: Option<String> =
        word_line_height_settings(&para.runs, style, line_grid_pitch);
    let has_para_style = needs_block_wrapper(style) || line_height_settings.is_some();

    // Word's `w:ind` offsets the paragraph's whole column, and paints
    // `w:shd` and `w:pBdr` from the indent rather than the margin, so the
    // indent goes on an outer block as an inset and the fill and border stay
    // on an inner block that spans only the inset content area (issue #464).
    let indent = paragraph_indent_pt(style);
    if indent.is_some() {
        out.push_str("#block(width: 100%");
        write_block_spacing_params(out, style);
        write_paragraph_indent_inset(out, indent);
        out.push_str(")[\n");
    }

    if has_para_style {
        // The wrapper must span the full line width: Typst blocks shrink to
        // their content by default, which would defeat the inner #align.
        // Word measures `w:spacing w:before/w:after` from the edges of the
        // full line box, which `word_line_height_settings` spans directly,
        // so those gaps reach the block unmodified (issues #394, #452).
        out.push_str("#block(width: 100%");
        if indent.is_none() {
            write_block_spacing_params(out, style);
        }
        write_block_decoration_params(out, style);
        out.push_str(")[\n");
        write_paragraph_double_border_overlays(
            out,
            &style.border,
            style.border_space.as_deref().copied().unwrap_or_default(),
        );
        write_line_box_settings(out, style.line_box);
        write_par_settings(out, style, &para.runs);
        if let Some(ref settings) = line_height_settings {
            out.push_str(settings);
        }
    }

    if para.runs.is_empty() {
        out.push_str("#v(12pt)");
        if has_para_style {
            out.push_str("\n]");
        }
        if indent.is_some() {
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

    // Whichever fixed line box the wrapper above put in force — the computed
    // Word line, or the paragraph's own `LineBox` — is what a framed eojeol
    // has to restore inside itself. The two are mutually exclusive:
    // `word_line_leading_pt` bails on a paragraph that declares a `LineBox`.
    let line_box_em: Option<(f64, f64)> = has_para_style
        .then(|| {
            word_line_box_em(&para.runs, style, line_grid_pitch).or_else(|| {
                style
                    .line_box
                    .map(|line_box| (line_box.ascent_em, line_box.descent_em))
            })
        })
        .flatten();

    generate_word_runs_with_tabs(
        out,
        &para.runs,
        style.tab_stops.as_deref(),
        paragraph_tab_width_pt,
        paragraph_eojeol_wrap(
            breaks_hangul_at_eojeol,
            style,
            line_box_em,
            available_measure_pt,
        ),
        style,
        line_grid_pitch,
    );

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }
    if indent.is_some() {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

/// The letter-space PowerPoint counts after a slide line's last glyph, when
/// the paragraph's alignment is one that the extra width moves the line by.
///
/// PowerPoint measures a line with one letter-space after *every* glyph, the
/// last one included, and places the line from that width: centring halves the
/// trailing space, right alignment consumes it whole, and left alignment —
/// which starts at the content edge whatever the width — cannot see it at all.
/// Typst drops the tracking of a shaped item's final glyph
/// ([`track_and_space`] adds it only where another glyph follows), so a
/// centred tracked line came out half a letter-space to the right of the
/// native export (issue #1075).
///
/// The space belongs to the last run that puts a glyph on the line, because
/// that run's `a:rPr/@spc` is the one PowerPoint applies after it.
///
/// [`track_and_space`]: https://github.com/typst/typst/blob/v0.14.2/crates/typst-layout/src/inline/shaping.rs
pub(super) fn powerpoint_trailing_letter_space_pt(
    style: &ParagraphStyle,
    runs: &[Run],
) -> Option<f64> {
    if !matches!(
        style.alignment,
        Some(Alignment::Center) | Some(Alignment::Right)
    ) {
        return None;
    }
    runs.iter()
        .rev()
        .find(|run| !run.text.is_empty())?
        .style
        .letter_spacing
        // Zero is not tracking: PowerPoint writes `spc="0"` routinely, and a
        // `0pt` spacer would be pure noise in every deck that does.
        .filter(|spacing| *spacing != 0.0)
}

pub(super) fn paragraph_default_tab_width_pt(style: &ParagraphStyle, fallback_pt: f64) -> f64 {
    style
        .default_tab_stop_pt
        .filter(|width_pt| *width_pt > 0.0)
        .unwrap_or(fallback_pt)
}

/// The paragraph's `(left, right)` indent in points, or `None` when it has
/// neither. Negative indents — Word lets a paragraph hang into the margin —
/// are clamped to zero, because a Typst inset cannot be negative.
pub(super) fn paragraph_indent_pt(style: &ParagraphStyle) -> Option<(f64, f64)> {
    let left: f64 = style.indent_left.unwrap_or(0.0).max(0.0);
    let right: f64 = style.indent_right.unwrap_or(0.0).max(0.0);
    (left > 0.0 || right > 0.0).then_some((left, right))
}

fn write_paragraph_indent_inset(out: &mut String, indent: Option<(f64, f64)>) {
    if let Some((left, right)) = indent {
        let _ = write!(
            out,
            ", inset: (left: {}pt, right: {}pt)",
            format_f64(left),
            format_f64(right)
        );
    }
}

/// Whether the paragraph needs its own block to carry style. The indent is
/// deliberately absent: `generate_paragraph` wraps flow indents directly,
/// while fixed-text paragraph and list paths compute their PowerPoint-specific
/// first-line origin separately. Counting it here would add a duplicate bare
/// wrapper that only leaks Typst's default block spacing.
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

/// Line-box settings for a body paragraph: a fixed box spanning Word's full
/// line advance — the font's hhea line, 1.3 times it when the line carries
/// East Asian text, or a snapping document grid's pitch — with zero leading.
/// Typst's glyph-tight default renders such documents 20-30% shorter and
/// shifts every page break (issue #354).
///
/// The baseline sits at a constant `hhea ascender + lineGap` below the box
/// top, never at the font's ascender/descender proportion of it: whatever
/// height the line gains over the font's own — the East Asian bonus's lower
/// half, or a grid slot's slack — accrues below the baseline, not around it
/// (issues #508, #518).
///
/// Carrying the advance inside the box, rather than recovering the
/// remainder as `par(leading:)`, is what makes a paragraph's height match
/// Word's. Typst inserts `leading` only *between* the lines of one
/// paragraph, so an n-line paragraph came out one whole leading short
/// however many lines it had, and consecutive 9pt Courier New paragraphs
/// advanced 28% tighter than Word (issue #452). It also lets `w:spacing
/// w:before/w:after` reach the block unchanged, because the block edges now
/// sit exactly where Word measures those gaps from (issue #394).
pub(super) fn word_line_height_settings(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<String> {
    let (top_em, bottom_em) = word_line_box_em(runs, style, line_grid_pitch)?;
    // Pin the line box to the nominal font's own metric edges as fixed em
    // values rather than the "ascender"/"descender" keywords. The keywords
    // let Typst resolve the box against the tallest font on each line, so a
    // bullet marker or em dash pulled from a taller fallback font inflated
    // that one line's advance past the grid/single-spacing (issue #398).
    Some(format!(
        "#set text(top-edge: {}em, bottom-edge: -{}em)\n#set par(leading: 0pt)\n",
        format_f64(top_em),
        format_f64(bottom_em)
    ))
}

/// The `(top-edge, bottom-edge)` in em behind [`word_line_height_settings`],
/// exposed so a framed eojeol can restore the same edges inside itself
/// (issue #626).
pub(super) fn word_line_box_em(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<(f64, f64)> {
    let (ascender_em, descender_em, leading_em) =
        word_line_box_and_leading(runs, style, line_grid_pitch)?;
    let metric_em: f64 = ascender_em + descender_em;
    if metric_em <= 0.0 {
        return None;
    }
    let pitch_em: f64 = metric_em + leading_em;
    let top_em: f64 = ascender_em + east_asian_ascent_excess_em(runs, metric_em);
    Some((top_em, pitch_em - top_em))
}

/// The Word line-box context applied to each declared font run in a mixed-face
/// body paragraph.
///
/// The paragraph-level fixed edges remain the nominal face's edges, which is
/// what prevents glyph fallback from inflating an otherwise uniform line
/// (issue #398). When the paragraph deliberately names more than one metric
/// family, however, Word takes the greatest ascent and descent among the runs
/// that actually land on each line. Giving those runs their own numeric edges
/// lets Typst perform that per-line maximum without consulting glyph fallback
/// or imposing a paragraph-wide maximum on every line (issue #638).
#[derive(Clone, Copy)]
struct WordRunLineMetrics<'a> {
    paragraph_style: &'a ParagraphStyle,
    line_grid_pitch: Option<f64>,
}

impl<'a> WordRunLineMetrics<'a> {
    fn for_mixed_declared_families(
        runs: &[Run],
        paragraph_style: &'a ParagraphStyle,
        line_grid_pitch: Option<f64>,
    ) -> Option<Self> {
        let mut families = runs
            .iter()
            .filter(|run| run.footnote.is_none() && !run.text.is_empty())
            .filter_map(|run| east_asian_aware_metric_family(std::slice::from_ref(run)));
        let first_family: &str = families.next()?;
        families
            .any(|family| !family.eq_ignore_ascii_case(first_family))
            .then_some(Self {
                paragraph_style,
                line_grid_pitch,
            })
    }

    fn line_box_em(self, run: &Run) -> Option<(f64, f64)> {
        word_line_box_em(
            std::slice::from_ref(run),
            self.paragraph_style,
            self.line_grid_pitch,
        )
    }
}

/// A line-box override attached to one run rather than the whole paragraph.
/// Typst then takes the greatest top and bottom among only the runs that land
/// on each physical line.
#[derive(Clone, Copy)]
enum RunLineBox {
    Em { top: f64, bottom: f64 },
    Points { top: f64, bottom: f64 },
}

#[derive(Clone, Copy)]
enum RunLineMetrics<'a> {
    Word(WordRunLineMetrics<'a>),
    PowerPoint(PowerPointRunLineMetrics),
}

impl RunLineMetrics<'_> {
    fn line_box(self, run: &Run) -> Option<RunLineBox> {
        match self {
            Self::Word(metrics) => metrics
                .line_box_em(run)
                .map(|(top, bottom)| RunLineBox::Em { top, bottom }),
            Self::PowerPoint(metrics) => metrics
                .line_box_pt(run)
                .map(|(top, bottom)| RunLineBox::Points { top, bottom }),
        }
    }
}

/// PowerPoint line metrics for an automatically wrapped paragraph whose runs
/// declare different sizes.
///
/// The paragraph-wide family set still decides the baseline share, preserving
/// the existing font model. Only the size varies per run, allowing Typst to
/// derive each physical line from the largest run that actually landed there
/// instead of imposing the paragraph maximum on every line (#1329).
#[derive(Clone, Copy)]
struct PowerPointRunLineMetrics {
    plain_ascent_em: f64,
    line_spacing: f64,
}

impl PowerPointRunLineMetrics {
    fn for_mixed_declared_sizes(runs: &[Run], style: &ParagraphStyle) -> Option<Self> {
        let visible_runs: Vec<&Run> = runs
            .iter()
            .filter(|run| run.footnote.is_none() && !run.text.is_empty())
            .collect();
        if visible_runs.is_empty() || visible_runs.iter().any(|run| run.style.font_size.is_none()) {
            return None;
        }
        let first_size: f64 = visible_runs[0].style.font_size?;
        if !visible_runs.iter().any(|run| {
            run.style
                .font_size
                .is_some_and(|size| (size - first_size).abs() > f64::EPSILON)
        }) {
            return None;
        }

        let (plain_ascent_em, line_spacing) = powerpoint_paragraph_line_model(runs, style)?;
        Some(Self {
            plain_ascent_em,
            line_spacing,
        })
    }

    fn line_box_pt(self, run: &Run) -> Option<(f64, f64)> {
        let font_size_pt: f64 = run.style.font_size?;
        let (top_em, bottom_em) = powerpoint_percentage_line_box_em(
            self.plain_ascent_em,
            font_size_pt,
            self.line_spacing,
        );
        Some((top_em * font_size_pt, bottom_em * font_size_pt))
    }
}

/// Line-box settings for a slide's text: PowerPoint's flat 1.2em line, split
/// by [`crate::render::pdf::powerpoint_line_box_split_em`], and zero leading
/// between lines.
///
/// This is the PPTX counterpart of [`word_line_height_settings`], and the two
/// models genuinely differ. Word's line is the font's own hhea pitch;
/// PowerPoint's ignores the font's metrics for the height and consults them
/// only for where inside it the baseline sits. Slide text used to take the Word
/// treatment, which is up to 4% short per line and accumulates down a bullet
/// list, and it seated the baseline by Typst's normalised ascender, which put
/// a bottom-anchored box's last baseline flat on the inset with no descent gap
/// at all (issue #513).
///
/// `<a:lnSpc><a:spcPct>` scales that line rather than replacing it: the advance
/// is `percent x 1.2em`, and [`powerpoint_percentage_line_box_em`] takes the
/// whole change off the ascent side. Carrying the percentage as `par(leading)`
/// instead moved nothing between single-line paragraphs — a slide's code block
/// is one `<a:p>` per line — so the lines overlapped (issue #541).
///
/// The seat inside that line is measured in whole points rather than in em.
/// Paragraph-wide callers therefore resolve it against the largest declared
/// size. An automatically wrapped mixed-size paragraph instead attaches the
/// resulting point edges to each run, so Typst chooses the largest run that
/// actually lands on each physical line.
///
/// `None` when the paragraph carries its own line box or when the font's
/// metrics are unknown. An absolute `a:spcPts` rule is expressed as the
/// equivalent scale of PowerPoint's plain 1.2em line so it retains the same
/// baseline-seating model while replacing the advance.
fn powerpoint_paragraph_line_box_em(runs: &[Run], style: &ParagraphStyle) -> Option<(f64, f64)> {
    let (ascent_em, percent) = match style.line_spacing {
        Some(LineSpacing::Exact(points)) if points > 0.0 && style.line_box.is_none() => {
            let families: Vec<&str> = powerpoint_line_families(runs, style);
            let (ascent_em, _descent_em) =
                crate::render::pdf::powerpoint_line_box_em_for_families(&families)?;
            let font_size_pt: f64 = paragraph_font_size_pt(runs);
            let plain_advance_pt: f64 =
                crate::render::pdf::POWERPOINT_LINE_HEIGHT_FACTOR * font_size_pt;
            (ascent_em, points / plain_advance_pt)
        }
        _ => powerpoint_paragraph_line_model(runs, style)?,
    };
    Some(powerpoint_percentage_line_box_em(
        ascent_em,
        paragraph_font_size_pt(runs),
        percent,
    ))
}

fn powerpoint_paragraph_line_model(runs: &[Run], style: &ParagraphStyle) -> Option<(f64, f64)> {
    if style.line_box.is_some() {
        return None;
    }
    let percent: f64 = match style.line_spacing {
        None => 1.0,
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => factor,
        Some(_) => return None,
    };
    let families: Vec<&str> = powerpoint_line_families(runs, style);
    let (ascent_em, _descent_em) =
        crate::render::pdf::powerpoint_line_box_em_for_families(&families)?;
    Some((ascent_em, percent))
}

/// Every font on a slide paragraph's line: the families its runs declare, plus
/// the one its paragraph mark ends up in.
///
/// PowerPoint shares one 1.2em box across all of them — see
/// [`crate::render::pdf::powerpoint_line_box_split_em`], which measures the
/// mark's part of it. A run declaring no family of its own rides whatever the
/// line already has rather than dragging the renderer's default face into the
/// box; when nothing names a family, that default is the only face there is.
fn powerpoint_line_families<'a>(runs: &'a [Run], style: &'a ParagraphStyle) -> Vec<&'a str> {
    let mut families: Vec<&str> = Vec::new();
    let mut push = |family: &'a str| {
        if !family.is_empty()
            && !families
                .iter()
                .any(|seen| seen.eq_ignore_ascii_case(family))
        {
            families.push(family);
        }
    };
    for run in runs {
        if let Some(family) = run.style.font_family.as_deref() {
            push(family);
        }
    }
    if let Some(mark) = style.paragraph_mark_font_family.as_deref() {
        push(mark);
    }
    if families.is_empty() {
        families.push(crate::defaults::TYPST_DEFAULT_FONT_FAMILY);
    }
    families
}

/// The `(above baseline, below baseline)` split, in em, of a line an
/// `<a:lnSpc><a:spcPct>` has resized to `percent` of PowerPoint's 1.2em box,
/// for text set at `font_size_pt`. `ascent_em` is the plain line's own
/// above-baseline share, as [`crate::render::pdf::powerpoint_line_box_em`]
/// splits it.
///
/// **PowerPoint seats a baseline a whole number of points below its line box's
/// top**, and the descent gap is whatever is left of the line. So the seat's
/// share of the em is not constant across sizes, and neither is the gap:
/// measured on native PowerPoint 16.112 exports, Avenir Next LT Pro keeps
/// 0.192em under its baseline at 10pt and 0.2625em at 32pt. A one-factor probe
/// deck of bottom-anchored boxes with every inset zeroed — Georgia (1.13623em,
/// fits the box), Verdana (1.21533em), Avenir Next LT Pro (1.21289em) and
/// Posterama (1.33008em) at 8, 11, 14, 18, 24, 28, 32, 36, 40, 44, 48, 54, 72
/// and 100pt — puts all 56 cells on `1.2 x size - round(share x size)` within
/// the exports' 0.12pt half-grid, and none of them within 0.12pt of the
/// unrounded share. Carrying the share as a plain em fraction left every slide
/// of the #841 Contoso deck that repeats its 10pt footer band 0.55pt high
/// (issue #1074).
///
/// Georgia is the one of the four that fits the box, and its cells land on the
/// same proportional share the other three do. Halving a fitting face's extra
/// leading instead — which is what
/// [`crate::render::pdf::powerpoint_line_box_split_em`] used to hand it — misses
/// 9 of Georgia's 14 cells, by up to 2.04pt at 72pt (issue #1118).
///
/// PowerPoint resizes the line from its **top**: the gap the face keeps below
/// its baseline is the plain line's `1.2em - ascent_em` whatever the
/// percentage, and the ascent side absorbs the whole change, so the percentage
/// enters only through the advance the seat is measured back from. Measured on
/// the same exports, against the plain-box control for the same face and size:
/// Arial 38pt drops its first baseline 36.96pt below the content top plain and
/// 30.00pt under `val="85000"`, a 6.96pt loss where the line loses 6.84pt. All 18
/// Posterama titles of the #841 Contoso deck agree across 30, 32, 36, 38, 46 and
/// 50pt. Scaling both sides by the percentage instead, which is what the even
/// division of a *taller* box implies, left every one of those titles 1.8-3.7pt
/// low (issues #1020, #1024).
///
/// The rounding lands on the **scaled** seat, measured back from the plain
/// line's unrounded gap, so there is one rounding and not two. Rounding the
/// plain seat to a point first and subtracting that from the scaled advance
/// predicts 10pt for slide 8's 14pt `Heraclitus` attribution under
/// `val="85000"`, where the export seats it 11.04pt below the content top; this
/// form predicts 11. The #841 deck's other `spcPct` frames cannot tell the two
/// apart — its 30pt centred and 38pt top-anchored Posterama titles land on 23pt
/// and 29pt either way — so that one attribution carries the distinction.
///
/// The seat cannot outgrow the line it sits in, so a percentage small enough to
/// close the box takes the ascent to zero rather than negative.
pub(super) fn powerpoint_percentage_line_box_em(
    ascent_em: f64,
    font_size_pt: f64,
    percent: f64,
) -> (f64, f64) {
    let advance_em: f64 = (crate::render::pdf::POWERPOINT_LINE_HEIGHT_FACTOR * percent).max(0.0);
    let below_em: f64 = crate::render::pdf::POWERPOINT_LINE_HEIGHT_FACTOR - ascent_em;
    if font_size_pt.is_nan() || font_size_pt <= 0.0 {
        // Nothing to round against; fall back to the bare em split.
        let below_em: f64 = below_em.clamp(0.0, advance_em);
        return (advance_em - below_em, below_em);
    }
    let advance_pt: f64 = advance_em * font_size_pt;
    let seat_pt: f64 = ((advance_em - below_em) * font_size_pt)
        .round()
        .clamp(0.0, advance_pt);
    (
        seat_pt / font_size_pt,
        (advance_pt - seat_pt) / font_size_pt,
    )
}

/// The same box, stated in the unit that survives the scope it lands in.
///
/// An `em` in a `#set text` edge resolves against whatever size is in force
/// where the rule applies, not against the size the box was derived from. The
/// paragraph emits a `#set text(size:)` of its own only when every one of its
/// runs declares the same size, and a `<a:br/>` reaches the IR as a run with no
/// run properties at all — so one hard break is enough to strip that rule and
/// leave the edges resolving against Typst's 11pt default. Every hard-broken
/// line under 11pt then advanced a flat `1.2 x 11pt` = 13.20pt, 89% too far for
/// a 6pt caption (issue #1115).
///
/// Restating the box in points pins it to the size it was computed from — the
/// paragraph's largest declared size, which is the one PowerPoint's line keys
/// on. When nothing declares a size there is none to pin to: the box was
/// derived from the same default an `em` resolves against, so keeping `em`
/// leaves an inherited size (a slide table's, say) still governing.
pub(super) fn powerpoint_line_height_settings(
    runs: &[Run],
    style: &ParagraphStyle,
) -> Option<String> {
    let (ascent_em, descent_em) = powerpoint_paragraph_line_box_em(runs, style)?;
    let Some(font_size_pt) = declared_paragraph_font_size_pt(runs) else {
        return Some(format!(
            "#set text(top-edge: {}em, bottom-edge: -{}em)\n#set par(leading: 0pt)\n",
            format_f64(ascent_em),
            format_f64(descent_em)
        ));
    };
    Some(format!(
        "#set text(top-edge: {}pt, bottom-edge: -{}pt)\n#set par(leading: 0pt)\n",
        format_f64(ascent_em * font_size_pt),
        format_f64(descent_em * font_size_pt)
    ))
}

/// The height, in points, of one blank PowerPoint line for `runs`.
///
/// The empty-paragraph strut in a slide's table cell has to be sized from the
/// same model as its neighbours, or the blank line keeps Word's hhea height
/// while the text beside it takes PowerPoint's 1.2em one (issues #625, #663).
pub(super) fn powerpoint_line_box_pt(runs: &[Run]) -> Option<f64> {
    let family: &str = runs
        .iter()
        .find_map(|run| run.style.font_family.as_deref())
        .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_FAMILY);
    let (ascent_em, descent_em) = crate::render::pdf::powerpoint_line_box_em(family)?;
    Some((ascent_em + descent_em) * paragraph_font_size_pt(runs))
}

/// The nominal font's `(above baseline, below baseline)` split plus the
/// leading, in em, that tops the line box up to Word's line advance. `None`
/// when the metric-edge treatment does not apply.
///
/// Since #508 the pair already sums to the font's single-spacing pitch, so the
/// leading is zero for a Latin line with no grid; it carries the East Asian
/// bonus and any grid slack (issue #518).
fn word_line_box_and_leading(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<(f64, f64, f64)> {
    let leading_pt: f64 = word_line_leading_pt(runs, style, line_grid_pitch)?;
    let family: &str = east_asian_aware_metric_family(runs)?;
    let (ascender_em, descender_em, _word_pitch_em) =
        crate::render::pdf::font_line_metrics_em(family)?;
    let font_size: f64 = paragraph_font_size_pt(runs);
    Some((ascender_em, descender_em, leading_pt / font_size))
}

/// Word gives a line set in an East Asian face 130% of the font's own hhea
/// line, and centres the bonus on the baseline: half above, half below.
///
/// The face decides, not the line's characters — see
/// [`line_takes_east_asian_metrics`] (issue #643).
///
/// Both halves are measured, not assumed. Against native Word exports an Arial
/// first baseline sits at `hhea ascender + lineGap` = 0.937988em below the text
/// top while a Malgun Gothic one at the same settings sits at 1.28786em, and
/// the difference is exactly `0.15 x` Malgun's 1.330078em hhea pitch — the term
/// #508 could not attribute to any font table. The matching lower half shows up
/// as the advance: every Korean fixture in the business corpus paces its
/// wrapped lines at `1.3 x` the hhea pitch (10.5pt Malgun measures 18.00-18.24
/// against 18.156 predicted), and 06_official_letter_ko's 9.5pt paragraphs
/// advance 16.43pt where the font's bare hhea line is 12.64pt (issue #518).
const EAST_ASIAN_LINE_HEIGHT_FACTOR: f64 = 1.3;

/// The half of that bonus which lands above the baseline.
const EAST_ASIAN_ASCENT_EXCESS: f64 = (EAST_ASIAN_LINE_HEIGHT_FACTOR - 1.0) / 2.0;

/// The family whose metrics pace these runs' lines.
///
/// A line carrying East Asian text is paced by the East Asian face, not by the
/// Latin one the same runs also name: the 1.3 factor above was measured
/// against Malgun Gothic's hhea pitch. Reading the Latin family was harmless
/// only while `w:eastAsia` was being dropped and the Latin family was the one
/// actually shaping the Hangul (issue #575).
fn east_asian_aware_metric_family(runs: &[Run]) -> Option<&str> {
    let latin = || runs.iter().find_map(|run| run.style.font_family.as_deref());
    if has_east_asian_text(runs) {
        runs.iter()
            .find_map(|run| run.style.east_asian_font_family.as_deref())
            .or_else(latin)
    } else {
        latin()
    }
}

fn has_east_asian_text(runs: &[Run]) -> bool {
    runs.iter().any(|run| run.text.chars().any(is_cjk_like))
}

/// Whether Word measures this line with its East Asian metrics.
///
/// Word keys that on the face the line is **set in**, not on the script of its
/// characters: a CJK family shapes its own Latin glyphs and the line keeps the
/// East Asian height. Three Heading2 paragraphs in `03_meeting_minutes_ko`
/// share `w:rPr` naming Malgun Gothic in every `w:rFonts` slot and differ only
/// in script; Word gives them the same height, and gating on the text left the
/// Latin-only one 2.37pt short (issue #643).
///
/// Reading the resolved metric family rather than `w:eastAsia` directly is what
/// keeps a Latin line Latin: the business fixtures declare
/// `w:eastAsia="Arial"`, and [`east_asian_aware_metric_family`] already resolves
/// a Latin-only line to its Latin family. So an Arial paragraph inside a Korean
/// document keeps its own line, which is what Word does — snapping those
/// inflated every Western document by 30-50% (issue #354).
pub(super) fn line_takes_east_asian_metrics(runs: &[Run]) -> bool {
    has_east_asian_text(runs)
        || east_asian_aware_metric_family(runs)
            .is_some_and(crate::render::font_subst::is_east_asian_family)
}

/// The extra ascent, in em, that Word gives a line set in an East Asian face.
///
/// `pitch_em` is the font's own hhea pitch, never the line's advance: under a
/// document grid the slot's extra height accrues entirely below the baseline,
/// so this term must not scale with the slot (issue #518).
fn east_asian_ascent_excess_em(runs: &[Run], pitch_em: f64) -> f64 {
    if line_takes_east_asian_metrics(runs) {
        EAST_ASIAN_ASCENT_EXCESS * pitch_em
    } else {
        0.0
    }
}

/// A line's sub-baseline share, in em: the East Asian advance minus the ascent
/// Word seats the glyphs at, rather than the face's plain descender.
///
/// Two callers want exactly this quantity. A footer story is measured **up**
/// from the `w:pgMar/@w:footer` line to its last line's box bottom, which is
/// the mirror of [`word_header_line_ascent_em`]. A header's `w:pBdr w:space`
/// gap is measured **down** from the same box bottom. The rationale below was
/// established for the footer case; the header case is noted at the end.
///
/// Which one it is decides the footer outright. The three golden mocks 01, 02
/// and 03 differ in their `footer1.xml` only in `w:rFonts`, and Word moves the
/// baseline with the font: 804.72pt for the Arial footer and 802.80pt for the
/// Malgun Gothic ones, both at `w:footer="708"` = 35.40pt on A4. Typst's
/// `bottom-edge: "descender"` answers with its *normalised* descender, which is
/// 0.199em for Malgun Gothic against the 0.4412em its line box actually carries
/// below the baseline — so every Korean footer in the corpus printed 2.10pt low
/// while the Arial one, whose two answers nearly agree, looked correct
/// (issue #630).
///
/// East Asian metrics key on the resolved face, not on the footer's characters:
/// `- 2 -` is ASCII in all three files and Word still gives the Malgun ones the
/// taller line. That is [`line_takes_east_asian_metrics`]'s rule, already
/// established for body lines by issue #643.
///
/// The header path asks the same question: `w:pBdr w:space` is measured from
/// the line's bottom, so a Korean header rule sat 1.98pt high on the same
/// normalised-descender answer until it read this instead (issue #737).
///
/// `None` when the face's metrics are unknown, which leaves the story on the
/// renderer's own seat.
pub(super) fn word_line_box_descent_em(runs: &[Run]) -> Option<f64> {
    let family: &str = east_asian_aware_metric_family(runs)?;
    let (ascender_em, descender_em, pitch_em) = crate::render::pdf::font_line_metrics_em(family)?;
    if ascender_em + descender_em <= 0.0 || pitch_em <= 0.0 {
        return None;
    }
    let advance_em: f64 = if line_takes_east_asian_metrics(runs) {
        EAST_ASIAN_LINE_HEIGHT_FACTOR * pitch_em
    } else {
        pitch_em
    };
    let top_em: f64 = ascender_em + east_asian_ascent_excess_em(runs, pitch_em);
    Some(advance_em - top_em)
}

/// What sits below the baseline of an Excel sheet header or footer line, in em
/// of the line's face.
///
/// The face's bare `hhea` descender, deliberately *not*
/// [`word_line_box_descent_em`]: that one carries Word's 1.3x East Asian line
/// box and the line gap Word keeps above the baseline, neither of which any
/// Excel export measured here shows. Native Excel-for-Mac exports put a footer
/// baseline exactly its `hhea` descent above the band Excel seats the text on —
/// Calibri 0.26855em, Arial 0.21191em, Verdana 0.20996em, Times New Roman
/// 0.21631em and Aptos 0.28174em all land on the whole point Excel prints
/// (issue #1142).
///
/// `None` when the face's metrics are unknown, which leaves the story on the
/// renderer's own seat.
pub(super) fn sheet_line_box_descent_em(runs: &[Run]) -> Option<f64> {
    let family: &str = east_asian_aware_metric_family(runs)?;
    let (_, descent_em, _) = crate::render::pdf::font_line_metrics_em(family)?;
    (descent_em > 0.0).then_some(descent_em)
}

/// Where Word seats a header story's first baseline, in em below the
/// `w:pgMar/@w:header` line the header is measured from.
///
/// That origin is the top of the first line's *ascent*, so the term is the
/// face's bare hhea ascender — not [`word_line_box_em`]'s top edge, which folds
/// in the hhea line gap Word keeps above the origin — plus the upper half of
/// the East Asian bonus when the line carries East Asian text.
///
/// Measured against native Word exports of the business corpus, all at
/// `w:header="708"` = 35.40pt: an 8pt Arial header baseline lands at 42.72pt
/// against 42.64pt predicted (`0.9053em`), and an 8pt Malgun Gothic one at
/// 45.60pt against 45.70pt predicted (`1.2879em`), both within the 0.24pt grid
/// those exports quantise positions to. Taking the gap-inclusive body ascent
/// instead would predict 42.90pt for the Arial case, a whole grid step past
/// what Word wrote (issue #629).
///
/// The bonus keys on the resolved face, not on the header's characters, like
/// every other line box (issues #643, #814): a one-factor native export of
/// `10_research_report_ko` with its header text swapped to `Monthly Customer
/// Satisfaction Trend Report` keeps the baseline at 45.60pt — exactly the
/// Korean control's seat — where the bare ascender would put it at 44.11pt.
fn word_header_line_ascent_em(runs: &[Run], family: &str) -> Option<f64> {
    let ascender_em: f64 = crate::render::pdf::font_hhea_ascender_em(family)?;
    if !line_takes_east_asian_metrics(runs) {
        return Some(ascender_em);
    }
    let (_, _, pitch_em) = crate::render::pdf::font_line_metrics_em(family)?;
    Some(ascender_em + EAST_ASIAN_ASCENT_EXCESS * pitch_em)
}

/// How far a pinned header band must move so its first baseline lands where
/// Word puts it, in points, positive downward.
///
/// The difference between [`word_header_line_ascent_em`] and the ascent the
/// compiler gives the same line — its default `top-edge` is `"cap-height"` —
/// resolved at the size that line is set in. Emitting the difference as a band
/// offset rather than as a `top-edge` is what keeps the story's own line advance
/// untouched: `top-edge` applies to every line of the paragraph it is set on, so
/// declaring Word's ascent there stretched a wrapped 8pt Arial header's advance
/// from 10.93pt to 12.44pt against Word's 9.20pt (issue #629).
///
/// Both terms read the same family and the same size — the largest among the
/// runs, which is the one whose ascent wins the line in either engine.
pub(super) fn word_header_band_shift_pt(runs: &[Run]) -> Option<f64> {
    let family: &str = east_asian_aware_metric_family(runs)?;
    let word_ascent_em: f64 = word_header_line_ascent_em(runs, family)?;
    let compiler_ascent_em: f64 = crate::render::pdf::font_cap_height_em(family)?;
    Some((word_ascent_em - compiler_ascent_em) * paragraph_font_size_pt(runs))
}

/// Leading that makes a header or footer line advance by Word's pitch.
///
/// Typst advances a line by `top-edge + bottom-edge + leading`, and the story
/// carried none of those, so it took the 0.65em default: a wrapped 8pt Arial
/// header advanced 10.9305pt = (0.65 + 0.71631) x 8 where Word advances
/// 9.1992pt = (2355/2048) x 8 (issue #735).
///
/// The leading is the remainder after the edges, not a replacement for them.
/// That matters: the top edge is what #629's `word_header_band_shift_pt` seats
/// the first baseline against, so leaving it alone keeps that seat exactly
/// where it was and confines this to the advance *between* lines. `bottom_edge_em`
/// is whatever the caller set, which is nonzero only where a `w:pBdr` wrapper
/// pinned it (#737).
///
/// Not to be confused with [`word_line_leading_pt`], which answers a different
/// question for body paragraphs: that one is folded into a fixed line box,
/// this one is emitted directly as `par(leading:)` for a header or footer
/// story.
///
/// `None` where the face's metrics are unknown, or where the edges already
/// exceed Word's pitch and no leading could shrink the advance to it.
pub(super) fn word_hf_line_leading_pt(runs: &[Run], bottom_edge_em: f64) -> Option<f64> {
    let family: &str = east_asian_aware_metric_family(runs)?;
    let (_ascender_em, _descender_em, pitch_em) = crate::render::pdf::font_line_metrics_em(family)?;
    if pitch_em <= 0.0 {
        return None;
    }
    // The story sets no `top-edge`, so Typst's default cap-height one is what
    // the advance already carries.
    let top_edge_em: f64 = crate::render::pdf::font_cap_height_em(family)?;
    let advance_em: f64 = word_natural_line_em(runs, pitch_em);
    let leading_em: f64 = advance_em - top_edge_em - bottom_edge_em;
    (leading_em >= 0.0).then(|| leading_em * paragraph_font_size_pt(runs))
}

/// The height one header or footer line takes, in points.
///
/// Word's natural line for the paragraph's resolved face and size — the hhea
/// line, or 1.3 times it for an East Asian one. A header taller than the band
/// `w:top - w:header` leaves has to grow the top margin, and this is the term
/// that measures it (issue #736).
///
/// `None` where the face's metrics are unknown, which leaves the band at the
/// declared size rather than guessing at a growth.
pub(super) fn word_line_advance_pt(runs: &[Run]) -> Option<f64> {
    let family: &str = east_asian_aware_metric_family(runs)?;
    let (_ascender_em, _descender_em, pitch_em) = crate::render::pdf::font_line_metrics_em(family)?;
    if pitch_em <= 0.0 {
        return None;
    }
    Some(word_natural_line_em(runs, pitch_em) * paragraph_font_size_pt(runs))
}

/// The line advance Word gives this paragraph before any grid is consulted:
/// the font's hhea line, or 1.3 times it when the line is set in an East Asian
/// face (issues #518, #643).
fn word_natural_line_em(runs: &[Run], word_pitch_em: f64) -> f64 {
    if line_takes_east_asian_metrics(runs) {
        EAST_ASIAN_LINE_HEIGHT_FACTOR * word_pitch_em
    } else {
        word_pitch_em
    }
}

/// The font size Word resolves a paragraph's line box against: the largest
/// size among its runs, falling back to the Word default when unset.
fn paragraph_font_size_pt(runs: &[Run]) -> f64 {
    largest_font_size_pt(runs.iter().filter_map(|run| run.style.font_size))
}

/// The same size, but only when a run actually declares one. `None` says the
/// paragraph inherits its size, so a length derived from it cannot be stated in
/// absolute points (issue #1115).
fn declared_paragraph_font_size_pt(runs: &[Run]) -> Option<f64> {
    runs.iter()
        .filter_map(|run| run.style.font_size)
        .reduce(f64::max)
}

/// The largest declared size, or Word's default when nothing declares one —
/// which, in that fallback alone, is also what an `em` resolves against, since
/// the generator emits no document-wide `#set text(size:)`.
///
/// The equivalence does not extend to a declared size. A paragraph states a
/// `#set text(size:)` only where every run agrees on one, so an `em` length
/// derived from the largest declared size can still resolve against the 11pt
/// default — which is what pinned every hard-broken line under 11pt to 13.20pt
/// (issue #1115). Use [`declared_paragraph_font_size_pt`] and emit points when
/// the length has to carry the size it was derived from.
fn largest_font_size_pt(sizes: impl Iterator<Item = f64>) -> f64 {
    let largest: f64 = sizes.fold(f64::NAN, f64::max);
    if largest.is_nan() { 11.0 } else { largest }
}

/// The row-level East Asian answer every cell in a table row shares, decided
/// once per row so the whole row sits on one baseline (issue #498).
///
/// The two gates deliberately differ, mirroring the body path's asymmetry:
/// the line *box* keys on the face the row's lines are set in (issues #643,
/// #814), while a snapping document grid keys on the characters — Word does
/// not stretch a Latin-only line to the grid pitch (issue #354).
#[derive(Clone, Copy)]
pub(super) struct RowEastAsianMetrics {
    /// Whether any cell in the row carries East Asian characters — what a
    /// snapping grid and its `w:spacing w:after` absorption key on (issues
    /// #354, #500).
    pub has_east_asian_text: bool,
    /// Whether the row's lines take Word's East Asian line box — keyed on the
    /// resolved face, not the script of the characters (issues #643, #814).
    pub takes_east_asian_metrics: bool,
}

/// Whether a cell's grid-snapped line box already contains the paragraph's
/// `w:spacing w:after`, so the caller must not emit it a second time.
///
/// Mirrors the guard inside [`word_cell_line_box_settings`] exactly, including
/// its early return for a paragraph carrying its own line box — gating on
/// `has_east_asian_text` alone would strip the gap from those.
///
/// A declared `w:spacing w:line` is deliberately *not* a bail here, and must
/// not become one again: since issue #727 the box scales by the multiple
/// rather than being abandoned, so it absorbs the gap like any other, and
/// re-adding the clause would emit `w:after` twice on a grid-snapped
/// line-spaced row.
pub(super) fn cell_grid_absorbs_space_after(
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
    row_east_asian: RowEastAsianMetrics,
) -> bool {
    row_east_asian.has_east_asian_text
        && style.line_box.is_none()
        && line_grid_pitch.is_some_and(|pitch| pitch > 0.0)
}

/// The one text line every cell of a *tight* spreadsheet row seats on
/// (issue #839).
///
/// Excel prints a single-line sheet row's cells on one baseline: the native
/// export of `09_expense_report_en` puts a `vertical="bottom"` amount column
/// and its `vertical="center"` neighbours all at y=143.00 in a 14pt track,
/// and `04_payroll_ko`'s fixed 합계 row seats its centred Korean label and
/// bottom-aligned numbers together at y=218.00. A track that cannot hold more
/// than one line leaves the alignments nowhere to differ, so the row's line
/// is resolved once — one metric family, one size — and every cell centres
/// that one box (see `generate_table_cell` for why centring is the anchor
/// taken). Reading each cell's own face and alignment instead split one row
/// across baselines 0.13–1.40pt apart, keyed to the cell's column.
#[derive(Clone)]
pub(super) struct SheetRowLine {
    /// The family whose metrics pace the row's shared line.
    pub metric_family: String,
    /// The size the row's line resolves at: the largest run size in the row.
    pub font_size_pt: f64,
}

/// The advance, in points, of one line set in `metric_family` at
/// `font_size_pt` under the row's East Asian answer — the height the
/// tightness gate compares the row's content box against (issue #839).
///
/// `None` when the family's metrics are unknown, in which case no shared row
/// line can be resolved either.
pub(super) fn sheet_row_line_advance_pt(
    metric_family: &str,
    font_size_pt: f64,
    takes_east_asian_metrics: bool,
) -> Option<f64> {
    let (_ascender_em, _descender_em, pitch_em) =
        crate::render::pdf::font_line_metrics_em(metric_family)?;
    if pitch_em <= 0.0 {
        return None;
    }
    let advance_em: f64 = if takes_east_asian_metrics {
        EAST_ASIAN_LINE_HEIGHT_FACTOR * pitch_em
    } else {
        pitch_em
    };
    Some(advance_em * font_size_pt)
}

/// The baseline-to-baseline advance Excel gives a sheet cell's lines when it
/// sets more than one — a wrapped cell, or one carrying its own line breaks —
/// for a face and size the sweep below measured (issue #1163).
///
/// **It is not the face's `hhea` line**, which is what the cell's box spans,
/// and no rounding of that line reproduces it: Arial and Times New Roman
/// declare the same 2355/2048 line yet advance 29pt against 27pt at 24pt,
/// while Verdana's taller 2489 advances 30pt there. The best fit over the
/// whole sweep of `ROUND(ascender) + ROUND(line gap) + ROUND(descender) + c`,
/// across every combination of floor/ceil/round and the `hhea`, `OS/2 win`
/// and `OS/2 typo` metric sets, misses 36 of 126 samples; a per-face
/// `ROUND(k x size + c)` misses 9 of Segoe UI's 21 alone. So this is a
/// measured table, exactly as the row-height recompute in
/// `parser::xlsx_cells` is, and for the same reason.
///
/// **Method.** Native Excel-for-Mac exports of purpose-built probe workbooks
/// (`/Volumes/T7/scratch/issue-1163`, built with openpyxl): one wrapped,
/// top-aligned cell per (face, size) in a fixed track far taller than its
/// text, read back with `mutool draw -F stext`. Every block advances by one
/// constant whole number of points.
///
/// Four factors were swept and none of them moves it:
///
/// - **the workbook's Normal font** — the same 126-cell sweep exported under
///   `fonts[0]` of Arial 10, Segoe UI 10 and Calibri 11 returns the identical
///   table, so the advance keys on the *cell's* face, and the printed grid's
///   `x0.92` compaction (which Calibri 11 turns on) does not reach it
/// - **the row's track** — Arial 14 advances 17.00pt in 60, 90, 120, 160, 200
///   and 240pt tracks alike
/// - **the characters** — a Malgun Gothic column set in Korean answers its
///   Latin twin at all fourteen sizes, the same independence issue #1060
///   measured for the line box itself
/// - **the column width** — the large-size batch was exported twice at two
///   different widths, wrapping at different points, and agreed everywhere
///
/// **Confirmed against a real workbook.** The sheet of #1163 prints at a 0.82
/// fit-to-page scale, and its three wrapped cells measure 17.22, 13.12 and
/// 31.16pt between baselines — 21.00, 16.00 and 38.00pt unscaled, which is
/// exactly what this table gives Segoe UI 14, Segoe UI 12 and Century Gothic
/// 30.
///
/// `None` for a family or size the sweep did not reach, which leaves the
/// caller on the face's `hhea` line. **Nothing here interpolates**, and no
/// family may be lent another's column: Arial and Times New Roman share a
/// `hhea` line and diverge here, Aptos and Aptos Narrow were each swept in
/// full before being paired, and so were `Malgun Gothic` and `맑은 고딕`.
pub(super) fn sheet_wrapped_line_advance_pt(family: &str, font_size_pt: f64) -> Option<f64> {
    let advances: &[f64; SHEET_ADVANCE_SIZES_PT.len()] = &SHEET_WRAPPED_LINE_ADVANCES
        .iter()
        .find(|face| {
            face.families
                .iter()
                .any(|candidate| candidate.eq_ignore_ascii_case(family))
        })?
        .advances_pt;
    SHEET_ADVANCE_SIZES_PT
        .iter()
        .position(|size_pt| (font_size_pt - size_pt).abs() < 0.01)
        .map(|index| advances[index])
}

/// The font sizes [`SHEET_WRAPPED_LINE_ADVANCES`] states an advance for, in
/// points. Every series carries one entry per size, in this order.
const SHEET_ADVANCE_SIZES_PT: [f64; 21] = [
    8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 20.0, 22.0, 24.0, 26.0, 28.0,
    30.0, 32.0, 36.0, 40.0, 48.0,
];

/// One face's measured advances, and the family names that select it.
struct SheetLineAdvances {
    /// Matched whole and case-insensitively, never as a prefix.
    families: &'static [&'static str],
    /// One advance in points per entry of [`SHEET_ADVANCE_SIZES_PT`].
    advances_pt: [f64; SHEET_ADVANCE_SIZES_PT.len()],
}

/// What the sweep described on [`sheet_wrapped_line_advance_pt`] measured, one
/// series per face.
#[rustfmt::skip]
const SHEET_WRAPPED_LINE_ADVANCES: [SheetLineAdvances; 13] = [
    SheetLineAdvances { families: &["Arial"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 21.0, 22.0, 24.0, 27.0, 29.0, 32.0, 33.0, 35.0, 38.0, 43.0, 46.0, 56.0] },
    SheetLineAdvances { families: &["Times New Roman"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 21.0, 23.0, 26.0, 27.0, 30.0, 32.0, 34.0, 37.0, 41.0, 46.0, 54.0] },
    SheetLineAdvances { families: &["Verdana"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0, 32.0, 35.0, 37.0, 40.0, 45.0, 49.0, 59.0] },
    SheetLineAdvances { families: &["Tahoma"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0, 32.0, 35.0, 37.0, 40.0, 44.0, 49.0, 59.0] },
    SheetLineAdvances { families: &["Georgia"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 20.0, 21.0, 22.0, 23.0, 26.0, 28.0, 31.0, 33.0, 36.0, 37.0, 42.0, 47.0, 56.0] },
    SheetLineAdvances { families: &["Courier New"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 21.0, 24.0, 26.0, 28.0, 31.0, 32.0, 35.0, 38.0, 42.0, 46.0, 55.0] },
    SheetLineAdvances { families: &["Helvetica"], advances_pt:
        [11.0, 12.0, 13.0, 15.0, 16.0, 17.0, 18.0, 19.0, 21.0, 22.0, 22.0, 25.0, 27.0, 30.0, 32.0, 34.0, 37.0, 39.0, 44.0, 49.0, 59.0] },
    SheetLineAdvances { families: &["Segoe UI"], advances_pt:
        [11.0, 13.0, 14.0, 16.0, 16.0, 20.0, 21.0, 23.0, 23.0, 25.0, 26.0, 28.0, 31.0, 33.0, 38.0, 40.0, 43.0, 45.0, 50.0, 57.0, 67.0] },
    SheetLineAdvances { families: &["Calibri"], advances_pt:
        [11.0, 12.0, 14.0, 14.0, 15.0, 16.0, 18.0, 19.0, 20.0, 22.0, 23.0, 25.0, 28.0, 30.0, 33.0, 36.0, 38.0, 40.0, 45.0, 50.0, 60.0] },
    SheetLineAdvances { families: &["Aptos", "Aptos Narrow"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 21.0, 22.0, 23.0, 26.0, 28.0, 31.0, 32.0, 35.0, 37.0, 40.0, 45.0, 50.0, 60.0] },
    SheetLineAdvances { families: &["Century Gothic"], advances_pt:
        [11.0, 12.0, 13.0, 14.0, 16.0, 17.0, 18.0, 19.0, 21.0, 22.0, 23.0, 25.0, 28.0, 30.0, 33.0, 35.0, 38.0, 40.0, 45.0, 50.0, 60.0] },
    SheetLineAdvances { families: &["Malgun Gothic", "맑은 고딕"], advances_pt:
        [13.0, 14.0, 15.0, 17.0, 18.0, 19.0, 20.0, 22.0, 23.0, 26.0, 27.0, 30.0, 32.0, 35.0, 37.0, 40.0, 44.0, 47.0, 52.0, 59.0, 69.0] },
    SheetLineAdvances { families: &["Gulim"], advances_pt:
        [12.0, 13.0, 14.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 22.0, 23.0, 25.0, 27.0, 30.0, 32.0, 34.0, 36.0, 40.0, 44.0, 49.0, 59.0] },
];

/// Where a spreadsheet cell's fixed row track sits, so the cell's line can be
/// seated on the baseline the native export prints (issue #1063).
///
/// `None` outside that regime: a Word table, an auto-height sheet row (whose
/// track is the content's own answer, not Excel's), a top-aligned cell (a
/// seat this issue did not settle), or a cell spanning several tracks.
#[derive(Clone, Copy)]
pub(super) struct SheetCellSeat {
    /// The row's printed track height, in points.
    pub track_pt: f64,
    /// The cell's own top inset — its padding plus its share of the border.
    pub inset_top_pt: f64,
    /// The cell's own bottom inset.
    pub inset_bottom_pt: f64,
    /// The gap the bottom seat keeps under the baseline however small the
    /// font, in points — this workbook's floor (issues #1097, #1199).
    pub descent_floor_pt: f64,
}

/// The baseline Excel prints for one line of a face whose bare `hhea` numbers
/// are `ascent_em`/`descent_em`/`line_gap_em`, set at `font_size_pt` in a
/// `track_pt` row, as an offset below the track's **top boundary**
/// (issues #1063, #1161).
///
/// Excel lays a sheet out in whole sheet-space points, and it rounds the three
/// `hhea` numbers into points **separately** before composing the line box:
/// the ascender truncated, the line gap rounded up, the descender rounded.
/// The box is the three added together, the baseline sits the gap plus the
/// ascent below its top, and the box is centred in the row's own track — not
/// in the cell's inset content box, which four native probe exports show has
/// no say in the vertical seat — with an odd leftover point going to the
/// space *above* the line. A fitted sheet scales that completed seat onto the
/// printed page (issue #1238).
///
/// Measured on native Excel-for-Mac exports of purpose-built probe workbooks
/// (`/Volumes/T7/scratch/issue-1063/probe`, reproduced in
/// `sheet_cell_line_seat_reproduces_the_native_excel_probe`): a row-height
/// sweep from 12pt to 60pt at Arial 10, a font-size sweep from 8pt to 44pt in
/// 40pt and 60pt tracks, and a border/no-border and Normal-font pairing that
/// changed nothing. All 28 samples land on this rule exactly.
///
/// Those probes are Arial-only, and Arial's 67/2048 line gap makes the
/// separate rounding indistinguishable from folding the gap into the ascender
/// first: the two readings differ by 1pt in the ascent and 2pt in the box,
/// which cancel in the centred seat. A face declaring **no** line gap
/// separates them, and the folded reading then seats every line a point low —
/// Segoe UI, Century Gothic and Calibri all measured that way on the workbook
/// of #1161, reproduced in
/// `sheet_cell_line_seat_reproduces_a_face_with_no_line_gap`.
pub(super) fn sheet_cell_baseline_from_track_top_pt(
    track_pt: f64,
    ascent_em: f64,
    descent_em: f64,
    line_gap_em: f64,
    font_size_pt: f64,
    print_scale: Option<f64>,
) -> f64 {
    // The parser has already folded a fit-to-page scale into the track and
    // font. Excel instead snaps their declared sheet-space values, then
    // scales that answer onto the printed page (issue #1238).
    let scale: f64 = print_scale.filter(|scale| *scale > 0.0).unwrap_or(1.0);
    let sheet_font_size_pt: f64 = font_size_pt / scale;
    let sheet_track_pt: f64 = track_pt / scale;
    let ascent_pt: f64 = (ascent_em * sheet_font_size_pt).floor();
    let line_gap_pt: f64 = (line_gap_em * sheet_font_size_pt).ceil();
    let descent_pt: f64 = (descent_em * sheet_font_size_pt).round();
    let above_baseline_pt: f64 = line_gap_pt + ascent_pt;
    let line_pt: f64 = above_baseline_pt + descent_pt;
    (((sheet_track_pt - line_pt) / 2.0).ceil() + above_baseline_pt) * scale
}

/// The gap Excel never closes between a bottom-aligned sheet cell's baseline
/// and its row's bottom boundary, in the workbooks whose printed grid keeps
/// its declared row tracks (issue #1097).
///
/// Flat across every factor probed: the row's track height (12, 13, 14, 15,
/// 16, 17, 18, 20, 22, 25, 30, 40, 45 and 60pt all give 4.00), the cell's own
/// border, its weight, and the workbook's Normal font (Arial 10, Calibri 11
/// and Arial 20 all give 4.00). So it is one distance, not a face's metric in
/// disguise — though only Arial cells were measured, so a face whose descent
/// is far from Arial's could yet show it as something else.
pub(crate) const SHEET_CELL_MIN_DESCENT_SEAT_PT: f64 = 4.0;

/// The 3pt gap measured for the remapped Calibri/Aptos family in the native
/// floor probe matrix, one point below the script-face theme family
/// (issue #1199). Face-specific row-track mappings measured later do not
/// choose this floor.
///
/// Measured on native Excel-for-Mac exports of `10_kpi_tracker_en.xlsx`, whose
/// A11 note is the corpus's one bottom-aligned cell small enough to tell a
/// floor from a bare descent. Ruling that cell with a thin box border makes
/// Excel print its row's own boundaries, so the seat is read rather than
/// inferred; a re-export with only the border added moves no text, and one
/// sweeping the note's `sz` over eleven sizes gives:
///
/// | size | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 16 | 18 | 20 | 24 |
/// | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
/// | seat | 3 | 3 | 3 | 3 | 3 | 3 | 3 | 3 | 4 | 4 | 5 |
///
/// which is `max(3, round(0.211914 x size))` at every one of them. Issue #1097
/// read this family as having no floor at all from its one available sample,
/// `09_expense_report_en`'s Arial Bold 14 title at 3.00pt — a size whose bare
/// rounded descent is already 3, so it cannot separate the two readings. The
/// four sizes at or under 11 can, and they floor.
pub(crate) const COMPACTED_SHEET_CELL_MIN_DESCENT_SEAT_PT: f64 = 3.0;

/// The descent Excel rests a bottom-aligned sheet cell's last line on: the
/// face's `hhea` descent at a whole number of points, sitting on the row's own
/// bottom boundary (issue #1063) — but never closer to that boundary than the
/// workbook's own floor (issues #1097, #1199), and never on that rounded
/// descent at all for a face [`sheet_cell_measured_seat_pt`] measured a seat
/// of its own for (issue #1208).
///
/// The boundary itself, with no inset under it: over the probe's size sweep
/// the fitted inset is 0.00-0.14pt.
///
/// `descent_floor_pt` is the workbook-wide distance
/// `xlsx_cells::bottom_aligned_descent_floor_pt` reads, which carries the
/// measurement behind it: [`COMPACTED_SHEET_CELL_MIN_DESCENT_SEAT_PT`] for the
/// remapped Calibri/Aptos family in the floor probe matrix, and
/// [`SHEET_CELL_MIN_DESCENT_SEAT_PT`] for its script-face theme family. Both
/// families reproduce their native probes at every size swept; later
/// face-specific printed-grid mappings do not choose this separately measured
/// floor. The two values differ only under a font small enough to separate
/// them, which on Arial is 11pt and below.
pub(super) fn sheet_cell_descent_pt(
    family: &str,
    descent_em: f64,
    font_size_pt: f64,
    print_scale: Option<f64>,
    descent_floor_pt: f64,
) -> f64 {
    // Excel reads every seat component at the size the cell declares and
    // prints the result through the sheet's fit-to-page scale, which the
    // parser has already folded into `font_size_pt`. That includes both the
    // measured face series and the rounded hhea descent/floor (issue #1238).
    let scale: f64 = print_scale.filter(|scale| *scale > 0.0).unwrap_or(1.0);
    let sheet_font_size_pt: f64 = font_size_pt / scale;
    sheet_cell_measured_seat_pt(family, sheet_font_size_pt)
        .unwrap_or_else(|| (descent_em * sheet_font_size_pt).round())
        .max(descent_floor_pt)
        * scale
}

/// The seat measured for `family` at `font_size_pt`, or `None` where no sweep
/// reached that face and size or where the sweep could only see the workbook's
/// own floor (issue #1208).
///
/// **Method.** Native Excel-for-Mac exports of purpose-built probe workbooks
/// (`/Volumes/T7/scratch/issue-1208/probe`, built with openpyxl): one
/// bottom-aligned cell per (face, size) in an auto-height row, each ruled by a
/// thin box border so Excel prints the row's own boundaries rather than
/// leaving them to be inferred, and the baseline read back with
/// `mutool draw -F trace`. A border-free control block reproduced the same
/// baselines, so ruling the cell moves nothing.
///
/// **The seat does not depend on the track.** Malgun Gothic's column was swept
/// a second time in fixed 60pt tracks and answered the same seat at all
/// fifteen sizes of that run. That is what rules out reading the seat as an
/// ascent measured down from the row's *top*, which fits the auto rows of
/// issue #1208's fixture just as well but would put a 60pt track's baseline
/// 25pt lower than a 35pt one's.
///
/// **No rounded descent fits these three faces.** Malgun Gothic seats 4pt at
/// 14pt, which needs a descent under 0.3214em, and 14pt at 40pt, which needs
/// one of at least 0.3375em. So this is a measured table, exactly as
/// [`SHEET_WRAPPED_LINE_ADVANCES`] is, and it shares that sweep's size axis.
///
/// **It is an exception list, not a replacement.** The same probe swept Arial,
/// Verdana, Georgia, Times New Roman, Tahoma, Courier New, Century Gothic,
/// Calibri, Aptos and MS Gothic over the same twenty-one sizes, and
/// `max(floor, round(descent x size))` reproduces all two hundred and ten of
/// those samples. MS Gothic conforming is why the exception cannot be stated
/// as "an East Asian face".
///
/// `None` below each series' first entry, because the probe workbook floors at
/// [`SHEET_CELL_MIN_DESCENT_SEAT_PT`] and so cannot tell a face value of 4pt
/// from a floored one — and the floor is exactly what differs between the two
/// workbook families (issue #1199). The rounded descent stands in there, and
/// on all three faces it lands at or under both floors (3pt at the largest
/// such size on each), so the family's own floor decides as it did before.
///
/// **Nothing here interpolates**, and no family may be lent another's column:
/// Gulim and Batang share their `hhea` metrics but were each swept in full
/// before being given the same numbers. Each Korean alias rides with the face
/// `font_subst` resolves it to rather than a sweep of its own, the same
/// pairing [`SHEET_WRAPPED_LINE_ADVANCES`] makes for `맑은 고딕`.
fn sheet_cell_measured_seat_pt(family: &str, font_size_pt: f64) -> Option<f64> {
    let seats: &[Option<f64>; SHEET_ADVANCE_SIZES_PT.len()] = &SHEET_CELL_SEATS
        .iter()
        .find(|face| {
            face.families
                .iter()
                .any(|candidate| candidate.eq_ignore_ascii_case(family))
        })?
        .seats_pt;
    SHEET_ADVANCE_SIZES_PT
        .iter()
        .position(|size_pt| (font_size_pt - size_pt).abs() < 0.01)
        .and_then(|index| seats[index])
}

/// One face's measured bottom-aligned seats, and the family names that select
/// it.
struct SheetCellSeats {
    /// Matched whole and case-insensitively, never as a prefix.
    families: &'static [&'static str],
    /// One seat in points per entry of [`SHEET_ADVANCE_SIZES_PT`], which this
    /// series shares because both were swept on one probe.
    seats_pt: [Option<f64>; SHEET_ADVANCE_SIZES_PT.len()],
}

/// What the sweep described on [`sheet_cell_measured_seat_pt`] measured, one
/// series per face that the rounded descent does not reproduce.
#[rustfmt::skip]
const SHEET_CELL_SEATS: [SheetCellSeats; 3] = [
    SheetCellSeats { families: &["Malgun Gothic", "맑은 고딕"], seats_pt:
        [None, None, None, None, None, None, None, Some(5.0), Some(5.0), Some(6.0), Some(6.0), Some(7.0), Some(7.0), Some(8.0), Some(8.0), Some(9.0), Some(10.0), Some(11.0), Some(12.0), Some(14.0), Some(16.0)] },
    SheetCellSeats { families: &["Gulim", "굴림"], seats_pt:
        [None, None, None, None, None, None, None, None, None, None, None, None, None, None, Some(5.0), Some(5.0), Some(5.0), Some(7.0), Some(7.0), Some(8.0), Some(10.0)] },
    SheetCellSeats { families: &["Batang", "바탕"], seats_pt:
        [None, None, None, None, None, None, None, None, None, None, None, None, None, None, Some(5.0), Some(5.0), Some(5.0), Some(7.0), Some(7.0), Some(8.0), Some(10.0)] },
];

/// A table-cell paragraph's fixed line box, resolved at the paragraph's own
/// font size — or at the row's shared family and size when a tight
/// spreadsheet row supplies a [`SheetRowLine`] (issue #839). `top_em`/
/// `bottom_em` are the metric edges the cell emits; `leading_pt` is the gap
/// between line boxes.
pub(super) struct CellLineBox {
    pub top_em: f64,
    pub bottom_em: f64,
    /// Zero for a Word or PowerPoint cell whose box is not re-seated on the
    /// descender. It carries the surplus that seat removes from below the
    /// baseline, so multi-line advance is unchanged (issue #618) — and, in a
    /// spreadsheet cell, whatever Excel's measured wrapped-line pitch adds to
    /// the face's hhea box, which is independent of that seat and applies to
    /// top- and centre-aligned cells alike (issue #1163).
    pub leading_pt: f64,
    pub font_size_pt: f64,
}

/// Line-box settings for a **Word or sheet** table cell: a fixed box spanning
/// the font's full single-spacing (hhea) line — 1.3 times it for an East Asian row — seated at
/// the same constant ascent the body path uses. In the default symmetric
/// emission the box carries the whole line advance below the ascent with zero
/// leading, so a single-line cell occupies the full line height Word gives it
/// rather than only the tighter metric box (which left auto-height rows too
/// short, issue #396). When `seats_text_on_descender` is set (bottom-aligned
/// spreadsheet cells in fixed-height rows), the box instead ends at the
/// font's descender and the removed sub-baseline surplus moves into leading,
/// so the last line's descent rests on the row's bottom inset edge while
/// multi-line advance is unchanged (issue #618) — unless `sheet_seat` re-seats
/// it, below. `None` when the font's metrics are unknown or the paragraph
/// carries its own line box; a declared
/// `w:spacing w:line` scales the box instead of suppressing it (issue #727).
///
/// The box also carries the paragraph's `w:spacing w:after` when a snapping
/// grid is in force, because Word snaps the line and that gap together (issues
/// #500, #503).
///
/// When `sheet_row_line` is `Some` — a tight fixed-track spreadsheet row,
/// gated by `sheet_row_shared_line` in the table codegen — the box resolves
/// at the row's shared metric family and size instead of this paragraph's
/// own, so every cell of the row carries the same box and lands on one
/// baseline as Excel prints it (issue #839).
///
/// When `sheet_seat` is `Some` — a spreadsheet cell in a fixed row track,
/// gated by `generate_table_cell` — the box is redistributed around the
/// baseline so it seats where Excel prints it, overriding both seats above:
/// the ascent and descent are each rounded to a whole point and the line is
/// centred in the row's own *track*, or, under bottom alignment, its rounded
/// descent rests on the track's own bottom boundary. The box's height, and
/// with it the row's advance, is unchanged (issue #1063).
///
/// When `sheet_print_scale` is `Some` — any cell of a spreadsheet, carrying
/// the sheet's `fitToWidth` factor — a face and size
/// [`sheet_wrapped_line_advance_pt`] measured pace their *lines* on Excel's
/// own pitch instead of on that box, and the difference leaves the box
/// untouched and rides as leading. So a sheet cell's leading is non-zero
/// whatever its vertical alignment, where every seat above leaves it at zero
/// (issue #1163).
///
/// A **slide's** table cell does not come here: it paces on PowerPoint's flat
/// 1.2em line via [`powerpoint_line_height_settings`], like the slide's own
/// text boxes (issue #663).
// Same shape as `word_cell_line_box`, which this only formats the result of;
// grouping the arguments would have to split both, and the sheet's three are
// each read on their own inside the box.
#[allow(clippy::too_many_arguments)]
pub(super) fn word_cell_line_box_settings(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
    row_east_asian: RowEastAsianMetrics,
    seats_text_on_descender: bool,
    sheet_row_line: Option<&SheetRowLine>,
    sheet_seat: Option<SheetCellSeat>,
    sheet_print_scale: Option<f64>,
) -> Option<String> {
    let line_box: CellLineBox = word_cell_line_box(
        runs,
        style,
        line_grid_pitch,
        row_east_asian,
        seats_text_on_descender,
        sheet_row_line,
        sheet_seat,
        sheet_print_scale,
    )?;
    Some(format!(
        // The descent is negated here rather than written behind a literal
        // `-`: a sheet cell's descent can fall *short* of the cell's bottom
        // inset (issue #1063), and a negative edge behind that sign emitted
        // `--0.02em`.
        "#set text(top-edge: {}em, bottom-edge: {}em)\n#set par(leading: {}pt)\n",
        format_f64(line_box.top_em),
        format_f64(-line_box.bottom_em),
        format_f64(line_box.leading_pt)
    ))
}

/// The line box behind [`word_cell_line_box_settings`], exposed so the spill
/// wrapper can size its clip box and strut from the same numbers the block
/// emits (issue #618).
#[allow(clippy::too_many_arguments)]
pub(super) fn word_cell_line_box(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
    row_east_asian: RowEastAsianMetrics,
    seats_text_on_descender: bool,
    sheet_row_line: Option<&SheetRowLine>,
    sheet_seat: Option<SheetCellSeat>,
    sheet_print_scale: Option<f64>,
) -> Option<CellLineBox> {
    if style.line_box.is_some() {
        return None;
    }
    // A tight spreadsheet row's box resolves at the row's one family and
    // size, not this cell's: Excel prints every cell of such a row on one
    // baseline, and per-cell metrics split it by the descender difference
    // (issue #839).
    let (family, font_size): (&str, f64) = match sheet_row_line {
        Some(row_line) => (row_line.metric_family.as_str(), row_line.font_size_pt),
        None => (
            east_asian_aware_metric_family(runs)?,
            paragraph_font_size_pt(runs),
        ),
    };
    let (ascender_em, descender_em, word_pitch_em) =
        crate::render::pdf::font_line_metrics_em(family)?;
    let metric_em: f64 = ascender_em + descender_em;
    if metric_em <= 0.0 || word_pitch_em <= 0.0 {
        return None;
    }
    // The row decides whether its lines are East Asian, not this cell: reading
    // each cell's own text put a Korean label and its numeric neighbours on
    // line boxes of different heights, splitting one row across two baselines
    // 4.29pt apart (issue #498). So both the 1.3 line-height bonus and the
    // ascent excess it implies key on the row's answer.
    let natural_em: f64 = if row_east_asian.takes_east_asian_metrics {
        EAST_ASIAN_LINE_HEIGHT_FACTOR * word_pitch_em
    } else {
        word_pitch_em
    };
    // A grid-snapped row snaps the line *plus* the paragraph's own `w:spacing
    // w:after`, not the line alone. Snapping the bare line and then adding the
    // gap outside it made every grid-scoped row 1.06pt too tall, because
    // 12.64pt of Malgun and a 1.5pt gap both fit inside one 18pt line where
    // 18 + 1.5 does not (issues #500, #503). `cell_grid_absorbs_space_after`
    // gates the caller's matching suppression of the trailing gap; the two must
    // agree. The grid keys on the row's *text*, not its face: Word does not
    // stretch a Latin-only line to the grid pitch (issues #354, #814).
    let advance_em: f64 = match line_grid_pitch.filter(|pitch| *pitch > 0.0) {
        Some(pitch) if row_east_asian.has_east_asian_text => {
            // Same two-way choice as the body path (issue #508), with the
            // paragraph's `w:after` inside the quantity being compared.
            let natural_pt: f64 = natural_em * font_size + style.space_after.unwrap_or(0.0);
            let advance_pt: f64 = if natural_pt <= pitch {
                pitch
            } else {
                natural_pt
            };
            advance_pt / font_size
        }
        _ => natural_em,
    };
    // `w:spacing w:line` states the advance the same way inside a cell as it
    // does in the body: a proportional rule scales Word's own line, an exact
    // one replaces it outright. Bailing on any declared spacing left the
    // paragraph with no fixed box at all, so the multiple never applied and the
    // advance fell back to whatever Typst chose (issue #727).
    let advance_em: f64 = match style.line_spacing {
        None => advance_em,
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => advance_em * factor,
        Some(LineSpacing::Exact(points)) if points > 0.0 => points / font_size,
        // A non-positive rule states nothing usable; Word ignores it.
        Some(_) => advance_em,
    };
    let excess_em: f64 = if row_east_asian.takes_east_asian_metrics {
        EAST_ASIAN_ASCENT_EXCESS * word_pitch_em
    } else {
        0.0
    };
    let top_em: f64 = ascender_em + excess_em;
    // Excel rests a bottom-aligned cell's last line on its descender: the
    // descent bottom sits on the row's bottom inset edge with all slack above.
    // The symmetric box carries the East Asian 0.15-line surplus below the
    // baseline, which floated bottom-aligned Korean cells above where Excel
    // prints them (issue #618). Ending the box at the descender and moving the
    // surplus into leading keeps multi-line advance identical.
    //
    // This once cited the header/footer path as applying the same rule with
    // `bottom-edge: "descender"`. It no longer does: a header's `w:pBdr` gap
    // measures from the East Asian line box's bottom, not the descender, and
    // reads that from `word_line_box_descent_em` (issue #737). The seating
    // here is Excel's and is unaffected.
    // TODO(#618 follow-up: leading is one per-paragraph pt value derived from
    // the max run size, so mixed-font-size wrapped lines gain
    // 0.15*pitch*(max-line) advance error; needs per-line seating if a real
    // sheet exhibits it).
    let (bottom_em, leading_pt): (f64, f64) = if seats_text_on_descender {
        (
            descender_em,
            ((advance_em - top_em - descender_em) * font_size).max(0.0),
        )
    } else {
        (advance_em - top_em, 0.0)
    };
    // A sheet cell in a fixed track seats its line where Excel prints it, not
    // where the cell's own inset box would centre it (issue #1063). Typst
    // places the box inside that inset box, so the seat is expressed by
    // redistributing the box around the baseline — its height, and with it the
    // row's advance, is unchanged.
    let (top_em, bottom_em, leading_pt): (f64, f64, f64) = match sheet_seat {
        None => (top_em, bottom_em, leading_pt),
        Some(seat) if seats_text_on_descender => {
            // Typst rests the box's bottom edge on the inset content bottom;
            // Excel rests the descender on the row boundary itself, one inset
            // lower.
            let bottom_em: f64 = (sheet_cell_descent_pt(
                family,
                descender_em,
                font_size,
                sheet_print_scale,
                seat.descent_floor_pt,
            ) - seat.inset_bottom_pt)
                / font_size;
            (
                top_em,
                bottom_em,
                ((advance_em - top_em - bottom_em) * font_size).max(0.0),
            )
        }
        Some(seat) => {
            // Typst centres the box in the inset content box, whose centre is
            // this far below the track's top boundary.
            let content_mid_pt: f64 =
                (seat.inset_top_pt + (seat.track_pt - seat.inset_bottom_pt)) / 2.0;
            // `font_line_metrics_em` folds the line gap into `ascender_em`,
            // where Word wants it; Excel rounds the gap into whole points on
            // its own, so split it back out here (issue #1161).
            let line_gap_em: f64 = crate::render::pdf::font_line_gap_em(family).unwrap_or(0.0);
            let baseline_pt: f64 = sheet_cell_baseline_from_track_top_pt(
                seat.track_pt,
                ascender_em - line_gap_em,
                descender_em,
                line_gap_em,
                font_size,
                sheet_print_scale,
            );
            let top_em: f64 = advance_em / 2.0 + (baseline_pt - content_mid_pt) / font_size;
            (top_em, advance_em - top_em, leading_pt)
        }
    };
    // Excel paces a sheet cell's *lines* on a measured per-face advance that
    // is not the face's hhea line, so the surplus rides as leading rather than
    // inside the box: the box, and with it every seat measured against it,
    // stays exactly where issues #618/#839/#1063 put it, and only a cell that
    // sets a second line can see the difference (issue #1163).
    // Excel evaluates that advance at the size the cell *declares* and prints
    // it through the sheet's `fitToWidth` scale, which the parser has already
    // folded into `font_size`; the scaled advance is no whole number of points.
    let leading_pt: f64 = match sheet_print_scale
        .filter(|scale| *scale > 0.0)
        .and_then(|scale| {
            sheet_wrapped_line_advance_pt(family, font_size / scale)
                .map(|advance_pt| advance_pt * scale)
        }) {
        Some(advance_pt) => advance_pt - (top_em + bottom_em) * font_size,
        None => leading_pt,
    };
    Some(CellLineBox {
        top_em,
        bottom_em,
        leading_pt,
        font_size_pt: font_size,
    })
}

/// The top-up that raises the font's typographic metric box to Word's line
/// advance — its hhea single-space line, or the document grid pitch for East
/// Asian text under a `w:docGrid`. `word_line_height_settings` folds this
/// into the fixed line-box height rather than emitting it as `par(leading:)`
/// whitespace between boxes, because Typst inserts that only *between* the
/// lines of one paragraph and every paragraph then came up one top-up short
/// (issues #354, #452). A proportional `w:lineRule="auto"` scales the result
/// rather than replacing it, because that is what Word's own rule means.
/// `None` when the paragraph states an exact advance, carries its own line
/// box, or the font's metrics are unknown — the treatment does not apply
/// then.
pub(super) fn word_line_leading_pt(
    runs: &[Run],
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) -> Option<f64> {
    if style.line_box.is_some() {
        return None;
    }
    // `w:lineRule="auto"` scales Word's own line rather than replacing it:
    // `w:line="278"` means 1.158 of the line this function computes. Bailing
    // out on any `w:spacing w:line` left those paragraphs to Typst's default
    // leading, which knows nothing of the East Asian line — 15.4pt against
    // Word's 19.9pt on the technical brief (issue #575). An exact rule states
    // the advance outright and is still handled as a plain `par(leading:)`.
    let proportion: f64 = match style.line_spacing {
        None => 1.0,
        Some(LineSpacing::Proportional(factor)) if factor > 0.0 => factor,
        Some(_) => return None,
    };
    let family: &str = east_asian_aware_metric_family(runs)?;
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

    // Word's single spacing is the font's full hhea line, which the metric pair
    // sums to directly (issue #508) - so for Latin text this top-up is zero and
    // the subtraction below is just a guard for a face whose reported pitch
    // exceeds its own ascent-plus-descent. East Asian lines get 30% more
    // (issue #518).
    let natural_line_pt: f64 = word_natural_line_em(runs, word_pitch_em) * font_size;

    // Word only snaps East Asian *text* to the document grid, so this gate
    // stays keyed on the characters even though the line height above is keyed
    // on the face: a Latin-only paragraph keeps whatever advance its own line
    // asks for rather than being stretched to the grid pitch (native Word GT:
    // Arial 10.5 lines stay 12pt in a Korean document). Snapping Latin
    // paragraphs inflated every Western document by 30-50% (issue #354).
    //
    // That advance is no longer always the bare hhea line: a Latin-only
    // paragraph set in a CJK face now carries the 1.3 factor (issue #643). It
    // still does not snap.
    let advance_pt: f64 = match line_grid_pitch {
        Some(pitch) if pitch > 0.0 && has_east_asian_text(runs) => {
            // A grid line never compresses text below the height its font
            // needs, and Word chooses between exactly two advances, never a
            // multiple: the grid pitch when the natural line fits inside one
            // grid line, otherwise the natural line untouched (issues #402,
            // #508).
            if natural_line_pt <= pitch {
                pitch
            } else {
                natural_line_pt
            }
        }
        _ => natural_line_pt,
    };
    Some((advance_pt * proportion - line_box_pt).max(0.0))
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

/// The paragraph's `w:spacing` gaps, for a parameter list that already has a
/// first entry (every parameter is prefixed with a comma). They belong to
/// the outermost block, so an indent wrapper does not separate them from the
/// neighbouring paragraphs they collapse against.
fn write_block_spacing_params(out: &mut String, style: &ParagraphStyle) {
    if let Some(above) = style.space_before {
        let _ = write!(out, ", above: {}pt", format_f64(above));
    }
    if let Some(below) = style.space_after {
        let _ = write!(out, ", below: {}pt", format_f64(below));
    }
}

/// The paragraph's shading and borders, which Word paints across the
/// paragraph's own column — from the left indent to the right indent — so
/// they belong to the innermost block (issue #464).
fn write_block_decoration_params(out: &mut String, style: &ParagraphStyle) {
    let decorated: bool = style.background.is_some() || style.border.is_some();
    if let Some(background) = style.background {
        let _ = write!(out, ", fill: {}", rgb(&background));
    }
    if let Some(border) = &style.border {
        write_paragraph_border_params(
            out,
            border,
            style.border_space.as_deref().copied().unwrap_or_default(),
        );
    }
    if decorated {
        // `outset` widens what the block paints without moving the text in
        // it, which is what the overhang needs (issue #644).
        let _ = write!(
            out,
            ", outset: (x: {}pt)",
            format_f64(TEXT_COLUMN_DECORATION_OVERHANG_PT)
        );
    }
}

/// How far Word paints a rule or a shaded block past each edge of the text
/// column it belongs to: 0.02in. Applies to `w:pBdr` and `w:shd` on a body
/// paragraph and to a header or footer rule alike.
///
/// Measured, not assumed: across seven golden Word exports the widest rule on
/// page 1 spans 456.48pt against a 453.60pt text column — 2.88pt wider, or
/// 1.44pt at each edge — and the landscape file shows the same overhang on its
/// own wider column. The seven cover both `w:pBdr` rules and `w:shd` blocks,
/// and body paragraphs as well as header bands.
///
/// Whether the overhang is also independent of border weight and style is what
/// issue #644 reports, and it is consistent with these files all landing on one
/// value, but the per-style breakdown in that issue was not re-measured here.
pub(super) const TEXT_COLUMN_DECORATION_OVERHANG_PT: f64 = 1.44;

fn stroke_literal(side: &BorderSide) -> String {
    // Callers skip Double sides (drawn as overlays), so for every reachable
    // style this matches the table flavor of the shared stroke formatter.
    stroke_value(side, true)
}

/// Emit `stroke:`/`inset:` block parameters for the paragraph's borders.
/// Double rules are drawn as overlays (Typst strokes have no double style),
/// so those sides only reserve inset space here.
///
/// Each side reserves its own `w:space` plus *half* the rule's thickness,
/// because Typst centres a box stroke on the inset edge and the other half
/// falls outside it (issue #648). A double side is the exception: it emits no
/// box stroke at all, so it reserves the full three-width span its overlays
/// draw into. A fixed 4pt stood in for `w:space` until #520: a letterhead declaring 8pt then
/// pulled every line below it up by the difference, and the error is a step,
/// not a drift, so it survives to the bottom of the page.
fn write_paragraph_border_params(out: &mut String, border: &CellBorder, space: Insets) {
    let mut strokes: Vec<String> = Vec::new();
    let mut insets: Vec<String> = Vec::new();

    let mut push_side = |name: &str, side: &Option<BorderSide>, gap: f64| {
        let Some(side) = side else {
            return;
        };
        let reserved = if side.style == BorderLineStyle::Double {
            gap + double_rule_thickness(side.width)
        } else {
            strokes.push(format!("{name}: {}", stroke_literal(side)));
            // Typst centres a stroke on the inset edge, so half the rule
            // already falls outside the reserved band. Reserving the whole
            // width counted that half twice and put the rule half a width
            // low — 0.31pt measured on `04_resume_en`'s 0.75pt name rule,
            // against the native export (issue #648).
            gap + side.width / 2.0
        };
        insets.push(format!("{name}: {}pt", format_f64(reserved)));
    };
    push_side("top", &border.top, space.top);
    push_side("bottom", &border.bottom, space.bottom);
    push_side("left", &border.left, space.left);
    push_side("right", &border.right, space.right);

    if !strokes.is_empty() {
        let _ = write!(out, ", stroke: ({})", strokes.join(", "));
    }
    if !insets.is_empty() {
        let _ = write!(out, ", inset: ({})", insets.join(", "));
    }
}

/// A Word double rule draws two lines of the declared width separated by a gap
/// of the same width, so it stands three widths tall in total. Measured on
/// 06_official_letter_ko's `w:sz="8"` letterhead rule: 3pt, against the GT's
/// 2.93pt gap between the paragraph below it and the rule's far edge.
fn double_rule_thickness(width: f64) -> f64 {
    width * 3.0
}

/// Draw double-rule paragraph borders as two placed hairlines; Typst strokes
/// cannot render Word's double style. Only horizontal doubles occur in
/// practice (letterhead rules); vertical doubles fall back to a single
/// stroke drawn by `write_paragraph_border_params`.
fn write_paragraph_double_border_overlays(
    out: &mut String,
    border: &Option<Box<CellBorder>>,
    space: Insets,
) {
    let Some(border) = border else {
        return;
    };
    for (name, side, gap) in [
        ("top", &border.top, space.top),
        ("bottom", &border.bottom, space.bottom),
    ] {
        let Some(side) = side else {
            continue;
        };
        if side.style != BorderLineStyle::Double {
            continue;
        }
        let w = side.width;
        let near_dy = gap + w;
        let far_dy = gap + double_rule_thickness(w);
        let (align, sign) = if name == "top" {
            ("top", -1.0)
        } else {
            ("bottom", 1.0)
        };
        for dy in [near_dy, far_dy] {
            // A placed line spans the block's layout box, which `outset` does
            // not widen, so a double rule has to reach past both edges itself
            // to match the single-stroke case (issue #644).
            let _ = write!(
                out,
                "#place({align}, dx: -{}pt, dy: {}pt, line(length: 100% + {}pt, stroke: {}pt + {}))",
                format_f64(TEXT_COLUMN_DECORATION_OVERHANG_PT),
                format_f64(sign * dy),
                format_f64(2.0 * TEXT_COLUMN_DECORATION_OVERHANG_PT),
                format_f64(w),
                rgb(&side.color),
            );
        }
    }
}

pub(super) fn write_par_settings(out: &mut String, style: &ParagraphStyle, runs: &[Run]) {
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
        if justified_lines_take_natural_width_only(runs) {
            out.push_str("  #set par(linebreaks: \"simple\")\n");
        }
        write_east_asian_justification_limits(out, runs);
    }
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

/// State the ceiling Word puts on every expandable gap of a justified line
/// carrying the East Asian/Latin auto space.
///
/// Word distributes a line's stretch demand in three phases (measured over
/// demands from 6.2pt to 300.2pt, issue #1053): the word spaces fill to half
/// an em, then the auto spaces take the remainder, then *every* expandable
/// gap on the line sits at one common width. Typst's justifier fills the gaps
/// in proportion to their stretchability and, once that is spent, adds what
/// is left equally to every justifiable glyph — so a ceiling stated as a
/// length rather than a multiple of each gap's own width brings a 0.352em
/// word space and a 0.25em auto space to the same half em together, and
/// Word's common width falls out of the equal remainder. Measured on the
/// #1193 line: 6.803pt for both kinds against Word's 6.8065pt and 6.8028pt,
/// where the rigid quarter em gave 8.70pt and 2.62pt.
///
/// The phases in between are not separable — one ratio drives every gap — so
/// an ordinary justified line, whose demand never reaches the ceiling, now
/// moves its auto spaces where Word leaves them: 0.27pt per gap on the
/// fixture's second line, against 4.18pt per gap the other way before.
///
/// `justification-limits` rejects a ratio of zero, so the smallest positive
/// one stands in for "no ratio term": at 0.0001% it adds 4 nanopoints to a
/// 10.5pt space. The `em` resolves against each glyph's own size rather than
/// the paragraph's, so a line mixing sizes caps each gap at its own half em.
/// The floor stays [`JUSTIFIED_SPACING_FLOOR`], which is calibrated on the
/// corpus and has nothing to do with this ceiling.
///
/// Scoped to the paragraphs that carry the auto space: Typst's line breaker
/// prices a line against the allowance it is given, so a document-wide
/// ceiling would move breaks in every justified Latin paragraph as well.
pub(super) fn write_east_asian_justification_limits(out: &mut String, runs: &[Run]) {
    if !runs
        .iter()
        .any(|run| run.text.contains(EAST_ASIAN_AUTO_SPACE_CHAR))
    {
        return;
    }
    let _ = writeln!(
        out,
        "  #set par(justification-limits: (spacing: (min: {JUSTIFIED_SPACING_FLOOR}, max: 0.0001% + {EAST_ASIAN_JUSTIFIED_GAP_CEILING_EM}em)))"
    );
}

/// Whether this justified paragraph fills each line only up to the width its
/// content naturally occupies, never borrowing from the word spaces to seat one
/// more token.
///
/// Typst's justified line breaker is Knuth-Plass, which prices a slightly
/// squeezed line far below a very loose one and so takes that trade whenever
/// the shrink allowance permits it. Word's does the same under its post-2013
/// engine: at `compatibilityMode 15` a native export pulls one more eojeol onto
/// the line and compresses twelve word spaces to 0.9746 of natural to seat it,
/// which is our own 0.9746 to within 0.0001pt.
///
/// Word's pre-2013 East Asian justification has no such phase. Swept over
/// eleven measures either side of the fit boundary, a native legacy export
/// seats the extra token only while it fits at natural width and refuses a
/// 0.5pt overrun — even though the same package's Latin paragraph, in the same
/// export, takes up to 2.5pt of overrun. So the switch is scoped to East Asian
/// paragraphs: it is not that legacy Word cannot compress, it is that its East
/// Asian justification does not (issue #1130).
///
/// `linebreaks: "simple"` is Typst's first-fit breaker, which is what Word's
/// is; a line only ever shrinks under it when a single unbreakable token
/// overflows the measure, where Word overflows too. Typst already uses it for
/// every unjustified paragraph, which is why left-aligned Korean text has
/// always broken where Word breaks it.
fn justified_lines_take_natural_width_only(runs: &[Run]) -> bool {
    legacy_word_justification_is_active() && has_east_asian_text(runs)
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
    default_tab_width_pt: f64,
    eojeol_wrap: EojeolWrap,
) {
    generate_runs_with_tabs_and_metrics(
        out,
        runs,
        tab_stops,
        default_tab_width_pt,
        eojeol_wrap,
        None,
    );
}

/// Emit fixed-page PowerPoint runs, giving each explicitly broken segment
/// PowerPoint's own line box when every one of them fits a physical line.
///
/// Every line owns a 1.2em box at **its own** size, so a break advances by the
/// preceding line's descent plus the following line's seat and a column of
/// same-size lines paces at exactly `1.2 x size`. Typst would otherwise pace
/// the stack by the *face's* edges, which are not PowerPoint's box.
///
/// A native PowerPoint 16 export of nine-line hard-broken columns measures the
/// same-size case at `1.2 x size` for every size probed — 6, 6.5, 8, 8.5, 9,
/// 9.2, 9.5, 9.8, 10.5, 11.5, 12.25 and 20pt — in a text box and in a table
/// cell alike, which is what retired the 10pt line-box floor this path used to
/// carry (issue #1172). The floor was fitted to one 9.5pt table cell that
/// advances 12.00pt natively, but that cell follows a *13pt* line: the same
/// export paces a nine-line 9.5pt column at 11.37pt, and the 12.00pt is the
/// taller preceding line's descent, not a floor.
///
/// The width guard is essential: an edge set on a segment that wraps would
/// affect every physical line in it, not only the one before the hard break.
/// Unmeasured containers and segments too close to the wrapping boundary keep
/// their existing, internally consistent Typst line boxes.
///
/// `reserves_trailing_letter_space` says whether the caller places a line by
/// the width PowerPoint measured, letter-space after the last glyph included;
/// every decline above then routes through [`generate_powerpoint_inline_runs`],
/// which writes that reserve on the lines a hard break ends. The per-line box
/// path never needs it: [`powerpoint_hard_break_line_fits`] declines any
/// tracked run, so a paragraph reaching the stack has no letter-space to
/// reserve.
#[allow(clippy::too_many_arguments)]
pub(super) fn generate_powerpoint_runs_with_tabs(
    out: &mut String,
    runs: &[Run],
    style: &ParagraphStyle,
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
    eojeol_wrap: EojeolWrap,
    available_measure_pt: Option<f64>,
    reserves_trailing_letter_space: bool,
) {
    let inline = |out: &mut String, lines: Option<&[PowerPointHardBreakLine]>| {
        generate_powerpoint_inline_runs(
            out,
            runs,
            lines,
            style,
            tab_stops,
            default_tab_width_pt,
            eojeol_wrap,
            reserves_trailing_letter_space,
        );
    };

    let Some(lines) = split_runs_on_hard_breaks(runs) else {
        inline(out, None);
        return;
    };
    let Some((measure_pt, top_em, bottom_em)) =
        powerpoint_hard_break_stack_settings(&lines, runs, style, available_measure_pt)
    else {
        inline(out, Some(&lines));
        return;
    };

    let mut line_sources: Vec<String> = Vec::with_capacity(lines.len());
    for line in &lines {
        let mut source: String = String::new();
        generate_runs_with_tabs(
            &mut source,
            &line.runs,
            tab_stops,
            default_tab_width_pt,
            eojeol_wrap,
        );
        line_sources.push(source);
    }

    out.push_str("#stack(dir: ttb, spacing: 0pt,\n");
    // The full-measure line box swallows the surrounding paragraph alignment,
    // so each line re-anchors by the paragraph's own: a centred table cell
    // stayed centred in the reference while the first stack build flushed it
    // to the left edge.
    let anchor: &str = match style.alignment {
        Some(Alignment::Center) => "top + center",
        Some(Alignment::Right) => "top + right",
        _ => "top + left",
    };
    for (line, source) in lines.iter().zip(&line_sources) {
        let line_size_pt: f64 = line.max_font_size_pt();
        let top_pt: f64 = top_em * line_size_pt;
        let line_height_pt: f64 = top_pt + bottom_em * line_size_pt;
        let _ = writeln!(
            out,
            "  box(width: {width}pt, height: {height}pt)[#place({anchor}, dy: {top}pt)[#text(top-edge: \"baseline\", bottom-edge: \"baseline\")[{source}]]],",
            width = format_f64(measure_pt),
            height = format_f64(line_height_pt),
            top = format_f64(top_pt),
            anchor = anchor,
        );
    }
    out.push(')');
}

pub(super) fn powerpoint_hard_breaks_use_line_stack(
    runs: &[Run],
    style: &ParagraphStyle,
    available_measure_pt: Option<f64>,
) -> bool {
    let Some(lines) = split_runs_on_hard_breaks(runs) else {
        return false;
    };
    powerpoint_hard_break_stack_settings(&lines, runs, style, available_measure_pt).is_some()
}

fn powerpoint_hard_break_stack_settings(
    lines: &[PowerPointHardBreakLine],
    runs: &[Run],
    style: &ParagraphStyle,
    available_measure_pt: Option<f64>,
) -> Option<(f64, f64, f64)> {
    let measure_pt: f64 = available_measure_pt
        .map(|measure| {
            measure
                - style.indent_left.unwrap_or(0.0)
                - style.indent_right.unwrap_or(0.0)
                - style.indent_first_line.unwrap_or(0.0).max(0.0)
        })
        .filter(|measure| *measure > 0.0)?;
    let (top_em, bottom_em): (f64, f64) = powerpoint_paragraph_line_box_em(runs, style)?;
    if !lines
        .iter()
        .all(|line| powerpoint_hard_break_line_fits(line, measure_pt))
    {
        return None;
    }
    Some((measure_pt, top_em, bottom_em))
}

/// Emit the paragraph as ordinary inline markup, giving every hard-broken line
/// but the last the letter-space PowerPoint measures after its own last glyph.
///
/// PowerPoint measures *each* line that way, not just the paragraph's final
/// one, and places a centred or right-aligned line from that width
/// ([`powerpoint_trailing_letter_space_pt`]). #1120 trailed the space once,
/// after the whole paragraph, so a paragraph broken by `<a:br/>` reserved it on
/// its last line alone: slide 13 of the #841 deck — one centred 38pt title at
/// `spc="300"` — sat half a letter-space right of a native PowerPoint 16.112
/// export on both of its visible lines (issue #1174).
///
/// The last line is deliberately absent: its space is written by the caller,
/// after the paragraph's whole markup, so that a `no_wrap` box carries it
/// inside its own width.
///
/// `lines` is the hard-break split the caller already computed, or `None` when
/// the paragraph has no break at all. Splitting is otherwise pure overhead: a
/// paragraph that needs no per-line space emits exactly the markup one pass
/// over `runs` produces.
#[allow(clippy::too_many_arguments)]
fn generate_powerpoint_inline_runs(
    out: &mut String,
    runs: &[Run],
    lines: Option<&[PowerPointHardBreakLine]>,
    style: &ParagraphStyle,
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
    eojeol_wrap: EojeolWrap,
    reserves_trailing_letter_space: bool,
) {
    if lines.is_none() {
        let run_line_metrics = PowerPointRunLineMetrics::for_mixed_declared_sizes(runs, style)
            .map(RunLineMetrics::PowerPoint);
        generate_runs_with_tabs_and_metrics(
            out,
            runs,
            tab_stops,
            default_tab_width_pt,
            eojeol_wrap,
            run_line_metrics,
        );
        return;
    }

    let broken: Option<&[PowerPointHardBreakLine]> = reserves_trailing_letter_space
        .then_some(lines)
        .flatten()
        .filter(|lines| {
            lines.split_last().is_some_and(|(_, leading)| {
                leading
                    .iter()
                    .any(|line| powerpoint_trailing_letter_space_pt(style, &line.runs).is_some())
            })
        });
    let Some(lines) = broken else {
        generate_runs_with_tabs(out, runs, tab_stops, default_tab_width_pt, eojeol_wrap);
        return;
    };

    for (index, line) in lines.iter().enumerate() {
        if index > 0 {
            out.push_str("#linebreak()");
        }
        generate_runs_with_tabs(
            out,
            &line.runs,
            tab_stops,
            default_tab_width_pt,
            eojeol_wrap,
        );
        if index + 1 < lines.len()
            && let Some(spacing) = powerpoint_trailing_letter_space_pt(style, &line.runs)
        {
            let _ = write!(out, "#h({}pt)", format_f64(spacing));
        }
    }
}

/// Whether this explicit segment is safely narrower than its physical line.
///
/// The one-point reserve keeps font-substitution and PowerPoint's 1/8pt
/// nominal-advance rounding from turning a borderline static measurement into
/// an actual wrap. Unsupported width-affecting constructs decline correction
/// instead of guessing.
fn powerpoint_hard_break_line_fits(line: &PowerPointHardBreakLine, measure_pt: f64) -> bool {
    const WRAP_RESERVE_PT: f64 = 1.0;
    let mut advance_pt: f64 = 0.0;
    for run in &line.runs {
        if run.footnote.is_some()
            || run.text.contains('\t')
            || matches!(run.style.italic, Some(true))
            || matches!(run.style.small_caps, Some(true))
            || run.style.vertical_align.is_some()
            || run.style.baseline_shift.is_some()
            || run
                .style
                .letter_spacing
                .is_some_and(|spacing| spacing != 0.0)
        {
            return false;
        }
        let source: String = if matches!(run.style.all_caps, Some(true)) {
            run.text.to_uppercase()
        } else {
            run.text.clone()
        };
        let latin_family: &str = match run.style.font_family.as_deref() {
            Some(family) => family,
            None => return false,
        };
        let east_asian_family: &str = run
            .style
            .east_asian_font_family
            .as_deref()
            .unwrap_or(latin_family);
        let font_size_pt: f64 = match run.style.font_size {
            Some(size) => size,
            None => return false,
        };
        let is_bold: bool = effective_font_weight(&run.style)
            .is_some_and(|weight| weight != "regular" && weight != "light");
        let mut segment: String = String::new();
        let mut segment_is_east_asian: Option<bool> = None;
        for character in source.chars() {
            let is_east_asian: bool = is_cjk_like(character);
            if segment_is_east_asian != Some(is_east_asian) && !segment.is_empty() {
                let family: &str = if segment_is_east_asian == Some(true) {
                    east_asian_family
                } else {
                    latin_family
                };
                let Some(advance_em) =
                    crate::render::pdf::text_advance_em(family, is_bold, &segment)
                else {
                    return false;
                };
                advance_pt += advance_em * font_size_pt;
                segment.clear();
            }
            segment_is_east_asian = Some(is_east_asian);
            segment.push(character);
        }
        if !segment.is_empty() {
            let family: &str = if segment_is_east_asian == Some(true) {
                east_asian_family
            } else {
                latin_family
            };
            let Some(advance_em) = crate::render::pdf::text_advance_em(family, is_bold, &segment)
            else {
                return false;
            };
            advance_pt += advance_em * font_size_pt;
        }
    }
    advance_pt + WRAP_RESERVE_PT <= measure_pt
}

/// One source line delimited by a PPTX `<a:br/>` marker or a literal newline.
struct PowerPointHardBreakLine {
    runs: Vec<Run>,
    /// Style of the delimiter that opened an otherwise empty line.
    fallback_font_size_pt: Option<f64>,
}

impl PowerPointHardBreakLine {
    fn max_font_size_pt(&self) -> f64 {
        self.runs
            .iter()
            .filter(|run| !run.text.is_empty() || run.footnote.is_some())
            .filter_map(|run| run.style.font_size)
            .reduce(f64::max)
            .or(self.fallback_font_size_pt)
            .unwrap_or(12.0)
    }
}

/// Split without emitting the delimiter, preserving each visible fragment's
/// run styling. A delimiter's style is only a blank-line fallback: when the
/// next line has text, its own runs set its size (as PowerPoint does).
fn split_runs_on_hard_breaks(runs: &[Run]) -> Option<Vec<PowerPointHardBreakLine>> {
    if !runs.iter().any(|run| {
        run.text
            .chars()
            .any(|ch| matches!(ch, '\n' | PPTX_SOFT_LINE_BREAK_CHAR))
    }) {
        return None;
    }

    let mut lines: Vec<PowerPointHardBreakLine> = vec![PowerPointHardBreakLine {
        runs: Vec::new(),
        fallback_font_size_pt: None,
    }];
    for run in runs {
        let mut segment_start: usize = 0;
        for (offset, ch) in run.text.char_indices() {
            if !matches!(ch, '\n' | PPTX_SOFT_LINE_BREAK_CHAR) {
                continue;
            }
            if segment_start < offset {
                lines
                    .last_mut()
                    .expect("a line always exists")
                    .runs
                    .push(Run {
                        text: run.text[segment_start..offset].to_string(),
                        style: run.style.clone(),
                        href: run.href.clone(),
                        footnote: run.footnote.clone(),
                    });
            }
            lines.push(PowerPointHardBreakLine {
                runs: Vec::new(),
                fallback_font_size_pt: run.style.font_size,
            });
            segment_start = offset + ch.len_utf8();
        }
        if segment_start < run.text.len() {
            lines
                .last_mut()
                .expect("a line always exists")
                .runs
                .push(Run {
                    text: run.text[segment_start..].to_string(),
                    style: run.style.clone(),
                    href: run.href.clone(),
                    footnote: run.footnote.clone(),
                });
        } else if segment_start == 0 {
            lines
                .last_mut()
                .expect("a line always exists")
                .runs
                .push(run.clone());
        }
    }
    Some(lines)
}

fn generate_word_runs_with_tabs(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
    eojeol_wrap: EojeolWrap,
    style: &ParagraphStyle,
    line_grid_pitch: Option<f64>,
) {
    let metrics: Option<WordRunLineMetrics<'_>> =
        WordRunLineMetrics::for_mixed_declared_families(runs, style, line_grid_pitch);
    generate_runs_with_tabs_and_metrics(
        out,
        runs,
        tab_stops,
        default_tab_width_pt,
        eojeol_wrap,
        metrics.map(RunLineMetrics::Word),
    );
}

fn generate_runs_with_tabs_and_metrics(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
    eojeol_wrap: EojeolWrap,
    run_line_metrics: Option<RunLineMetrics<'_>>,
) {
    if !paragraph_contains_tabs(runs) {
        generate_runs_with_metrics(out, runs, eojeol_wrap, run_line_metrics);
        return;
    }

    let segments: Vec<Vec<Run>> = split_runs_on_tabs(runs);
    out.push_str("#context {\n");

    for (index, segment) in segments.iter().enumerate() {
        let _ = write!(out, "  let tab_segment_{index} = [");
        generate_runs_with_metrics(out, segment, eojeol_wrap, run_line_metrics);
        out.push_str("]\n");

        if index == 0 {
            out.push_str("  let tab_prefix_0 = tab_segment_0\n");
            continue;
        }

        write_tab_segment_bindings(out, index, segment, tab_stops, default_tab_width_pt);
    }

    let _ = writeln!(out, "  tab_prefix_{}", segments.len() - 1);
    out.push('}');
}

pub(super) fn generate_runs_with_tabs_no_wrap(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
) {
    let transformed_runs: Vec<Run> = runs
        .iter()
        .map(|run| {
            let mut transformed_run: Run = run.clone();
            if transformed_run.footnote.is_none() {
                transformed_run.text = no_wrap_text(&transformed_run.text);
            }
            transformed_run
        })
        .collect();

    // Slide text keeps PowerPoint's own breaking, which splits Korean
    // mid-word; this path additionally forbids every break outright.
    generate_runs_with_tabs(
        out,
        &transformed_runs,
        tab_stops,
        default_tab_width_pt,
        EojeolWrap::Syllable,
    );
}

/// Emits Typst variable bindings for a non-first tab segment: measurement,
/// decimal anchor (if applicable), default remainder, advance, fill, and
/// the accumulated prefix content variable.
fn write_tab_segment_bindings(
    out: &mut String,
    index: usize,
    segment: &[Run],
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
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
        // Measured for its width only, which a frame does not change, so the
        // anchor stays the plain emission.
        generate_runs(out, &anchor_runs, EojeolWrap::Syllable);
        out.push_str("]\n");
        let _ = writeln!(
            out,
            "  let tab_decimal_width_{index} = measure(tab_decimal_anchor_{index}).width"
        );
    }

    let _ = writeln!(
        out,
        "  let tab_default_remainder_{index} = calc.rem-euclid(tab_prefix_width_{index}.abs.pt(), {})",
        format_f64(default_tab_width_pt)
    );
    let _ = writeln!(
        out,
        "  let tab_advance_{index} = {}",
        build_tab_advance_expr(index, segment, tab_stops, default_tab_width_pt)
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

/// Whether a run list keeps each Hangul eojeol — a space-delimited Korean
/// word — whole when a line has to break (issue #626).
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(super) enum EojeolWrap {
    /// Typst's UAX #14 default, which permits a break between any two Hangul
    /// syllable blocks. PowerPoint breaks Korean mid-word, and our output
    /// already matches it there, so slides and sheets stay on this — as does
    /// any Word paragraph whose effective `w:wordWrap` is off.
    #[default]
    Syllable,
    /// Emit each eojeol inside an inline `#box`. A frame is a single object to
    /// UAX #14, so no break opportunity survives inside it. Typst 0.14 offers
    /// no other lever for *removing* one: `text(lang: "ko")`,
    /// `par(linebreaks:)` and `text(costs:)` were each measured to leave the
    /// breakpoints untouched, because typst-layout builds its ICU4X segmenters
    /// with default options and never consults `Lang::KOREAN`.
    /// `par(linebreaks:)` only picks among the opportunities that exist, which
    /// is what issue #1130 uses it for. The repo already uses the same
    /// mechanism in the opposite direction, where `#box[]` *creates* a
    /// contingent break for PowerPoint kinsoku (issue #438).
    ///
    /// A no-break marker between the syllables — U+2060 WORD JOINER, the
    /// obvious alternative — does suppress the breaks, but it lands in the PDF
    /// text layer and makes the text unsearchable. The no-wrap path emitted
    /// exactly that until issue #664; the marker turned out to be redundant
    /// there, because the paragraph was already inside a box. A frame leaves
    /// the text layer untouched, so it is the mechanism here too.
    ///
    /// U+200B, the kinsoku break marker, is stripped before emission and
    /// never reaches the text layer. U+E001 is stripped too, but what
    /// replaces it does not: since issue #1193 the auto space is drawn with a
    /// U+00A0, so a boundary does put a character in the text stream. That is
    /// not the #664 hazard, because the gap already extracted as a space
    /// without one — `pdftotext` reads the fixture's justified line
    /// identically from Word's own export, from our output before the change
    /// and from our output after it, and neither `pdftotext` nor `mutool`
    /// reports a U+00A0 in any of the three.
    ///
    /// `line_box_em` is the paragraph's fixed `(top-edge, bottom-edge)` when
    /// it declares them, so the frame can restore them; see
    /// [`write_eojeol_frame_open`]. `measure_pt` is the width one line of the
    /// paragraph has, which bounds how wide a token may be and still be
    /// framed; see [`is_framed_eojeol`].
    Eojeol {
        line_box_em: Option<(f64, f64)>,
        measure_pt: Option<f64>,
    },
}

/// The longest token still framed when its width cannot be measured, in
/// characters.
///
/// A token wider than the line cannot break inside its frame, so it starts a
/// line of its own and then overflows it — one line more than Word spends,
/// plus ink outside the column. [`is_framed_eojeol`] therefore compares the
/// token's measured advance against the paragraph's own measure. This cap is
/// only the fallback for when one of the two is unknown: on `wasm32`, where
/// [`text_advance_em`](crate::render::pdf::text_advance_em) always returns
/// `None`, for a run that names no family or size, and for a container whose
/// measure did not reach codegen. An eojeol is a stem plus its particles and
/// rarely reaches ten syllables, so twenty is a generous ceiling.
const MAX_UNMEASURED_EOJEOL_CHARS: usize = 20;

/// How a paragraph breaks its Hangul lines.
///
/// Alignment is not part of the decision. A one-factor `w:jc` probe measured
/// Word breaking the same Korean sentence at the same two eojeol boundaries
/// for `left`, `both`, `center` and `right` when the package defines a default
/// paragraph style, and at the same syllable boundary for all four when it
/// defines none — the #732 style rule either way. The justified line is
/// stretched 55.12pt to the measure rather than pulled tighter by borrowing a
/// syllable, so keeping the eojeol whole is a deliberate choice Word makes
/// under justification too. The contract fixture that once read as
/// "justification breaks syllables" defines no paragraph style, so its
/// paragraphs arrive here already carrying `word_wrap == Some(false)`
/// (issue #1084).
///
/// `container_measure_pt` is the width the enclosing container gives a line —
/// the page's text width, a table column, a text box — before this
/// paragraph's own indents are taken off it.
pub(super) fn paragraph_eojeol_wrap(
    breaks_hangul_at_eojeol: bool,
    style: &ParagraphStyle,
    line_box_em: Option<(f64, f64)>,
    container_measure_pt: Option<f64>,
) -> EojeolWrap {
    // `w:wordWrap w:val="0"` asks for character-level breaking outright, and
    // it wins over the style chain, so it is checked before anything the
    // paragraph inherits (issue #730).
    if style.word_wrap == Some(false) {
        return EojeolWrap::Syllable;
    }
    if !breaks_hangul_at_eojeol {
        return EojeolWrap::Syllable;
    }
    // A hanging first line (`indent_first_line < 0`) is wider than the rest,
    // so the continuation lines — the ones a frame can be pushed onto — are
    // the binding measure and the negative first-line indent is ignored.
    let measure_pt: Option<f64> = container_measure_pt
        .map(|measure| {
            measure - style.indent_left.unwrap_or(0.0) - style.indent_right.unwrap_or(0.0)
        })
        .filter(|measure| *measure > 0.0);
    EojeolWrap::Eojeol {
        line_box_em,
        measure_pt,
    }
}

pub(super) fn generate_runs(out: &mut String, runs: &[Run], eojeol_wrap: EojeolWrap) {
    generate_runs_with_metrics(out, runs, eojeol_wrap, None);
}

fn generate_runs_with_metrics(
    out: &mut String,
    runs: &[Run],
    eojeol_wrap: EojeolWrap,
    run_line_metrics: Option<RunLineMetrics<'_>>,
) {
    let EojeolWrap::Eojeol {
        line_box_em,
        measure_pt,
    } = eojeol_wrap
    else {
        for (index, run) in runs.iter().enumerate() {
            generate_run_at_with_metrics(out, run, index == 0, run_line_metrics);
        }
        return;
    };

    // Everything between two frames is coalesced and spliced back into whole
    // runs before it is emitted, so a paragraph in which no eojeol is framed —
    // every Latin one, and every Korean one whose words are all single
    // syllables — keeps byte-identical markup.
    let mut units: Vec<(bool, Vec<EojeolPiece>)> = Vec::new();
    for token in split_runs_into_eojeol_tokens(runs) {
        match (is_framed_eojeol(&token, measure_pt), units.last_mut()) {
            (false, Some((false, unframed))) => unframed.extend(token),
            (framed, _) => units.push((framed, token)),
        }
    }

    for (framed, pieces) in &units {
        // A framed eojeol restores the paragraph's edges inside itself, so a
        // synthetic-oblique box within it has to claim the same descent or
        // the frame's baseline shift over-corrects by exactly that much.
        let seat_bottom_pt: Option<f64> = if *framed {
            write_eojeol_frame_open(out, pieces, line_box_em)
        } else {
            None
        };
        write_eojeol_pieces(out, pieces, seat_bottom_pt, run_line_metrics);
        if *framed {
            write_eojeol_frame_close(out, line_box_em);
        }
    }
}

/// A slice of one run, tagged with the run it was cut from.
///
/// The tag is what lets [`write_eojeol_pieces`] splice neighbouring slices
/// back together: `escape_typst` reads its whole input — a run of spaces
/// becomes a code-mode string, a leading `<digits>.` an escaped enum marker —
/// so re-joining pieces of *different* runs could change the markup where
/// concatenating pieces of the same run never can.
struct EojeolPiece {
    run_index: usize,
    run: Run,
}

/// Emits pieces, re-joining every neighbouring pair cut from the same run.
fn write_eojeol_pieces(
    out: &mut String,
    pieces: &[EojeolPiece],
    seat_bottom_pt: Option<f64>,
    run_line_metrics: Option<RunLineMetrics<'_>>,
) {
    let mut pending: Option<(usize, Run)> = None;
    for piece in pieces {
        match pending {
            Some((run_index, ref mut previous)) if run_index == piece.run_index => {
                previous.text.push_str(&piece.run.text);
            }
            _ => {
                if let Some((_, previous)) = pending.take() {
                    generate_run_seated_with_metrics(
                        out,
                        &previous,
                        false,
                        seat_bottom_pt,
                        run_line_metrics,
                    );
                }
                pending = Some((piece.run_index, piece.run.clone()));
            }
        }
    }
    if let Some((_, previous)) = pending {
        generate_run_seated_with_metrics(out, &previous, false, seat_bottom_pt, run_line_metrics);
    }
}

/// The characters that close an eojeol, so a frame never spans one.
///
/// Most are the characters a Word line may end at. A tab is among them, so a
/// frame can never straddle a [`split_runs_on_tabs`] segment. A no-break
/// space cannot host a break at all, but it still separates words, and
/// treating it as a boundary keeps a whole run of them out of one token.
///
/// The East Asian auto space is the one entry that is not a break
/// opportunity. It closes a token because a frame is laid out at its natural
/// width, which freezes every gap inside it: left within the frame the auto
/// space could take no part in the line's justification, and the whole of a
/// stretched line's demand landed in its word spaces (issue #1193). Outside
/// it the boundary is still unbreakable, because the space
/// [`east_asian_auto_space`] emits there is U+00A0.
fn is_eojeol_delimiter(ch: char) -> bool {
    matches!(
        ch,
        ' ' | '\u{00A0}' | '\t' | '\n' | PPTX_SOFT_LINE_BREAK_CHAR | EAST_ASIAN_AUTO_SPACE_CHAR
    )
}

/// Hangul: a precomposed syllable block, a conjoining jamo, or a
/// compatibility jamo. Han and kana are deliberately absent — Chinese and
/// Japanese really do break between characters, and framing them would
/// destroy correct output.
fn is_hangul(ch: char) -> bool {
    matches!(ch as u32, 0x1100..=0x11FF | 0x3130..=0x318F | 0xAC00..=0xD7A3)
}

/// Splits a run list into the tokens Word may break between: maximal stretches
/// of delimiter-free text, each possibly spanning several runs, alternating
/// with the delimiters themselves.
///
/// Grouping happens here, structurally, rather than through a marker pair in
/// the text: [`extract_decimal_anchor_runs`] slices a run sub-list for a
/// decimal tab stop and would sever an open/close pair, emitting unbalanced
/// markup. Spanning runs matters because a bold or coloured fragment inside a
/// word would otherwise leave a frame boundary — itself a break opportunity —
/// in the middle of the word.
///
/// A footnote run is a token of its own: the reference is an anchor, not part
/// of any word.
fn split_runs_into_eojeol_tokens(runs: &[Run]) -> Vec<Vec<EojeolPiece>> {
    let mut tokens: Vec<Vec<EojeolPiece>> = Vec::new();
    let mut token: Vec<EojeolPiece> = Vec::new();

    for (run_index, run) in runs.iter().enumerate() {
        if run.footnote.is_some() {
            if !token.is_empty() {
                tokens.push(std::mem::take(&mut token));
            }
            tokens.push(vec![EojeolPiece {
                run_index,
                run: run.clone(),
            }]);
            continue;
        }
        // An empty run still emits its wrappers, so it must survive the split.
        if run.text.is_empty() {
            token.push(EojeolPiece {
                run_index,
                run: run.clone(),
            });
            continue;
        }

        let mut piece_start: usize = 0;
        let mut piece_is_delimiter: bool = is_eojeol_delimiter(
            run.text
                .chars()
                .next()
                .expect("a non-empty run has a first char"),
        );
        for (offset, ch) in run.text.char_indices() {
            let ch_is_delimiter: bool = is_eojeol_delimiter(ch);
            if ch_is_delimiter == piece_is_delimiter {
                continue;
            }
            push_eojeol_piece(
                &mut tokens,
                &mut token,
                run_index,
                run,
                &run.text[piece_start..offset],
                piece_is_delimiter,
            );
            piece_start = offset;
            piece_is_delimiter = ch_is_delimiter;
        }
        push_eojeol_piece(
            &mut tokens,
            &mut token,
            run_index,
            run,
            &run.text[piece_start..],
            piece_is_delimiter,
        );
    }

    if !token.is_empty() {
        tokens.push(token);
    }
    tokens
}

fn push_eojeol_piece(
    tokens: &mut Vec<Vec<EojeolPiece>>,
    token: &mut Vec<EojeolPiece>,
    run_index: usize,
    run: &Run,
    text: &str,
    is_delimiter: bool,
) {
    let piece: EojeolPiece = EojeolPiece {
        run_index,
        run: Run {
            text: text.to_string(),
            style: run.style.clone(),
            href: run.href.clone(),
            footnote: None,
        },
    };
    if !is_delimiter {
        token.push(piece);
        return;
    }
    if !token.is_empty() {
        tokens.push(std::mem::take(token));
    }
    tokens.push(vec![piece]);
}

/// Whether a token is an eojeol Word would keep whole: it carries Hangul, it
/// is long enough for a break to fall inside it, and it is narrow enough that
/// a line can hold it.
///
/// `measure_pt` is the width one line of the paragraph has. A frame is opaque
/// to line breaking, so a token wider than that would be pushed onto a line of
/// its own and still overflow it — one line more than Word spends, with ink
/// outside the column. Word itself breaks such a token at character level, so
/// this returns `false` and the token keeps the engine's syllable breaking.
fn is_framed_eojeol(token: &[EojeolPiece], measure_pt: Option<f64>) -> bool {
    if token.iter().any(|piece| piece.run.footnote.is_some()) {
        return false;
    }
    // Letter spacing survives a frame boundary by a rule this generator cannot
    // predict. Measured on typst 0.14 at `tracking: 0.7pt`: framing the four
    // Korean words of a 13pt centred heading made it 1.4pt *narrower*, while
    // framing the four words of a 9pt one made it 3.0pt *wider* — the shaper
    // does not simply trim one step per item. A tracked run is decorative
    // display text, short enough that it does not wrap and so has nothing to
    // gain here, so it keeps today's emission until the rule is measured.
    if token
        .iter()
        .any(|piece| piece.run.style.letter_spacing.is_some_and(|s| s != 0.0))
    {
        return false;
    }
    let mut visible_chars: usize = 0;
    let mut has_hangul: bool = false;
    for ch in token.iter().flat_map(|piece| piece.run.text.chars()) {
        // The in-text markers stand for spacing and break opportunities, not
        // glyphs, so they cannot make a one-syllable token breakable.
        if matches!(ch, EAST_ASIAN_AUTO_SPACE_CHAR | HANGUL_KINSOKU_BREAK_CHAR) {
            continue;
        }
        has_hangul |= is_hangul(ch);
        visible_chars += 1;
    }
    if !has_hangul || visible_chars < 2 {
        return false;
    }
    match (measure_pt, eojeol_advance_pt(token)) {
        (Some(measure), Some(advance)) => advance <= measure,
        // Either the container's measure or the token's advance is unknown;
        // fall back to the character ceiling, which at least keeps a
        // pathologically long token out of a frame.
        _ => visible_chars <= MAX_UNMEASURED_EOJEOL_CHARS,
    }
}

/// The advance a token takes on a line, in points, measured with the same
/// machinery the auto-layout column widths use (issue #624): each piece's
/// resolved family — the `w:eastAsia` face for East Asian codepoints — its
/// weight, and its own size.
///
/// `None` when a piece names no family or no size, or when a glyph is missing
/// from the resolved face; on `wasm32`
/// [`text_advance_em`](crate::render::pdf::text_advance_em) is always `None`,
/// so the whole guard degrades to its character ceiling there.
///
/// The in-text markers are skipped: they stand for a `#h()` the shaper never
/// sees as a glyph. That under-counts an auto-space boundary by a quarter em,
/// which only matters for a token already within a quarter em of the measure.
fn eojeol_advance_pt(token: &[EojeolPiece]) -> Option<f64> {
    let mut advance_pt: f64 = 0.0;
    for piece in token {
        let latin_family: &str = piece.run.style.font_family.as_deref()?;
        let east_asian_family: &str = piece
            .run
            .style
            .east_asian_font_family
            .as_deref()
            .unwrap_or(latin_family);
        let font_size_pt: f64 = piece.run.style.font_size?;
        let is_bold: bool = effective_font_weight(&piece.run.style)
            .is_some_and(|weight| weight != "regular" && weight != "light");
        // One `text_advance_em` call per maximal same-face segment: the call
        // takes a global face-cache lock, so a per-character loop would be
        // needlessly hot on long Korean paragraphs.
        let mut segment: String = String::new();
        let mut segment_is_east_asian: Option<bool> = None;
        for character in piece.run.text.chars() {
            if matches!(
                character,
                EAST_ASIAN_AUTO_SPACE_CHAR | HANGUL_KINSOKU_BREAK_CHAR
            ) {
                continue;
            }
            let is_east_asian: bool = is_cjk_like(character);
            if segment_is_east_asian != Some(is_east_asian) && !segment.is_empty() {
                let family: &str = if segment_is_east_asian == Some(true) {
                    east_asian_family
                } else {
                    latin_family
                };
                advance_pt +=
                    crate::render::pdf::text_advance_em(family, is_bold, &segment)? * font_size_pt;
                segment.clear();
            }
            segment_is_east_asian = Some(is_east_asian);
            segment.push(character);
        }
        if !segment.is_empty() {
            let family: &str = if segment_is_east_asian == Some(true) {
                east_asian_family
            } else {
                latin_family
            };
            advance_pt +=
                crate::render::pdf::text_advance_em(family, is_bold, &segment)? * font_size_pt;
        }
    }
    Some(advance_pt)
}

/// Opens an eojeol's frame.
///
/// Under Word's fixed line box (issues #354, #508) a bare `#box` seats its
/// baseline on its own *bottom* edge, which would drop the framed text by the
/// descent while the spaces around it stayed put. The frame therefore restores
/// the paragraph's edges inside itself and shifts its baseline back up by the
/// descent.
///
/// Those edges are re-emitted in points rather than the `em` the paragraph
/// declares: an `em` resolves against each run's own size, so a size change
/// inside one eojeol would leave the frame's height and its baseline shift
/// disagreeing. Resolving them at the token's own largest size reproduces
/// exactly what the same text contributes to the line unframed.
fn write_eojeol_frame_open(
    out: &mut String,
    token: &[EojeolPiece],
    line_box_em: Option<(f64, f64)>,
) -> Option<f64> {
    let Some((top_em, bottom_em)) = line_box_em else {
        out.push_str("#box[");
        return None;
    };
    let font_size_pt: f64 =
        largest_font_size_pt(token.iter().filter_map(|piece| piece.run.style.font_size));
    let top_pt: f64 = top_em * font_size_pt;
    let bottom_pt: f64 = bottom_em * font_size_pt;
    let _ = write!(
        out,
        "#box(baseline: {}pt)[#text(top-edge: {}pt, bottom-edge: -{}pt)[",
        format_f64(bottom_pt),
        format_f64(top_pt),
        format_f64(bottom_pt)
    );
    Some(bottom_pt)
}

fn write_eojeol_frame_close(out: &mut String, line_box_em: Option<(f64, f64)>) {
    out.push_str(if line_box_em.is_some() { "]]" } else { "]" });
}

/// Strip the kinsoku break markers from a run bound for a no-wrap box.
///
/// The paragraph is already emitted inside a `#box[...]`, and a box is a
/// single object to UAX #14, so no break opportunity survives inside it. This
/// therefore has nothing to add to break suppression and only has to keep the
/// zero-width kinsoku marker out of the text layer.
///
/// It used to also replace every space with U+00A0 and insert a U+2060 WORD
/// JOINER between adjacent characters. Those were redundant with the enclosing
/// box, and they reached the PDF text layer: slide text came out as
/// `p<WJ>a<WJ>n<WJ>i<WJ>c`, which no search for `panic` can match (issue #664).
fn no_wrap_text(text: &str) -> String {
    text.chars()
        .filter(|ch| *ch != HANGUL_KINSOKU_BREAK_CHAR)
        .collect()
}

pub(crate) fn is_cjk_like(ch: char) -> bool {
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

fn build_tab_advance_expr(
    index: usize,
    segment: &[Run],
    tab_stops: Option<&[TabStop]>,
    default_tab_width_pt: f64,
) -> String {
    let prefix_width_var = format!("tab_prefix_width_{index}");
    let segment_width_var = format!("tab_segment_width_{index}");
    let decimal_width_var =
        extract_decimal_anchor_runs(segment).map(|_| format!("tab_decimal_width_{index}"));
    let default_expr = build_default_tab_advance_expr(index, default_tab_width_pt);

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

fn build_default_tab_advance_expr(index: usize, default_tab_width_pt: f64) -> String {
    format!(
        "if tab_default_remainder_{index} == 0 {{ {}pt }} else {{ ({} - tab_default_remainder_{index}) * 1pt }}",
        format_f64(default_tab_width_pt),
        format_f64(default_tab_width_pt)
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
    generate_run_at(out, run, false);
}

/// As [`generate_run`], but stating whether this run opens its markup line.
///
/// Typst drops a space that opens a line, so a run leading the paragraph with
/// exactly one loses it (issue #752). A run-leading space further along sits
/// between siblings, survives literally, and must stay a break opportunity so
/// the line can still wrap there — hence the flag rather than a blanket rule
/// inside [`escape_typst`].
///
/// Two of the three paths that emit runs carry the flag: the plain loop in
/// [`generate_runs`], and the segment split around PPTX break markers, where
/// only the segment starting the run can open the line.
///
/// [`write_eojeol_pieces`] deliberately does not. Each eojeol is emitted into
/// its own frame, so "first piece" there means first-in-frame rather than
/// first-in-paragraph, and threading the flag through it put a code-mode
/// space in front of every framed eojeol. A Korean paragraph opening with a
/// single space therefore still loses it; that path needs a paragraph-level
/// notion of line start, which this does not add.
pub(super) fn generate_run_at(out: &mut String, run: &Run, opens_line: bool) {
    generate_run_seated(out, run, opens_line, None);
}

fn generate_run_at_with_metrics(
    out: &mut String,
    run: &Run,
    opens_line: bool,
    run_line_metrics: Option<RunLineMetrics<'_>>,
) {
    generate_run_seated_with_metrics(out, run, opens_line, None, run_line_metrics);
}

/// As [`generate_run_at`], plus the descent the run's line box carries below
/// the baseline, in points, when the caller knows it.
///
/// Only the synthetic-oblique path reads it, and only to keep its slant box
/// claiming the same space below the baseline the unslanted text would have.
/// `None` means "not stated", and the box then ends at the baseline — exact
/// wherever nothing outside it depends on that descent, which is every path
/// but the eojeol frame's (issue #686).
fn generate_run_seated(out: &mut String, run: &Run, opens_line: bool, seat_bottom_pt: Option<f64>) {
    generate_run_seated_with_metrics(out, run, opens_line, seat_bottom_pt, None);
}

fn generate_run_seated_with_metrics(
    out: &mut String,
    run: &Run,
    opens_line: bool,
    seat_bottom_pt: Option<f64>,
    run_line_metrics: Option<RunLineMetrics<'_>>,
) {
    if let Some(ref content) = run.footnote {
        // The note's runs carry the style its `w:pStyle` and `w:rPr` resolved
        // to, so they emit through the ordinary run path rather than as a bare
        // string that would take the engine's own footnote styling (#580).
        out.push_str("#footnote[");
        // The note's own line box is the engine's, not the referring
        // paragraph's, so a frame here has no edges to restore.
        generate_runs(out, content, EojeolWrap::Syllable);
        out.push(']');
        return;
    }

    let run_line_box: Option<RunLineBox> =
        run_line_metrics.and_then(|metrics| metrics.line_box(run));
    match run_line_box {
        Some(RunLineBox::Em { top, bottom }) => {
            let _ = write!(
                out,
                "#text(top-edge: {}em, bottom-edge: -{}em)[",
                format_f64(top),
                format_f64(bottom)
            );
        }
        Some(RunLineBox::Points { top, bottom }) => {
            let _ = write!(
                out,
                "#text(top-edge: {}pt, bottom-edge: -{}pt)[",
                format_f64(top),
                format_f64(bottom)
            );
        }
        None => {}
    }

    if run.text.contains(PPTX_SOFT_LINE_BREAK_CHAR)
        || run.text.contains(HANGUL_KINSOKU_BREAK_CHAR)
        || run.text.contains(EAST_ASIAN_AUTO_SPACE_CHAR)
    {
        write_run_with_break_markers(out, run, opens_line, seat_bottom_pt);
    } else {
        write_run_segment(out, run, &run.text, opens_line, seat_bottom_pt);
    }

    if run_line_box.is_some() {
        out.push(']');
    }
}

/// Expands the PPTX in-text markers: a soft line break becomes
/// `#linebreak()`, and a kinsoku break marker (issue #438) becomes an
/// empty `#box[]`. An inline frame is a Contingent Break in UAX #14
/// (U+FFFC), so the line may end between a Hangul syllable and its
/// trailing punctuation — which LB13 otherwise forbids. LB13 still glues
/// the mark to the frame, so the two move to the next line together, and
/// the zero-size frame neither disturbs line metrics nor leaves a
/// zero-width space in the PDF text layer.
fn write_run_with_break_markers(
    out: &mut String,
    run: &Run,
    opens_line: bool,
    seat_bottom_pt: Option<f64>,
) {
    let mut segment_start: usize = 0;

    for (offset, ch) in run.text.char_indices() {
        let auto_space: String;
        let replacement: &str = match ch {
            PPTX_SOFT_LINE_BREAK_CHAR => "#linebreak()",
            HANGUL_KINSOKU_BREAK_CHAR => "#box[]",
            EAST_ASIAN_AUTO_SPACE_CHAR => {
                auto_space = east_asian_auto_space(run);
                &auto_space
            }
            _ => continue,
        };
        if segment_start < offset {
            // Only the segment that starts the run can open the line; the
            // ones after a marker are mid-line by construction (issue #752).
            write_run_segment(
                out,
                run,
                &run.text[segment_start..offset],
                opens_line && segment_start == 0,
                seat_bottom_pt,
            );
        }
        out.push_str(replacement);
        segment_start = offset + ch.len_utf8();
    }

    if segment_start < run.text.len() {
        write_run_segment(
            out,
            run,
            &run.text[segment_start..],
            opens_line && segment_start == 0,
            seat_bottom_pt,
        );
    }
}

fn write_run_segment(
    out: &mut String,
    run: &Run,
    text: &str,
    opens_line: bool,
    seat_bottom_pt: Option<f64>,
) {
    let style = &run.style;

    let needs_all_caps: bool = matches!(style.all_caps, Some(true));
    let source: String = if needs_all_caps {
        text.to_uppercase()
    } else {
        text.to_string()
    };
    // A lone space opening the line has to become a code-mode string here,
    // inside the run's own `#text(...)`, so it takes the run's font and size —
    // emitted outside the wrapper it would be set in the ambient face and come
    // out the wrong width (issue #752).
    let escaped: String = match source.strip_prefix(' ') {
        Some(rest) if opens_line && !rest.starts_with(' ') => {
            format!("#\" \";{}", escape_typst(rest))
        }
        _ => escape_typst(&source),
    };

    let wrappers: Vec<String> = collect_formatting_wrappers(run);

    for wrapper in &wrappers {
        out.push_str(wrapper);
    }

    match synthetic_oblique_units(style, &source) {
        Some(units) => {
            write_synthetic_oblique_content(out, style, &source, &units, opens_line, seat_bottom_pt)
        }
        // A tracked run keeps one shaped item: the grid path splits words and
        // spaces into separate Typst items, and Typst trims tracking at every
        // item boundary, so the inter-word gaps lost their tracking — the
        // #841 deck's footer line came out 7.93pt narrow over two spaces
        // while every intra-word advance matched (issue #1023). Tracked runs
        // are short decorative display text, the same trade the framed-eojeol
        // exemption makes, so forgoing 1/8pt snapping costs less than the
        // gaps.
        None if powerpoint_advance_grid_is_active()
            && can_snap_powerpoint_run(&source)
            && !style.letter_spacing.is_some_and(|spacing| spacing != 0.0) =>
        {
            write_powerpoint_grid_run_content(out, &source, &escaped, style)
        }
        None => write_run_content(out, &source, &escaped, style),
    }

    for _ in &wrappers {
        out.push(']');
    }
}

/// Whether this run can take the PowerPoint advance-grid treatment without
/// changing its native break rules.
///
/// PowerPoint's Latin runs break at spaces and ASCII hyphens, both modelled by
/// [`write_powerpoint_grid_run_content`]. Common Latin punctuation is safe to
/// shape inside the same words. CJK, bidi, and other Unicode scripts keep their
/// existing Typst shaping and line breaking until the helper can preserve
/// their script-specific opportunities just as precisely.
fn can_snap_powerpoint_run(text: &str) -> bool {
    !text.is_empty()
        && text.chars().all(|ch| {
            ch == ' '
                || (ch.is_ascii_graphic() && ch != '\u{7f}')
                || matches!(
                    ch,
                    '\u{2010}'..='\u{2015}' | '\u{2018}'..='\u{201f}' | '\u{2026}'
                )
        })
}

/// Write one Latin slide run on PowerPoint's 1/8pt advance grid.
///
/// Words are shaped whole and scaled only by the accumulated nominal rounding
/// delta. That retains pair kerning; a per-glyph box would lose it and
/// fragment the PDF text layer. A zero-size box restores the native
/// break opportunity after a hyphen without adding a character to extraction.
fn write_powerpoint_grid_run_content(
    out: &mut String,
    source: &str,
    escaped: &str,
    style: &TextStyle,
) {
    let wrapped = has_text_properties(style) || needs_kerning_wrapper(style, escaped, source);
    if wrapped {
        out.push_str("#text(");
        write_text_params_for_run(out, style, escaped, source);
        out.push_str(")[");
    }

    let mut token_start = 0;
    for (offset, ch) in source.char_indices() {
        match ch {
            ' ' => {
                write_powerpoint_grid_word(out, &source[token_start..offset]);
                out.push_str("#o2p-pptx-space()");
                token_start = offset + ch.len_utf8();
            }
            '-' => {
                let end = offset + ch.len_utf8();
                write_powerpoint_grid_word(out, &source[token_start..end]);
                if end < source.len() {
                    out.push_str("#box[]");
                }
                token_start = end;
            }
            _ => {}
        }
    }
    write_powerpoint_grid_word(out, &source[token_start..]);

    if wrapped {
        out.push(']');
    }
}

fn write_powerpoint_grid_word(out: &mut String, word: &str) {
    if word.is_empty() {
        return;
    }
    let _ = write!(out, "#o2p-pptx-word([{}], (", escape_typst(word));
    for glyph in word.chars() {
        let _ = write!(out, "\"{}\",", escape_typst_string(&glyph.to_string()));
    }
    out.push_str("))");
}

/// The shear Word and PowerPoint apply to a run marked italic whose resolved
/// face ships no italic member.
///
/// Measured off a native Word export of a Malgun Gothic `<w:i/>` run: its text
/// matrix reads `trm="38 0 12.91406 38"`, a slope of 12.91406/38 = 0.340
/// (issue #686). Typst has no synthetic style of its own — it selects the
/// upright face and the emphasis disappears without a warning.
///
/// Typst's `skew` takes an angle, so the slope is carried as `atan(0.34)` =
/// 18.778 degrees, whose tangent is 0.3399994 — 6.5e-7 of slope from the
/// measured value.
const SYNTHETIC_OBLIQUE_ANGLE_DEG: f64 = 18.778;

/// One stretch of a run that the synthetic-oblique path treats alike.
enum ObliqueUnit<'a> {
    /// Shaped by a family that has a real italic face; the engine handles it.
    RealItalic(&'a str),
    /// Slanted by hand, inside one atomic box.
    Slanted(&'a str),
    /// Whitespace, kept outside the boxes so a line can still break on it.
    Space(&'a str),
}

/// How `text` has to be split so every part is slanted the way its own face
/// requires, or `None` when the ordinary `style: "italic"` path covers it.
///
/// `None` is the overwhelmingly common answer — a non-italic run, or an italic
/// one on a face that has the variant — and keeps those runs emitting exactly
/// what they emitted before.
fn synthetic_oblique_units<'a>(style: &TextStyle, text: &'a str) -> Option<Vec<ObliqueUnit<'a>>> {
    if !matches!(style.italic, Some(true)) || text.is_empty() {
        return None;
    }
    let family: Option<&str> = style.font_family.as_deref();
    let east_asian: Option<&str> = style.east_asian_font_family.as_deref();
    let latin_needs_slant: bool = font_subst::needs_synthetic_oblique(
        family,
        east_asian,
        text,
        font_subst::TextScript::Latin,
    );
    // The East Asian half of a mixed run resolves to a different face than its
    // Latin half, so the two are asked separately — that split is exactly what
    // the issue measured: Calibri-Italic for the Latin, a synthesised slant
    // for the Hangul.
    let east_asian_script: font_subst::TextScript = font_subst::text_script(text);
    let east_asian_needs_slant: bool = east_asian_script != font_subst::TextScript::Latin
        && font_subst::needs_synthetic_oblique(family, east_asian, text, east_asian_script);
    if !latin_needs_slant && !east_asian_needs_slant {
        return None;
    }

    let needs_slant = |character: char| -> bool {
        if font_subst::is_east_asian_char(character) {
            east_asian_needs_slant
        } else {
            latin_needs_slant
        }
    };
    Some(split_oblique_units(text, needs_slant))
}

/// Group `text` into [`ObliqueUnit`]s.
///
/// A slant box is a single object to UAX #14, so nothing inside one can break.
/// The grouping therefore keeps every break opportunity the text already had:
/// whitespace stays outside the boxes, each East Asian character gets its own
/// box — Korean and Japanese break between characters — and a Latin word,
/// which is atomic anyway, stays whole so its kerning and ligatures survive.
fn split_oblique_units(text: &str, needs_slant: impl Fn(char) -> bool) -> Vec<ObliqueUnit<'_>> {
    fn push_unit<'a>(units: &mut Vec<ObliqueUnit<'a>>, chunk: &'a str, kind: (bool, bool)) {
        if chunk.is_empty() {
            return;
        }
        units.push(match kind {
            (true, _) => ObliqueUnit::Space(chunk),
            (false, true) => ObliqueUnit::Slanted(chunk),
            (false, false) => ObliqueUnit::RealItalic(chunk),
        });
    }

    let mut units: Vec<ObliqueUnit<'_>> = Vec::new();
    let mut start: usize = 0;
    let mut current: Option<(bool, bool)> = None; // (is_space, is_slanted)

    for (offset, character) in text.char_indices() {
        let is_space: bool = character.is_whitespace();
        let is_slanted: bool = !is_space && needs_slant(character);
        let kind: (bool, bool) = (is_space, is_slanted);
        // An East Asian character is its own break opportunity, so it always
        // starts a new box rather than joining the one before it.
        let breaks_group: bool =
            current != Some(kind) || (is_slanted && font_subst::is_east_asian_char(character));
        if breaks_group {
            if let Some(kind) = current {
                push_unit(&mut units, &text[start..offset], kind);
            }
            start = offset;
            current = Some(kind);
        }
    }
    if let Some(kind) = current {
        push_unit(&mut units, &text[start..], kind);
    }
    units
}

/// Write a run whose slant the engine cannot supply on its own.
///
/// One outer `#text(...)` carries the run's font list, size and colour — minus
/// the italic the resolved face cannot honour — so the whitespace between the
/// boxes keeps the run's own metrics, and each part inside picks the slant it
/// needs.
fn write_synthetic_oblique_content(
    out: &mut String,
    style: &TextStyle,
    source: &str,
    units: &[ObliqueUnit<'_>],
    opens_line: bool,
    seat_bottom_pt: Option<f64>,
) {
    let upright_style: TextStyle = TextStyle {
        italic: None,
        ..style.clone()
    };
    out.push_str("#text(");
    write_text_params_for_text(out, &upright_style, source);
    out.push_str(")[");

    for (index, unit) in units.iter().enumerate() {
        match unit {
            // The line-opening space has to be a code-mode string for the same
            // reason it does on the ordinary path: markup collapses a space at
            // the start of a line (issue #752).
            ObliqueUnit::Space(spaces) if index == 0 && opens_line => {
                let _ = write!(out, "#\"{spaces}\";");
            }
            ObliqueUnit::Space(spaces) => out.push_str(&escape_typst(spaces)),
            ObliqueUnit::RealItalic(chunk) => {
                let _ = write!(out, "#text(style: \"italic\")[{}]", escape_typst(chunk));
            }
            ObliqueUnit::Slanted(chunk) => {
                // Three details keep the box from moving the text it slants.
                //
                // `bottom-edge: "baseline"` makes the box end at the baseline,
                // and an inline box sits on the line by its bottom edge, so
                // the glyphs keep their seat: with the paragraph's own bottom
                // edge they came out a descent high. `origin: bottom + left`
                // then pivots the shear on that same baseline, so the glyphs
                // lean without sliding — the default centre pivot moved a 11pt
                // Malgun Gothic syllable 1.34pt left, and a descender pivot
                // moved it 0.88pt right.
                //
                // A stated seat adds the descent back as padding below the
                // box and shifts the box down by the same amount, which
                // cancels for the glyphs and leaves the box occupying exactly
                // what the unslanted text did. An eojeol frame needs that: it
                // shifts its own baseline up by the descent it expects its
                // content to carry (issue #626), and a box that ended at the
                // baseline dropped every framed Korean italic 3.97pt.
                let seat: String = match seat_bottom_pt {
                    Some(bottom_pt) => format!(
                        "inset: (bottom: {bottom}pt), baseline: {bottom}pt, ",
                        bottom = format_f64(bottom_pt)
                    ),
                    None => String::new(),
                };
                let _ = write!(
                    out,
                    "#box({seat}skew(ax: -{}deg, origin: bottom + left)[#text(bottom-edge: \"baseline\")[{}]])",
                    format_f64(SYNTHETIC_OBLIQUE_ANGLE_DEG),
                    escape_typst(chunk)
                );
            }
        }
    }
    out.push(']');
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
        wrappers.push(format!("#highlight(fill: {})[", rgb(highlight)));
    }
    if matches!(style.strikethrough, Some(true)) {
        wrappers.push("#strike[".to_string());
    }
    if matches!(style.underline, Some(true)) {
        // Word draws the rule as one filled rectangle straight through any
        // descender that crosses it. Typst's `underline` skips ink by default,
        // which broke a single run's rule into segments: on the audited offer
        // letter it emitted three pieces totalling 84.49pt where Word draws
        // one 89.28pt rectangle (issue #641).
        wrappers.push("#underline(evade: false)[".to_string());
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
fn write_run_content(out: &mut String, source: &str, escaped: &str, style: &TextStyle) {
    if has_text_properties(style) || needs_kerning_wrapper(style, escaped, source) {
        out.push_str("#text(");
        write_text_params_for_run(out, style, escaped, source);
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
        || style.baseline_shift.is_some()
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

/// Emit text parameters for content the caller cannot name.
///
/// The text matters to two of the parameters — the font fallback list and the
/// kerning decision — so a caller that cannot name it gets the answer that is
/// safe for any script: the fallback list of the family alone, and kerning
/// left on. See [`kerning_param`] for why "on" is the safe side.
pub(super) fn write_text_params(out: &mut String, style: &TextStyle) {
    write_text_params_inner(out, style, KerningText::Unknown);
}

/// As [`write_text_params`], but told what the run holds.
///
/// The font list has to answer for the script the text is written in, not only
/// for the family it names: a run can declare a face that has no glyph for its
/// own content (issues #537, #543).
pub(super) fn write_text_params_for_text(out: &mut String, style: &TextStyle, text: &str) {
    write_text_params_inner(out, style, KerningText::known(text));
}

/// As [`write_text_params_for_text`], but for a run whose markup form differs
/// from the text the engine will shape.
///
/// The escaping a run goes through adds backslashes and can lift a leading
/// space into a code-mode string, so the markup is the wrong thing to measure
/// advances on even though it still answers for the script.
fn write_text_params_for_run(out: &mut String, style: &TextStyle, escaped: &str, shaped: &str) {
    write_text_params_inner(out, style, KerningText::Known { escaped, shaped });
}

/// What the emitter knows about the text a `#text(...)` will cover.
#[derive(Clone, Copy)]
enum KerningText<'a> {
    /// The exact text, so its script decides. `escaped` is the markup form;
    /// `shaped` is what the engine lays out, and the only one whose glyphs can
    /// be measured.
    Known { escaped: &'a str, shaped: &'a str },
    /// The emission site cannot name the text — a list marker's numbering
    /// result, a header field's page number — so the safe answer stands.
    Unknown,
}

impl<'a> KerningText<'a> {
    /// Text whose markup and shaped forms are the same.
    fn known(text: &'a str) -> Self {
        KerningText::Known {
            escaped: text,
            shaped: text,
        }
    }
}

fn write_text_params_inner(out: &mut String, style: &TextStyle, kerning_text: KerningText<'_>) {
    let mut first = true;
    let text: &str = match kerning_text {
        KerningText::Known { escaped, .. } => escaped,
        KerningText::Unknown => "",
    };

    if let Some(ref family) = style.font_family {
        let font_value = match style.east_asian_font_family {
            Some(ref east_asian) if !east_asian.eq_ignore_ascii_case(family) => {
                font_subst::font_with_east_asian_fallbacks(family, east_asian, text)
            }
            _ => font_subst::font_with_fallbacks_for_text(family, text),
        };
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
        write_param(out, &mut first, &format_run_fill(color, style.color_alpha));
    }
    if let Some(spacing) = effective_letter_spacing(style, kerning_text) {
        write_param(
            out,
            &mut first,
            &format!("tracking: {}pt", format_f64(spacing)),
        );
        // Tracking and ligation are mutually exclusive: a ligature replaces
        // several glyphs with one and swallows the inter-glyph spacing the
        // tracking should have added. PowerPoint disables ligatures under
        // `a:rPr/@spc` for that reason, which is also the ordinary
        // typographic rule. Leaving `liga` on merged the `ffi` of "office2pdf"
        // into one glyph, so the tracking between those three letters was
        // never applied and the text layer extracted as "o ffi c e 2 p d f" —
        // 17 characters instead of 10, matching no search for the word. On the
        // audited deck that affected 24 of 52 occurrences (issue #684).
        //
        // Zero is not tracking. PowerPoint writes `spc="0"` routinely, so
        // keying this on `is_some()` would strip ligatures from whole decks
        // that never asked for it.
        if spacing != 0.0 {
            write_param(out, &mut first, "ligatures: false");
        }
    }
    if let Some(BaselineShiftEm(shift_em)) = style.baseline_shift {
        // Typst's text baseline parameter is positive downward, opposite to
        // DrawingML. Resolve against the effective run size when available.
        let shift: String = match style.font_size {
            Some(font_size_pt) => format!("{}pt", format_f64(-shift_em * font_size_pt)),
            None => format!("{}em", format_f64(-shift_em)),
        };
        write_param(out, &mut first, &format!("baseline: {shift}"));
    }
    if let Some(param) = kerning_param(style, kerning_text) {
        write_param(out, &mut first, &param);
    }
}

/// The `kerning:` parameter this text needs, or `None` when the format states
/// no rule and the engine's own default stands.
///
/// The threshold is resolved here rather than in the parser because it is
/// compared against the run's *effective* size, which only exists once the
/// style chain has been merged.
///
/// Every emission carries the decision explicitly rather than leaning on an
/// enclosing rule: the whole point of the RTL exemption below is that some
/// text must not take the surrounding answer, and a parameter that is merely
/// omitted takes whatever the nearest `#set text` says.
fn kerning_param(style: &TextStyle, kerning_text: KerningText<'_>) -> Option<String> {
    // A tracked run whose format answers the kerning question for itself takes
    // that answer, not the fallback below: PowerPoint reads `kern` and `spc`
    // off the same `a:rPr` and applies both, so a 38pt title under the deck's
    // `kern="1200"` is kerned *and* tracked (issue #1073).
    if !tracked_run_states_its_own_kerning(style) {
        // Otherwise a tracked run states only its own inter-glyph spacing. A
        // pair kern lands on top of that, and where the face is a substitute
        // its pairs are not the ones the document was set in — the combined
        // advance can exceed the gap a PDF text extractor reads as a word
        // break. On the deck in issue #864 that split five titles: `ANSATTE`
        // extracted as `ANSAT TE` while the glyphs rendered continuously,
        // because the T/T pair came to 0.115em against the reference's
        // 0.078em.
        //
        // The RTL exemption still wins: switching the `kern` feature off there
        // costs glyphs, which is worse than a word break.
        //
        // A spreadsheet run's grid correction is measured from the face's bare
        // `hmtx` sum, which is also what Excel accumulates: leaving pair
        // kerning on would move every glyph after a kerned pair off the grid
        // the correction just put it on (issue #1088).
        if effective_letter_spacing(style, kerning_text).is_some_and(|spacing| spacing != 0.0)
            && !rtl_shaping_exemption_is_active()
            && matches!(kerning_text, KerningText::Known { .. })
        {
            return Some("kerning: false".to_string());
        }
    }
    let pair_kerning: PairKerning = style.pair_kerning?;
    let kerns: bool = pair_kerning.applies_at(style.font_size)
        || rtl_shaping_exemption_is_active()
        || match kerning_text {
            KerningText::Known { .. } => false,
            // Unknown text may be RTL, and switching kerning off there costs
            // glyphs; switching it on costs a fraction of a point of advance.
            KerningText::Unknown => true,
        };
    Some(format!("kerning: {kerns}"))
}

/// Whether a tracked run's own format already decides pair kerning, so the
/// blanket rule of issue #864 must stand aside for it.
///
/// Two things have to hold together. The format must state a threshold at all
/// — `w:kern`, or DrawingML's `kern`, which every PowerPoint master writes on
/// its `titleStyle` — because absence leaves us guessing and the guess that
/// protects the text layer is "do not kern". And the run's own face must be
/// present, because #864's failure is specifically a *substitute's* kern pairs
/// riding on top of tracking the document sized for a different face: the
/// stated threshold says what PowerPoint would do with the real font, and says
/// nothing about the one we actually reach for.
fn tracked_run_states_its_own_kerning(style: &TextStyle) -> bool {
    style.pair_kerning.is_some()
        && style
            .font_family
            .as_deref()
            .is_some_and(font_subst::is_primary_font_available)
}

thread_local! {
    /// Whether the document being generated is one the `kern` feature may not
    /// be switched off in. See [`with_rtl_shaping_exemption`].
    static RTL_SHAPING_EXEMPTION: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    /// Whether run emission is currently inside a fixed PowerPoint page and
    /// should use the 1/8pt nominal-advance grid (issue #661).
    static POWERPOINT_ADVANCE_GRID: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    /// The printed-to-declared scale while run emission is inside a
    /// spreadsheet coordinate system. `None` off a sheet; `Some(1.0)` on an
    /// unscaled sheet or inside a drawing that is scaled as a whole.
    static SHEET_ADVANCE_GRID_SCALE: std::cell::Cell<Option<f64>> =
        const { std::cell::Cell::new(None) };
    /// Whether the document being generated is laid out by Word's pre-2013
    /// engine, whose East Asian justification never compresses a line to fit
    /// one more token. See [`with_legacy_word_justification`].
    static LEGACY_WORD_JUSTIFICATION: std::cell::Cell<bool> =
        const { std::cell::Cell::new(false) };
}

/// Run `operation` with Word's pre-2013 justification in the requested state,
/// restoring the enclosing state even when generation panics.
///
/// This is a *document*-wide switch because `compatibilityMode` is declared
/// once for the package, in `word/settings.xml`. Which paragraphs it then
/// governs is a per-paragraph question, answered where the setting is emitted;
/// see [`write_par_settings`].
pub(super) fn with_legacy_word_justification<T>(active: bool, operation: impl FnOnce() -> T) -> T {
    LEGACY_WORD_JUSTIFICATION.with(|legacy| {
        let previous: bool = legacy.replace(active);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        legacy.set(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

fn legacy_word_justification_is_active() -> bool {
    LEGACY_WORD_JUSTIFICATION.with(std::cell::Cell::get)
}

/// Run `operation` with Excel's whole-point advance grid at `print_scale`,
/// restoring the enclosing state even when generation panics.
pub(super) fn with_sheet_advance_grid<T>(
    print_scale: Option<f64>,
    operation: impl FnOnce() -> T,
) -> T {
    SHEET_ADVANCE_GRID_SCALE.with(|grid| {
        let previous: Option<f64> = grid.replace(print_scale.filter(|scale| *scale > 0.0));
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        grid.set(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

fn sheet_advance_grid_scale() -> Option<f64> {
    SHEET_ADVANCE_GRID_SCALE.with(std::cell::Cell::get)
}

/// The uniform inter-glyph spacing that puts `text` on
/// [`SHEET_ADVANCE_GRID_PT`] in sheet space, or `None` when this run is not one
/// the grid applies to. A fitted cell rounds at its declared size and scales
/// that grid onto the page (issue #1238).
///
/// Typst lays a run out on the face's exact advances, and offers no per-glyph
/// override; `tracking` is the one lever, and it is uniform. Spreading the
/// run's rounding delta over its gaps stops the 5% deficit accumulating and
/// leaves each glyph within the rounding noise: measured against the ten
/// golden-mock exports, the worst origin lands 1.35pt from the native one and
/// the median 0.13pt, where the unquantized line reached 19.5pt.
///
/// The delta is summed over the advances that *carry a gap* — every glyph but
/// the last — so the run's last origin lands exactly where Excel puts it. The
/// last glyph's own advance is deliberately left out: folding it in would make
/// the run's total width exact at the cost of pushing that whole rounding into
/// the visible gaps, which on a two-glyph cell is the entire correction in the
/// one gap there is (`OK` at Arial 10 came out 8.33pt against the export's
/// 8.00pt). [`sheet_trailing_advance_space_pt`] carries that final rounding
/// separately for a right-aligned cell, so it changes the width the line is
/// placed from without moving any visible glyph origin (issue #1233).
///
/// A one-glyph run gets `None`: Typst drops the tracking after a shaped item's
/// last glyph, so there is no gap to carry the correction.
fn sheet_advance_grid_tracking_pt(style: &TextStyle, text: &str) -> Option<f64> {
    let scale: f64 = sheet_advance_grid_scale()?;
    let size_pt: f64 = style.font_size.filter(|size| *size > 0.0)?;
    let sheet_size_pt: f64 = size_pt / scale;
    let advances_em: Vec<f64> = sheet_advance_grid_glyph_advances_em(style, text)?;
    let gap_advances_em: &[f64] = advances_em.split_last()?.1;
    if gap_advances_em.is_empty() {
        return None;
    }

    let natural_pt: f64 = gap_advances_em
        .iter()
        .map(|advance| advance * size_pt)
        .sum();
    let quantized_pt: f64 = gap_advances_em
        .iter()
        .map(|advance| {
            round_half_up_to_grid(advance * sheet_size_pt, SHEET_ADVANCE_GRID_PT) * scale
        })
        .sum();
    let gaps: f64 = gap_advances_em.len() as f64;
    let tracking_pt: f64 = (quantized_pt - natural_pt) / gaps;
    // An exact fit needs no correction, and emitting `tracking: 0pt` would
    // still cost the run its ligatures and kerning below.
    (tracking_pt != 0.0).then_some(tracking_pt)
}

/// The final rounded glyph advance Excel includes when placing a right-aligned
/// sheet line from the cell's trailing edge.
///
/// The inter-glyph tracking above stops at the last glyph so every visible
/// origin stays on Excel's whole-point sheet grid. Typst therefore measures
/// the line's final advance at its natural width, while Excel includes that
/// advance rounded to [`SHEET_ADVANCE_GRID_PT`] before applying any print
/// scale. A trailing `#h` carries only the printed difference: positive values
/// move a right-aligned line left, and the rarer negative values move it right,
/// without disturbing glyph pitch.
///
/// Native Excel-for-Mac traces across the ten business workbooks predict the
/// measured right-aligned origin residual run by run from this value (issue
/// #1233). Centred lines deliberately stay unchanged: Excel's separate
/// whole-point origin snap absorbs the half-residual in the measured corpus.
pub(super) fn sheet_trailing_advance_space_pt(style: &ParagraphStyle, runs: &[Run]) -> Option<f64> {
    let scale: f64 = sheet_advance_grid_scale()?;
    if !matches!(style.alignment, Some(Alignment::Right)) {
        return None;
    }

    // This call site can append only one reserve after the whole paragraph.
    // An explicit line or tab boundary would need a reserve at each segment's
    // own trailing edge, so decline the paragraph instead of correcting only
    // its final segment and leaving the earlier ones inconsistent.
    if runs.iter().any(|run| {
        run.text
            .chars()
            .any(|ch| matches!(ch, '\n' | '\t' | PPTX_SOFT_LINE_BREAK_CHAR))
    }) {
        return None;
    }

    let run: &Run = runs.iter().rev().find(|run| !run.text.is_empty())?;
    if matches!(run.style.small_caps, Some(true)) {
        return None;
    }
    let all_caps: String;
    let shaped: &str = if matches!(run.style.all_caps, Some(true)) {
        all_caps = run.text.to_uppercase();
        &all_caps
    } else {
        &run.text
    };
    let size_pt: f64 = run.style.font_size.filter(|size| *size > 0.0)?;
    let sheet_size_pt: f64 = size_pt / scale;
    let advances_em: Vec<f64> = sheet_advance_grid_glyph_advances_em(&run.style, shaped)?;
    let natural_pt: f64 = advances_em.last()? * size_pt;
    let rounded_pt: f64 =
        round_half_up_to_grid(advances_em.last()? * sheet_size_pt, SHEET_ADVANCE_GRID_PT) * scale;
    let space_pt: f64 = rounded_pt - natural_pt;
    (space_pt != 0.0).then_some(space_pt)
}

fn sheet_advance_grid_glyph_advances_em(style: &TextStyle, text: &str) -> Option<Vec<f64>> {
    let bold: bool = matches!(style.bold, Some(true));
    let family: &str = style.font_family.as_deref()?;
    // A Korean cell names its face in `font_family` like any other, but a run
    // that carries a separate East Asian family is measured on the face its
    // glyphs will actually come from — the Latin one has no glyph for them and
    // reports nothing.
    crate::render::pdf::glyph_advances_em(family, bold, text).or_else(|| {
        let east_asian: &str = style.east_asian_font_family.as_deref()?;
        (!east_asian.eq_ignore_ascii_case(family))
            .then(|| crate::render::pdf::glyph_advances_em(east_asian, bold, text))
            .flatten()
    })
}

/// `value` rounded to the nearest multiple of `grid`, halves away from zero —
/// the rule Excel's own column metrics take (issue #621).
fn round_half_up_to_grid(value_pt: f64, grid_pt: f64) -> f64 {
    (value_pt / grid_pt).round() * grid_pt
}

/// The inter-glyph spacing this run is set with: what the source states, or
/// the spreadsheet grid's correction where the source states nothing.
///
/// Both answers reach the emitted `tracking:` and the kerning decision through
/// this one function, so a grid-corrected run takes the same "tracking states
/// its own spacing" rule a `spc`-tracked slide run does.
fn effective_letter_spacing(style: &TextStyle, kerning_text: KerningText<'_>) -> Option<f64> {
    if let Some(spacing) = style.letter_spacing {
        return Some(spacing);
    }
    match kerning_text {
        KerningText::Known { shaped, .. } => sheet_advance_grid_tracking_pt(style, shaped),
        KerningText::Unknown => None,
    }
}

/// Run `operation` with PowerPoint advance snapping in the requested state,
/// restoring the enclosing state even when generation panics.
pub(super) fn with_powerpoint_advance_grid<T>(active: bool, operation: impl FnOnce() -> T) -> T {
    POWERPOINT_ADVANCE_GRID.with(|grid| {
        let previous = grid.replace(active);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        grid.set(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

fn powerpoint_advance_grid_is_active() -> bool {
    POWERPOINT_ADVANCE_GRID.with(std::cell::Cell::get)
}

/// Run `operation` with the RTL kerning exemption in the given state, then
/// restore whatever it was.
///
/// TODO(typst 0.14.2 mis-orders RTL glyph ranges without `kern`; report
/// upstream): shaping a right-to-left segment of two or more characters with
/// the feature disabled walks `infos` backwards from index 0 and hands the
/// first glyph the whole segment's text range, leaving the next glyph an
/// inverted one. A debug build trips `assert_glyph_ranges_in_order` in
/// `typst-layout`'s `inline/shaping.rs`; a release build keeps the broken
/// ranges, which are what krilla writes `ActualText` from — krilla 0.6.0 then
/// panics with "byte range starts at 3 but ends at 0", which is how
/// FDO76312.docx failed the bulk gate. Measured on Arabic and Hebrew at two
/// characters and up; a single character, Latin, Hangul, Han, Thai and
/// Devanagari all shape correctly, and every one of them is fine with the
/// feature left on. Word's own kerning rule is therefore honoured everywhere
/// except in documents that shape right-to-left, where it would cost correct
/// text to gain at most a fraction of a point of advance. The defect is
/// upstream's, not ours — per the reference-project rule it belongs in a typst
/// issue with this reproduction, and the exemption here should be removed once
/// a release carries the fix.
///
/// The exemption is a *document*-wide switch rather than a per-run one because
/// bidi reordering is decided over a whole shaped paragraph: the run that
/// loses its glyph order need carry no right-to-left codepoint at all. It is
/// enough that a sibling run does (`مرحبا` between two Latin runs), or that
/// the paragraph's base direction is right-to-left and the run holds two
/// neutral characters — `w:bidi` plus a row of full stops, which is exactly
/// FDO76312.docx. Narrower scopes have to model which emissions end up in one
/// shaped paragraph; this one does not, so no emission site can bypass it.
pub(super) fn with_rtl_shaping_exemption<T>(active: bool, operation: impl FnOnce() -> T) -> T {
    RTL_SHAPING_EXEMPTION.with(|exemption| {
        let previous: bool = exemption.replace(active);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        exemption.set(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

fn rtl_shaping_exemption_is_active() -> bool {
    RTL_SHAPING_EXEMPTION.with(std::cell::Cell::get)
}

/// Whether generated Typst markup will be shaped right-to-left anywhere.
///
/// This reads the emitted source rather than the IR it came from, because the
/// source is what the engine shapes: no IR shape, and no future emission site,
/// can carry a right-to-left segment past it. Both routes to one are visible
/// there — a `dir: rtl` the emitter wrote for a paragraph whose base direction
/// is right-to-left, and the strong right-to-left codepoints themselves, which
/// `escape_typst` passes through unchanged.
///
/// Text that merely *reads* `dir: rtl` costs the document its kerning rule and
/// nothing else, which is the side to be wrong on.
pub(super) fn source_shapes_right_to_left(source: &str) -> bool {
    // The strong right-to-left blocks, as `xlsx_cells::strong_direction` reads
    // them: Hebrew, Arabic and its supplements, Syriac, Thaana, and the Arabic
    // presentation forms.
    source
        .chars()
        .any(|ch| matches!(ch as u32, 0x0590..=0x08FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF))
        || source.contains("dir: rtl")
}

/// Whether a run needs a `#text(...)` of its own purely to state kerning.
///
/// A run that states no other text property is emitted bare, and bare text
/// takes the engine's default — which kerns. That is the wrong answer for the
/// body text of a document Word does not kern, so the decision is stated on
/// the run rather than document-wide: a document-wide `kerning: false` would
/// also reach the emission sites that cannot name their text.
fn needs_kerning_wrapper(style: &TextStyle, escaped: &str, shaped: &str) -> bool {
    !escaped.is_empty()
        && kerning_param(style, KerningText::Known { escaped, shaped })
            .is_some_and(|param| param.ends_with("false"))
}

pub(super) fn write_param(out: &mut String, first: &mut bool, param: &str) {
    if !*first {
        out.push_str(", ");
    }
    out.push_str(param);
    *first = false;
}

pub(super) fn format_color(color: &Color) -> String {
    format!("fill: {}", rgb(color))
}

/// A run's fill, composited at the opacity its colour declares.
///
/// PowerPoint draws a run whose `a:solidFill` colour carries `<a:alpha>` at
/// exactly that fraction of the backdrop, so the ink has to reach the page with
/// its alpha channel intact rather than as the flattened base colour
/// (issue #1121).
fn format_run_fill(color: &Color, alpha: Option<f64>) -> String {
    match alpha {
        Some(alpha) => format!(
            "fill: {}",
            rgb_with_alpha(color, (alpha.clamp(0.0, 1.0) * 255.0).round() as u8)
        ),
        None => format_color(color),
    }
}

/// The char index Typst reads a *line-leading* markup marker at.
///
/// Typst recognises those markers through one leading space, so the scan
/// steps over one. Which text lands at the start of an escaping unit is the
/// generator's choice, not the document's: a run is cut at every tab, at
/// every in-text marker, and since #626 at every eojeol boundary, so a
/// paragraph's ` + ` or ` = ` reaches [`escape_typst`] as a unit of its own.
///
/// Exactly one space, because that is the only leading whitespace
/// [`escape_typst`] emits as markup. A run of two or more — and any run after
/// a hard break — leaves as a code-mode string, which cannot open a marker.
/// Measured on typst: `[ 2026. 7. 17.]`, `[ + x]` and `[ = x]` become an
/// enumeration, a list item and a heading; `[#"  ";+ x]`, `[#"  ";= x]` and a
/// leading U+00A0 do not.
fn line_leading_markup_index(text: &str) -> usize {
    usize::from(text.starts_with(' ') && !text[1..].starts_with(' '))
}

/// Whether `text` opens with a Typst line-leading marker whose first
/// character must be escaped to neutralise it.
///
/// The full set of Typst markup that is only meaningful at a line start:
///
/// | Marker | Handling |
/// | --- | --- |
/// | `- ` bullet list | here, and also escaped everywhere else (`--` ligates) |
/// | `+ ` numbered list | here — `+` is otherwise a literal |
/// | `= ` heading (any run of `=`) | here — `=` is otherwise a literal |
/// | `/ ` term list | already escaped unconditionally below |
/// | `<digits>. ` enumeration | [`enum_marker_dot`], which escapes the dot |
///
/// Every other Typst shorthand (`#`, `*`, `_`, `` ` ``, `$`, `<`, `>`, `@`,
/// `~`, `\`, `[`, `]`, `{`, `}`, `"`, `'`) is markup wherever it appears and
/// is escaped unconditionally, so a line start needs no extra rule for it.
///
/// A marker also needs trailing whitespace to be one: `[ =x]` and `[+]` stay
/// literal. Escaping only the *first* character of a `==`-style run is
/// enough — measured on typst, `[ \== ]` renders ` == ` — because what
/// remains no longer starts the line.
fn opens_line_leading_marker(text: &str) -> bool {
    match text.chars().next() {
        // A one-byte marker char, so the byte slice is safe.
        Some('-' | '+') => text[1..].chars().next().is_some_and(char::is_whitespace),
        Some('=') => text
            .trim_start_matches('=')
            .chars()
            .next()
            .is_some_and(char::is_whitespace),
        _ => false,
    }
}

pub(super) fn escape_typst(text: &str) -> String {
    let normalized_text: String = text.nfc().collect();
    let leading_space: usize = line_leading_markup_index(&normalized_text);
    let after_space: &str = &normalized_text[leading_space..];

    // A leading `-`/`+` bullet or `=` heading run would be re-typeset as that
    // marker, deleting the character from the page: ` + ` between two Korean
    // eojeol became an enumeration item and ` = ` an empty heading (#626).
    let line_leading_marker: Option<usize> =
        opens_line_leading_marker(after_space).then_some(leading_space);

    // A leading "<digits>. " would be re-typeset as a Typst numbered-list
    // marker (e.g. "2026. 07. 17." became "2026. 7. 17."); escape its dot.
    // `"시행일자: 2026. 7. 17."` reached this function as `" 2026. 7. 17."`
    // once #626 cut the run at the eojeol boundary, and Typst put the date on
    // an enumeration line of its own.
    let enum_marker_dot: Option<usize> = {
        let digit_count = after_space
            .chars()
            .take_while(|c| c.is_ascii_digit())
            .count();
        let rest = &after_space[digit_count..];
        if digit_count > 0 && (rest.starts_with(". ") || rest == ".") {
            Some(leading_space + digit_count)
        } else {
            None
        }
    };

    let mut result = String::with_capacity(normalized_text.len());
    let mut chars = normalized_text.chars().peekable();
    let mut char_index: usize = 0;

    let mut after_linebreak = false;
    while let Some(ch) = chars.next() {
        let should_escape_list_prefix: bool = line_leading_marker == Some(char_index);

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
        char_index += 1;
    }
    result
}

/// The Typst label a heading's contents-entry marker carries.
pub(super) const TOC_ENTRY_LABEL: &str = "o2p-toc";

/// A heading's text with no markup, for the contents entry that points at it.
fn paragraph_plain_text(runs: &[Run]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

/// Escape a Rust string for use inside a Typst double-quoted string literal.
fn escape_typst_string(text: &str) -> String {
    text.replace('\\', "\\\\").replace('"', "\\\"")
}

/// The family the heading's first run names, for the contents entry's own
/// fallback chain — a Korean entry needs the Korean face even though the entry
/// is laid out at body size (issue #610).
fn first_run_family(runs: &[Run]) -> Option<&str> {
    runs.iter().find_map(|run| run.style.font_family.as_deref())
}
