use super::*;

pub(super) struct MeasuredFooters {
    values: [Option<String>; 2],
}

pub(super) fn needs_measurement(page: &FlowPage) -> bool {
    [page.footer.as_ref(), page.first_footer.as_ref()]
        .into_iter()
        .flatten()
        .any(|story| story.distance_from_edge.is_some() && hf_has_flow_content(story))
}

/// Measure the same wrapped content that is painted, rather than counting lines.
/// OOXML pgMar@bottom >= 0 reserves its full extent; a negative value fixes the
/// body's boundary at abs(bottom), explicitly allowing footer overlap.
pub(super) fn write_measurements(
    out: &mut String,
    page: &FlowPage,
    size: &PageSize,
    ctx: &mut GenCtx,
) -> Option<MeasuredFooters> {
    if !needs_measurement(page) {
        return None;
    }
    let width = format_f64((size.width - page.margins.left - page.margins.right).max(0.0));
    let base = format!("{}pt", format_f64(page.margins.bottom.abs()));
    let mut margins = [base.clone(), base.clone()];
    let mut values = [None, None];
    for (index, footer) in [page.footer.as_ref(), page.first_footer.as_ref()]
        .into_iter()
        .enumerate()
    {
        let Some(footer) = footer.filter(|story| hf_has_flow_content(story)) else {
            continue;
        };
        let Some(distance) = footer.distance_from_edge else {
            continue;
        };
        let name = format!("o2p_footer_{index}");
        let bottom_edge = footer
            .paragraphs
            .last()
            .map(hf_paragraph_metric_runs)
            .and_then(|runs| text::word_line_box_descent_em(&runs))
            .map(|descent| format!("-{}em", format_f64(descent)))
            .unwrap_or_else(|| "\"descender\"".into());
        let _ = write!(
            out,
            "#let {name} = block(width: {width}pt, above: 0pt, below: 0pt)[#set text(bottom-edge: {bottom_edge}); "
        );
        generate_flow_hf_content(out, footer, ctx);
        let _ = writeln!(
            out,
            "]\n#let {name}_height = measure({name}, width: {width}pt).height"
        );
        let edge = format_f64(distance + hf_trailing_space_after_pt(footer));
        if page.margins.bottom >= 0.0 {
            margins[index] = format!("calc.max({base}, {edge}pt + {name}_height)");
        }
        // Out-of-flow placement also handles fixed margins whose footer extends
        // above the body boundary. Growing the body margin cannot move this anchor.
        values[index] = Some(format!(
            "place(top, dy: o2p_footer_bottom - {edge}pt - {name}_height, {name})"
        ));
    }
    if page.first_footer.is_none() {
        margins[1] = margins[0].clone();
    }
    let _ = writeln!(
        out,
        "#let o2p_footer_bottom = {}\n#let o2p_first_footer_bottom = {}",
        margins[0], margins[1]
    );
    tracing::debug!(
        section = ctx.flow_section_index,
        bottom_pt = page.margins.bottom,
        "Measured footer content for body clearance"
    );
    Some(MeasuredFooters { values })
}

impl MeasuredFooters {
    pub(super) fn preserve_unmeasured(
        &self,
        mut value: FlowFooterValue,
        original_bottom: f64,
    ) -> FlowFooterValue {
        // footer-descent is shared by both variants. Restore Typst's default
        // 30% descent for an unanchored story when its peer sets zero.
        let origin = original_bottom.abs() * if value.wants_zero_descent { 1.0 } else { 0.7 };
        value.markup = format!(
            "move(dy: o2p_footer_bottom - {}pt, {})",
            format_f64(origin),
            value.markup
        );
        value.wants_zero_descent = true;
        value
    }

    pub(super) fn value(&self, first: bool) -> Option<FlowFooterValue> {
        self.values[usize::from(first)]
            .as_ref()
            .map(|markup| FlowFooterValue {
                markup: markup.clone(),
                wants_zero_descent: true,
            })
    }

    pub(super) fn write_first_page_adjustment(&self, out: &mut String) {
        // A signed parent-region float adjusts only the first page, including
        // its footnote region. Later pages retain the default footer's capacity.
        out.push_str("#if o2p_first_footer_bottom != o2p_footer_bottom { place(bottom, scope: \"parent\", float: true, clearance: 0pt, block(height: o2p_first_footer_bottom - o2p_footer_bottom)[]) }\n");
    }
}
