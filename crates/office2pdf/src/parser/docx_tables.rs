use super::contexts::{DocxConversionContext, ResolvedTableStyle, apply_table_text_style};
use super::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Color,
    HyperlinkMap, ImageMap, Insets, LineJoin, MAX_TABLE_DEPTH, StyleMap, Table, TableCell,
    TableRow, convert_paragraph_blocks, parse_hex_color,
};
use crate::ir::TableBorderPaintModel;
use crate::parser::units::{emu_to_pt, twips_to_pt};

#[derive(Clone)]
struct RawCell {
    content: Vec<Block>,
    col_span: u32,
    col_index: usize,
    preferred_width: Option<f64>,
    vmerge: Option<String>,
    border: Option<CellBorder>,
    background: Option<Color>,
    has_explicit_background: bool,
    vertical_align: Option<CellVerticalAlign>,
    padding: Option<Insets>,
}

struct RawRow {
    cells: Vec<RawCell>,
    /// Exact row height in points, already converted from `w:trHeight`'s twips.
    height: Option<f64>,
    /// Minimum row height in points, from the same `w:trHeight` read under
    /// `atLeast` (issue #965).
    minimum_height: Option<f64>,
}

fn extract_margin_side_points(side_json: &serde_json::Value) -> Option<f64> {
    let width_type = side_json
        .get("widthType")
        .and_then(|v| v.as_str())
        .unwrap_or("dxa");
    let value = side_json.get("val").and_then(|v| v.as_f64())?;

    match width_type {
        "dxa" => Some(twips_to_pt(value)),
        _ => None,
    }
}

fn extract_insets_from_margins_json(margins_json: &serde_json::Value) -> Option<Insets> {
    let top = margins_json.get("top").and_then(extract_margin_side_points);
    let right = margins_json
        .get("right")
        .and_then(extract_margin_side_points);
    let bottom = margins_json
        .get("bottom")
        .and_then(extract_margin_side_points);
    let left = margins_json
        .get("left")
        .and_then(extract_margin_side_points);

    if top.is_none() && right.is_none() && bottom.is_none() && left.is_none() {
        return None;
    }

    Some(Insets {
        top: top.unwrap_or_default(),
        right: right.unwrap_or_default(),
        bottom: bottom.unwrap_or_default(),
        left: left.unwrap_or_default(),
    })
}

fn extract_table_alignment(prop_json: Option<&serde_json::Value>) -> Option<Alignment> {
    prop_json
        .and_then(|j| j.get("justification"))
        .and_then(|v| v.as_str())
        .and_then(|value| match value {
            "center" => Some(Alignment::Center),
            "right" | "end" => Some(Alignment::Right),
            _ => None,
        })
}

fn extract_table_default_cell_padding(prop_json: Option<&serde_json::Value>) -> Option<Insets> {
    prop_json
        .and_then(|j| j.get("margins"))
        .and_then(extract_insets_from_margins_json)
}

fn extract_cell_padding(
    prop_json: Option<&serde_json::Value>,
    inherited_padding: Option<Insets>,
) -> Option<Insets> {
    let margins_json = prop_json.and_then(|j| j.get("margins"))?;
    let mut merged_padding = inherited_padding.unwrap_or_default();

    if let Some(top) = margins_json.get("top").and_then(extract_margin_side_points) {
        merged_padding.top = top;
    }
    if let Some(right) = margins_json
        .get("right")
        .and_then(extract_margin_side_points)
    {
        merged_padding.right = right;
    }
    if let Some(bottom) = margins_json
        .get("bottom")
        .and_then(extract_margin_side_points)
    {
        merged_padding.bottom = bottom;
    }
    if let Some(left) = margins_json
        .get("left")
        .and_then(extract_margin_side_points)
    {
        merged_padding.left = left;
    }

    Some(merged_padding)
}

fn extract_table_cell_width(prop_json: Option<&serde_json::Value>) -> Option<f64> {
    let width_json = prop_json.and_then(|j| j.get("width"))?;
    let width_type = width_json
        .get("widthType")
        .and_then(|v| v.as_str())
        .unwrap_or("dxa");
    let width = width_json.get("width").and_then(|v| v.as_f64())?;

    match width_type {
        "dxa" => Some(twips_to_pt(width)),
        _ => None,
    }
}

pub(super) fn convert_table(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    depth: usize,
) -> Table {
    let header_info = ctx.table_headers.consume_next();
    let table_style = ctx.table_styles.consume_next();
    let table_prop_json = serde_json::to_value(&table.property).ok();
    let alignment = extract_table_alignment(table_prop_json.as_ref());
    // Direct table properties win, but a table that states no `w:tblStyle`
    // still inherits the package's default table style. Word's built-in
    // `TableNormal` commonly carries 0pt vertical and 5.4pt horizontal cell
    // margins there; losing it lets Typst's unrelated 5pt inset leak into all
    // four sides (issue #1466).
    let default_cell_padding = extract_table_default_cell_padding(table_prop_json.as_ref())
        .or_else(|| {
            table_style
                .as_ref()
                .and_then(ResolvedTableStyle::default_cell_padding)
        });

    let mut raw_rows = extract_raw_rows(
        table,
        images,
        hyperlinks,
        style_map,
        ctx,
        depth,
        default_cell_padding,
    );
    // The table's own `w:tblBorders`, as tri-states: a stated `none` has to
    // reach the style merge as a suppression rather than as silence
    // (issue #931).
    let direct_table_borders: TableBorderSpec = table_prop_json
        .as_ref()
        .and_then(|json| json.get("borders"))
        .filter(|borders| !borders.is_null())
        .map(extract_table_border_spec)
        .unwrap_or_default();
    if let Some(table_style) = table_style.as_ref() {
        apply_conditional_table_style(&mut raw_rows, table_style, &direct_table_borders);
    }

    let mut column_widths: Vec<f64> = if table.grid.is_empty() {
        derive_column_widths_from_cells(&raw_rows).unwrap_or_default()
    } else {
        let grid: Vec<f64> = table.grid.iter().map(|&w| twips_to_pt(w as f64)).collect();
        // `w:tblW` shares `TableWidth`'s JSON shape with `w:tcW`; docx-rs
        // serializes an absent element as `{width: 0, widthType: "auto"}`,
        // which the extractor (auto) and the filter (0) both reject.
        let declared_table_width_pt: Option<f64> =
            extract_table_cell_width(table_prop_json.as_ref()).filter(|width| *width > 0.0);
        reconcile_auto_layout_widths(&grid, &raw_rows, declared_table_width_pt)
    };

    if header_info.is_visual_rtl {
        let column_count: usize = raw_table_column_count(&raw_rows).max(column_widths.len());
        reverse_raw_rows_for_visual_rtl(&mut raw_rows, column_count);
        column_widths.reverse();
    }

    let mut rows = resolve_vmerge_and_build_rows(&raw_rows);
    apply_table_level_borders(&mut rows, table_prop_json.as_ref());

    Table {
        rows,
        column_widths,
        header_row_count: header_info.repeat_rows.min(table.rows.len()),
        non_repeating_header_row_count: 0,
        alignment,
        default_cell_padding,
        use_content_driven_row_heights: false,
        default_vertical_align: None,
        // Word GT has not verified descender seating for bottom cells (#618).
        seats_bottom_aligned_text_on_descender: false,
        bottom_aligned_descent_floor_pt: 0.0,
        border_paint_model: TableBorderPaintModel::WordPositiveAxisBands,
        prints_gridlines: false,
        prints_headings: false,
        centers_between_print_margins: false,
        print_scale: None,
    }
}

fn reverse_raw_rows_for_visual_rtl(raw_rows: &mut [RawRow], column_count: usize) {
    for row in raw_rows {
        for cell in &mut row.cells {
            let cell_end: usize = cell.col_index + cell.col_span as usize;
            cell.col_index = column_count.saturating_sub(cell_end);
            if let Some(border) = &mut cell.border {
                std::mem::swap(&mut border.left, &mut border.right);
            }
            if let Some(padding) = &mut cell.padding {
                std::mem::swap(&mut padding.left, &mut padding.right);
            }
        }
        row.cells.reverse();
    }
}

fn extract_raw_rows(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    depth: usize,
    default_cell_padding: Option<Insets>,
) -> Vec<RawRow> {
    let mut raw_rows: Vec<RawRow> = Vec::new();

    for table_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = table_child;
        let row_prop_json = serde_json::to_value(&row.property).ok();
        // docx-rs stores `w:trHeight/@w:val` verbatim, and the schema types it
        // as ST_TwipsMeasure — forwarding it as points made every exact-height
        // row 20x too tall (issue #842).
        let declared_height: Option<f64> = row_prop_json
            .as_ref()
            .and_then(|j| j.get("rowHeight"))
            .and_then(|v| v.as_f64())
            .map(twips_to_pt);
        // `@w:hRule` decides what that value means. `exact` pins the row;
        // `auto` discards it; anything else — including the absent attribute,
        // whose schema default is `atLeast` — makes it a floor (issue #965).
        let height_rule: Option<&str> = row_prop_json
            .as_ref()
            .and_then(|j| j.get("heightRule"))
            .and_then(|v| v.as_str());
        let (height, mut minimum_height): (Option<f64>, Option<f64>) = match height_rule {
            Some("exact") => (declared_height, None),
            Some("auto") => (None, None),
            _ => (None, declared_height),
        };
        let mut cells: Vec<RawCell> = Vec::new();
        let mut col_index: usize = 0;
        let mut in_cell_picture_floor: Option<f64> = None;

        for row_child in &row.cells {
            let docx_rs::TableRowChild::TableCell(cell) = row_child;

            let prop_json = serde_json::to_value(&cell.property).ok();
            let grid_span = prop_json
                .as_ref()
                .and_then(|j| j.get("gridSpan"))
                .and_then(|v| v.as_u64())
                .unwrap_or(1) as u32;

            let vmerge = prop_json
                .as_ref()
                .and_then(|j| j.get("verticalMerge"))
                .and_then(|v| v.as_str())
                .map(String::from);
            let preferred_width = extract_table_cell_width(prop_json.as_ref());

            let content = extract_cell_content(cell, images, hyperlinks, style_map, ctx, depth);
            if let Some(picture_height) = in_cell_floating_picture_height(cell) {
                in_cell_picture_floor = Some(
                    in_cell_picture_floor.map_or(picture_height, |floor| floor.max(picture_height)),
                );
            }
            let border = prop_json
                .as_ref()
                .and_then(|j| j.get("borders"))
                .and_then(extract_cell_borders);
            let shading = prop_json
                .as_ref()
                .and_then(|j| j.get("shading"))
                .filter(|value| !value.is_null());
            let background = shading.and_then(extract_cell_shading);
            let has_explicit_background = shading.is_some();
            let vertical_align = prop_json
                .as_ref()
                .and_then(|j| j.get("verticalAlign"))
                .and_then(|v| v.as_str())
                .and_then(|s| match s {
                    "center" => Some(CellVerticalAlign::Center),
                    "bottom" => Some(CellVerticalAlign::Bottom),
                    _ => None,
                });
            let padding = extract_cell_padding(prop_json.as_ref(), default_cell_padding);

            cells.push(RawCell {
                content,
                col_span: grid_span,
                col_index,
                preferred_width,
                vmerge,
                border,
                background,
                has_explicit_background,
                vertical_align,
                padding,
            });

            col_index += grid_span as usize;
        }

        // A floating picture with `wp:anchor/@layoutInCell="1"` still belongs
        // to its table cell's layout boundary. For an auto/at-least row Word
        // therefore grows the track to the picture's `wp:extent/@cy`; treating
        // it as a zero-height overlay collapsed #1368's 343.81pt picture back
        // to the declared 283.5pt floor and kept the following row on page 1.
        // An exact row remains exact and may let the drawing overhang.
        if height.is_none()
            && let Some(picture_floor) = in_cell_picture_floor
        {
            minimum_height =
                Some(minimum_height.map_or(picture_floor, |floor| floor.max(picture_floor)));
        }

        align_top_oriented_cells_to_row_vertical_margins(&mut cells, default_cell_padding);

        raw_rows.push(RawRow {
            cells,
            height,
            minimum_height,
        });
    }

    raw_rows
}

fn in_cell_floating_picture_height(cell: &docx_rs::TableCell) -> Option<f64> {
    cell.children
        .iter()
        .filter_map(|content| match content {
            docx_rs::TableCellContent::Paragraph(paragraph) => {
                paragraph_in_cell_floating_picture_height(paragraph)
            }
            docx_rs::TableCellContent::StructuredDataTag(sdt) => {
                sdt_in_cell_floating_picture_height(sdt)
            }
            // A nested table calculates the constraint on its own rows.
            _ => None,
        })
        .reduce(f64::max)
}

fn sdt_in_cell_floating_picture_height(sdt: &docx_rs::StructuredDataTag) -> Option<f64> {
    sdt.children
        .iter()
        .filter_map(|child| match child {
            docx_rs::StructuredDataTagChild::Run(run) => run_in_cell_floating_picture_height(run),
            docx_rs::StructuredDataTagChild::Paragraph(paragraph) => {
                paragraph_in_cell_floating_picture_height(paragraph)
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                sdt_in_cell_floating_picture_height(nested)
            }
            // A nested table calculates the constraint on its own rows.
            _ => None,
        })
        .reduce(f64::max)
}

fn paragraph_in_cell_floating_picture_height(paragraph: &docx_rs::Paragraph) -> Option<f64> {
    super::flatten_tracked_changes(&paragraph.children)
        .into_iter()
        .filter_map(|item| match item {
            super::ParagraphItem::Run(run) => run_in_cell_floating_picture_height(run),
            super::ParagraphItem::Hyperlink(hyperlink) => hyperlink
                .children
                .iter()
                .filter_map(|child| match child {
                    docx_rs::ParagraphChild::Run(run) => run_in_cell_floating_picture_height(run),
                    _ => None,
                })
                .reduce(f64::max),
            _ => None,
        })
        .reduce(f64::max)
}

fn run_in_cell_floating_picture_height(run: &docx_rs::Run) -> Option<f64> {
    run.children
        .iter()
        .filter_map(|child| match child {
            docx_rs::RunChild::Drawing(drawing) => match drawing.data.as_ref() {
                Some(docx_rs::DrawingData::Pic(picture))
                    if picture.position_type == docx_rs::DrawingPositionType::Anchor
                        && picture.layout_in_cell
                        && picture.size.1 > 0 =>
                {
                    Some(emu_to_pt(picture.size.1))
                }
                _ => None,
            },
            _ => None,
        })
        .reduce(f64::max)
}

fn align_top_oriented_cells_to_row_vertical_margins(
    cells: &mut [RawCell],
    default_cell_padding: Option<Insets>,
) {
    let Some(default_cell_padding) = default_cell_padding else {
        return;
    };

    let top_oriented =
        |cell: &&RawCell| matches!(cell.vertical_align, None | Some(CellVerticalAlign::Top));
    let max_top = cells
        .iter()
        .filter(top_oriented)
        .map(|cell| cell.padding.unwrap_or(default_cell_padding).top)
        .fold(default_cell_padding.top, f64::max);
    let max_bottom = cells
        .iter()
        .filter(top_oriented)
        .map(|cell| cell.padding.unwrap_or(default_cell_padding).bottom)
        .fold(default_cell_padding.bottom, f64::max);

    for cell in cells
        .iter_mut()
        .filter(|cell| matches!(cell.vertical_align, None | Some(CellVerticalAlign::Top)))
    {
        let mut effective_padding = cell.padding.unwrap_or(default_cell_padding);
        if effective_padding.top != max_top || effective_padding.bottom != max_bottom {
            effective_padding.top = max_top;
            effective_padding.bottom = max_bottom;
            cell.padding = Some(effective_padding);
        }
    }
}

fn apply_conditional_table_style(
    raw_rows: &mut [RawRow],
    table_style: &ResolvedTableStyle,
    direct_borders: &TableBorderSpec,
) {
    let row_count = raw_rows.len();
    let column_count = raw_table_column_count(raw_rows);
    for (row_index, row) in raw_rows.iter_mut().enumerate() {
        for cell in &mut row.cells {
            let style = table_style.cell_style(
                row_index,
                row_count,
                cell.col_index,
                cell.col_span as usize,
                column_count,
                direct_borders,
            );
            if !cell.has_explicit_background {
                cell.background = style.background;
            }
            // Explicit tcBorders on the cell win over the style's borders.
            if cell.border.is_none() {
                cell.border = style.border.clone();
            }
            // A conditional region's `w:tcMar` applies only where that region
            // is active. Direct per-cell margins were resolved before this
            // pass and keep precedence.
            if cell.padding.is_none() {
                cell.padding = style.padding;
            }
            apply_table_text_style(&mut cell.content, &style);
        }
    }
}

/// Reconcile a table's `w:tblGrid` with the per-row `w:tcW` preferences its
/// cells declare, the way Word's auto table layout does.
///
/// `w:tblGrid` is only a starting point: without `<w:tblLayout
/// w:type="fixed"/>` Word treats the per-cell widths as preferences and
/// resolves them across every row, which can contradict the grid outright.
/// The invoice fixture's item rows ask for `700/4200/1200/1450/1476` twips
/// while its Subtotal/VAT/Total rows put a 4200-twip value cell in the last
/// column, so Word widens Amount from the grid's 73.8pt to 153.3pt. Taking
/// the grid verbatim left it less than half Word's (issue #355).
///
/// Word resolves the conflict by compressing each column in proportion to
/// its compressible slack above min-content, not by a uniform scale: with
/// `pref_i` the widest single-column `w:tcW` on grid column `i`, `min_i` its
/// widest unbreakable token plus cell side margins, and `W` the fit width,
///
/// ```text
/// k = (W - Σmin) / (Σpref - Σmin);   width_i = min_i + (pref_i - min_i)·k
/// ```
///
/// Derived on the invoice's Word GT (issue #624): the uniform scale put
/// Description and Amount at an identical 161.32pt where Word prints 156.9
/// and 153.3, while this rule lands every column within 0.10pt. The model runs
/// ONLY in that direction — `Σpref > W` compression with `k < 1`, which is
/// what GT verified.
///
/// `Σpref <= W` is not a conflict at all, so `w:tblGrid` stands, scaled to the
/// fit width (issue #925, measured on a second Word GT). Rescaling the tcW
/// maxima there only ever agreed with Word at `Σpref == W`.
///
/// The pre-#624 uniform scale over the per-column tcW maxima survives as the
/// degrade target for the cases that cannot be measured at all: any token
/// without a face or a glyph (wasm, missing font), tables whose cells are all
/// empty (no measurement to anchor the minima), and a compression whose
/// columns already sit at min-content. Font-less and wasm output is unchanged
/// by #624; #925 does move it, since the surplus branch runs before any
/// measurement.
fn reconcile_auto_layout_widths(
    grid: &[f64],
    raw_rows: &[RawRow],
    declared_table_width_pt: Option<f64>,
) -> Vec<f64> {
    let Some(cell_maxima) = derive_column_widths_from_cells(raw_rows) else {
        return grid.to_vec();
    };
    if cell_maxima.len() != grid.len() {
        return grid.to_vec();
    }
    let cell_maxima_total: f64 = cell_maxima.iter().sum();
    let grid_total: f64 = grid.iter().sum();
    if cell_maxima_total <= 0.0
        || grid_total <= 0.0
        || cell_maxima.iter().any(|width| *width <= 0.0)
    {
        return grid.to_vec();
    }
    // The pre-#624 result: one uniform scale over the per-column tcW maxima.
    // Kept verbatim as the degrade target so environments that cannot measure
    // text (wasm, missing fonts) keep producing today's output.
    let uniform_scale: f64 = grid_total / cell_maxima_total;
    let uniformly_scaled: Vec<f64> = cell_maxima
        .iter()
        .map(|width| width * uniform_scale)
        .collect();

    // Word fits the table to `w:tblW` when stated, else to the grid total —
    // but only the SHRINK direction is verified against GT: every #624
    // measurement (the invoice) has 0 < tblW <= grid total and Σpref > W. A
    // tblW beyond the grid total is where Word starts clamping to the section
    // content width, which is not modeled (section geometry is not threaded
    // into tables), so such tables keep the grid total as their fit target.
    let fit_width_pt: f64 = match declared_table_width_pt {
        Some(declared)
            if declared > 0.0 && declared <= grid_total + AUTO_LAYOUT_WIDTH_EPSILON_PT =>
        {
            declared
        }
        _ => grid_total,
    };

    let preferred: Vec<f64> = derive_grid_column_preferences(grid, raw_rows);
    let preferred_total: f64 = preferred.iter().sum();
    // Σpref <= W is the surplus direction: nothing is over-subscribed, so
    // there is no conflict for the slack model to resolve and `w:tblGrid` —
    // Word's own last layout of this table — stands, scaled to the fit width.
    //
    // Rescaling the `w:tcW` maxima instead only ever agreed with Word at
    // Σpref == W. Below it the scale is not ≈1 and the grid does NOT come
    // back: the invoice of #925 states a 2403/2656/2997/2550 grid against
    // stale 2092/2313/2610/1620 cells, and the 1.2283 rescale handed its last
    // column 99.5pt where Word prints 127.5, so a 12pt `FORFALLSDATO` header
    // spilled past the table's right edge. Only Σpref > W (k < 1 compression,
    // issue #624) needs the slack model.
    if preferred_total <= fit_width_pt + AUTO_LAYOUT_WIDTH_EPSILON_PT {
        let grid_scale: f64 = fit_width_pt / grid_total;
        return grid.iter().map(|width| width * grid_scale).collect();
    }
    let Some(min_content) = derive_grid_column_min_content_widths(raw_rows, grid.len()) else {
        return uniformly_scaled;
    };

    // A preference below min-content carries no compressible slack: the
    // column floors at min-content and takes no share of the surplus.
    let clamped_preferred: Vec<f64> = preferred
        .iter()
        .zip(&min_content)
        .map(|(preference, min)| preference.max(*min))
        .collect();
    let min_total: f64 = min_content.iter().sum();
    let compressible_slack: f64 = clamped_preferred.iter().sum::<f64>() - min_total;
    if compressible_slack <= AUTO_LAYOUT_WIDTH_EPSILON_PT {
        // Every column already sits at min-content; how Word grows such a
        // table to a wider tblW is unmeasured, so keep today's output.
        return uniformly_scaled;
    }
    // k < 1 always holds here (Σpref > W was gated above, so the extrapolating
    // k >= 1 branch never reaches this point); k clamps at 0 when W < Σmin,
    // flooring every column at min-content and letting the table overflow W,
    // which is untested.
    let slack_share: f64 = ((fit_width_pt - min_total) / compressible_slack).max(0.0);
    clamped_preferred
        .iter()
        .zip(&min_content)
        .map(|(preference, min)| min + (preference - min) * slack_share)
        .collect()
}

/// One twip (0.05pt) — the resolution of every source value. A conflict
/// smaller than one twip is dxa rounding noise, not an authored disagreement,
/// so it falls on the no-conflict side of the surplus gate: the grid scaled to
/// the fit width, reached without measuring a single token. The same epsilon
/// guards the min-content saturation check, whose degrade target is still the
/// uniform scale.
const AUTO_LAYOUT_WIDTH_EPSILON_PT: f64 = 0.05;

/// The preferred width of each grid column: the widest `w:tcW` any
/// single-column cell states on it, falling back to the declared `gridCol`.
///
/// Occupancy is tracked through `w:gridSpan` — a cell following a span-4 cell
/// sits on grid column 5 and its `w:tcW` claims that column, which is exactly
/// how the invoice's Subtotal rows hand their 4200-twip value cell to the
/// last column. A spanned cell's own `w:tcW` is ignored unless it exceeds the
/// sum of its spanned columns' preferences (untested by fixtures: the excess
/// is spread proportionally).
fn derive_grid_column_preferences(grid: &[f64], raw_rows: &[RawRow]) -> Vec<f64> {
    let mut stated: Vec<Option<f64>> = vec![None; grid.len()];
    for row in raw_rows {
        for cell in &row.cells {
            if cell.col_span != 1 || cell.col_index >= grid.len() {
                continue;
            }
            let Some(preferred_width) = cell.preferred_width else {
                continue;
            };
            let slot: &mut Option<f64> = &mut stated[cell.col_index];
            *slot = Some(slot.map_or(preferred_width, |width| width.max(preferred_width)));
        }
    }
    let mut preferred: Vec<f64> = stated
        .iter()
        .zip(grid)
        .map(|(stated_width, grid_width)| stated_width.unwrap_or(*grid_width))
        .collect();
    raise_spanned_ranges_to_spanning_cells(&mut preferred, raw_rows, |cell| cell.preferred_width);
    preferred
}

/// The min-content width of each grid column: over its single-column cells,
/// the widest unbreakable token plus the cell's left and right margins.
///
/// Word never compresses a column below this in auto layout. Borders are NOT
/// added — measured on the invoice, adding them degrades the fit. Returns
/// `None` when any cell's text cannot be measured, and also when NO cell
/// produced a font measurement at all (every cell empty or whitespace-only):
/// a margins-only minimum is unverified against GT, and running the slack
/// model from it would move empty form-skeleton tables away from today's
/// output on native while wasm — which never measures — kept the uniform
/// scale. Degrading keeps both targets identical.
fn derive_grid_column_min_content_widths(
    raw_rows: &[RawRow],
    column_count: usize,
) -> Option<Vec<f64>> {
    let mut any_token_measured: bool = false;

    let mut min_content: Vec<f64> = vec![0.0; column_count];
    for row in raw_rows {
        for cell in &row.cells {
            if cell.col_span != 1 || cell.col_index >= column_count {
                continue;
            }
            let cell_min: f64 = measured_cell_min_content_pt(cell, &mut any_token_measured)?;
            min_content[cell.col_index] = min_content[cell.col_index].max(cell_min);
        }
    }
    // Spanned cells' mins are ignored unless exceeding the spanned columns'
    // min sum (untested by fixtures) — but their text must still be
    // measurable, or the whole table degrades consistently.
    let mut spanned_all_measured: bool = true;
    raise_spanned_ranges_to_spanning_cells(&mut min_content, raw_rows, |cell| {
        match measured_cell_min_content_pt(cell, &mut any_token_measured) {
            Some(cell_min) => Some(cell_min),
            None => {
                spanned_all_measured = false;
                None
            }
        }
    });
    (spanned_all_measured && any_token_measured).then_some(min_content)
}

/// One cell's min-content: its widest unbreakable token plus its left and
/// right margins (`w:tcMar` default 108 twips = 5.4pt per writing side when
/// neither the cell nor the table states one). Sets `any_token_measured`
/// when at least one real font measurement backed the result.
fn measured_cell_min_content_pt(cell: &RawCell, any_token_measured: &mut bool) -> Option<f64> {
    const DEFAULT_CELL_SIDE_MARGIN_PT: f64 = 5.4;
    let widest_token_pt: f64 = max_unbreakable_token_advance_pt(&cell.content, any_token_measured)?;
    let (left_margin, right_margin): (f64, f64) = cell.padding.map_or(
        (DEFAULT_CELL_SIDE_MARGIN_PT, DEFAULT_CELL_SIDE_MARGIN_PT),
        |padding| (padding.left, padding.right),
    );
    Some(widest_token_pt + left_margin + right_margin)
}

/// Shared spanned-cell rule for preferences and min-content: when a
/// `w:gridSpan` cell's own requirement exceeds the sum its spanned columns
/// already carry, raise those columns proportionally to cover it. No fixture
/// exercises this branch; the invoice's span-4 label cells all require less
/// than their columns' sums and are ignored here.
fn raise_spanned_ranges_to_spanning_cells(
    column_values: &mut [f64],
    raw_rows: &[RawRow],
    mut cell_requirement: impl FnMut(&RawCell) -> Option<f64>,
) {
    for row in raw_rows {
        for cell in &row.cells {
            let span: usize = cell.col_span as usize;
            if span < 2 {
                continue;
            }
            let range_end: usize = (cell.col_index + span).min(column_values.len());
            if cell.col_index >= range_end {
                continue;
            }
            let Some(required_width) = cell_requirement(cell) else {
                continue;
            };
            let range = &mut column_values[cell.col_index..range_end];
            let range_sum: f64 = range.iter().sum();
            if required_width > range_sum && range_sum > 0.0 {
                let scale: f64 = required_width / range_sum;
                for value in range.iter_mut() {
                    *value *= scale;
                }
            }
        }
    }
}

/// Word breaks tokens at ordinary whitespace, but a no-break space (U+00A0),
/// narrow no-break space (U+202F), or figure space (U+2007) stays inside the
/// token — "1 240,00" with an NBSP thousands separator is one token.
fn is_token_breaking_whitespace(character: char) -> bool {
    character.is_whitespace() && !matches!(character, '\u{00A0}' | '\u{202F}' | '\u{2007}')
}

/// The advance of the widest unbreakable token in a cell's paragraphs, in
/// points, measured with each run's resolved family, weight, and size — bold
/// runs use the bold face, East Asian codepoints the `w:eastAsia` face.
///
/// Token boundaries mirror Word's line breaking: breaking whitespace closes a
/// token (no-break spaces do not — see [`is_token_breaking_whitespace`]), and
/// a CJK character is ALWAYS a token of its own because Word may break
/// between any two CJK characters — a Korean phrase's min-content is its
/// widest single glyph, and "모델A" splits as 모/델/A. Everything else forms
/// maximal non-CJK segments that accumulate across run boundaries (a price
/// like "$1,240.00" split over runs is one unbreakable token).
/// TODO(issue #624): Word also breaks after hyphens, and kinsoku forbids
/// breaks before CJK closing punctuation / after opening punctuation; no
/// fixture exercises either, so neither is modeled here.
///
/// Each contiguous same-family segment is measured with ONE
/// `text_advance_em` call, keeping the global face-cache mutex out of the
/// per-character path. `any_token_measured` is set when at least one call
/// succeeded, so callers can tell a real measurement from the vacuous 0 of an
/// empty cell.
///
/// Non-paragraph blocks (images, nested tables) also bound Word's
/// min-content, but no auto-layout fixture carries them, so they contribute
/// nothing rather than blocking measurement. Returns `None` when a run's
/// family or size is unresolved or a glyph is missing, so the caller can
/// degrade to a measurement-free path.
fn max_unbreakable_token_advance_pt(
    blocks: &[Block],
    any_token_measured: &mut bool,
) -> Option<f64> {
    let mut widest_token_pt: f64 = 0.0;
    for block in blocks {
        let Block::Paragraph(paragraph) = block else {
            continue;
        };
        let mut current_token_pt: f64 = 0.0;
        for run in &paragraph.runs {
            if run.text.chars().all(is_token_breaking_whitespace) {
                if !run.text.is_empty() {
                    widest_token_pt = widest_token_pt.max(current_token_pt);
                    current_token_pt = 0.0;
                }
                continue;
            }
            let font_family: &str = run.style.font_family.as_deref()?;
            let font_size: f64 = run.style.font_size?;
            let is_bold: bool = run.style.bold == Some(true);

            let text: &str = &run.text;
            // Start of the current maximal non-CJK, non-breaking segment.
            let mut segment_start: Option<usize> = None;
            for (byte_index, character) in text.char_indices() {
                if !is_token_breaking_whitespace(character)
                    && !crate::render::typst_gen::is_cjk_like(character)
                {
                    segment_start.get_or_insert(byte_index);
                    continue;
                }
                if let Some(start) = segment_start.take() {
                    current_token_pt += measured_segment_advance_pt(
                        font_family,
                        is_bold,
                        font_size,
                        &text[start..byte_index],
                        any_token_measured,
                    )?;
                }
                if is_token_breaking_whitespace(character) {
                    widest_token_pt = widest_token_pt.max(current_token_pt);
                    current_token_pt = 0.0;
                    continue;
                }
                // A CJK character: break before and after, so it is a
                // singleton token measured with the `w:eastAsia` face.
                widest_token_pt = widest_token_pt.max(current_token_pt);
                current_token_pt = 0.0;
                let cjk_family: &str = run
                    .style
                    .east_asian_font_family
                    .as_deref()
                    .unwrap_or(font_family);
                let character_end: usize = byte_index + character.len_utf8();
                let singleton_pt: f64 = measured_segment_advance_pt(
                    cjk_family,
                    is_bold,
                    font_size,
                    &text[byte_index..character_end],
                    any_token_measured,
                )?;
                widest_token_pt = widest_token_pt.max(singleton_pt);
            }
            if let Some(start) = segment_start.take() {
                // The token stays open: it may continue into the next run.
                current_token_pt += measured_segment_advance_pt(
                    font_family,
                    is_bold,
                    font_size,
                    &text[start..],
                    any_token_measured,
                )?;
            }
        }
        widest_token_pt = widest_token_pt.max(current_token_pt);
    }
    Some(widest_token_pt)
}

/// One `text_advance_em` call for a whole token segment, converted to points.
/// Marks `any_token_measured` on success so callers can tell a real
/// measurement from the vacuous zero of an empty cell.
fn measured_segment_advance_pt(
    family: &str,
    is_bold: bool,
    font_size_pt: f64,
    segment: &str,
    any_token_measured: &mut bool,
) -> Option<f64> {
    let advance_em: f64 = crate::render::pdf::text_advance_em(family, is_bold, segment)?;
    *any_token_measured = true;
    Some(advance_em * font_size_pt)
}

fn derive_column_widths_from_cells(raw_rows: &[RawRow]) -> Option<Vec<f64>> {
    let num_cols: usize = raw_table_column_count(raw_rows);

    if num_cols == 0 {
        return None;
    }

    let mut widths: Vec<f64> = vec![0.0; num_cols];
    let mut saw_width = false;

    for row in raw_rows {
        for cell in &row.cells {
            let Some(preferred_width) = cell.preferred_width else {
                continue;
            };
            if cell.col_span == 0 {
                continue;
            }

            let per_column_width = preferred_width / cell.col_span as f64;
            for width in widths
                .iter_mut()
                .skip(cell.col_index)
                .take(cell.col_span as usize)
            {
                *width = width.max(per_column_width);
            }
            saw_width = true;
        }
    }

    saw_width.then_some(widths)
}

fn raw_table_column_count(raw_rows: &[RawRow]) -> usize {
    raw_rows
        .iter()
        .flat_map(|row| {
            row.cells
                .iter()
                .map(|cell| cell.col_index + cell.col_span as usize)
        })
        .max()
        .unwrap_or_default()
}

fn resolve_vmerge_and_build_rows(raw_rows: &[RawRow]) -> Vec<TableRow> {
    let mut rows: Vec<TableRow> = Vec::new();

    for (row_idx, raw_row) in raw_rows.iter().enumerate() {
        let mut cells: Vec<TableCell> = Vec::new();

        for raw_cell in &raw_row.cells {
            match raw_cell.vmerge.as_deref() {
                Some("continue") => continue,
                Some("restart") => {
                    let row_span = count_vmerge_span(raw_rows, row_idx, raw_cell.col_index);
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                        background_alpha: None,
                        data_bar: None,
                        sparkline: None,
                        icon_text: None,
                        icon_color: None,
                        icon_shading: None,
                        spill_width: None,
                        vertical_align: raw_cell.vertical_align,
                        padding: raw_cell.padding,
                    });
                }
                _ => {
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span: 1,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                        background_alpha: None,
                        data_bar: None,
                        sparkline: None,
                        icon_text: None,
                        icon_color: None,
                        icon_shading: None,
                        spill_width: None,
                        vertical_align: raw_cell.vertical_align,
                        padding: raw_cell.padding,
                    });
                }
            }
        }

        rows.push(TableRow {
            cells,
            height: raw_row.height,
            minimum_height: raw_row.minimum_height,
        });
    }

    rows
}

fn count_vmerge_span(raw_rows: &[RawRow], start_row: usize, col_index: usize) -> u32 {
    let mut span: u32 = 1;
    for row in raw_rows.iter().skip(start_row + 1) {
        let has_continue = row
            .cells
            .iter()
            .any(|c| c.col_index == col_index && c.vmerge.as_deref() == Some("continue"));
        if has_continue {
            span += 1;
        } else {
            break;
        }
    }
    span
}

fn extract_cell_content(
    cell: &docx_rs::TableCell,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    depth: usize,
) -> Vec<Block> {
    let mut blocks: Vec<Block> = Vec::new();
    for content in &cell.children {
        match content {
            docx_rs::TableCellContent::Paragraph(para) => {
                convert_paragraph_blocks(para, &mut blocks, images, hyperlinks, style_map, ctx);
            }
            docx_rs::TableCellContent::Table(nested_table) if depth < MAX_TABLE_DEPTH => {
                blocks.push(Block::Table(convert_table(
                    nested_table,
                    images,
                    hyperlinks,
                    style_map,
                    ctx,
                    depth + 1,
                )));
            }
            // A content control wrapping whole paragraphs — the block-level
            // form Word writes for a placeholder line. Its children are
            // ordinary cell content; dropping the wrapper dropped them with it
            // (issue #844).
            docx_rs::TableCellContent::StructuredDataTag(sdt) => {
                extend_with_cell_sdt_content(
                    &mut blocks,
                    sdt,
                    images,
                    hyperlinks,
                    style_map,
                    ctx,
                    depth,
                );
            }
            _ => {}
        }
    }
    blocks
}

/// Append a block-level `w:sdt`'s paragraphs and tables to a cell's content,
/// in document order, descending through nested controls.
fn extend_with_cell_sdt_content(
    blocks: &mut Vec<Block>,
    sdt: &docx_rs::StructuredDataTag,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
    depth: usize,
) {
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                convert_paragraph_blocks(para, blocks, images, hyperlinks, style_map, ctx);
            }
            docx_rs::StructuredDataTagChild::Table(nested_table) if depth < MAX_TABLE_DEPTH => {
                blocks.push(Block::Table(convert_table(
                    nested_table,
                    images,
                    hyperlinks,
                    style_map,
                    ctx,
                    depth + 1,
                )));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                extend_with_cell_sdt_content(
                    blocks, nested, images, hyperlinks, style_map, ctx, depth,
                );
            }
            _ => {}
        }
    }
}

/// Expand table-level `w:tblBorders` onto cells that carry no explicit
/// borders of their own: outer sides on edge cells, insideH/insideV between
/// cells. Previously these tables relied on Typst's default grid, which the
/// renderer no longer paints.
fn apply_table_level_borders(rows: &mut [TableRow], table_prop_json: Option<&serde_json::Value>) {
    let Some(borders) = table_prop_json.and_then(|j| j.get("borders")) else {
        return;
    };
    if borders.is_null() {
        return;
    }
    let outer: Option<CellBorder> = extract_cell_borders(borders);
    let inside_h: Option<BorderSide> = extract_border_side(borders, "insideH");
    let inside_v: Option<BorderSide> = extract_border_side(borders, "insideV");
    if outer.is_none() && inside_h.is_none() && inside_v.is_none() {
        return;
    }

    let row_count = rows.len();
    for (row_index, row) in rows.iter_mut().enumerate() {
        let cell_count = row.cells.len();
        for (cell_index, cell) in row.cells.iter_mut().enumerate() {
            if cell.border.is_some() {
                continue;
            }
            let is_first_row = row_index == 0;
            let is_last_row = row_index + 1 == row_count;
            let is_first_col = cell_index == 0;
            let is_last_col = cell_index + 1 == cell_count;
            let border = CellBorder {
                top: if is_first_row {
                    outer.as_ref().and_then(|b| b.top.clone())
                } else {
                    inside_h.clone()
                },
                bottom: if is_last_row {
                    outer.as_ref().and_then(|b| b.bottom.clone())
                } else {
                    inside_h.clone()
                },
                left: if is_first_col {
                    outer.as_ref().and_then(|b| b.left.clone())
                } else {
                    inside_v.clone()
                },
                right: if is_last_col {
                    outer.as_ref().and_then(|b| b.right.clone())
                } else {
                    inside_v.clone()
                },
            };
            if border.top.is_some()
                || border.bottom.is_some()
                || border.left.is_some()
                || border.right.is_some()
            {
                cell.border = Some(border);
            }
        }
    }
}

fn extract_border_side(borders_json: &serde_json::Value, key: &str) -> Option<BorderSide> {
    extract_cell_borders(&serde_json::json!({ "top": borders_json.get(key)? }))
        .and_then(|border| border.top)
}

/// One side of a `w:tblBorders`/`w:tcBorders`, keeping "stated as none" apart
/// from "not stated at all".
///
/// ECMA-376 gives `w:val="none"` (and its legacy spelling `nil`) the meaning
/// "no border", which outranks whatever the table style would have drawn;
/// only an *absent* side inherits. Collapsing the two onto `None` made the
/// invoice of #841 draw six grey rules its `w:tblBorders` asks not to have
/// (issue #931).
#[derive(Debug, Clone, Default)]
pub(super) enum BorderSideSpec {
    /// The element states nothing; the style decides.
    #[default]
    Unstated,
    /// `w:val="none"` or `"nil"` — draw nothing, and stop the style lookup.
    Suppressed,
    /// A border to draw.
    Stated(BorderSide),
}

impl BorderSideSpec {
    /// The side to draw, once the resolution above has run.
    pub(super) fn drawn(&self) -> Option<BorderSide> {
        match self {
            Self::Stated(side) => Some(side.clone()),
            Self::Unstated | Self::Suppressed => None,
        }
    }

    pub(super) fn is_stated(&self) -> bool {
        !matches!(self, Self::Unstated)
    }
}

/// The six sides a `w:tblBorders` can state, each as a tri-state.
#[derive(Debug, Clone, Default)]
pub(super) struct TableBorderSpec {
    pub(super) top: BorderSideSpec,
    pub(super) bottom: BorderSideSpec,
    pub(super) left: BorderSideSpec,
    pub(super) right: BorderSideSpec,
    pub(super) inside_h: BorderSideSpec,
    pub(super) inside_v: BorderSideSpec,
}

impl TableBorderSpec {
    /// The side this spec states for a cell's top or bottom edge, given where
    /// the cell sits: the outer side on a boundary row, `insideH` between
    /// rows. `outer` picks which of `top`/`bottom` the caller means.
    pub(super) fn horizontal<'a>(
        &'a self,
        at_boundary: bool,
        outer: &'a BorderSideSpec,
    ) -> &'a BorderSideSpec {
        if at_boundary { outer } else { &self.inside_h }
    }

    /// As [`Self::horizontal`], for a cell's left and right edges.
    pub(super) fn vertical<'a>(
        &'a self,
        at_boundary: bool,
        outer: &'a BorderSideSpec,
    ) -> &'a BorderSideSpec {
        if at_boundary { outer } else { &self.inside_v }
    }
}

/// Read a `w:tblBorders`/`w:tcBorders` element as tri-states, so a stated
/// `none` survives to the style merge (issue #931).
pub(super) fn extract_table_border_spec(borders_json: &serde_json::Value) -> TableBorderSpec {
    let side = |key: &str| -> BorderSideSpec {
        let Some(value) = borders_json.get(key) else {
            return BorderSideSpec::Unstated;
        };
        if value.is_null() {
            return BorderSideSpec::Unstated;
        }
        match extract_border_side(borders_json, key) {
            Some(border) => BorderSideSpec::Stated(border),
            None => BorderSideSpec::Suppressed,
        }
    };
    TableBorderSpec {
        top: side("top"),
        bottom: side("bottom"),
        left: side("left"),
        right: side("right"),
        inside_h: side("insideH"),
        inside_v: side("insideV"),
    }
}

fn extract_cell_borders(borders_json: &serde_json::Value) -> Option<CellBorder> {
    if borders_json.is_null() {
        return None;
    }

    let extract_side = |key: &str| -> Option<BorderSide> {
        let side = borders_json.get(key)?;
        if side.is_null() {
            return None;
        }
        let border_type = side
            .get("borderType")
            .and_then(|v| v.as_str())
            .unwrap_or("none");
        if border_type == "none" || border_type == "nil" {
            return None;
        }
        let size = side.get("size").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let color_hex = side
            .get("color")
            .and_then(|v| v.as_str())
            .unwrap_or("000000");
        let color = parse_hex_color(color_hex).unwrap_or(Color::black());
        let style = match border_type {
            "dashed" | "dashSmallGap" => BorderLineStyle::Dashed,
            "dotted" => BorderLineStyle::Dotted,
            "dashDotStroked" | "dotDash" => BorderLineStyle::DashDot,
            "dotDotDash" => BorderLineStyle::DashDotDot,
            "double"
            | "thinThickSmallGap"
            | "thickThinSmallGap"
            | "thinThickMediumGap"
            | "thickThinMediumGap"
            | "thinThickLargeGap"
            | "thickThinLargeGap"
            | "thinThickThinSmallGap"
            | "thinThickThinMediumGap"
            | "thinThickThinLargeGap"
            | "triple" => BorderLineStyle::Double,
            _ => BorderLineStyle::Solid,
        };
        Some(BorderSide {
            width: size / 8.0,
            color,
            style,
            join: LineJoin::Round,
        })
    };

    let top = extract_side("top");
    let bottom = extract_side("bottom");
    let left = extract_side("left");
    let right = extract_side("right");

    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }

    Some(CellBorder {
        top,
        bottom,
        left,
        right,
    })
}

fn extract_cell_shading(shading_json: &serde_json::Value) -> Option<Color> {
    if shading_json.is_null() {
        return None;
    }
    let fill = shading_json.get("fill").and_then(|v| v.as_str())?;
    if fill == "auto" || fill.is_empty() {
        return None;
    }
    parse_hex_color(fill)
}
