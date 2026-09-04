use super::*;

pub(super) fn generate_table(
    out: &mut String,
    table: &Table,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    ctx.table_depth += 1;
    // A nested table decides its cells' effective vertical alignment from its
    // own defaults; restore the enclosing table's answers afterwards, like
    // `row_east_asian` below.
    let enclosing_default_vertical_align: Option<CellVerticalAlign> =
        ctx.table_default_vertical_align;
    let enclosing_seats_on_descender: bool = ctx.table_seats_bottom_aligned_text_on_descender;
    let enclosing_box_is_aligned: bool = ctx.table_box_is_aligned;
    let enclosing_descent_floor_pt: f64 = ctx.table_bottom_aligned_descent_floor_pt;
    let enclosing_print_scale: Option<f64> = ctx.table_print_scale;
    ctx.table_default_vertical_align = table.default_vertical_align;
    ctx.table_seats_bottom_aligned_text_on_descender = table.seats_bottom_aligned_text_on_descender;
    ctx.table_bottom_aligned_descent_floor_pt = table.bottom_aligned_descent_floor_pt;
    ctx.table_print_scale = table.print_scale;
    // `w:tblPr/w:jc` places the table box on the page and says nothing about
    // the text inside it, but Typst inherits `align` into the cells. The cells
    // undo it; a nested table's own answer must not outlive it (issue #843).
    ctx.table_box_is_aligned = matches!(
        table.alignment,
        Some(Alignment::Center) | Some(Alignment::Right)
    );
    let result = match table.alignment {
        Some(Alignment::Center) => {
            out.push_str("#align(center)[\n");
            let result = generate_table_inner(out, table, ctx);
            out.push_str("]\n");
            result
        }
        Some(Alignment::Right) => {
            out.push_str("#align(right)[\n");
            let result = generate_table_inner(out, table, ctx);
            out.push_str("]\n");
            result
        }
        _ => generate_table_inner(out, table, ctx),
    };
    ctx.table_default_vertical_align = enclosing_default_vertical_align;
    ctx.table_seats_bottom_aligned_text_on_descender = enclosing_seats_on_descender;
    ctx.table_bottom_aligned_descent_floor_pt = enclosing_descent_floor_pt;
    ctx.table_print_scale = enclosing_print_scale;
    ctx.table_box_is_aligned = enclosing_box_is_aligned;
    ctx.table_depth -= 1;
    result
}

fn generate_table_inner(
    out: &mut String,
    table: &Table,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    out.push_str("#table(\n");

    // Only explicitly set borders render: Excel does not print gridlines,
    // and Word/PowerPoint borderless tables have none either. Typst's
    // default 1pt grid painted spurious borders on every unbordered table.
    out.push_str("  stroke: none,\n");

    if let Some(ref default_vertical_align) = table.default_vertical_align {
        let align_str: &str = match default_vertical_align {
            CellVerticalAlign::Top => "top",
            CellVerticalAlign::Center => "horizon",
            CellVerticalAlign::Bottom => "bottom",
        };
        let _ = writeln!(out, "  align: {align_str},");
    }

    if let Some(padding) = table.default_cell_padding {
        let _ = writeln!(out, "  inset: {},", format_insets(&padding));
    }

    let num_cols = if !table.column_widths.is_empty() {
        table.column_widths.len()
    } else {
        table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    };

    if !table.column_widths.is_empty() {
        out.push_str("  columns: (");
        for (i, w) in table.column_widths.iter().enumerate() {
            if i > 0 {
                out.push_str(", ");
            }
            let _ = write!(out, "{}pt", format_f64(*w));
        }
        out.push_str("),\n");
    } else if num_cols > 1 {
        let _ = writeln!(out, "  columns: {num_cols},");
    }

    if !table.use_content_driven_row_heights && table.rows.iter().any(|row| row.height.is_some()) {
        out.push_str("  rows: (");
        for (i, row) in table.rows.iter().enumerate() {
            if i > 0 {
                out.push_str(", ");
            }
            match row.height {
                Some(height) => {
                    let _ = write!(out, "{}pt", format_f64(height));
                }
                None => out.push_str("auto"),
            }
        }
        out.push_str("),\n");
    }

    let mut rowspan_remaining = vec![0usize; num_cols];
    // The printed-headings letter strip is `rows[0]` when the XLSX parser
    // materialized it (issue #623); the print-title header counts below start
    // after it.
    let heading_strip_row_count: usize =
        usize::from(table.prints_headings && !table.rows.is_empty());
    let countable_rows: usize = table.rows.len() - heading_strip_row_count;
    let declared_header_row_count = table.header_row_count.min(countable_rows);
    let default_cell_padding = table.default_cell_padding.unwrap_or(Insets {
        top: 5.0,
        right: 5.0,
        bottom: 5.0,
        left: 5.0,
    });

    let fixed_row_heights = !table.use_content_driven_row_heights;

    // Rows above a print-title range belong to the header block but print only
    // once, so they go in a `repeat: false` header. The repeating title rows
    // then need a higher level to keep repeating alongside it.
    let lead_row_count = table
        .non_repeating_header_row_count
        .min(countable_rows.saturating_sub(declared_header_row_count));
    let lead_start: usize = heading_strip_row_count;
    let title_start: usize = lead_start + lead_row_count;
    let header_row_count: usize =
        header_row_count_covering_rowspans(&table.rows[title_start..], declared_header_row_count);

    // Grid boundaries whose upper side repeats on every page while the lower
    // side prints once: the printed-headings letter strip's bottom (issue
    // #623) and the boundary between the last repeating print-title row and
    // the first body row. Border-band ties there must resolve toward the
    // repeating side (issue #619 review, remediation 2).
    let mut repeating_header_boundaries: Vec<usize> = Vec::new();
    if heading_strip_row_count > 0 && heading_strip_row_count < table.rows.len() {
        repeating_header_boundaries.push(heading_strip_row_count);
    }
    if header_row_count > 0
        && heading_strip_row_count + lead_row_count + header_row_count < table.rows.len()
    {
        repeating_header_boundaries
            .push(heading_strip_row_count + lead_row_count + header_row_count);
    }

    // Excel and Word boundary-band tables resolve which cell paints each
    // shared boundary before emission: the bands are boundary-anchored and
    // declaration-independent, so a boundary declared by both neighbours must
    // paint exactly once (issues #619 and #724). Resolved separately from
    // `TableCell::border` so the layout inset of #500/#503 keeps following
    // each cell's own declaration. Word's content-seat translation (#649)
    // follows the resolved owner later without changing that layout inset.
    let boundary_band_model: Option<TableBorderPaintModel> = match table.border_paint_model {
        TableBorderPaintModel::CenteredStroke => None,
        model => Some(model),
    };
    let painted_borders: Option<Vec<Vec<Option<CellBorder>>>> = boundary_band_model
        .map(|_| resolve_boundary_painted_borders(table, num_cols, &repeating_header_boundaries));
    // A later cell's +x fill bleed can cross a horizontal band emitted by an
    // earlier row, while an upper-right cell's +y bleed can cross a
    // differently painted left neighbour. Precompute both junction trims so
    // Excel fills meet without repainting the corner owner (#1475, #1495).
    let background_bleed_trims: Option<ExcelBackgroundBleedTrims> =
        (boundary_band_model == Some(TableBorderPaintModel::ExcelBoundaryBands)).then(|| {
            excel_background_bleed_trims(
                table,
                num_cols,
                painted_borders
                    .as_deref()
                    .expect("Excel boundary bands must be resolved"),
            )
        });
    if heading_strip_row_count > 0 {
        // GT prints the column-letter strip on every page (issue #623); the
        // outermost header level repeats above the print-title headers below.
        out.push_str("  table.header(repeat: true,\n");
        generate_table_rows(
            out,
            &table.rows[..heading_strip_row_count],
            painted_borders
                .as_deref()
                .map(|p| &p[..heading_strip_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.top[..heading_strip_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.bottom_left[..heading_strip_row_count]),
            boundary_band_model,
            &table.column_widths,
            num_cols,
            &mut rowspan_remaining,
            "    ",
            default_cell_padding,
            fixed_row_heights,
            ctx,
        )?;
        out.push_str("  ),\n");
    }

    if lead_row_count > 0 {
        if heading_strip_row_count > 0 {
            out.push_str("  table.header(repeat: false, level: 2,\n");
        } else {
            out.push_str("  table.header(repeat: false,\n");
        }
        generate_table_rows(
            out,
            &table.rows[lead_start..lead_start + lead_row_count],
            painted_borders
                .as_deref()
                .map(|p| &p[lead_start..lead_start + lead_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.top[lead_start..lead_start + lead_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.bottom_left[lead_start..lead_start + lead_row_count]),
            boundary_band_model,
            &table.column_widths,
            num_cols,
            &mut rowspan_remaining,
            "    ",
            default_cell_padding,
            fixed_row_heights,
            ctx,
        )?;
        out.push_str("  ),\n");
    }

    if header_row_count > 0 {
        // Consecutive Typst headers need strictly increasing levels: the
        // strip (when present) takes level 1 and the lead block the next one,
        // so the print-title header lands below both.
        let title_header_level: usize =
            1 + heading_strip_row_count + usize::from(lead_row_count > 0);
        if title_header_level > 1 {
            let _ = writeln!(out, "  table.header(level: {title_header_level},");
        } else {
            out.push_str("  table.header(\n");
        }
        generate_table_rows(
            out,
            &table.rows[title_start..title_start + header_row_count],
            painted_borders
                .as_deref()
                .map(|p| &p[title_start..title_start + header_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.top[title_start..title_start + header_row_count]),
            background_bleed_trims
                .as_ref()
                .map(|trims| &trims.bottom_left[title_start..title_start + header_row_count]),
            boundary_band_model,
            &table.column_widths,
            num_cols,
            &mut rowspan_remaining,
            "    ",
            default_cell_padding,
            fixed_row_heights,
            ctx,
        )?;
        out.push_str("  ),\n");
    }

    generate_table_rows(
        out,
        &table.rows[title_start + header_row_count..],
        painted_borders
            .as_deref()
            .map(|p| &p[title_start + header_row_count..]),
        background_bleed_trims
            .as_ref()
            .map(|trims| &trims.top[title_start + header_row_count..]),
        background_bleed_trims
            .as_ref()
            .map(|trims| &trims.bottom_left[title_start + header_row_count..]),
        boundary_band_model,
        &table.column_widths,
        num_cols,
        &mut rowspan_remaining,
        "  ",
        default_cell_padding,
        fixed_row_heights,
        ctx,
    )?;

    out.push_str(")\n");
    Ok(())
}

/// Keep rows touched by a header-originating rowspan in the same Typst header
/// block. Typst cannot continue a header cell in the body; splitting such a
/// merge makes cells in the following rows overflow the available columns.
pub(crate) fn header_row_count_covering_rowspans(
    rows: &[TableRow],
    declared_count: usize,
) -> usize {
    let mut covered_count: usize = declared_count.min(rows.len());
    let mut row_index: usize = 0;
    while row_index < covered_count {
        for cell in &rows[row_index].cells {
            covered_count = covered_count
                .max(row_index.saturating_add(cell.row_span as usize))
                .min(rows.len());
        }
        row_index += 1;
    }
    covered_count
}

#[allow(clippy::too_many_arguments)]
fn generate_table_rows(
    out: &mut String,
    rows: &[TableRow],
    painted_borders: Option<&[Vec<Option<CellBorder>>]>,
    background_bleed_top_trims: Option<&[Vec<f64>]>,
    background_bleed_bottom_left_trims: Option<&[Vec<f64>]>,
    boundary_band_model: Option<TableBorderPaintModel>,
    // The table's declared column widths, in points, so each cell can bound
    // how wide a framed eojeol may be (issue #626). Empty when the table
    // declares none.
    column_widths: &[f64],
    num_cols: usize,
    rowspan_remaining: &mut [usize],
    indent: &str,
    default_cell_padding: Insets,
    fixed_row_heights: bool,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // A nested table decides its own rows; restore the enclosing row's answer
    // so the outer cells that follow keep sharing their baseline.
    let enclosing_row_east_asian: RowEastAsianMetrics = ctx.row_east_asian;
    for (row_index, row) in rows.iter().enumerate() {
        for rs in rowspan_remaining.iter_mut() {
            if *rs > 0 {
                *rs -= 1;
            }
        }

        // Word sizes a row's lines from the whole row: if any cell holds East
        // Asian text, every cell in it takes the East Asian line height, and a
        // snapping grid applies to all of them. Asking each cell separately
        // split mixed-script rows across two baselines (issue #498).
        //
        // The line box itself keys on the face, not the text: a native export
        // with twelve `10_research_report_ko` rows relabelled `2025-01`..
        // `2025-12` — Latin-only rows in an all-slots Malgun Gothic face —
        // keeps Word's 25.44pt row pitch where the bare hhea line would give
        // 21.64pt (issue #814).
        //
        // A **spreadsheet** row takes no East Asian box at all, whatever its
        // characters. A native Excel-for-Mac export of the probe workbook
        // `tests/fixtures/xlsx/issue_1060_sheet_row_line_box_probe.xlsx`
        // prints its paired Korean and Latin-only blocks — same Malgun Gothic
        // face, size, row-height mode and vertical alignment, differing only
        // in script — 0.00pt apart on all four pairs, and prints the 14pt auto
        // rows at a 20.00pt track that Word's 24.20pt East Asian box does not
        // fit. So the row's answer must not vary with the text (it did:
        // a Korean top-aligned row seated 2.79pt below its Latin twin), and
        // the value it must not vary to is the bare hhea line (issue #1060).
        let has_east_asian_text: bool = row_has_east_asian_text(row);
        ctx.row_east_asian = RowEastAsianMetrics {
            has_east_asian_text,
            takes_east_asian_metrics: !ctx.table_seats_bottom_aligned_text_on_descender
                && (has_east_asian_text || row_is_set_in_east_asian_face(row)),
        };

        // A spreadsheet row whose fixed track cannot hold more than one line
        // seats every cell on that one line, whatever each cell's declared
        // vertical alignment says — Excel's native exports print such rows on
        // a single baseline (issue #839).
        let row_shared_line: Option<SheetRowLine> = sheet_row_shared_line(
            row,
            row.height.filter(|_| fixed_row_heights),
            ctx.row_east_asian,
            default_cell_padding,
            ctx.table_seats_bottom_aligned_text_on_descender,
        );

        // The auto-row frame estimate walks every cell's font metrics, so
        // computing it per cell is O(cells²) on wrap-text rows; it depends
        // only on the row, so it is computed lazily at most once per row and
        // shared by every cell that needs a vertical band length (issue #619
        // review, remediation 3).
        let mut row_frame_estimate_cache: Option<Option<f64>> = None;
        let mut col_pos: usize = 0;
        for (cell_index, cell) in row.cells.iter().enumerate() {
            if cell.col_span == 0 || cell.row_span == 0 {
                continue;
            }

            while col_pos < num_cols && rowspan_remaining[col_pos] > 0 {
                col_pos += 1;
            }
            if col_pos >= num_cols {
                break;
            }

            let remaining = num_cols - col_pos;
            let clamped_colspan = (cell.col_span as usize).min(remaining).max(1) as u32;
            // `Some` selects the boundary-band regime even when this cell
            // paints nothing; `None` keeps the stroke regime.
            let boundary_band: Option<BoundaryBandCell> =
                painted_borders.map(|p| BoundaryBandCell {
                    painted_border: &p[row_index][cell_index],
                    background_bleed_top_trim_pt: background_bleed_top_trims
                        .map_or(0.0, |trims| trims[row_index][cell_index]),
                    background_bleed_bottom_left_trim_pt: background_bleed_bottom_left_trims
                        .map_or(0.0, |trims| trims[row_index][cell_index]),
                    paint_model: boundary_band_model
                        .expect("painted borders require a boundary-band model"),
                    vertical_extent: vertical_band_extent(
                        rows,
                        row_index,
                        cell,
                        fixed_row_heights,
                        default_cell_padding,
                        ctx,
                        &mut row_frame_estimate_cache,
                    ),
                });
            // A cell's own text column: the columns it spans, less the inset
            // that keeps its text off the border (issue #626).
            let enclosing_measure_pt: Option<f64> = ctx.available_measure_pt;
            if !column_widths.is_empty() {
                let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
                let span_width_pt: f64 = column_widths
                    .iter()
                    .skip(col_pos)
                    .take(clamped_colspan as usize)
                    .sum();
                ctx.available_measure_pt =
                    Some(span_width_pt - inset.left - inset.right).filter(|measure| *measure > 0.0);
            }
            generate_table_cell(
                out,
                cell,
                boundary_band,
                clamped_colspan,
                indent,
                default_cell_padding,
                row.height.filter(|_| fixed_row_heights),
                fixed_spanned_row_height(rows, row_index, cell.row_span, fixed_row_heights),
                // The floor belongs to the row, and a Typst table row is as
                // tall as its tallest cell, so one strut carries it. Putting
                // it in every cell would only repeat the same constraint.
                row.minimum_height.filter(|_| cell_index == 0),
                row_shared_line.as_ref(),
                ctx,
            )?;
            ctx.available_measure_pt = enclosing_measure_pt;

            if cell.row_span > 1 {
                for rs in rowspan_remaining
                    .iter_mut()
                    .skip(col_pos)
                    .take(clamped_colspan as usize)
                {
                    *rs = cell.row_span as usize;
                }
            }
            col_pos += clamped_colspan as usize;
        }

        while col_pos < num_cols {
            if rowspan_remaining[col_pos] == 0 {
                let _ = writeln!(out, "{indent}[],");
            }
            col_pos += 1;
        }
    }
    ctx.row_east_asian = enclosing_row_east_asian;

    Ok(())
}

/// Whether any cell in the row carries East Asian text.
///
/// Nested tables are excluded: they run their own row loop and decide each of
/// their rows on their own content.
fn row_has_east_asian_text(row: &TableRow) -> bool {
    row.cells
        .iter()
        .flat_map(|cell| cell.content.iter())
        .any(block_has_east_asian_text)
}

fn block_has_east_asian_text(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => paragraph
            .runs
            .iter()
            .any(|run| run.text.chars().any(is_cjk_like)),
        Block::List(list) => {
            list.items
                .iter()
                .flat_map(|item| item.content.iter())
                .any(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .any(|run| run.text.chars().any(is_cjk_like))
                })
        }
        _ => false,
    }
}

/// Whether any line in the row resolves to an East Asian metric family, which
/// gives the whole row Word's East Asian line box even when every character is
/// Latin (issues #643, #814). Nested tables are excluded like
/// [`row_has_east_asian_text`]'s.
fn row_is_set_in_east_asian_face(row: &TableRow) -> bool {
    row.cells
        .iter()
        .flat_map(|cell| cell.content.iter())
        .any(block_is_set_in_east_asian_face)
}

fn block_is_set_in_east_asian_face(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => line_takes_east_asian_metrics(&paragraph.runs),
        Block::List(list) => list
            .items
            .iter()
            .flat_map(|item| item.content.iter())
            .any(|paragraph| line_takes_east_asian_metrics(&paragraph.runs)),
        _ => false,
    }
}

/// Excel's printed grid quantises row tracks to whole PDF points (see
/// `native_excel_pdf_row_height` in the XLSX parser), so a row's content box
/// can read up to half a point of slack the native export does not actually
/// have. The tightness gate absorbs that quantum before deciding a row has
/// room for per-cell vertical alignment (issue #839).
const SHEET_ROW_TRACK_QUANTISATION_SLACK_PT: f64 = 0.5;

/// The one line a tight spreadsheet row seats every cell on, or `None` when
/// the row is not in that regime (issue #839).
///
/// Excel prints every cell of a single-line sheet row on one baseline: the
/// native export of `09_expense_report_en` puts a `vertical="bottom"` amount
/// column and its `vertical="center"` neighbours all at y=143.00, and
/// `04_payroll_ko`'s fixed 합계 row seats its centred Korean label with its
/// bottom-aligned numbers at y=218.00. The alignments coincide because the
/// track holds essentially the line alone: there is no slack to distribute,
/// so the declared alignments have nowhere to differ.
///
/// Both gates key on the row's *bare hhea* line, not the 1.3-factor East
/// Asian box (#518): the tight regime is a property of Excel's geometry, and
/// Excel's own line never carries that Word factor — judging a 23pt Korean
/// title track against the inflated box would misread its real ~3pt of slack
/// as none. Two regimes stay out. A row with more content room than the line
/// — a tall header, a spanned merge — keeps per-cell alignment, which Excel
/// honours and #618 measured. A row whose track is *shorter* than the line
/// holds text deliberately oversized for it (a 42pt title in a 23pt track),
/// where the alignments pick which part of the overflowing line shows and
/// stay honoured too; a track an auto-fit produced is never shorter than its
/// own font's line.
fn sheet_row_shared_line(
    row: &TableRow,
    row_track_pt: Option<f64>,
    row_east_asian: RowEastAsianMetrics,
    default_cell_padding: Insets,
    table_seats_bottom_aligned_text_on_descender: bool,
) -> Option<SheetRowLine> {
    if !table_seats_bottom_aligned_text_on_descender {
        return None;
    }
    let track_pt: f64 = row_track_pt?;
    let metric_family: String = row_metric_family(row, row_east_asian.has_east_asian_text)?;
    let font_size_pt: f64 = row_font_size_pt(row);
    let bare_line_pt: f64 = sheet_row_line_advance_pt(&metric_family, font_size_pt, false)?;
    if track_pt + SHEET_ROW_TRACK_QUANTISATION_SLACK_PT < bare_line_pt {
        return None;
    }
    // The tallest content box among the cells this track alone holds; a cell
    // spanning several tracks has more room than the track and is judged out
    // of scope by its `row_span` in `generate_table_cell` anyway.
    let max_content_pt: f64 = row
        .cells
        .iter()
        .filter(|cell| cell.col_span > 0 && cell.row_span == 1)
        .map(|cell| {
            let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
            track_pt - inset.top - inset.bottom
        })
        .fold(f64::NAN, f64::max);
    // A NaN (no cells hold this track alone) fails the comparison and bails.
    let row_is_tight: bool = max_content_pt <= bare_line_pt + SHEET_ROW_TRACK_QUANTISATION_SLACK_PT;
    if !row_is_tight {
        return None;
    }
    Some(SheetRowLine {
        metric_family,
        font_size_pt,
    })
}

/// The family whose metrics pace a spreadsheet row's shared line: the
/// row-level mirror of `east_asian_aware_metric_family` (issue #839). A row
/// carrying East Asian text is paced by the family shaping that text — the
/// declared `east_asian_font_family` slot, or failing that the `font_family`
/// of a run that carries East Asian characters. A Latin row is paced by its
/// first declared family.
///
/// Which family paces the row keys on the *characters*, not on the row's line
/// box: a sheet row never takes an East Asian box (issue #1060), and reading
/// that answer here would have paced `김민준 | E-1021` on whichever cell came
/// first instead of on the face shaping the Hangul.
fn row_metric_family(row: &TableRow, has_east_asian_text: bool) -> Option<String> {
    let runs = || {
        row.cells
            .iter()
            .flat_map(|cell| cell.content.iter())
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph.runs.as_slice()),
                _ => None,
            })
            .flatten()
    };
    let run: &Run = if has_east_asian_text {
        runs()
            .find(|run| {
                run.text.chars().any(is_cjk_like)
                    && (run.style.font_family.is_some()
                        || run.style.east_asian_font_family.is_some())
            })
            .or_else(|| runs().find(|run| run.style.east_asian_font_family.is_some()))?
    } else {
        runs().find(|run| run.style.font_family.is_some())?
    };
    let latin_family: &str = run
        .style
        .font_family
        .as_deref()
        .or(run.style.east_asian_font_family.as_deref())?;
    Some(crate::render::font_subst::painted_family_for_text(
        latin_family,
        run.style.east_asian_font_family.as_deref(),
        &run.text,
    ))
}

/// The size a spreadsheet row's shared line resolves at: the largest run size
/// in the row, mirroring `paragraph_font_size_pt`'s largest-run rule at row
/// scope (issue #839).
fn row_font_size_pt(row: &TableRow) -> f64 {
    let largest: f64 = row
        .cells
        .iter()
        .flat_map(|cell| cell.content.iter())
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph.runs.as_slice()),
            _ => None,
        })
        .flatten()
        .filter_map(|run| run.style.font_size)
        .fold(f64::NAN, f64::max);
    if largest.is_nan() { 11.0 } else { largest }
}

/// Excel does not fill the cell with a data bar: it insets the bar from the
/// row's top and bottom edges. Native Excel PDF exports of the business corpus
/// print a 10 pt bar in every 14 pt row, which is 2 pt of clearance per side.
const DATA_BAR_VERTICAL_INSET_PT: f64 = 2.0;

/// Floor for rows shorter than the inset, so a bar never vanishes or inverts.
const DATA_BAR_MIN_HEIGHT_PT: f64 = 1.0;

/// Horizontal clearance between an Excel data bar and its cell boundaries.
///
/// This is independent of the text inset. Native Excel exports start the bars
/// about 2pt inside the left boundary and leave about 1pt on the right. With a
/// 3pt/side text inset, resolving the percentage against the text box made the
/// track 3pt too narrow and started it 1pt too far right (issue #655).
const DATA_BAR_LEFT_INSET_PT: f64 = 2.0;
const DATA_BAR_RIGHT_INSET_PT: f64 = 1.0;
/// Where an icon-set icon starts, measured from its cell's left boundary.
///
/// Excel anchors the icon itself, not the text: in the native export of
/// `10_kpi_tracker_en` column E spans x 384-456pt and every `3Arrows` sprite
/// is placed `transform="11 0 0 11 386 …"` — 2pt in — while the column is
/// centred, so the alignment moves the value and leaves the icon alone. That
/// 2pt matches `DATA_BAR_LEFT_INSET_PT` but is a separate measurement of a
/// separate feature; they are kept apart so a later correction to one does
/// not silently move the other (issue #1087).
///
/// Extracting those sprites puts the arrow's ink flush with the placement
/// box's left edge — the up and down arrows trim to 11x12 of a 12x12 bitmap
/// at +0+0, the padding #651 measured falling on the right — so the drawn
/// polygon's own left edge belongs at this inset. The circle sets have no
/// ground truth in this corpus and share the anchor.
const ICON_SET_LEFT_INSET_PT: f64 = 2.0;

/// The side of the square box Excel prints an icon-set sprite in, in points.
///
/// Every `3Arrows` sprite of the native `10_kpi_tracker_en` export is placed
/// `transform="11 0 0 11 386 …"`, carrying a 12 x 12px bitmap. Where the ink
/// sits inside that box is the sprite's own business: the up and down arrows
/// fill it, the right one leaves its bottom row blank (issue #651).
const ICON_SET_BOX_SIZE_PT: f64 = 11.0;

/// How far below its row's top boundary Excel seats that box, in points.
///
/// Measured on the same export (issue #1202): the six sprites sit at
/// y = 133, 147, 161, 175, 189 and 203, and the thin row rules over them fill
/// 132-133, 146-147, … — Excel's boundary-anchored `[B, B + 1]` band, so the
/// boundaries are 132, 146, … and every box starts one whole point below one.
/// The 14pt tracks do not centre it: an 11pt box centred there would start
/// 1.5pt down, and centring the *ink* — which is what `horizon` did — put the
/// vertical arrows 0.25pt low and the right one, 0.92pt short of its box, a
/// further 0.46pt.
const ICON_SET_TOP_INSET_PT: f64 = 1.0;

/// Excel's arrow icon sets are drawn shapes, not characters. Native Excel PDFs
/// print them as sprites, ramped along the icon box's diagonal under a flat
/// outline; these constants size the vector stand-in, whose paint comes from
/// the band's [`IconShading`] where one was measured.
///
/// Measured from the Excel export of `10_kpi_tracker_en`: the sheet places six
/// 11 x 11pt `fill_image` sprites, but that is the placement box. Extracting
/// them gives 12 x 12px bitmaps whose non-white ink spans 11 x 12px for the up
/// arrow and 12 x 11px for the right one — 10.08 x 11.00pt of actual arrow,
/// with about a pixel of padding on the narrow axis. Sizing to the 11 x 11 box
/// instead would give the ink the size of the whole sprite (issue #651).
///
/// Along its length the arrow does span the whole box, so this is
/// [`ICON_SET_BOX_SIZE_PT`] itself; only the breadth is short of it.
const ARROW_ICON_LENGTH_PT: f64 = ICON_SET_BOX_SIZE_PT;
/// Across the shaft the arrow is narrower than it is long.
const ARROW_ICON_BREADTH_PT: f64 = 10.08;

/// How the arrow's silhouette divides that box, as fractions of the ink the
/// sprite's soft mask carries.
///
/// The mask *is* the silhouette, so it gives the split directly. The up
/// arrow's is a 12 x 12px bitmap with 11 x 12px of ink:
///
/// ```text
///  0 .....#......
///  1 ....###.....
///  2 ...#####....
///  3 ..#######...
///  4 .#########..
///  5 ###########.
///  6 ...#####....
///  … ...#####....
/// 11 ...#####....
/// ```
///
/// The head occupies rows 0-5, half the length; the shaft columns 3-7, 5 of
/// the 11 ink columns. The hand-picked fractions these replace — 0.45 of the
/// length, and a 0.28 half-width, so 0.56 of the breadth — gave the shaft 23%
/// too much breadth under a head 10% too short, which read chunkier than the
/// native arrow at any resolution that resolves the shaft (issue #1135). The right
/// arrow's mask is this silhouette transposed, which is how the polygon is
/// built, so the one split covers all five orientations.
const ARROW_ICON_SHAFT_BREADTH_FRACTION: f64 = 5.0 / 11.0;
const ARROW_ICON_HEAD_LENGTH_FRACTION: f64 = 6.0 / 12.0;

/// Diameter of a circular icon-set icon, in points.
///
/// Measured from Excel's export of the audited workbook: 6.72pt printed at
/// that sheet's 75% scale, so 8.96pt at 100%. The `●` character it used to
/// print is a little over half that (#536).
const CIRCLE_ICON_DIAMETER_PT: f64 = 8.96;

/// How wide Excel draws an arrow's outline, in points.
///
/// The sprite's outline is exactly one of its own pixels wide everywhere and
/// lies wholly inside the silhouette: in `10_kpi_tracker_en`'s up arrow the
/// bottom row is the outline hue across, the shaft's side columns likewise,
/// and the interior ramp starts one pixel in. The bitmap is 12px over the
/// [`ARROW_ICON_LENGTH_PT`] box, so that pixel is 0.917pt (issue #1201).
const ARROW_ICON_OUTLINE_WIDTH_PT: f64 = ARROW_ICON_LENGTH_PT / 12.0;

/// Excel shades an arrow sprite along the icon box's diagonal, not down it:
/// every interior pixel is a function of `x + y`, and the sprite's pixel is
/// square, so the ramp runs at 45 degrees (issue #1134).
const ARROW_ICON_GRADIENT_ANGLE_DEG: i32 = 45;

/// Seat an icon whose ink is shorter than Excel's sprite box in the middle of
/// that box, so every set shares one placement.
///
/// The box is what the row seats, and only the arrows' masks say their ink
/// hangs from its top. A set with no native export to read keeps the middle,
/// which is where a shape drawn in a square sprite lands (issue #1202).
fn icon_sprite_box(content: &str) -> String {
    format!(
        "box(height: {}pt)[#place(left + horizon, {content})]",
        format_f64(ICON_SET_BOX_SIZE_PT),
    )
}

/// The drawn shape for an icon-set glyph, or `None` for the sets that stay
/// characters — symbols, flags, stars.
fn icon_shape(glyph: &str, color: Option<Color>, shading: Option<IconShading>) -> Option<String> {
    if glyph == crate::ir::ICON_CIRCLE {
        let radius: f64 = CIRCLE_ICON_DIAMETER_PT / 2.0;
        let paint: String = color
            .map(|c| rgb(&c))
            .unwrap_or_else(|| "black".to_string());
        return Some(icon_sprite_box(&format!(
            "circle(radius: {}pt, fill: {paint}, stroke: none)",
            format_f64(radius)
        )));
    }
    arrow_icon_polygon(glyph, color, shading)
}

/// Build the Typst `polygon` for one of the arrow icon-set glyphs, or `None`
/// for any other glyph.
fn arrow_icon_polygon(
    glyph: &str,
    color: Option<Color>,
    shading: Option<IconShading>,
) -> Option<String> {
    // The head spans the full breadth; `shaft` is the shaft's half-width and
    // `neck` is where the head meets it, measured from the tip.
    let breadth: f64 = ARROW_ICON_BREADTH_PT;
    let length: f64 = ARROW_ICON_LENGTH_PT;
    let shaft: f64 = breadth * ARROW_ICON_SHAFT_BREADTH_FRACTION / 2.0;
    let neck: f64 = length * ARROW_ICON_HEAD_LENGTH_FRACTION;

    // Points of an up arrow, clockwise from the tip.
    let up: Vec<(f64, f64)> = vec![
        (breadth / 2.0, 0.0),
        (breadth, neck),
        (breadth / 2.0 + shaft, neck),
        (breadth / 2.0 + shaft, length),
        (breadth / 2.0 - shaft, length),
        (breadth / 2.0 - shaft, neck),
        (0.0, neck),
    ];
    let flip_y = |points: &[(f64, f64)]| -> Vec<(f64, f64)> {
        points.iter().map(|(x, y)| (*x, length - *y)).collect()
    };
    let transpose = |points: &[(f64, f64)]| -> Vec<(f64, f64)> {
        points.iter().map(|(x, y)| (length - *y, *x)).collect()
    };

    let (points, rotation): (Vec<(f64, f64)>, Option<i32>) = match glyph {
        crate::ir::ICON_ARROW_UP => (up, None),
        crate::ir::ICON_ARROW_DOWN => (flip_y(&up), None),
        crate::ir::ICON_ARROW_RIGHT => (transpose(&up), None),
        crate::ir::ICON_ARROW_UP_RIGHT => (up, Some(45)),
        crate::ir::ICON_ARROW_DOWN_RIGHT => (flip_y(&up), Some(-45)),
        _ => return None,
    };

    let coordinates = |points: &[(f64, f64)]| -> String {
        points
            .iter()
            .map(|(x, y)| format!("({}pt, {}pt)", format_f64(*x), format_f64(*y)))
            .collect::<Vec<String>>()
            .join(", ")
    };
    // A band Excel exported carries its own ramp and its own outline hue; one
    // that was never measured keeps the flat stand-in under an outline derived
    // from it, which is close for green and olive for amber (issue #1134).
    let (fill, outline): (String, String) = match shading {
        Some(shading) => (
            format!(
                "gradient.linear(angle: {ARROW_ICON_GRADIENT_ANGLE_DEG}deg, space: rgb, {}, {})",
                rgb(&shading.fill_start),
                rgb(&shading.fill_end),
            ),
            rgb(&shading.outline),
        ),
        None => {
            let paint: String = color
                .map(|c| rgb(&c))
                .unwrap_or_else(|| "black".to_string());
            (paint.clone(), format!("{paint}.darken(30%)"))
        }
    };
    // Typst has no inset stroke, so the outline is stroked on a path inset by
    // half its width: mitring the corners that way puts the stroke's outer
    // edge back on the silhouette, with none of it over the page. Stroking
    // the silhouette itself painted less than a quarter of Excel's outline
    // inward and spilled the rest outward (issue #1201).
    let ring: Vec<(f64, f64)> = inset_polygon(&points, ARROW_ICON_OUTLINE_WIDTH_PT / 2.0);
    // The ramp stays on the silhouette rather than the inset path: a 45deg
    // gradient runs over its own shape's box, and the fractions #1134 measured
    // are of the sprite's full box, not of what the outline leaves (issue
    // #1201).
    //
    // Superimposing the two takes a box, and it is the sprite's own 11 x 11pt
    // one rather than the silhouette's extent. The silhouette is flush with
    // its top-left corner already — the padding column of the up and down
    // masks falls on the right and the padding row of the transposed one at
    // the bottom (#651, #1135) — so seating the box seats the ink, and a
    // sprite short of the box keeps its ink against the top instead of being
    // re-centred there (issue #1202).
    let shape: String = format!(
        "box(width: {size}pt, height: {size}pt)\
         [#place(top + left, polygon(fill: {fill}, stroke: none, {silhouette}))\
         #place(top + left, polygon(stroke: {stroke}pt + {outline}, {ring}))]",
        size = format_f64(ICON_SET_BOX_SIZE_PT),
        silhouette = coordinates(&points),
        stroke = format_f64(ARROW_ICON_OUTLINE_WIDTH_PT),
        ring = coordinates(&ring),
    );
    Some(match rotation {
        Some(degrees) => format!("rotate({degrees}deg, {shape})"),
        None => shape,
    })
}

/// A closed polygon's vertices moved `distance` toward its interior, each
/// corner landing where its two neighbouring offset edges meet.
///
/// Offsetting the result back out by the same distance reproduces the input,
/// which is what makes a `2 * distance` stroke on it cover exactly the input's
/// outermost band and nothing beyond it. The interior's side of each edge is
/// read from the signed area rather than assumed, so the flipped and
/// transposed arrows inset inward too.
fn inset_polygon(points: &[(f64, f64)], distance: f64) -> Vec<(f64, f64)> {
    let count: usize = points.len();
    let twice_signed_area: f64 = (0..count)
        .map(|index| {
            let (x0, y0) = points[index];
            let (x1, y1) = points[(index + 1) % count];
            x0 * y1 - x1 * y0
        })
        .sum();
    let winding: f64 = if twice_signed_area > 0.0 { 1.0 } else { -1.0 };
    // The inward unit normal of the edge leaving vertex `index`.
    let normal = |index: usize| -> (f64, f64) {
        let (x0, y0) = points[index];
        let (x1, y1) = points[(index + 1) % count];
        let (dx, dy) = (x1 - x0, y1 - y0);
        let length: f64 = dx.hypot(dy);
        (-dy / length * winding, dx / length * winding)
    };
    (0..count)
        .map(|index| {
            let (x, y) = points[index];
            let (arriving_x, arriving_y) = normal((index + count - 1) % count);
            let (leaving_x, leaving_y) = normal(index);
            // Both offset edges are `distance` along their own normal, so
            // their meeting point lies on the normals' sum, scaled to keep
            // that distance. A straight-through vertex has the two normals
            // equal, which the same expression carries.
            let (sum_x, sum_y) = (arriving_x + leaving_x, arriving_y + leaving_y);
            let projection: f64 = sum_x * arriving_x + sum_y * arriving_y;
            if projection.abs() < 1e-9 {
                // An edge doubling straight back on itself has no interior to
                // move into; leave the vertex where it is.
                return (x, y);
            }
            let scale: f64 = distance / projection;
            (x + sum_x * scale, y + sum_y * scale)
        })
        .collect()
}

#[allow(clippy::too_many_arguments)]
fn generate_table_cell(
    out: &mut String,
    cell: &TableCell,
    boundary_band: Option<BoundaryBandCell>,
    clamped_colspan: u32,
    indent: &str,
    default_cell_padding: Insets,
    row_height: Option<f64>,
    cell_track_height: Option<f64>,
    row_minimum_height: Option<f64>,
    row_shared_line: Option<&SheetRowLine>,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // Whether this cell joins its tight row's one baseline (issue #839). A
    // cell spanning several tracks has more room than the row's single track,
    // and a cell stacking several blocks holds more than the row's one line;
    // both keep their declared alignment, which is what Excel honours when
    // there is room. A cell with no paragraph has no line to seat.
    let seats_on_row_line: bool = row_shared_line.is_some()
        && cell.row_span <= 1
        && cell
            .content
            .iter()
            .filter(|block| {
                !matches!(
                    block,
                    Block::TableOfContents(_) | Block::PageBreak | Block::ColumnBreak
                )
            })
            .count()
            <= 1
        && cell
            .content
            .iter()
            .any(|block| matches!(block, Block::Paragraph(_)));

    let needs_cell_fn = clamped_colspan > 1
        || cell.row_span > 1
        || cell.border.is_some()
        || cell.background.is_some()
        || cell.vertical_align.is_some()
        || cell.padding.is_some()
        || seats_on_row_line;

    // The alignment the cell actually renders with: its own, or the table's
    // default (Excel's bottom). The paragraph codegen needs the effective
    // answer, not the cell's declaration, because Excel's untouched default
    // cells are exactly the bottom-aligned ones (issue #618).
    //
    // In a tight spreadsheet row the declared alignment has no room to act in
    // Excel, so the choice of anchor here is free — and it is taken as the
    // *centred* symmetric box for every cell, not the descender seat. The
    // centred baseline — `sheet_cell_baseline_from_track_top_pt`'s rounded,
    // track-centred seat since issue #1063 — carries no dependence on the
    // East Asian 1.3 line factor at all, so it is immune to that factor's
    // known overshoot (#709); the bottom seat inherits the row-track error in
    // full. Measured on the business corpus
    // baseline gate, centring is the anchor that moves every deviating cell
    // toward its GT — the descender seat moved every Korean page 1.2–1.8pt
    // further away (issue #839).
    let effective_vertical_align: Option<CellVerticalAlign> = if seats_on_row_line {
        Some(CellVerticalAlign::Center)
    } else {
        cell.vertical_align.or(ctx.table_default_vertical_align)
    };
    let enclosing_cell_seats_on_descender: bool = ctx.cell_seats_text_on_descender;
    let enclosing_cell_vertical_align: Option<CellVerticalAlign> = ctx.cell_vertical_align;
    let enclosing_cell_sheet_row_line: Option<SheetRowLine> = ctx.cell_sheet_row_line.take();
    let enclosing_cell_sheet_seat: Option<SheetCellSeat> = ctx.cell_sheet_seat.take();
    // The fixed sheet track this cell seats its line in (issue #1063). A
    // centred multi-row merge uses the full joined track: Excel rounds and
    // centres the text block against that one declared-space height, whereas
    // Typst otherwise centres from the paragraph's own metric edges (#1497).
    // Bottom-aligned multi-row cells keep the legacy answer because their
    // measured descender seat belongs to a single row; top alignment is also
    // outside the measured regime. Auto rows have no fixed grid track at all.
    ctx.cell_sheet_seat = cell_track_height
        .filter(|_| ctx.table_seats_bottom_aligned_text_on_descender)
        .filter(|_| {
            cell.row_span <= 1 || effective_vertical_align == Some(CellVerticalAlign::Center)
        })
        .filter(|_| effective_vertical_align != Some(CellVerticalAlign::Top))
        .map(|track_pt| {
            let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
            SheetCellSeat {
                track_pt,
                inset_top_pt: inset.top,
                inset_bottom_pt: inset.bottom,
                descent_floor_pt: ctx.table_bottom_aligned_descent_floor_pt,
                is_horizontally_merged: cell.col_span > 1,
                is_centered_multi_row: cell.row_span > 1
                    && effective_vertical_align == Some(CellVerticalAlign::Center),
            }
        });
    ctx.cell_sheet_row_line = row_shared_line.filter(|_| seats_on_row_line).cloned();
    // Descender seating applies only to FIXED-height rows (`row_height` is
    // `Some` only then). In auto rows the renderer sizes the row from the
    // content itself, whose intrinsic height was calibrated against Excel GT
    // (#396/#411/#498) with the symmetric box; only fixed rows have slack for
    // alignment to distribute, and only they were measured in #618.
    ctx.cell_seats_text_on_descender = ctx.table_seats_bottom_aligned_text_on_descender
        && effective_vertical_align == Some(CellVerticalAlign::Bottom)
        && row_height.is_some();
    ctx.cell_vertical_align = effective_vertical_align;

    let paints_boundary_bands: bool = boundary_band.is_some();

    if needs_cell_fn {
        out.push_str(indent);
        out.push_str("table.cell(");
        write_cell_params(
            out,
            cell,
            clamped_colspan,
            default_cell_padding,
            paints_boundary_bands,
            seats_on_row_line.then_some(CellVerticalAlign::Center),
        );
        out.push_str(")[");
    } else {
        out.push_str(indent);
        out.push('[');
    }

    // The `#align(...)` that places an aligned table's box is inherited by
    // everything inside it, so a cell paragraph that declares no alignment of
    // its own would be laid out centred or right. Reset it at the cell, where
    // a paragraph's own `#set align(...)` still nests deeper and wins
    // (issue #843).
    if ctx.table_box_is_aligned {
        out.push_str("#set align(start)\n");
    }

    if let Some(band) = &boundary_band {
        let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
        // The shading goes down before the bands, so a border still paints
        // over the strip the two share.
        if band.paint_model == TableBorderPaintModel::ExcelBoundaryBands
            && let Some(background) = &cell.background
        {
            write_excel_background_bleed(
                out,
                background,
                cell.background_alpha,
                inset,
                &band.vertical_extent,
                (
                    band.background_bleed_top_trim_pt,
                    band.background_bleed_bottom_left_trim_pt,
                ),
                ctx.sheet_print_scale(),
            );
        }
        if let Some(border) = band.painted_border {
            match band.paint_model {
                TableBorderPaintModel::ExcelBoundaryBands => {
                    write_boundary_anchored_border_overlays(
                        out,
                        border,
                        inset,
                        &band.vertical_extent,
                    );
                }
                TableBorderPaintModel::WordPositiveAxisBands => {
                    write_word_positive_axis_border_overlays(
                        out,
                        border,
                        inset,
                        &band.vertical_extent,
                    );
                }
                TableBorderPaintModel::CenteredStroke => {
                    unreachable!("centred strokes do not select the boundary-band path")
                }
            }
        }
    } else if let Some(border) = &cell.border {
        write_double_border_overlays(out, border, cell.padding.unwrap_or(default_cell_padding));
    }

    if let Some(ref db) = cell.data_bar {
        // Excel draws the bar behind the value on the same line (no track),
        // with a horizontal fade of the bar color; #place keeps it out of
        // layout so the value renders on top at its normal position. The bar
        // height must be concrete: in auto-height rows a relative height has
        // no cell frame to resolve against and blows up to the page height,
        // smearing over neighboring rows (issue #362).
        //
        // Where Excel's fade ends was read off its own export of
        // `06_sales_dashboard_en` rather than picked (issue #654). Sampling
        // along two bars and fitting gives a straight ramp to 0.84 of the way
        // to white, worst residual under 1.2%.
        //
        // 83 rather than 84 because the fit is of the *rendered* page, and our
        // own rendering reads a little light: the 70% this replaced measured
        // 0.706 back. 84% was tried and landed about three levels past Excel
        // at the bar's tail; 83% reproduces every sampled pixel on both bars
        // to within one level. The earlier 70% stopped short of Excel
        // altogether and left brief bars reading near-solid.
        let pct = db.fill_pct.clamp(0.0, 100.0);
        let cell_inset: Insets = cell_inset_with_border(cell, default_cell_padding);
        let (bar_dx, bar_width): (f64, String) = match ctx.available_measure_pt {
            Some(content_width) => {
                // `available_measure_pt` is the column span less the text
                // inset. Restore the full span, apply Excel's independent bar
                // inset, then quantise the painted width to whole PDF points.
                // This reproduces all 5 bars in `06_sales_dashboard_en` and
                // all 45 page-1 bars in `03_inventory_en` exactly (#655).
                let track_width = (content_width + cell_inset.left + cell_inset.right
                    - DATA_BAR_LEFT_INSET_PT
                    - DATA_BAR_RIGHT_INSET_PT)
                    .max(0.0);
                let painted_width = (track_width * pct / 100.0).round();
                (
                    DATA_BAR_LEFT_INSET_PT - cell_inset.left,
                    format!("{}pt", format_f64(painted_width)),
                )
            }
            None => (0.0, format!("{}%", format_f64(pct))),
        };
        let bar_height: String = match row_height {
            Some(height) => {
                let inset_height =
                    (height - 2.0 * DATA_BAR_VERTICAL_INSET_PT).max(DATA_BAR_MIN_HEIGHT_PT);
                format!("{}pt", format_f64(inset_height))
            }
            // Excel sizes default rows to the font's line box; 1.2em tracks
            // that for single-line numeric cells, less the same inset.
            None => format!("1.2em - {}pt", format_f64(2.0 * DATA_BAR_VERTICAL_INSET_PT)),
        };
        let _ = write!(
            out,
            "#place(left + horizon, dx: {}pt, box(width: {}, height: {}, fill: gradient.linear({}, {}.lighten(83%))))",
            format_geometry(bar_dx),
            bar_width,
            bar_height,
            rgb(&db.color),
            rgb(&db.color),
        );
    }

    if let Some(ref icon) = cell.icon_text {
        // Excel draws icon set glyphs in their band color, independent of
        // the cell's font color, anchored at the cell's left edge on the
        // value's own line. Placing the icon out of layout keeps narrow
        // cells from wrapping the value onto a second line, which doubled
        // the row height (issue #367). Because it takes no width here, the
        // cell carries `ICON_SET_VALUE_RESERVE_PT` of extra left inset so the
        // value still aligns to the icon's right, as Excel does (issue #652).
        // Excel's arrow sets are drawn shapes rather than characters: a shaft
        // with a triangular head, outlined and filling most of the row. The
        // triangle characters the parser records are only a third that size,
        // so arrows are re-drawn as polygons.
        // The circle sets are drawn discs for the same reason (#536).
        //
        // `#place` starts at the *content* box, which the cell's inset — and,
        // on the left, that value reserve — has already pushed inward, so both
        // offsets are the cell's own inset undone and Excel's put in its
        // place. Left alone, the icon rides along with the value and the gap
        // Excel leaves between the two collapses (issue #1087).
        //
        // Vertically the anchor is the row's top boundary, not its centre
        // line: Excel seats the sprite's box `ICON_SET_TOP_INSET_PT` below the
        // boundary whatever the track's height, and the ink hangs inside that
        // box wherever the sprite puts it (issue #1202).
        let cell_inset: Insets = cell_inset_with_border(cell, default_cell_padding);
        let icon_dx: f64 = ICON_SET_LEFT_INSET_PT - cell_inset.left;
        let icon_dy: f64 = ICON_SET_TOP_INSET_PT - cell_inset.top;
        let anchor: String = format!(
            "#place(top + left, dx: {}pt, dy: {}pt, ",
            format_geometry(icon_dx),
            format_geometry(icon_dy),
        );
        match (
            icon_shape(icon, cell.icon_color, cell.icon_shading),
            cell.icon_color,
        ) {
            (Some(polygon), _) => {
                let _ = write!(out, "{anchor}{polygon})");
            }
            (None, Some(color)) => {
                let _ = write!(
                    out,
                    "{anchor}{})",
                    icon_sprite_box(&format!(
                        "text(fill: {}, weight: \"bold\")[{}]",
                        rgb(&color),
                        icon
                    ))
                );
            }
            (None, None) => {
                let _ = write!(
                    out,
                    "{anchor}{})",
                    icon_sprite_box(&format!("text(weight: \"bold\")[{icon}]"))
                );
            }
        }
    }

    if let Some(sparkline) = &cell.sparkline {
        generate_sparkline_overlay(
            out,
            sparkline,
            cell,
            default_cell_padding,
            row_height,
            cell_track_height,
            ctx,
        );
    }

    // Writer's positive-axis origin and Excel's filled-merge centring are
    // visual seats, not new layout measure. Keeping them as translations
    // preserves row pitch, column width, margins, and wrapping (#649, #1488,
    // #1493).
    let content_shift: Option<(f64, f64)> = word_cell_content_shift(&boundary_band)
        .or_else(|| excel_merged_cell_content_shift(&boundary_band, cell));
    let wraps_content_shift: bool = content_shift.is_some() && cell.spill_width.is_none();
    if wraps_content_shift {
        let (dx, dy) = content_shift.expect("the content-shift wrapper requires a seat");
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[",
            format_geometry(dx),
            format_geometry(dy),
        );
    }

    if let Some(spill_width) = cell.spill_width {
        // An unwrapped cell keeps its text on one line: lay the content out in
        // a clipped box via #place (out of layout) and hold the row height with
        // a zero-width strut, so the line does not grow the row.
        //
        // The clip box does *not* keep the line unwrapped — a Typst box breaks
        // its content at the width it states, and this one states one. What
        // holds the line together is the inner box sized from `measure()`
        // further down; see the comment there (issue #811).
        //
        // The box is anchored where the cell's own alignment puts it. A
        // general/left cell paints rightwards across empty neighbours from its
        // left edge; a centred or right-aligned one is clipped at its own edge,
        // so anchoring it left would slide its text out of the column
        // (issue #615).
        let horizontal_alignment: Option<Alignment> = cell_horizontal_alignment(cell);
        let anchor = match horizontal_alignment {
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            _ => "left",
        };
        // Excel cuts the line on the cell's own gridline, and `#place` starts
        // at the *content* box the inset has already pushed in. A box given
        // the whole spill width therefore overhangs that gridline by the inset
        // on the side it is anchored to, and the overhang admits glyphs Excel
        // does not print: `Wrapping paper` blocked in a 65pt column printed
        // `Wrapping pap` against the native export's `Wrapping pa`, the extra
        // `p` fitting in exactly the 3pt of left inset (issue #1105).
        //
        // Probe-measured on a native Excel for Mac export (5pt and 12pt runs
        // blocked left and right in one 65pt column): the last glyph Excel
        // draws is the last one that *starts* before the gridline, so a glyph
        // may overhang it but one starting past it is dropped. Clipping at the
        // inset content edge instead would have cut two glyphs earlier than
        // the export does, so the gridline is the boundary, not the inset.
        //
        // Nothing is lost on the anchored side: the line starts at that same
        // content edge, so the width given up is width the text never occupies.
        let spill_inset: Insets = cell.padding.unwrap_or(default_cell_padding);
        let clip_width_pt: f64 = (spill_width
            - match horizontal_alignment {
                // A centred box sits on the content box's centre, which is the
                // cell's own centre while the two insets match. Only the
                // difference between them offsets it — as on an icon-set cell,
                // whose left inset carries the icon's reserve (issue #652) —
                // and both sides have to give that up to stay inside the cell.
                Some(Alignment::Center) => (spill_inset.left - spill_inset.right).abs(),
                Some(Alignment::Right) => spill_inset.right,
                _ => spill_inset.left,
            })
        .max(0.0);
        // `#place` ignores the table's `align:`, so the wrapper must anchor
        // where the cell's effective vertical alignment puts the line. The
        // hardcoded `horizon` centred bottom-aligned titles in tall rows
        // (issue #618). A bottom anchor needs the box and strut sized from
        // the paragraph's own line box at the run's font size. The bottom
        // anchor applies only to FIXED-height rows — auto rows are
        // content-sized against the legacy shape (see the seating gate above)
        // — and top-aligned seating is unverified against Excel GT and out of
        // #618's measured scope, so Top shares Centred's `horizon` anchor.
        let vertical_anchor: &str = match effective_vertical_align {
            Some(CellVerticalAlign::Bottom) if row_height.is_some() => "bottom",
            _ => "horizon",
        };
        // Every anchor sizes its clip box from the cell's own line. `1.3em`
        // resolves against the *ambient* text size, so a cell set larger than
        // its surroundings was clipped mid-glyph: an 18.9pt title on an 11pt
        // sheet got a 14.30pt box against the 21.74pt its glyphs span, cutting
        // every descender off flat at the baseline (issue #927). The anchor
        // itself is unchanged — #618 measured the centred position correct,
        // and only the box around it was wrong.
        let line_box_height_pt: Option<f64> = spill_line_box_height_pt(cell, ctx);
        // The clip box states a width, and a Typst box **wraps** its content
        // at the width it states. The line therefore broke into several, the
        // clip hid all but one of them, and the one left visible was the tail:
        // a merged title rendered starting mid-sentence, with its opening
        // words gone (issue #811).
        //
        // Binding the content and sizing an inner box to `measure()`'s answer
        // is what keeps it on one line — measure lays out in an unbounded
        // region, so the inner box is the text's natural width and has nothing
        // to break at. The clip then cuts that single line at the spill edge,
        // which is where Excel cuts it, and each anchor keeps the fragment
        // Excel leaves visible: the head for a left cell, the tail for a right
        // one, the middle for a centred one.
        let height: String = match line_box_height_pt {
            Some(height_pt) => format!("{}pt", format_f64(height_pt)),
            // Unknown font metrics: keep the legacy ambient-sized shape.
            None => "1.3em".to_string(),
        };
        out.push_str("#context {let o2p-spill = [");
        let enclosing_in_spill_cell = ctx.in_spill_cell;
        ctx.in_spill_cell = true;
        let spill_content = generate_sheet_cell_content(out, &cell.content, ctx);
        ctx.in_spill_cell = enclosing_in_spill_cell;
        spill_content?;
        // Translate the placed line itself. Wrapping this whole `#context`
        // in `#move` changes the measurement region that `place(center)`
        // resolves against and can move a wide title off the page (#1493).
        if let Some((content_dx, content_dy)) = content_shift {
            let _ = write!(
                out,
                "]; place({anchor} + {vertical_anchor}, dx: {}pt, dy: {}pt, box(width: {}pt, height: {height}, clip: true)\
                 [#box(width: measure(o2p-spill).width)[#o2p-spill]])}}#box(width: 0pt, height: {height})",
                format_geometry(content_dx),
                format_geometry(content_dy),
                format_f64(clip_width_pt),
            );
        } else {
            let _ = write!(
                out,
                "]; place({anchor} + {vertical_anchor}, box(width: {}pt, height: {height}, clip: true)\
                 [#box(width: measure(o2p-spill).width)[#o2p-spill]])}}#box(width: 0pt, height: {height})",
                format_f64(clip_width_pt),
            );
        }
    } else {
        // A `w:trHeight` floor is `max(floor, content)`, which no Typst row
        // length expresses — a stated length pins the row, and `auto` drops
        // the floor. A one-row grid beside a strut of the floor's height is
        // exactly that maximum, and it costs nothing when the content already
        // wins (issue #965).
        let strut_height_pt: Option<f64> = row_minimum_height.map(|floor| {
            let inset: Insets = cell.padding.unwrap_or(default_cell_padding);
            (floor - inset.top - inset.bottom).max(0.0)
        });
        if let Some(height) = strut_height_pt {
            let _ = write!(
                out,
                "#grid(columns: (0pt, 1fr), rows: (auto,), box(width: 0pt, height: {}pt), [",
                format_f64(height)
            );
        }
        generate_sheet_cell_content(out, &cell.content, ctx)?;
        if strut_height_pt.is_some() {
            out.push_str("])");
        }
    }
    if wraps_content_shift {
        out.push(']');
    }
    ctx.cell_seats_text_on_descender = enclosing_cell_seats_on_descender;
    ctx.cell_vertical_align = enclosing_cell_vertical_align;
    ctx.cell_sheet_row_line = enclosing_cell_sheet_row_line;
    ctx.cell_sheet_seat = enclosing_cell_sheet_seat;
    out.push_str("],\n");
    Ok(())
}

/// Excel renders an x14 line sparkline to an opaque cell-sized bitmap before
/// fitting the worksheet to paper. The issue #1261 ground truth carries a
/// 42 x 18px image in a 42 x 19pt destination cell; its frame starts 1pt in
/// from the left and top boundaries and runs to the far edges.
/// Reconstructing the declared geometry before rasterising preserves that
/// image density on fitted sheets instead of generating a smaller bitmap from
/// the already-scaled tracks.
fn generate_sparkline_overlay(
    out: &mut String,
    sparkline: &SparklineInfo,
    cell: &TableCell,
    default_cell_padding: Insets,
    row_height: Option<f64>,
    cell_track_height: Option<f64>,
    ctx: &mut GenCtx,
) {
    let scale = ctx.sheet_print_scale().unwrap_or(1.0).max(0.01);
    let inset = cell_inset_with_border(cell, default_cell_padding);
    let Some(content_width_pt) = ctx.available_measure_pt else {
        return;
    };
    let track_width_pt = content_width_pt + inset.left + inset.right;
    let track_height_pt = cell_track_height.or(row_height).unwrap_or(20.0 * scale);
    let declared_width_pt = track_width_pt / scale;
    let declared_height_pt = track_height_pt / scale;
    let pixel_width = declared_width_pt.round().max(1.0) as u32;
    let pixel_height = (declared_height_pt - 1.0).round().max(1.0) as u32;
    let frame_width_pt = (declared_width_pt - 1.0).max(1.0) * scale;
    let frame_height_pt = (declared_height_pt - 1.0).max(1.0) * scale;
    let Some(data) = sparkline_png(sparkline, pixel_width, pixel_height) else {
        return;
    };
    let path = ctx.add_generated_png(data);
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt, image(\"{}\", width: {}pt, height: {}pt, fit: \"stretch\"))",
        format_geometry(scale - inset.left),
        format_geometry(2.0 * scale - inset.top),
        path,
        format_f64(frame_width_pt),
        format_f64(frame_height_pt),
    );
}

/// Rasterise a one-pixel butt-capped line with the quarter-coverage palette
/// Excel's embedded sparkline images use. Four-by-four subpixel sampling is
/// quantised to 25% coverage, yielding the exact white/25/50/75/100% colour
/// ladder visible in the native issue #1261 images.
fn sparkline_png(sparkline: &SparklineInfo, width: u32, height: u32) -> Option<Vec<u8>> {
    if width < 2 || height < 2 || sparkline.values.len() < 2 {
        return None;
    }
    let present: Vec<f64> = sparkline.values.iter().flatten().copied().collect();
    if present.len() < 2 {
        return None;
    }
    let minimum = present.iter().copied().fold(f64::INFINITY, f64::min);
    let maximum = present.iter().copied().fold(f64::NEG_INFINITY, f64::max);
    let left = 3.5_f64.min(f64::from(width) / 2.0);
    let right = (f64::from(width) - 4.5).max(left);
    let top = 4.5_f64.min(f64::from(height) / 2.0);
    let bottom = (f64::from(height) - 4.5).max(top);
    let last_index = (sparkline.values.len() - 1) as f64;
    let points: Vec<Option<(f64, f64)>> = sparkline
        .values
        .iter()
        .enumerate()
        .map(|(index, value)| {
            value.map(|value| {
                let x = left + (right - left) * index as f64 / last_index;
                let y = if (maximum - minimum).abs() < f64::EPSILON {
                    top
                } else {
                    top + (bottom - top) * (maximum - value) / (maximum - minimum)
                };
                (x, y)
            })
        })
        .collect();
    let segments: Vec<((f64, f64), (f64, f64))> = points
        .windows(2)
        .filter_map(|pair| Some((pair[0]?, pair[1]?)))
        .collect();
    if segments.is_empty() {
        return None;
    }
    let interior_vertices: Vec<(f64, f64)> = points
        .windows(3)
        .filter_map(|triple| match triple {
            [Some(_), Some(middle), Some(_)] => Some(*middle),
            _ => None,
        })
        .collect();
    let mut image = image::RgbImage::from_pixel(width, height, image::Rgb([255, 255, 255]));
    for y in 0..height {
        for x in 0..width {
            let mut covered = 0_u32;
            for sample_y in 0..4 {
                for sample_x in 0..4 {
                    let px = f64::from(x) + (f64::from(sample_x) + 0.5) / 4.0;
                    let py = f64::from(y) + (f64::from(sample_y) + 0.5) / 4.0;
                    let on_segment = segments
                        .iter()
                        .any(|&(start, end)| sample_hits_butt_segment(px, py, start, end));
                    let on_join = interior_vertices
                        .iter()
                        .any(|&(vx, vy)| (px - vx).hypot(py - vy) <= 0.5);
                    covered += u32::from(on_segment || on_join);
                }
            }
            let coverage_quarters = ((covered + 2) / 4).min(4);
            if coverage_quarters == 0 {
                continue;
            }
            let blend = |ink: u8| -> u8 {
                ((u32::from(ink) * coverage_quarters + 255 * (4 - coverage_quarters) + 2) / 4) as u8
            };
            image.put_pixel(
                x,
                y,
                image::Rgb([
                    blend(sparkline.color.r),
                    blend(sparkline.color.g),
                    blend(sparkline.color.b),
                ]),
            );
        }
    }
    let mut bytes = Cursor::new(Vec::new());
    image::DynamicImage::ImageRgb8(image)
        .write_to(&mut bytes, RasterImageFormat::Png)
        .ok()?;
    Some(bytes.into_inner())
}

fn sample_hits_butt_segment(px: f64, py: f64, start: (f64, f64), end: (f64, f64)) -> bool {
    let (dx, dy) = (end.0 - start.0, end.1 - start.1);
    let length_squared = dx * dx + dy * dy;
    if length_squared <= f64::EPSILON {
        return false;
    }
    let projection = ((px - start.0) * dx + (py - start.1) * dy) / length_squared;
    if !(0.0..=1.0).contains(&projection) {
        return false;
    }
    let nearest_x = start.0 + projection * dx;
    let nearest_y = start.1 + projection * dy;
    (px - nearest_x).hypot(py - nearest_y) <= 0.5
}

#[cfg(test)]
mod sparkline_tests {
    use super::*;

    #[test]
    fn flat_series_matches_excels_42_by_18_quarter_coverage_raster() {
        let sparkline = SparklineInfo {
            values: vec![Some(565.0); 12],
            color: Color::new(0x29, 0x74, 0x4F),
        };
        let png = sparkline_png(&sparkline, 42, 18).expect("a line must rasterise");
        let image = image::load_from_memory(&png).unwrap().to_rgb8();
        assert_eq!(image.dimensions(), (42, 18));
        assert_eq!(image.get_pixel(2, 4).0, [255, 255, 255]);
        assert_eq!(image.get_pixel(3, 4).0, [148, 186, 167]);
        assert_eq!(image.get_pixel(4, 4).0, [41, 116, 79]);
        assert_eq!(image.get_pixel(36, 4).0, [41, 116, 79]);
        assert_eq!(image.get_pixel(37, 4).0, [148, 186, 167]);
        assert_eq!(image.get_pixel(38, 4).0, [255, 255, 255]);
        assert_eq!(image.get_pixel(4, 5).0, [255, 255, 255]);
    }
}

/// Writer's borderless DOCX table-cell text origin sits this far into the
/// positive x side of the cell track. The #1219 native PDF trace measures the
/// same 0.10pt delta for left, centre, and right paragraph alignment, so this
/// is a content-box seat rather than a margin or text-width correction.
const WRITER_TABLE_CELL_X_ORIGIN_SEAT_PT: f64 = 0.1;

/// The translation from Typst's cell track to Writer's positive-axis content
/// seat. A resolved boundary owner adds its painted band only to the following
/// cell, exactly like the band itself; the borderless x seat applies to every
/// DOCX cell without changing its layout measure.
fn word_cell_content_shift(boundary_band: &Option<BoundaryBandCell<'_>>) -> Option<(f64, f64)> {
    let band = boundary_band.as_ref()?;
    if band.paint_model != TableBorderPaintModel::WordPositiveAxisBands {
        return None;
    }
    let border_dx = band
        .painted_border
        .as_ref()
        .and_then(|border| border.left.as_ref())
        .map_or(0.0, |side| word_pdf_border_side(side).width / 2.0);
    let dy = band
        .painted_border
        .as_ref()
        .and_then(|border| border.top.as_ref())
        .map_or(0.0, |side| word_pdf_border_side(side).width);
    Some((WRITER_TABLE_CELL_X_ORIGIN_SEAT_PT + border_dx, dy))
}

/// Excel centres a filled horizontal merge on the visible positive-axis
/// region, which includes its background bleed. The native #1493 probes keep
/// left alignment and an unfilled merge on the nominal track, so this is a
/// content-only seat rather than a change to the cell's width.
fn excel_merged_cell_content_shift(
    boundary_band: &Option<BoundaryBandCell<'_>>,
    cell: &TableCell,
) -> Option<(f64, f64)> {
    let band = boundary_band.as_ref()?;
    (band.paint_model == TableBorderPaintModel::ExcelBoundaryBands
        && cell.col_span > 1
        && cell.background.is_some()
        && cell_horizontal_alignment(cell) == Some(Alignment::Center))
    .then_some((BAND_RUN_END_EXTENSION_PT, 0.0))
}

/// Height, in points, of the single line box a spill cell's paragraph emits —
/// the same metric edges the block carries, times the run's own font size.
/// `None` when the font's metrics are unknown.
fn spill_line_box_height_pt(cell: &TableCell, ctx: &GenCtx) -> Option<f64> {
    let paragraph: &Paragraph = cell.content.iter().find_map(|block| match block {
        Block::Paragraph(paragraph) => Some(paragraph),
        _ => None,
    })?;
    let line_box: CellLineBox = word_cell_line_box(
        &paragraph.runs,
        &paragraph.style,
        ctx.line_grid_pitch,
        ctx.row_east_asian,
        ctx.cell_vertical_align,
        ctx.cell_seats_text_on_descender,
        ctx.cell_sheet_row_line.as_ref(),
        ctx.cell_sheet_seat,
        ctx.sheet_print_scale(),
    )?;
    Some((line_box.top_em + line_box.bottom_em) * line_box.font_size_pt)
}

/// The horizontal alignment a cell's own paragraph declares, if any.
fn cell_horizontal_alignment(cell: &TableCell) -> Option<Alignment> {
    cell.content.iter().find_map(|block| match block {
        Block::Paragraph(paragraph) => paragraph.style.alignment,
        _ => None,
    })
}

fn write_double_border_overlays(out: &mut String, border: &CellBorder, padding: Insets) {
    if let Some(side) = border
        .top
        .as_ref()
        .filter(|side| side.style == BorderLineStyle::Double)
    {
        write_horizontal_double_border(out, side, padding, true);
    }
    if let Some(side) = border
        .bottom
        .as_ref()
        .filter(|side| side.style == BorderLineStyle::Double)
    {
        write_horizontal_double_border(out, side, padding, false);
    }
    if let Some(side) = border
        .left
        .as_ref()
        .filter(|side| side.style == BorderLineStyle::Double)
    {
        write_vertical_double_border(out, side, padding, true);
    }
    if let Some(side) = border
        .right
        .as_ref()
        .filter(|side| side.style == BorderLineStyle::Double)
    {
        write_vertical_double_border(out, side, padding, false);
    }
}

fn write_horizontal_double_border(
    out: &mut String,
    side: &BorderSide,
    padding: Insets,
    is_top: bool,
) {
    let align = if is_top {
        "top + left"
    } else {
        "bottom + left"
    };
    let first_dy = if is_top {
        -padding.top - side.width
    } else {
        padding.bottom - side.width
    };
    let second_dy = if is_top {
        -padding.top + side.width
    } else {
        padding.bottom + side.width
    };
    let dx = -padding.left;
    let length_extra = padding.left + padding.right;
    write_double_border_line(out, align, dx, first_dy, "0deg", length_extra, side);
    write_double_border_line(out, align, dx, second_dy, "0deg", length_extra, side);
}

fn write_vertical_double_border(
    out: &mut String,
    side: &BorderSide,
    padding: Insets,
    is_left: bool,
) {
    let align = if is_left { "top + left" } else { "top + right" };
    let first_dx = if is_left {
        -padding.left - side.width
    } else {
        padding.right - side.width
    };
    let second_dx = if is_left {
        -padding.left + side.width
    } else {
        padding.right + side.width
    };
    let dy = -padding.top;
    let length_extra = padding.top + padding.bottom;
    write_double_border_line(out, align, first_dx, dy, "90deg", length_extra, side);
    write_double_border_line(out, align, second_dx, dy, "90deg", length_extra, side);
}

fn write_double_border_line(
    out: &mut String,
    align: &str,
    dx: f64,
    dy: f64,
    angle: &str,
    length_extra: f64,
    side: &BorderSide,
) {
    let _ = write!(
        out,
        "#place({align}, dx: {}pt, dy: {}pt, line(length: 100% + {}pt, angle: {angle}, stroke: {}pt + {}))",
        format_geometry(dx),
        format_geometry(dy),
        format_geometry(length_extra),
        format_geometry(side.width),
        rgb(&side.color),
    );
}

pub(super) fn format_geometry(value: f64) -> String {
    let rounded = (value * 1_000.0).round() / 1_000.0;
    format_f64(if rounded == -0.0 { 0.0 } else { rounded })
}

/// The cell's inset, with the layout space its horizontal borders occupy.
///
/// Typst draws our per-cell strokes without reserving room for them, but Word
/// counts a border's width in the row height. Each horizontal border is shared
/// between the rows above and below it, so each cell takes half (issues #500,
/// #503).
fn cell_inset_with_border(cell: &TableCell, default_cell_padding: Insets) -> Insets {
    let padding: Insets = cell.padding.unwrap_or(default_cell_padding);
    let Some(border) = &cell.border else {
        return padding;
    };
    let half = |side: &Option<BorderSide>| side.as_ref().map_or(0.0, |s| s.width / 2.0);
    Insets {
        top: padding.top + half(&border.top),
        bottom: padding.bottom + half(&border.bottom),
        ..padding
    }
}

/// Excel extends every border band 1pt past its end boundary — the
/// `[A_start, A_end + 1]` run rule, measured independent of weight (issue
/// #619). It is what lets horizontal bands own the corner blocks.
const BAND_RUN_END_EXTENSION_PT: f64 = 1.0;

/// Inward overlap that keeps three same-colour fill rectangles from meeting
/// at only one antialiased corner. The measured outer bleed remains 1pt; this
/// quarter point stays inside the cell and closes the raster seam from #1397.
const BACKGROUND_BLEED_SEAM_OVERLAP_PT: f64 = 0.25;

/// Width of Excel's printed gridline band, in points.
///
/// Measured on native Excel exports of NumberFormatTests (issue #622,
/// /Volumes/T7/scratch/issue-622/nft2-p1.rects.txt and nft2-p2.trace): every
/// gridline is an axis-aligned fill rect exactly 1.0pt thick filling the
/// boundary band [B, B+1] — no stroke ops and no fractional hairlines exist
/// anywhere in the traces.
const PRINTED_GRIDLINE_WIDTH_PT: f64 = 1.0;

/// The side a printed gridline paints on an unowned boundary.
///
/// Pure black, not gray and not a theme colour: the GT traces fill every
/// gridline with "0 0 0" in ICCBased sRGB (issue #622 measurement — the
/// common assumption of gray printed gridlines is wrong for Excel GT).
fn printed_gridline_side() -> BorderSide {
    BorderSide {
        width: PRINTED_GRIDLINE_WIDTH_PT,
        color: Color::black(),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
    }
}

/// One edge of the black 1pt print frame that `<printOptions headings="1"/>`
/// draws around the heading bands and the data grid (issue #623).
///
/// GT (nft-sheet-0002 trace): the frame is four 1pt pure-black fill bands on
/// the table's exterior boundaries — [54,538]x[72,73] top, [54,55]x[72,710]
/// left, [537,538]x[72,710] right, [54,538]x[709,710] bottom — each on the
/// same [B, B+1] band convention as the #619/#622 rules, and everything else
/// on the page is clipped to the frame's interior.
fn print_heading_frame_side() -> BorderSide {
    BorderSide {
        width: 1.0,
        color: Color::black(),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
    }
}

/// Total order for Excel's shared-boundary conflict rule (issue #619 review,
/// remediation 1). Derived `PartialOrd` compares the fields lexicographically
/// in declaration order.
///
/// Heaviness is a style precedence, not the stored stroke width: Excel paints
/// a double rule on top of every single band — including thick — even though
/// each of a double's two strokes is stored at the thin 1pt weight (Excel's
/// double-on-top conflict behaviour). Below double, the total painted band
/// width decides (thick 3 > medium 2 > thin/hair 1), and at equal width a
/// solid rule beats a patterned one (hair/dotted/dashed). Exact rank ties
/// fall back to the caller's positional rule — the lower/right cell's
/// top/left slot keeps the boundary — which never consults colour, so
/// ownership is colour-stable: declarations differing only in colour resolve
/// the same way regardless of which side declares which colour.
#[derive(Clone, Copy, PartialEq, PartialOrd)]
struct BoundaryConflictRank {
    /// Excel's double-on-top rule: any double outranks any single band.
    is_double: bool,
    /// Total painted band width in points (thick 3 > medium 2 > thin 1).
    band_width: f64,
    /// Solid beats patterned at equal band width.
    is_solid: bool,
}

/// Rank one declared side for shared-boundary conflict resolution.
fn boundary_conflict_rank(side: &BorderSide) -> BoundaryConflictRank {
    BoundaryConflictRank {
        is_double: side.style == BorderLineStyle::Double,
        band_width: side.width,
        is_solid: side.style == BorderLineStyle::Solid,
    }
}

/// Which border sides each cell paints in the boundary-band regime, parallel
/// to `table.rows[r].cells[c]`.
///
/// A printed border belongs to the grid *boundary*, not to the declaring
/// cell: the band is anchored to the boundary whichever neighbour declares
/// it, so a boundary declared by both neighbours must paint exactly once.
/// Conflicting declarations resolve to the heavier style per
/// [`BoundaryConflictRank`]. Each internal boundary is therefore left on the
/// highest-ranked declaration; equal declarations use the lower/right cell's
/// top/left slot. That slot paints toward positive x/y for both Excel's
/// measured bands and Word's filled rectangles.
///
/// Suppression is whole-side: when merged cells overlap a boundary only
/// partially, both declarations are kept and the equal bands overlap
/// invisibly. Partial overlaps of *differing* weight would need per-track
/// resolution (known limitation, deliberately skipped).
///
/// Each of the `repeating_header_boundaries` is a grid boundary whose upper
/// side lives in a repeating header block (the #623 letter strip's bottom,
/// the last print-title row's bottom) while its lower side prints once. The
/// lower row renders once but the header repeats on every page, so a band
/// left on the lower side would vanish under the repeated header on pages
/// 2+. At those boundaries ties therefore go to the *header's* declaration,
/// and a strictly heavier lower declaration is additionally adopted into the
/// header cell's bottom slot: both sides then paint the same
/// boundary-anchored band — coincident and invisible where they overlap on
/// page 1, while the header's copy repeats with it.
///
/// A sheet that prints headings is the second exception, and a broader one:
/// there "exactly one declaration" does not hold at any ordinary horizontal
/// boundary, because a band left on one side alone closes only one side of a
/// page break codegen cannot see. Both are kept. See
/// `keeps_coincident_horizontal_bands` in the body (issue #722).
pub(super) fn resolve_boundary_painted_borders(
    table: &Table,
    num_cols: usize,
    repeating_header_boundaries: &[usize],
) -> Vec<Vec<Option<CellBorder>>> {
    use std::collections::{HashMap, HashSet};

    /// Grid footprint of one emitted cell.
    struct CellPlacement {
        row_index: usize,
        cell_index: usize,
        first_col: usize,
        row_span: usize,
        col_span: usize,
    }

    // Mirror the emission walk in `generate_table_rows` exactly, so the
    // painted set stays parallel to what actually renders.
    let mut placements: Vec<CellPlacement> = Vec::new();
    let mut rowspan_remaining: Vec<usize> = vec![0usize; num_cols];
    for (row_index, row) in table.rows.iter().enumerate() {
        for rs in rowspan_remaining.iter_mut() {
            if *rs > 0 {
                *rs -= 1;
            }
        }
        let mut col_pos: usize = 0;
        for (cell_index, cell) in row.cells.iter().enumerate() {
            if cell.col_span == 0 || cell.row_span == 0 {
                continue;
            }
            while col_pos < num_cols && rowspan_remaining[col_pos] > 0 {
                col_pos += 1;
            }
            if col_pos >= num_cols {
                break;
            }
            let remaining: usize = num_cols - col_pos;
            let col_span: usize = (cell.col_span as usize).min(remaining).max(1);
            placements.push(CellPlacement {
                row_index,
                cell_index,
                first_col: col_pos,
                row_span: (cell.row_span as usize).max(1),
                col_span,
            });
            if cell.row_span > 1 {
                for rs in rowspan_remaining.iter_mut().skip(col_pos).take(col_span) {
                    *rs = cell.row_span as usize;
                }
            }
            col_pos += col_span;
        }
    }

    let cell_of = |placement: &CellPlacement| -> &TableCell {
        &table.rows[placement.row_index].cells[placement.cell_index]
    };

    // Declared sides per (grid boundary index, crossing track). A horizontal
    // boundary `b` separates grid rows `b-1` and `b`; its declarations are
    // the bottoms of cells ending at `b` and the tops of cells starting
    // there. Vertical boundaries likewise with columns. Whole sides are kept
    // (not just widths) because ranking needs the style and the header/body
    // boundary below adopts the winning side wholesale.
    let mut bottom_sides: HashMap<(usize, usize), BorderSide> = HashMap::new();
    let mut top_sides: HashMap<(usize, usize), BorderSide> = HashMap::new();
    let mut right_sides: HashMap<(usize, usize), BorderSide> = HashMap::new();
    let mut left_sides: HashMap<(usize, usize), BorderSide> = HashMap::new();
    for placement in &placements {
        let Some(border) = &cell_of(placement).border else {
            continue;
        };
        let column_tracks = placement.first_col..placement.first_col + placement.col_span;
        let row_tracks = placement.row_index..placement.row_index + placement.row_span;
        if let Some(side) = &border.bottom {
            for col in column_tracks.clone() {
                bottom_sides.insert(
                    (placement.row_index + placement.row_span, col),
                    side.clone(),
                );
            }
        }
        if let Some(side) = &border.top {
            for col in column_tracks {
                top_sides.insert((placement.row_index, col), side.clone());
            }
        }
        if let Some(side) = &border.right {
            for row in row_tracks.clone() {
                right_sides.insert(
                    (placement.first_col + placement.col_span, row),
                    side.clone(),
                );
            }
        }
        if let Some(side) = &border.left {
            for row in row_tracks {
                left_sides.insert((placement.first_col, row), side.clone());
            }
        }
    }

    // Whether a horizontal boundary keeps *both* its declarations instead of
    // resolving them down to one.
    //
    // Codegen cannot see the page breaks Typst chooses, so a boundary painted
    // by only one of its two owners is closed on only one side of a break. At
    // a tie the rule below hands the boundary to the top owner — the row
    // *below* — which on an intermediate page is the first row of the *next*
    // page, leaving the previous page's bottom edge open. Excel frames every
    // page across the full block width, and the row-number gutter's span was
    // the part left hanging: the data span survived only because #622's
    // gridline seeds happen to put a band there independently, and with
    // gridlines off the whole bottom edge would go (issue #722).
    //
    // Inverting the tie would only move the hole to the top of the next page,
    // so both bands are kept. They are one rule drawn twice at the same
    // coordinate, which is what `augment_page_with_print_headings` already
    // assumes when it declares a bottom on every gutter cell: "adjacent
    // cells' coincident bands overlap invisibly, as in #622". The cost is a
    // doubled draw op per interior boundary, visible in a rect census but not
    // in the ink.
    //
    // Scoped to sheets that print headings, which is where the frame is a
    // stated Excel behaviour. Word tables keep one owner per boundary.
    let keeps_coincident_horizontal_bands: bool = table.prints_headings;
    // The repeating-header ownership exception was measured for Excel's
    // split header/body emission. Word paints an equal shared rule inside the
    // following body row even when the row above repeats (#724).
    let applies_excel_repeating_header_ownership: bool =
        table.border_paint_model == TableBorderPaintModel::ExcelBoundaryBands;

    let mut painted: Vec<Vec<Option<CellBorder>>> = table
        .rows
        .iter()
        .map(|row| vec![None; row.cells.len()])
        .collect();
    for placement in &placements {
        let Some(border) = &cell_of(placement).border else {
            continue;
        };
        let column_tracks = placement.first_col..placement.first_col + placement.col_span;
        let row_tracks = placement.row_index..placement.row_index + placement.row_span;
        let mut resolved: CellBorder = border.clone();
        // A bottom/right yields to a neighbour's declaration ranked at least
        // as heavy (ties keep the top/left owner); a top/left yields only to
        // a strictly heavier one. Exactly one side survives per fully shared
        // boundary in every rank combination. The repeating-header boundary
        // inverts the tie direction (see the function docs).
        let bottom_boundary: usize = placement.row_index + placement.row_span;
        let bottom_is_repeating_header_boundary: bool = applies_excel_repeating_header_ownership
            && repeating_header_boundaries.contains(&bottom_boundary);
        if let Some(side) = &resolved.bottom
            && (bottom_is_repeating_header_boundary || !keeps_coincident_horizontal_bands)
            && column_tracks.clone().all(|col| {
                top_sides
                    .get(&(bottom_boundary, col))
                    .is_some_and(|neighbour| {
                        let neighbour_rank = boundary_conflict_rank(neighbour);
                        let own_rank = boundary_conflict_rank(side);
                        if bottom_is_repeating_header_boundary {
                            neighbour_rank > own_rank
                        } else {
                            neighbour_rank >= own_rank
                        }
                    })
            })
        {
            resolved.bottom = if bottom_is_repeating_header_boundary {
                // The body's strictly heavier band must also repeat with the
                // header: adopt the highest-ranked body declaration into this
                // header cell's bottom slot instead of dropping it.
                column_tracks
                    .clone()
                    .filter_map(|col| top_sides.get(&(bottom_boundary, col)))
                    .max_by(|a, b| {
                        boundary_conflict_rank(a)
                            .partial_cmp(&boundary_conflict_rank(b))
                            .unwrap_or(std::cmp::Ordering::Equal)
                    })
                    .cloned()
            } else {
                None
            };
        }
        if let Some(side) = &resolved.top
            && column_tracks.clone().all(|col| {
                bottom_sides
                    .get(&(placement.row_index, col))
                    .is_some_and(|neighbour| {
                        let neighbour_rank = boundary_conflict_rank(neighbour);
                        let own_rank = boundary_conflict_rank(side);
                        if applies_excel_repeating_header_ownership
                            && repeating_header_boundaries.contains(&placement.row_index)
                        {
                            // Ties at the repeating-header boundary stay with
                            // the header's bottom declaration.
                            neighbour_rank >= own_rank
                        } else {
                            neighbour_rank > own_rank
                        }
                    })
            })
        {
            resolved.top = None;
        }
        if let Some(side) = &resolved.right
            && row_tracks.clone().all(|row| {
                left_sides
                    .get(&(placement.first_col + placement.col_span, row))
                    .is_some_and(|neighbour| {
                        boundary_conflict_rank(neighbour) >= boundary_conflict_rank(side)
                    })
            })
        {
            resolved.right = None;
        }
        if let Some(side) = &resolved.left
            && row_tracks.clone().all(|row| {
                right_sides
                    .get(&(placement.first_col, row))
                    .is_some_and(|neighbour| {
                        boundary_conflict_rank(neighbour) > boundary_conflict_rank(side)
                    })
            })
        {
            resolved.left = None;
        }
        if resolved.top.is_some()
            || resolved.bottom.is_some()
            || resolved.left.is_some()
            || resolved.right.is_some()
        {
            painted[placement.row_index][placement.cell_index] = Some(resolved);
        }
    }

    // Printed gridlines (issue #622): `<printOptions gridLines="1"/>` rules
    // every cell boundary of the printed range with Excel's gridline band,
    // strictly below any explicit declaration — a boundary owned by any
    // declared side (either neighbour's) keeps that side alone, hair borders
    // included, which the #619 rank would otherwise wrongly outrank. Every
    // placement seeds all four of its unowned sides: the two seeds of an
    // interior boundary are boundary-anchored to the same [B, B+1] strip and
    // coincide invisibly, and the redundant bottom seed is what closes the
    // grid at a page break, where GT draws the bottom rule (which row ends a
    // page only the renderer knows).
    if table.prints_gridlines {
        // A cell fill suppresses all four adjacent gridline segments: GT
        // truncates the interior verticals at a filled row and omits the
        // horizontal at the fill's bottom boundary (Tests p1 vs the
        // fill-free p2 control), because fills paint after gridlines.
        //
        // TODO(#622 follow-up: a background-filled row that lands as the
        // first row of a page under natural pagination leaves the previous
        // page's grid open at that boundary — GT closes it; suppression is
        // kept because an unsuppressed band would paint over the fill's top
        // edge on every within-page filled row, the far more common case).
        let mut fill_suppressed_horizontal: HashSet<(usize, usize)> = HashSet::new();
        let mut fill_suppressed_vertical: HashSet<(usize, usize)> = HashSet::new();
        for placement in &placements {
            if cell_of(placement).background.is_none() {
                continue;
            }
            for col in placement.first_col..placement.first_col + placement.col_span {
                fill_suppressed_horizontal.insert((placement.row_index, col));
                fill_suppressed_horizontal.insert((placement.row_index + placement.row_span, col));
            }
            for row in placement.row_index..placement.row_index + placement.row_span {
                fill_suppressed_vertical.insert((placement.first_col, row));
                fill_suppressed_vertical.insert((placement.first_col + placement.col_span, row));
            }
        }
        // Printed headings (issue #623): boundary 0 in each direction is the
        // heading exterior — the strip row's top and the gutter column's
        // left. GT rules those edges as the black print FRAME, which the
        // forcing pass below paints; excluding GRIDLINE-styled seeds here
        // keeps the frame band the boundary's only owner (replace, not
        // stack). The data area starts at row/column 1, so its seeding is
        // untouched.
        let heading_exterior_is_excluded: bool = table.prints_headings;
        // Stated as "no disqualifier applies" rather than as a chain of ANDed
        // negations: the two are equivalent by De Morgan, but only this form
        // satisfies clippy::nonminimal_bool.
        let horizontal_boundary_is_free = |boundary: usize, col: usize| -> bool {
            !((heading_exterior_is_excluded && boundary == 0)
                || top_sides.contains_key(&(boundary, col))
                || bottom_sides.contains_key(&(boundary, col))
                || fill_suppressed_horizontal.contains(&(boundary, col)))
        };
        let vertical_boundary_is_free = |boundary: usize, row: usize| -> bool {
            !((heading_exterior_is_excluded && boundary == 0)
                || left_sides.contains_key(&(boundary, row))
                || right_sides.contains_key(&(boundary, row))
                || fill_suppressed_vertical.contains(&(boundary, row)))
        };
        for placement in &placements {
            let column_tracks = placement.first_col..placement.first_col + placement.col_span;
            let row_tracks = placement.row_index..placement.row_index + placement.row_span;
            let mut seeded: CellBorder = painted[placement.row_index][placement.cell_index]
                .take()
                .unwrap_or_default();
            // Whole-side seeding: a side whose boundary is even partially
            // declared or fill-suppressed stays unseeded — the merged-cell
            // partial-overlap simplification of #619, erring toward fewer
            // rules, which is also GT's direction for fills.
            if seeded.top.is_none()
                && column_tracks
                    .clone()
                    .all(|col| horizontal_boundary_is_free(placement.row_index, col))
            {
                seeded.top = Some(printed_gridline_side());
            }
            if seeded.bottom.is_none()
                && column_tracks.clone().all(|col| {
                    horizontal_boundary_is_free(placement.row_index + placement.row_span, col)
                })
            {
                seeded.bottom = Some(printed_gridline_side());
            }
            if seeded.left.is_none()
                && row_tracks
                    .clone()
                    .all(|row| vertical_boundary_is_free(placement.first_col, row))
            {
                seeded.left = Some(printed_gridline_side());
            }
            if seeded.right.is_none()
                && row_tracks.clone().all(|row| {
                    vertical_boundary_is_free(placement.first_col + placement.col_span, row)
                })
            {
                seeded.right = Some(printed_gridline_side());
            }
            if seeded.top.is_some()
                || seeded.bottom.is_some()
                || seeded.left.is_some()
                || seeded.right.is_some()
            {
                painted[placement.row_index][placement.cell_index] = Some(seeded);
            }
        }
    }

    // Printed headings (issue #623): GT draws a 1pt black frame enclosing
    // the heading bands and the data grid, on the table's exterior
    // boundaries — the corner box's top and left edges ARE the frame. Forced
    // here, after declaration resolution and gridline seeding, so the frame
    // band REPLACES whatever landed on a frame boundary (a heading gray
    // rule, a #622 closure seed) instead of stacking on it; painting is
    // band-only, so no layout inset moves. The strip-top edge rides the
    // repeating header block (pages 2+ carry it) and the left/right edges
    // ride each row, so they close on every rendered page; the bottom edge
    // exists only on the LAST table row — a Typst page break inside the
    // table leaves that page's bottom open unless printed gridlines already
    // close it with their own black band (tracked in #722).
    // TODO(#623 follow-up: whether a gridLines-only sheet — headings off —
    // prints this frame is unmeasured; the frame is gated on
    // prints_headings alone until a GT probe answers it).
    if table.prints_headings {
        let frame_rank: BoundaryConflictRank = boundary_conflict_rank(&print_heading_frame_side());
        // A strictly heavier declared band keeps its boundary (GT for a
        // heavy cell border meeting the frame is unmeasured; erring toward
        // the author's declaration); equal-rank sides — the heading gray
        // rules — yield to the frame, as GT clips them to its interior.
        let force_frame = |slot: &mut Option<BorderSide>| {
            if !slot
                .as_ref()
                .is_some_and(|side| boundary_conflict_rank(side) > frame_rank)
            {
                *slot = Some(print_heading_frame_side());
            }
        };
        for placement in &placements {
            let on_top_exterior: bool = placement.row_index == 0;
            let on_left_exterior: bool = placement.first_col == 0;
            let on_right_exterior: bool = placement.first_col + placement.col_span == num_cols;
            let on_bottom_exterior: bool =
                placement.row_index + placement.row_span == table.rows.len();
            if !(on_top_exterior || on_left_exterior || on_right_exterior || on_bottom_exterior) {
                continue;
            }
            let mut framed: CellBorder = painted[placement.row_index][placement.cell_index]
                .take()
                .unwrap_or_default();
            if on_top_exterior {
                force_frame(&mut framed.top);
            }
            if on_left_exterior {
                force_frame(&mut framed.left);
            }
            if on_right_exterior {
                force_frame(&mut framed.right);
            }
            if on_bottom_exterior {
                force_frame(&mut framed.bottom);
            }
            painted[placement.row_index][placement.cell_index] = Some(framed);
        }
    }

    painted
}

struct ExcelBackgroundBleedTrims {
    top: Vec<Vec<f64>>,
    bottom_left: Vec<Vec<f64>>,
}

#[derive(Clone, Copy, PartialEq)]
struct ExcelBackgroundPaint {
    color: Color,
    alpha: Option<f64>,
}

type ExcelBackgroundPaintGrid = Vec<Vec<Option<ExcelBackgroundPaint>>>;

/// How much ink each Excel background must preserve at its top-right and
/// bottom-left junctions after shared-boundary conflict resolution.
///
/// Cell children paint in table order. A right-edge background bleed in row
/// `r` therefore paints after a bottom band owned by row `r - 1`, even though
/// the band wins the Excel boundary conflict. Trimming the later vertical
/// bleed by the winning band's positive extent preserves the final colour at
/// that crossing (#1475). Reading the resolved paint plan is important: a
/// losing declaration can have a different extent and must not create a gap.
/// At the opposite corner, adjacent differently painted cells reserve the
/// left cell's positive-axis block in the later cell's bottom strip (#1495).
fn excel_background_bleed_trims(
    table: &Table,
    num_cols: usize,
    painted_borders: &[Vec<Option<CellBorder>>],
) -> ExcelBackgroundBleedTrims {
    struct CellPlacement {
        row_index: usize,
        cell_index: usize,
        first_col: usize,
        row_span: usize,
        col_span: usize,
    }

    let mut placements: Vec<CellPlacement> = Vec::new();
    let mut rowspan_remaining: Vec<usize> = vec![0; num_cols];
    for (row_index, row) in table.rows.iter().enumerate() {
        for remaining in &mut rowspan_remaining {
            *remaining = remaining.saturating_sub(1);
        }
        let mut col_pos: usize = 0;
        for (cell_index, cell) in row.cells.iter().enumerate() {
            if cell.col_span == 0 || cell.row_span == 0 {
                continue;
            }
            while col_pos < num_cols && rowspan_remaining[col_pos] > 0 {
                col_pos += 1;
            }
            if col_pos >= num_cols {
                break;
            }
            let col_span: usize = (cell.col_span as usize).min(num_cols - col_pos).max(1);
            let row_span: usize = (cell.row_span as usize).max(1);
            placements.push(CellPlacement {
                row_index,
                cell_index,
                first_col: col_pos,
                row_span,
                col_span,
            });
            if row_span > 1 {
                for remaining in rowspan_remaining.iter_mut().skip(col_pos).take(col_span) {
                    *remaining = row_span;
                }
            }
            col_pos += col_span;
        }
    }

    // Track each occupied grid slot's effective paint. When adjacent cells
    // use different paints, the later cell's bottom strip must begin after
    // the preceding cell's +x corner block instead of overpainting it
    // (#1495). Row spans are expanded here so a cell beside a merge receives
    // the same answer as one beside an ordinary cell.
    let mut background_paints: ExcelBackgroundPaintGrid =
        vec![vec![None; num_cols]; table.rows.len()];
    for placement in &placements {
        let cell: &TableCell = &table.rows[placement.row_index].cells[placement.cell_index];
        let paint: Option<ExcelBackgroundPaint> =
            cell.background.map(|color| ExcelBackgroundPaint {
                color,
                alpha: cell.background_alpha,
            });
        if paint.is_none() {
            continue;
        }
        for row_paints in background_paints
            .iter_mut()
            .skip(placement.row_index)
            .take(placement.row_span)
        {
            for slot in row_paints
                .iter_mut()
                .skip(placement.first_col)
                .take(placement.col_span)
            {
                *slot = paint;
            }
        }
    }

    let mut horizontal_border_positive_extents: Vec<Vec<f64>> =
        vec![vec![0.0; num_cols]; table.rows.len() + 1];
    let mut horizontal_background_positive_extents: Vec<Vec<f64>> =
        vec![vec![0.0; num_cols]; table.rows.len() + 1];
    for placement in &placements {
        let column_tracks = placement.first_col..placement.first_col + placement.col_span;
        let bottom_boundary: usize = placement.row_index + placement.row_span;
        let cell: &TableCell = &table.rows[placement.row_index].cells[placement.cell_index];
        // An upper cell's own bottom background bleed is visible as the title
        // rule in #1475 even without a declared border. It owns the same
        // positive 1pt boundary band. It also owns the corner where an
        // adjacent upper-right fill begins; otherwise the lower-left cell's
        // later vertical bleed cuts a differently coloured notch (#1495).
        if bottom_boundary < horizontal_background_positive_extents.len()
            && cell.background.is_some()
        {
            for col in column_tracks.clone() {
                horizontal_background_positive_extents[bottom_boundary][col] =
                    BAND_RUN_END_EXTENSION_PT;
            }
        }

        if let Some(border) = painted_borders[placement.row_index][placement.cell_index].as_ref() {
            if let Some(side) = &border.top {
                let extent: f64 = positive_axis_band_extent(side);
                for col in column_tracks.clone() {
                    horizontal_border_positive_extents[placement.row_index][col] =
                        horizontal_border_positive_extents[placement.row_index][col].max(extent);
                }
            }
            if bottom_boundary < horizontal_border_positive_extents.len()
                && let Some(side) = &border.bottom
            {
                let extent: f64 = positive_axis_band_extent(side);
                for col in column_tracks {
                    horizontal_border_positive_extents[bottom_boundary][col] =
                        horizontal_border_positive_extents[bottom_boundary][col].max(extent);
                }
            }
        }
    }

    let mut top_trims: Vec<Vec<f64>> = table
        .rows
        .iter()
        .map(|row| vec![0.0; row.cells.len()])
        .collect();
    let mut bottom_left_trims: Vec<Vec<f64>> = table
        .rows
        .iter()
        .map(|row| vec![0.0; row.cells.len()])
        .collect();
    for placement in placements {
        let rightmost_col: usize = placement.first_col + placement.col_span - 1;
        top_trims[placement.row_index][placement.cell_index] = horizontal_border_positive_extents
            [placement.row_index][rightmost_col]
            .max(horizontal_background_positive_extents[placement.row_index][rightmost_col]);
        if placement.first_col > 0 {
            let current_paint = background_paints[placement.row_index][placement.first_col];
            let left_paint = background_paints[placement.row_index][placement.first_col - 1];
            if current_paint.is_some() && left_paint.is_some() && current_paint != left_paint {
                bottom_left_trims[placement.row_index][placement.cell_index] =
                    BAND_RUN_END_EXTENSION_PT;
            }
        }
    }
    ExcelBackgroundBleedTrims {
        top: top_trims,
        bottom_left: bottom_left_trims,
    }
}

/// Furthest point a centred Excel band reaches on the positive side of its
/// nominal grid boundary. Double borders use the second band's far edge.
fn positive_axis_band_extent(side: &BorderSide) -> f64 {
    band_centre_offsets(side)
        .into_iter()
        .map(|centre| centre + side.width / 2.0)
        .fold(0.0, f64::max)
}

/// One cell's share of the boundary-band regime, threaded from the row walk
/// into the cell writer.
struct BoundaryBandCell<'a> {
    /// The sides this cell paints after shared-boundary resolution. `None`
    /// paints nothing but still selects the band regime (no cell stroke).
    painted_border: &'a Option<CellBorder>,
    /// Positive-axis horizontal ink already occupying this cell's top-right
    /// junction. A later Excel fill starts below it instead of overpainting
    /// the rule emitted by an earlier row (#1475).
    background_bleed_top_trim_pt: f64,
    /// Space reserved at the left end of this cell's bottom strip for a
    /// differently coloured left neighbour's +x corner block (#1495).
    background_bleed_bottom_left_trim_pt: f64,
    /// The source application's placement convention for those bands.
    paint_model: TableBorderPaintModel,
    /// How far this cell's vertical bands may extend.
    vertical_extent: VerticalBandExtent,
}

/// How a cell's vertical border bands obtain their length (issue #619).
///
/// A Typst-relative length (`100%`) inside a `#place` resolves against the
/// measurement region — the remaining page, not the cell — whenever any row
/// the cell spans is auto-sized (measured on typst 0.14/0.15), painting
/// page-long spears. Vertical bands therefore always use concrete lengths.
enum VerticalBandExtent {
    /// Every spanned row's height is fixed: one top-anchored band of the
    /// summed frame height covers boundary to boundary exactly.
    FrameHeight(f64),
    /// The span includes auto-sized rows, whose final height only the
    /// renderer knows. Two twin bands anchored at the cell's top and bottom
    /// edges are painted instead, each sized from the row's tallest
    /// single-line frame: the twins coincide exactly on single-line rows
    /// (the row is sized by that same line box) and cover a wrapped row from
    /// both ends without overshooting, because a row is at least as tall as
    /// its tallest cell's first line. Rows wrapping past roughly twice the
    /// estimate keep a mid-row gap (known limitation).
    TwinBands(f64),
    /// No cell in the row has usable line metrics: twin bands sized by the
    /// ambient text size, following the data-bar `1.2em` precedent.
    TwinBandsEmFallback,
}

/// Decide [`VerticalBandExtent`] for one cell of a boundary-band table.
///
/// `row_frame_estimate_cache` is the calling row-walk's per-row memo for
/// [`auto_row_frame_height_estimate_pt`]: the estimate is row-wide and
/// costly (it reads every cell's font metrics), so each row computes it at
/// most once however many cells land here.
#[allow(clippy::too_many_arguments)]
fn vertical_band_extent(
    rows: &[TableRow],
    row_index: usize,
    cell: &TableCell,
    fixed_row_heights: bool,
    default_cell_padding: Insets,
    ctx: &GenCtx,
    row_frame_estimate_cache: &mut Option<Option<f64>>,
) -> VerticalBandExtent {
    if let Some(frame_height_pt) =
        fixed_spanned_row_height(rows, row_index, cell.row_span, fixed_row_heights)
    {
        return VerticalBandExtent::FrameHeight(frame_height_pt);
    }
    // Estimate from the anchor row only; a multi-row span over auto rows
    // keeps whatever the twins cover (known limitation).
    let frame_estimate: Option<f64> = *row_frame_estimate_cache.get_or_insert_with(|| {
        auto_row_frame_height_estimate_pt(&rows[row_index], default_cell_padding, ctx)
    });
    match frame_estimate {
        Some(frame_estimate_pt) => VerticalBandExtent::TwinBands(frame_estimate_pt),
        None => VerticalBandExtent::TwinBandsEmFallback,
    }
}

/// The total fixed height of every row a cell joins. A span reaching outside
/// this row group (for example across a Typst header/body split), an auto-row
/// table, or any row without an exact height states no complete track.
fn fixed_spanned_row_height(
    rows: &[TableRow],
    row_index: usize,
    row_span: u32,
    fixed_row_heights: bool,
) -> Option<f64> {
    if !fixed_row_heights {
        return None;
    }
    let row_span: usize = (row_span as usize).max(1);
    let spanned_rows: &[TableRow] = &rows[row_index..(row_index + row_span).min(rows.len())];
    if spanned_rows.len() != row_span {
        return None;
    }
    spanned_rows
        .iter()
        .try_fold(0.0_f64, |sum, row| row.height.map(|height| sum + height))
}

// Test-only probe counting `auto_row_frame_height_estimate_pt` calls, so
// the once-per-row caching contract (issue #619 review, remediation 3) is
// assertable: the estimate walks every cell's font metrics, and calling it
// per cell made vertical-band preparation O(cells²) per row. (A regular
// comment: rustc discards doc comments attached to macro invocations.)
#[cfg(test)]
thread_local! {
    pub(super) static AUTO_ROW_FRAME_ESTIMATE_CALLS: std::cell::Cell<usize> =
        const { std::cell::Cell::new(0) };
}

/// The frame height the renderer will give an auto-sized row, estimated as
/// the tallest cell's single-line box plus that cell's insets — exact when
/// the sizing cell holds one line (the common spreadsheet case), an
/// underestimate when it wraps.
fn auto_row_frame_height_estimate_pt(
    row: &TableRow,
    default_cell_padding: Insets,
    ctx: &GenCtx,
) -> Option<f64> {
    #[cfg(test)]
    AUTO_ROW_FRAME_ESTIMATE_CALLS.with(|calls| calls.set(calls.get() + 1));
    row.cells
        .iter()
        .filter_map(|cell| {
            let paragraph: &Paragraph = cell.content.iter().find_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })?;
            // Auto rows never seat text on the descender — the seating gate
            // keys on a fixed row height — so the estimate must not either,
            // and no shared row line exists for it to resolve against.
            let line_box: CellLineBox = word_cell_line_box(
                &paragraph.runs,
                &paragraph.style,
                ctx.line_grid_pitch,
                ctx.row_east_asian,
                cell.vertical_align.or(ctx.table_default_vertical_align),
                false,
                None,
                None,
                ctx.sheet_print_scale(),
            )?;
            let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
            Some(
                (line_box.top_em + line_box.bottom_em) * line_box.font_size_pt
                    + inset.top
                    + inset.bottom,
            )
        })
        .fold(None, |tallest: Option<f64>, height| {
            Some(tallest.map_or(height, |t| t.max(height)))
        })
}

/// Word's PDF graphics grid is 1/300 inch (0.24 pt). Its native exports snap
/// a declared half-point table border to two grid units, or 0.48 pt (#724).
const WORD_PDF_GRAPHICS_GRID_PT: f64 = 72.0 / 300.0;

fn word_pdf_border_side(side: &BorderSide) -> BorderSide {
    let grid_units: f64 = (side.width / WORD_PDF_GRAPHICS_GRID_PT).round().max(1.0);
    BorderSide {
        width: grid_units * WORD_PDF_GRAPHICS_GRID_PT,
        ..side.clone()
    }
}

/// Paint Word table borders as rectangles whose leading edge is the nominal
/// grid boundary and whose ink extends along positive x/y (#724). Solid and
/// double rules use filled rectangles, matching Word's PDF primitives;
/// patterned rules keep a line solely to preserve their dash sequence.
fn write_word_positive_axis_border_overlays(
    out: &mut String,
    border: &CellBorder,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
) {
    if let Some(side) = &border.top {
        write_word_horizontal_border(out, side, inset, true);
    }
    if let Some(side) = &border.bottom {
        write_word_horizontal_border(out, side, inset, false);
    }
    if let Some(side) = &border.left {
        write_word_vertical_border(out, side, inset, vertical_extent, true);
    }
    if let Some(side) = &border.right {
        write_word_vertical_border(out, side, inset, vertical_extent, false);
    }
}

fn write_word_horizontal_border(out: &mut String, side: &BorderSide, inset: Insets, is_top: bool) {
    let painted_side: BorderSide = word_pdf_border_side(side);
    let horizontal_length: String =
        format!("100% + {}pt", format_geometry(inset.left + inset.right));
    let align: &str = if is_top {
        "top + left"
    } else {
        "bottom + left"
    };
    let offsets: &[f64] = if painted_side.style == BorderLineStyle::Double {
        &[0.0, painted_side.width * 2.0]
    } else {
        &[0.0]
    };
    for inward_offset in offsets {
        if matches!(
            painted_side.style,
            BorderLineStyle::Solid | BorderLineStyle::Double | BorderLineStyle::None
        ) {
            let dy: f64 = if is_top {
                -inset.top + inward_offset
            } else {
                inset.bottom + painted_side.width + inward_offset
            };
            write_band_rect(
                out,
                align,
                &format!("{}pt", format_geometry(-inset.left)),
                &format!("{}pt", format_geometry(dy)),
                &horizontal_length,
                &format!("{}pt", format_geometry(painted_side.width)),
                &painted_side.color,
            );
        } else {
            let centre: f64 = painted_side.width / 2.0;
            let dy: f64 = if is_top {
                -inset.top + centre
            } else {
                inset.bottom + centre
            };
            write_boundary_band_line(
                out,
                align,
                -inset.left,
                dy,
                "0deg",
                &horizontal_length,
                &painted_side,
            );
        }
    }
}

fn write_word_vertical_border(
    out: &mut String,
    side: &BorderSide,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
    is_left: bool,
) {
    let painted_side: BorderSide = word_pdf_border_side(side);
    let align: &str = if is_left { "top + left" } else { "top + right" };
    let offsets: &[f64] = if painted_side.style == BorderLineStyle::Double {
        &[0.0, painted_side.width * 2.0]
    } else {
        &[0.0]
    };
    for inward_offset in offsets {
        let band_anchor_x: f64 = if is_left {
            -inset.left + inward_offset
        } else {
            inset.right + painted_side.width + inward_offset
        };
        if matches!(
            painted_side.style,
            BorderLineStyle::Solid | BorderLineStyle::Double | BorderLineStyle::None
        ) {
            write_word_vertical_band_rect(
                out,
                align,
                band_anchor_x,
                painted_side.width,
                inset,
                vertical_extent,
                &painted_side.color,
            );
        } else {
            let centre_x: f64 = if is_left {
                band_anchor_x + painted_side.width / 2.0
            } else {
                band_anchor_x - painted_side.width / 2.0
            };
            write_word_patterned_vertical_band(
                out,
                align,
                centre_x,
                inset,
                vertical_extent,
                &painted_side,
            );
        }
    }
}

fn write_word_vertical_band_rect(
    out: &mut String,
    align: &str,
    dx: f64,
    width: f64,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
    color: &Color,
) {
    let dx: String = format!("{}pt", format_geometry(dx));
    let width: String = format!("{}pt", format_geometry(width));
    match *vertical_extent {
        VerticalBandExtent::FrameHeight(frame_height_pt) => {
            write_band_rect(
                out,
                align,
                &dx,
                &format!("{}pt", format_geometry(-inset.top)),
                &width,
                &format!("{}pt", format_geometry(frame_height_pt)),
                color,
            );
        }
        VerticalBandExtent::TwinBands(frame_estimate_pt) => {
            let height: String = format!("{}pt", format_geometry(frame_estimate_pt));
            write_band_rect(
                out,
                align,
                &dx,
                &format!("{}pt", format_geometry(-inset.top)),
                &width,
                &height,
                color,
            );
            write_band_rect(
                out,
                &align.replacen("top", "bottom", 1),
                &dx,
                &format!("{}pt", format_geometry(inset.bottom)),
                &width,
                &height,
                color,
            );
        }
        VerticalBandExtent::TwinBandsEmFallback => {
            let height: String = format!("1.2em + {}pt", format_geometry(inset.top + inset.bottom));
            write_band_rect(
                out,
                align,
                &dx,
                &format!("{}pt", format_geometry(-inset.top)),
                &width,
                &height,
                color,
            );
            write_band_rect(
                out,
                &align.replacen("top", "bottom", 1),
                &dx,
                &format!("{}pt", format_geometry(inset.bottom)),
                &width,
                &height,
                color,
            );
        }
    }
}

fn write_word_patterned_vertical_band(
    out: &mut String,
    align: &str,
    dx: f64,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
    side: &BorderSide,
) {
    match *vertical_extent {
        VerticalBandExtent::FrameHeight(frame_height_pt) => write_boundary_band_line(
            out,
            align,
            dx,
            -inset.top,
            "90deg",
            &format!("{}pt", format_geometry(frame_height_pt)),
            side,
        ),
        VerticalBandExtent::TwinBands(frame_estimate_pt) => {
            let length: String = format!("{}pt", format_geometry(frame_estimate_pt));
            write_boundary_band_line(out, align, dx, -inset.top, "90deg", &length, side);
            write_boundary_band_line(
                out,
                &align.replacen("top", "bottom", 1),
                dx,
                inset.bottom,
                "-90deg",
                &length,
                side,
            );
        }
        VerticalBandExtent::TwinBandsEmFallback => {
            let length: String = format!("1.2em + {}pt", format_geometry(inset.top + inset.bottom));
            write_boundary_band_line(out, align, dx, -inset.top, "90deg", &length, side);
            write_boundary_band_line(
                out,
                &align.replacen("top", "bottom", 1),
                dx,
                inset.bottom,
                "-90deg",
                &length,
                side,
            );
        }
    }
}

/// Place one axis-aligned band rectangle, out of layout, at `align` shifted by
/// `dx`/`dy`. Every boundary-anchored paint paints one of these: Word's border
/// rectangles and Excel's background bleed alike.
fn write_band_rect(
    out: &mut String,
    align: &str,
    dx: &str,
    dy: &str,
    width: &str,
    height: &str,
    color: &Color,
) {
    let _ = write!(
        out,
        "#place({align}, dx: {dx}, dy: {dy}, rect(width: {width}, height: {height}, fill: {}, stroke: none))",
        rgb(color),
    );
}

/// Paint a cell's borders as filled bands anchored to the nominal grid
/// boundaries, as Excel prints them (issue #619; native Excel 16.111
/// one-factor probe + golden-mock GT traces):
///
/// - 1pt styles (`thin`, `hair`, dashes): band `[B, B+1]`, on the +x/+y side
///   even at the table's outer right/bottom edge;
/// - `medium` (2pt): `[B-1, B+1]`; `thick` (3pt): `[B-1, B+2]` — an odd
///   leftover point always lands on the +x/+y side;
/// - `double`: two 1pt bands `[B-1, B]` and `[B+1, B+2]`, the boundary strip
///   `[B, B+1]` being the gap.
///
/// The `dx`/`dy` offsets back out the cell's effective inset so `B` is the
/// grid boundary the table already lays out. Horizontal bands run from the
/// cell's left boundary to 1pt past its right boundary, owning the corner
/// blocks; verticals span the same extended run instead of Excel's trim to
/// strictly between the horizontals — the overlap is same-colour in the GT
/// regimes and therefore invisible. Corners whose crossing rules differ in
/// colour would need that trim (known limitation, deliberately skipped).
fn write_boundary_anchored_border_overlays(
    out: &mut String,
    border: &CellBorder,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
) {
    // Horizontal bands can stay relative: a line's `100%` length resolves
    // against the cell's width, which the spreadsheet's fixed column tracks
    // always determine (and colspans span correctly through it).
    let horizontal_length: String = format!(
        "100% + {}pt",
        format_geometry(inset.left + inset.right + BAND_RUN_END_EXTENSION_PT)
    );
    if let Some(side) = &border.top {
        for centre in band_centre_offsets(side) {
            write_boundary_band_line(
                out,
                "top + left",
                -inset.left,
                -inset.top + centre,
                "0deg",
                &horizontal_length,
                side,
            );
        }
    }
    if let Some(side) = &border.bottom {
        for centre in band_centre_offsets(side) {
            write_boundary_band_line(
                out,
                "bottom + left",
                -inset.left,
                inset.bottom + centre,
                "0deg",
                &horizontal_length,
                side,
            );
        }
    }
    if let Some(side) = &border.left {
        for centre in band_centre_offsets(side) {
            write_vertical_boundary_band(
                out,
                side,
                "left",
                -inset.left + centre,
                inset,
                vertical_extent,
            );
        }
    }
    if let Some(side) = &border.right {
        for centre in band_centre_offsets(side) {
            write_vertical_boundary_band(
                out,
                side,
                "right",
                inset.right + centre,
                inset,
                vertical_extent,
            );
        }
    }
}

/// Paint one vertical band rule at `dx` from the cell's `horizontal_anchor`
/// edge, spanning from the row's top boundary to 1pt past its bottom boundary
/// per [`VerticalBandExtent`]'s answer for this cell.
fn write_vertical_boundary_band(
    out: &mut String,
    side: &BorderSide,
    horizontal_anchor: &str,
    dx: f64,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
) {
    let top_anchor: String = format!("top + {horizontal_anchor}");
    let (length, twins): (String, bool) = vertical_band_run(vertical_extent, inset);
    if twins {
        write_vertical_twin_bands(out, side, dx, inset, &top_anchor, &length);
    } else {
        write_boundary_band_line(out, &top_anchor, dx, -inset.top, "90deg", &length, side);
    }
}

/// The concrete length one boundary-anchored vertical run takes, and whether
/// it must be painted as twins because that length only estimates the row's
/// frame.
fn vertical_band_run(vertical_extent: &VerticalBandExtent, inset: Insets) -> (String, bool) {
    match *vertical_extent {
        VerticalBandExtent::FrameHeight(frame_height_pt) => (
            format!(
                "{}pt",
                format_geometry(frame_height_pt + BAND_RUN_END_EXTENSION_PT)
            ),
            false,
        ),
        VerticalBandExtent::TwinBands(frame_estimate_pt) => (
            format!(
                "{}pt",
                format_geometry(frame_estimate_pt + BAND_RUN_END_EXTENSION_PT)
            ),
            true,
        ),
        VerticalBandExtent::TwinBandsEmFallback => (
            format!(
                "1.2em + {}pt",
                format_geometry(inset.top + inset.bottom + BAND_RUN_END_EXTENSION_PT)
            ),
            true,
        ),
    }
}

/// Paint the two strips by which Excel's cell background overruns its own grid
/// rect: the shading covers its box **plus** the 1pt boundary band on the
/// bottom and right edges, so neighbouring shadings overlap by exactly the
/// strip a border then paints over (issue #1190, the fill half of the #619
/// probe). Typst's cell `fill:` covers the track exactly, so the overrun is
/// painted here. With Typst's bottom/right placement alignment, growing each
/// strip inward while leaving its placement offset unchanged keeps the
/// measured outer edge fixed; without the overlap the three paths meet at one
/// antialiased corner and can leave a one-pixel pinhole (issue #1397).
///
/// Only the +y/+x edges bleed, which is what keeps the strips harmless: every
/// band sharing one of them — this cell's own, and both neighbours' — is
/// emitted after these rects and stays on top of them.
fn write_excel_background_bleed(
    out: &mut String,
    background: &Color,
    background_alpha: Option<f64>,
    inset: Insets,
    vertical_extent: &VerticalBandExtent,
    trims_pt: (f64, f64),
    sheet_print_scale: Option<f64>,
) {
    let (top_trim_pt, bottom_left_trim_pt) = trims_pt;
    // A fitted worksheet paints these lengths in declared sheet space and
    // scales the result. Word/PowerPoint and an unfitted sheet keep the
    // measured physical lengths unchanged (#1538).
    let scale: f64 = sheet_print_scale
        .filter(|scale| *scale > 0.0 && *scale < 1.0)
        .unwrap_or(1.0);
    let end_extension_pt: f64 = BAND_RUN_END_EXTENSION_PT * scale;
    let seam_overlap_pt: f64 = BACKGROUND_BLEED_SEAM_OVERLAP_PT * scale;
    let bleed_with_overlap: String =
        format!("{}pt", format_geometry(end_extension_pt + seam_overlap_pt));
    // The bottom strip runs the cell's full width plus the corner block it
    // shares with the right one, exactly as a horizontal border band does.
    // A differently painted left neighbour reserves its own corner block, so
    // shift this strip's left edge inward while keeping its right edge fixed.
    write_background_rect(
        out,
        "bottom + left",
        &format!("{}pt", format_geometry(-inset.left + bottom_left_trim_pt)),
        &format!("{}pt", format_geometry(inset.bottom + end_extension_pt)),
        &format!(
            "100% + {}pt",
            format_geometry(inset.left + inset.right + end_extension_pt - bottom_left_trim_pt)
        ),
        &bleed_with_overlap,
        background,
        background_alpha,
    );
    // The right strip spans the row frame, and takes a concrete length for the
    // same reason [`VerticalBandExtent`] gives.
    let dx: String = format!("{}pt", format_geometry(inset.right + end_extension_pt));
    // Keep the lower edge fixed while moving the upper edge below a
    // horizontal band already painted by the preceding row. This applies to
    // both auto-row twins: reducing the bottom-anchored twin's height moves
    // only its top, so it cannot reintroduce the same overpaint (#1475).
    let top_dy: String = format!("{}pt", format_geometry(-inset.top + top_trim_pt));
    let (height, twins): (String, bool) =
        background_bleed_vertical_run(vertical_extent, inset, top_trim_pt, end_extension_pt);
    write_background_rect(
        out,
        "top + right",
        &dx,
        &top_dy,
        &bleed_with_overlap,
        &height,
        background,
        background_alpha,
    );
    if twins {
        write_background_rect(
            out,
            "bottom + right",
            &dx,
            &format!("{}pt", format_geometry(inset.bottom + end_extension_pt)),
            &bleed_with_overlap,
            &height,
            background,
            background_alpha,
        );
    }
}

/// Vertical background-bleed length after preserving `top_trim_pt` of ink at
/// the row's top boundary. The untrimmed lower edge remains unchanged.
fn background_bleed_vertical_run(
    vertical_extent: &VerticalBandExtent,
    inset: Insets,
    top_trim_pt: f64,
    end_extension_pt: f64,
) -> (String, bool) {
    match *vertical_extent {
        VerticalBandExtent::FrameHeight(frame_height_pt) => (
            format!(
                "{}pt",
                format_geometry((frame_height_pt + end_extension_pt - top_trim_pt).max(0.0))
            ),
            false,
        ),
        VerticalBandExtent::TwinBands(frame_estimate_pt) => (
            format!(
                "{}pt",
                format_geometry((frame_estimate_pt + end_extension_pt - top_trim_pt).max(0.0))
            ),
            true,
        ),
        VerticalBandExtent::TwinBandsEmFallback => (
            format!(
                "1.2em + {}pt",
                format_geometry(inset.top + inset.bottom + end_extension_pt - top_trim_pt)
            ),
            true,
        ),
    }
}

#[allow(clippy::too_many_arguments)]
fn write_background_rect(
    out: &mut String,
    align: &str,
    dx: &str,
    dy: &str,
    width: &str,
    height: &str,
    color: &Color,
    alpha: Option<f64>,
) {
    let paint: String = alpha.map_or_else(
        || rgb(color),
        |alpha| rgb_with_alpha(color, (alpha.clamp(0.0, 1.0) * 255.0).round() as u8),
    );
    let _ = write!(
        out,
        "#place({align}, dx: {dx}, dy: {dy}, rect(width: {width}, height: {height}, fill: {paint}, stroke: none))",
    );
}

/// Two same-length rules: one hanging from the row's top boundary, one rising
/// from 1pt past its bottom boundary. On a single-line auto row they coincide
/// exactly; on a wrapped row they cover it from both ends.
fn write_vertical_twin_bands(
    out: &mut String,
    side: &BorderSide,
    dx: f64,
    inset: Insets,
    top_anchor: &str,
    length: &str,
) {
    let bottom_anchor: String = top_anchor.replacen("top", "bottom", 1);
    write_boundary_band_line(out, top_anchor, dx, -inset.top, "90deg", length, side);
    write_boundary_band_line(
        out,
        &bottom_anchor,
        dx,
        inset.bottom + BAND_RUN_END_EXTENSION_PT,
        "-90deg",
        length,
        side,
    );
}

/// Offsets of each painted rule's centre line from the boundary `B`, for a
/// band of the side's width `w`: a single band `[B - floor(w/2), ...]` puts
/// the centre at `w/2 - floor(w/2)` (0.5 for thin/thick, 0 for medium); a
/// double paints one rule per band.
fn band_centre_offsets(side: &BorderSide) -> Vec<f64> {
    if side.style == BorderLineStyle::Double {
        vec![-side.width / 2.0, side.width * 1.5]
    } else {
        vec![side.width / 2.0 - (side.width / 2.0).floor()]
    }
}

#[allow(clippy::too_many_arguments)]
fn write_boundary_band_line(
    out: &mut String,
    align: &str,
    dx: f64,
    dy: f64,
    angle: &str,
    length: &str,
    side: &BorderSide,
) {
    // `stroke_value` keeps the dash dict of patterned styles (dashed, dotted,
    // hair) on the overlay line; a double side's two rules are each plain.
    let _ = write!(
        out,
        "#place({align}, dx: {}pt, dy: {}pt, line(length: {length}, angle: {angle}, stroke: {}))",
        format_geometry(dx),
        format_geometry(dy),
        stroke_value(side, true),
    );
}

fn write_cell_params(
    out: &mut String,
    cell: &TableCell,
    clamped_colspan: u32,
    default_cell_padding: Insets,
    paints_boundary_bands: bool,
    // `Some` replaces whatever vertical alignment the cell declares or
    // inherits: a tight spreadsheet row anchors every cell on its one centred
    // line (issue #839). Emitted even for a cell declaring nothing, because
    // the sheet table's default it would inherit is bottom.
    forced_vertical_align: Option<CellVerticalAlign>,
) {
    let mut first = true;

    if clamped_colspan > 1 {
        write_param(out, &mut first, &format!("colspan: {clamped_colspan}"));
    }
    if cell.row_span > 1 {
        write_param(out, &mut first, &format!("rowspan: {}", cell.row_span));
    }
    if let Some(ref bg) = cell.background {
        let fill: String = cell.background_alpha.map_or_else(
            || format_color(bg),
            |alpha| {
                format!(
                    "fill: {}",
                    rgb_with_alpha(bg, (alpha.clamp(0.0, 1.0) * 255.0).round() as u8)
                )
            },
        );
        write_param(out, &mut first, &fill);
    }
    let inset: Insets = cell_inset_with_border(cell, default_cell_padding);
    if cell.padding.is_some() || cell.border.is_some() {
        write_param(
            out,
            &mut first,
            &format!("inset: {}", format_insets(&inset)),
        );
    }
    // A boundary-band cell paints its borders as overlays instead: a Typst
    // stroke is centred on the track boundary, which cannot reproduce either
    // Excel's measured bands (#619) or Word's positive-axis rectangles
    // (#724). The
    // `inset` above still reserves the border's layout space either way.
    if !paints_boundary_bands && let Some(ref border) = cell.border {
        let stroke = format_cell_stroke(border);
        if !stroke.is_empty() {
            write_param(out, &mut first, &stroke);
        }
    }
    let emitted_vertical_align: Option<CellVerticalAlign> =
        forced_vertical_align.or(cell.vertical_align);
    if let Some(va) = emitted_vertical_align {
        let align_str: &str = match va {
            CellVerticalAlign::Top => "top",
            CellVerticalAlign::Center => "horizon",
            CellVerticalAlign::Bottom => "bottom",
        };
        write_param(out, &mut first, &format!("align: {align_str}"));
    }
}

fn format_cell_stroke(border: &CellBorder) -> String {
    let mut parts = Vec::with_capacity(4);

    if let Some(ref side) = border.top
        && side.style != BorderLineStyle::Double
    {
        parts.push(format!("top: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.bottom
        && side.style != BorderLineStyle::Double
    {
        parts.push(format!("bottom: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.left
        && side.style != BorderLineStyle::Double
    {
        parts.push(format!("left: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.right
        && side.style != BorderLineStyle::Double
    {
        parts.push(format!("right: {}", format_border_side(side)));
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("stroke: ({})", parts.join(", "))
    }
}

fn format_border_side(side: &BorderSide) -> String {
    stroke_value(side, true)
}

/// Generate a table cell's content at its established seat when the fitted
/// worksheet's paint layer has been translated underneath it.
///
/// `#move` leaves layout dimensions unchanged, so the inverse translation
/// cancels only the visual table-paint offset from issue #1538. Off a fitted
/// sheet the helper is byte-for-byte the old direct generation path.
fn generate_sheet_cell_content(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let paint_offset_pt: Option<(f64, f64)> = ctx.sheet_paint_offset_pt.take();
    if let Some((paint_dx_pt, paint_dy_pt)) = paint_offset_pt {
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[",
            format_geometry(-paint_dx_pt),
            format_geometry(-paint_dy_pt),
        );
    }
    let result = generate_cell_content(out, blocks, ctx);
    ctx.sheet_paint_offset_pt = paint_offset_pt;
    if paint_offset_pt.is_some() {
        out.push(']');
    }
    result
}

fn generate_cell_content(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    // Word separates stacked cell paragraphs only by the resolved
    // `w:spacing w:after`/`w:before` — the explicit `#v` emissions — but
    // sibling `#block` wrappers otherwise pick up Typst's ambient default
    // block spacing (1.2em at the document size), adding ~13pt Word never
    // shows (issue #625). This counts the stacked blocks; whether a given
    // paragraph may actually drop that ambient spacing is decided in
    // `generate_cell_paragraph`, which zeroes it only for paragraphs that
    // emit a fixed line box of their own. A lone block keeps today's exact
    // emission, since its boundary spacing vanishes at the cell edge anyway.
    let rendered_block_count: usize = blocks
        .iter()
        .filter(|block| {
            !matches!(
                block,
                Block::TableOfContents(_) | Block::PageBreak | Block::ColumnBreak
            )
        })
        .count();
    let stacks_multiple_blocks: bool = rendered_block_count > 1;
    for (i, block) in blocks.iter().enumerate() {
        if i > 0 {
            out.push('\n');
        }
        let paragraph_ctx = |para: &Paragraph| CellParagraphCtx {
            default_tab_width_pt: ctx.default_tab_width_pt,
            line_grid_pitch: ctx.line_grid_pitch,
            row_east_asian: ctx.row_east_asian,
            vertical_align: ctx.cell_vertical_align,
            seats_text_on_descender: ctx.cell_seats_text_on_descender,
            sheet_row_line: ctx.cell_sheet_row_line.clone(),
            sheet_seat: ctx.cell_sheet_seat,
            sheet_print_scale: ctx.sheet_print_scale(),
            in_spill_cell: ctx.in_spill_cell,
            uses_powerpoint_line_box: ctx.table_uses_powerpoint_line_box,
            stacks_multiple_blocks,
            paragraph_mark_metric_runs: para
                .runs
                .is_empty()
                .then(|| empty_cell_paragraph_metric_runs(blocks, i))
                .flatten(),
            breaks_hangul_at_eojeol: ctx.breaks_hangul_at_eojeol,
            available_measure_pt: ctx.available_measure_pt,
        };
        match block {
            // A `TOC` field inside a table cell is not a shape Word produces.
            Block::TableOfContents(_) => {}
            Block::Caption(caption) => {
                generate_cell_paragraph(out, &caption.paragraph, &paragraph_ctx(&caption.paragraph))
            }
            Block::Paragraph(para) => generate_cell_paragraph(out, para, &paragraph_ctx(para)),
            Block::Table(table) => {
                if ctx.table_depth < MAX_TABLE_DEPTH {
                    generate_table(out, table, ctx)?;
                }
            }
            Block::Image(img) => generate_image(out, img, ctx),
            Block::InlineImages(images) => {
                for image in images {
                    generate_image(out, image, ctx);
                }
            }
            Block::FloatingImage(fi) => generate_floating_image(out, fi, ctx),
            Block::FloatingTextBox(ftb) => generate_floating_text_box(out, ftb, ctx)?,
            Block::FloatingShape(fs) => generate_floating_shape(out, fs, ctx),
            Block::List(list) => {
                if can_render_fixed_text_list_inline(list) {
                    generate_fixed_text_list(out, list, true, None, false, false)?;
                } else {
                    // No wrapper settings reach a cell list, so it has no
                    // fixed text edges of its own to restore (issue #626).
                    generate_list(
                        out,
                        list,
                        None,
                        ListEojeolWrap {
                            breaks_hangul_at_eojeol: ctx.breaks_hangul_at_eojeol,
                            line_box_em: None,
                            available_measure_pt: ctx.available_measure_pt,
                            baseline_snap: None,
                        },
                    )?;
                }
            }
            Block::MathEquation(math) => generate_math_equation(out, math),
            Block::Chart(chart) => generate_chart(out, chart),
            Block::PageBreak | Block::ColumnBreak => {}
        }
    }
    Ok(())
}

/// The cell-level facts a paragraph's emission needs beyond its own IR.
struct CellParagraphCtx<'a> {
    default_tab_width_pt: f64,
    line_grid_pitch: Option<f64>,
    /// Decided once per row so every cell in it shares a baseline (issue #498).
    row_east_asian: RowEastAsianMetrics,
    /// The cell's effective Word vertical anchor, including the table default.
    vertical_align: Option<CellVerticalAlign>,
    seats_text_on_descender: bool,
    /// The one line the cell's tight spreadsheet row seats every cell on, so
    /// this paragraph's box resolves at the row's family and size rather than
    /// its own (issue #839). `None` outside that regime.
    sheet_row_line: Option<SheetRowLine>,
    /// The fixed sheet track the cell sits in, so its line seats where Excel
    /// prints it (issue #1063). `None` outside that regime.
    sheet_seat: Option<SheetCellSeat>,
    /// `Some(fit-to-page scale)` when this cell belongs to a spreadsheet.
    /// Excel reads its measured line pitch (#1163) and whole-point line seats
    /// (#1238) at the cell's declared size before scaling them. `None` off a
    /// sheet; the marker is the same one the descender seat keys on.
    sheet_print_scale: Option<f64>,
    /// Whether this paragraph is inside a spill cell's clipped wrapper, where
    /// the `#place` anchor already carries the cell's horizontal alignment.
    /// A `width: 100%` block inside that wrapper is not just redundant: the
    /// wrapper sizes itself from `measure()`, which lays out in an unbounded
    /// region where a percentage width has nothing to resolve against, and the
    /// paragraph came back so narrow that every word took a line of its own
    /// (issue #811).
    in_spill_cell: bool,
    /// Whether the cell stacks more than one rendered block, so this
    /// paragraph has a sibling to leak Typst's default block spacing against.
    stacks_multiple_blocks: bool,
    /// Runs standing in for the paragraph mark's own font when the paragraph
    /// has none of its own — see [`empty_cell_paragraph_metric_runs`].
    paragraph_mark_metric_runs: Option<&'a [Run]>,
    /// Whether the enclosing page is a Word flow page, whose Hangul lines
    /// break only at eojeol boundaries (issue #626). False for a slide or a
    /// sheet, which keep the engine's syllable breaking.
    breaks_hangul_at_eojeol: bool,
    /// Whether this cell paces its lines on PowerPoint's flat 1.2em line
    /// instead of Word's hhea one — true inside a slide's `<a:tbl>`
    /// (issue #663).
    uses_powerpoint_line_box: bool,
    /// The width one line of this cell has, in points: the column width less
    /// the cell's own inset. Bounds how wide a framed eojeol may be.
    available_measure_pt: Option<f64>,
}

/// The runs an empty `<w:p>` in a cell borrows its line box from.
///
/// Word lays a blank cell paragraph out on a full line, sized from the
/// paragraph mark's own `w:rPr`. The IR carries no runs — and so no font or
/// size — for such a paragraph, so the nearest sibling paragraph in the same
/// cell stands in: the one above by preference, since a spacer line follows
/// the text it separates (issue #625).
///
/// `None` when the cell holds no other text at all — a wholly blank cell,
/// whose height Word takes from the row and the cell insets rather than from
/// any run this codegen could measure.
/// TODO(#625 follow-up: a wholly blank cell keeps today's zero-height
/// emission, so a blank auto-height row is still one line short of Word;
/// sizing it needs the table/style default font, which the IR does not carry
/// to codegen — measure against a Word GT before inventing one).
fn empty_cell_paragraph_metric_runs(blocks: &[Block], index: usize) -> Option<&[Run]> {
    fn paragraph_runs(block: &Block) -> Option<&[Run]> {
        match block {
            Block::Paragraph(paragraph) => Some(paragraph.runs.as_slice()),
            Block::Caption(caption) => Some(caption.paragraph.runs.as_slice()),
            _ => None,
        }
    }
    let preceding = blocks[..index]
        .iter()
        .rev()
        .filter_map(paragraph_runs)
        .find(|runs| !runs.is_empty());
    preceding.or_else(|| {
        blocks[index + 1..]
            .iter()
            .filter_map(paragraph_runs)
            .find(|runs| !runs.is_empty())
    })
}

fn generate_cell_paragraph(out: &mut String, para: &Paragraph, cell: &CellParagraphCtx) {
    let style: &ParagraphStyle = &para.style;
    let alignment = style.alignment;
    let align_str: Option<&str> = match alignment {
        Some(Alignment::Left) => Some("left"),
        Some(Alignment::Center) => Some("center"),
        Some(Alignment::Right) => Some("right"),
        _ => None,
    };
    let line_height_settings: Option<String> = if cell.uses_powerpoint_line_box {
        // A slide's table cell paces on PowerPoint's flat 1.2em line, the same
        // model its own text boxes use, not on Word's hhea line (issue #663).
        powerpoint_line_height_settings(&para.runs, style)
    } else {
        // Off-slide, table-cell text occupies the font's full single-spacing
        // (hhea) line as a fixed box: a single-line cell must fill the whole
        // line height Word gives it rather than only the tighter metric box,
        // or auto-height rows come out short (issue #396). A cell whose *row*
        // holds East Asian text takes 1.3 times that line, like body text, and
        // a snapping grid's pitch above it — decided once per row so every
        // cell in it shares a baseline, the numeric ones included
        // (issues #498, #518).
        word_cell_line_box_settings(
            &para.runs,
            style,
            cell.line_grid_pitch,
            cell.row_east_asian,
            cell.vertical_align,
            cell.seats_text_on_descender,
            cell.sheet_row_line.as_ref(),
            cell.sheet_seat,
            cell.sheet_print_scale,
        )
    };
    // Whichever fixed edges the block wrapper below puts in force — the
    // computed cell line box, or the paragraph's own `LineBox` — is what a
    // framed eojeol has to restore inside itself (issue #626). The two are
    // mutually exclusive: `word_cell_line_box` bails on a paragraph that
    // declares a `LineBox`.
    let cell_line_box_em: Option<(f64, f64)> = word_cell_line_box(
        &para.runs,
        style,
        cell.line_grid_pitch,
        cell.row_east_asian,
        cell.vertical_align,
        cell.seats_text_on_descender,
        cell.sheet_row_line.as_ref(),
        cell.sheet_seat,
        cell.sheet_print_scale,
    )
    .map(|line_box| (line_box.top_em, line_box.bottom_em))
    .or_else(|| {
        style
            .line_box
            .map(|line_box| (line_box.ascent_em, line_box.descent_em))
    });
    // An empty `<w:p>` has no runs, so it resolves no line box above and would
    // otherwise emit nothing at all — zero height, where Word gives the
    // paragraph mark a full blank line (issue #625). Size that line from the
    // neighbours' metrics and hold it with a zero-width strut, the same shape
    // the spill wrapper uses. This mirrors the body path's `#v` branch for an
    // empty paragraph, at the cell's fixed line box instead of a flat 12pt.
    //
    // The blank line has to come from the same model as its neighbours, or a
    // slide's empty cell keeps Word's hhea height while the cell beside it
    // takes PowerPoint's 1.2em one (issue #663).
    let paragraph_mark_line_pt: Option<f64> = cell.paragraph_mark_metric_runs.and_then(|runs| {
        if cell.uses_powerpoint_line_box {
            powerpoint_line_box_pt(runs)
        } else {
            word_cell_line_box(
                runs,
                style,
                cell.line_grid_pitch,
                cell.row_east_asian,
                cell.vertical_align,
                cell.seats_text_on_descender,
                cell.sheet_row_line.as_ref(),
                cell.sheet_seat,
                cell.sheet_print_scale,
            )
            .map(|line_box| (line_box.top_em + line_box.bottom_em) * line_box.font_size_pt)
        }
    });
    // Typst's default block spacing may only be dropped where this paragraph
    // supplies a fixed line box of its own, which carries the whole advance;
    // adding Typst's gap on top would count the line twice. A paragraph that
    // resolves no box — an unknown face, say — keeps the default, or zeroing
    // its wrapper would leave it no vertical separation at all and collapse the
    // stack onto itself. A `w:spacing w:line` used to land in that second case
    // and advance short of Word for want of a box; it now scales one
    // (issue #727).
    let emits_fixed_line_box: bool =
        line_height_settings.is_some() || paragraph_mark_line_pt.is_some();
    let suppress_default_block_spacing: bool = cell.stacks_multiple_blocks && emits_fixed_line_box;
    // Word's `w:ind` offsets a paragraph's whole column wherever the paragraph
    // sits, cells included, so it rides the cell wrapper as an inset the same
    // way the body path puts it on its own outer block (issues #464, #938).
    // Unlike the body path this needs no separate outer block: a cell paragraph
    // paints no `w:shd` or `w:pBdr` of its own, so nothing has to span the
    // un-inset width.
    let indent: Option<(f64, f64)> = paragraph_indent_pt(style);
    let has_block_wrapper = cell_paragraph_needs_block_wrapper(style)
        || align_str.is_some()
        || line_height_settings.is_some()
        || suppress_default_block_spacing
        || indent.is_some();

    // Excel assigns the one-point residual of an odd, wrapped line stack to
    // the space above the block, which seats the ink one point lower than a
    // geometric centre. Typst centres the measured block exactly. Measure the
    // laid-out paragraph so automatic wraps participate, then translate only
    // odd stacks longer than one line; `move` preserves the row, width, wrap,
    // and following content geometry (issue #1494).
    let centered_sheet_odd_line_seat: Option<(f64, f64, f64, f64)> =
        centered_sheet_odd_line_seat(para, cell);
    if centered_sheet_odd_line_seat.is_some() {
        out.push_str("#context {\n  let o2p-centered-sheet-body = [");
    }

    if has_block_wrapper {
        out.push_str("#block(");
        write_cell_paragraph_block_params(
            out,
            align_str.is_some() && !cell.in_spill_cell,
            suppress_default_block_spacing,
            indent,
        );
        out.push_str(")[\n");
        write_line_box_settings(out, style.line_box);
        write_par_settings(out, style, &para.runs);
        if let Some(align_str) = align_str {
            let _ = writeln!(out, "  #set align({align_str})");
        }
        if let Some(ref settings) = line_height_settings {
            out.push_str(settings);
        }
    }

    if let Some(space_before) = style.space_before {
        let _ = writeln!(out, "#v({}pt)", format_f64(space_before));
    }

    match paragraph_mark_line_pt {
        Some(height_pt) => {
            let _ = write!(out, "#box(width: 0pt, height: {}pt)", format_f64(height_pt));
        }
        None => {
            // A spill wrapper deliberately lays its content out at natural
            // width and clips one physical line at the sheet boundary
            // (#811). It must never inherit Word's measured token breaker:
            // that breaker inserts real line boundaries for an overlong Latin
            // token (#1454), defeating the wrapper's no-wrap contract before
            // the outer box gets a chance to clip it.
            let eojeol_wrap = if cell.in_spill_cell {
                EojeolWrap::Syllable
            } else {
                paragraph_eojeol_wrap(
                    cell.breaks_hangul_at_eojeol,
                    style,
                    cell_line_box_em,
                    cell.available_measure_pt,
                )
            };
            if cell.uses_powerpoint_line_box {
                generate_powerpoint_runs_with_tabs(
                    out,
                    &para.runs,
                    style,
                    style.tab_stops.as_deref(),
                    paragraph_default_tab_width_pt(style, cell.default_tab_width_pt),
                    eojeol_wrap,
                    cell.available_measure_pt,
                    true,
                );
                // PowerPoint includes the final glyph's `a:rPr/@spc` when it
                // places a centred or right-aligned table-cell line, just as
                // it does in a text box. The run generator reserves that space
                // on lines ended by hard breaks; append it here for the last
                // physical line (issue #1256).
                if let Some(spacing) = powerpoint_trailing_letter_space_pt(style, &para.runs) {
                    let _ = write!(out, "#h({}pt)", format_f64(spacing));
                }
            } else {
                generate_runs_with_tabs(
                    out,
                    &para.runs,
                    style.tab_stops.as_deref(),
                    paragraph_default_tab_width_pt(style, cell.default_tab_width_pt),
                    eojeol_wrap,
                );
                if let Some(space_pt) = sheet_trailing_advance_space_pt(style, &para.runs) {
                    let _ = write!(out, "#h({}pt)", format_geometry(space_pt));
                }
            }
        }
    }

    // Suppressed when the grid-snapped line box already contains it, or the
    // gap would be counted twice (issues #500, #503).
    // TODO(#625 follow-up: cells compose w:after + w:before additively via
    // strong #v while body flow max-collapses them; Word's in-cell rule is
    // unmeasured — probe before changing).
    if let Some(space_after) = style.space_after
        && !cell_grid_absorbs_space_after(style, cell.line_grid_pitch, cell.row_east_asian)
    {
        let _ = write!(out, "\n#v({}pt)", format_f64(space_after));
    }

    if has_block_wrapper {
        out.push_str("\n]");
    }

    if let Some((line_advance_pt, leading_pt, shift_pt, measure_width_pt)) =
        centered_sheet_odd_line_seat
    {
        let _ = write!(
            out,
            "];\n  let o2p-centered-sheet-lines = calc.round((measure(o2p-centered-sheet-body, width: {}pt).height + {}pt) / {}pt);\n  move(dy: if o2p-centered-sheet-lines > 1 and calc.rem-euclid(o2p-centered-sheet-lines, 2) == 1 {{ {}pt }} else {{ 0pt }}, o2p-centered-sheet-body)\n}}",
            format_geometry(measure_width_pt),
            format_geometry(leading_pt),
            format_geometry(line_advance_pt),
            format_geometry(shift_pt),
        );
    }
}

/// The measured line model needed to distinguish an odd wrapped spreadsheet
/// paragraph from its one- and even-line peers without predicting wrapping in
/// Rust. The returned values are `(advance, leading, shift, measure width)` in
/// printed points.
fn centered_sheet_odd_line_seat(
    para: &Paragraph,
    cell: &CellParagraphCtx,
) -> Option<(f64, f64, f64, f64)> {
    if cell.vertical_align != Some(CellVerticalAlign::Center)
        || cell.sheet_seat.is_none()
        || cell.sheet_print_scale.is_none()
        || cell.in_spill_cell
        || cell.stacks_multiple_blocks
        || para.runs.is_empty()
        || para.style.space_before.is_some()
        || para.style.space_after.is_some()
    {
        return None;
    }

    let line_box: CellLineBox = word_cell_line_box(
        &para.runs,
        &para.style,
        cell.line_grid_pitch,
        cell.row_east_asian,
        cell.vertical_align,
        cell.seats_text_on_descender,
        cell.sheet_row_line.as_ref(),
        cell.sheet_seat,
        cell.sheet_print_scale,
    )?;
    let line_advance_pt: f64 =
        (line_box.top_em + line_box.bottom_em) * line_box.font_size_pt + line_box.leading_pt;
    if line_advance_pt <= 0.0 {
        return None;
    }
    let shift_pt: f64 = cell.sheet_print_scale.filter(|scale| *scale > 0.0)?;
    let measure_width_pt: f64 = cell.available_measure_pt.filter(|width| *width > 0.0)?;
    Some((
        line_advance_pt,
        line_box.leading_pt,
        shift_pt,
        measure_width_pt,
    ))
}

fn cell_paragraph_needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.line_spacing.is_some()
        || style.line_box.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
}

fn write_cell_paragraph_block_params(
    out: &mut String,
    needs_full_width: bool,
    suppress_default_block_spacing: bool,
    indent: Option<(f64, f64)>,
) {
    let mut first = true;

    if needs_full_width {
        write_param(out, &mut first, "width: 100%");
    }
    // Stacked cell paragraphs: the inter-paragraph gap is carried entirely by
    // the explicit `#v(space_before)`/`#v(space_after)` emissions (which are
    // the resolved Word values), so the wrapper must contribute nothing —
    // Typst's default `block` spacing is 1.2em of engine whitespace Word does
    // not have (issue #625). The trailing `#v(space_after)` stays *inside* the
    // block rather than becoming a weak `below:`, because Word counts it into
    // the row height and weak spacing would vanish at the cell's edge.
    if suppress_default_block_spacing {
        write_param(out, &mut first, "above: 0pt");
        write_param(out, &mut first, "below: 0pt");
    }
    // Word's `w:ind`, as the block's own padding: the wrapper shrinks to its
    // content, so a left inset shifts the text right by exactly that much and a
    // right inset takes the same width off the measure the text wraps in
    // (issue #938).
    if let Some((left, right)) = indent {
        write_param(
            out,
            &mut first,
            &format!(
                "inset: (left: {}pt, right: {}pt)",
                format_f64(left),
                format_f64(right)
            ),
        );
    }
}
