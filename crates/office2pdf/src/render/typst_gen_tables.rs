use super::*;

pub(super) fn generate_table(
    out: &mut String,
    table: &Table,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    ctx.table_depth += 1;
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
    ctx.table_depth -= 1;
    result
}

fn generate_table_inner(
    out: &mut String,
    table: &Table,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    out.push_str("#table(\n");

    if table_has_explicit_cell_borders(table) {
        out.push_str("  stroke: none,\n");
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
    let header_row_count = table.header_row_count.min(table.rows.len());

    if header_row_count > 0 {
        out.push_str("  table.header(\n");
        generate_table_rows(
            out,
            &table.rows[..header_row_count],
            num_cols,
            &mut rowspan_remaining,
            "    ",
            ctx,
        )?;
        out.push_str("  ),\n");
    }

    generate_table_rows(
        out,
        &table.rows[header_row_count..],
        num_cols,
        &mut rowspan_remaining,
        "  ",
        ctx,
    )?;

    out.push_str(")\n");
    Ok(())
}

fn table_has_explicit_cell_borders(table: &Table) -> bool {
    table.rows.iter().any(|row| {
        row.cells
            .iter()
            .any(|cell| cell.border.as_ref().is_some_and(cell_border_has_side))
    })
}

fn cell_border_has_side(border: &CellBorder) -> bool {
    border.top.is_some()
        || border.bottom.is_some()
        || border.left.is_some()
        || border.right.is_some()
}

fn generate_table_rows(
    out: &mut String,
    rows: &[TableRow],
    num_cols: usize,
    rowspan_remaining: &mut [usize],
    indent: &str,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    for row in rows {
        for rs in rowspan_remaining.iter_mut() {
            if *rs > 0 {
                *rs -= 1;
            }
        }

        let mut col_pos: usize = 0;
        for cell in &row.cells {
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
            generate_table_cell(out, cell, clamped_colspan, indent, ctx)?;

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

    Ok(())
}

fn generate_table_cell(
    out: &mut String,
    cell: &TableCell,
    clamped_colspan: u32,
    indent: &str,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let needs_cell_fn = clamped_colspan > 1
        || cell.row_span > 1
        || cell.border.is_some()
        || cell.background.is_some()
        || cell.vertical_align.is_some()
        || cell.horizontal_align.is_some()
        || cell.padding.is_some();

    if needs_cell_fn {
        out.push_str(indent);
        out.push_str("table.cell(");
        write_cell_params(out, cell, clamped_colspan);
        out.push_str(")[");
    } else {
        out.push_str(indent);
        out.push('[');
    }

    if let Some(ref db) = cell.data_bar {
        let pct = db.fill_pct.clamp(0.0, 100.0);
        let _ = write!(
            out,
            "#box(width: 100%, height: 0.8em, fill: rgb(240, 240, 240))[#box(width: {}%, height: 100%, fill: rgb({}, {}, {}))]",
            format_f64(pct),
            db.color.r,
            db.color.g,
            db.color.b,
        );
    }

    if let Some(ref icon) = cell.icon_text {
        let _ = write!(out, "{} ", icon);
    }

    generate_cell_content(out, &cell.content, ctx)?;

    out.push_str("],\n");
    Ok(())
}

fn write_cell_params(out: &mut String, cell: &TableCell, clamped_colspan: u32) {
    let mut first = true;

    if clamped_colspan > 1 {
        write_param(out, &mut first, &format!("colspan: {clamped_colspan}"));
    }
    if cell.row_span > 1 {
        write_param(out, &mut first, &format!("rowspan: {}", cell.row_span));
    }
    if let Some(ref bg) = cell.background {
        write_param(out, &mut first, &format_color(bg));
    }
    if let Some(ref padding) = cell.padding {
        write_param(
            out,
            &mut first,
            &format!("inset: {}", format_insets(padding)),
        );
    }
    if let Some(ref border) = cell.border {
        let stroke = format_cell_stroke(border);
        if !stroke.is_empty() {
            write_param(out, &mut first, &stroke);
        }
    }
    let align_str: Option<String> = match (&cell.vertical_align, &cell.horizontal_align) {
        (Some(va), Some(ha)) => {
            let v = match va {
                CellVerticalAlign::Top => "top",
                CellVerticalAlign::Center => "horizon",
                CellVerticalAlign::Bottom => "bottom",
            };
            let h = match ha {
                CellHorizontalAlign::Left => "left",
                CellHorizontalAlign::Center => "center",
                CellHorizontalAlign::Right => "right",
            };
            Some(format!("align: {h} + {v}"))
        }
        (Some(va), None) => {
            let v = match va {
                CellVerticalAlign::Top => "top",
                CellVerticalAlign::Center => "horizon",
                CellVerticalAlign::Bottom => "bottom",
            };
            Some(format!("align: {v}"))
        }
        (None, Some(ha)) => {
            let h = match ha {
                CellHorizontalAlign::Left => "left",
                CellHorizontalAlign::Center => "center",
                CellHorizontalAlign::Right => "right",
            };
            Some(format!("align: {h}"))
        }
        (None, None) => None,
    };
    if let Some(ref a) = align_str {
        write_param(out, &mut first, a);
    }
}

fn format_cell_stroke(border: &CellBorder) -> String {
    let mut parts = Vec::with_capacity(4);

    if let Some(ref side) = border.top {
        parts.push(format!("top: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.bottom {
        parts.push(format!("bottom: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.left {
        parts.push(format!("left: {}", format_border_side(side)));
    }
    if let Some(ref side) = border.right {
        parts.push(format!("right: {}", format_border_side(side)));
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("stroke: ({})", parts.join(", "))
    }
}

fn format_border_side(side: &BorderSide) -> String {
    let base = format!(
        "{}pt + rgb({}, {}, {})",
        format_f64(side.width),
        side.color.r,
        side.color.g,
        side.color.b
    );
    match side.style {
        BorderLineStyle::Solid | BorderLineStyle::None => base,
        BorderLineStyle::Double => format!(
            "(paint: rgb({}, {}, {}), thickness: {}pt)",
            side.color.r,
            side.color.g,
            side.color.b,
            format_f64(side.width * 2.5),
        ),
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

fn generate_cell_content(
    out: &mut String,
    blocks: &[Block],
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    for (i, block) in blocks.iter().enumerate() {
        if i > 0 {
            out.push('\n');
        }
        match block {
            Block::Paragraph(para) => generate_cell_paragraph(out, para),
            Block::Table(table) => {
                if ctx.table_depth < MAX_TABLE_DEPTH {
                    generate_table(out, table, ctx)?;
                }
            }
            Block::Image(img) => generate_image(out, img, ctx),
            Block::FloatingImage(fi) => generate_floating_image(out, fi, ctx),
            Block::FloatingTextBox(ftb) => generate_floating_text_box(out, ftb, ctx)?,
            Block::FloatingShape(fs) => generate_floating_shape(out, fs),
            Block::List(list) => {
                if can_render_fixed_text_list_inline(list) {
                    generate_fixed_text_list(out, list, true, None)?;
                } else {
                    generate_list(out, list)?;
                }
            }
            Block::MathEquation(math) => generate_math_equation(out, math),
            Block::Chart(chart) => generate_chart(out, chart),
            Block::PageBreak | Block::ColumnBreak => {}
        }
    }
    Ok(())
}

fn generate_cell_paragraph(out: &mut String, para: &Paragraph) {
    let style: &ParagraphStyle = &para.style;
    let alignment = style.alignment;
    let align_str: Option<&str> = match alignment {
        Some(Alignment::Left) => Some("left"),
        Some(Alignment::Center) => Some("center"),
        Some(Alignment::Right) => Some("right"),
        _ => None,
    };
    let has_block_wrapper = cell_paragraph_needs_block_wrapper(style) || align_str.is_some();

    if has_block_wrapper {
        out.push_str("#block(");
        write_cell_paragraph_block_params(out, align_str.is_some());
        out.push_str(")[\n");
        write_par_settings(out, style);
        if let Some(align_str) = align_str {
            let _ = writeln!(out, "  #set align({align_str})");
        }
    }

    if let Some(space_before) = style.space_before {
        let _ = writeln!(out, "#v({}pt)", format_f64(space_before));
    }

    generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());

    if let Some(space_after) = style.space_after {
        let _ = write!(out, "\n#v({}pt)", format_f64(space_after));
    }

    if has_block_wrapper {
        out.push_str("\n]");
    }
}

fn cell_paragraph_needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.line_spacing.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
}

fn write_cell_paragraph_block_params(out: &mut String, needs_full_width: bool) {
    let mut first = true;

    if needs_full_width {
        write_param(out, &mut first, "width: 100%");
    }
}
