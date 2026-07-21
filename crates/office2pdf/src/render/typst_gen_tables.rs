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
    let header_row_count = table.header_row_count.min(table.rows.len());
    let default_cell_padding = table.default_cell_padding.unwrap_or(Insets {
        top: 5.0,
        right: 5.0,
        bottom: 5.0,
        left: 5.0,
    });

    if header_row_count > 0 {
        out.push_str("  table.header(\n");
        generate_table_rows(
            out,
            &table.rows[..header_row_count],
            num_cols,
            &mut rowspan_remaining,
            "    ",
            default_cell_padding,
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
        default_cell_padding,
        ctx,
    )?;

    out.push_str(")\n");
    Ok(())
}

fn generate_table_rows(
    out: &mut String,
    rows: &[TableRow],
    num_cols: usize,
    rowspan_remaining: &mut [usize],
    indent: &str,
    default_cell_padding: Insets,
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
            generate_table_cell(
                out,
                cell,
                clamped_colspan,
                indent,
                default_cell_padding,
                ctx,
            )?;

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
    default_cell_padding: Insets,
    ctx: &mut GenCtx,
) -> Result<(), ConvertError> {
    let needs_cell_fn = clamped_colspan > 1
        || cell.row_span > 1
        || cell.border.is_some()
        || cell.background.is_some()
        || cell.vertical_align.is_some()
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

    if let Some(border) = &cell.border {
        write_double_border_overlays(out, border, cell.padding.unwrap_or(default_cell_padding));
    }

    if let Some(ref db) = cell.data_bar {
        // Excel draws the bar behind the value on the same line (no track),
        // with a horizontal fade of the bar color; #place keeps it out of
        // layout so the value renders on top at its normal position.
        let pct = db.fill_pct.clamp(0.0, 100.0);
        let _ = write!(
            out,
            "#place(left + horizon, box(width: {}%, height: 100%, fill: gradient.linear(rgb({}, {}, {}), rgb({}, {}, {}).lighten(70%))))",
            format_f64(pct),
            db.color.r,
            db.color.g,
            db.color.b,
            db.color.r,
            db.color.g,
            db.color.b,
        );
    }

    if let Some(ref icon) = cell.icon_text {
        // Excel draws icon set glyphs in their band color, independent of the
        // cell's font color.
        if let Some(color) = cell.icon_color {
            let _ = write!(
                out,
                "#text(fill: rgb({}, {}, {}))[{}] ",
                color.r, color.g, color.b, icon
            );
        } else {
            let _ = write!(out, "{} ", icon);
        }
    }

    if let Some(spill_width) = cell.spill_width {
        // Excel paints unwrapped text across empty right neighbors without
        // growing the row: lay the content out on one clipped line via
        // #place (out of layout) and hold the row height with a zero-width
        // strut.
        let _ = write!(
            out,
            "#place(left + horizon, box(width: {}pt, height: 1.3em, clip: true)[",
            format_f64(spill_width),
        );
        generate_cell_content(out, &cell.content, ctx)?;
        out.push_str("])#box(width: 0pt, height: 1.3em)");
    } else {
        generate_cell_content(out, &cell.content, ctx)?;
    }
    out.push_str("],\n");
    Ok(())
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
        "#place({align}, dx: {}pt, dy: {}pt, line(length: 100% + {}pt, angle: {angle}, stroke: {}pt + rgb({}, {}, {})))",
        format_geometry(dx),
        format_geometry(dy),
        format_geometry(length_extra),
        format_geometry(side.width),
        side.color.r,
        side.color.g,
        side.color.b,
    );
}

fn format_geometry(value: f64) -> String {
    let rounded = (value * 1_000.0).round() / 1_000.0;
    format_f64(if rounded == -0.0 { 0.0 } else { rounded })
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
    if let Some(ref va) = cell.vertical_align {
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
    let base = format!(
        "{}pt + rgb({}, {}, {})",
        format_f64(side.width),
        side.color.r,
        side.color.g,
        side.color.b
    );
    match side.style {
        BorderLineStyle::Solid | BorderLineStyle::Double | BorderLineStyle::None => base,
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
            Block::InlineImages(images) => {
                for image in images {
                    generate_image(out, image, ctx);
                }
            }
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
        write_line_box_settings(out, style.line_box);
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
        || style.line_box.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
}

fn write_cell_paragraph_block_params(out: &mut String, needs_full_width: bool) {
    let mut first = true;

    if needs_full_width {
        write_param(out, &mut first, "width: 100%");
    }
}
