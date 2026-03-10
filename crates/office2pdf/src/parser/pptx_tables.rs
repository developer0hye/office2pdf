use super::*;

/// Parse a `<a:tbl>` element from the reader into a Table IR.
///
/// The reader should be positioned right after the `<a:tbl>` Start event.
/// Reads until the matching `</a:tbl>` End event.
pub(super) fn parse_pptx_table(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Result<Table, ConvertError> {
    let mut column_widths = Vec::new();
    let mut rows: Vec<TableRow> = Vec::new();

    let mut in_row = false;
    let mut row_height_emu: i64 = 0;
    let mut cells: Vec<TableCell> = Vec::new();

    let mut in_cell = false;
    let mut cell_col_span: u32 = 1;
    let mut cell_row_span: u32 = 1;
    let mut is_h_merge = false;
    let mut is_v_merge = false;
    let mut cell_text_entries: Vec<PptxParagraphEntry> = Vec::new();
    let mut cell_background: Option<Color> = None;
    let mut cell_vertical_align: Option<CellVerticalAlign> = None;
    let mut cell_padding: Option<Insets> = None;

    let mut in_txbody = false;
    let mut text_body_style_defaults = PptxTextBodyStyleDefaults::default();
    let mut in_para = false;
    let mut para_style = ParagraphStyle::default();
    let mut para_level: u32 = 0;
    let mut para_default_run_style = TextStyle::default();
    let mut para_end_run_style = TextStyle::default();
    let mut para_bullet_definition = PptxBulletDefinition::default();
    let mut in_line_spacing = false;
    let mut runs: Vec<Run> = Vec::new();
    let mut in_run = false;
    let mut run_style = TextStyle::default();
    let mut run_text = String::new();
    let mut in_text = false;
    let mut in_run_properties = false;
    let mut in_end_paragraph_run_properties = false;
    let mut solid_fill_context = SolidFillCtx::None;

    let mut in_table_cell_properties = false;
    let mut border_left: Option<BorderSide> = None;
    let mut border_right: Option<BorderSide> = None;
    let mut border_top: Option<BorderSide> = None;
    let mut border_bottom: Option<BorderSide> = None;
    let mut in_border_line = false;
    let mut border_line_width_emu: i64 = 0;
    let mut border_line_color: Option<Color> = None;
    let mut border_line_dash_style: BorderLineStyle = BorderLineStyle::Solid;

    #[derive(Clone, Copy, PartialEq, Eq)]
    enum BorderDir {
        None,
        Left,
        Right,
        Top,
        Bottom,
    }

    let mut current_border_dir = BorderDir::None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"gridCol" => {
                        if let Some(width) = get_attr_i64(e, b"w") {
                            column_widths.push(emu_to_pt(width));
                        }
                    }
                    b"tr" => {
                        in_row = true;
                        row_height_emu = get_attr_i64(e, b"h").unwrap_or(0);
                        cells.clear();
                    }
                    b"tc" if in_row => {
                        in_cell = true;
                        cell_col_span = get_attr_i64(e, b"gridSpan").map(|v| v as u32).unwrap_or(1);
                        cell_row_span = get_attr_i64(e, b"rowSpan").map(|v| v as u32).unwrap_or(1);
                        is_h_merge = get_attr_str(e, b"hMerge").is_some();
                        is_v_merge = get_attr_str(e, b"vMerge").is_some();
                        cell_text_entries.clear();
                        cell_background = None;
                        cell_vertical_align = None;
                        cell_padding = None;
                        in_table_cell_properties = false;
                        border_left = None;
                        border_right = None;
                        border_top = None;
                        border_bottom = None;
                    }
                    b"txBody" if in_cell => {
                        in_txbody = true;
                        text_body_style_defaults = PptxTextBodyStyleDefaults::default();
                    }
                    b"lstStyle" if in_txbody => {
                        let local_defaults = parse_pptx_list_style(reader, theme, color_map);
                        text_body_style_defaults.merge_from(&local_defaults);
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_level = 0;
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        in_line_spacing = false;
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_line_spacing = true;
                    }
                    b"tab" if in_para && !in_run => {
                        extract_pptx_tab_stop(e, &mut para_style);
                    }
                    b"spcPct" if in_line_spacing => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_line_spacing => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_context = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"r" if in_para => {
                        in_run = true;
                        run_style = para_default_run_style.clone();
                        run_text.clear();
                    }
                    b"rPr" if in_run => {
                        in_run_properties = true;
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        in_end_paragraph_run_properties = true;
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"solidFill" if in_run_properties => {
                        solid_fill_context = SolidFillCtx::RunFill;
                    }
                    b"solidFill" if in_end_paragraph_run_properties => {
                        solid_fill_context = SolidFillCtx::EndParaFill;
                    }
                    b"solidFill" if in_table_cell_properties && !in_border_line => {
                        solid_fill_context = SolidFillCtx::ShapeFill;
                    }
                    b"solidFill" if in_border_line => {
                        solid_fill_context = SolidFillCtx::LineFill;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_context != SolidFillCtx::None =>
                    {
                        let color = parse_color_from_start(reader, e, theme, color_map).color;
                        match solid_fill_context {
                            SolidFillCtx::ShapeFill => cell_background = color,
                            SolidFillCtx::LineFill => border_line_color = color,
                            SolidFillCtx::RunFill => run_style.color = color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }
                    b"tcPr" if in_cell => {
                        in_table_cell_properties = true;
                        extract_pptx_table_cell_props(
                            e,
                            &mut cell_vertical_align,
                            &mut cell_padding,
                        );
                    }
                    b"lnL" if in_table_cell_properties => {
                        in_border_line = true;
                        current_border_dir = BorderDir::Left;
                        border_line_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_line_color = None;
                        border_line_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnR" if in_table_cell_properties => {
                        in_border_line = true;
                        current_border_dir = BorderDir::Right;
                        border_line_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_line_color = None;
                        border_line_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnT" if in_table_cell_properties => {
                        in_border_line = true;
                        current_border_dir = BorderDir::Top;
                        border_line_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_line_color = None;
                        border_line_dash_style = BorderLineStyle::Solid;
                    }
                    b"lnB" if in_table_cell_properties => {
                        in_border_line = true;
                        current_border_dir = BorderDir::Bottom;
                        border_line_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        border_line_color = None;
                        border_line_dash_style = BorderLineStyle::Solid;
                    }
                    b"prstDash" if in_border_line => {
                        border_line_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"gridCol" => {
                        if let Some(width) = get_attr_i64(e, b"w") {
                            column_widths.push(emu_to_pt(width));
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_context != SolidFillCtx::None =>
                    {
                        let color = parse_color_from_empty(e, theme, color_map).color;
                        match solid_fill_context {
                            SolidFillCtx::ShapeFill => cell_background = color,
                            SolidFillCtx::LineFill => border_line_color = color,
                            SolidFillCtx::RunFill => run_style.color = color,
                            SolidFillCtx::EndParaFill => para_end_run_style.color = color,
                            SolidFillCtx::BulletFill => {
                                para_bullet_definition.color =
                                    color.map(PptxBulletColorSource::Explicit);
                            }
                            SolidFillCtx::None => {}
                        }
                    }
                    b"prstDash" if in_border_line => {
                        border_line_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"tcPr" if in_cell => {
                        extract_pptx_table_cell_props(
                            e,
                            &mut cell_vertical_align,
                            &mut cell_padding,
                        );
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_line_spacing = true;
                    }
                    b"tab" if in_para && !in_run => {
                        extract_pptx_tab_stop(e, &mut para_style);
                    }
                    b"spcPct" if in_line_spacing => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_line_spacing => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_context = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"latin" | b"ea" | b"cs" if in_run_properties => {
                        apply_typeface_to_style(e, &mut run_style, theme, true);
                    }
                    b"latin" | b"ea" | b"cs" if in_end_paragraph_run_properties => {
                        apply_typeface_to_style(e, &mut para_end_run_style, theme, true);
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_text && let Some(text) = decode_pptx_text_event(t) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(ref reference)) => {
                if in_text && let Some(text) = decode_pptx_general_ref(reference) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"tbl" => break,
                    b"tr" if in_row => {
                        let height = if row_height_emu > 0 {
                            Some(emu_to_pt(row_height_emu))
                        } else {
                            None
                        };
                        rows.push(TableRow {
                            cells: std::mem::take(&mut cells),
                            height,
                        });
                        in_row = false;
                    }
                    b"tc" if in_cell => {
                        let has_border = border_left.is_some()
                            || border_right.is_some()
                            || border_top.is_some()
                            || border_bottom.is_some();

                        let (col_span, row_span) = if is_h_merge {
                            (0, 1)
                        } else if is_v_merge {
                            (1, 0)
                        } else {
                            (cell_col_span, cell_row_span)
                        };

                        cells.push(TableCell {
                            content: group_pptx_text_blocks(std::mem::take(&mut cell_text_entries)),
                            col_span,
                            row_span,
                            border: if has_border {
                                Some(CellBorder {
                                    left: border_left.take(),
                                    right: border_right.take(),
                                    top: border_top.take(),
                                    bottom: border_bottom.take(),
                                })
                            } else {
                                None
                            },
                            background: cell_background.take(),
                            data_bar: None,
                            icon_text: None,
                            vertical_align: cell_vertical_align.take(),
                            padding: cell_padding.take(),
                        });
                        in_cell = false;
                        in_table_cell_properties = false;
                    }
                    b"txBody" if in_txbody => {
                        in_txbody = false;
                    }
                    b"p" if in_para => {
                        let resolved_list_marker = resolve_pptx_list_marker(
                            &para_bullet_definition,
                            para_level,
                            &runs,
                            &para_end_run_style,
                            &para_default_run_style,
                        );
                        let paragraph_runs = std::mem::take(&mut runs);
                        cell_text_entries.push(PptxParagraphEntry {
                            paragraph: Paragraph {
                                style: para_style.clone(),
                                runs: paragraph_runs,
                            },
                            list_marker: resolved_list_marker,
                        });
                        in_para = false;
                    }
                    b"r" if in_run => {
                        if !run_text.is_empty() {
                            push_pptx_run(
                                &mut runs,
                                Run {
                                    text: std::mem::take(&mut run_text),
                                    style: run_style.clone(),
                                    href: None,
                                    footnote: None,
                                },
                            );
                        }
                        in_run = false;
                    }
                    b"rPr" if in_run_properties => {
                        in_run_properties = false;
                    }
                    b"endParaRPr" if in_end_paragraph_run_properties => {
                        in_end_paragraph_run_properties = false;
                    }
                    b"lnSpc" if in_line_spacing => {
                        in_line_spacing = false;
                    }
                    b"solidFill" if solid_fill_context != SolidFillCtx::None => {
                        solid_fill_context = SolidFillCtx::None;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }
                    b"tcPr" if in_table_cell_properties => {
                        in_table_cell_properties = false;
                    }
                    b"lnL" | b"lnR" | b"lnT" | b"lnB" if in_border_line => {
                        if let Some(color) = border_line_color.take() {
                            let side = BorderSide {
                                width: border_line_width_emu as f64 / 12700.0,
                                color,
                                style: border_line_dash_style,
                            };
                            match current_border_dir {
                                BorderDir::Left => border_left = Some(side),
                                BorderDir::Right => border_right = Some(side),
                                BorderDir::Top => border_top = Some(side),
                                BorderDir::Bottom => border_bottom = Some(side),
                                BorderDir::None => {}
                            }
                        }
                        in_border_line = false;
                        current_border_dir = BorderDir::None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(ConvertError::Parse(format!("XML error in table: {error}"))),
            _ => {}
        }
    }

    Ok(Table {
        rows,
        column_widths,
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(default_pptx_table_cell_padding()),
        use_content_driven_row_heights: true,
    })
}

pub(super) fn scale_pptx_table_geometry_to_frame(
    table: &mut Table,
    frame_width_pt: f64,
    frame_height_pt: f64,
) {
    let intrinsic_width_pt: f64 = table.column_widths.iter().sum();
    if intrinsic_width_pt > 0.0 && frame_width_pt > 0.0 {
        let x_scale: f64 = frame_width_pt / intrinsic_width_pt;
        for width in &mut table.column_widths {
            *width *= x_scale;
        }
    }

    let intrinsic_height_pt: f64 = table.rows.iter().filter_map(|row| row.height).sum();
    if intrinsic_height_pt > 0.0 && frame_height_pt > 0.0 {
        let y_scale: f64 = frame_height_pt / intrinsic_height_pt;
        for row in &mut table.rows {
            if let Some(height) = row.height.as_mut() {
                *height *= y_scale;
            }
        }
    }
}
