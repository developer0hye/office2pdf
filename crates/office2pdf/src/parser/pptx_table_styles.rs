use super::*;

// ── Table style data structures ─────────────────────────────────────────

/// Styling for a table cell region (e.g., firstRow, band1H, wholeTbl).
#[derive(Debug, Clone, Default)]
pub(super) struct TableCellRegionStyle {
    pub(super) fill: Option<Color>,
    pub(super) text_color: Option<Color>,
    pub(super) text_bold: Option<bool>,
}

/// Parsed definition of a single `<a:tblStyle>` element.
#[derive(Debug, Clone, Default)]
pub(super) struct PptxTableStyleDef {
    pub(super) whole_table: Option<TableCellRegionStyle>,
    pub(super) band1_h: Option<TableCellRegionStyle>,
    pub(super) band2_h: Option<TableCellRegionStyle>,
    pub(super) first_row: Option<TableCellRegionStyle>,
    pub(super) last_row: Option<TableCellRegionStyle>,
    pub(super) first_col: Option<TableCellRegionStyle>,
    pub(super) last_col: Option<TableCellRegionStyle>,
}

/// Map from style ID (GUID string) to parsed table style definition.
pub(super) type TableStyleMap = HashMap<String, PptxTableStyleDef>;

/// Attributes from `<a:tblPr>` that control which style regions are active.
#[derive(Debug, Clone, Default)]
pub(super) struct PptxTableProps {
    pub(super) style_id: Option<String>,
    pub(super) first_row: bool,
    pub(super) last_row: bool,
    pub(super) first_col: bool,
    pub(super) last_col: bool,
    pub(super) band_row: bool,
    pub(super) band_col: bool,
}

// ── Parsing ─────────────────────────────────────────────────────────────

/// Parse `ppt/tableStyles.xml` into a map of table style definitions.
pub(super) fn parse_table_styles_xml(
    xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> TableStyleMap {
    let mut styles: TableStyleMap = HashMap::new();
    let mut reader = Reader::from_str(xml);

    let mut current_style_id: Option<String> = None;
    let mut current_def = PptxTableStyleDef::default();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"tblStyle" => {
                    current_style_id = get_attr_str(e, b"styleId");
                    current_def = PptxTableStyleDef::default();
                }
                b"wholeTbl" if current_style_id.is_some() => {
                    current_def.whole_table = Some(parse_region_style(
                        &mut reader,
                        b"wholeTbl",
                        theme,
                        color_map,
                    ));
                }
                b"band1H" if current_style_id.is_some() => {
                    current_def.band1_h =
                        Some(parse_region_style(&mut reader, b"band1H", theme, color_map));
                }
                b"band2H" if current_style_id.is_some() => {
                    current_def.band2_h =
                        Some(parse_region_style(&mut reader, b"band2H", theme, color_map));
                }
                b"firstRow" if current_style_id.is_some() => {
                    current_def.first_row = Some(parse_region_style(
                        &mut reader,
                        b"firstRow",
                        theme,
                        color_map,
                    ));
                }
                b"lastRow" if current_style_id.is_some() => {
                    current_def.last_row = Some(parse_region_style(
                        &mut reader,
                        b"lastRow",
                        theme,
                        color_map,
                    ));
                }
                b"firstCol" if current_style_id.is_some() => {
                    current_def.first_col = Some(parse_region_style(
                        &mut reader,
                        b"firstCol",
                        theme,
                        color_map,
                    ));
                }
                b"lastCol" if current_style_id.is_some() => {
                    current_def.last_col = Some(parse_region_style(
                        &mut reader,
                        b"lastCol",
                        theme,
                        color_map,
                    ));
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tblStyle" => {
                if let Some(id) = current_style_id.take() {
                    styles.insert(id, std::mem::take(&mut current_def));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    styles
}

/// Parse a region element (e.g., `<a:firstRow>`) and extract fill, text color, and bold.
fn parse_region_style(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> TableCellRegionStyle {
    let mut style = TableCellRegionStyle::default();
    let mut in_tc_style = false;
    let mut in_tc_tx_style = false;
    let mut in_fill = false;
    let mut in_solid_fill = false;
    let mut in_font_ref = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"tcStyle" => in_tc_style = true,
                b"tcTxStyle" => {
                    in_tc_tx_style = true;
                    if let Some(bold) = get_attr_str(e, b"b") {
                        style.text_bold = Some(bold == "on");
                    }
                }
                b"fill" if in_tc_style => in_fill = true,
                b"solidFill" if in_fill || in_tc_style => in_solid_fill = true,
                b"fontRef" if in_tc_tx_style => in_font_ref = true,
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_solid_fill => {
                    let parsed: ParsedColor = parse_color_from_start(reader, e, theme, color_map);
                    style.fill = parsed.color;
                }
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_font_ref => {
                    let parsed: ParsedColor = parse_color_from_start(reader, e, theme, color_map);
                    style.text_color = parsed.color;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_solid_fill => {
                    let parsed: ParsedColor = parse_color_from_empty(e, theme, color_map);
                    style.fill = parsed.color;
                }
                b"srgbClr" | b"schemeClr" | b"sysClr" if in_font_ref => {
                    let parsed: ParsedColor = parse_color_from_empty(e, theme, color_map);
                    style.text_color = parsed.color;
                }
                b"tcTxStyle" => {
                    if let Some(bold) = get_attr_str(e, b"b") {
                        style.text_bold = Some(bold == "on");
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == end_tag {
                    break;
                }
                match local.as_ref() {
                    b"tcStyle" => in_tc_style = false,
                    b"tcTxStyle" => in_tc_tx_style = false,
                    b"fill" => in_fill = false,
                    b"solidFill" => in_solid_fill = false,
                    b"fontRef" => in_font_ref = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    style
}

// ── Style application ───────────────────────────────────────────────────

/// Apply table style colors/formatting to cells that don't have explicit overrides.
///
/// Priority (highest wins): cell-level explicit → firstRow/lastRow/firstCol/lastCol → band → wholeTbl
pub(super) fn apply_table_style(table: &mut Table, props: &PptxTableProps, styles: &TableStyleMap) {
    let style_id: &str = match props.style_id.as_deref() {
        Some(id) => id,
        None => return,
    };
    let style_def: &PptxTableStyleDef = match styles.get(style_id) {
        Some(def) => def,
        None => return,
    };

    let total_rows: usize = table.rows.len();
    let total_cols: usize = table.column_widths.len();
    let header_rows: usize = if props.first_row { 1 } else { 0 };
    let footer_rows: usize = if props.last_row { 1 } else { 0 };

    for (row_idx, row) in table.rows.iter_mut().enumerate() {
        let is_first_row: bool = props.first_row && row_idx < header_rows;
        let is_last_row: bool =
            props.last_row && total_rows > header_rows && row_idx == total_rows - 1;

        // Data row index for banding (excludes first/last special rows)
        let data_row_idx: Option<usize> = if !is_first_row && !is_last_row {
            Some(row_idx.saturating_sub(header_rows))
        } else {
            None
        };

        for (col_idx, cell) in row.cells.iter_mut().enumerate() {
            let is_first_col: bool = props.first_col && col_idx == 0;
            let is_last_col: bool = props.last_col && total_cols > 0 && col_idx == total_cols - 1;

            // Determine which region style applies (highest priority first)
            let region_style: Option<&TableCellRegionStyle> = if is_first_row {
                style_def.first_row.as_ref()
            } else if is_last_row {
                style_def.last_row.as_ref()
            } else if is_first_col {
                style_def.first_col.as_ref()
            } else if is_last_col {
                style_def.last_col.as_ref()
            } else if props.band_row
                && let Some(data_idx) = data_row_idx
            {
                if data_idx % 2 == 0 {
                    style_def.band1_h.as_ref()
                } else {
                    style_def.band2_h.as_ref()
                }
            } else {
                style_def.whole_table.as_ref()
            };

            let Some(region) = region_style else {
                // Fall back to wholeTbl if the region is defined but has no style
                if let Some(whole) = style_def.whole_table.as_ref() {
                    apply_region_to_cell(cell, whole);
                }
                continue;
            };

            apply_region_to_cell(cell, region);
        }
    }

    // Suppress footer_rows warning
    let _ = footer_rows;
}

/// Apply a region style to a cell, respecting explicit cell-level overrides.
fn apply_region_to_cell(cell: &mut TableCell, region: &TableCellRegionStyle) {
    // Only apply fill if cell doesn't have an explicit background
    if cell.background.is_none() {
        cell.background = region.fill;
    }

    // Apply text color and bold to all runs that don't have explicit overrides
    if region.text_color.is_some() || region.text_bold.is_some() {
        for block in &mut cell.content {
            if let Block::Paragraph(paragraph) = block {
                for run in &mut paragraph.runs {
                    if region.text_color.is_some() && run.style.color.is_none() {
                        run.style.color = region.text_color;
                    }
                    if let Some(bold) = region.text_bold
                        && run.style.bold.is_none()
                    {
                        run.style.bold = Some(bold);
                    }
                }
            }
        }
    }
}
