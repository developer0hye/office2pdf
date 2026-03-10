use super::{
    Alignment, BidiContext, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign,
    Color, DrawingTextBoxContext, HyperlinkMap, ImageMap, Insets, MAX_TABLE_DEPTH, NoteContext,
    SmallCapsContext, StyleMap, Table, TableCell, TableHeaderContext, TableRow, VmlTextBoxContext,
    WrapContext, convert_paragraph_blocks, parse_hex_color,
};

#[derive(Clone)]
struct RawCell {
    content: Vec<Block>,
    col_span: u32,
    col_index: usize,
    preferred_width: Option<f64>,
    vmerge: Option<String>,
    border: Option<CellBorder>,
    background: Option<Color>,
    vertical_align: Option<CellVerticalAlign>,
    padding: Option<Insets>,
}

struct RawRow {
    cells: Vec<RawCell>,
    height: Option<f64>,
}

fn extract_margin_side_points(side_json: &serde_json::Value) -> Option<f64> {
    let width_type = side_json
        .get("widthType")
        .and_then(|v| v.as_str())
        .unwrap_or("dxa");
    let value = side_json.get("val").and_then(|v| v.as_f64())?;

    match width_type {
        "dxa" => Some(value / 20.0),
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
    extract_insets_from_margins_json(margins_json)?;
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
        "dxa" => Some(width / 20.0),
        _ => None,
    }
}

fn extract_table_width_points(prop_json: Option<&serde_json::Value>) -> Option<f64> {
    let width_json = prop_json.and_then(|json| json.get("width"))?;
    let width_type = width_json
        .get("widthType")
        .and_then(|value| value.as_str())?;
    if width_type != "dxa" {
        return None;
    }
    width_json
        .get("width")
        .and_then(|value| value.as_f64())
        .map(|width| width / 20.0)
}

fn is_auto_table_width(prop_json: Option<&serde_json::Value>) -> bool {
    prop_json
        .and_then(|json| json.get("width"))
        .and_then(|width| width.get("widthType"))
        .and_then(|value| value.as_str())
        == Some("auto")
}

fn has_placeholder_autofit_grid(
    table: &docx_rs::Table,
    prop_json: Option<&serde_json::Value>,
) -> bool {
    if table.grid.is_empty() || !is_auto_table_width(prop_json) {
        return false;
    }

    // Some generators emit tiny placeholder gridCol widths for auto-fit tables.
    const PLACEHOLDER_GRID_MAX_TWIPS: usize = 200;
    table
        .grid
        .iter()
        .all(|column_width_twips| *column_width_twips <= PLACEHOLDER_GRID_MAX_TWIPS)
}

#[allow(clippy::too_many_arguments)]
pub(super) fn convert_table(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
) -> Table {
    let header_info = table_headers.consume_next();
    let table_prop_json = serde_json::to_value(&table.property).ok();
    let alignment = extract_table_alignment(table_prop_json.as_ref());
    let default_cell_padding = extract_table_default_cell_padding(table_prop_json.as_ref());

    let raw_rows = extract_raw_rows(
        table,
        images,
        hyperlinks,
        style_map,
        notes,
        wraps,
        drawing_text_boxes,
        table_headers,
        vml_text_boxes,
        bidi,
        small_caps,
        depth,
        default_cell_padding,
    );

    let inferred_column_widths: Option<Vec<f64>> = derive_column_widths_from_cells(&raw_rows);
    let mut grid_widths_points: Vec<f64> = table
        .grid
        .iter()
        .map(|&width| width as f64 / 20.0)
        .collect();

    let column_widths: Vec<f64> = if table.grid.is_empty()
        || has_placeholder_autofit_grid(table, table_prop_json.as_ref())
    {
        inferred_column_widths.unwrap_or_default()
    } else if let Some(table_width_points) = extract_table_width_points(table_prop_json.as_ref()) {
        let grid_total_points: f64 = grid_widths_points.iter().sum();
        if grid_total_points > 0.0 {
            let scale = table_width_points / grid_total_points;
            if (scale - 1.0).abs() > 0.05 {
                for width in &mut grid_widths_points {
                    *width *= scale;
                }
            }
            grid_widths_points
        } else if let Some(widths) = inferred_column_widths {
            widths
        } else {
            let column_count = table.grid.len().max(1);
            vec![table_width_points / column_count as f64; column_count]
        }
    } else {
        grid_widths_points
    };

    let rows = resolve_vmerge_and_build_rows(&raw_rows);

    Table {
        rows,
        column_widths,
        header_row_count: header_info.repeat_rows.min(table.rows.len()),
        alignment,
        default_cell_padding,
        use_content_driven_row_heights: false,
    }
}

#[allow(clippy::too_many_arguments)]
fn extract_raw_rows(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
    default_cell_padding: Option<Insets>,
) -> Vec<RawRow> {
    let mut raw_rows: Vec<RawRow> = Vec::new();

    for table_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = table_child;
        let row_prop_json = serde_json::to_value(&row.property).ok();
        let height = row_prop_json
            .as_ref()
            .filter(|j| j.get("heightRule").and_then(|v| v.as_str()) == Some("exact"))
            .and_then(|j| j.get("rowHeight"))
            .and_then(|v| v.as_f64());
        let mut cells: Vec<RawCell> = Vec::new();
        let mut col_index: usize = 0;

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

            let content = extract_cell_content(
                cell,
                images,
                hyperlinks,
                style_map,
                notes,
                wraps,
                drawing_text_boxes,
                table_headers,
                vml_text_boxes,
                bidi,
                small_caps,
                depth,
            );
            let border = prop_json
                .as_ref()
                .and_then(|j| j.get("borders"))
                .and_then(extract_cell_borders);
            let background = prop_json
                .as_ref()
                .and_then(|j| j.get("shading"))
                .and_then(extract_cell_shading);
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
                vertical_align,
                padding,
            });

            col_index += grid_span as usize;
        }

        raw_rows.push(RawRow { cells, height });
    }

    raw_rows
}

fn derive_column_widths_from_cells(raw_rows: &[RawRow]) -> Option<Vec<f64>> {
    let num_cols = raw_rows
        .iter()
        .flat_map(|row| {
            row.cells
                .iter()
                .map(|cell| cell.col_index + cell.col_span as usize)
        })
        .max()
        .unwrap_or(0);

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
                        data_bar: None,
                        icon_text: None,
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
                        data_bar: None,
                        icon_text: None,
                        vertical_align: raw_cell.vertical_align,
                        padding: raw_cell.padding,
                    });
                }
            }
        }

        rows.push(TableRow {
            cells,
            height: raw_row.height,
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

#[allow(clippy::too_many_arguments)]
fn extract_cell_content(
    cell: &docx_rs::TableCell,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
) -> Vec<Block> {
    let mut blocks: Vec<Block> = Vec::new();
    for content in &cell.children {
        match content {
            docx_rs::TableCellContent::Paragraph(para) => {
                convert_paragraph_blocks(
                    para,
                    &mut blocks,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                );
            }
            docx_rs::TableCellContent::Table(nested_table) => {
                if depth < MAX_TABLE_DEPTH {
                    blocks.push(Block::Table(convert_table(
                        nested_table,
                        images,
                        hyperlinks,
                        style_map,
                        notes,
                        wraps,
                        drawing_text_boxes,
                        table_headers,
                        vml_text_boxes,
                        bidi,
                        small_caps,
                        depth + 1,
                    )));
                }
            }
            _ => {}
        }
    }
    blocks
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
    if fill == "auto" || fill == "FFFFFF" || fill.is_empty() {
        return None;
    }
    parse_hex_color(fill)
}
