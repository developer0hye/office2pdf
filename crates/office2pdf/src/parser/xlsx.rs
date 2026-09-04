use std::io::Cursor;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Document, ImageData, Margins, Metadata, Page, PageSize, SheetPage, StyleSheet, Table,
    TableBorderPaintModel, TableRow,
};
use crate::parser::Parser;

#[path = "xlsx_cond_fmt_raw.rs"]
pub(crate) mod cond_fmt_raw;

#[path = "xlsx_chartsheet.rs"]
mod chartsheet;
#[path = "xlsx_fit_to_page.rs"]
mod fit_to_page;
#[path = "xlsx_indent.rs"]
mod indent;
#[path = "xlsx_paper_state.rs"]
mod paper_state;
#[path = "xlsx_print_headings.rs"]
mod print_headings;
#[path = "xlsx_print_options.rs"]
mod print_options;
#[path = "xlsx_row_boundaries.rs"]
mod row_boundaries;
#[path = "xlsx_sparklines.rs"]
mod sparklines;
#[path = "xlsx_tables.rs"]
mod tables;
#[path = "xlsx_cells.rs"]
mod xlsx_cells;
#[path = "xlsx_drawing.rs"]
pub(in crate::parser) mod xlsx_drawing;
#[path = "xlsx_hf.rs"]
mod xlsx_hf;
#[path = "xlsx_pagination.rs"]
mod xlsx_pagination;
#[path = "xlsx_style.rs"]
pub(crate) mod xlsx_style;

use self::xlsx_cells::*;
use self::xlsx_drawing::*;
use self::xlsx_hf::*;

// Re-export cell address types for cond_fmt module.
pub(crate) use self::xlsx_cells::{CellPos, CellRange, parse_cell_ref};

/// Parser for XLSX (Office Open XML Excel) spreadsheets.
/// The margins Excel prints a sheet at when it states none: 0.7" left and
/// right, 0.75" top and bottom.
const DEFAULT_PRINT_MARGINS: Margins = Margins {
    top: 54.0,
    bottom: 54.0,
    left: 50.4,
    right: 50.4,
};

/// Print margins for a sheet: the worksheet's explicit `<pageMargins>` when
/// present, otherwise Excel's defaults, each snapped to the whole device point
/// Excel prints against.
///
/// umya leaves absent margin attributes at 0.0, which is not a value Excel
/// ever writes, so ≤0 means "not specified".
///
/// Excel lays a printed sheet out on whole device points, so a margin the file
/// states in inches reaches the paper floored. Measured on an Excel-for-Mac
/// export of `tests/fixtures/xlsx/ExcelTables.xlsx`, which declares a
/// 0.78740157in top (56.69pt) and a 0.7in left (50.4pt) over 17pt rows: the
/// grid boundaries are 56, 73, 90, 107 and 124 rather than 56.69, 73.69, …,
/// and the first table column starts at x 464 = 50 + six 69pt columns rather
/// than 464.4. Rounding does not fit — 56.69 would round up to 57 — and the
/// row pitch itself is untouched, so it is the origin alone that moves
/// (issue #1127).
///
/// The far margins follow the same rule, which is what puts the printed page's
/// far edges on whole points too. No native export measured here paginates
/// against a fractional bottom margin, so the ≤1pt of printable extent that
/// gains is inferred from the model rather than observed.
fn sheet_print_margins(sheet: &umya_spreadsheet::Worksheet) -> Margins {
    let page_margins = sheet.get_page_margins();
    let inches_to_printed_pt = |inches: f64, default_pt: f64| -> f64 {
        let declared_pt: f64 = if inches > 0.0 {
            inches * 72.0
        } else {
            default_pt
        };
        declared_pt.floor()
    };
    Margins {
        top: inches_to_printed_pt(*page_margins.get_top(), DEFAULT_PRINT_MARGINS.top),
        bottom: inches_to_printed_pt(*page_margins.get_bottom(), DEFAULT_PRINT_MARGINS.bottom),
        left: inches_to_printed_pt(*page_margins.get_left(), DEFAULT_PRINT_MARGINS.left),
        right: inches_to_printed_pt(*page_margins.get_right(), DEFAULT_PRINT_MARGINS.right),
    }
}

/// The footer margin Excel prints a sheet at when it states none: 0.3".
const DEFAULT_PRINT_FOOTER_MARGIN_PT: f64 = 21.6;

/// What Excel leaves between `<pageMargins>/@footer` and the bottom of the
/// footer text's line box, in points.
///
/// Measured on Excel-for-Mac exports of one-factor variants of
/// `tests/fixtures/xlsx/headerFooterTest.xlsx`. A 12pt Calibri footer over a
/// 0.5in (36pt) footer margin puts its baseline 41pt above the page's bottom
/// edge; Calibri's `hhea` descender is 0.26855em, so 36 + 2 + 3.22 = 41.22
/// lands on the whole point Excel prints it at. The same 2pt holds across the
/// series 6, 8, 12, 14, 20, 40 and 80pt, across Arial, Verdana, Times New
/// Roman, Aptos and Segoe UI, and at footer margins of 0.3, 0.5, 0.75 and
/// 1.0in (issue #1142).
const SHEET_FOOTER_BAND_INSET_PT: f64 = 2.0;

/// Where Excel seats a printed sheet footer: the distance from the page's
/// bottom edge to the bottom of the footer text's line box.
///
/// The band is measured up from the paper, through `<pageMargins>/@footer`,
/// and neither the bottom margin nor the sheet's own body height moves it —
/// both probed one factor at a time against native exports. The margin floors
/// to a whole device point for the same reason [`sheet_print_margins`] does.
fn sheet_footer_distance_from_edge(sheet: &umya_spreadsheet::Worksheet) -> f64 {
    let declared_in: f64 = *sheet.get_page_margins().get_footer();
    let footer_margin_pt: f64 = if declared_in > 0.0 {
        declared_in * 72.0
    } else {
        DEFAULT_PRINT_FOOTER_MARGIN_PT
    };
    footer_margin_pt.floor() + SHEET_FOOTER_BAND_INSET_PT
}

/// Parse a sheet's odd footer and seat it on the page's bottom edge.
fn sheet_print_footer(
    sheet: &umya_spreadsheet::Worksheet,
    sheet_name: &str,
    normal_font: Option<&xlsx_cells::NormalFont>,
    warnings: &mut Vec<ConvertWarning>,
) -> Option<crate::ir::HeaderFooter> {
    let mut footer = parse_hf_format_string(
        sheet.get_header_footer().get_odd_footer().get_value(),
        sheet_name,
        normal_font,
        warnings,
    )?;
    footer.distance_from_edge = Some(sheet_footer_distance_from_edge(sheet));
    Some(footer)
}

/// Map an OOXML worksheet paper-size code to portrait dimensions in points.
///
/// Code 0 is not a paper size — it is the zero umya leaves in an unset
/// `UInt32Value`. ECMA-376 defaults an initialised but silent `paperSize` to 1,
/// US Letter, and Excel prints such a sheet on Letter; A4 left it 16.7pt narrow
/// and 49.9pt tall, which repaginated the whole sheet (issue #717). A fully
/// pristine sheet is routed around this schema default by [`sheet_page_size`]
/// because Excel instead follows its current application paper (issue #1382).
///
/// An unrecognised *positive* code is a different case: the file names a paper
/// this table does not model, and nothing in the schema says what to
/// substitute, so those keep the renderer's A4 default.
fn worksheet_paper_size(code: u32) -> PageSize {
    let (width, height) = match code {
        0 => (612.0, 792.0),        // `paperSize` omitted — schema default
        1 | 2 => (612.0, 792.0),    // Letter / Letter Small
        3 => (792.0, 1224.0),       // Tabloid
        4 => (1224.0, 792.0),       // Ledger
        5 => (612.0, 1008.0),       // Legal
        6 => (396.0, 612.0),        // Statement
        7 => (522.0, 756.0),        // Executive
        8 => (841.89, 1190.55),     // A3
        9 | 10 => (595.28, 841.89), // A4 / A4 Small
        11 => (419.53, 595.28),     // A5
        12 => (728.50, 1031.81),    // B4 (JIS)
        13 => (515.91, 728.50),     // B5 (JIS)
        _ => return PageSize::default(),
    };
    PageSize { width, height }
}

/// What a sheet asks its fit-to-page scale to be measured against, if it asks.
///
/// `fitToWidth` and `fitToHeight` count for nothing unless `<pageSetUpPr
/// fitToPage="1"/>` is also set — Excel writes them into sheets that print at
/// 100% too (issue #530). `fit_to_page::sheets_fit_to_page` has already
/// applied that gate and ECMA-376's default of one page in each direction. A
/// zero is Excel's "as many pages as it takes", so it bounds nothing.
///
/// The row bound is measured against the whole printed sheet rather than the
/// page handed to pagination, which may hold one streaming chunk or one
/// explicit-break segment of it. The height itself is only summed when that
/// bound binds, since it walks every printed row.
fn sheet_fit(
    sheet: &umya_spreadsheet::Worksheet,
    sheet_name: &str,
    fitting_sheets: &std::collections::HashMap<String, fit_to_page::SheetFitToPage>,
    printed_rows: (u32, u32),
    ctx: &SheetContext,
) -> xlsx_pagination::SheetFit {
    let declared: Option<&fit_to_page::SheetFitToPage> = fitting_sheets.get(sheet_name);
    let bound = |pages: u32| -> Option<u32> { (pages > 0).then_some(pages) };
    let pages_tall: Option<u32> = declared.and_then(|fit| bound(fit.pages_tall));
    xlsx_pagination::SheetFit {
        pages_wide: declared.and_then(|fit| bound(fit.pages_wide)),
        pages_tall,
        sheet_height_pt: pages_tall
            .map_or(0.0, |_| printed_sheet_height_pt(sheet, printed_rows, ctx)),
    }
}

/// The printed grid height of every row a sheet prints, in points.
///
/// This is the same track a drawing anchor is measured against, not the
/// heights the worksheet holds: Excel prints its rows compacted or truncated
/// to whole device points and paginates against what it printed.
fn printed_sheet_height_pt(
    sheet: &umya_spreadsheet::Worksheet,
    (row_start, row_end): (u32, u32),
    ctx: &SheetContext,
) -> f64 {
    (row_start..=row_end)
        .map(|row| {
            xlsx_cells::printed_grid_row_height_pt(
                sheet,
                row,
                ctx.normal_font.as_ref(),
                Some(&ctx.row_boundary_points),
            )
        })
        .sum()
}

/// Whether the sheet's header and footer shrink with its fit-to-page scale.
///
/// `headerFooter/@scaleWithDoc` defaults to 1, so a sheet that states nothing —
/// including one with no `<headerFooter>` at all — scales (issue #940).
fn sheet_header_footer_scales_with_doc(
    sheet_name: &str,
    fitting_sheets: &std::collections::HashMap<String, fit_to_page::SheetFitToPage>,
) -> bool {
    fitting_sheets
        .get(sheet_name)
        .is_none_or(|fit| fit.header_footer_scales_with_doc)
}

/// Whether a worksheet contributes pages to the export.
///
/// A sheet that `xl/workbook.xml` declares `state="hidden"` — or
/// `"veryHidden"`, which additionally keeps it out of the unhide dialog — is
/// not printed by Excel: hiding a lookup sheet is the whole point of the
/// state, and neither the native export nor LibreOffice pages it (issue
/// #1065).
///
/// Naming a sheet through `--sheets` overrides that. The caller has asked for
/// that sheet by name, which is the one way to get a hidden one on paper.
fn sheet_prints(sheet: &umya_spreadsheet::Worksheet, options: &ConvertOptions) -> bool {
    match options.sheet_names.as_deref() {
        Some(names) => names.iter().any(|name| name == sheet.get_name()),
        None => !matches!(
            sheet.get_state(),
            umya_spreadsheet::SheetStateValues::Hidden
                | umya_spreadsheet::SheetStateValues::VeryHidden
        ),
    }
}

/// The first sheet that prints, honouring `sheet_names` and sheet visibility.
fn first_printed_sheet<'a>(
    book: &'a umya_spreadsheet::Spreadsheet,
    options: &ConvertOptions,
) -> Option<&'a umya_spreadsheet::Worksheet> {
    book.get_sheet_collection()
        .iter()
        .find(|sheet| sheet_prints(sheet, options))
}

/// The single page a workbook prints when none of its sheets has a used range.
///
/// The sheet loop skips a sheet with no used cells and no drawings, so such a
/// workbook reached codegen with no pages at all and the Typst default — a
/// blank A4 — stood in for it. That default answers to nothing in the file, so
/// a sheet declaring `<pageSetup paperSize="1"/>` printed on A4 (issue #632).
/// A blank page comes out either way; this one is the size the file asks for.
///
/// The page stays blank. The sheet's header and footer are deliberately not
/// carried onto it: the ground truth for an empty sheet is a blank page, and
/// Excel itself refuses to print one at all ("nothing found to print"), so
/// there is no observed behaviour that puts running text on a page with no
/// cells behind it. Rendering the paper the file asks for is what the evidence
/// supports; inventing content for it is not.
fn empty_workbook_page(
    book: &umya_spreadsheet::Spreadsheet,
    options: &ConvertOptions,
    pristine_paper_sheets: &std::collections::HashSet<String>,
) -> Option<SheetPage> {
    let sheet = first_printed_sheet(book, options)?;
    Some(SheetPage {
        name: sheet.get_name().to_string(),
        size: sheet_page_size(sheet, pristine_paper_sheets.contains(sheet.get_name())),
        margins: sheet_print_margins(sheet),
        table: Table::default(),
        header: None,
        footer: None,
        charts: Vec::new(),
        images: Vec::new(),
        text_boxes: Vec::new(),
    })
}

/// Preserve a worksheet's paper size and landscape orientation in the IR.
fn sheet_page_size(
    sheet: &umya_spreadsheet::Worksheet,
    uses_converter_paper_default: bool,
) -> PageSize {
    let page_setup = sheet.get_page_setup();
    let paper_size_code: u32 = *page_setup.get_paper_size();
    let size: PageSize = if paper_size_code == 0 && uses_converter_paper_default {
        PageSize::default()
    } else {
        worksheet_paper_size(paper_size_code)
    };
    if matches!(
        page_setup.get_orientation(),
        umya_spreadsheet::structs::OrientationValues::Landscape
    ) {
        PageSize {
            width: size.height,
            height: size.width,
        }
    } else {
        size
    }
}

/// Convert absolute print-title columns to 0-based indices within the
/// rendered column range, half-open. None when the titles fall outside it.
fn title_column_indices(print_titles: PrintTitles, ctx: &SheetContext) -> Option<(usize, usize)> {
    let (col_start, col_end) = print_titles.cols?;
    if col_end < ctx.col_start || col_start > ctx.col_end {
        return None;
    }
    let start_idx = col_start.max(ctx.col_start) - ctx.col_start;
    let end_idx = col_end.min(ctx.col_end) - ctx.col_start + 1;
    Some((start_idx as usize, end_idx as usize))
}

/// Excel's reading of a `twoCellAnchor`'s `to` `xdr:colOff`, which is a
/// position inside the `to` column and not an unbounded extension of it.
///
/// The offset counts as written while it stays inside the column. Past its far
/// edge Excel gives the whole-point part of the overrun back, leaving the
/// effective offset within half a point of the column width however far the
/// anchor overruns:
///
/// ```text
/// eff = off                              off <= width
/// eff = off - round(off - width)         otherwise
/// ```
///
/// Measured on Excel for Mac exports of `issue_1066_blip_effect_picture.xlsx`,
/// one factor per variant, reading `width of picture 1` back through
/// AppleScript: a 9-width column sweep under a fixed 963720 EMU offset, and a
/// 16-sample offset sweep over 65pt columns where the picture stops growing at
/// 830000 EMU and then wanders half a point either side of 194pt all the way
/// to 2000000. Adding the offset as written left that fixture's picture 11.00pt
/// wider than the export, the deviation issue #1102 measured and could not
/// attribute.
///
/// A whole-point quantity is being subtracted from a fractional one somewhere
/// in Excel's normalisation; the rule is empirical and no claim is made about
/// which of its internal quantities that is. The `from` offset and both row
/// offsets are left alone — no sweep covers an overrun on those, and every
/// well-formed anchor keeps its offsets inside their track anyway (issue
/// #1149).
fn normalized_to_col_offset_pt(offset_pt: f64, column_width_pt: f64) -> f64 {
    if offset_pt <= column_width_pt {
        return offset_pt;
    }
    // The overrun is positive here, so rounding away from zero is half-up -
    // the convention the rest of the column model rounds by. No sample sits
    // exactly on a half point, so the tie itself is unmeasured.
    offset_pt - (offset_pt - column_width_pt).round()
}

/// Convert a raw drawing anchor into a render-ready image: 1-indexed anchor
/// row plus a size in points resolved against the sheet's column widths and
/// row heights (twoCellAnchor) or the declared extent (oneCellAnchor).
fn anchored_image(
    anchor: xlsx_drawing::RawImageAnchor,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
) -> crate::ir::SheetImage {
    const EMU_PER_PT: f64 = 12_700.0;

    let column_width_at = |col_zero_based: u32| -> f64 {
        let col: u32 = col_zero_based + 1;
        if col >= ctx.col_start && col <= ctx.col_end {
            ctx.column_widths
                .get((col - ctx.col_start) as usize)
                .copied()
                .unwrap_or(0.0)
        } else {
            ctx.default_column_width_pt
        }
    };
    // Excel prints a drawing against the same grid it prints the cells
    // against, so an anchor spans printed tracks rather than the heights the
    // worksheet holds. The two agree wherever a workbook's grid does not
    // compact — `theme_color_drawing.xlsx` spans six 18pt rows for 108pt
    // either way (issue #460) — and part where it does: the picture of
    // `issue_1066_blip_effect_picture.xlsx` sits 96.00pt down and 112.00pt
    // tall in the worksheet over 16pt rows, and 90.00pt down and 105.00pt
    // tall in the export over the 15pt track (issue #1102).
    let row_height_at = |row_zero_based: u32| -> f64 {
        xlsx_cells::printed_grid_row_height_pt(
            sheet,
            row_zero_based + 1,
            ctx.normal_font.as_ref(),
            Some(&ctx.row_boundary_points),
        )
    };

    let (width, height): (f64, f64) =
        if let Some((to_col, to_col_off, to_row, to_row_off)) = anchor.to {
            let to_column_pt: f64 = column_width_at(to_col);
            let width: f64 = (anchor.from_col..to_col).map(column_width_at).sum::<f64>()
                - anchor.from_col_off_emu as f64 / EMU_PER_PT
                + normalized_to_col_offset_pt(to_col_off as f64 / EMU_PER_PT, to_column_pt);
            let height: f64 = (anchor.from_row..to_row).map(row_height_at).sum::<f64>()
                - anchor.from_row_off_emu as f64 / EMU_PER_PT
                + to_row_off as f64 / EMU_PER_PT;
            (width.max(1.0), height.max(1.0))
        } else if let Some((cx, cy)) = anchor.ext_emu {
            (
                (cx as f64 / EMU_PER_PT).max(1.0),
                (cy as f64 / EMU_PER_PT).max(1.0),
            )
        } else {
            (100.0, 100.0)
        };

    let x_offset_pt: f64 = (0..anchor.from_col).map(column_width_at).sum::<f64>()
        + anchor.from_col_off_emu as f64 / EMU_PER_PT;
    // Excel places a drawing at absolute worksheet coordinates, so the
    // vertical origin is the summed height of every row above the anchor
    // row plus its `xdr:rowOff` - the same geometry the width and height
    // already use (issue #474).
    let y_offset_pt: f64 = (0..anchor.from_row).map(row_height_at).sum::<f64>()
        + anchor.from_row_off_emu as f64 / EMU_PER_PT;

    let image = ImageData {
        rotation_deg: None,
        flip_h: false,
        flip_v: false,
        data: anchor.data,
        format: anchor.format,
        width: Some(width),
        height: Some(height),
        crop: None,
        stroke: None,
        alignment: None,
        clip_shape: None,
        shadow: None,
        paragraph_spacing: None,
    };
    crate::ir::SheetImage {
        anchor_row: anchor.from_row + 1,
        x_offset_pt,
        y_offset_pt,
        image,
        clip_width_pt: None,
    }
}

/// Context stand-in for sheets with no used cells, so drawing anchors can
/// still resolve against the sheet's column widths and row heights.
///
/// Such a sheet may still declare `<cols>`, and those widths are read here the
/// way `prepare_sheet_context` reads them for a populated sheet. Only a sheet
/// declaring none falls back to the default width for every column.
///
/// The column metric must come from the workbook Normal font exactly as it
/// does for populated sheets: hardcoding a 7px digit metric laid every
/// drawing-only sheet out on 44.2575pt columns while the workbook's own
/// Calibri-11 metric prices default columns at 53pt, shrinking anchors and
/// distorting picture aspect ratios (issue #620). Without a readable Normal
/// font the shared fallback inspects cell fonts, finds none on an empty
/// sheet, and keeps the legacy 5.25pt unit.
fn empty_sheet_context(
    sheet: &umya_spreadsheet::Worksheet,
    normal_font: Option<&NormalFont>,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
    row_boundary_points: Option<&row_boundaries::RowBoundaryPoints>,
) -> SheetContext {
    let unit_pt: f64 = resolve_column_unit_pt(sheet, normal_font);
    let default_width_pt: f64 = default_column_width_pt(
        declared_default_column_width(sheet),
        declared_base_column_width(sheet),
        unit_pt,
    );

    // A sheet with no used cells can still declare `<cols>`, and a drawing
    // anchored to those columns is placed against their widths. Leaving the
    // window empty priced every column at the default: on a probe declaring
    // width=20, an anchored picture came out 141pt wide at x=144.40 where a
    // reference render puts it 340pt wide at x=280.63 (issue #714).
    let declared: Vec<(u32, f64)> = sheet
        .get_column_dimensions()
        .iter()
        .map(|column| (*column.get_col_num(), *column.get_width()))
        .collect();
    let (col_start, col_end) = match (
        declared.iter().map(|(col, _)| *col).min(),
        declared.iter().map(|(col, _)| *col).max(),
    ) {
        (Some(first), Some(last)) => (first, last),
        // No `<cols>` either: keep the empty window, so every column falls
        // through to the default width as before.
        _ => (1, 0),
    };
    let column_widths: Vec<f64> = (col_start..=col_end)
        .map(|col| {
            sheet
                .get_column_dimension_by_number(&col)
                .map(|column| column_width_to_pt(*column.get_width(), unit_pt))
                .unwrap_or(default_width_pt)
        })
        .collect();

    SheetContext {
        col_start,
        col_end,
        num_cols: column_widths.len(),
        column_widths,
        default_column_width_pt: default_width_pt,
        merge_tops: std::collections::HashMap::new(),
        merge_skips: std::collections::HashSet::new(),
        cond_fmt_overrides: std::collections::HashMap::new(),
        sparklines: std::collections::HashMap::new(),
        normal_font: normal_font.cloned(),
        table_styles: Vec::new(),
        theme: theme.cloned(),
        // A sheet with no used cells has no cell to indent; drawings anchor
        // to the grid, which the indent never moves.
        cell_indents: std::collections::HashMap::new(),
        row_boundary_points: row_boundary_points.cloned().unwrap_or_default(),
        indent_unit_pt: resolve_indent_unit_pt(normal_font),
    }
}

/// Convert a raw text-box anchor into a render-ready box, sized like images.
fn anchored_text_box(
    anchor: xlsx_drawing::RawTextBoxAnchor,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
) -> crate::ir::SheetTextBox {
    let placed = anchored_image(
        xlsx_drawing::RawImageAnchor {
            from_row: anchor.geometry.from_row,
            from_col: anchor.geometry.from_col,
            from_col_off_emu: anchor.geometry.from_col_off_emu,
            from_row_off_emu: anchor.geometry.from_row_off_emu,
            to: anchor.geometry.to,
            ext_emu: anchor.geometry.ext_emu,
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
        },
        sheet,
        ctx,
    );
    crate::ir::SheetTextBox {
        anchor_row: placed.anchor_row,
        x_offset_pt: placed.x_offset_pt,
        y_offset_pt: placed.y_offset_pt,
        width: placed.image.width.unwrap_or(100.0),
        height: placed.image.height.unwrap_or(50.0),
        paragraphs: anchor.paragraphs,
        fill: anchor.fill,
        border: anchor.border,
        vertical_center: anchor.vertical_center,
    }
}

/// Convert a raw chart anchor into a render-ready sheet chart: the anchor's
/// absolute placement resolved exactly as a picture's is, or no placement at
/// all for a chart no drawing references (issue #982).
fn anchored_chart(
    anchor: xlsx_drawing::RawChartAnchor,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
) -> crate::ir::SheetChart {
    let Some(geometry) = anchor.geometry else {
        return crate::ir::SheetChart {
            anchor_row: u32::MAX,
            placement: None,
            chart: anchor.chart,
        };
    };
    let placed = anchored_image(
        xlsx_drawing::RawImageAnchor {
            from_row: geometry.from_row,
            from_col: geometry.from_col,
            from_col_off_emu: geometry.from_col_off_emu,
            from_row_off_emu: geometry.from_row_off_emu,
            to: geometry.to,
            ext_emu: geometry.ext_emu,
            data: Vec::new(),
            format: crate::ir::ImageFormat::Png,
        },
        sheet,
        ctx,
    );
    crate::ir::SheetChart {
        anchor_row: placed.anchor_row,
        placement: Some(crate::ir::SheetChartPlacement {
            x_offset_pt: placed.x_offset_pt,
            y_offset_pt: placed.y_offset_pt,
            width: placed.image.width.unwrap_or(100.0),
            height: placed.image.height.unwrap_or(50.0),
            print_scale: 1.0,
        }),
        chart: anchor.chart,
    }
}

/// The one page a chartsheet prints: its chart alone, seated inside the
/// printable area.
///
/// The chart's own drawing anchor states a position and an extent, and Excel
/// ignores both — the anchor plays no part in where the chart lands. Measured
/// on Excel for Mac 16.100 exports of `tests/fixtures/xlsx/any_sheets.xlsx`:
/// halving the drawing's `xdr:ext` and moving its `xdr:pos` each left the
/// exported page byte-identical, while widening a margin moved and resized the
/// printed chart (issue #1099). The page setup therefore fixes the chart's
/// origin. `chartsheet::printed_chart_box` also gives its deterministic
/// page-only extent; Excel's internal-grid far-edge residual remains an
/// explicit error term rather than a fitted page constant (issues #1147 and
/// #1221).
///
/// A chartsheet has no header or footer of its own in any audited package, and
/// no cells at all, so the page carries an empty grid.
fn chartsheet_page(
    sheet_name: String,
    setup: &chartsheet::ChartsheetPrintSetup,
    raw_charts: Vec<xlsx_drawing::RawChartAnchor>,
) -> SheetPage {
    let chart_box: chartsheet::ChartsheetChartBox = chartsheet::printed_chart_box(setup);
    let charts: Vec<crate::ir::SheetChart> = raw_charts
        .into_iter()
        .map(|anchor| {
            let mut chart = anchor.chart;
            chart.host = crate::ir::ChartHost::SpreadsheetChartsheet;
            crate::ir::SheetChart {
                anchor_row: 0,
                placement: Some(crate::ir::SheetChartPlacement {
                    x_offset_pt: chart_box.x_offset_pt,
                    y_offset_pt: chart_box.y_offset_pt,
                    width: chart_box.width,
                    height: chart_box.height,
                    print_scale: 1.0,
                }),
                chart,
            }
        })
        .collect();
    SheetPage {
        name: sheet_name,
        size: setup.size,
        margins: setup.margins,
        table: Table::default(),
        header: None,
        footer: None,
        charts,
        images: Vec::new(),
        text_boxes: Vec::new(),
    }
}

pub struct XlsxParser;

impl XlsxParser {
    /// Parse XLSX in streaming mode, returning one `Document` per chunk of rows.
    ///
    /// Each chunk contains a single `SheetPage` with at most `chunk_size` rows.
    /// This allows the caller to compile each chunk independently, bounding peak
    /// memory during Typst compilation.
    pub fn parse_streaming(
        &self,
        data: &[u8],
        options: &ConvertOptions,
        chunk_size: usize,
    ) -> Result<(Vec<Document>, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).map_err(|e| {
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        let metadata = extract_xlsx_metadata(&book);
        let cond_fmt_hints = cond_fmt_raw::extract_cond_fmt_hints(data);
        let sparkline_groups = sparklines::extract_sparkline_groups(data);
        let sparklines_by_sheet = sparklines::resolve_sparklines(&book, &sparkline_groups);
        // umya drops `<alignment indent="N"/>`, so the levels come from the
        // package itself (issue #1109).
        let cell_indents = indent::extract_cell_indents(data);
        // crates.io umya v2 drops `thickTop`, so read both row-boundary flags
        // from the package as one printed-track input (issue #1228).
        let row_boundary_points = row_boundaries::extract_row_boundary_points(data);
        // A `cfRule type="expression"` names the workbook's defined names
        // rather than repeating their formulas (issue #852).
        let defined_names = cond_fmt_raw::extract_defined_names(data);
        let fitting_sheets = fit_to_page::sheets_fit_to_page(data);
        let pristine_paper_sheets = paper_state::pristine_paper_sheets(data);
        let print_options_by_sheet = print_options::sheets_print_options(data);
        let mut table_styles = tables::extract_table_styles(data);
        let normal_font = extract_normal_font(data, Some(book.get_theme()));

        let chartsheet_setups = chartsheet::chartsheet_print_setups(data);
        let mut warnings = Vec::new();

        let mut chart_map = extract_charts_with_anchors(data);
        let mut image_map = extract_images_with_anchors(data, &mut warnings);
        let mut text_box_map = extract_text_boxes_with_anchors(data);

        let mut chunks = Vec::new();

        for sheet in book.get_sheet_collection() {
            // Sheets the caller excluded by name, and hidden ones nobody
            // asked for, contribute nothing.
            if !sheet_prints(sheet, options) {
                tracing::debug!(
                    sheet = sheet.get_name(),
                    state = ?sheet.get_state(),
                    "skipping sheet that does not print"
                );
                continue;
            }

            // A chartsheet holds no cells, so it never reaches the grid path.
            if let Some(setup) = chartsheet_setups.get(sheet.get_name()) {
                let sheet_name = sheet.get_name().to_string();
                let raw_charts = chart_map.remove(&sheet_name).unwrap_or_default();
                chunks.push(Document {
                    metadata: metadata.clone(),
                    pages: vec![Page::Sheet(chartsheet_page(sheet_name, setup, raw_charts))],
                    styles: StyleSheet::default(),
                });
                continue;
            }

            let Some((ctx, row_start, row_end)) = prepare_sheet_context(
                sheet,
                normal_font.as_ref(),
                cond_fmt_hints.get(sheet.get_name()),
                &defined_names,
                table_styles.remove(sheet.get_name()).unwrap_or_default(),
                Some(book.get_theme()),
                cell_indents.get(sheet.get_name()),
                row_boundary_points.get(sheet.get_name()),
                sparklines_by_sheet.get(sheet.get_name()),
            ) else {
                // A sheet without used cells can still carry drawings; give
                // its images a page instead of dropping them.
                let sheet_name = sheet.get_name().to_string();
                let raw_images = image_map.remove(&sheet_name);
                let raw_text_boxes = text_box_map.remove(&sheet_name);
                let raw_charts = chart_map.remove(&sheet_name);
                if raw_images.is_some() || raw_text_boxes.is_some() || raw_charts.is_some() {
                    let stub_ctx = empty_sheet_context(
                        sheet,
                        normal_font.as_ref(),
                        Some(book.get_theme()),
                        row_boundary_points.get(sheet.get_name()),
                    );
                    let images: Vec<crate::ir::SheetImage> = raw_images
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_image(anchor, sheet, &stub_ctx))
                        .collect();
                    let text_boxes: Vec<crate::ir::SheetTextBox> = raw_text_boxes
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_text_box(anchor, sheet, &stub_ctx))
                        .collect();
                    let charts: Vec<crate::ir::SheetChart> = raw_charts
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_chart(anchor, sheet, &stub_ctx))
                        .collect();
                    if !images.is_empty() || !text_boxes.is_empty() || !charts.is_empty() {
                        chunks.push(Document {
                            metadata: metadata.clone(),
                            // Drawings past the printable width split into
                            // page-columns as Excel prints them (issue #713).
                            pages: xlsx_pagination::split_drawing_only_page(SheetPage {
                                name: sheet_name,
                                size: sheet_page_size(
                                    sheet,
                                    pristine_paper_sheets.contains(sheet.get_name()),
                                ),
                                margins: sheet_print_margins(sheet),
                                table: Table::default(),
                                header: None,
                                footer: None,
                                charts,
                                images,
                                text_boxes,
                            })
                            .into_iter()
                            .map(Page::Sheet)
                            .collect(),
                            styles: StyleSheet::default(),
                        });
                    }
                }
                continue;
            };

            let sheet_name = sheet.get_name().to_string();
            let sheet_print_options: print_options::SheetPrintOptions = print_options_by_sheet
                .get(&sheet_name)
                .copied()
                .unwrap_or_default();

            // Extract sheet header/footer
            let hf = sheet.get_header_footer();
            let sheet_header = parse_hf_format_string(
                hf.get_odd_header().get_value(),
                &sheet_name,
                normal_font.as_ref(),
                &mut warnings,
            );
            let sheet_footer =
                sheet_print_footer(sheet, &sheet_name, normal_font.as_ref(), &mut warnings);

            // Pull charts for this sheet
            let mut sheet_charts: Vec<crate::ir::SheetChart> = chart_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_chart(anchor, sheet, &ctx))
                .collect();
            for sheet_chart in &sheet_charts {
                let title = sheet_chart
                    .chart
                    .title
                    .as_deref()
                    .unwrap_or("untitled")
                    .to_string();
                warnings.push(ConvertWarning::FallbackUsed {
                    format: "XLSX".to_string(),
                    from: format!("chart ({title})"),
                    to: "data table".to_string(),
                });
            }
            sheet_charts.sort_by_key(|sheet_chart| sheet_chart.anchor_row);
            let mut sheet_images: Vec<crate::ir::SheetImage> = image_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_image(anchor, sheet, &ctx))
                .collect();
            sheet_images.sort_by_key(|sheet_image| sheet_image.anchor_row);
            let mut sheet_text_boxes: Vec<crate::ir::SheetTextBox> = text_box_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_text_box(anchor, sheet, &ctx))
                .collect();
            sheet_text_boxes.sort_by_key(|text_box| text_box.anchor_row);

            let print_titles = find_print_titles(&book, sheet);
            let title_columns: Option<(usize, usize)> =
                print_headings::heading_adjusted_title_columns(
                    title_column_indices(print_titles, &ctx),
                    sheet_print_options.prints_headings,
                );
            let fit: xlsx_pagination::SheetFit = sheet_fit(
                sheet,
                &sheet_name,
                &fitting_sheets,
                (row_start, row_end),
                &ctx,
            );
            let header_footer_scales_with_doc: bool =
                sheet_header_footer_scales_with_doc(&sheet_name, &fitting_sheets);

            // Process rows in chunks
            let mut chunk_start = row_start;
            let mut first_chunk = true;
            while chunk_start <= row_end {
                let chunk_end = (chunk_start + chunk_size as u32 - 1).min(row_end);

                let mut rows = build_rows_for_range(sheet, &ctx, chunk_start, chunk_end);
                // Worksheet row number of each built row, for the printed
                // heading gutter (issue #623).
                let mut sheet_row_numbers: Option<Vec<u32>> = sheet_print_options
                    .prints_headings
                    .then(|| (chunk_start..=chunk_end).collect());
                let mut header_row_count: usize = 0;
                // Rows above the print-title range print once; only the title
                // rows themselves repeat.
                let mut non_repeating_header_row_count: usize = 0;
                if let Some((title_start, title_end)) = print_titles.rows
                    && title_end < chunk_start
                {
                    // Later chunks don't contain the title rows — prepend them.
                    let mut title_rows = build_rows_for_range(sheet, &ctx, title_start, title_end);
                    header_row_count = title_rows.len();
                    title_rows.append(&mut rows);
                    rows = title_rows;
                    if let Some(numbers) = sheet_row_numbers.as_mut() {
                        numbers.splice(0..0, title_start..=title_end);
                    }
                } else if let Some((title_start, title_end)) = print_titles.rows
                    && title_end >= chunk_start
                    && title_end <= chunk_end
                {
                    non_repeating_header_row_count =
                        title_start.saturating_sub(chunk_start) as usize;
                    header_row_count =
                        (title_end + 1).saturating_sub(title_start.max(chunk_start)) as usize;
                }

                let mut sheet_page = SheetPage {
                    name: sheet_name.clone(),
                    size: sheet_page_size(sheet, pristine_paper_sheets.contains(sheet.get_name())),
                    margins: sheet_print_margins(sheet),
                    table: Table {
                        rows,
                        column_widths: ctx.column_widths.clone(),
                        header_row_count,
                        non_repeating_header_row_count,
                        alignment: None,
                        default_cell_padding: Some(xlsx_cells::default_cell_padding(
                            ctx.normal_font.as_ref(),
                        )),
                        use_content_driven_row_heights: false,
                        default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
                        seats_bottom_aligned_text_on_descender: true,
                        bottom_aligned_descent_floor_pt: bottom_aligned_descent_floor_pt(
                            ctx.normal_font.as_ref(),
                        ),
                        border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
                        prints_gridlines: sheet_print_options.prints_gridlines,
                        prints_headings: false,
                        centers_between_print_margins: sheet_print_options.centers_horizontally,
                        print_scale: None,
                    },
                    header: sheet_header.clone(),
                    footer: sheet_footer.clone(),
                    charts: if first_chunk {
                        std::mem::take(&mut sheet_charts)
                    } else {
                        vec![]
                    },
                    images: if first_chunk {
                        std::mem::take(&mut sheet_images)
                    } else {
                        vec![]
                    },
                    text_boxes: if first_chunk {
                        first_chunk = false;
                        std::mem::take(&mut sheet_text_boxes)
                    } else {
                        vec![]
                    },
                };
                if let Some(numbers) = sheet_row_numbers.as_deref() {
                    print_headings::augment_page_with_print_headings(
                        &mut sheet_page,
                        numbers,
                        ctx.col_start,
                        normal_font.as_ref(),
                    );
                }
                let doc = Document {
                    metadata: metadata.clone(),
                    pages: xlsx_pagination::split_sheet_page_by_width(
                        sheet_page,
                        title_columns,
                        fit,
                        header_footer_scales_with_doc,
                    )
                    .into_iter()
                    .map(Page::Sheet)
                    .collect(),
                    styles: StyleSheet::default(),
                };

                chunks.push(doc);
                chunk_start = chunk_end + 1;
            }
        }

        if chunks.is_empty()
            && let Some(page) = empty_workbook_page(&book, options, &pristine_paper_sheets)
        {
            chunks.push(Document {
                metadata,
                pages: vec![Page::Sheet(page)],
                styles: StyleSheet::default(),
            });
        }

        Ok((chunks, warnings))
    }
}

impl Parser for XlsxParser {
    fn parse(
        &self,
        data: &[u8],
        options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        let cursor = Cursor::new(data);
        let book = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).map_err(|e| {
            crate::parser::parse_err(format!("Failed to parse XLSX (umya-spreadsheet): {e}"))
        })?;

        // Extract metadata from umya-spreadsheet properties
        let metadata = extract_xlsx_metadata(&book);
        let cond_fmt_hints = cond_fmt_raw::extract_cond_fmt_hints(data);
        let sparkline_groups = sparklines::extract_sparkline_groups(data);
        let sparklines_by_sheet = sparklines::resolve_sparklines(&book, &sparkline_groups);
        // umya drops `<alignment indent="N"/>`, so the levels come from the
        // package itself (issue #1109).
        let cell_indents = indent::extract_cell_indents(data);
        // crates.io umya v2 drops `thickTop`, so read both row-boundary flags
        // from the package as one printed-track input (issue #1228).
        let row_boundary_points = row_boundaries::extract_row_boundary_points(data);
        // A `cfRule type="expression"` names the workbook's defined names
        // rather than repeating their formulas (issue #852).
        let defined_names = cond_fmt_raw::extract_defined_names(data);
        let fitting_sheets = fit_to_page::sheets_fit_to_page(data);
        let pristine_paper_sheets = paper_state::pristine_paper_sheets(data);
        let print_options_by_sheet = print_options::sheets_print_options(data);
        let mut table_styles = tables::extract_table_styles(data);
        let normal_font = extract_normal_font(data, Some(book.get_theme()));

        let chartsheet_setups = chartsheet::chartsheet_print_setups(data);
        let mut warnings = Vec::new();

        // Extract charts with anchor positions per sheet
        let mut chart_map = extract_charts_with_anchors(data);
        let mut image_map = extract_images_with_anchors(data, &mut warnings);
        let mut text_box_map = extract_text_boxes_with_anchors(data);

        let sheet_count = book.get_sheet_collection().len();
        let mut pages = Vec::with_capacity(sheet_count);

        for sheet in book.get_sheet_collection() {
            // Sheets the caller excluded by name, and hidden ones nobody
            // asked for, contribute nothing.
            if !sheet_prints(sheet, options) {
                tracing::debug!(
                    sheet = sheet.get_name(),
                    state = ?sheet.get_state(),
                    "skipping sheet that does not print"
                );
                continue;
            }

            // A chartsheet holds no cells, so it never reaches the grid path.
            if let Some(setup) = chartsheet_setups.get(sheet.get_name()) {
                let sheet_name = sheet.get_name().to_string();
                let raw_charts = chart_map.remove(&sheet_name).unwrap_or_default();
                pages.push(Page::Sheet(chartsheet_page(sheet_name, setup, raw_charts)));
                continue;
            }

            let Some((ctx, row_start, row_end)) = prepare_sheet_context(
                sheet,
                normal_font.as_ref(),
                cond_fmt_hints.get(sheet.get_name()),
                &defined_names,
                table_styles.remove(sheet.get_name()).unwrap_or_default(),
                Some(book.get_theme()),
                cell_indents.get(sheet.get_name()),
                row_boundary_points.get(sheet.get_name()),
                sparklines_by_sheet.get(sheet.get_name()),
            ) else {
                // A sheet without used cells can still carry drawings; give
                // its images a page instead of dropping them.
                let sheet_name = sheet.get_name().to_string();
                let raw_images = image_map.remove(&sheet_name);
                let raw_text_boxes = text_box_map.remove(&sheet_name);
                let raw_charts = chart_map.remove(&sheet_name);
                if raw_images.is_some() || raw_text_boxes.is_some() || raw_charts.is_some() {
                    let stub_ctx = empty_sheet_context(
                        sheet,
                        normal_font.as_ref(),
                        Some(book.get_theme()),
                        row_boundary_points.get(sheet.get_name()),
                    );
                    let images: Vec<crate::ir::SheetImage> = raw_images
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_image(anchor, sheet, &stub_ctx))
                        .collect();
                    let text_boxes: Vec<crate::ir::SheetTextBox> = raw_text_boxes
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_text_box(anchor, sheet, &stub_ctx))
                        .collect();
                    let charts: Vec<crate::ir::SheetChart> = raw_charts
                        .unwrap_or_default()
                        .into_iter()
                        .map(|anchor| anchored_chart(anchor, sheet, &stub_ctx))
                        .collect();
                    if !images.is_empty() || !text_boxes.is_empty() || !charts.is_empty() {
                        // Drawings past the printable width split into
                        // page-columns as Excel prints them (issue #713).
                        pages.extend(
                            xlsx_pagination::split_drawing_only_page(SheetPage {
                                name: sheet_name,
                                size: sheet_page_size(
                                    sheet,
                                    pristine_paper_sheets.contains(sheet.get_name()),
                                ),
                                margins: sheet_print_margins(sheet),
                                table: Table::default(),
                                header: None,
                                footer: None,
                                charts,
                                images,
                                text_boxes,
                            })
                            .into_iter()
                            .map(Page::Sheet),
                        );
                    }
                }
                continue;
            };

            let rows = build_rows_for_range(sheet, &ctx, row_start, row_end);

            let sheet_name = sheet.get_name().to_string();
            let sheet_print_options: print_options::SheetPrintOptions = print_options_by_sheet
                .get(&sheet_name)
                .copied()
                .unwrap_or_default();

            let print_titles = find_print_titles(&book, sheet);
            let title_columns: Option<(usize, usize)> =
                print_headings::heading_adjusted_title_columns(
                    title_column_indices(print_titles, &ctx),
                    sheet_print_options.prints_headings,
                );
            let fit: xlsx_pagination::SheetFit = sheet_fit(
                sheet,
                sheet.get_name(),
                &fitting_sheets,
                (row_start, row_end),
                &ctx,
            );
            let header_footer_scales_with_doc: bool =
                sheet_header_footer_scales_with_doc(sheet.get_name(), &fitting_sheets);
            // Only the rows named by `_xlnm.Print_Titles` repeat on later
            // pages. Rows above them still lead the table, but print once, so
            // they go into a non-repeating header block.
            let (non_repeating_header_row_count, header_row_count): (usize, usize) = print_titles
                .rows
                .filter(|(_, title_end)| *title_end >= row_start)
                .map(|(title_start, title_end)| {
                    let lead: usize = title_start.saturating_sub(row_start) as usize;
                    let repeat: usize = (title_end.min(row_end) + 1)
                        .saturating_sub(title_start.max(row_start))
                        as usize;
                    (lead, repeat)
                })
                .unwrap_or((0, 0));

            // Collect row page breaks and split rows into page segments
            let row_breaks = collect_row_breaks(sheet);

            // Extract sheet header/footer
            let hf = sheet.get_header_footer();
            let sheet_header = parse_hf_format_string(
                hf.get_odd_header().get_value(),
                &sheet_name,
                normal_font.as_ref(),
                &mut warnings,
            );
            let sheet_footer =
                sheet_print_footer(sheet, &sheet_name, normal_font.as_ref(), &mut warnings);

            // Pull charts for this sheet (if any)
            let mut sheet_charts: Vec<crate::ir::SheetChart> = chart_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_chart(anchor, sheet, &ctx))
                .collect();
            for sheet_chart in &sheet_charts {
                let title = sheet_chart
                    .chart
                    .title
                    .as_deref()
                    .unwrap_or("untitled")
                    .to_string();
                warnings.push(ConvertWarning::FallbackUsed {
                    format: "XLSX".to_string(),
                    from: format!("chart ({title})"),
                    to: "data table".to_string(),
                });
            }
            // Sort by anchor row
            sheet_charts.sort_by_key(|sheet_chart| sheet_chart.anchor_row);
            let mut sheet_images: Vec<crate::ir::SheetImage> = image_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_image(anchor, sheet, &ctx))
                .collect();
            sheet_images.sort_by_key(|sheet_image| sheet_image.anchor_row);
            let mut sheet_text_boxes: Vec<crate::ir::SheetTextBox> = text_box_map
                .remove(&sheet_name)
                .unwrap_or_default()
                .into_iter()
                .map(|anchor| anchored_text_box(anchor, sheet, &ctx))
                .collect();
            sheet_text_boxes.sort_by_key(|text_box| text_box.anchor_row);

            if row_breaks.is_empty() {
                // No page breaks — single page
                let sheet_row_numbers: Option<Vec<u32>> = sheet_print_options
                    .prints_headings
                    .then(|| (row_start..=row_end).collect());
                let mut sheet_page = SheetPage {
                    name: sheet_name,
                    size: sheet_page_size(sheet, pristine_paper_sheets.contains(sheet.get_name())),
                    margins: sheet_print_margins(sheet),
                    table: Table {
                        rows,
                        column_widths: ctx.column_widths.clone(),
                        header_row_count,
                        non_repeating_header_row_count,
                        alignment: None,
                        default_cell_padding: Some(xlsx_cells::default_cell_padding(
                            ctx.normal_font.as_ref(),
                        )),
                        use_content_driven_row_heights: false,
                        default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
                        seats_bottom_aligned_text_on_descender: true,
                        bottom_aligned_descent_floor_pt: bottom_aligned_descent_floor_pt(
                            ctx.normal_font.as_ref(),
                        ),
                        border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
                        prints_gridlines: sheet_print_options.prints_gridlines,
                        prints_headings: false,
                        centers_between_print_margins: sheet_print_options.centers_horizontally,
                        print_scale: None,
                    },
                    header: sheet_header.clone(),
                    footer: sheet_footer.clone(),
                    charts: sheet_charts,
                    images: sheet_images,
                    text_boxes: sheet_text_boxes,
                };
                if let Some(numbers) = sheet_row_numbers.as_deref() {
                    print_headings::augment_page_with_print_headings(
                        &mut sheet_page,
                        numbers,
                        ctx.col_start,
                        normal_font.as_ref(),
                    );
                }
                pages.extend(
                    xlsx_pagination::split_sheet_page_by_width(
                        sheet_page,
                        title_columns,
                        fit,
                        header_footer_scales_with_doc,
                    )
                    .into_iter()
                    .map(Page::Sheet),
                );
            } else {
                // Split rows at break points
                // Breaks are 1-indexed row numbers; break after that row
                let mut segments: Vec<(u32, Vec<TableRow>)> = Vec::new();
                let mut current_segment: Vec<TableRow> = Vec::new();
                let mut current_segment_start: u32 = row_start;
                let mut break_idx = 0;

                for (i, row) in rows.into_iter().enumerate() {
                    let actual_row = row_start + i as u32; // 1-indexed row number
                    if current_segment.is_empty() {
                        current_segment_start = actual_row;
                    }
                    current_segment.push(row);

                    // Check if this row is a break point
                    if break_idx < row_breaks.len() && actual_row == row_breaks[break_idx] {
                        segments
                            .push((current_segment_start, std::mem::take(&mut current_segment)));
                        break_idx += 1;
                    }
                }
                // Push remaining rows as the last segment
                if !current_segment.is_empty() {
                    segments.push((current_segment_start, current_segment));
                }

                // For page-break segments, attach all charts to the first segment
                let mut first_segment = true;
                for (segment_start_row, mut segment) in segments {
                    let mut segment_header_rows: usize = 0;
                    let mut segment_lead_rows: usize = 0;
                    // Title rows a later segment repeats, with their original
                    // worksheet numbers for the heading gutter.
                    let mut prepended_title_range: Option<(u32, u32)> = None;
                    if first_segment {
                        segment_lead_rows = non_repeating_header_row_count.min(segment.len());
                        segment_header_rows =
                            header_row_count.min(segment.len() - segment_lead_rows);
                    } else if let Some((title_start, title_end)) = print_titles.rows
                        && title_end >= row_start
                    {
                        // Later segments don't contain the title rows — prepend.
                        let clamped_title_start: u32 = title_start.max(row_start);
                        let mut title_rows =
                            build_rows_for_range(sheet, &ctx, clamped_title_start, title_end);
                        segment_header_rows = title_rows.len();
                        prepended_title_range = Some((clamped_title_start, title_end));
                        title_rows.append(&mut segment);
                        segment = title_rows;
                    }
                    let sheet_row_numbers: Option<Vec<u32>> =
                        sheet_print_options.prints_headings.then(|| {
                            let prepended_rows: usize = prepended_title_range
                                .map(|(start, end)| (end - start + 1) as usize)
                                .unwrap_or(0);
                            let data_rows: u32 = (segment.len() - prepended_rows) as u32;
                            prepended_title_range
                                .map(|(start, end)| (start..=end).collect::<Vec<u32>>())
                                .unwrap_or_default()
                                .into_iter()
                                .chain(segment_start_row..segment_start_row + data_rows)
                                .collect()
                        });
                    let mut sheet_page = SheetPage {
                        name: sheet_name.clone(),
                        size: sheet_page_size(
                            sheet,
                            pristine_paper_sheets.contains(sheet.get_name()),
                        ),
                        margins: sheet_print_margins(sheet),
                        table: Table {
                            rows: segment,
                            column_widths: ctx.column_widths.clone(),
                            header_row_count: segment_header_rows,
                            non_repeating_header_row_count: segment_lead_rows,
                            alignment: None,
                            default_cell_padding: Some(xlsx_cells::default_cell_padding(
                                ctx.normal_font.as_ref(),
                            )),
                            use_content_driven_row_heights: false,
                            default_vertical_align: Some(crate::ir::CellVerticalAlign::Bottom),
                            seats_bottom_aligned_text_on_descender: true,
                            bottom_aligned_descent_floor_pt: bottom_aligned_descent_floor_pt(
                                ctx.normal_font.as_ref(),
                            ),
                            border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
                            prints_gridlines: sheet_print_options.prints_gridlines,
                            prints_headings: false,
                            centers_between_print_margins: sheet_print_options.centers_horizontally,
                            print_scale: None,
                        },
                        header: sheet_header.clone(),
                        footer: sheet_footer.clone(),
                        charts: if first_segment {
                            std::mem::take(&mut sheet_charts)
                        } else {
                            vec![]
                        },
                        images: if first_segment {
                            std::mem::take(&mut sheet_images)
                        } else {
                            vec![]
                        },
                        text_boxes: if first_segment {
                            first_segment = false;
                            std::mem::take(&mut sheet_text_boxes)
                        } else {
                            vec![]
                        },
                    };
                    if let Some(numbers) = sheet_row_numbers.as_deref() {
                        print_headings::augment_page_with_print_headings(
                            &mut sheet_page,
                            numbers,
                            ctx.col_start,
                            normal_font.as_ref(),
                        );
                    }
                    pages.extend(
                        xlsx_pagination::split_sheet_page_by_width(
                            sheet_page,
                            title_columns,
                            fit,
                            header_footer_scales_with_doc,
                        )
                        .into_iter()
                        .map(Page::Sheet),
                    );
                }
            }
        }

        if pages.is_empty()
            && let Some(page) = empty_workbook_page(&book, options, &pristine_paper_sheets)
        {
            pages.push(Page::Sheet(page));
        }

        Ok((
            Document {
                metadata,
                pages,
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

/// Extract metadata from umya-spreadsheet Properties.
/// Empty strings are converted to None.
fn extract_xlsx_metadata(book: &umya_spreadsheet::Spreadsheet) -> Metadata {
    let props = book.get_properties();
    let non_empty = |s: &str| {
        if s.is_empty() {
            None
        } else {
            Some(s.to_string())
        }
    };
    Metadata {
        title: non_empty(props.get_title()),
        author: non_empty(props.get_creator()),
        subject: non_empty(props.get_subject()),
        description: non_empty(props.get_description()),
        created: non_empty(props.get_created()),
        modified: non_empty(props.get_modified()),
    }
}

#[cfg(test)]
#[path = "xlsx_tests.rs"]
mod tests;
