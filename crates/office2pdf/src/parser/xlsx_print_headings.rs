//! Printed row/column headings for `<printOptions headings="1"/>` (issue
//! #623).
//!
//! Native Excel prints a row-number gutter down the left edge and a
//! column-letter strip across the top of every page. Both are materialized in
//! the IR here, before pagination: the gutter as a prepended first column (so
//! the numbers flow with Typst's row pagination and the width packer accounts
//! for the track) and the letter strip as `rows[0]`, which codegen re-emits
//! as a `table.header(repeat: true)`. Codegen also paints GT's 1pt black
//! print frame around the heading bands and the data grid (see
//! `resolve_boundary_painted_borders`).
//!
//! All geometry below is measured on native Excel exports of the
//! NumberFormatTests fixture (issue #623 evidence,
//! /Volumes/T7/scratch/issue-623/trace-*.xml against
//! /Volumes/T7/scratch/issue-621/evidence/gt/nft-sheet-000{1,2}.pdf).

use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Color, Insets,
    LineCap, LineJoin, Paragraph, ParagraphStyle, Run, SheetPage, TableCell, TableRow, TextStyle,
};

use super::xlsx_cells::NormalFont;
#[cfg(test)]
use super::xlsx_cells::{ThemeFontSlot, ThemeUiScriptFaces};

/// Layout width of the row-number gutter track, in points.
///
/// GT: the gutter interior is a FIXED 22pt (x=[55,77] on both measured
/// sheets, 1-digit and 3-digit row numbers alike — it does not scale with
/// digit count) and the black gutter|data separator fills [77,78], so the
/// gutter|data grid boundary sits exactly 23pt right of the left margin
/// (54pt). A 23pt track puts that boundary where GT paints it under the
/// renderer's [B, B+1] band convention (#619), reproducing the separator band
/// and the 78pt data origin exactly.
pub(super) const PRINT_HEADING_GUTTER_WIDTH_PT: f64 = 23.0;

/// Layout height of the column-letter strip track, in points.
///
/// GT: 12pt interior (device y=[73,85]) plus the 1pt black strip|data
/// separator [85,86]; the strip|data boundary sits 13pt below the top margin
/// (72pt), so a 13pt track lands the separator band at GT's device position.
pub(super) const PRINT_HEADING_STRIP_HEIGHT_PT: f64 = 13.0;

/// The gray of the 1pt rules between heading cells.
///
/// GT paints them as degenerate axial shadings whose sampled function is a
/// constant: every byte of the 1365-sample table is 0x56, i.e. flat
/// RGB(86,86,86), confirmed by raster sampling srgb(86,86,86) on both the
/// gutter and strip separators — so a flat fill reproduces them.
pub(super) const PRINT_HEADING_RULE_GRAY: Color = Color {
    r: 86,
    g: 86,
    b: 86,
};

/// Heading cell insets. GT seats the heading baseline 3pt above the bottom
/// band's bottom at every row height — descent bottom resting ON the grid
/// boundary, i.e. a ZERO effective bottom inset (issue #623 evidence: digit
/// baseline 96.0 against row boundary 98, letter baseline 83.0 against strip
/// boundary 85, Verdana-10 descent 2.1pt). Every heading cell declares a 1pt
/// bottom rule, whose half width the renderer folds into the layout inset
/// (#500/#503), so the stored bottom padding is -0.5pt to cancel it exactly.
/// Slim 0.5pt horizontal insets: three Verdana-10 digits measure 19.07pt and
/// would wrap inside the 22pt gutter interior under the default 3pt sides,
/// which leave 16pt, while GT centers them on the gutter without shrinking.
const PRINT_HEADING_CELL_PADDING: Insets = Insets {
    top: 1.0,
    right: 0.5,
    bottom: -0.5,
    left: 0.5,
};

/// The black 1pt separator between a heading track and the data grid.
///
/// GT: gutter|data [77,78]x[85,710] and strip|data [77,538]x[85,86] are pure
/// black 1pt fill bands, the same convention as the #619/#622 rules.
fn heading_separator_side() -> BorderSide {
    BorderSide {
        width: 1.0,
        color: Color::black(),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    }
}

/// The gray 1pt rule between heading cells: under every gutter cell (one per
/// row boundary) and at every column boundary inside the strip.
fn heading_rule_side() -> BorderSide {
    BorderSide {
        width: 1.0,
        color: PRINT_HEADING_RULE_GRAY,
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    }
}

/// One heading cell: centred text in the Normal face, unpainted background
/// (GT samples pure white inside every heading cell), bottom-aligned
/// EXPLICITLY so the #711 descender seat holds even where a table copy drops
/// the sheet default (GT's rule: baseline 3pt above the band bottom).
fn heading_cell(text: String, text_style: TextStyle, border: CellBorder) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                // GT centres digits and letters on the cell (measured glyph
                // centres within 0.6pt of the track centres).
                alignment: Some(Alignment::Center),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text,
                style: text_style,
                href: None,
                footnote: None,
                inline_box: None,
            }],
        })],
        border: Some(border),
        padding: Some(PRINT_HEADING_CELL_PADDING),
        vertical_align: Some(CellVerticalAlign::Bottom),
        ..TableCell::default()
    }
}

/// GT: heading digits and letters are the workbook Normal face at the Normal
/// size (Verdana Regular 10pt on the measured sheets), pure black, not bold —
/// the same face and size as default cell text, so a theme-scheme Normal font
/// puts them in the face it resolves to (issue #1380).
fn heading_text_style(normal_font: Option<&NormalFont>) -> TextStyle {
    TextStyle {
        font_family: normal_font.map(|font| font.resolved_family().to_string()),
        font_size: normal_font.map(|font| font.size_pt),
        ..TextStyle::default()
    }
}

/// Excel's column letters: bijective base-26 over A..Z (1 → A, 26 → Z,
/// 27 → AA, 703 → AAA).
pub(super) fn column_letters(column_one_based: u32) -> String {
    let mut remaining: u32 = column_one_based;
    let mut letters: Vec<u8> = Vec::new();
    while remaining > 0 {
        remaining -= 1;
        letters.push(b'A' + (remaining % 26) as u8);
        remaining /= 26;
    }
    letters.reverse();
    String::from_utf8(letters).unwrap_or_default()
}

/// Adjust the repeated print-title column range for the prepended gutter.
///
/// GT prints the gutter (with its row numbers) on every page, so on width
/// overflow pages it must repeat like a title column. Left-anchored title
/// columns extend the same contiguous block with the gutter, so the combined
/// range repeats. A non-left-anchored range cannot merge with the gutter
/// into the single contiguous range `prepend_title_columns` supports; the
/// user's title columns keep repeating at their gutter-shifted indices and
/// the gutter numbers appear on the first column group only — dropping the
/// user's repeat would regress behavior that predates the headings.
/// TODO(#623 follow-up): Excel repeats BOTH the gutter and a
/// non-left-anchored title range on overflow pages; carrying both needs
/// multi-range support in `split_sheet_page_by_width`.
pub(super) fn heading_adjusted_title_columns(
    title_columns: Option<(usize, usize)>,
    prints_headings: bool,
) -> Option<(usize, usize)> {
    if !prints_headings {
        return title_columns;
    }
    match title_columns {
        Some((0, end)) => Some((0, end + 1)),
        Some((start, end)) => Some((start + 1, end + 1)),
        None => Some((0, 1)),
    }
}

/// Materialize the printed headings on a fully built sheet page, in place.
///
/// `sheet_row_numbers` holds the 1-based worksheet row number of each
/// existing table row in order (print-title rows a later chunk/segment
/// prepended included, with their own numbers — Excel prints a repeated
/// title row under its original number). `first_sheet_column` is the 1-based
/// worksheet column of `column_widths[0]`, so a print area starting at C
/// letters its first column "C".
pub(super) fn augment_page_with_print_headings(
    page: &mut SheetPage,
    sheet_row_numbers: &[u32],
    first_sheet_column: u32,
    normal_font: Option<&NormalFont>,
) {
    debug_assert_eq!(
        page.table.rows.len(),
        sheet_row_numbers.len(),
        "each table row needs its worksheet row number"
    );
    let text_style: TextStyle = heading_text_style(normal_font);
    let data_column_count: usize = page.table.column_widths.len();

    // Gutter cells: black separator on the data side; gray rules at both row
    // boundaries (GT rules every row boundary of the gutter, frame bottom
    // included — adjacent cells' coincident bands overlap invisibly, as in
    // #622).
    for (row, row_number) in page.table.rows.iter_mut().zip(sheet_row_numbers) {
        row.cells.insert(
            0,
            heading_cell(
                row_number.to_string(),
                text_style.clone(),
                CellBorder {
                    top: Some(heading_rule_side()),
                    bottom: Some(heading_rule_side()),
                    left: None,
                    right: Some(heading_separator_side()),
                },
            ),
        );
    }

    // Corner box: empty interior; GT gives it a gray right edge (spanning the
    // strip height) and a gray bottom edge — the black separators start only
    // where the data grid does. Its top and left edges ARE the print frame,
    // which codegen paints as exterior bands on every prints_headings table
    // (the frame is paint-only there, so it never shifts layout insets).
    let mut strip_cells: Vec<TableCell> = Vec::with_capacity(data_column_count + 1);
    strip_cells.push(TableCell {
        border: Some(CellBorder {
            top: None,
            bottom: Some(heading_rule_side()),
            left: None,
            right: Some(heading_rule_side()),
        }),
        padding: Some(PRINT_HEADING_CELL_PADDING),
        ..TableCell::default()
    });
    // Letter cells: black separator under each (they join into the full
    // strip|data rule), gray rule at every column boundary — codegen's
    // exterior frame band overrides the last one, where GT paints the black
    // frame right edge.
    for data_column_index in 0..data_column_count {
        strip_cells.push(heading_cell(
            column_letters(first_sheet_column + data_column_index as u32),
            text_style.clone(),
            CellBorder {
                top: None,
                bottom: Some(heading_separator_side()),
                left: None,
                right: Some(heading_rule_side()),
            },
        ));
    }
    page.table.rows.insert(
        0,
        TableRow {
            cant_split: false,
            minimum_height: None,
            cells: strip_cells,
            height: Some(PRINT_HEADING_STRIP_HEIGHT_PT),
        },
    );
    page.table
        .column_widths
        .insert(0, PRINT_HEADING_GUTTER_WIDTH_PT);
    page.table.prints_headings = true;

    // Drawings sit at absolute worksheet coordinates measured from cell A1's
    // top-left corner; the headings move that corner right by the gutter
    // track and down by the strip track, so the overlay offsets must follow
    // the grid.
    for image in &mut page.images {
        image.x_offset_pt += PRINT_HEADING_GUTTER_WIDTH_PT;
        image.y_offset_pt += PRINT_HEADING_STRIP_HEIGHT_PT;
    }
    for text_box in &mut page.text_boxes {
        text_box.x_offset_pt += PRINT_HEADING_GUTTER_WIDTH_PT;
        text_box.y_offset_pt += PRINT_HEADING_STRIP_HEIGHT_PT;
    }
    for sheet_shape in &mut page.shapes {
        sheet_shape.x_offset_pt += PRINT_HEADING_GUTTER_WIDTH_PT;
        sheet_shape.y_offset_pt += PRINT_HEADING_STRIP_HEIGHT_PT;
    }
    for placement in page
        .charts
        .iter_mut()
        .filter_map(|chart| chart.placement.as_mut())
    {
        placement.x_offset_pt += PRINT_HEADING_GUTTER_WIDTH_PT;
        placement.y_offset_pt += PRINT_HEADING_STRIP_HEIGHT_PT;
    }
}

#[cfg(test)]
#[path = "xlsx_print_headings_tests.rs"]
mod tests;
