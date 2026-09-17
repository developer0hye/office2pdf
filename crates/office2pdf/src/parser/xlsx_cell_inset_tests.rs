//! The horizontal box Excel lays a cell's text out in (issues #1157, #1165).
//!
//! The box's left edge steps with the cell's own font — 37 rows of a
//! one-factor native Excel for Mac probe, tabulated on `cell_left_inset_pt`,
//! which the sweep below asserts on the families whose digit advance the
//! reference table carries. The figures under it fix the Calibri 11 pair the
//! step reduces to at the workbook default.
//!
//! Measured against the ten native Excel for Mac exports under
//! `tests/golden_mocks/business/expected/xlsx/`, one `fill_text` per cell in a
//! `mutool draw -F trace` of both sides, matched by string and baseline. Each
//! run is classified by how far it moves when the right inset alone changes:
//! a left-aligned run does not move, a centred one moves half a point, a
//! right-aligned one a whole point.
//!
//! | run class | n | comparable edge | error at 3/3 | error at 3/2 |
//! | --- | ---: | --- | ---: | ---: |
//! | left-aligned | 139 | pen origin | median 0.000 | median 0.000 |
//! | centred | 570 | run centre | mean +0.012 | mean +0.512 |
//! | right-aligned | 135 | pen end | mean -0.994 | mean +0.006 |
//!
//! The right-aligned figures are the ones that carry a correction, because
//! Excel rounds every glyph advance to a whole point: its last glyph is laid
//! down `round(adv) - adv` short of where the true advance would end it, and
//! we place the same glyph on the unrounded advance. Subtracting that per-run
//! offset — read off the GT's own trace, not fitted — leaves +0.006pt (sd
//! 0.086) at a 2pt right inset against -0.994pt at 3pt: one whole point out.
//!
//! The centred figures carry no inset correction. Excel centres the run on
//! the *column*, not in that asymmetric box, so a symmetric split of the same
//! 5pt total is what puts our centred runs on its own — which is why issue
//! #657's warning that an asymmetric pair moves every centred run by half the
//! difference is right, and only its conclusion that both sides are therefore
//! 3pt is not. The remaining sub-point residual of a centred run is Excel's
//! whole-point origin seat, which the renderer applies from the cell's
//! `wrapText` flag (issue #1600, `wraps_text` below).

use super::*;

/// A one-cell workbook whose cell carries `horizontal`.
fn workbook_with_cell_alignment(
    horizontal: umya_spreadsheet::HorizontalAlignmentValues,
) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("B3");
        cell.set_value("2026");
        cell.get_style_mut()
            .get_alignment_mut()
            .set_horizontal(horizontal);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// The insets the workbook's one value-bearing cell is laid out with: its own
/// when it states them, otherwise the table default it inherits.
fn first_cell_padding(data: &[u8]) -> Insets {
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("workbook should parse");
    let sheet = get_sheet_page(&doc, 0);
    let default: Insets = sheet
        .table
        .default_cell_padding
        .expect("a sheet states the cell padding its table is laid out with");
    sheet
        .table
        .rows
        .iter()
        .flat_map(|row| row.cells.iter())
        .find(|cell| !cell.content.is_empty())
        .expect("the workbook has a value-bearing cell")
        .padding
        .unwrap_or(default)
}

#[test]
fn test_left_aligned_cell_starts_three_points_inside_its_column() {
    let data = workbook_with_cell_alignment(umya_spreadsheet::HorizontalAlignmentValues::Left);

    let padding: Insets = first_cell_padding(&data);

    assert!(
        (padding.left - 3.0).abs() < 0.01,
        "Excel starts a left-aligned value 3pt inside the column, got {}",
        padding.left
    );
}

#[test]
fn test_right_aligned_cell_ends_two_points_inside_its_column() {
    let data = workbook_with_cell_alignment(umya_spreadsheet::HorizontalAlignmentValues::Right);

    let padding: Insets = first_cell_padding(&data);

    assert!(
        (padding.right - 2.0).abs() < 0.01,
        "Excel ends a right-aligned value 2pt inside the column, got {}",
        padding.right
    );
}

#[test]
fn test_centred_cell_stays_on_its_columns_own_centre() {
    let data = workbook_with_cell_alignment(umya_spreadsheet::HorizontalAlignmentValues::Center);

    let padding: Insets = first_cell_padding(&data);

    assert!(
        (padding.left - padding.right).abs() < 0.01,
        "an asymmetric box moves a centred run off the column's centre by half \
         the difference, got left {} right {}",
        padding.left,
        padding.right
    );
    assert!(
        ((padding.left + padding.right) - 5.0).abs() < 0.01,
        "a centred cell splits Excel's own 5pt total, got {}",
        padding.left + padding.right
    );
}

#[test]
fn test_sheet_table_carries_excels_cell_text_box() {
    let data = build_xlsx_bytes("Sheet1", &[("B3", "2026")]);
    let (doc, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("workbook should parse");

    let source: String = crate::render::typst_gen::generate_typst(&doc)
        .expect("sheet should generate Typst")
        .source;

    assert!(
        source.contains("inset: (top: 1pt, right: 2pt, bottom: 1.5pt, left: 3pt)"),
        "the measured text box should reach the renderer, got:\n{source}"
    );
}

#[test]
fn test_sheet_default_padding_follows_the_workbook_normal_fonts_column_unit() {
    let data = build_xlsx_with_normal_font("Arial", 32.0);
    let (doc, _warnings) = XlsxParser
        .parse(&data, &ConvertOptions::default())
        .expect("workbook should parse");
    let sheet = get_sheet_page(&doc, 0);
    let padding = sheet
        .table
        .default_cell_padding
        .expect("the table carries its Normal-font cell box");

    assert_eq!(padding.left, 6.0);
    assert_eq!(padding.right, 5.0);
    assert!(
        sheet.table.rows[0].cells[0].padding.is_none(),
        "an unstyled cell inherits the Normal-font table box instead of repeating it"
    );
}

/// A one-cell workbook whose cell states `family`, `size_pt`, and `alignment`.
fn workbook_with_cell_font_and_alignment(
    family: &str,
    size_pt: f64,
    alignment: umya_spreadsheet::HorizontalAlignmentValues,
) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("B3");
        cell.set_value("2026");
        let style = cell.get_style_mut();
        style.get_alignment_mut().set_horizontal(alignment);
        let font = style.get_font_mut();
        font.set_name(family);
        font.set_size(size_pt);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// A one-cell workbook whose cell states `family` at `size_pt`, left-aligned.
fn workbook_with_cell_font(family: &str, size_pt: f64) -> Vec<u8> {
    workbook_with_cell_font_and_alignment(
        family,
        size_pt,
        umya_spreadsheet::HorizontalAlignmentValues::Left,
    )
}

/// A one-cell, left-aligned workbook whose cell states `family`, `size_pt`,
/// and `bold`.
fn workbook_with_cell_font_and_weight(family: &str, size_pt: f64, bold: bool) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("B3");
        cell.set_value("2026");
        let style = cell.get_style_mut();
        style
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Left);
        let font = style.get_font_mut();
        font.set_name(family);
        font.set_size(size_pt);
        font.set_bold(bold);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// Every row of the issue #1165 probe whose family the reference digit table
/// carries, so the expectation is the same on any machine. Century Gothic and
/// Segoe UI are measured in the module doc but resolve through the live font
/// set, which CI does not ship.
const MEASURED_LEFT_INSETS: &[(&str, f64, f64)] = &[
    ("Calibri", 6.0, 2.0),
    ("Calibri", 7.0, 2.0),
    ("Calibri", 8.0, 2.0),
    ("Calibri", 9.0, 3.0),
    ("Calibri", 10.0, 3.0),
    ("Calibri", 11.0, 3.0),
    ("Calibri", 12.0, 3.0),
    ("Calibri", 14.0, 3.0),
    ("Calibri", 16.0, 3.0),
    ("Calibri", 17.0, 4.0),
    ("Calibri", 20.0, 4.0),
    ("Calibri", 24.0, 4.0),
    ("Calibri", 25.0, 5.0),
    ("Calibri", 28.0, 5.0),
    ("Calibri", 32.0, 5.0),
    ("Calibri", 33.0, 6.0),
    ("Calibri", 36.0, 6.0),
    ("Arial", 8.0, 2.0),
    ("Arial", 10.0, 3.0),
    ("Arial", 12.0, 3.0),
    ("Arial", 14.0, 3.0),
    ("Arial", 16.0, 4.0),
    ("Arial", 18.0, 4.0),
    ("Arial", 20.0, 4.0),
    ("Arial", 24.0, 5.0),
    ("Arial", 32.0, 6.0),
    ("Times New Roman", 11.0, 3.0),
    ("Times New Roman", 16.0, 3.0),
    ("Times New Roman", 18.0, 4.0),
    ("Verdana", 10.0, 3.0),
    ("Verdana", 11.0, 3.0),
];

#[test]
fn test_left_inset_steps_with_the_cell_fonts_own_column_unit() {
    for &(family, size_pt, expected_left) in MEASURED_LEFT_INSETS {
        let padding: Insets = first_cell_padding(&workbook_with_cell_font(family, size_pt));

        assert!(
            (padding.left - expected_left).abs() < 0.01,
            "{family} {size_pt} starts {expected_left}pt inside its column in Excel's own \
             export, got {}",
            padding.left,
        );
    }
}

#[test]
fn test_right_inset_steps_one_point_behind_the_cell_fonts_left_inset() {
    for &(family, size_pt, expected_left) in MEASURED_LEFT_INSETS {
        let data = workbook_with_cell_font_and_alignment(
            family,
            size_pt,
            umya_spreadsheet::HorizontalAlignmentValues::Right,
        );
        let padding: Insets = first_cell_padding(&data);
        let expected_right: f64 = expected_left - 1.0;

        assert!(
            (padding.right - expected_right).abs() < 0.01,
            "{family} {size_pt} ends {expected_right}pt inside its column in Excel's own \
             export, got {}",
            padding.right,
        );
    }
}

#[test]
fn test_a_title_cell_starts_further_in_than_the_body_line_below_it() {
    // The shape the issue reported: one column, a body line and a title, the
    // title's origin further right than the body's. Both sizes are probe rows,
    // and their 3pt step is the one the reported workbook's column B carries.
    let body: Insets = first_cell_padding(&workbook_with_cell_font("Arial", 14.0));
    let title: Insets = first_cell_padding(&workbook_with_cell_font("Arial", 32.0));

    assert!(
        (title.left - body.left - 3.0).abs() < 0.01,
        "Excel starts a 32pt title three points right of a 14pt body line in the same \
         column, got {} against {}",
        title.left,
        body.left,
    );
}

#[test]
fn test_the_cell_font_drives_the_inset_not_the_workbook_normal_font() {
    // The workbook Normal font stays Calibri 11, whose inset is 3pt; only the
    // cell's own font is larger.
    let padding: Insets = first_cell_padding(&workbook_with_cell_font("Arial", 32.0));

    assert!(
        (padding.left - 6.0).abs() < 0.01,
        "the cell's own Arial 32 takes a 6pt inset whatever the Normal font is, got {}",
        padding.left,
    );
}

#[test]
fn test_a_centred_cell_holds_the_column_centre_at_any_cell_font() {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("B3");
        cell.set_value("2026");
        let style = cell.get_style_mut();
        style
            .get_alignment_mut()
            .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Center);
        // Name the face outright: umya's default cell font defers to the
        // theme's minor scheme, which resolves to the theme's UI-script face
        // and a different column unit (issue #1380), while this measurement
        // is Calibri 32's.
        style.get_font_mut().set_name("Calibri");
        style.get_font_mut().set_size(32.0);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();

    let padding: Insets = first_cell_padding(&cursor.into_inner());

    assert!(
        (padding.left - padding.right).abs() < 0.01,
        "a centred run sits on the column's own centre whatever its font size, got \
         left {} right {}",
        padding.left,
        padding.right,
    );
    assert!(
        ((padding.left + padding.right) - 9.0).abs() < 0.01,
        "a centred Calibri 32 cell keeps the font-sized 5pt/4pt box total, got left {} \
         right {}",
        padding.left,
        padding.right,
    );
}

/// A one-cell workbook whose centred cell states `wrapText`.
fn workbook_with_centred_cell_wrap(wrap_text: bool) -> Vec<u8> {
    let mut book = umya_spreadsheet::new_file();
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        let cell = sheet.get_cell_mut("B3");
        cell.set_value("Yes");
        let alignment = cell.get_style_mut().get_alignment_mut();
        alignment.set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Center);
        alignment.set_wrap_text(wrap_text);
    }
    let mut cursor = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&book, &mut cursor).unwrap();
    cursor.into_inner()
}

/// The workbook's one value-bearing cell.
fn first_value_cell(data: &[u8]) -> TableCell {
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("workbook should parse");
    get_sheet_page(&doc, 0)
        .table
        .rows
        .iter()
        .flat_map(|row| row.cells.iter())
        .find(|cell| !cell.content.is_empty())
        .expect("the workbook has a value-bearing cell")
        .clone()
}

/// Excel starts a centred line one sheet point further left in a wrapped
/// cell than in an unwrapped one of the same width and text (issue #1600), so
/// the flag has to reach the renderer even when the text fits on one line and
/// leaves no spill behind.
#[test]
fn test_centred_cell_carries_its_wrap_text_flag_to_the_renderer() {
    let wrapped: TableCell = first_value_cell(&workbook_with_centred_cell_wrap(true));
    let unwrapped: TableCell = first_value_cell(&workbook_with_centred_cell_wrap(false));

    assert!(
        wrapped.wraps_text,
        "a wrapText=\"1\" cell must carry the flag: {wrapped:?}"
    );
    assert!(
        !unwrapped.wraps_text,
        "an unwrapped cell must not carry the flag: {unwrapped:?}"
    );
    assert!(
        wrapped.spill_width.is_none() && unwrapped.spill_width.is_none(),
        "text that fits its column leaves no spill on either cell, so the spill cannot stand in for the flag"
    );
}

/// Cambria 11/24/36/42/48, regular and bold, bracketed by a Calibri 11
/// control (issue #1623). This is the already-validated `ceil(unit / 4) + 1`
/// inset formula (#1165, #1232) evaluated at Cambria's regular and bold
/// digit advances, which are read directly from Excel's own font files —
/// `Cambria.ttc`'s 1134/2048em and `Cambriab.ttf`'s 1213/2048em hmtx maxima
/// over U+0030..=U+0039 (verified with `fontTools`, not estimated).
///
/// Three of these rows also have an independent native Excel for Mac
/// confirmation, not just the formula:
/// - 42pt bold is the issue's own fixture title cell: a fresh native export
///   places it at physical x 85.80pt, matching this model's 8pt inset to
///   within the +0.475pt residual issue #1719 already tracks for every run on
///   that fitted sheet (a shared paint/text origin snap, not this rule).
/// - 36 and 48 are a differential probe (2026-09-17): the same wide-column
///   workbook exported regular and bold, reading only the bold-minus-regular
///   shift in the "2026" digits' start x so the (unmodelled) column boundary
///   cancels out. Native gave +1.0pt at 36 and +0.0pt at 48, matching this
///   table exactly.
///
/// 11 and 24 are not independently probed: at those sizes the rounded
/// whole-point column unit lands in the same `ceil(unit / 4)` bracket
/// regardless of weight, so they are the formula's own negative control
/// rather than a new empirical claim.
const MEASURED_CAMBRIA_LEFT_INSETS: &[(&str, f64, bool, f64)] = &[
    ("Calibri", 11.0, false, 3.0),
    ("Cambria", 11.0, false, 3.0),
    ("Cambria", 11.0, true, 3.0),
    ("Cambria", 24.0, false, 5.0),
    ("Cambria", 24.0, true, 5.0),
    ("Cambria", 36.0, false, 6.0),
    ("Cambria", 36.0, true, 7.0),
    ("Cambria", 42.0, false, 7.0),
    ("Cambria", 42.0, true, 8.0),
    ("Cambria", 48.0, false, 8.0),
    ("Cambria", 48.0, true, 8.0),
    ("Calibri", 11.0, false, 3.0),
];

#[test]
fn test_bold_cell_font_prices_its_own_bold_digit_advance() {
    for &(family, size_pt, bold, expected_left) in MEASURED_CAMBRIA_LEFT_INSETS {
        let padding: Insets =
            first_cell_padding(&workbook_with_cell_font_and_weight(family, size_pt, bold));

        assert!(
            (padding.left - expected_left).abs() < 0.01,
            "{family} {size_pt} bold={bold} starts {expected_left}pt inside its column in \
             Excel's own export, got {}",
            padding.left,
        );
    }
}

/// A family whose digits are the same width at every weight (Calibri, Arial,
/// Times New Roman all measured equal) must not move its inset just because
/// the cell is bold — ruling out a hack that always adds a point for bold
/// rather than actually pricing the bold face's own digit advance.
#[test]
fn test_bold_does_not_move_the_inset_for_a_tabular_digit_family() {
    let regular: Insets =
        first_cell_padding(&workbook_with_cell_font_and_weight("Calibri", 32.0, false));
    let bold: Insets =
        first_cell_padding(&workbook_with_cell_font_and_weight("Calibri", 32.0, true));

    assert!(
        (regular.left - bold.left).abs() < 0.01,
        "Calibri's digits are the same width bold or regular, so the inset must not move: \
         regular {} bold {}",
        regular.left,
        bold.left,
    );
}

/// Malgun Gothic's bold digits (1187/2048em, read from Excel's own
/// `malgunbd.ttf`) are ~5% wider than regular (1128/2048em) — the same class
/// of defect as Cambria and Verdana (issue #1623), on a family this codebase
/// already knew was not weight-invariant on other metrics (issues #1097,
/// #1199, #1208, #1627). 11pt is the workbook-default size and lands both
/// weights in the same `ceil(unit / 4)` bracket (a negative control); 22pt is
/// where the wider bold digit measurably moves the inset. Formula-derived
/// from the measured em values, not an independent native probe.
#[test]
fn test_bold_cell_font_prices_its_own_bold_malgun_gothic_digit_advance() {
    for &(size_pt, bold, expected_left) in &[
        (11.0_f64, false, 3.0_f64),
        (11.0, true, 3.0),
        (22.0, false, 4.0),
        (22.0, true, 5.0),
    ] {
        let padding: Insets = first_cell_padding(&workbook_with_cell_font_and_weight(
            "Malgun Gothic",
            size_pt,
            bold,
        ));

        assert!(
            (padding.left - expected_left).abs() < 0.01,
            "Malgun Gothic {size_pt} bold={bold} should start {expected_left}pt inside its \
             column, got {}",
            padding.left,
        );
    }
}
