use super::SHEET_TABLE_LABEL_PREFIX;
use crate::parser::Parser;
use crate::parser::xlsx::XlsxParser;
use crate::render::pdf::{PlacedGlyphRun, compiled_glyph_runs};
use crate::render::typst_gen::generate_typst;

/// Positions closer together than this are the same point.
const EPSILON_PT: f64 = 1e-6;

/// How far a glyph origin may sit from the whole sheet point Excel paces it
/// on. Tight enough that a face's own fractional advance fails it: the
/// narrowest advance in the line below is 2.87pt at 12pt, so a fractional pen
/// is at least 0.1pt off a whole point.
const GRID_TOLERANCE_PT: f64 = 0.001;

/// The occupation this issue was measured on, and the size it prints at
/// (`tests/fixtures/xlsx/customers_overflow_strip.xlsx`, row 17).
const MEASURED_LINE: &str = "Chief Configuration Representative";
const MEASURED_SIZE_PT: f64 = 12.0;

/// Workbooks that name the families a real Excel user does. Which face a host
/// resolves for those decides how many of their runs codegen's reservation
/// reaches, so they carry the invariants that hold either way rather than a
/// count.
///
/// The repository workbook prices no reservation for 856 of its 4,014 runs
/// even where every declared face is installed, so it is what holds the pass
/// to the reservation it redistributes (issue #1854).
const INSTALLED_FACE_SHEETS: [&str; 3] = [
    "customers_overflow_strip.xlsx",
    "temperature.xlsx",
    "office2pdf_repository_workbook.xlsx",
];

fn sheet_source(fixture: &str) -> String {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx")
        .join(fixture);
    let data = std::fs::read(path).expect("fixture should be available");
    let (document, _) = XlsxParser
        .parse(&data, &crate::ConvertOptions::default())
        .expect("fixture should parse");
    generate_typst(&document)
        .expect("fixture should generate Typst")
        .source
}

/// Runs of at least two glyphs, which are the only ones with an interior
/// glyph origin to place.
fn paced_runs(runs: Vec<PlacedGlyphRun>) -> Vec<PlacedGlyphRun> {
    runs.into_iter()
        .filter(|run| run.advances_pt.len() > 1)
        .collect()
}

/// The glyph origins of `run` that do not sit on a whole point of a sheet
/// printed at `scale`, as `(index, origin)`.
fn origins_off_grid(run: &PlacedGlyphRun, scale: f64) -> Vec<(usize, f64)> {
    run.glyph_origins_pt()
        .into_iter()
        .map(|origin| origin / scale)
        .enumerate()
        .filter(|(_, origin)| (origin - origin.round()).abs() > GRID_TOLERANCE_PT)
        .collect()
}

/// One worksheet cell as codegen emits it: a table carrying the sheet grid's
/// label, whose trailing field is the print scale, holding one run spaced by
/// `tracking_pt` with the shaper's own kerning and ligatures off.
fn sheet_cell_source(tracking_pt: f64, scale: f64) -> String {
    format!(
        "#set page(width: 500pt, height: 120pt, margin: 10pt)\n\
         #table(stroke: none, columns: (400pt))[\
         #text(size: {MEASURED_SIZE_PT}pt, tracking: {tracking_pt}pt, kerning: false, \
         ligatures: false)[{MEASURED_LINE}]]<{SHEET_TABLE_LABEL_PREFIX}0-{scale}>\n"
    )
}

/// The uniform correction codegen reserves for a run whose shaped advances are
/// `natural_pt`: the rounding its gaps owe Excel's grid, spread over them.
///
/// Mirrors `sheet_advance_grid_tracking_pt`, including its stop before the
/// last advance, which places no glyph of the run. Reading the advances off
/// the face that actually shaped the run is what makes these tests say the
/// same thing on every host, whatever face it resolves.
fn reserved_tracking_pt(natural_pt: &[f64], scale: f64) -> f64 {
    let gaps: &[f64] = natural_pt.split_last().expect("at least two glyphs").1;
    let quantized_pt: f64 = gaps.iter().map(|pt| (pt / scale).round() * scale).sum();
    let natural_sum_pt: f64 = gaps.iter().sum();
    (quantized_pt - natural_sum_pt) / gaps.len() as f64
}

/// The one run a [`sheet_cell_source`] paints, laid out with or without the
/// completed-frame passes.
fn compiled_cell_run(source: &str, apply_frame_passes: bool) -> PlacedGlyphRun {
    let mut runs = paced_runs(
        compiled_glyph_runs(source, 0, apply_frame_passes).expect("the cell source should compile"),
    );
    assert_eq!(runs.len(), 1, "the source paints exactly one run");
    runs.remove(0)
}

/// Excel for Mac advances every sheet glyph by a whole point, so each glyph
/// after a run's first starts a whole number of points from its origin
/// (issue #1659). Asserted over every glyph of the line this issue measured,
/// so no single correction can satisfy it.
#[test]
fn a_reserved_sheet_run_seats_every_glyph_on_the_grid() {
    let natural: PlacedGlyphRun = compiled_cell_run(&sheet_cell_source(0.0, 1.0), false);
    assert!(
        origins_off_grid(&natural, 1.0).len() > 20,
        "the face's own advances must be fractional for this to test anything: {:?}",
        natural.advances_pt
    );

    let source: String = sheet_cell_source(reserved_tracking_pt(&natural.advances_pt, 1.0), 1.0);
    let paced: PlacedGlyphRun = compiled_cell_run(&source, true);
    let off_grid: Vec<(usize, f64)> = origins_off_grid(&paced, 1.0);
    assert!(
        off_grid.is_empty(),
        "the run places glyphs off the whole-point grid: {off_grid:?} (advances {:?})",
        paced.advances_pt
    );

    let reserved: PlacedGlyphRun = compiled_cell_run(&source, false);
    assert!(
        (reserved.width_pt() - paced.width_pt()).abs() < EPSILON_PT,
        "the pass changed the run's width from {} to {}",
        reserved.width_pt(),
        paced.width_pt()
    );
    assert!(
        (reserved.left_pt - paced.left_pt).abs() < EPSILON_PT,
        "the pass moved the run's origin from {} to {}",
        reserved.left_pt,
        paced.left_pt
    );
}

/// A fitted sheet rounds at its declared size and scales that grid onto the
/// page, the same way its column widths and its tracking correction do, so its
/// glyph origins land on multiples of the print scale rather than on whole
/// printed points.
#[test]
fn a_fitted_sheet_paces_on_whole_points_of_its_own_coordinate_space() {
    const SCALE: f64 = 0.75;
    let natural: PlacedGlyphRun = compiled_cell_run(&sheet_cell_source(0.0, SCALE), false);
    let source: String =
        sheet_cell_source(reserved_tracking_pt(&natural.advances_pt, SCALE), SCALE);
    let paced: PlacedGlyphRun = compiled_cell_run(&source, true);

    let off_sheet_grid: Vec<(usize, f64)> = origins_off_grid(&paced, SCALE);
    assert!(
        off_sheet_grid.is_empty(),
        "the run places glyphs off the sheet's own grid: {off_sheet_grid:?} (advances {:?})",
        paced.advances_pt
    );
    assert!(
        !origins_off_grid(&paced, 1.0).is_empty(),
        "the grid must be the sheet's declared one, not the printed one: {:?}",
        paced.advances_pt
    );
}

/// The grid is Excel's, so a table that is not a worksheet keeps the advances
/// the layout engine placed. A DOCX or PPTX table reaches the same cell
/// emission code and must not be paced on whole points.
#[test]
fn a_table_outside_a_worksheet_keeps_its_placed_advances() {
    let natural: PlacedGlyphRun = compiled_cell_run(&sheet_cell_source(0.0, 1.0), false);
    let label: String = format!("<{SHEET_TABLE_LABEL_PREFIX}0-1>");
    let unlabelled: String =
        sheet_cell_source(reserved_tracking_pt(&natural.advances_pt, 1.0), 1.0).replace(&label, "");
    let placed: PlacedGlyphRun = compiled_cell_run(&unlabelled, true);
    assert!(
        !origins_off_grid(&placed, 1.0).is_empty(),
        "an ordinary table must keep the advances the layout engine placed: {:?}",
        placed.advances_pt
    );
}

/// A workbook reaches the same grid through codegen, and a run whose
/// reservation it never priced keeps exactly the pacing the layout engine gave
/// it. Half-pacing a run — moving its glyphs onto the grid from a width the
/// grid never priced — is what this forbids.
#[test]
fn a_sheet_run_is_either_seated_on_the_grid_or_left_as_placed() {
    for fixture in INSTALLED_FACE_SHEETS {
        let source = sheet_source(fixture);
        // A fitted sheet keeps its grid in its own declared points, so the
        // seat is read in that space (`office2pdf_repository_workbook.xlsx`
        // prints fitted).
        let scale: f64 = fitted_sheet_scale(&source).unwrap_or(1.0);
        let before =
            paced_runs(compiled_glyph_runs(&source, 0, false).expect("source should compile"));
        let after =
            paced_runs(compiled_glyph_runs(&source, 0, true).expect("source should compile"));
        assert!(
            after.len() > 4,
            "{fixture} should paint several multi-glyph runs, got {}",
            after.len()
        );
        for (before, after) in before.iter().zip(&after) {
            if before.advances_pt == after.advances_pt {
                continue;
            }
            let off_grid: Vec<(usize, f64)> = origins_off_grid(after, scale);
            assert!(
                off_grid.is_empty(),
                "{fixture}: run {:?} was re-paced but places glyphs off the whole-point \
                 grid: {off_grid:?} (advances {:?})",
                after.text,
                after.advances_pt
            );
        }
    }
}

/// The pass redistributes the rounding codegen already reserved through its
/// uniform tracking, so it must not move a run's origin or change its width.
/// Everything placed from that width — the right-aligned trailing reserve of
/// issue #1233, the centred seat of issue #1600, wrapping and the spill
/// clip — keeps the geometry it was measured with.
#[test]
fn sheet_pacing_conserves_every_run_origin_and_width() {
    for fixture in INSTALLED_FACE_SHEETS {
        let source = sheet_source(fixture);
        let before =
            compiled_glyph_runs(&source, 0, false).expect("the sheet source should compile");
        let after = compiled_glyph_runs(&source, 0, true).expect("the sheet source should compile");
        assert_eq!(
            before.len(),
            after.len(),
            "{fixture}: the pass must not add or drop runs"
        );
        for (before, after) in before.iter().zip(&after) {
            assert_eq!(before.text, after.text, "{fixture}: run order must hold");
            assert!(
                (before.left_pt - after.left_pt).abs() < EPSILON_PT,
                "{fixture}: run {:?} moved from {} to {}",
                before.text,
                before.left_pt,
                after.left_pt
            );
            assert!(
                (before.width_pt() - after.width_pt()).abs() < EPSILON_PT,
                "{fixture}: run {:?} changed width from {} to {}",
                before.text,
                before.width_pt(),
                after.width_pt()
            );
        }
    }
}

/// The print scale the sheet's own table label states.
fn fitted_sheet_scale(source: &str) -> Option<f64> {
    let label: &str = source
        .split(&format!("<{SHEET_TABLE_LABEL_PREFIX}"))
        .nth(1)?;
    label[..label.find('>')?].rsplit('-').next()?.parse().ok()
}
