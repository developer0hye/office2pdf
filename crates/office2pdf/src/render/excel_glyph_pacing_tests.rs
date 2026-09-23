use super::SHEET_TABLE_LABEL_PREFIX;
use crate::parser::Parser;
use crate::parser::xlsx::XlsxParser;
use crate::render::pdf::{PlacedGlyphRun, compiled_glyph_runs};
use crate::render::typst_gen::generate_typst;

/// Positions closer together than this are the same point.
const EPSILON_PT: f64 = 1e-6;

/// How far a glyph origin may sit from the whole sheet point Excel paces it
/// on. Tight enough that a face's own fractional advance fails it: the
/// smallest advance on the customer fixture's 12pt lines is 2.87pt, so a
/// fractional pen is at least 0.1pt off a whole point.
const GRID_TOLERANCE_PT: f64 = 0.001;

/// Workbooks whose cells name a family that ships inside this crate, so the
/// face that shapes them is the same one on every host.
///
/// That matters because the grid correction is priced by codegen's own metric
/// lookup: where a host has to substitute a face for the declared family, the
/// run is placed from a width this grid never priced and the pass declines it
/// (see `RESERVED_WIDTH_TOLERANCE_PT`). Only a workbook whose face travels
/// with the crate can carry the assertion over *every* run.
const EMBEDDED_FACE_SHEET: &str = "sheet_pacing_embedded_face.xlsx";
const EMBEDDED_FACE_FITTED_SHEET: &str = "sheet_pacing_embedded_face_fitted.xlsx";

/// Workbooks that name the families a real Excel user does. Whether a host
/// has those installed decides how many of their runs the grid reaches, so
/// they carry the assertions that hold either way.
///
/// The repository workbook is the one that priced no correction for 856 of
/// its 4,014 runs even where every declared face is installed, so it is what
/// holds the pass to the reservation it redistributes.
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

/// The glyph origins of `run` that do not sit on a whole point of the sheet's
/// own coordinate space, as `(index, origin)`.
fn origins_off_grid(run: &PlacedGlyphRun, scale: f64) -> Vec<(usize, f64)> {
    run.glyph_origins_pt()
        .into_iter()
        .map(|origin| origin / scale)
        .enumerate()
        .filter(|(_, origin)| (origin - origin.round()).abs() > GRID_TOLERANCE_PT)
        .collect()
}

/// Excel for Mac advances every sheet glyph by a whole point, so each glyph
/// after a run's first starts a whole number of points from its origin
/// (issue #1659). Asserted over every multi-glyph run of the sheet rather
/// than one measured string, so no per-string correction can satisfy it.
#[test]
fn sheet_glyph_origins_land_on_whole_points() {
    let source = sheet_source(EMBEDDED_FACE_SHEET);
    let runs =
        paced_runs(compiled_glyph_runs(&source, 0, true).expect("the sheet source should compile"));
    assert!(
        runs.len() > 4,
        "{EMBEDDED_FACE_SHEET} should paint several multi-glyph runs, got {}",
        runs.len()
    );
    for run in &runs {
        let off_grid: Vec<(usize, f64)> = origins_off_grid(run, 1.0);
        assert!(
            off_grid.is_empty(),
            "{EMBEDDED_FACE_SHEET}: run {:?} places glyphs off the whole-point grid: \
             {off_grid:?} (advances {:?})",
            run.text,
            run.advances_pt
        );
    }
}

/// A workbook naming an installed family reaches the grid the same way, and
/// where a host substitutes a face for it the run keeps exactly the pacing
/// the layout engine gave it. Half-pacing a run — moving its glyphs onto the
/// grid from a width the grid never priced — is what this forbids.
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
    for fixture in [EMBEDDED_FACE_SHEET, EMBEDDED_FACE_FITTED_SHEET]
        .into_iter()
        .chain(INSTALLED_FACE_SHEETS)
    {
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

/// The grid is Excel's, so a table that is not a worksheet keeps the face's
/// own fractional advances. A DOCX or PPTX table reaches the same cell
/// emission code and must not be paced on whole points.
#[test]
fn a_table_outside_a_worksheet_keeps_its_fractional_advances() {
    let source = "#set page(width: 300pt, height: 120pt, margin: 10pt)\n\
                  #table(columns: (200pt), stroke: none)[\
                  #text(font: \"Liberation Sans\", size: 12pt)[Configuration]]\n";
    let runs = paced_runs(compiled_glyph_runs(source, 0, true).expect("source should compile"));
    assert_eq!(runs.len(), 1, "the source paints exactly one run");
    let fractional: bool = runs[0]
        .glyph_origins_pt()
        .iter()
        .any(|origin| (origin - origin.round()).abs() > GRID_TOLERANCE_PT);
    assert!(
        fractional,
        "an ordinary table must keep the face's advances: {:?}",
        runs[0].advances_pt
    );
}

/// A fitted sheet rounds at its declared size and scales that grid onto the
/// page, the same way its column widths and its tracking correction do, so
/// its glyph origins land on multiples of the print scale rather than on
/// whole printed points.
#[test]
fn a_fitted_sheet_paces_on_whole_points_of_its_own_coordinate_space() {
    let source = sheet_source(EMBEDDED_FACE_FITTED_SHEET);
    let scale: f64 = fitted_sheet_scale(&source).expect("the fixture should print fitted");
    assert!(
        scale < 1.0,
        "the fixture must actually scale for this to test anything, got {scale}"
    );
    let runs =
        paced_runs(compiled_glyph_runs(&source, 0, true).expect("the sheet source should compile"));
    assert!(runs.len() > 4, "the sheet should paint several runs");
    let mut off_printed_grid: usize = 0;
    for run in &runs {
        let off_sheet_grid: Vec<(usize, f64)> = origins_off_grid(run, scale);
        assert!(
            off_sheet_grid.is_empty(),
            "run {:?} places a glyph off the sheet's own grid: {off_sheet_grid:?} \
             (advances {:?})",
            run.text,
            run.advances_pt
        );
        off_printed_grid += run
            .glyph_origins_pt()
            .into_iter()
            .filter(|origin| (origin - origin.round()).abs() > GRID_TOLERANCE_PT)
            .count();
    }
    assert!(
        off_printed_grid > 0,
        "the sheet grid must be the declared one, not the printed one"
    );
}

/// The print scale the sheet's own table label states.
fn fitted_sheet_scale(source: &str) -> Option<f64> {
    let label: &str = source
        .split(&format!("<{SHEET_TABLE_LABEL_PREFIX}"))
        .nth(1)?;
    label[..label.find('>')?].rsplit('-').next()?.parse().ok()
}
