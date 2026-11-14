use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::cond_fmt_raw::{
    parse_relationships, parse_sheet_relationships, read_zip_text, worksheet_path,
};

/// ECMA-376's default for `<pageSetup fitToWidth>` and `<pageSetup
/// fitToHeight>`, used when either attribute is absent (§18.3.1.63).
///
/// Excel omits `fitToWidth` from a sheet that fits onto one page wide, which
/// is the most common shape of all (issue #850). A sheet that omits
/// `fitToHeight` is asking for the same bound in the row direction, and Excel
/// honours it: the reported college-budget workbook names neither attribute
/// and its native export is one A3 page (issue #1181).
const DEFAULT_FIT_TO_PAGES: u32 = 1;

/// Percentages `<pageSetup scale>` is allowed to name (ECMA-376 §18.3.1.63).
///
/// Excel's Page Setup dialog clamps the box to this range, and a file naming
/// anything outside it — including the `0` umya reports for an absent
/// attribute — is asking for no scaling rather than for a degenerate one.
pub(super) const PRINT_PERCENTAGE_RANGE: std::ops::RangeInclusive<u32> = 10..=400;

/// The percentage that prints a sheet at its declared size.
const UNSCALED_PRINT_PERCENTAGE: u32 = 100;

/// Worksheet print scaling mode and header/footer scaling policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SheetPrintSetup {
    /// Whether fit-to-page overrides an explicit print percentage.
    pub(crate) fits_to_page: bool,
    /// Active `fitToWidth`, defaulting to [`DEFAULT_FIT_TO_PAGES`] when omitted.
    /// Zero leaves width unconstrained, including when fit-to-page is disabled.
    pub(crate) pages_wide: u32,
    /// Active `fitToHeight`, defaulting to [`DEFAULT_FIT_TO_PAGES`] when omitted.
    /// Zero leaves height unconstrained, including when fit-to-page is disabled.
    pub(crate) pages_tall: u32,
    /// `<pageSetup scale>` when it asks for a size other than 100%, which is
    /// what Excel's "Adjust to: N% normal size" writes. `None` covers an
    /// absent attribute, an out-of-range one, and a declared 100 — none of
    /// which changes a printed dimension.
    pub(crate) print_percentage: Option<u32>,
    /// `headerFooter/@scaleWithDoc`, which defaults to `1` — Excel shrinks the
    /// header and footer with the sheet unless the file opts out
    /// (ECMA-376 §18.3.1.46, issue #940).
    pub(crate) header_footer_scales_with_doc: bool,
}

/// What each sheet's page setup asks for, keyed by sheet name.
///
/// Every worksheet appears so explicit percentage scaling also preserves
/// `headerFooter/@scaleWithDoc`. Fit bounds are active only when
/// `<sheetPr><pageSetUpPr fitToPage="1"/>` selects fit-to-page (issue #530).
///
/// umya-spreadsheet models `<pageSetup>` but not `<sheetPr>`, and it cannot
/// tell an absent `fitToWidth`, `fitToHeight` or `scale` from an explicit zero
/// — all read back as 0 — so all four are read from the archive directly.
/// `<headerFooter>`'s
/// `scaleWithDoc` is read in the same pass for the same reason: the struct
/// umya exposes carries only the section strings.
pub(crate) fn sheets_print_setup(data: &[u8]) -> HashMap<String, SheetPrintSetup> {
    let mut setups: HashMap<String, SheetPrintSetup> = HashMap::new();
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return setups;
    };
    let Some(workbook_xml) = read_zip_text(&mut archive, "xl/workbook.xml") else {
        return setups;
    };
    let Some(relationships_xml) = read_zip_text(&mut archive, "xl/_rels/workbook.xml.rels") else {
        return setups;
    };

    let relationships = parse_relationships(&relationships_xml);
    for (sheet_name, relationship_id) in parse_sheet_relationships(&workbook_xml) {
        let Some(target) = relationships.get(&relationship_id) else {
            continue;
        };
        let Some(worksheet_xml) = read_zip_text(&mut archive, &worksheet_path(target)) else {
            continue;
        };
        let fit = worksheet_fit_to_page(&worksheet_xml);
        let (pages_wide, pages_tall) = fit.unwrap_or((0, 0));
        setups.insert(
            sheet_name,
            SheetPrintSetup {
                fits_to_page: fit.is_some(),
                pages_wide,
                pages_tall,
                print_percentage: worksheet_print_percentage(&worksheet_xml),
                header_footer_scales_with_doc: worksheet_header_footer_scales_with_doc(
                    &worksheet_xml,
                ),
            },
        );
    }
    setups
}

/// `(fitToWidth, fitToHeight)` for one worksheet part, or `None` when it does
/// not fit to page.
///
/// `<pageSetUpPr>` precedes `<sheetData>` and `<pageSetup>` follows it, so the
/// whole part is scanned rather than stopping at the cells.
fn worksheet_fit_to_page(worksheet_xml: &str) -> Option<(u32, u32)> {
    let mut reader = Reader::from_str(worksheet_xml);
    let mut fits_to_page: bool = false;
    let mut declared_pages_wide: Option<u32> = None;
    let mut declared_pages_tall: Option<u32> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"pageSetUpPr" => {
                        fits_to_page = element.attributes().flatten().any(|attribute| {
                            attribute.key.local_name().as_ref() == b"fitToPage"
                                && matches!(attribute.value.as_ref(), b"1" | b"true")
                        });
                    }
                    b"pageSetup" => {
                        declared_pages_wide = page_count_attribute(element, b"fitToWidth");
                        declared_pages_tall = page_count_attribute(element, b"fitToHeight");
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    fits_to_page.then(|| {
        (
            declared_pages_wide.unwrap_or(DEFAULT_FIT_TO_PAGES),
            declared_pages_tall.unwrap_or(DEFAULT_FIT_TO_PAGES),
        )
    })
}

/// The sheet's explicit print percentage, or `None` when it prints unscaled.
///
/// The `<customSheetViews>` subtree is skipped for the same reason
/// `worksheet_header_footer_scales_with_doc` skips it: each CT_CustomSheetView
/// nests its own `<pageSetup>`, so a saved view's percentage would otherwise
/// be read as the sheet's and rescale the whole printed grid.
fn worksheet_print_percentage(worksheet_xml: &str) -> Option<u32> {
    let mut reader = Reader::from_str(worksheet_xml);
    // Element depth inside the skipped `<customSheetViews>` subtree; 0 means
    // the scan is at sheet level.
    let mut skipped_subtree_depth: usize = 0;
    let mut declared: Option<u32> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element)) => {
                if skipped_subtree_depth > 0 {
                    skipped_subtree_depth += 1;
                } else if element.local_name().as_ref() == b"customSheetViews" {
                    skipped_subtree_depth = 1;
                } else if element.local_name().as_ref() == b"pageSetup" {
                    declared = page_count_attribute(element, b"scale");
                }
            }
            Ok(Event::Empty(ref element)) => {
                if skipped_subtree_depth == 0 && element.local_name().as_ref() == b"pageSetup" {
                    declared = page_count_attribute(element, b"scale");
                }
            }
            Ok(Event::End(_)) => {
                skipped_subtree_depth = skipped_subtree_depth.saturating_sub(1);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    declared.filter(|percentage| {
        PRINT_PERCENTAGE_RANGE.contains(percentage) && *percentage != UNSCALED_PRINT_PERCENTAGE
    })
}

/// One `<pageSetup>` page-count attribute, or `None` when it is absent or not
/// a number. An absent attribute has to stay distinguishable from a declared
/// zero: zero leaves that direction unconstrained where absence asks for one
/// page.
fn page_count_attribute(element: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<u32> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            std::str::from_utf8(attribute.value.as_ref())
                .ok()
                .and_then(|value| value.trim().parse::<u32>().ok())
        })
}

/// Whether the worksheet's `<headerFooter>` scales with the sheet.
///
/// `scaleWithDoc` defaults to `1`, so a part with no `<headerFooter>` at all —
/// or one that omits the attribute — scales. Only an explicit false opts out.
/// The `<customSheetViews>` subtree is skipped for the same reason
/// `worksheet_print_options` skips it: each CT_CustomSheetView nests its own
/// `<headerFooter>`, describing that saved view rather than the sheet.
fn worksheet_header_footer_scales_with_doc(worksheet_xml: &str) -> bool {
    let mut reader = Reader::from_str(worksheet_xml);
    // Element depth inside the skipped `<customSheetViews>` subtree; 0 means
    // the scan is at sheet level.
    let mut skipped_subtree_depth: usize = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element)) => {
                if skipped_subtree_depth > 0 {
                    skipped_subtree_depth += 1;
                } else if element.local_name().as_ref() == b"customSheetViews" {
                    skipped_subtree_depth = 1;
                } else if element.local_name().as_ref() == b"headerFooter" {
                    return scales_with_doc(element);
                }
            }
            Ok(Event::Empty(ref element)) => {
                if skipped_subtree_depth == 0 && element.local_name().as_ref() == b"headerFooter" {
                    return scales_with_doc(element);
                }
            }
            Ok(Event::End(_)) => {
                skipped_subtree_depth = skipped_subtree_depth.saturating_sub(1);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    true
}

/// `scaleWithDoc` on a `<headerFooter>` element: true unless explicitly false.
fn scales_with_doc(element: &quick_xml::events::BytesStart<'_>) -> bool {
    !element.attributes().flatten().any(|attribute| {
        attribute.key.local_name().as_ref() == b"scaleWithDoc"
            && matches!(attribute.value.as_ref(), b"0" | b"false")
    })
}

#[cfg(test)]
#[path = "xlsx_fit_to_page_tests.rs"]
mod tests;
