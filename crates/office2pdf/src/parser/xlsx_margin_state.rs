use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::cond_fmt_raw::worksheets_map;

/// Which print margin attributes a worksheet's `<pageMargins>` states,
/// whatever value each one states.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) struct DeclaredPrintMargins {
    pub(crate) left: bool,
    pub(crate) right: bool,
    pub(crate) top: bool,
    pub(crate) bottom: bool,
    pub(crate) header: bool,
    pub(crate) footer: bool,
}

/// Each worksheet name against the print margins its own part declares.
///
/// Excel writes `left="0"` for a sheet printed flush against the paper, and
/// crates.io umya-spreadsheet maps both that zero and a missing attribute to
/// the same `0.0`. Reading the value therefore cannot tell a borderless sheet
/// from a silent one, which substituted the 0.7in default and moved the first
/// glyph from x22 to x53 (issue #1812). Take the provenance from the package
/// part, the way [`super::paper_state`] takes the paper state, rather than
/// from a fork-only accessor: the published library resolves against the
/// crates.io graph, where a workspace patch does not apply (issue #1041).
///
/// A sheet is absent from the map only when its part could not be read or
/// parsed at all. A well-formed part with no `<pageMargins>` maps to "nothing
/// declared", so the two cases stay distinguishable and an unreadable package
/// can fall back to the value test instead of collapsing every margin to zero.
pub(crate) fn declared_print_margins(data: &[u8]) -> HashMap<String, DeclaredPrintMargins> {
    worksheets_map(data, worksheet_declared_print_margins)
}

/// The print margins one worksheet declares, or `None` when its part is
/// malformed.
///
/// Only a direct child of `<worksheet>` counts. A saved custom view carries
/// its own `<pageMargins>`, and that view-local state must not decide how the
/// sheet prints — the same rule [`super::paper_state`] applies to the paper.
fn worksheet_declared_print_margins(worksheet_xml: &str) -> Option<DeclaredPrintMargins> {
    let mut reader = Reader::from_str(worksheet_xml);
    let mut depth: usize = 0;
    let mut saw_worksheet: bool = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref element)) => {
                if depth == 0 && element.local_name().as_ref() == b"worksheet" {
                    saw_worksheet = true;
                } else if depth == 1 && element.local_name().as_ref() == b"pageMargins" {
                    return Some(declared_edges(element));
                }
                depth += 1;
            }
            Ok(Event::Empty(ref element)) => {
                if depth == 0 && element.local_name().as_ref() == b"worksheet" {
                    return Some(DeclaredPrintMargins::default());
                }
                if depth == 1 && element.local_name().as_ref() == b"pageMargins" {
                    return Some(declared_edges(element));
                }
            }
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) => return saw_worksheet.then(DeclaredPrintMargins::default),
            Err(_) => return None,
            _ => {}
        }
    }
}

/// The margin attributes one `<pageMargins>` element names.
///
/// Attribute keys are matched on their local name so a worksheet serialized
/// with a namespace prefix — which the reader itself accepts (issue #1803) —
/// does not silently read as declaring nothing.
fn declared_edges(element: &BytesStart<'_>) -> DeclaredPrintMargins {
    let mut declared = DeclaredPrintMargins::default();
    for attribute in element.attributes().flatten() {
        match attribute.key.local_name().as_ref() {
            b"left" => declared.left = true,
            b"right" => declared.right = true,
            b"top" => declared.top = true,
            b"bottom" => declared.bottom = true,
            b"header" => declared.header = true,
            b"footer" => declared.footer = true,
            _ => {}
        }
    }
    declared
}

#[cfg(test)]
mod tests {
    use super::{DeclaredPrintMargins, worksheet_declared_print_margins};

    const ALL: DeclaredPrintMargins = DeclaredPrintMargins {
        left: true,
        right: true,
        top: true,
        bottom: true,
        header: true,
        footer: true,
    };
    const NONE: DeclaredPrintMargins = DeclaredPrintMargins {
        left: false,
        right: false,
        top: false,
        bottom: false,
        header: false,
        footer: false,
    };

    #[test]
    fn a_declared_zero_is_declared_and_an_absent_element_declares_nothing() {
        assert_eq!(
            worksheet_declared_print_margins(
                "<worksheet><pageMargins left=\"0\" right=\"0.7\" top=\"1.15\" bottom=\"1\" header=\"0.3\" footer=\"0.3\"/></worksheet>"
            ),
            Some(ALL),
            "a zero value is still a declared attribute"
        );
        assert_eq!(
            worksheet_declared_print_margins("<worksheet><sheetData/></worksheet>"),
            Some(NONE),
            "a well-formed sheet with no pageMargins declares nothing"
        );
        assert_eq!(
            worksheet_declared_print_margins("<worksheet/>"),
            Some(NONE),
            "an empty worksheet element declares nothing"
        );
    }

    #[test]
    fn only_the_named_edges_of_a_direct_page_margins_count() {
        assert_eq!(
            worksheet_declared_print_margins(
                "<worksheet><pageMargins left=\"0\" bottom=\"0\"/></worksheet>"
            ),
            Some(DeclaredPrintMargins {
                left: true,
                right: false,
                top: false,
                bottom: true,
                header: false,
                footer: false,
            }),
            "an edge the element omits is not declared by its siblings"
        );
        assert_eq!(
            worksheet_declared_print_margins("<worksheet><pageMargins/></worksheet>"),
            Some(NONE),
            "an empty pageMargins element declares no margins"
        );
    }

    #[test]
    fn header_and_footer_presence_is_tracked_separately_from_body_margins() {
        assert_eq!(
            worksheet_declared_print_margins(
                "<worksheet><pageMargins header=\"0\" footer=\"0\"/></worksheet>"
            ),
            Some(DeclaredPrintMargins {
                header: true,
                footer: true,
                ..NONE
            }),
            "declared zero header/footer values must remain distinguishable from omissions"
        );
    }

    #[test]
    fn a_custom_view_does_not_declare_the_sheets_own_margins() {
        assert_eq!(
            worksheet_declared_print_margins(
                "<worksheet><customSheetViews><customSheetView><pageMargins left=\"0\" right=\"0\" top=\"0\" bottom=\"0\"/></customSheetView></customSheetViews></worksheet>"
            ),
            Some(NONE)
        );
    }

    #[test]
    fn a_prefixed_worksheet_is_read_and_a_malformed_one_is_not_guessed() {
        assert_eq!(
            worksheet_declared_print_margins(
                "<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:pageMargins x:left=\"0\" x:top=\"1\"/></x:worksheet>"
            ),
            Some(DeclaredPrintMargins {
                left: true,
                right: false,
                top: true,
                bottom: false,
                header: false,
                footer: false,
            }),
            "a namespace prefix must not hide a declared edge"
        );
        assert_eq!(
            worksheet_declared_print_margins("<worksheet><pageMargins left=\"0\""),
            None,
            "a part that cannot be parsed states nothing either way"
        );
    }
}
