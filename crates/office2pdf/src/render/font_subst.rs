//! Metric-compatible font substitution table.
//!
//! Maps common Microsoft fonts to open-source metric-compatible alternatives.
//! When the requested font is unavailable, the substitutes are tried in order.
//! Uses a `match` statement for zero-cost static lookup (no runtime allocation).

use std::cell::RefCell;
use std::collections::BTreeSet;
use std::path::PathBuf;

use crate::ir::{
    Block, Document, FixedElementKind, HFInline, HeaderFooter, Page, Paragraph, Table,
};

use super::font_context::{FontSearchContext, resolve_font_search_context};

thread_local! {
    static ACTIVE_FONT_CONTEXT: RefCell<Option<FontSearchContext>> = const { RefCell::new(None) };
}

fn normalized_lookup_key(font_family: &str) -> String {
    let lower = font_family.trim().to_ascii_lowercase();
    if lower.starts_with("pretendard") {
        "pretendard".to_string()
    } else {
        lower
    }
}

fn alias_family(font_family: &str) -> Option<&'static str> {
    let lower = font_family.trim().to_ascii_lowercase();
    if lower.starts_with("pretendard") && lower != "pretendard" {
        Some("Pretendard")
    } else {
        None
    }
}

fn fallback_candidates(font_family: &str, context: Option<&FontSearchContext>) -> Vec<String> {
    let mut candidates: Vec<String> = Vec::new();
    let requested = font_family.trim();

    if let Some(alias) = alias_family(requested)
        && !alias.eq_ignore_ascii_case(requested)
    {
        candidates.push(alias.to_string());
    }

    if let Some(subs) = substitutes(requested) {
        let mut ranked_subs: Vec<(u8, usize, &'static str)> = subs
            .iter()
            .enumerate()
            .filter_map(|(index, sub)| {
                if sub.eq_ignore_ascii_case(requested)
                    || candidates
                        .iter()
                        .any(|candidate| candidate.eq_ignore_ascii_case(sub))
                {
                    return None;
                }
                let rank = context.map(|ctx| ctx.family_source_rank(sub)).unwrap_or(2);
                Some((rank, index, *sub))
            })
            .collect();
        ranked_subs.sort_by_key(|(rank, index, _)| (*rank, *index));
        for (_, _, sub) in ranked_subs {
            candidates.push(sub.to_string());
        }
    }

    candidates
}

/// Return metric-compatible substitute font names for the given font family.
///
/// Returns `None` if no substitution is defined for the font (i.e., it is not
/// a known Microsoft font that has metric-compatible open-source alternatives).
///
/// The returned slice is ordered by preference — the first entry is the best
/// metric-compatible match.
pub fn substitutes(font_family: &str) -> Option<&'static [&'static str]> {
    match normalized_lookup_key(font_family).as_str() {
        "calibri" => Some(&["Carlito", "Liberation Sans"]),
        "cambria" => Some(&["Caladea", "Liberation Serif"]),
        "arial" => Some(&["Liberation Sans", "Arimo"]),
        "times new roman" => Some(&["Liberation Serif", "Tinos"]),
        "courier new" => Some(&["Liberation Mono", "Cousine"]),
        "comic sans ms" => Some(&["Comic Neue"]),
        "verdana" => Some(&["DejaVu Sans"]),
        "georgia" => Some(&["DejaVu Serif"]),
        "consolas" => Some(&["Inconsolata"]),
        "trebuchet ms" => Some(&["Ubuntu"]),
        "impact" => Some(&["Oswald"]),
        "raleway" => Some(&[
            "Helvetica",
            "Arial",
            "Arial Unicode MS",
            "Apple SD Gothic Neo",
            "Noto Sans CJK KR",
            "Malgun Gothic",
            "Liberation Sans",
        ]),
        "lato" => Some(&[
            "Helvetica",
            "Arial",
            "Arial Unicode MS",
            "Apple SD Gothic Neo",
            "Noto Sans CJK KR",
            "Malgun Gothic",
            "Liberation Sans",
        ]),
        "pretendard" => Some(&[
            "Apple SD Gothic Neo",
            "Noto Sans CJK KR",
            "Malgun Gothic",
            "Arial Unicode MS",
            "Helvetica",
            "Arial",
            "Liberation Sans",
        ]),
        _ => None,
    }
}

/// Build a Typst font fallback list string for the given font family.
///
/// If substitutions exist, returns a Typst array literal like
/// `("Calibri", "Carlito", "Liberation Sans")`.
/// If no substitutions exist, returns a simple quoted name like `"Helvetica"`.
pub fn font_with_fallbacks(font_family: &str) -> String {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        font_with_fallbacks_for_context(font_family, active_context.borrow().as_ref())
    })
}

fn font_with_fallbacks_for_context(
    font_family: &str,
    context: Option<&FontSearchContext>,
) -> String {
    let fallbacks = fallback_candidates(font_family, context);
    if fallbacks.is_empty() {
        let mut result = String::with_capacity(font_family.len() + 2);
        result.push('"');
        result.push_str(font_family);
        result.push('"');
        return result;
    }

    let mut result = String::with_capacity(64);
    result.push('(');
    result.push('"');
    result.push_str(font_family);
    result.push('"');
    for sub in fallbacks {
        result.push_str(", \"");
        result.push_str(&sub);
        result.push('"');
    }
    result.push(')');
    result
}

pub(crate) fn with_font_search_context<T>(
    context: Option<&FontSearchContext>,
    operation: impl FnOnce() -> T,
) -> T {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let previous = active_context.replace(context.cloned());
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        active_context.replace(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

fn push_font_family(fonts: &mut BTreeSet<String>, font_family: Option<&str>) {
    if let Some(font_family) = font_family.map(str::trim).filter(|font| !font.is_empty()) {
        fonts.insert(font_family.to_string());
    }
}

fn has_font_family(font_family: Option<&str>) -> bool {
    font_family
        .map(str::trim)
        .is_some_and(|font_family| !font_family.is_empty())
}

fn paragraph_requests_font_family(paragraph: &Paragraph) -> bool {
    paragraph
        .runs
        .iter()
        .any(|run| has_font_family(run.style.font_family.as_deref()))
}

fn table_requests_font_family(table: &Table) -> bool {
    table.rows.iter().any(|row| {
        row.cells
            .iter()
            .any(|cell| cell.content.iter().any(block_requests_font_family))
    })
}

fn header_footer_requests_font_family(header_footer: &HeaderFooter) -> bool {
    header_footer.paragraphs.iter().any(|paragraph| {
        paragraph.elements.iter().any(|inline| match inline {
            HFInline::Run(run) => has_font_family(run.style.font_family.as_deref()),
            HFInline::PageNumber | HFInline::TotalPages => false,
        })
    })
}

fn block_requests_font_family(block: &Block) -> bool {
    match block {
        Block::Paragraph(paragraph) => paragraph_requests_font_family(paragraph),
        Block::Table(table) => table_requests_font_family(table),
        Block::FloatingTextBox(text_box) => text_box.content.iter().any(block_requests_font_family),
        Block::List(list) => list
            .items
            .iter()
            .any(|item| item.content.iter().any(paragraph_requests_font_family)),
        Block::Image(_)
        | Block::FloatingImage(_)
        | Block::MathEquation(_)
        | Block::Chart(_)
        | Block::PageBreak
        | Block::ColumnBreak => false,
    }
}

fn collect_paragraph_fonts(paragraph: &Paragraph, fonts: &mut BTreeSet<String>) {
    for run in &paragraph.runs {
        push_font_family(fonts, run.style.font_family.as_deref());
    }
}

fn collect_table_fonts(table: &Table, fonts: &mut BTreeSet<String>) {
    for row in &table.rows {
        for cell in &row.cells {
            for block in &cell.content {
                collect_block_fonts(block, fonts);
            }
        }
    }
}

fn collect_header_footer_fonts(header_footer: &HeaderFooter, fonts: &mut BTreeSet<String>) {
    for paragraph in &header_footer.paragraphs {
        for inline in &paragraph.elements {
            if let HFInline::Run(run) = inline {
                push_font_family(fonts, run.style.font_family.as_deref());
            }
        }
    }
}

fn collect_block_fonts(block: &Block, fonts: &mut BTreeSet<String>) {
    match block {
        Block::Paragraph(paragraph) => collect_paragraph_fonts(paragraph, fonts),
        Block::Table(table) => collect_table_fonts(table, fonts),
        Block::FloatingTextBox(text_box) => {
            for block in &text_box.content {
                collect_block_fonts(block, fonts);
            }
        }
        Block::List(list) => {
            for item in &list.items {
                for paragraph in &item.content {
                    collect_paragraph_fonts(paragraph, fonts);
                }
            }
        }
        Block::Image(_)
        | Block::FloatingImage(_)
        | Block::MathEquation(_)
        | Block::Chart(_)
        | Block::PageBreak
        | Block::ColumnBreak => {}
    }
}

fn collect_document_font_families(doc: &Document) -> BTreeSet<String> {
    let mut fonts = BTreeSet::new();

    for page in &doc.pages {
        match page {
            Page::Flow(page) => {
                if let Some(header) = &page.header {
                    collect_header_footer_fonts(header, &mut fonts);
                }
                if let Some(footer) = &page.footer {
                    collect_header_footer_fonts(footer, &mut fonts);
                }
                for block in &page.content {
                    collect_block_fonts(block, &mut fonts);
                }
            }
            Page::Fixed(page) => {
                for element in &page.elements {
                    match &element.kind {
                        FixedElementKind::TextBox(text_box) => {
                            for block in &text_box.content {
                                collect_block_fonts(block, &mut fonts);
                            }
                        }
                        FixedElementKind::Table(table) => collect_table_fonts(table, &mut fonts),
                        FixedElementKind::Image(_)
                        | FixedElementKind::Shape(_)
                        | FixedElementKind::SmartArt(_)
                        | FixedElementKind::Chart(_) => {}
                    }
                }
            }
            Page::Table(page) => {
                if let Some(header) = &page.header {
                    collect_header_footer_fonts(header, &mut fonts);
                }
                if let Some(footer) = &page.footer {
                    collect_header_footer_fonts(footer, &mut fonts);
                }
                collect_table_fonts(&page.table, &mut fonts);
            }
        }
    }

    fonts
}

pub(crate) fn document_requests_font_families(doc: &Document) -> bool {
    doc.pages.iter().any(|page| match page {
        Page::Flow(page) => {
            page.header
                .as_ref()
                .is_some_and(header_footer_requests_font_family)
                || page
                    .footer
                    .as_ref()
                    .is_some_and(header_footer_requests_font_family)
                || page.content.iter().any(block_requests_font_family)
        }
        Page::Fixed(page) => page.elements.iter().any(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => {
                text_box.content.iter().any(block_requests_font_family)
            }
            FixedElementKind::Table(table) => table_requests_font_family(table),
            FixedElementKind::Image(_)
            | FixedElementKind::Shape(_)
            | FixedElementKind::SmartArt(_)
            | FixedElementKind::Chart(_) => false,
        }),
        Page::Table(page) => {
            page.header
                .as_ref()
                .is_some_and(header_footer_requests_font_family)
                || page
                    .footer
                    .as_ref()
                    .is_some_and(header_footer_requests_font_family)
                || table_requests_font_family(&page.table)
        }
    })
}

fn resolve_available_fallback(font_family: &str, context: &FontSearchContext) -> Option<String> {
    if context.has_family(font_family) {
        return None;
    }

    fallback_candidates(font_family, Some(context))
        .into_iter()
        .find(|candidate| context.has_family(candidate))
}

#[cfg(not(target_arch = "wasm32"))]
pub fn detect_missing_font_fallbacks(
    doc: &Document,
    font_paths: &[PathBuf],
) -> Vec<(String, String)> {
    let context = resolve_font_search_context(font_paths);
    detect_missing_font_fallbacks_with_context(doc, &context)
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn detect_missing_font_fallbacks_with_context(
    doc: &Document,
    context: &FontSearchContext,
) -> Vec<(String, String)> {
    let requested_fonts = collect_document_font_families(doc);
    if requested_fonts.is_empty() {
        return Vec::new();
    }

    requested_fonts
        .into_iter()
        .filter_map(|font| resolve_available_fallback(&font, context).map(|to| (font, to)))
        .collect()
}

#[cfg(target_arch = "wasm32")]
pub fn detect_missing_font_fallbacks(
    _doc: &Document,
    _font_paths: &[PathBuf],
) -> Vec<(String, String)> {
    Vec::new()
}

#[cfg(test)]
#[path = "font_subst_tests.rs"]
mod tests;
