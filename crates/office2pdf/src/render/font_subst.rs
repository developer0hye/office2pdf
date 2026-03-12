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
        "microsoft yahei" | "microsoft yahei ui" | "微软雅黑" | "微软雅黑ui" | "微软雅黑 ui" => {
            Some(&[
                "Microsoft YaHei",
                "Microsoft YaHei UI",
                "HYQiHei-55J",
                "HYZhongJianHeiJ",
                "SIL Hei",
                "Heiti SC",
                "Hiragino Sans GB",
                "Arial Unicode MS",
                "Noto Sans CJK SC",
            ])
        }
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
mod tests {
    use super::*;

    // --- substitutes() tests ---

    #[test]
    fn test_calibri_substitutes() {
        let subs = substitutes("Calibri").expect("Calibri should have substitutes");
        assert!(subs.contains(&"Carlito"), "Calibri should map to Carlito");
        assert!(
            subs.contains(&"Liberation Sans"),
            "Calibri should have Liberation Sans as fallback"
        );
        assert_eq!(subs[0], "Carlito", "Carlito should be first preference");
    }

    #[test]
    fn test_cambria_substitutes() {
        let subs = substitutes("Cambria").expect("Cambria should have substitutes");
        assert!(subs.contains(&"Caladea"));
        assert!(subs.contains(&"Liberation Serif"));
    }

    #[test]
    fn test_arial_substitutes() {
        let subs = substitutes("Arial").expect("Arial should have substitutes");
        assert!(subs.contains(&"Liberation Sans"));
        assert!(subs.contains(&"Arimo"));
    }

    #[test]
    fn test_times_new_roman_substitutes() {
        let subs = substitutes("Times New Roman").expect("TNR should have substitutes");
        assert!(subs.contains(&"Liberation Serif"));
        assert!(subs.contains(&"Tinos"));
    }

    #[test]
    fn test_courier_new_substitutes() {
        let subs = substitutes("Courier New").expect("Courier New should have substitutes");
        assert!(subs.contains(&"Liberation Mono"));
        assert!(subs.contains(&"Cousine"));
    }

    #[test]
    fn test_comic_sans_substitutes() {
        let subs = substitutes("Comic Sans MS").expect("Comic Sans MS should have substitutes");
        assert!(subs.contains(&"Comic Neue"));
    }

    #[test]
    fn test_verdana_substitutes() {
        let subs = substitutes("Verdana").expect("Verdana should have substitutes");
        assert!(subs.contains(&"DejaVu Sans"));
    }

    #[test]
    fn test_georgia_substitutes() {
        let subs = substitutes("Georgia").expect("Georgia should have substitutes");
        assert!(subs.contains(&"DejaVu Serif"));
    }

    #[test]
    fn test_unknown_font_returns_none() {
        assert!(
            substitutes("Papyrus").is_none(),
            "Unknown fonts should return None"
        );
        assert!(substitutes("Helvetica").is_none());
        assert!(substitutes("").is_none());
    }

    #[test]
    fn test_case_insensitive_lookup() {
        assert!(substitutes("calibri").is_some(), "lowercase should match");
        assert!(substitutes("CALIBRI").is_some(), "uppercase should match");
        assert!(substitutes("Calibri").is_some(), "title case should match");
        assert!(substitutes("cAlIbRi").is_some(), "mixed case should match");
        assert!(
            substitutes("times new roman").is_some(),
            "lowercase multi-word should match"
        );
        assert!(
            substitutes("TIMES NEW ROMAN").is_some(),
            "uppercase multi-word should match"
        );
    }

    #[test]
    fn test_at_least_8_fonts_mapped() {
        let known_fonts = [
            "Calibri",
            "Cambria",
            "Arial",
            "Times New Roman",
            "Courier New",
            "Comic Sans MS",
            "Verdana",
            "Georgia",
        ];
        let mut mapped = 0;
        for font in &known_fonts {
            if substitutes(font).is_some() {
                mapped += 1;
            }
        }
        assert!(
            mapped >= 8,
            "At least 8 common Microsoft fonts should be mapped, got {mapped}"
        );
    }

    #[test]
    fn test_consolas_substitutes() {
        let subs = substitutes("Consolas").expect("Consolas should have substitutes");
        assert!(subs.contains(&"Inconsolata"));
    }

    #[test]
    fn test_trebuchet_ms_substitutes() {
        let subs = substitutes("Trebuchet MS").expect("Trebuchet MS should have substitutes");
        assert!(subs.contains(&"Ubuntu"));
    }

    #[test]
    fn test_impact_substitutes() {
        let subs = substitutes("Impact").expect("Impact should have substitutes");
        assert!(subs.contains(&"Oswald"));
    }

    #[test]
    fn test_raleway_substitutes() {
        let subs = substitutes("Raleway").expect("Raleway should have substitutes");
        assert!(subs.contains(&"Helvetica"));
        assert!(subs.contains(&"Arial"));
        assert!(subs.contains(&"Arial Unicode MS"));
        assert!(subs.contains(&"Apple SD Gothic Neo"));
        assert_eq!(subs[0], "Helvetica");
    }

    #[test]
    fn test_lato_substitutes() {
        let subs = substitutes("Lato").expect("Lato should have substitutes");
        assert!(subs.contains(&"Helvetica"));
        assert!(subs.contains(&"Arial"));
        assert!(subs.contains(&"Arial Unicode MS"));
        assert!(subs.contains(&"Apple SD Gothic Neo"));
    }

    #[test]
    fn test_pretendard_substitutes() {
        let subs = substitutes("Pretendard").expect("Pretendard should have substitutes");
        assert_eq!(subs[0], "Apple SD Gothic Neo");
        assert!(subs.contains(&"Noto Sans CJK KR"));
        assert!(subs.contains(&"Malgun Gothic"));
    }

    #[test]
    fn test_microsoft_yahei_substitutes() {
        let subs = substitutes("微软雅黑").expect("微软雅黑 should have substitutes");
        assert_eq!(subs[0], "Microsoft YaHei");
        assert!(subs.contains(&"Hiragino Sans GB"));
        assert!(subs.contains(&"Heiti SC"));
    }

    // --- font_with_fallbacks() tests ---

    #[test]
    fn test_font_with_fallbacks_known_font() {
        let result = font_with_fallbacks("Calibri");
        assert_eq!(
            result, r#"("Calibri", "Carlito", "Liberation Sans")"#,
            "Known font should produce Typst array with original + substitutes"
        );
    }

    #[test]
    fn test_font_with_fallbacks_unknown_font() {
        let result = font_with_fallbacks("Helvetica");
        assert_eq!(
            result, "\"Helvetica\"",
            "Unknown font should produce simple quoted string"
        );
    }

    #[test]
    fn test_font_with_fallbacks_single_substitute() {
        let result = font_with_fallbacks("Comic Sans MS");
        assert_eq!(result, r#"("Comic Sans MS", "Comic Neue")"#);
    }

    #[test]
    fn test_font_with_fallbacks_preserves_original_case() {
        // The original font name should appear as-is (not lowercased)
        let result = font_with_fallbacks("CALIBRI");
        assert!(
            result.starts_with("(\"CALIBRI\""),
            "Original case should be preserved: {result}"
        );
    }

    #[test]
    fn test_font_with_fallbacks_pretendard_variant_includes_base_family() {
        let result = font_with_fallbacks("Pretendard SemiBold");
        assert!(
            result.contains("\"Pretendard\""),
            "Pretendard variants should fall back to the base family: {result}"
        );
        assert!(
            result.contains("\"Apple SD Gothic Neo\""),
            "Pretendard variants should include Korean-capable fallbacks: {result}"
        );
    }

    #[test]
    fn test_resolve_available_fallback_prefers_alias_before_system_fallback() {
        let context = FontSearchContext::for_test(
            Vec::new(),
            &["Pretendard", "Apple SD Gothic Neo"],
            &[],
            &[],
        );
        let fallback = resolve_available_fallback("Pretendard Medium", &context);
        assert_eq!(fallback.as_deref(), Some("Pretendard"));
    }

    #[test]
    fn test_font_with_fallbacks_prefers_office_source_rank_over_static_substitute_order() {
        let context = FontSearchContext::for_test(
            Vec::new(),
            &["Apple SD Gothic Neo", "Malgun Gothic"],
            &["Malgun Gothic"],
            &[],
        );
        let result = with_font_search_context(Some(&context), || font_with_fallbacks("Pretendard"));
        let apple_index = result
            .find("\"Apple SD Gothic Neo\"")
            .expect("Apple SD Gothic Neo should appear in fallback list");
        let malgun_index = result
            .find("\"Malgun Gothic\"")
            .expect("Malgun Gothic should appear in fallback list");
        assert!(
            malgun_index < apple_index,
            "office-resolved font should outrank static substitute order: {result}"
        );
    }

    #[test]
    fn test_font_with_fallbacks_prefers_wps_chinese_font_for_microsoft_yahei() {
        let context = FontSearchContext::for_test(
            Vec::new(),
            &["HYQiHei-55J", "Heiti SC"],
            &["HYQiHei-55J"],
            &[],
        );
        let result = with_font_search_context(Some(&context), || font_with_fallbacks("微软雅黑"));
        let wps_index = result
            .find("\"HYQiHei-55J\"")
            .expect("WPS Chinese font should appear in fallback list");
        let heiti_index = result
            .find("\"Heiti SC\"")
            .expect("Heiti SC should appear in fallback list");
        assert!(
            wps_index < heiti_index,
            "WPS font should outrank generic macOS fallback for 微软雅黑: {result}"
        );
    }

    #[test]
    fn test_detect_missing_font_fallbacks_with_context_prefers_office_font() {
        let context = FontSearchContext::for_test(
            Vec::new(),
            &["Malgun Gothic", "Apple SD Gothic Neo"],
            &["Malgun Gothic"],
            &[],
        );
        let doc = Document {
            metadata: crate::ir::Metadata::default(),
            pages: vec![Page::Flow(crate::ir::FlowPage {
                size: crate::ir::PageSize::default(),
                margins: crate::ir::Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: crate::ir::ParagraphStyle::default(),
                    runs: vec![crate::ir::Run {
                        text: "Title".to_string(),
                        style: crate::ir::TextStyle {
                            font_family: Some("Pretendard Medium".to_string()),
                            ..crate::ir::TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                columns: None,
                page_number_start: None,
            })],
            styles: crate::ir::StyleSheet::default(),
        };

        let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);
        assert_eq!(
            fallbacks,
            vec![("Pretendard Medium".to_string(), "Malgun Gothic".to_string())]
        );
    }

    #[test]
    fn test_document_requests_font_families_false_when_all_runs_use_defaults() {
        let doc = Document {
            metadata: crate::ir::Metadata::default(),
            pages: vec![Page::Flow(crate::ir::FlowPage {
                size: crate::ir::PageSize::default(),
                margins: crate::ir::Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: crate::ir::ParagraphStyle::default(),
                    runs: vec![crate::ir::Run {
                        text: "Plain text".to_string(),
                        style: crate::ir::TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                columns: None,
                page_number_start: None,
            })],
            styles: crate::ir::StyleSheet::default(),
        };

        assert!(!document_requests_font_families(&doc));
    }

    #[test]
    fn test_document_requests_font_families_true_when_any_run_sets_family() {
        let doc = Document {
            metadata: crate::ir::Metadata::default(),
            pages: vec![Page::Flow(crate::ir::FlowPage {
                size: crate::ir::PageSize::default(),
                margins: crate::ir::Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: crate::ir::ParagraphStyle::default(),
                    runs: vec![crate::ir::Run {
                        text: "Styled text".to_string(),
                        style: crate::ir::TextStyle {
                            font_family: Some("Pretendard".to_string()),
                            ..crate::ir::TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                columns: None,
                page_number_start: None,
            })],
            styles: crate::ir::StyleSheet::default(),
        };

        assert!(document_requests_font_families(&doc));
    }
}
