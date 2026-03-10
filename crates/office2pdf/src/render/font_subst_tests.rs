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
        })],
        styles: crate::ir::StyleSheet::default(),
    };

    assert!(document_requests_font_families(&doc));
}
