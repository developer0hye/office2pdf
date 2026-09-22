#![cfg(not(target_arch = "wasm32"))] // native-only unit tests (filesystem, system fonts)
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
fn calibri_light_resolves_through_the_calibri_chain() {
    // `Calibri Light` is the `majorHAnsi` face of every Office theme since
    // 2013, so it is the face of every built-in `Heading N`. Keyed on the exact
    // name it matched nothing, so a theme-headed run got no substitute chain
    // and no metrics at all (issue #1197).
    let subs = substitutes("Calibri Light").expect("Calibri Light should have substitutes");
    assert_eq!(subs, &["Calibri", "Carlito", "Liberation Sans"]);
    assert_eq!(
        subs[0], "Calibri",
        "the base family carries Calibri Light's own hhea metrics, so it leads"
    );
}

#[test]
fn calibri_light_and_calibri_read_the_same_line_metrics() {
    // Both faces declare hhea 1950/-550/0 on 2048 upem, so whichever member of
    // the shared chain a host resolves, the two families must agree — that
    // equality is what lets a theme-headed paragraph take Word's line box
    // (issue #1197).
    let Some(calibri) = crate::render::pdf::font_line_metrics_em("Calibri") else {
        return; // no Calibri-compatible face on this host
    };
    let light = crate::render::pdf::font_line_metrics_em("Calibri Light")
        .expect("Calibri Light must resolve wherever Calibri does");
    assert_eq!(light, calibri);
}

#[test]
fn test_carlito_substitutes_stay_sans_serif() {
    let subs = substitutes("Carlito").expect("Carlito should have substitutes");
    assert_eq!(subs, &["Calibri", "Liberation Sans", "Arimo", "Arial"]);
    assert!(subs.iter().all(|family| !family.contains("Serif")));
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
fn named_monospace_families_get_a_monospace_fallback_chain() {
    for family in ["Lucida Sans Typewriter", "JetBrains Mono", "IBM Plex Mono"] {
        let substitutes = substitutes(family)
            .unwrap_or_else(|| panic!("{family} should have class-preserving substitutes"));
        assert_eq!(
            substitutes,
            &[
                "DejaVu Sans Mono",
                "Noto Sans Mono",
                "Liberation Mono",
                "Cousine",
            ],
            "{family} must not fall through to a proportional serif"
        );
    }
}

#[test]
fn monotype_brand_name_is_not_misclassified_as_monospace() {
    assert_eq!(substitutes("Monotype Corsiva"), None);
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
fn office_poster_fonts_use_the_bundled_noto_serif_fallback() {
    // LibreOffice 25.2 resolves the unembedded poster-template faces in
    // issue #1458 to Noto Serif 2.015. Keeping that exact OFL face as the
    // first named substitute makes widths and Word line metrics independent
    // of whichever serif a host happens to install.
    for family in [
        "Avenir Next LT Pro",
        "Avenir Next W1G Medium",
        "The Hand Black",
    ] {
        assert_eq!(
            substitutes(family),
            Some(&["Noto Serif"][..]),
            "{family} should resolve through the reproducible Noto Serif face"
        );
    }
}

#[test]
fn segoe_ui_uses_the_bundled_sans_fallback() {
    // The issue #982 workbook prints its instruction panel in Segoe UI. On a
    // host without that proprietary face, a one-family Typst request reaches
    // the engine's serif default and changes the native nine-line wrap
    // boundaries (issue #1472). Keep the replacement deterministic and in the
    // requested family's sans-serif class.
    assert_eq!(
        substitutes("Segoe UI"),
        Some(&["Selawik"][..]),
        "missing Segoe UI must not fall through to Libertinus Serif"
    );
}

#[test]
fn segoe_ui_metric_substitute_reproduces_native_instruction_line_widths() {
    // `mutool draw -F trace` on Excel for Mac 16.112's page-2 export of Gift
    // Budget and Tracker reads these nine Segoe UI advances. They are the
    // exact wrap boundaries that issue #1472 lost when the run reached
    // Libertinus Serif. Resolve through only the bundled faces so the test is
    // independent of whether the runner itself has Segoe UI installed.
    let measured = [
        ("Fill out as much as you can at the", 14.667001),
        ("beginning of each new year. Enter", 14.975001),
        ("names of family and friends, select", 15.202001),
        ("gift occasion, select month of", 12.963001),
        ("event, and type in how much you", 14.663001),
        ("want to spend. Each person may", 14.271001),
        ("appear multiple times within the", 14.259001),
        ("table for each different occasion", 14.197001),
        ("(birthday, holiday, shower, etc.).", 13.888001),
    ];
    let context = crate::render::font_context::resolve_font_search_context_from_fonts(
        crate::bundled_fonts::selawik_fonts(),
    );

    with_font_search_context(Some(&context), || {
        for (text, native_advance_em) in measured {
            let actual = crate::render::pdf::text_advance_em("Segoe UI", false, text)
                .unwrap_or_else(|| panic!("the bundled Selawik face must measure {text:?}"));
            assert!(
                (actual - native_advance_em).abs() < 0.005,
                "{text:?}: native Segoe UI advances {native_advance_em}em, bundled fallback gives {actual}em"
            );
        }
    });
}

#[test]
fn an_available_segoe_ui_face_still_leads_the_bundled_substitute() {
    // Selawik is a missing-face replacement, not an override for an explicit
    // Segoe UI conversion input. Preserve the declared face whenever the
    // active font context can actually supply it (issue #1472).
    let context =
        FontSearchContext::for_test(Vec::new(), &["Segoe UI", "Selawik"], &["Segoe UI"], &[]);
    let chain = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("Segoe UI", "Gift Budget")
    });

    assert!(
        chain.starts_with("(\"Segoe UI\""),
        "an available Segoe UI face must remain authoritative: {chain}"
    );
    assert_eq!(
        resolve_available_fallback("Segoe UI", TextScript::Latin, &context),
        None
    );
}

#[test]
fn only_office_poster_requests_select_the_bundled_noto_serif() {
    use crate::ir::{
        Block, Document, FlowPage, Margins, Metadata, Page, PageSize, Paragraph, ParagraphStyle,
        Run, StyleSheet, TextStyle,
    };

    let document = |family: &str| Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Poster".to_string(),
                    style: TextStyle {
                        font_family: Some(family.to_string()),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: StyleSheet::default(),
    };

    for family in [
        "Avenir Next LT Pro",
        "Avenir Next W1G Medium",
        "The Hand Black",
    ] {
        assert!(document_requests_bundled_noto_serif(&document(family)));
    }
    assert!(!document_requests_bundled_noto_serif(&document("Calibri")));
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

// --- font_with_fallbacks_for_text() tests ---
// Latin-only text, so these exercise the family chain alone.

#[test]
fn test_font_with_fallbacks_known_font() {
    let result = font_with_fallbacks_for_text("Calibri", "");
    assert_eq!(
        result, r#"("Calibri", "Carlito", "Liberation Sans", "Arimo", "DejaVu Sans", "Helvetica")"#,
        "Known font should produce Typst array with original + substitutes, \
         ending on its own class so an absent Carlito cannot land it on a serif"
    );
}

#[test]
fn test_carlito_font_with_fallbacks_emits_sans_chain() {
    let result = font_with_fallbacks_for_text("Carlito", "");
    assert_eq!(
        result,
        r#"("Carlito", "Calibri", "Liberation Sans", "Arimo", "Arial", "DejaVu Sans", "Helvetica")"#
    );
}

#[test]
fn test_carlito_installed_system_fallback_is_ranked_first() {
    let context = FontSearchContext::for_test(Vec::new(), &["Arial"], &[], &[]);
    let result = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("Carlito", "")
    });
    let arial_index = result
        .find("\"Arial\"")
        .expect("Arial should remain in the fallback list");
    let calibri_index = result
        .find("\"Calibri\"")
        .expect("Calibri should remain in the fallback list");
    assert!(
        arial_index < calibri_index,
        "an installed system sans font should outrank unavailable candidates: {result}"
    );
}

#[test]
fn unembedded_aptos_ignores_the_incidental_host_face() {
    // `Place your event title here.docx` declares Aptos for its 8pt footer,
    // but does not embed it. LibreOffice resolves that run to Noto Sans. The
    // converter must make the same choice whether the converting host happens
    // to have Microsoft Office's Aptos installed or not (issue #1463).
    let host_context =
        FontSearchContext::for_test(Vec::new(), &["Aptos", "Noto Sans"], &["Aptos"], &[]);
    let chain = with_font_search_context(Some(&host_context), || {
        font_with_fallbacks_for_text("Aptos", "Sensitivity: Internal")
    });

    assert!(
        chain.starts_with("(\"Noto Sans\""),
        "an incidental host Aptos must not lead the paint chain: {chain}"
    );
    assert_eq!(
        resolve_available_fallback("Aptos", TextScript::Latin, &host_context).as_deref(),
        Some("Noto Sans"),
        "the warning path must report the same family that paints"
    );
}

#[test]
fn caller_supplied_aptos_still_leads_its_own_chain() {
    // A caller-provided font is an explicit conversion input, unlike a face
    // discovered incidentally from the host. Preserve that opt-in path while
    // making the unembedded document deterministic (issue #1463).
    let user_context =
        FontSearchContext::for_test(Vec::new(), &["Aptos", "Noto Sans"], &[], &["Aptos"]);
    let chain = with_font_search_context(Some(&user_context), || {
        font_with_fallbacks_for_text("Aptos", "Sensitivity: Internal")
    });

    assert!(
        chain.starts_with("(\"Aptos\""),
        "an explicitly supplied Aptos face remains authoritative: {chain}"
    );
    assert_eq!(
        resolve_available_fallback("Aptos", TextScript::Latin, &user_context),
        None
    );
}

#[test]
fn missing_typewriter_face_resolves_to_an_available_monospace_face() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Libertinus Serif", "DejaVu Sans Mono"],
        &[],
        &[],
    );

    let fallback =
        resolve_available_fallback("Lucida Sans Typewriter", TextScript::Latin, &context);

    assert_eq!(fallback.as_deref(), Some("DejaVu Sans Mono"));
}

#[test]
fn a_listed_family_whose_substitutes_are_all_absent_keeps_its_own_class() {
    // Every name in the substitution table is a real font, so a chain of them
    // can run out: a host with neither Carlito nor Liberation Sans left a
    // Calibri run — the `minorFont` of every default Office theme, and so the
    // face an XLSX chart draws its labels in — with nothing but the engine's
    // default, which is a serif for a family whose own PANOSE says sans
    // (issue #1213).
    //
    // One case per class, and more than one family per class, so the rule is
    // "end on the family's own class" rather than a patch for Calibri.
    let sans_only_host: FontSearchContext =
        FontSearchContext::for_test(Vec::new(), &["Libertinus Serif", "DejaVu Sans"], &[], &[]);
    for family in ["Calibri", "Calibri Light", "Verdana", "Corbel"] {
        assert_eq!(
            resolve_available_fallback(family, TextScript::Latin, &sans_only_host).as_deref(),
            Some("DejaVu Sans"),
            "{family} must reach an available sans rather than the default serif"
        );
    }

    // The class has to be the family's own, not whichever face happens to be
    // installed: a serif family on the same host must not take the sans.
    let mixed_host: FontSearchContext =
        FontSearchContext::for_test(Vec::new(), &["DejaVu Sans", "DejaVu Serif"], &[], &[]);
    for family in ["Cambria", "Times New Roman"] {
        assert_eq!(
            resolve_available_fallback(family, TextScript::Latin, &mixed_host).as_deref(),
            Some("DejaVu Serif"),
            "{family} is a serif and must stay one"
        );
    }

    let fixed_pitch_host: FontSearchContext =
        FontSearchContext::for_test(Vec::new(), &["DejaVu Sans", "Cousine"], &[], &[]);
    for family in ["Consolas", "Courier New"] {
        assert_eq!(
            resolve_available_fallback(family, TextScript::Latin, &fixed_pitch_host).as_deref(),
            Some("Cousine"),
            "{family} is fixed pitch and must stay fixed pitch"
        );
    }
}

#[test]
fn the_class_tail_follows_the_metric_compatible_substitutes() {
    // The tail is a last resort, not a preference: Carlito and Liberation Sans
    // are metric-compatible with Calibri and a generic sans is not, so the
    // order the table states has to survive the addition (issue #1213).
    let painted: String = font_with_fallbacks_for_text("Calibri", "Sales");
    let position = |family: &str| {
        painted
            .find(&format!("\"{family}\""))
            .unwrap_or_else(|| panic!("{family} should be in the chain: {painted}"))
    };

    assert!(position("Carlito") < position("Arimo"));
    assert!(position("Liberation Sans") < position("Arimo"));
    assert!(position("Arimo") < position("Helvetica"));
    assert!(
        !painted.contains("Serif"),
        "a sans family must not gain a serif candidate: {painted}"
    );
}

#[test]
fn the_class_tail_cannot_outrank_an_available_family_substitute() {
    // Source priority chooses between the family's own substitutes, but the
    // generic class tail is a last resort. An Office-bundled generic must not
    // jump ahead of an installed metric-compatible stand-in (issue #1213).
    let context =
        FontSearchContext::for_test(Vec::new(), &["Carlito", "Helvetica"], &["Helvetica"], &[]);
    let painted: String = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("Calibri", "Sales")
    });
    let carlito: usize = painted
        .find("\"Carlito\"")
        .expect("Calibri's own substitute is present");
    let helvetica: usize = painted
        .find("\"Helvetica\"")
        .expect("the generic sans tail is present");

    assert!(
        carlito < helvetica,
        "the generic class tail must follow the family's substitutes: {painted}"
    );
}

#[test]
fn a_metrics_lookup_stops_at_the_metric_compatible_substitutes() {
    // A metrics lookup resolves one face and reads its numbers whole, so the
    // generic tail must not reach it: those numbers are the stand-in's, not the
    // family's. Trebuchet MS is the case that showed it — its only substitute
    // is Ubuntu, and on a host without it the tail handed a fit-to-page sheet's
    // column unit Liberation Sans's digit advance, re-scaling the whole sheet
    // by 13% (issue #1213).
    // Families whose own substitutes name no generic, so a generic appearing in
    // the metrics list can only have come from the tail. Calibri is excluded
    // for the opposite reason: Liberation Sans is one of its own.
    for family in ["Trebuchet MS", "Comic Sans MS", "Consolas", "Malgun Gothic"] {
        let measured: Vec<String> = family_candidates(family);
        let generic: Option<&String> = measured.iter().find(|candidate| {
            SANS_SERIF_SUBSTITUTES.contains(&candidate.as_str())
                || SERIF_SUBSTITUTES.contains(&candidate.as_str())
                || MONOSPACE_SUBSTITUTES.contains(&candidate.as_str())
        });
        assert_eq!(
            generic, None,
            "{family} must not read its metrics from a generic face: {measured:?}"
        );

        // The list Typst paints from still ends on the class, which is the
        // whole point of the tail.
        let painted: String = font_with_fallbacks_for_text(family, "Sales");
        assert!(
            painted.contains("Helvetica") || painted.contains("DejaVu"),
            "{family} must still paint in its own class: {painted}"
        );
    }
}

#[test]
fn test_font_with_fallbacks_unknown_font() {
    let result = font_with_fallbacks_for_text("Helvetica", "");
    assert_eq!(
        result, "\"Helvetica\"",
        "Unknown font should produce simple quoted string"
    );
}

#[test]
fn configured_last_resort_is_appended_after_every_regular_candidate() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Source Han Sans SC"],
        &[],
        &["Source Han Sans SC"],
    )
    .with_last_resort_family(Some("Source Han Sans SC"));

    let result = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("SimSun", "中文")
    });

    assert!(
        result.ends_with(", \"Source Han Sans SC\")"),
        "the caller-selected face must be the last candidate: {result}"
    );
}

#[test]
fn configured_last_resort_is_appended_to_east_asian_split_chain() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Source Han Sans SC"],
        &[],
        &["Source Han Sans SC"],
    )
    .with_last_resort_family(Some("Source Han Sans SC"));

    let result = with_font_search_context(Some(&context), || {
        font_with_east_asian_fallbacks("Calibri", "SimSun", "Hello 中文")
    });

    assert!(
        result.ends_with(", \"Source Han Sans SC\")"),
        "the caller-selected face must be the last candidate: {result}"
    );
}

#[test]
fn missing_cjk_coverage_reports_notdef_instead_of_staying_silent() {
    let context = FontSearchContext::for_test(Vec::new(), &[], &[], &[]);
    let doc = korean_document_requesting("SimSun", "中文");

    let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);

    assert_eq!(
        fallbacks,
        vec![("SimSun".to_string(), ".notdef".to_string())]
    );
}

#[test]
fn test_font_with_fallbacks_single_named_substitute_then_the_class_tail() {
    let result = font_with_fallbacks_for_text("Comic Sans MS", "");
    assert_eq!(
        result,
        r#"("Comic Sans MS", "Comic Neue", "Liberation Sans", "Arimo", "DejaVu Sans", "Helvetica")"#
    );
}

// Family names come from parsed OOXML, i.e. document-controlled input. A name
// containing `"` or `\` must not break out of the generated Typst string
// literal (it would corrupt the whole generated source).

#[test]
fn test_font_with_fallbacks_escapes_quotes_in_family_name() {
    let result = font_with_fallbacks_for_text(r#"Weird "Quoted" Font"#, "");
    assert_eq!(result, r#""Weird \"Quoted\" Font""#);
}

#[test]
fn test_font_with_fallbacks_escapes_backslashes_in_family_name() {
    let result = font_with_fallbacks_for_text(r"Fonts\Custom", "");
    assert_eq!(result, r#""Fonts\\Custom""#);
}

#[test]
fn test_font_with_fallbacks_escapes_quotes_in_fallback_array_head() {
    // "Pretendard <anything>" resolves to the Pretendard substitute chain, so
    // the raw (quote-carrying) name lands at the head of a Typst array literal.
    let result = font_with_fallbacks_for_text(r#"Pretendard "Display""#, "");
    assert!(
        result.starts_with(r#"("Pretendard \"Display\"""#),
        "array head must escape document-supplied quotes: {result}"
    );
}

#[test]
fn test_font_with_fallbacks_preserves_original_case() {
    // The original font name should appear as-is (not lowercased)
    let result = font_with_fallbacks_for_text("CALIBRI", "");
    assert!(
        result.starts_with("(\"CALIBRI\""),
        "Original case should be preserved: {result}"
    );
}

#[test]
fn test_font_with_fallbacks_pretendard_variant_includes_base_family() {
    let result = font_with_fallbacks_for_text("Pretendard SemiBold", "");
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
    let context =
        FontSearchContext::for_test(Vec::new(), &["Pretendard", "Apple SD Gothic Neo"], &[], &[]);
    let fallback = resolve_available_fallback("Pretendard Medium", TextScript::Latin, &context);
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
    let result = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("Pretendard", "")
    });
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
            first_header: None,
            first_footer: None,
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
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };

    let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);
    assert_eq!(
        fallbacks,
        vec![("Pretendard Medium".to_string(), "Malgun Gothic".to_string())]
    );
}

/// A Korean workbook whose styles name a Simplified Chinese Latin face must
/// report the face the text actually lands on.
///
/// `Noto Sans CJK SC`'s substitute chain reaches Microsoft YaHei, but the run
/// holds Hangul, so the renderer's script chain puts Malgun Gothic ahead of it
/// and that is what the PDF embeds. Reporting the substitute sends anyone
/// debugging Korean output after a YaHei substitution that never happened
/// (issue #617).
#[test]
fn test_detect_missing_font_fallbacks_reports_script_resolved_face() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Microsoft YaHei", "Malgun Gothic"],
        &["Microsoft YaHei", "Malgun Gothic"],
        &[],
    );
    let doc = korean_document_requesting("Noto Sans CJK SC", "견적서");

    let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);
    assert_eq!(
        fallbacks,
        vec![("Noto Sans CJK SC".to_string(), "Malgun Gothic".to_string())]
    );
}

/// The same family over Latin-only text has no script chain to consult, so the
/// metric substitute is what renders and what the warning must name.
#[test]
fn test_detect_missing_font_fallbacks_reports_substitute_for_latin_text() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Microsoft YaHei", "Malgun Gothic"],
        &["Microsoft YaHei", "Malgun Gothic"],
        &[],
    );
    let doc = korean_document_requesting("Noto Sans CJK SC", "Quotation");

    let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);
    assert_eq!(
        fallbacks,
        vec![(
            "Noto Sans CJK SC".to_string(),
            "Microsoft YaHei".to_string()
        )]
    );
}

/// One family carrying two scripts resolves differently per script, so both
/// resolutions are reported rather than one standing in for the other.
#[test]
fn test_detect_missing_font_fallbacks_reports_each_script_separately() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Microsoft YaHei", "Malgun Gothic"],
        &["Microsoft YaHei", "Malgun Gothic"],
        &[],
    );
    let mut doc = korean_document_requesting("Noto Sans CJK SC", "견적서");
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!("korean_document_requesting builds a flow page")
    };
    page.content.push(Block::Paragraph(Paragraph {
        style: crate::ir::ParagraphStyle::default(),
        runs: vec![crate::ir::Run {
            text: "Quotation".to_string(),
            style: crate::ir::TextStyle {
                font_family: Some("Noto Sans CJK SC".to_string()),
                ..crate::ir::TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    }));

    let fallbacks = detect_missing_font_fallbacks_with_context(&doc, &context);
    assert_eq!(
        fallbacks,
        vec![
            ("Noto Sans CJK SC".to_string(), "Malgun Gothic".to_string()),
            (
                "Noto Sans CJK SC".to_string(),
                "Microsoft YaHei".to_string()
            ),
        ]
    );
}

fn korean_document_requesting(font_family: &str, text: &str) -> Document {
    Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Flow(crate::ir::FlowPage {
            first_header: None,
            first_footer: None,
            size: crate::ir::PageSize::default(),
            margins: crate::ir::Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: crate::ir::ParagraphStyle::default(),
                runs: vec![crate::ir::Run {
                    text: text.to_string(),
                    style: crate::ir::TextStyle {
                        font_family: Some(font_family.to_string()),
                        ..crate::ir::TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                }],
            })],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    }
}

#[test]
fn test_document_requests_font_families_false_when_all_runs_use_defaults() {
    let doc = Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Flow(crate::ir::FlowPage {
            first_header: None,
            first_footer: None,
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
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };

    assert!(!document_requests_font_families(&doc));
}

/// Both families whose static substitution chain is complete enough to skip
/// the availability scan: Arial, and the Times New Roman a DOCX naming no
/// `w:rFonts` resolves (issue #1196).
///
/// The skip is what keeps a package that states no font at all from reporting
/// a substitution on every host that does not install the face — the scan is
/// the only thing that emits `FallbackUsed`, and before #1196 the synthetic
/// default was Arial, so no such package ever reached it.
#[test]
fn test_document_requests_font_families_false_for_context_free_families() {
    for family in ["Arial", "Times New Roman"] {
        let doc = Document {
            metadata: crate::ir::Metadata::default(),
            pages: vec![Page::Flow(crate::ir::FlowPage {
                first_header: None,
                first_footer: None,
                size: crate::ir::PageSize::default(),
                margins: crate::ir::Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: crate::ir::ParagraphStyle::default(),
                    runs: vec![crate::ir::Run {
                        text: "DOCX default text".to_string(),
                        style: crate::ir::TextStyle {
                            font_family: Some(family.to_string()),
                            ..crate::ir::TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                columns: None,
                line_grid_pitch: None,
                line_grid_snaps_lines: false,
                page_numbering: None,
            })],
            styles: crate::ir::StyleSheet::default(),
        };

        assert!(
            !document_requests_font_families(&doc),
            "{family} resolves through its static chain and needs no scan"
        );
    }
}

#[test]
fn test_document_requests_font_families_true_when_any_run_sets_family() {
    let doc = Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Flow(crate::ir::FlowPage {
            first_header: None,
            first_footer: None,
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
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };

    assert!(document_requests_font_families(&doc));
}

// --- Korean / CJK font name tests ---

#[test]
fn test_korean_malgun_gothic_name_has_substitutes() {
    let subs = substitutes("맑은 고딕").expect("Korean Malgun Gothic name should have substitutes");
    assert!(
        subs.contains(&"Malgun Gothic"),
        "Should include English name as fallback: {subs:?}"
    );
}

#[test]
fn test_korean_gulim_name_has_substitutes() {
    let subs = substitutes("굴림").expect("Korean Gulim name should have substitutes");
    assert!(subs.contains(&"Gulim"));
}

#[test]
fn test_font_with_fallbacks_korean_malgun_gothic_includes_english_name() {
    let result = font_with_fallbacks_for_text("맑은 고딕", "");
    assert!(
        result.contains("\"Malgun Gothic\""),
        "Should include English name in fallback list: {result}"
    );
    assert!(
        result.starts_with("(\"맑은 고딕\""),
        "Original name should be preserved first: {result}"
    );
}

#[test]
fn test_japanese_font_name_has_substitutes() {
    let subs = substitutes("メイリオ").expect("Japanese Meiryo name should have substitutes");
    assert!(subs.contains(&"Meiryo"));
}

#[test]
fn test_chinese_font_name_has_substitutes() {
    let subs = substitutes("微软雅黑").expect("Chinese YaHei name should have substitutes");
    assert!(subs.contains(&"Microsoft YaHei"));
}

// --- is_primary_font_available() tests ---

#[test]
fn test_is_primary_font_available_returns_true_when_no_context() {
    // When no font context is active (e.g. WASM), assume available.
    assert!(is_primary_font_available("Anything"));
}

#[test]
fn test_is_primary_font_available_returns_true_when_font_exists() {
    let context = FontSearchContext::for_test(Vec::new(), &["Pretendard"], &[], &[]);
    let result =
        with_font_search_context(Some(&context), || is_primary_font_available("Pretendard"));
    assert!(result);
}

#[test]
fn test_is_primary_font_available_returns_true_via_alias() {
    // "Pretendard ExtraBold" → alias "Pretendard" → available
    let context = FontSearchContext::for_test(Vec::new(), &["Pretendard"], &[], &[]);
    let result = with_font_search_context(Some(&context), || {
        is_primary_font_available("Pretendard ExtraBold")
    });
    assert!(result);
}

#[test]
fn test_is_primary_font_available_returns_false_when_missing() {
    let context = FontSearchContext::for_test(Vec::new(), &["Arial"], &[], &[]);
    let result = with_font_search_context(Some(&context), || {
        is_primary_font_available("Pretendard ExtraBold")
    });
    assert!(!result);
}

// --- Noto CJK family substitutes (issue #290) ---

#[test]
fn test_noto_sans_cjk_kr_substitutes() {
    let subs = substitutes("Noto Sans CJK KR").expect("Noto Sans CJK KR should have substitutes");
    assert!(subs.contains(&"Apple SD Gothic Neo"));
    assert!(subs.contains(&"Malgun Gothic"));
}

#[test]
fn test_noto_sans_cjk_sc_substitutes() {
    let subs = substitutes("Noto Sans CJK SC").expect("Noto Sans CJK SC should have substitutes");
    assert!(subs.contains(&"PingFang SC"));
}

#[test]
fn test_noto_sans_cjk_jp_substitutes() {
    let subs = substitutes("Noto Sans CJK JP").expect("Noto Sans CJK JP should have substitutes");
    assert!(subs.contains(&"Hiragino Sans"));
}

#[test]
fn test_noto_sans_cjk_tc_substitutes() {
    let subs = substitutes("Noto Sans CJK TC").expect("Noto Sans CJK TC should have substitutes");
    assert!(subs.contains(&"PingFang TC"));
}

#[test]
fn test_noto_serif_cjk_kr_substitutes() {
    let subs = substitutes("Noto Serif CJK KR").expect("Noto Serif CJK KR should have substitutes");
    assert!(subs.contains(&"Apple Myungjo") || subs.contains(&"Batang"));
}

#[test]
fn test_noto_sans_kr_short_name_substitutes() {
    // Google Fonts ships the short-name variants ("Noto Sans KR"); they must
    // resolve the same way as the CJK superfamily names.
    let subs = substitutes("Noto Sans KR").expect("Noto Sans KR should have substitutes");
    assert!(subs.contains(&"Apple SD Gothic Neo"));
}

#[test]
fn east_asian_family_follows_the_latin_one_in_the_font_list() {
    // Typst resolves a font list per glyph, so listing the Latin family first
    // and the East Asian family straight after reproduces Word's split: a
    // Latin face has no Hangul, so the Hangul lands on the declared East
    // Asian face (issue #575).
    let list = font_with_east_asian_fallbacks("Calibri", "맑은 고딕", "본문");

    let latin = list.find("\"Calibri\"").expect("the Latin family leads");
    let east_asian = list
        .find("\"Malgun Gothic\"")
        .expect("the localized name resolves to the English one");
    assert!(latin < east_asian, "the Latin family comes first: {list}");
    assert!(
        list.starts_with('('),
        "a two-family run emits a list: {list}"
    );
}

#[test]
fn a_run_naming_the_same_family_twice_does_not_repeat_it() {
    let list = font_with_east_asian_fallbacks("Batang", "Batang", "본문");
    assert_eq!(list.matches("\"Batang\"").count(), 1, "{list}");
}

#[test]
fn a_run_whose_family_cannot_write_its_script_reaches_a_face_that_can() {
    // A PowerPoint run may declare `<a:ea typeface="Calibri"/>` over Hangul,
    // and a workbook a Simplified Chinese family over Korean text. Neither
    // declared family has the glyphs, and neither carries a chain that does
    // (issues #537, #543).
    let latin_over_hangul = font_with_fallbacks_for_text("Calibri", "클라우드 변환");
    assert!(
        latin_over_hangul.contains("\"Malgun Gothic\""),
        "Hangul reaches a Korean face: {latin_over_hangul}"
    );
    assert!(
        latin_over_hangul.starts_with("(\"Calibri\""),
        "the declared family still leads, so Latin keeps it: {latin_over_hangul}"
    );

    let chinese_family_over_hangul = font_with_fallbacks_for_text("Noto Sans CJK SC", "구현 완료");
    let korean = chinese_family_over_hangul
        .find("\"Malgun Gothic\"")
        .expect("a Korean face is offered");
    for chinese in ["\"Microsoft YaHei\"", "\"PingFang SC\"", "\"SimSun\""] {
        if let Some(position) = chinese_family_over_hangul.find(chinese) {
            assert!(
                korean < position,
                "the script outranks the family's substitute chain: {chinese_family_over_hangul}"
            );
        }
    }
}

#[test]
fn painted_family_tracks_the_first_emitted_face_that_covers_the_text() {
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Calibri", "Noto Sans CJK SC", "Malgun Gothic"],
        &[],
        &[],
    )
    .with_italic_and_scripts(
        &[],
        &[
            ("Calibri", &[TextScript::Latin]),
            (
                "Noto Sans CJK SC",
                &[TextScript::Latin, TextScript::Chinese],
            ),
            ("Malgun Gothic", &[TextScript::Latin, TextScript::Korean]),
        ],
    );

    with_font_search_context(Some(&context), || {
        assert_eq!(
            painted_family_for_text("Noto Sans CJK SC", None, "금액"),
            "Malgun Gothic"
        );
        assert_eq!(
            painted_family_for_text("Noto Sans CJK SC", None, "Amount"),
            "Noto Sans CJK SC"
        );
        assert_eq!(
            painted_family_for_text("Calibri", Some("Noto Sans CJK SC"), "금액"),
            "Malgun Gothic"
        );
    });
}

#[test]
fn a_run_whose_family_can_write_its_script_keeps_it() {
    // The script chain must not preempt a family that is already right.
    let list = font_with_fallbacks_for_text("Batang", "구현 완료");
    assert!(list.starts_with("(\"Batang\""), "{list}");
}

#[test]
fn latin_only_text_is_offered_no_east_asian_face() {
    let list = font_with_fallbacks_for_text("Calibri", "Introduction");
    assert!(!list.contains("Malgun Gothic"), "{list}");
    assert!(!list.contains("Yu Gothic"), "{list}");
    assert!(!list.contains("Microsoft YaHei"), "{list}");
}

#[test]
fn kana_and_han_pick_their_own_scripts() {
    let kana = font_with_fallbacks_for_text("Calibri", "テキスト");
    assert!(kana.contains("\"Yu Gothic\""), "{kana}");
    // Han alone is ambiguous between the three, and is only decisive when no
    // script-specific character appears.
    let han = font_with_fallbacks_for_text("Calibri", "文書");
    assert!(han.contains("\"Microsoft YaHei\""), "{han}");
    // Korean text that also carries Han stays Korean.
    let mixed = font_with_fallbacks_for_text("Calibri", "文書 변환");
    assert!(mixed.contains("\"Malgun Gothic\""), "{mixed}");
}

/// A symbol the declared CJK family does not carry must reach a Latin face
/// before a CJK one.
///
/// U+25E6 WHITE BULLET is absent from Malgun Gothic. Word resolves it to
/// ArialMT and draws a 0.3545em ring; we resolved it through the family's
/// Korean substitute chain to a CJK face, whose U+25E6 is a full-width 1.0em
/// glyph — nearly three times the advance, and legible as a lowercase "o"
/// rather than a bullet (issue #642).
#[test]
fn test_symbol_missing_from_a_korean_family_falls_back_to_a_latin_face() {
    let result = font_with_fallbacks_for_text("Malgun Gothic", "\u{25E6}");

    let latin = ["Arial", "Liberation Sans", "Helvetica", "Arimo"]
        .iter()
        .filter_map(|face| result.find(face))
        .min()
        .unwrap_or_else(|| panic!("no Latin face in the chain for U+25E6: {result}"));
    let cjk = ["Apple SD Gothic Neo", "Noto Sans CJK", "Arial Unicode MS"]
        .iter()
        .filter_map(|face| result.find(face))
        .min();

    if let Some(cjk) = cjk {
        assert!(
            latin < cjk,
            "a Latin face must precede any CJK face for U+25E6: {result}"
        );
    }
}

/// The Latin detour must not displace CJK text's own faces.
#[test]
fn test_korean_text_still_reaches_korean_faces_first() {
    let result = font_with_fallbacks_for_text("Malgun Gothic", "가나");
    let korean = result
        .find("Malgun Gothic")
        .unwrap_or_else(|| panic!("Korean text must keep its declared family: {result}"));
    if let Some(arial) = result.find("Arial\"") {
        assert!(
            korean < arial,
            "Korean text must not be routed through a Latin face first: {result}"
        );
    }
}

/// The eastAsia path needs the same detour.
///
/// A DOCX run that names both `w:ascii` and `w:eastAsia` goes through
/// [`font_with_east_asian_fallbacks`] instead, and a marker missing from both
/// declared families fell through the East Asian family's substitutes to a
/// full-width CJK glyph exactly as the single-family path did (issue #642).
#[test]
fn test_symbol_missing_from_both_declared_families_falls_back_to_a_latin_face() {
    let result = font_with_east_asian_fallbacks("Malgun Gothic", "Malgun Gothic", "\u{25E6}");

    let latin = ["Arial", "Liberation Sans", "Helvetica", "Arimo"]
        .iter()
        .filter_map(|face| result.find(face))
        .min()
        .unwrap_or_else(|| panic!("no Latin face in the chain for U+25E6: {result}"));
    let cjk = ["Apple SD Gothic Neo", "Noto Sans CJK", "Arial Unicode MS"]
        .iter()
        .filter_map(|face| result.find(face))
        .min();

    if let Some(cjk) = cjk {
        assert!(
            latin < cjk,
            "a Latin face must precede any CJK face for U+25E6: {result}"
        );
    }
}

/// CJK fallback keeps the requested face's serif/sans class.
///
/// PowerPoint substitutes a serif Hangul face for a serif Latin request, to
/// keep the typographic voice. We reached Malgun Gothic — a geometric sans —
/// whatever was asked for, so 29 slide titles declaring `<a:ea
/// typeface="Cambria"/>` read in the wrong voice (issue #687).
#[test]
fn test_korean_under_a_serif_family_reaches_a_serif_hangul_face() {
    let result = font_with_fallbacks_for_text("Cambria", "가나다");

    let serif = ["Batang", "Noto Serif CJK KR", "Apple Myungjo", "Gungsuh"]
        .iter()
        .filter_map(|face| result.find(face))
        .min()
        .unwrap_or_else(|| panic!("no serif Hangul face for a serif request: {result}"));
    if let Some(sans) = ["Malgun Gothic", "Noto Sans CJK KR", "Apple SD Gothic Neo"]
        .iter()
        .filter_map(|face| result.find(face))
        .min()
    {
        assert!(
            serif < sans,
            "a serif request must reach a serif Hangul face first: {result}"
        );
    }
}

/// A sans request keeps the sans chain, so #537's behaviour is unchanged.
#[test]
fn test_korean_under_a_sans_family_still_reaches_the_sans_chain() {
    let result = font_with_fallbacks_for_text("Calibri", "가나다");
    let sans = result
        .find("Malgun Gothic")
        .unwrap_or_else(|| panic!("sans request must keep the sans chain: {result}"));
    if let Some(serif) = result.find("Noto Serif CJK KR") {
        assert!(
            sans < serif,
            "a sans request must not be routed to a serif Hangul face: {result}"
        );
    }
}

// ----- One face covering several scripts at once (issue #668) -----

#[test]
fn a_mixed_script_chain_substitutes_the_declared_face_before_the_script_faces() {
    // A Korean chart's Latin labels have to reach the stand-in for the face the
    // chart asked for. Ordering the script faces first sent `DOCX` to Malgun
    // Gothic, which covers Latin too, so the substitute was never consulted.
    let context = FontSearchContext::for_test(Vec::new(), &["Carlito", "Malgun Gothic"], &[], &[]);
    let list: String = with_font_search_context(Some(&context), || {
        font_for_mixed_script_text("Calibri", "DOCX 매출")
    });

    let carlito = list
        .find("Carlito")
        .expect("Calibri's substitute is listed");
    let malgun = list
        .find("Malgun Gothic")
        .expect("the Korean face is still listed");
    assert!(
        carlito < malgun,
        "the declared face's substitute must precede the script faces, got {list}"
    );
}

#[test]
fn a_mixed_script_chain_still_reaches_the_script_face() {
    // Dropping the script faces would leave the Hangul to the engine's own pick.
    let context = FontSearchContext::for_test(Vec::new(), &["Carlito", "Malgun Gothic"], &[], &[]);
    let list: String = with_font_search_context(Some(&context), || {
        font_for_mixed_script_text("Calibri", "DOCX 매출")
    });
    assert!(list.contains("Malgun Gothic"), "got {list}");
}

#[test]
fn test_document_requests_font_families_true_for_a_chart_only_document() {
    // A deck whose only font request comes from a chart's resolved theme face
    // still needs the font search context, or the compiler never sees the
    // directories that hold it and the chart falls back to the engine's
    // default anyway — the same trap shape labels fell into (issue #461).
    let mut chart = crate::ir::Chart {
        chart_type: crate::ir::ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string()],
        series: Vec::new(),
        grouping: crate::ir::ChartGrouping::default(),
        legend_position: crate::ir::LegendPosition::default(),
        has_legend: false,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: crate::ir::AxisTickMark::default(),
        value_axis_major_tick_mark: crate::ir::AxisTickMark::default(),
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: crate::ir::BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: crate::ir::ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: Some("Pretendard".to_string()),
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    let doc = Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Flow(crate::ir::FlowPage {
            first_header: None,
            first_footer: None,
            size: crate::ir::PageSize::default(),
            margins: crate::ir::Margins::default(),
            content: vec![Block::Chart(Box::new(chart.clone()))],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };
    assert!(
        document_requests_font_families(&doc),
        "a flowed chart's face must reach the font-context gate"
    );

    // The slide path is the one PPTX charts actually take.
    let doc = Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Fixed(crate::ir::FixedPage {
            size: crate::ir::PageSize {
                width: 720.0,
                height: 540.0,
            },
            elements: vec![crate::ir::FixedElement {
                x: 0.0,
                y: 0.0,
                width: 480.0,
                height: 320.0,
                kind: FixedElementKind::Chart(Box::new(chart.clone())),
            }],
            background_color: None,
            background_gradient: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };
    assert!(
        document_requests_font_families(&doc),
        "a slide chart's face must reach the font-context gate"
    );

    // An axis-only request must still trigger font discovery (#1563).
    chart.text_font_family = None;
    chart.category_axis_text_font_family = Some("Cambria".to_string());
    chart.value_axis_text_font_family = Some("Arial".to_string());
    let mut faces = Vec::new();
    assert!(visit_chart_fonts(&chart, &mut |family, _| {
        faces.push(family.to_string());
        true
    }));
    assert_eq!(faces, ["Cambria", "Arial"]);
    assert!(!visit_chart_fonts(&chart, &mut |family, _| family != "Arial"));

    // A secondary axis' own face is a request like the primary's (#1374).
    chart.secondary_value_axis = Some(crate::ir::ChartSecondaryValueAxis {
        side: crate::ir::ValueAxisSide::Right,
        deleted: false,
        number_format: None,
        major_unit: None,
        min: None,
        max: None,
        major_tick_mark: crate::ir::AxisTickMark::None,
        line: crate::ir::ChartLine::Automatic,
        text_font_family: Some("Georgia".to_string()),
        text_style: crate::ir::ChartTextStyle::default(),
    });
    let mut faces = Vec::new();
    assert!(visit_chart_fonts(&chart, &mut |family, _| {
        faces.push(family.to_string());
        true
    }));
    assert_eq!(faces, ["Cambria", "Arial", "Georgia"]);
    chart.secondary_value_axis = None;

    // A chart that names nothing still asks for nothing.
    chart.category_axis_text_font_family = None;
    chart.value_axis_text_font_family = None;
    chart.text_font_family = None;
    let doc = Document {
        metadata: crate::ir::Metadata::default(),
        pages: vec![Page::Flow(crate::ir::FlowPage {
            first_header: None,
            first_footer: None,
            size: crate::ir::PageSize::default(),
            margins: crate::ir::Margins::default(),
            content: vec![Block::Chart(Box::new(chart))],
            header: None,
            footer: None,
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })],
        styles: crate::ir::StyleSheet::default(),
    };
    assert!(!document_requests_font_families(&doc));
}

// --- sans-serif class fallback (issue #848) ---

#[test]
fn named_sans_serif_families_get_a_sans_serif_fallback_chain() {
    // The families the templates on issue #841 actually ask for, plus one
    // grotesque, none of which the explicit table names.
    for family in [
        "Microsoft Sans Serif",
        "Franklin Gothic Demi",
        "Century Gothic",
        "Akzidenz Grotesk",
    ] {
        let substitutes = substitutes(family)
            .unwrap_or_else(|| panic!("{family} should have class-preserving substitutes"));
        assert_eq!(
            substitutes,
            &["Liberation Sans", "Arimo", "DejaVu Sans", "Helvetica"],
            "{family} must not fall through to a proportional serif"
        );
    }
}

/// `mono` wins over `sans` when a name carries both, or every mono face whose
/// name says "Sans Mono" would come back proportional.
#[test]
fn a_sans_mono_family_stays_monospace() {
    let substitutes = substitutes("Noto Sans Mono CJK KR").expect("class-preserving substitutes");
    assert_eq!(
        substitutes,
        &[
            "DejaVu Sans Mono",
            "Noto Sans Mono",
            "Liberation Mono",
            "Cousine",
        ]
    );
}

/// A serif family carries no sans token and must keep falling through to the
/// document default rather than being pushed onto a sans chain.
#[test]
fn a_serif_family_is_not_misclassified_as_sans_serif() {
    for family in ["Bookman Old Style", "Baskerville", "Palatino Linotype"] {
        assert_eq!(substitutes(family), None, "{family} is not sans-serif");
    }
}

/// Corbel is a humanist sans with no class token in its name, so the table
/// has to name it — the Gantt template on #841 sets it as its major face.
#[test]
fn corbel_resolves_to_a_sans_serif_face() {
    let subs = substitutes("Corbel").expect("Corbel should have substitutes");
    assert!(
        subs.contains(&"Liberation Sans"),
        "Corbel must resolve to a sans face, got {subs:?}"
    );
}

/// Aptos has been Microsoft 365's default face since 2024, so it turns up in
/// every document a current Office build creates. It is a sans with no class
/// token in its name, and an XLSX header/footer names it as a bare `&"Aptos"`
/// with no declared class to fall back on, so the table has to name it — the
/// Gantt template on #841 sets its footer in it (issue #949).
#[test]
fn the_aptos_family_resolves_to_a_sans_serif_face() {
    assert_eq!(
        substitutes("Aptos"),
        Some(&["Noto Sans"][..]),
        "base Aptos must use the pinned LibreOffice-compatible face"
    );
    for family in ["Aptos Display", "Aptos Narrow", "Aptos SemiBold"] {
        let subs = substitutes(family).unwrap_or_else(|| panic!("{family} should substitute"));
        assert!(
            subs.contains(&"Liberation Sans"),
            "{family} must resolve to a sans face, got {subs:?}"
        );
    }
}

/// `Aptos Mono` is the family's fixed-pitch member, and the monospace token
/// already outranks the sans classification (issue #949).
#[test]
fn aptos_mono_stays_fixed_pitch() {
    let subs = substitutes("Aptos Mono").expect("Aptos Mono should substitute");
    assert_eq!(subs, MONOSPACE_SUBSTITUTES);
}

#[test]
fn a_declared_sans_family_stops_falling_back_to_a_serif() {
    // `Posterama` carries no `sans` token, so before the declaration was read
    // it fell through every name heuristic to the serif default (issue #891).
    assert_eq!(substitutes("Posterama"), None);
    let courier_before = substitutes("Courier New");

    let previous = set_declared_font_classes(HashMap::from([(
        "posterama".to_string(),
        DeclaredFontClass::SansSerif,
    )]));
    let declared = substitutes("Posterama").expect("a declared class must yield substitutes");
    assert_eq!(declared, SANS_SERIF_SUBSTITUTES);
    // The lookup normalizes, so the spelling at the call site does not matter.
    assert_eq!(substitutes("  POSTERAMA "), Some(SANS_SERIF_SUBSTITUTES));

    // A family the map says nothing about keeps whatever it resolved to.
    assert_eq!(substitutes("Courier New"), courier_before);

    set_declared_font_classes(previous);
    assert_eq!(substitutes("Posterama"), None);
}

/// The brand list matches the family's first token, so an unrelated family
/// that merely ends in one is left alone (issue #949).
#[test]
fn a_brand_token_only_counts_at_the_start_of_the_family() {
    assert_eq!(substitutes("Old Aptos"), None);
    assert_eq!(substitutes("Aptosia"), None);
}

// --- Style-suffixed family names (issue #1286) ---

use crate::render::font_context::test_faces::noto_serif_at_weight;

fn in_memory_context(fonts: &[typst::text::Font]) -> FontSearchContext {
    crate::render::font_context::resolve_font_search_context_from_fonts(fonts)
}

#[test]
fn typographic_family_trims_the_suffixes_typst_trims() {
    // Typst indexes every face under its name-table family with style
    // suffixes trimmed, so a request has to be trimmed the same way before it
    // can be compared with what the book holds.
    assert_eq!(typographic_family("Calibri Light"), "Calibri");
    assert_eq!(typographic_family("Segoe UI Semibold"), "Segoe UI");
    assert_eq!(typographic_family("Pretendard ExtraBold"), "Pretendard");
    assert_eq!(typographic_family("Pretendard Extra Bold"), "Pretendard");
    assert_eq!(
        typographic_family("Franklin Gothic Demi"),
        "Franklin Gothic"
    );
    assert_eq!(typographic_family("Arial Narrow"), "Arial");
    assert_eq!(typographic_family("Arial Black"), "Arial");
    // No suffix, and a style word that is not at the end, stay whole.
    assert_eq!(typographic_family("Calibri"), "Calibri");
    assert_eq!(typographic_family("Malgun Gothic"), "Malgun Gothic");
    assert_eq!(typographic_family("Blackadder ITC"), "Blackadder ITC");
}

#[test]
fn typographic_family_trims_variable_font_suffixes_as_typst_0_15_does() {
    // typst 0.15 files a variable face under its family with a `Variable`,
    // `Var` or `VF` suffix trimmed: `Pretendard Variable` is `Pretendard`.
    assert_eq!(typographic_family("Pretendard Variable"), "Pretendard");
    assert_eq!(typographic_family("Inter Var"), "Inter");
    assert_eq!(typographic_family("Roboto Flex VF"), "Roboto Flex");
    assert_eq!(typographic_family("Noto Sans KR Variable"), "Noto Sans KR");
}

#[test]
fn a_variable_font_request_paints_and_measures_through_its_trimmed_family() {
    // The book holds `Pretendard Variable` as `Pretendard` since typst 0.15,
    // and the suffix names that very face, so both chains must reach it.
    for purpose in [ChainPurpose::Paint, ChainPurpose::Metrics] {
        let candidates: Vec<String> = fallback_candidates("Pretendard Variable", None, purpose);
        assert_eq!(
            candidates.first().map(String::as_str),
            Some("Pretendard"),
            "{purpose:?}: {candidates:?}"
        );
    }
    assert_eq!(variable_font_base_family("Inter Var"), Some("Inter"));
    // A stretch suffix is not a variable-font one, and a name that merely ends
    // in the letters has no separator before them.
    assert_eq!(variable_font_base_family("Arial Narrow"), None);
    assert_eq!(variable_font_base_family("Kovar"), None);
}

#[test]
fn a_weight_suffixed_request_reaches_its_base_family_for_metrics_too() {
    // The book files `segoeuisb.ttf` under `Segoe UI` at 600 and never under
    // `Segoe UI Semibold`, so a chain keyed on the untrimmed name resolves no
    // face at all and every metric answers `None`. Native Word measures the
    // member the name denotes — twelve single-spaced 20pt `Arial Black`
    // paragraphs advance 28.189pt against Arial's 22.996pt, which is Arial
    // Black's own 1.4102em `hhea` sum and not Arial's 1.1499em — so the
    // metrics chain has to reach the base family too (issue #1643).
    for purpose in [ChainPurpose::Paint, ChainPurpose::Metrics] {
        for (requested, base_family) in [
            ("Segoe UI Semibold", "Segoe UI"),
            ("Arial Black", "Arial"),
            ("Calibri Light", "Calibri"),
        ] {
            let candidates: Vec<String> = fallback_candidates(requested, None, purpose);
            assert_eq!(
                candidates.first().map(String::as_str),
                Some(base_family),
                "{purpose:?} chain for {requested:?}: {candidates:?}"
            );
        }
    }
}

#[test]
fn a_family_name_stating_no_weight_keeps_its_base_family_out_of_both_chains() {
    // A stretch suffix names a family of its own — `Arial Narrow` is not a
    // member of Arial — so neither chain may fall back on the base family.
    for purpose in [ChainPurpose::Paint, ChainPurpose::Metrics] {
        let candidates: Vec<String> = fallback_candidates("Arial Narrow", None, purpose);
        assert!(
            !candidates.iter().any(|candidate| candidate == "Arial"),
            "{purpose:?} chain for Arial Narrow must not reach Arial: {candidates:?}"
        );
    }
}

#[test]
fn a_weight_suffix_states_the_weight_of_the_member_it_names() {
    use typst::text::FontWeight;
    assert_eq!(
        weight_stated_by_family_name("Calibri Light"),
        Some(FontWeight::LIGHT)
    );
    assert_eq!(
        weight_stated_by_family_name("Segoe UI Semibold"),
        Some(FontWeight::SEMIBOLD)
    );
    assert_eq!(
        weight_stated_by_family_name("Pretendard Extra Bold"),
        Some(FontWeight::EXTRABOLD)
    );
    assert_eq!(
        weight_stated_by_family_name("Franklin Gothic Medium"),
        Some(FontWeight::MEDIUM)
    );
    assert_eq!(
        weight_stated_by_family_name("Arial Black"),
        Some(FontWeight::BLACK)
    );
    assert_eq!(
        weight_stated_by_family_name("Segoe UI Bold"),
        Some(FontWeight::BOLD)
    );
}

#[test]
fn a_family_name_stating_no_weight_suffix_states_no_weight() {
    // A stretch suffix is not a weight, and a weight word that is not the
    // trailing suffix names a family, not a member of one.
    assert_eq!(weight_stated_by_family_name("Arial Narrow"), None);
    assert_eq!(weight_stated_by_family_name("Calibri"), None);
    assert_eq!(weight_stated_by_family_name("Blackadder ITC"), None);
    assert_eq!(weight_stated_by_family_name("Lightning Sans"), None);
}

#[test]
fn a_suffixed_family_is_available_where_the_base_family_ships_that_weight() {
    // The book indexes the rewritten face as `Noto Serif` at 300; a request
    // for `Noto Serif Light` is that face.
    let context = in_memory_context(&[noto_serif_at_weight(300)]);
    let available = with_font_search_context(Some(&context), || {
        is_primary_font_available("Noto Serif Light")
    });
    assert!(
        available,
        "the light member is indexed under the base family"
    );
}

#[test]
fn a_suffixed_family_is_unavailable_where_the_base_family_lacks_that_weight() {
    // Only the regular member is indexed, so nothing answers to the light
    // request and the stated weight must stay unclaimed.
    let context = in_memory_context(&[noto_serif_at_weight(400)]);
    let available = with_font_search_context(Some(&context), || {
        is_primary_font_available("Noto Serif Light")
    });
    assert!(!available, "no face at weight 300 is indexed");

    let other_family = in_memory_context(&[noto_serif_at_weight(300)]);
    let available = with_font_search_context(Some(&other_family), || {
        is_primary_font_available("Noto Sans Light")
    });
    assert!(!available, "the base family itself is not indexed");
}

#[test]
fn a_suffixed_request_paints_through_its_base_family() {
    // Typst never finds a family called `Noto Serif Light`, so the stated
    // weight can only land on the member it names if the base family follows
    // the request in the paint chain.
    let context = in_memory_context(&[noto_serif_at_weight(300)]);
    let chain = with_font_search_context(Some(&context), || {
        font_with_fallbacks_for_text("Noto Serif Light", "Heading")
    });
    assert!(
        chain.starts_with("(\"Noto Serif Light\", \"Noto Serif\""),
        "the base family must follow the suffixed request, got {chain}"
    );
}
