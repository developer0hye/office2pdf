#![cfg(not(target_arch = "wasm32"))] // native-only unit tests (filesystem, system fonts)
use super::*;

#[test]
fn typst_cache_state_evicts_after_the_bounded_document_interval() {
    let mut state = TypstCacheState {
        active_compilations: 0,
        completed_since_eviction: 0,
    };

    for _ in 0..TYPST_CACHE_EVICTION_INTERVAL {
        assert!(!state.begin_compilation());
        state.finish_compilation();
    }

    assert!(state.begin_compilation());
    assert_eq!(state.completed_since_eviction, 0);
    state.finish_compilation();
}

#[test]
fn typst_cache_state_defers_eviction_while_compilations_overlap() {
    let mut state = TypstCacheState {
        active_compilations: 1,
        completed_since_eviction: TYPST_CACHE_EVICTION_INTERVAL,
    };

    assert!(!state.begin_compilation());
    state.finish_compilation();
    state.finish_compilation();
    assert!(state.begin_compilation());
    state.finish_compilation();
}
use crate::test_support::make_test_svg;

/// Index into `data` of the face the compiler would shape `family` with over
/// an explicit font set, walking the same alias and substitute chain as
/// `best_face` but without a conversion context's in-memory faces.
fn best_face_index(data: &CachedFontData, family: &str) -> Option<usize> {
    let variant: typst::text::FontVariant = chain_variant(family, None);
    crate::render::font_subst::family_candidates(family)
        .iter()
        .find_map(|candidate| select_face_index(data, candidate, variant))
}

/// `line_metric_face` over an explicit font set: `office_font_dirs` hold the
/// application-bundled faces and `system_font_dirs` the faces that may shadow
/// them.
fn line_metric_face_in(
    data: &CachedFontData,
    office_font_dirs: &[PathBuf],
    system_font_dirs: &[PathBuf],
    family: &str,
) -> Option<typst::text::Font> {
    let shaped_index: usize = best_face_index(data, family)?;
    line_metric_face_at(data, office_font_dirs, system_font_dirs, shaped_index)
}

#[test]
fn test_default_font_search_paths_match_the_resolved_context() {
    // Every measurement site substitutes the memoized paths for a resolved
    // context's search paths, so the two must stay the same font set or the
    // metrics are taken against faces the compiler never sees.
    let resolved = crate::render::font_context::resolve_font_search_context(&[]);
    assert_eq!(
        crate::render::font_context::default_font_search_paths(),
        resolved.search_paths()
    );
}

#[test]
fn test_fallback_glyph_advances_do_not_resolve_a_font_context() {
    // Each resolution of a full font search context parses every face on
    // the host, and the fallback metric runs once per measured run: a
    // 35-page report spent 111 s of a 112 s conversion re-indexing fonts.
    // Outside an active context the lookup needs only the default paths.
    let families: Vec<String> = vec!["Calibri".to_string()];
    crate::render::font_context::FONT_CONTEXT_RESOLUTIONS.with(|count| count.set(0));
    let _ = glyph_advances_em_with_typst_fallback(&families, false, "overlong-token");
    let resolutions: usize =
        crate::render::font_context::FONT_CONTEXT_RESOLUTIONS.with(|count| count.get());
    assert_eq!(
        resolutions, 0,
        "a default-path lookup must not re-index the host's fonts"
    );
}

#[test]
fn test_compile_simple_text() {
    let result = compile_to_pdf("Hello, World!", &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty(), "PDF bytes should not be empty");
    assert!(
        result.starts_with(b"%PDF"),
        "PDF should start with %PDF magic bytes"
    );
}

#[test]
fn test_compile_with_page_setup() {
    let source = r#"#set page(width: 612pt, height: 792pt)
Hello from a US Letter page."#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_styled_text() {
    let source = r#"#text(weight: "bold", size: 16pt)[Bold Title]

#text(style: "italic")[Italic body text]

#underline[Underlined text]"#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_colored_text() {
    let source = r#"#text(fill: rgb(255, 0, 0))[Red text]
#text(fill: rgb(0, 128, 255))[Blue text]"#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_alignment() {
    let source = r#"#align(center)[Centered text]

#align(right)[Right-aligned text]"#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_invalid_source_returns_error() {
    // Invalid Typst source should produce a compilation error
    let result = compile_to_pdf(
        "#invalid-func-that-does-not-exist()",
        &[],
        None,
        &[],
        false,
        false,
    );
    assert!(result.is_err(), "Invalid source should produce an error");
}

#[test]
fn test_compile_empty_source() {
    // Empty source should still produce valid PDF (empty page)
    let result = compile_to_pdf("", &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_multiple_paragraphs() {
    let source = "First paragraph.\n\nSecond paragraph.\n\nThird paragraph.";
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

/// Compute CRC32 over PNG chunk type + data.
fn png_crc32(chunk_type: &[u8], data: &[u8]) -> u32 {
    let mut crc: u32 = 0xFFFF_FFFF;
    for &byte in chunk_type.iter().chain(data.iter()) {
        crc ^= byte as u32;
        for _ in 0..8 {
            if crc & 1 != 0 {
                crc = (crc >> 1) ^ 0xEDB8_8320;
            } else {
                crc >>= 1;
            }
        }
    }
    crc ^ 0xFFFF_FFFF
}

/// Build a minimal valid 1x1 red PNG with correct CRC checksums.
fn make_test_png() -> Vec<u8> {
    let mut png = Vec::new();
    // PNG signature
    png.extend_from_slice(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

    // IHDR: 1x1, 8-bit RGB
    let ihdr_data: [u8; 13] = [
        0x00, 0x00, 0x00, 0x01, // width=1
        0x00, 0x00, 0x00, 0x01, // height=1
        0x08, // bit depth=8
        0x02, // color type=RGB
        0x00, 0x00, 0x00, // compression, filter, interlace
    ];
    let ihdr_type = b"IHDR";
    png.extend_from_slice(&(ihdr_data.len() as u32).to_be_bytes());
    png.extend_from_slice(ihdr_type);
    png.extend_from_slice(&ihdr_data);
    png.extend_from_slice(&png_crc32(ihdr_type, &ihdr_data).to_be_bytes());

    // IDAT: zlib-compressed row [filter=0, R=255, G=0, B=0]
    let idat_data: [u8; 15] = [
        0x78, 0x01, // zlib header
        0x01, // BFINAL=1, BTYPE=00 (stored)
        0x04, 0x00, 0xFB, 0xFF, // LEN=4, NLEN
        0x00, 0xFF, 0x00, 0x00, // filter + RGB
        0x03, 0x01, 0x01, 0x00, // adler32
    ];
    let idat_type = b"IDAT";
    png.extend_from_slice(&(idat_data.len() as u32).to_be_bytes());
    png.extend_from_slice(idat_type);
    png.extend_from_slice(&idat_data);
    png.extend_from_slice(&png_crc32(idat_type, &idat_data).to_be_bytes());

    // IEND
    let iend_type = b"IEND";
    png.extend_from_slice(&0u32.to_be_bytes());
    png.extend_from_slice(iend_type);
    png.extend_from_slice(&png_crc32(iend_type, &[]).to_be_bytes());

    png
}

#[cfg(feature = "embedded-fonts")]
#[test]
fn test_embedded_fonts_are_available() {
    // MinimalWorld should always have embedded fallback fonts available
    // (Libertinus Serif, New Computer Modern, DejaVu Sans Mono)
    let world = MinimalWorld::new("", &[], &[]);
    assert!(
        world.font_source.len() > 0,
        "MinimalWorld should have at least the embedded fallback fonts"
    );
}

#[test]
fn test_system_fonts_enabled() {
    // System discovery adds to the feature-selected embedded set, which
    // can be empty on native builds without default features.
    let world = MinimalWorld::new("", &[], &[]);
    let embedded_only_count = { embedded_fonts().1.len() };
    assert!(
        world.font_source.len() >= embedded_only_count,
        "System font discovery should not reduce available fonts: total {} vs embedded-only {}",
        world.font_source.len(),
        embedded_only_count
    );
}

#[test]
fn test_compile_with_system_font_name() {
    // A document specifying a common system font should compile successfully.
    // If Arial is unavailable, another installed font or the optional
    // Typst embedded set supplies the fallback.
    let source = r#"#set text(font: "Arial")
Hello with a system font."#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[cfg(feature = "embedded-fonts")]
#[test]
fn test_embedded_fonts_still_available_as_fallback() {
    // Embedded fonts (Libertinus Serif) must still be available even with
    // system font discovery enabled.
    let source = r#"#set text(font: "Libertinus Serif")
Text in Libertinus Serif."#;
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_pdfa2b_produces_valid_pdf() {
    let result = compile_to_pdf(
        "Hello PDF/A!",
        &[],
        Some(crate::config::PdfStandard::PdfA2b),
        &[],
        false,
        false,
    )
    .unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_pdfa2b_contains_xmp_metadata() {
    let result = compile_to_pdf(
        "PDF/A metadata test",
        &[],
        Some(crate::config::PdfStandard::PdfA2b),
        &[],
        false,
        false,
    )
    .unwrap();
    // PDF/A-2b requires XMP metadata with pdfaid namespace
    let pdf_str = String::from_utf8_lossy(&result);
    assert!(
        pdf_str.contains("pdfaid") || pdf_str.contains("PDF/A"),
        "PDF/A output should contain PDF/A identification metadata"
    );
}

#[test]
fn test_compile_default_no_pdfa_metadata() {
    let result = compile_to_pdf("Regular PDF", &[], None, &[], false, false).unwrap();
    let pdf_str = String::from_utf8_lossy(&result);
    // A regular PDF should not have pdfaid conformance metadata
    assert!(
        !pdf_str.contains("pdfaid:conformance"),
        "Regular PDF should not contain PDF/A conformance metadata"
    );
}

#[test]
fn test_compile_with_font_paths_empty() {
    // Empty font paths should work the same as without
    let result = compile_to_pdf("Hello!", &[], None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_with_caller_provided_in_memory_font() {
    let carrier = include_bytes!("../../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
    let embedded = crate::parser::embedded_fonts::extract_embedded_font_data(
        carrier,
        crate::config::Format::Docx,
    )
    .expect("the #943 fixture should carry an embedded font");
    let fonts = load_fonts_from_bytes(embedded.font_bytes());
    assert!(!fonts.is_empty(), "the fixture font should parse");

    let pdf = compile_to_pdf_with_fonts(
        r#"#text(font: ("No Such Font", "Noto Sans SC"))[Hello 中文测试文档]"#,
        &[],
        None,
        &[],
        &fonts,
        false,
        false,
    )
    .expect("the in-memory face should compile on native targets too");

    assert!(
        pdf.windows(b"NotoSansSC".len())
            .any(|window| window == b"NotoSansSC"),
        "the output PDF should embed the caller-provided face"
    );
}

#[test]
fn test_in_memory_last_resort_bypasses_process_metric_cache() {
    let carrier = include_bytes!("../../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
    let embedded = crate::parser::embedded_fonts::extract_embedded_font_data(
        carrier,
        crate::config::Format::Docx,
    )
    .expect("the #943 fixture should carry an embedded font");
    let fonts = load_fonts_from_bytes(embedded.font_bytes());
    let font = fonts.first().expect("the fixture font should parse");
    let instance = measured_instance(font);
    let ttf = instance.ttf();
    let upem = f64::from(ttf.units_per_em()).max(1.0);
    let expected_pitch =
        (f64::from(ttf.ascender()) - f64::from(ttf.descender()) + f64::from(ttf.line_gap())) / upem;

    // A family no host ships and no substitute table names: its chain is the
    // family itself and the conversion's last resort. (`SimSun` would not do:
    // its substitutes precede the last resort, and a host shipping one of
    // them — macOS ships Songti SC — paints and measures that face, which is
    // the compiler's order, not a cache defect; issue #1629.)
    let missing_family = "office2pdf issue 943 embedded last resort";
    // Populate the process cache from the machine's ordinary font set first.
    assert!(font_line_metrics_em(missing_family).is_none());

    let context = crate::render::font_context::resolve_font_search_context_from_fonts(&fonts)
        .with_last_resort_family(Some("Noto Sans SC"));
    let actual = crate::render::font_subst::with_font_search_context(Some(&context), || {
        font_line_metrics_em(missing_family)
    })
    .expect("the active in-memory last resort should provide metrics");

    assert!(
        (actual.2 - expected_pitch).abs() < 1e-12,
        "active metrics must come from Noto Sans SC even after the family was cached as missing: {actual:?}"
    );
}

#[test]
fn test_path_font_last_resort_bypasses_process_metric_caches() {
    let carrier = include_bytes!("../../../../tests/fixtures/docx/wasm_embedded_cjk.docx");
    let embedded_data = crate::parser::embedded_fonts::extract_embedded_font_data(
        carrier,
        crate::config::Format::Docx,
    )
    .expect("the #969 fixture should carry an embedded font");
    let fonts = load_fonts_from_bytes(embedded_data.font_bytes());
    let font = fonts.first().expect("the fixture font should parse");
    let instance = measured_instance(font);
    let ttf = instance.ttf();
    let upem = f64::from(ttf.units_per_em()).max(1.0);
    let expected = (
        (f64::from(ttf.ascender()) + f64::from(ttf.line_gap())) / upem,
        -f64::from(ttf.descender()) / upem,
        (f64::from(ttf.ascender()) - f64::from(ttf.descender()) + f64::from(ttf.line_gap())) / upem,
    );
    let expected_hhea_ascender = f64::from(ttf.tables().hhea.ascender) / upem;
    let expected_cap_height = measured_instance(font).metrics().cap_height.get();
    // PowerPoint measures a line box from OS/2's usWin pair (issue #1176).
    let (ascent, descent) = match ttf.tables().os2 {
        Some(os2) if os2.windows_ascender() > 0 => (
            f64::from(os2.windows_ascender()).abs() / upem,
            f64::from(os2.windows_descender()).abs() / upem,
        ),
        _ => (
            f64::from(ttf.tables().hhea.ascender).abs() / upem,
            f64::from(ttf.tables().hhea.descender).abs() / upem,
        ),
    };
    // What this test pins is which *face* answers, not which split rule: the
    // rule itself is covered by
    // `test_powerpoint_line_box_shares_every_font_on_the_line` and
    // `an_overflowing_face_shares_the_line_box_in_its_own_proportion`.
    let expected_powerpoint = powerpoint_line_box_split_em([(ascent, descent)])
        .expect("the fixture face declares an ascent");

    // Native document fonts are materialized into a conversion-local path.
    // Prime the family-only process cache with a miss before that path becomes
    // active: the conversion context must still be allowed to resolve its own
    // final fallback face.
    let missing_family = "office2pdf issue 969 cache miss";
    assert!(font_line_metrics_em(missing_family).is_none());
    assert!(font_hhea_ascender_em(missing_family).is_none());
    assert!(font_cap_height_em(missing_family).is_none());
    let cached_powerpoint = powerpoint_line_box_em(missing_family)
        .expect("the ordinary PowerPoint metric falls back to Typst's default face");

    let embedded_dir =
        crate::parser::embedded_fonts::extract_embedded_fonts(carrier, crate::config::Format::Docx)
            .expect("the #969 fixture font should materialize");
    let context = crate::render::font_context::resolve_font_search_context(&[embedded_dir
        .path()
        .to_path_buf()])
    .with_last_resort_family(Some("Noto Sans SC"));
    let (actual, hhea_ascender, cap_height, powerpoint) =
        crate::render::font_subst::with_font_search_context(Some(&context), || {
            (
                font_line_metrics_em(missing_family),
                font_hhea_ascender_em(missing_family),
                font_cap_height_em(missing_family),
                powerpoint_line_box_em(missing_family),
            )
        });
    let actual = actual.expect("the active path font should provide line metrics");
    let hhea_ascender = hhea_ascender.expect("the active path font should provide an ascender");
    let cap_height = cap_height.expect("the active path font should provide a cap height");
    let powerpoint = powerpoint.expect("the active path font should provide a PowerPoint split");

    assert!(
        (actual.0 - expected.0).abs() < 1e-12
            && (actual.1 - expected.1).abs() < 1e-12
            && (actual.2 - expected.2).abs() < 1e-12,
        "active metrics must come from the materialized Noto Sans SC face: {actual:?}"
    );
    assert!((hhea_ascender - expected_hhea_ascender).abs() < 1e-12);
    assert!((cap_height - expected_cap_height).abs() < 1e-12);
    assert!(
        (powerpoint.0 - expected_powerpoint.0).abs() < 1e-12
            && (powerpoint.1 - expected_powerpoint.1).abs() < 1e-12,
        "active PowerPoint metrics must come from Noto Sans SC: {powerpoint:?}"
    );
    assert_ne!(
        powerpoint, cached_powerpoint,
        "the materialized face should replace the cached default-font split"
    );
}

/// The line box, bare `hhea` ascender and line gap a face declares, in em.
#[cfg(not(target_arch = "wasm32"))]
fn declared_line_box_em(font: &typst::text::Font) -> (f64, f64, f64) {
    let instance = measured_instance(font);
    let ttf = instance.ttf();
    let upem = f64::from(ttf.units_per_em()).max(1.0);
    let pitch =
        (f64::from(ttf.ascender()) - f64::from(ttf.descender()) + f64::from(ttf.line_gap())) / upem;
    let top = (f64::from(ttf.ascender()) + f64::from(ttf.line_gap())) / upem;
    (top, pitch - top, pitch)
}

#[test]
fn a_weight_suffixed_family_measures_the_member_its_name_denotes() {
    // The book files a weight member under the base family with the suffix
    // trimmed, so `Segoe UI Semibold` and `Arial Black` match no family key and
    // every line metric answered `None` — leaving `word_line_height_settings`
    // with no fixed Word line box, the shape of #1197 (issue #1643).
    //
    // Reaching the base family is only half of it: the member the name denotes
    // is the one the document asked for, and a family's members need not share
    // vertical metrics. Native Word measures that member — twelve
    // single-spaced 20pt `Arial Black` paragraphs advance 28.189pt, which is
    // Arial Black's own 1.4102em `hhea` sum, not the 22.996pt (1.1499em) Arial
    // regular declares; a `w:b` run on plain Arial advances 22.996pt, so the
    // stated weight alone does not move the box, the named member does.
    //
    // The tracked Noto Serif ships one `hhea` for every weight, so the light
    // member here carries a rewritten ascender: the regular declares
    // 1069/-293/0 on 1000 upem and the light 1500/-293/0. Reading the regular's
    // box for a light request is then off by 0.431em, the same shape as reading
    // Arial's for Arial Black.
    use crate::render::font_context::test_faces::{
        noto_serif_at_weight, noto_serif_at_weight_with_ascender,
    };

    let regular: typst::text::Font = noto_serif_at_weight(400);
    let light: typst::text::Font = noto_serif_at_weight_with_ascender(300, 1500);
    let expected_regular = declared_line_box_em(&regular);
    let expected_light = declared_line_box_em(&light);
    assert_ne!(
        expected_regular.2, expected_light.2,
        "the rewritten members must declare different line boxes for the test to discriminate"
    );

    let context = crate::render::font_context::resolve_font_search_context_from_fonts(&[
        regular.clone(),
        light.clone(),
    ]);
    let (suffixed_line, base_line, suffixed_ascender, suffixed_gap) =
        crate::render::font_subst::with_font_search_context(Some(&context), || {
            (
                font_line_metrics_em("Noto Serif Light"),
                font_line_metrics_em("Noto Serif"),
                font_hhea_ascender_em("Noto Serif Light"),
                font_line_gap_em("Noto Serif Light"),
            )
        });

    let suffixed_line =
        suffixed_line.expect("a weight-suffixed request resolves through its base family");
    let base_line = base_line.expect("the base family resolves its own regular member");
    let suffixed_ascender = suffixed_ascender.expect("the suffixed request resolves an ascender");
    let suffixed_gap = suffixed_gap.expect("the suffixed request resolves a line gap");

    let close = |actual: (f64, f64, f64), expected: (f64, f64, f64)| {
        (actual.0 - expected.0).abs() < 1e-12
            && (actual.1 - expected.1).abs() < 1e-12
            && (actual.2 - expected.2).abs() < 1e-12
    };
    assert!(
        close(suffixed_line, expected_light),
        "`Noto Serif Light` must measure the 300 member's own line box: \
         {suffixed_line:?} against {expected_light:?}"
    );
    assert!(
        close(base_line, expected_regular),
        "`Noto Serif` must still measure the regular member: \
         {base_line:?} against {expected_regular:?}"
    );

    let light_instance = measured_instance(&light);
    let light_ttf = light_instance.ttf();
    let upem = f64::from(light_ttf.units_per_em()).max(1.0);
    assert!(
        (suffixed_ascender - f64::from(light_ttf.tables().hhea.ascender) / upem).abs() < 1e-12,
        "the bare ascender must come from the 300 member too: {suffixed_ascender}"
    );
    assert!(
        (suffixed_gap - f64::from(light_ttf.line_gap()) / upem).abs() < 1e-12,
        "the line gap must come from the 300 member too: {suffixed_gap}"
    );
}

#[test]
fn a_stretch_suffixed_family_does_not_borrow_its_base_family_metrics() {
    // `Arial Narrow` is its own family, not a member of Arial, and the trimmer
    // that files `Arial Black` under `Arial` files it there too. Only a weight
    // suffix names a member, so a stretch suffix must keep answering from its
    // own chain (issue #1643).
    use crate::render::font_context::test_faces::noto_serif_at_weight_with_ascender;

    let regular: typst::text::Font =
        crate::render::font_context::test_faces::noto_serif_at_weight(400);
    let bold: typst::text::Font = noto_serif_at_weight_with_ascender(700, 1500);
    let bold_line = declared_line_box_em(&bold);
    let context = crate::render::font_context::resolve_font_search_context_from_fonts(&[
        regular.clone(),
        bold,
    ]);
    let (narrow, suffixed) =
        crate::render::font_subst::with_font_search_context(Some(&context), || {
            (
                font_line_metrics_em("Noto Serif Condensed"),
                font_line_metrics_em("Noto Serif Bold"),
            )
        });

    // The weight-suffixed control proves the context discriminates at all: a
    // `Bold` request reaches the 700 member's taller box.
    let suffixed = suffixed.expect("a weight-suffixed request resolves through its base family");
    assert!(
        (suffixed.2 - bold_line.2).abs() < 1e-12,
        "`Noto Serif Bold` must measure the 700 member: {suffixed:?}"
    );

    // Nothing declares a condensed member, so the request falls through its own
    // chain — never onto a heavier member picked because the name happened to
    // end in a suffix the book trims.
    if let Some(narrow) = narrow {
        assert!(
            (narrow.2 - bold_line.2).abs() > 1e-9,
            "a stretch suffix must not select a weight member: {narrow:?}"
        );
    }
}

#[test]
fn an_exact_face_on_a_search_path_outranks_an_in_memory_chain_tail_for_metrics() {
    // Naming a face whose reproducible fallback is bundled loads that fallback
    // into the conversion's in-memory fonts, and the metric chain ends on it.
    // The paint chain still resolves the exact face first when a search path
    // holds it, so every metric must read that face too (issue #1629).
    let exact_face_bytes: &[u8] = include_bytes!("../../fonts/Selawik-Regular.ttf");
    let exact_face = typst::text::Font::new(typst::foundations::Bytes::new(exact_face_bytes), 0)
        .expect("the bundled Selawik face parses");
    let exact_family: &str = "Selawik";
    assert_eq!(exact_face.info().family, exact_family);

    let unique = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("system time should be valid")
        .as_nanos();
    let font_dir = std::env::temp_dir().join(format!("office2pdf-issue-1629-{unique}"));
    std::fs::create_dir_all(&font_dir).expect("the search path directory is created");
    std::fs::write(font_dir.join("Selawik-Regular.ttf"), exact_face_bytes)
        .expect("the face is written to the search path");

    let in_memory: &[typst::text::Font] = crate::bundled_fonts::noto_serif_fonts();
    let chain_tail = in_memory
        .iter()
        .find(|font| font.info().variant == typst::text::FontVariant::default())
        .expect("the bundled Noto Serif ships a regular face");
    let chain_tail_family: &str = crate::bundled_fonts::NOTO_SERIF_FAMILY;

    let context =
        crate::render::font_context::resolve_font_search_context(std::slice::from_ref(&font_dir))
            .with_in_memory_fonts(in_memory)
            .with_last_resort_family(Some(chain_tail_family));
    let (line, hhea_ascender, cap_height, powerpoint, tail_line) =
        crate::render::font_subst::with_font_search_context(Some(&context), || {
            (
                font_line_metrics_em(exact_family),
                font_hhea_ascender_em(exact_family),
                font_cap_height_em(exact_family),
                powerpoint_line_box_em(exact_family),
                font_line_metrics_em(chain_tail_family),
            )
        });
    let _ = std::fs::remove_dir_all(&font_dir);

    let hhea_line = |font: &typst::text::Font| {
        let instance = measured_instance(font);
        let ttf = instance.ttf();
        let upem = f64::from(ttf.units_per_em()).max(1.0);
        let pitch = (f64::from(ttf.ascender()) - f64::from(ttf.descender())
            + f64::from(ttf.line_gap()))
            / upem;
        let top = (f64::from(ttf.ascender()) + f64::from(ttf.line_gap())) / upem;
        (top, pitch - top, pitch)
    };
    let expected_line = hhea_line(&exact_face);
    let tail_expected_line = hhea_line(chain_tail);
    assert_ne!(
        expected_line, tail_expected_line,
        "the two faces must disagree for the test to discriminate"
    );
    let expected_hhea_ascender = {
        let instance = measured_instance(&exact_face);
        let ttf = instance.ttf();
        f64::from(ttf.tables().hhea.ascender) / f64::from(ttf.units_per_em()).max(1.0)
    };
    let expected_cap_height = measured_instance(&exact_face).metrics().cap_height.get();
    let expected_powerpoint =
        powerpoint_line_box_split_em([powerpoint_face_metrics_em(&exact_face)])
            .expect("the exact face declares an ascent");

    let line = line.expect("the exact face on the search path provides line metrics");
    assert!(
        (line.0 - expected_line.0).abs() < 1e-12
            && (line.1 - expected_line.1).abs() < 1e-12
            && (line.2 - expected_line.2).abs() < 1e-12,
        "line metrics must come from the exact face on the search path, not the in-memory chain tail: {line:?} vs {expected_line:?}"
    );
    let hhea_ascender = hhea_ascender.expect("the exact face provides an ascender");
    assert!(
        (hhea_ascender - expected_hhea_ascender).abs() < 1e-12,
        "hhea ascender must come from the exact face: {hhea_ascender} vs {expected_hhea_ascender}"
    );
    let cap_height = cap_height.expect("the exact face provides a cap height");
    assert!(
        (cap_height - expected_cap_height).abs() < 1e-12,
        "cap height must come from the exact face: {cap_height} vs {expected_cap_height}"
    );
    let powerpoint = powerpoint.expect("the exact face provides a PowerPoint split");
    assert!(
        (powerpoint.0 - expected_powerpoint.0).abs() < 1e-12
            && (powerpoint.1 - expected_powerpoint.1).abs() < 1e-12,
        "the PowerPoint line box must come from the exact face: {powerpoint:?} vs {expected_powerpoint:?}"
    );

    // The in-memory face still answers for the family it actually is.
    let tail_line = tail_line.expect("the in-memory chain tail provides its own metrics");
    assert!(
        (tail_line.2 - tail_expected_line.2).abs() < 1e-12,
        "a family held in memory keeps resolving to that face: {tail_line:?} vs {tail_expected_line:?}"
    );
}

#[test]
fn test_compile_with_nonexistent_font_path() {
    // Non-existent font path should not crash — font discovery skips invalid dirs
    let paths = vec![PathBuf::from("/nonexistent/font/path")];
    let result = compile_to_pdf("Hello!", &[], None, &paths, false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_with_embedded_image() {
    let png_data = make_test_png();
    let images = vec![ImageAsset {
        path: "img-0.png".to_string(),
        data: png_data,
    }];
    let source = r#"#image("img-0.png", width: 100pt)"#;
    let result = compile_to_pdf(source, &images, None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[test]
fn test_compile_with_embedded_svg_image() {
    let svg_data = make_test_svg();
    let images = vec![ImageAsset {
        path: "img-0.svg".to_string(),
        data: svg_data,
    }];
    let source = r#"#image("img-0.svg", width: 100pt)"#;
    let result = compile_to_pdf(source, &images, None, &[], false, false).unwrap();
    assert!(!result.is_empty());
    assert!(result.starts_with(b"%PDF"));
}

#[cfg(feature = "embedded-fonts")]
#[test]
fn test_embedded_only_world_produces_valid_pdf() {
    // Simulates the WASM code path: embedded fonts only, no system fonts.
    // This verifies that the embedded-only MinimalWorld can produce valid PDFs.
    let world = MinimalWorld::new_embedded_only("Hello from embedded-only world!", &[]);
    assert!(
        world.font_source.len() > 0,
        "Embedded-only world should have fonts"
    );

    let warned = typst::compile::<PagedDocument>(&world);
    let document = warned.output.expect("Compilation should succeed");
    let pdf = typst_pdf::pdf(&document, &typst_pdf::PdfOptions::default())
        .expect("PDF export should succeed");
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_embedded_only_world_has_fonts() {
    // The native simulation follows the Cargo feature; real WASM builds
    // always retain their fallback set.
    let world = MinimalWorld::new_embedded_only("", &[]);
    let embedded_count = { embedded_fonts().1.len() };
    assert_eq!(
        world.font_source.len(),
        embedded_count,
        "Embedded-only world should have exactly the embedded fonts"
    );
}

#[test]
fn test_pdfa_timestamp_is_not_hardcoded() {
    // PDF/A output should contain the actual conversion timestamp,
    // not the previously hardcoded 2024-01-01.
    let result = compile_to_pdf(
        "Timestamp test",
        &[],
        Some(crate::config::PdfStandard::PdfA2b),
        &[],
        false,
        false,
    )
    .unwrap();
    let pdf_str = String::from_utf8_lossy(&result);
    // The old hardcoded date was 2024-01-01T00:00:00 — it should no longer appear
    assert!(
        !pdf_str.contains("2024-01-01T00:00:00"),
        "PDF/A timestamp should not be the hardcoded 2024-01-01T00:00:00"
    );
}

#[test]
fn test_current_utc_datetime_is_valid() {
    // The helper should produce a valid Datetime that can create a Timestamp.
    let dt = current_utc_datetime();
    let _ts = typst_pdf::Timestamp::new_utc(dt);
}

#[test]
fn test_pdfa_timestamp_has_recent_date() {
    // The PDF/A XMP metadata should contain a date from the current
    // decade, not a hardcoded past date.
    let result = compile_to_pdf(
        "Year test",
        &[],
        Some(crate::config::PdfStandard::PdfA2b),
        &[],
        false,
        false,
    )
    .unwrap();
    let pdf_str = String::from_utf8_lossy(&result);
    // The XMP metadata should contain a CreateDate field
    assert!(
        pdf_str.contains("xmp:CreateDate") || pdf_str.contains("CreateDate"),
        "PDF/A should contain creation date metadata"
    );
    // The date should NOT be the hardcoded 2024-01-01
    assert!(
        !pdf_str.contains("2024-01-01"),
        "PDF/A timestamp should not contain hardcoded 2024-01-01"
    );
}

// --- PDF output size optimization tests (US-089) ---

#[test]
fn test_pdf_uses_flate_compression() {
    // typst-pdf (via krilla) compresses content streams with FLATE by default.
    // Verify that the output PDF contains FlateDecode filter references.
    let source = "Hello, compressed world! ".repeat(100);
    let result = compile_to_pdf(&source, &[], None, &[], false, false).unwrap();
    let pdf_str = String::from_utf8_lossy(&result);
    assert!(
        pdf_str.contains("FlateDecode"),
        "PDF content streams should use FlateDecode compression"
    );
}

#[test]
fn test_font_subsetting_reduces_size() {
    // A PDF using only a few glyphs should be significantly smaller than
    // one using many distinct glyphs, demonstrating font subsetting is active.
    // "Few glyphs" document: only ASCII letters a-z
    let few_glyphs = compile_to_pdf("abcdefghij", &[], None, &[], false, false).unwrap();

    // "Many glyphs" document: diverse characters force more glyph data.
    // Avoid Typst special characters (#, $, *, _, etc.) to keep it valid markup.
    let many_glyphs_source = "abcdefghijklmnopqrstuvwxyz \
        ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789 \
        The quick brown fox jumps over the lazy dog. \
        SPHINX OF BLACK QUARTZ, JUDGE MY VOW. \
        Pack my box with five dozen liquor jugs. \
        How vexingly quick daft zebras jump.";
    let many_glyphs = compile_to_pdf(many_glyphs_source, &[], None, &[], false, false).unwrap();

    // With font subsetting, the "few glyphs" PDF should be noticeably smaller.
    // Without subsetting, both would embed the full font and be similar in size.
    assert!(
        few_glyphs.len() < many_glyphs.len(),
        "PDF with fewer glyphs ({} bytes) should be smaller than PDF with many glyphs ({} bytes), \
         indicating font subsetting is active",
        few_glyphs.len(),
        many_glyphs.len()
    );
}

#[test]
fn test_multipage_text_pdf_size_reasonable() {
    // A 10-page text-only document should produce a PDF well under 500KB.
    // This verifies that compression and font subsetting keep output compact.
    //
    // typst-pdf behavior (verified):
    // - Content streams use FLATE compression (compress_content_streams: true)
    // - Fonts are automatically subset to include only used glyphs
    // - No unnecessary re-encoding of embedded data
    let mut source = String::new();
    for i in 1..=10 {
        if i > 1 {
            source.push_str("#pagebreak()\n");
        }
        source.push_str(&format!(
            "= Page {i}\n\n\
             This is page {i} of a multi-page document used to verify \
             that PDF output size remains reasonable with compression \
             and font subsetting enabled.\n\n\
             Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
             Sed do eiusmod tempor incididunt ut labore et dolore magna \
             aliqua. Ut enim ad minim veniam, quis nostrud exercitation \
             ullamco laboris nisi ut aliquip ex ea commodo consequat.\n\n"
        ));
    }
    let result = compile_to_pdf(&source, &[], None, &[], false, false).unwrap();

    // 500KB = 512_000 bytes — generous upper bound for 10 pages of text
    assert!(
        result.len() < 512_000,
        "10-page text-only PDF should be under 500KB, actual size: {} bytes ({:.1} KB)",
        result.len(),
        result.len() as f64 / 1024.0
    );
}

#[test]
fn test_pdf_with_image_size_proportional() {
    // A PDF with an embedded image should not inflate the image size
    // significantly. The output PDF should be proportional to the input
    // image data size (not orders of magnitude larger from re-encoding).
    let png_data = make_test_png();
    let png_size = png_data.len();
    let images = vec![ImageAsset {
        path: "img-0.png".to_string(),
        data: png_data,
    }];
    let source = r#"#image("img-0.png", width: 100pt)"#;
    let result = compile_to_pdf(source, &images, None, &[], false, false).unwrap();

    // The PDF has overhead (fonts, structure, metadata) beyond the image.
    // But the total should not be unreasonably large for a tiny 1x1 image.
    // A 1x1 PNG is ~70 bytes; the PDF overhead is typically 10-30KB (fonts).
    // We assert the total is under 100KB to catch re-encoding issues.
    assert!(
        result.len() < 100_000,
        "PDF with tiny 1x1 image should be under 100KB, actual: {} bytes ({:.1} KB). \
         Image was {} bytes. Possible image re-encoding issue.",
        result.len(),
        result.len() as f64 / 1024.0,
        png_size
    );
}

#[test]
fn test_empty_page_pdf_baseline_size() {
    // An empty page PDF establishes the baseline overhead (fonts, structure).
    // This helps verify that additional content adds proportional size, not
    // excessive bloat from uncompressed data.
    let result = compile_to_pdf("", &[], None, &[], false, false).unwrap();

    // Empty page PDF should be compact — mostly font data and PDF structure.
    // Typically 10-30KB depending on embedded font data.
    assert!(
        result.len() < 100_000,
        "Empty page PDF should be under 100KB (baseline), actual: {} bytes ({:.1} KB)",
        result.len(),
        result.len() as f64 / 1024.0
    );
}

#[test]
fn test_compression_effective_for_repetitive_content() {
    // FLATE compression is especially effective on repetitive content.
    // A document with highly repetitive text should compress well,
    // producing a PDF not much larger than a document with less text.
    let short_source = "Hello world.\n\n";
    let short_pdf = compile_to_pdf(short_source, &[], None, &[], false, false).unwrap();

    // 100x the text content, but should compress to much less than 100x the size
    let long_source = "Hello world.\n\n".repeat(100);
    let long_pdf = compile_to_pdf(&long_source, &[], None, &[], false, false).unwrap();

    // With compression, 100x content should produce far less than 10x the PDF size.
    // The ratio demonstrates that content streams are being compressed.
    let size_ratio = long_pdf.len() as f64 / short_pdf.len() as f64;
    assert!(
        size_ratio < 10.0,
        "100x content should produce less than 10x PDF size with compression. \
         Short: {} bytes, Long: {} bytes, Ratio: {:.1}x",
        short_pdf.len(),
        long_pdf.len(),
        size_ratio
    );
}

// --- Tagged PDF and PDF/UA tests (US-096) ---

#[test]
fn test_tagged_pdf_contains_structure_tags() {
    // A tagged PDF with headings should contain StructTreeRoot and heading tags
    let source = "= My Heading\n\nSome paragraph text.\n\n== Sub Heading\n\nMore text.";
    let result = compile_to_pdf(source, &[], None, &[], true, false).unwrap();
    assert!(result.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result);
    // Tagged PDFs must contain a StructTreeRoot
    assert!(
        pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
        "Tagged PDF should contain structure tree or mark info"
    );
}

#[test]
fn test_untagged_pdf_no_structure_tree() {
    // Without tagging, there should be no StructTreeRoot
    let source = "= My Heading\n\nSome text.";
    let result = compile_to_pdf(source, &[], None, &[], false, false).unwrap();
    assert!(result.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result);
    assert!(
        !pdf_str.contains("StructTreeRoot"),
        "Untagged PDF should not contain StructTreeRoot"
    );
}

#[test]
fn test_pdf_ua_produces_valid_pdf() {
    // PDF/UA mode should produce a valid PDF with tagging enabled.
    // PDF/UA-1 requires a document title.
    let source = "#set document(title: \"Accessible Document\")\n= Accessible Document\n\nThis document is PDF/UA compliant.";
    let result = compile_to_pdf(source, &[], None, &[], false, true).unwrap();
    assert!(result.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result);
    // PDF/UA output should contain pdfuaid metadata
    assert!(
        pdf_str.contains("pdfuaid"),
        "PDF/UA output should contain pdfuaid metadata"
    );
}

#[test]
fn test_pdf_ua_implies_tagged() {
    // PDF/UA should produce a tagged PDF even if tagged=false.
    // PDF/UA-1 requires a document title.
    let source = "#set document(title: \"Test\")\n= Heading\n\nParagraph.";
    let result = compile_to_pdf(source, &[], None, &[], false, true).unwrap();
    let pdf_str = String::from_utf8_lossy(&result);
    assert!(
        pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
        "PDF/UA should produce tagged PDF"
    );
}

#[test]
fn test_tagged_pdf_with_table() {
    let source = "#table(columns: 2, [A], [B], [C], [D])";
    let result = compile_to_pdf(source, &[], None, &[], true, false).unwrap();
    assert!(result.starts_with(b"%PDF"));
    // Should be a valid PDF (compilation doesn't fail with tagging)
}

#[test]
fn test_tagged_pdf_with_pdfa_combined() {
    // Tagged + PDF/A should work together
    let source = "= Archival Accessible\n\nBoth standards combined.";
    let result = compile_to_pdf(
        source,
        &[],
        Some(crate::config::PdfStandard::PdfA2b),
        &[],
        true,
        false,
    )
    .unwrap();
    assert!(result.starts_with(b"%PDF"));
    let pdf_str = String::from_utf8_lossy(&result);
    assert!(pdf_str.contains("pdfaid"), "Should contain PDF/A metadata");
    assert!(
        pdf_str.contains("StructTreeRoot") || pdf_str.contains("MarkInfo"),
        "Should contain structure tags"
    );
}

/// The embedded Libertinus Serif faces make the token measurement
/// deterministic on every target (like the digit-advance pin for #621).
/// Ground truth from fontTools `hmtx` sums on the typst-assets faces:
/// "Total" is 2.138em regular and 2.392em bold at 1000 upem — the bold face
/// must be selected for bold runs, not the regular one (issue #624).
#[test]
fn test_text_advance_em_reads_regular_and_bold_faces() {
    let regular: f64 = text_advance_em("Libertinus Serif", false, "Total")
        .expect("the embedded Libertinus Serif regular face must resolve");
    assert!(
        (regular - 2.138).abs() < 1e-6,
        "regular 'Total' should be 2.138em, got {regular}"
    );

    let bold: f64 = text_advance_em("Libertinus Serif", true, "Total")
        .expect("the embedded Libertinus Serif bold face must resolve");
    assert!(
        (bold - 2.392).abs() < 1e-6,
        "bold 'Total' should be 2.392em, got {bold}"
    );
}

/// A character without a glyph (U+E000 private use) yields `None` so the
/// caller can degrade to a measurement-free path; an empty string is a valid
/// zero-width measurement.
#[test]
fn test_text_advance_em_is_none_for_missing_glyphs() {
    assert_eq!(text_advance_em("Libertinus Serif", false, "\u{E000}"), None);
    assert_eq!(text_advance_em("Libertinus Serif", false, ""), Some(0.0));
}

/// A caller that quantizes advances one at a time needs them one at a time.
///
/// Ground truth from the typst-assets Libertinus Serif regular `hmtx`, at its
/// 1000-unit em: `T` 597, `o` 504, `t` 316, `a` 457, `l` 264. Their sum is the
/// 2.138em [`text_advance_em`] reports, but Excel rounds each one to a whole
/// point before adding it, and at 10pt those two orders differ by 0.62pt
/// (issue #1088).
#[test]
fn test_glyph_advances_em_reports_each_glyph_separately() {
    let advances: Vec<f64> = glyph_advances_em("Libertinus Serif", false, "Total")
        .expect("the embedded Libertinus Serif regular face must resolve");
    let expected: [f64; 5] = [0.597, 0.504, 0.316, 0.457, 0.264];
    assert_eq!(advances.len(), expected.len(), "one advance per glyph");
    for (glyph, (advance, want)) in advances.iter().zip(expected).enumerate() {
        assert!(
            (advance - want).abs() < 1e-9,
            "glyph {glyph} of 'Total' should advance {want}em, got {advance}"
        );
    }
    let sum: f64 = advances.iter().sum();
    let total: f64 = text_advance_em("Libertinus Serif", false, "Total")
        .expect("the embedded Libertinus Serif regular face must resolve");
    assert!(
        (sum - total).abs() < 1e-9,
        "the parts must sum to the whole: {sum} against {total}"
    );
}

/// PowerPoint shares its 1.2em line across **every font on the line**, each
/// normalised to a unit line by its OS/2 `usWin*` pair first.
///
/// Arial is the face that shows the metric source, since it is one of the few
/// in the corpus declaring an hhea line gap (1854/-434/**67** per 2048 upem,
/// usWin 1854/434). Its candidate shares are:
///
/// - gap-free proportional `1854/2288 x 1.2` = **0.97238em**
/// - even split `(1.2 + 0.9053 - 0.2119) / 2` = 0.94668em
/// - gap-inclusive proportional `1854/2355 x 1.2` = 0.94471em
///
/// A native PowerPoint 16.111 probe of 14 top-anchored, zero-inset boxes at
/// 11-61pt, every one of whose paragraph marks declares Arial too, lands on the
/// gap-free share at all 14 sizes and outside the export's 0.12pt half-grid of
/// the other two at 12 of them (issue #1176).
///
/// The mark is what the golden mocks change: theirs declare no typeface, so the
/// line also carries the theme's minor Latin font, and the shared box moves.
/// [`arial_slide_seats_match_the_golden_mock_exports`] pins that pair against
/// the committed native exports.
#[test]
fn test_powerpoint_line_box_shares_every_font_on_the_line() {
    let Some((above, below)) = powerpoint_line_box_em("Arial") else {
        return; // no Arial-compatible face on this host
    };

    assert!(
        (above + below - POWERPOINT_LINE_HEIGHT_FACTOR).abs() < 1e-9,
        "the split must still span the 1.2em line, got {above} + {below}"
    );
    assert!(
        (above - 0.97238).abs() < 0.001,
        "a line whose only face is Arial must seat its first baseline 0.97238em \
         below the box top — Arial's share of the box counting no line gap — \
         not {above}em"
    );

    // The mark's face is stated as a metric pair rather than resolved by name:
    // a host with no Calibri-compatible face substitutes one whose metrics are
    // not Calibri's, and the substitution rather than the rule would decide the
    // assertion. Resolution by name is covered by
    // `the_paragraph_mark_face_moves_the_emitted_line_box`, which uses only
    // faces Typst embeds.
    let upem: f64 = 2048.0;
    let arial: (f64, f64) = (1854.0 / upem, 434.0 / upem);
    let calibri: (f64, f64) = (1950.0 / upem, 550.0 / upem);
    let (shared_above, _) = powerpoint_line_box_split_em([arial, calibri])
        .expect("a positive ascent splits the line box");
    assert!(
        shared_above < above - 0.02,
        "adding Calibri — whose 0.936em share is the deeper of the two — must \
         pull the shared box down from Arial's own {above}em, not leave it at \
         {shared_above}em"
    );
    assert!(
        (shared_above - 0.94377).abs() < 0.001,
        "an Arial line whose paragraph mark is Calibri seats at 0.94377em, not \
         {shared_above}em"
    );
}

/// Every Malgun Gothic seat the committed Korean golden-mock exports carry, at
/// the eleven sizes those four decks use.
///
/// Malgun Gothic overflows the box (usWin 2229/495 per 2048 upem, a 1.33008em
/// line) and its own share of it is 0.98194em, which seats every one of these
/// frames a whole point low — 1.0-1.1pt, growing with size, so a share error
/// rather than a rounding one (issue #1176). Inverting the eleven whole-point
/// seats bounds the share to `[0.9375, 0.94643)`.
///
/// What lands inside that interval is the box Malgun Gothic shares with the
/// paragraph mark. These decks declare no typeface on any `<a:endParaRPr>`, so
/// the mark falls to the theme's minor Latin font, Calibri (usWin 1950/550):
/// Malgun keeps the taller normalised ascent and Calibri the deeper normalised
/// descent, and renormalising the pair to 1.2em gives 0.94573em.
///
/// The figures come from the native PowerPoint 16.111 exports under
/// `tests/golden_mocks/business/expected/pptx/`, traced with
/// `mutool draw -F trace`, the same way
/// [`arial_slide_seats_match_the_golden_mock_exports`] reads the English decks.
#[test]
fn malgun_slide_seats_match_the_korean_golden_mock_exports() {
    let upem: f64 = 2048.0;
    // Malgun Gothic: usWinAscent 2229, usWinDescent 495.
    let malgun: (f64, f64) = (2229.0 / upem, 495.0 / upem);
    // Calibri, the theme's minor Latin font: usWinAscent 1950, usWinDescent 550.
    let calibri: (f64, f64) = (1950.0 / upem, 550.0 / upem);

    let (above, below) = powerpoint_line_box_split_em([malgun, calibri])
        .expect("a positive ascent splits the line box");
    assert!(
        (above + below - POWERPOINT_LINE_HEIGHT_FACTOR).abs() < 1e-9,
        "the split must still span the 1.2em line, got {above} + {below}"
    );

    // (font size pt, the seat the exports put inside the line, in pt).
    const EXPORTED: [(f64, f64); 11] = [
        (11.0, 9.96),
        (12.5, 12.06),
        (13.0, 11.88),
        (15.0, 14.04),
        (16.0, 14.88),
        (17.0, 16.08),
        (18.0, 17.04),
        (28.0, 26.04),
        (32.0, 30.00),
        (38.0, 36.00),
        (40.0, 37.92),
    ];
    // The exports quantise a position to a 0.24pt grid, so a whole-point seat
    // is within half of that of the measured one or it is a different model.
    const HALF_GRID_PT: f64 = 0.12 + 1e-9;

    let malgun_alone: f64 = powerpoint_line_box_split_em([malgun])
        .expect("a positive ascent splits the line box")
        .0;
    let mut malgun_alone_misses: usize = 0;
    for (size_pt, export_pt) in EXPORTED {
        let seat_pt: f64 = (above * size_pt).round();
        assert!(
            (seat_pt - export_pt).abs() <= HALF_GRID_PT,
            "at {size_pt}pt the Korean exports seat the baseline {export_pt}pt \
             into the line; the shared box predicts {seat_pt}pt"
        );
        if ((malgun_alone * size_pt).round() - export_pt).abs() > HALF_GRID_PT {
            malgun_alone_misses += 1;
        }
    }

    // Triangulation: Malgun's own share has to be a model this table rules out,
    // or the table would pass without the mark's face in the box at all.
    assert_eq!(
        malgun_alone_misses, 10,
        "Malgun Gothic's own {malgun_alone}em share must miss ten of the eleven \
         cells — 12.5pt is the one size where the two round together"
    );
}

/// Every Arial seat the committed golden-mock exports carry, at the twelve
/// sizes those decks use.
///
/// The figures come from the native PowerPoint 16.111 exports under
/// `tests/golden_mocks/business/expected/pptx/`, traced with
/// `mutool draw -F trace`. A frame's content top is its `a:off` plus the
/// `a:bodyPr` top inset; a centred single-line frame seats its 1.2em line
/// `(content height - 1.2 x size) / 2` further down. Subtracting that leaves
/// the seat inside the line, which the exports put on a whole point (#1074).
///
/// None of these frames is set in Arial alone. Their paragraph marks declare no
/// typeface, so each line also carries the theme's minor Latin font, Calibri,
/// and the box the two share seats at 0.94377em rather than Arial's own
/// 0.97238em (issue #1176). That shared value is 0.001em from the gap-inclusive
/// 0.94471em #1118 fitted to this same table, and rounds to the same point at
/// all twelve sizes — which is why reading the gap as the difference held for
/// as long as it did.
#[test]
fn arial_slide_seats_match_the_golden_mock_exports() {
    let upem: f64 = 2048.0;
    // Arial: usWinAscent 1854, usWinDescent 434 (hhea also declares a 67-unit
    // line gap, which PowerPoint's box does not read).
    let arial: (f64, f64) = (1854.0 / upem, 434.0 / upem);
    // Calibri, the theme's minor Latin font: usWinAscent 1950, usWinDescent 550.
    let calibri: (f64, f64) = (1950.0 / upem, 550.0 / upem);

    let (above, below) = powerpoint_line_box_split_em([arial, calibri])
        .expect("a positive ascent splits the line box");
    assert!(
        (above + below - POWERPOINT_LINE_HEIGHT_FACTOR).abs() < 1e-9,
        "the split must still span the 1.2em line, got {above} + {below}"
    );

    // (font size pt, the seat the exports put inside the line, in pt).
    const EXPORTED: [(f64, f64); 12] = [
        (12.0, 11.04),
        (12.5, 12.06),
        (13.0, 11.88),
        (14.5, 13.92),
        (15.0, 13.92),
        (17.0, 16.08),
        (18.0, 17.04),
        (19.0, 17.88),
        (28.0, 26.04),
        (30.0, 28.08),
        (32.0, 30.00),
        (40.0, 37.92),
    ];
    // The exports quantise a position to a 0.24pt grid, so a whole-point seat
    // is within half of that of the measured one or it is a different model.
    const HALF_GRID_PT: f64 = 0.12 + 1e-9;

    let ascent_em: f64 = arial.0;
    let descent_em: f64 = arial.1;
    let line_gap_em: f64 = 67.0 / upem;
    let natural_em: f64 = ascent_em + descent_em + line_gap_em;
    let rivals: [(&str, f64); 3] = [
        (
            "the even split",
            (POWERPOINT_LINE_HEIGHT_FACTOR + ascent_em - descent_em) / 2.0,
        ),
        (
            "Arial's own gap-free share, with no mark on the line",
            POWERPOINT_LINE_HEIGHT_FACTOR * ascent_em / (ascent_em + descent_em),
        ),
        (
            "the gap given to the ascent side",
            POWERPOINT_LINE_HEIGHT_FACTOR * (ascent_em + line_gap_em) / natural_em,
        ),
    ];
    let mut rival_misses: [usize; 3] = [0; 3];

    for (size_pt, export_pt) in EXPORTED {
        let seat_pt: f64 = (above * size_pt).round();
        assert!(
            (seat_pt - export_pt).abs() <= HALF_GRID_PT,
            "at {size_pt}pt the exports seat the baseline {export_pt}pt into the \
             line; the split predicts {seat_pt}pt"
        );
        for (index, (_, share_em)) in rivals.iter().enumerate() {
            if ((share_em * size_pt).round() - export_pt).abs() > HALF_GRID_PT {
                rival_misses[index] += 1;
            }
        }
    }

    // Triangulation: each rival has to be ruled out by this table, or the table
    // would pass on it too. The even split survives all but the 28pt cells,
    // which is why it stood until #1118.
    for ((name, _), misses) in rivals.iter().zip(rival_misses) {
        assert!(
            misses > 0,
            "{name} must be a model this table rules out, but it misses none of \
             the {} cells",
            EXPORTED.len()
        );
    }
    assert_eq!(
        rival_misses[0], 1,
        "only the 28pt cells separate the even split from the shared box"
    );
}

/// A face whose own line *fits* inside the 1.2em box is shared like any other
/// — the extra leading is not halved.
///
/// Measured on a native PowerPoint 16.112 export of a one-factor probe deck:
/// bottom-anchored text boxes with every inset zeroed, traced with
/// `mutool draw -F trace`, at the 14 sizes below. Georgia is the probe's only
/// face that fits the box (hhea 1878/-449 per 2048 upem, no line gap, so a
/// 1.13623em line), which is what lets it tell the two shares apart: its
/// proportional share is 0.968457em and its even share 0.948877em. A
/// bottom-anchored box keeps `1.2 x size - round(share x size)` under its last
/// baseline, the seat being rounded to a whole point (issue #1074).
///
/// The even split is outside the export's 0.12pt half-grid at 9 of the 14
/// sizes and 2.04pt out at 72pt; the proportional share is inside it at all 14.
/// The five sizes where they agree — 8, 18, 24, 28 and 48 — are the ones where
/// they round to the same point, which is how the branch survived #660 (issue
/// #1118).
#[test]
fn a_face_that_fits_the_line_box_is_shared_like_any_other() {
    // Georgia: hhea ascender 1878, descender -449, no line gap, 2048 upem.
    let upem: f64 = 2048.0;
    let ascent_em: f64 = 1878.0 / upem;
    let descent_em: f64 = 449.0 / upem;
    assert!(
        ascent_em + descent_em < POWERPOINT_LINE_HEIGHT_FACTOR,
        "this test needs a face that fits the box, got {}em",
        ascent_em + descent_em
    );

    let (above, below) = powerpoint_line_box_split_em([(ascent_em, descent_em)])
        .expect("a positive ascent splits the line box");
    assert!(
        (above + below - POWERPOINT_LINE_HEIGHT_FACTOR).abs() < 1e-9,
        "the split must still span the 1.2em line, got {above} + {below}"
    );

    // (font size pt, the gap the export keeps below the last baseline in pt).
    const PROBE: [(f64, f64); 14] = [
        (8.0, 1.680),
        (11.0, 2.160),
        (14.0, 2.800),
        (18.0, 4.560),
        (24.0, 5.920),
        (28.0, 6.560),
        (32.0, 7.400),
        (36.0, 8.200),
        (40.0, 8.960),
        (44.0, 9.880),
        (48.0, 11.560),
        (54.0, 12.760),
        (72.0, 16.360),
        (100.0, 23.120),
    ];
    // The exports quantise a position to a 0.24pt grid, so a modelled gap is
    // within half of that of the measured one or it is a different model. Two
    // of the 14 sizes land exactly on that half-grid, hence the float slack.
    const HALF_GRID_PT: f64 = 0.12 + 1e-9;
    let gap_pt = |share_em: f64, size_pt: f64| -> f64 {
        POWERPOINT_LINE_HEIGHT_FACTOR * size_pt - (share_em * size_pt).round()
    };

    let even_em: f64 = (POWERPOINT_LINE_HEIGHT_FACTOR + ascent_em - descent_em) / 2.0;
    let mut even_misses: usize = 0;
    for (size_pt, export_pt) in PROBE {
        let modelled_pt: f64 = gap_pt(above, size_pt);
        assert!(
            (modelled_pt - export_pt).abs() <= HALF_GRID_PT,
            "at {size_pt}pt the export keeps {export_pt}pt under the baseline; \
             the split predicts {modelled_pt}pt"
        );
        if (gap_pt(even_em, size_pt) - export_pt).abs() > HALF_GRID_PT {
            even_misses += 1;
        }
    }

    // Triangulation: without this the even split would pass the loop above at
    // the five sizes where the two shares round to the same point.
    assert_eq!(
        even_misses, 9,
        "the even split must still be the model this probe rules out"
    );
}

#[test]
fn the_world_hands_typst_the_face_that_kerns_from_the_legacy_table() {
    // Every face Typst shapes with comes through `World::font`, so that is
    // where the kern-source choice has to land: a face carrying both sources
    // must reach the shaper without its GPOS `kern` feature (issue #1116).
    let base: &[u8] = include_bytes!("../../fonts/NotoSansCJKsc-GB2312.otf");
    let font = Font::new(
        Bytes::new(crate::test_support::make_face_carrying_both_kern_sources(
            base,
        )),
        0,
    )
    .expect("the rebuilt face parses");
    assert!(crate::test_support::states_a_gpos_kern_feature(&font));

    let world = MinimalWorld::new_embedded_with_fonts("", &[], std::slice::from_ref(&font));
    let handed_over = world.font(0).expect("the in-memory face leads the book");

    assert!(
        !crate::test_support::states_a_gpos_kern_feature(&handed_over),
        "the compiler must be handed the face that kerns from the legacy table"
    );
    assert!(
        measured_instance(&handed_over)
            .ttf()
            .tables()
            .kern
            .is_some()
    );
}

#[test]
fn the_world_hands_typst_a_one_source_face_unchanged() {
    // A face that states its pairs in GPOS alone has nothing to fall back to,
    // so it must reach the shaper exactly as it was loaded.
    let base: &[u8] = include_bytes!("../../fonts/NotoSansCJKsc-GB2312.otf");
    let font = Font::new(Bytes::new(base.to_vec()), 0).expect("the bundled face parses");

    let world = MinimalWorld::new_embedded_with_fonts("", &[], std::slice::from_ref(&font));
    let handed_over = world.font(0).expect("the in-memory face leads the book");

    assert!(crate::test_support::states_a_gpos_kern_feature(
        &handed_over
    ));
    assert_eq!(handed_over.data().len(), font.data().len());
}

/// One expandable gap of a compiled line: a space glyph and the width the
/// line gave it, in page points.
struct CompiledGap {
    left_pt: f64,
    width_pt: f64,
    is_auto_space: bool,
}

/// The gaps of the first line of `frame`, left to right, read from the glyph
/// advances the exporter will paint. The auto space is the run codegen emits
/// for it: one no-break space of its own.
fn first_line_gaps(frame: &typst::layout::Frame) -> Vec<CompiledGap> {
    use typst::layout::{FrameItem, Transform};

    fn collect(
        frame: &typst::layout::Frame,
        transform: Transform,
        out: &mut Vec<(f64, CompiledGap)>,
    ) {
        for (position, item) in frame.items() {
            let at = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    collect(&group.frame, at.pre_concat(group.transform), out)
                }
                FrameItem::Text(text) => {
                    let mut cursor: f64 = at.tx.to_pt();
                    for glyph in &text.glyphs {
                        let width: f64 = glyph.x_advance.at(text.size).to_pt();
                        let is_space: bool = text.text[glyph.range()]
                            .chars()
                            .all(|c| matches!(c, ' ' | '\u{00A0}' | '\u{3000}'));
                        if is_space && width > 0.0 {
                            out.push((
                                at.ty.to_pt(),
                                CompiledGap {
                                    left_pt: cursor,
                                    width_pt: width,
                                    is_auto_space: text.text == "\u{00A0}",
                                },
                            ));
                        }
                        cursor += width;
                    }
                }
                _ => {}
            }
        }
    }

    let mut gaps: Vec<(f64, CompiledGap)> = Vec::new();
    collect(frame, Transform::identity(), &mut gaps);
    let first_baseline: f64 = gaps
        .iter()
        .map(|(baseline, _)| *baseline)
        .fold(f64::MAX, f64::min);
    let mut line: Vec<CompiledGap> = gaps
        .into_iter()
        .filter(|(baseline, _)| (baseline - first_baseline).abs() < 0.001)
        .map(|(_, gap)| gap)
        .collect();
    line.sort_by(|a, b| a.left_pt.total_cmp(&b.left_pt));
    line
}

/// A justified Latin line in the embedded Libertinus Serif, carrying the same
/// auto-space emission and ceiling codegen writes for a Korean paragraph, and
/// stretched by exactly `demand` (a Typst length expression that may use
/// `sp`, the face's word space). The face makes the numbers the same on every
/// target; the phases do not depend on the script.
fn justified_auto_space_line_source(demand: &str, tail: &str) -> String {
    format!(
        r##"
#set page(width: 400pt, height: 200pt, margin: 0pt)
#set text(font: "Libertinus Serif", size: 10.5pt)
#set par(justify: true, linebreaks: "simple", leading: 0pt)
#set par(justification-limits: (spacing: (min: 80%, max: 0.0001% + 0.5em)))
#let gap = [#metadata("office2pdf-docx-auto-space")#text(spacing: 2.625pt)[\u{{00A0}}]]
#let first = [alpha beta gamma delta#gap;epsilon zeta eta theta#gap;{tail}]
#context {{
  let sp = measure([n n]).width - measure([nn]).width
  block(width: measure(first).width + ({demand}))[#first wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww]
}}
"##
    )
}

fn compile_first_line_gaps(source: &str) -> (Vec<CompiledGap>, Vec<CompiledGap>) {
    let world = MinimalWorld::new(source, &[], &[]);
    let mut document = typst::compile::<PagedDocument>(&world).output.unwrap();
    let before: Vec<CompiledGap> = first_line_gaps(&document.pages()[0].frame);
    apply_completed_frame_passes(&mut document);
    let after: Vec<CompiledGap> = first_line_gaps(&document.pages()[0].frame);
    (before, after)
}

/// Word's phase 1 (issue #1280): while the word spaces of a justified line
/// are short of half an em, they take the whole stretch demand and the East
/// Asian/Latin auto spaces stay at their quarter em. Typst's one ratio moves
/// both, so the completed line is re-spread.
#[test]
fn a_lightly_stretched_justified_line_keeps_its_auto_spaces_at_a_quarter_em() {
    let source = justified_auto_space_line_source("1.5pt", "iota kappa");
    let (before, after) = compile_first_line_gaps(&source);

    assert_eq!(before.len(), 9, "seven word spaces and two auto spaces");
    assert_eq!(after.len(), 9);
    assert!(
        before
            .iter()
            .filter(|gap| gap.is_auto_space)
            .all(|gap| gap.width_pt > 2.625 + 0.01),
        "Typst stretches the auto spaces from the first point of demand"
    );

    let auto_widths: Vec<f64> = after
        .iter()
        .filter(|gap| gap.is_auto_space)
        .map(|gap| gap.width_pt)
        .collect();
    let word_widths: Vec<f64> = after
        .iter()
        .filter(|gap| !gap.is_auto_space)
        .map(|gap| gap.width_pt)
        .collect();
    assert_eq!(auto_widths.len(), 2);
    assert_eq!(word_widths.len(), 7);
    for width in &auto_widths {
        assert!(
            (width - 2.625).abs() < 1e-6,
            "an auto space stays at its quarter em: {width}"
        );
    }
    for width in &word_widths {
        assert!(
            (width - word_widths[0]).abs() < 1e-6,
            "the word spaces share the demand equally: {word_widths:?}"
        );
    }
    let total_before: f64 = before.iter().map(|gap| gap.width_pt).sum();
    let total_after: f64 = after.iter().map(|gap| gap.width_pt).sum();
    assert!(
        (total_before - total_after).abs() < 1e-6,
        "the line keeps its width"
    );
    // The word spaces took back what the auto spaces had been given.
    assert!(
        word_widths[0]
            > before
                .iter()
                .find(|gap| !gap.is_auto_space)
                .unwrap()
                .width_pt
    );
}

/// Word's phase 2: past the word spaces' half em the auto spaces take the
/// remainder equally while the word spaces stay pinned.
#[test]
fn a_heavier_justified_line_pins_its_word_spaces_at_half_an_em() {
    let source = justified_auto_space_line_source("7 * (5.25pt - sp) + 1pt", "iota kappa");
    let (before, after) = compile_first_line_gaps(&source);

    let auto_widths: Vec<f64> = after
        .iter()
        .filter(|gap| gap.is_auto_space)
        .map(|gap| gap.width_pt)
        .collect();
    let word_widths: Vec<f64> = after
        .iter()
        .filter(|gap| !gap.is_auto_space)
        .map(|gap| gap.width_pt)
        .collect();
    assert_eq!(auto_widths.len(), 2);
    assert_eq!(word_widths.len(), 7);
    for width in &word_widths {
        assert!(
            (width - 5.25).abs() < 1e-6,
            "a word space sits at its half em: {width}"
        );
    }
    let total_before: f64 = before.iter().map(|gap| gap.width_pt).sum();
    let remainder: f64 = total_before - 7.0 * 5.25 - 2.0 * 2.625;
    assert!(remainder > 0.5, "the demand reaches phase 2: {remainder}");
    for width in &auto_widths {
        assert!(
            (width - (2.625 + remainder / 2.0)).abs() < 1e-6,
            "the auto spaces share the remainder: {width}"
        );
    }
}

/// Re-spreading moves every later item of the line by the same amount its
/// text moved: a link and an underline drawn over a word that sits between
/// the gaps stay on it. (The last word never moves: the line keeps its
/// width.)
#[test]
fn the_re_spread_line_moves_links_and_underlines_with_their_text() {
    use typst::layout::{Frame, FrameItem, Transform};
    use typst::visualize::Geometry;

    /// `(x, width)` of the run named, its link, and its underline.
    fn anchors(frame: &Frame, transform: Transform, out: &mut [Vec<(f64, f64)>; 3]) {
        for (position, item) in frame.items() {
            let at = transform.pre_concat(Transform::translate(position.x, position.y));
            match item {
                FrameItem::Group(group) => {
                    anchors(&group.frame, at.pre_concat(group.transform), out)
                }
                FrameItem::Text(text) if text.text.as_str() == "iota" => {
                    out[0].push((at.tx.to_pt(), text.width().to_pt()));
                }
                FrameItem::Link(_, size) => out[1].push((at.tx.to_pt(), size.x.to_pt())),
                FrameItem::Shape(shape, _) => {
                    if let Geometry::Line(to) = shape.geometry {
                        out[2].push((at.tx.to_pt(), to.x.to_pt()));
                    }
                }
                _ => {}
            }
        }
    }

    let source = justified_auto_space_line_source(
        "1.5pt",
        r#"#underline(evade: false)[#link("https://example.com")[iota]] kappa"#,
    );
    let world = MinimalWorld::new(&source, &[], &[]);
    let mut document = typst::compile::<PagedDocument>(&world).output.unwrap();
    let mut before = [Vec::new(), Vec::new(), Vec::new()];
    anchors(
        &document.pages()[0].frame,
        Transform::identity(),
        &mut before,
    );
    apply_completed_frame_passes(&mut document);
    let mut after = [Vec::new(), Vec::new(), Vec::new()];
    anchors(
        &document.pages()[0].frame,
        Transform::identity(),
        &mut after,
    );

    assert_eq!(before[0].len(), 1, "one `iota` run");
    assert_eq!(after[1].len(), 1, "one link");
    assert_eq!(after[2].len(), 1, "one underline");
    let (text_x, text_width) = after[0][0];
    assert!(
        (text_x - before[0][0].0).abs() > 0.01,
        "a word between the gaps moves when they are re-spread"
    );
    for (kind, name) in [(1, "link"), (2, "underline")] {
        let (x, width) = after[kind][0];
        assert!(
            (x - text_x).abs() < 1e-6,
            "the {name} starts where its text starts: {x} vs {text_x}"
        );
        assert!(
            (width - text_width).abs() < 1e-6,
            "the {name} is as wide as its text: {width} vs {text_width}"
        );
    }
}

/// A private directory of font files that is removed when the test ends.
struct TempFontDir {
    path: PathBuf,
}

impl TempFontDir {
    fn new(label: &str) -> Self {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time should be valid")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("office2pdf-{label}-{unique}"));
        std::fs::create_dir_all(&path).expect("temp font dir should be creatable");
        Self { path }
    }
}

impl Drop for TempFontDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.path);
    }
}

/// The tracked Noto Serif face, which every runner has.
fn tracked_noto_serif_bytes() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("fonts/NotoSerif-Regular.ttf");
    std::fs::read(path).expect("the tracked Noto Serif face should be readable")
}

/// Index the given directories in order, the way [`get_fonts_for_extra_paths`]
/// indexes the Office bundle ahead of the system, without the host's fonts.
fn font_data_for_dirs(font_dirs: &[PathBuf]) -> CachedFontData {
    let (book, fonts) = discover_fonts(font_dirs, false, false);
    CachedFontData {
        book: LazyHash::new(book),
        fonts,
    }
}

fn hhea_line_gap(font: &Font) -> i16 {
    measured_instance(font).ttf().tables().hhea.line_gap
}

const NOTO_SERIF_LINE_GAP: i16 = 0;
const REWRITTEN_LINE_GAP: i16 = 200;

/// Word draws with the face an Office application bundles, but it paces the
/// line on the system-shipped face that shares that face's PostScript name.
/// Measured on Times New Roman, whose bundled copy declares no `hhea` line gap
/// while the system copy declares 87/2048 (issue #1284).
#[test]
fn a_system_face_sharing_the_bundled_faces_postscript_name_supplies_the_line_metrics() {
    let bundle = TempFontDir::new("office-bundle");
    let system = TempFontDir::new("system-fonts");
    std::fs::write(
        bundle.path.join("NotoSerif-Regular.ttf"),
        crate::test_support::make_face_with_hhea_line_gap(
            tracked_noto_serif_bytes(),
            REWRITTEN_LINE_GAP,
        ),
    )
    .unwrap();
    std::fs::write(
        system.path.join("NotoSerif-Regular.ttf"),
        tracked_noto_serif_bytes(),
    )
    .unwrap();
    let data = font_data_for_dirs(&[bundle.path.clone(), system.path.clone()]);

    // The premise: the bundled copy is the face the compiler shapes with.
    let shaped_index = best_face_index(&data, "Noto Serif").expect("the family is indexed");
    let shaped = data.fonts[shaped_index]
        .get()
        .expect("the bundled face loads");
    assert!(
        data.fonts[shaped_index]
            .path()
            .unwrap()
            .starts_with(&bundle.path)
    );
    assert_eq!(hhea_line_gap(&shaped), REWRITTEN_LINE_GAP);

    let measured = line_metric_face_in(
        &data,
        std::slice::from_ref(&bundle.path),
        std::slice::from_ref(&system.path),
        "Noto Serif",
    )
    .expect("a line-metric face resolves");
    assert_eq!(
        hhea_line_gap(&measured),
        NOTO_SERIF_LINE_GAP,
        "the line box must follow the system copy's hhea, not the bundled one's"
    );
}

/// A copy outside the system font directories does not shadow the bundle: a
/// user-installed Rockwell sharing the bundled face's PostScript name left
/// Word's pitch on the bundled metrics, even after a relaunch (issue #1284).
#[test]
fn a_user_installed_face_sharing_the_postscript_name_does_not_shadow_the_bundle() {
    let bundle = TempFontDir::new("office-bundle");
    let user = TempFontDir::new("user-fonts");
    std::fs::write(
        bundle.path.join("NotoSerif-Regular.ttf"),
        crate::test_support::make_face_with_hhea_line_gap(
            tracked_noto_serif_bytes(),
            REWRITTEN_LINE_GAP,
        ),
    )
    .unwrap();
    std::fs::write(
        user.path.join("NotoSerif-Regular.ttf"),
        tracked_noto_serif_bytes(),
    )
    .unwrap();
    let data = font_data_for_dirs(&[bundle.path.clone(), user.path.clone()]);

    let measured =
        line_metric_face_in(&data, std::slice::from_ref(&bundle.path), &[], "Noto Serif")
            .expect("a line-metric face resolves");
    assert_eq!(hhea_line_gap(&measured), REWRITTEN_LINE_GAP);
}

/// When the compiler already shapes with a face outside the bundle, that face
/// is the one measured; nothing is swapped underneath it.
#[test]
fn a_face_resolved_outside_the_bundle_keeps_its_own_line_metrics() {
    let bundle = TempFontDir::new("office-bundle");
    let system = TempFontDir::new("system-fonts");
    std::fs::write(
        bundle.path.join("NotoSerif-Regular.ttf"),
        crate::test_support::make_face_with_hhea_line_gap(
            tracked_noto_serif_bytes(),
            REWRITTEN_LINE_GAP,
        ),
    )
    .unwrap();
    std::fs::write(
        system.path.join("NotoSerif-Regular.ttf"),
        crate::test_support::make_face_with_hhea_line_gap(tracked_noto_serif_bytes(), 300),
    )
    .unwrap();
    // System first: the resolved face is the system copy itself.
    let data = font_data_for_dirs(&[system.path.clone(), bundle.path.clone()]);

    let measured = line_metric_face_in(
        &data,
        std::slice::from_ref(&bundle.path),
        std::slice::from_ref(&system.path),
        "Noto Serif",
    )
    .expect("a line-metric face resolves");
    assert_eq!(hhea_line_gap(&measured), 300);
}

/// The defect as filed: on a Mac with Office installed, `Times New Roman`
/// resolves Word's bundled `times.ttf` (`hhea` 1825 / -443 / 0) while Word
/// paces on the system copy's 1825 / -443 / 87, so a native export sets a
/// 12pt body line 13.80pt apart and ours set it 13.29pt apart (issue #1284).
#[test]
fn times_new_roman_line_box_follows_the_system_copy_beside_the_office_bundle() {
    let bundled = std::path::Path::new(
        "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf",
    );
    let system = std::path::Path::new("/System/Library/Fonts/Supplemental/Times New Roman.ttf");
    if !bundled.exists() || !system.exists() {
        return;
    }
    let (top_em, bottom_em, pitch_em) =
        font_line_metrics_em("Times New Roman").expect("the face is installed");
    let tolerance = 1e-9;
    assert!((top_em - 1912.0 / 2048.0).abs() < tolerance, "top {top_em}");
    assert!(
        (bottom_em - 443.0 / 2048.0).abs() < tolerance,
        "bottom {bottom_em}"
    );
    assert!(
        (pitch_em - 2355.0 / 2048.0).abs() < tolerance,
        "pitch {pitch_em}"
    );
    let gap_em = font_line_gap_em("Times New Roman").expect("the face is installed");
    assert!((gap_em - 87.0 / 2048.0).abs() < tolerance, "gap {gap_em}");
}

#[test]
fn test_embedded_font_feature_controls_native_font_book() {
    // Exclude disk fonts so an installed copy cannot conceal a feature leak.
    let (book, fonts) = discover_fonts(&[], false, true);
    assert_eq!(
        !fonts.is_empty(),
        cfg!(feature = "embedded-fonts"),
        "Typst's embedded fonts must follow the native Cargo feature"
    );
    assert_eq!(
        book.select_family("new computer modern").next().is_some(),
        cfg!(feature = "embedded-fonts")
    );
}

#[test]
fn test_font_feature_caller_fonts_compile_without_discovered_fonts() {
    let font_data = font_data_for_dirs(&[]);
    let source = "#set text(font: \"Noto Serif\")\nOffice report: Revenue 2026";
    let world = MinimalWorld::new_with_font_source(
        source,
        &[],
        FontSource::InMemory(InMemoryFontData::new(
            crate::bundled_fonts::noto_serif_fonts(),
            FallbackFontData::Shared(Arc::new(font_data)),
        )),
    );
    let document = typst::compile::<PagedDocument>(&world)
        .output
        .expect("caller fonts suffice without any discovered fonts");
    let pdf = typst_pdf::pdf(&document, &typst_pdf::PdfOptions::default()).unwrap();
    let text = pdf_extract::extract_text_from_mem(&pdf).unwrap();
    assert!(text.contains("Office report: Revenue 2026"));
}
