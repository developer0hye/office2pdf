//! Metric-compatible font substitution table.
//!
//! Maps common Microsoft fonts to open-source metric-compatible alternatives.
//! When the requested font is unavailable, the substitutes are tried in order.
//! Uses a `match` statement for zero-cost static lookup (no runtime allocation).

/// Return metric-compatible substitute font names for the given font family.
///
/// Returns `None` if no substitution is defined for the font (i.e., it is not
/// a known Microsoft font that has metric-compatible open-source alternatives).
///
/// The returned slice is ordered by preference — the first entry is the best
/// metric-compatible match.
pub fn substitutes(font_family: &str) -> Option<&'static [&'static str]> {
    // Case-insensitive matching: normalize to lowercase for comparison.
    // Stack-allocated buffer avoids heap allocation for typical font names.
    let lower = font_family.to_ascii_lowercase();
    match lower.as_str() {
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
        _ => None,
    }
}

/// Build a Typst font fallback list string for the given font family.
///
/// If substitutions exist, returns a Typst array literal like
/// `("Calibri", "Carlito", "Liberation Sans")`.
/// If no substitutions exist, returns a simple quoted name like `"Helvetica"`.
pub fn font_with_fallbacks(font_family: &str) -> String {
    match substitutes(font_family) {
        Some(subs) => {
            let mut result = String::with_capacity(64);
            result.push('(');
            result.push('"');
            result.push_str(font_family);
            result.push('"');
            for sub in subs {
                result.push_str(", \"");
                result.push_str(sub);
                result.push('"');
            }
            result.push(')');
            result
        }
        None => {
            let mut result = String::with_capacity(font_family.len() + 2);
            result.push('"');
            result.push_str(font_family);
            result.push('"');
            result
        }
    }
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
}
