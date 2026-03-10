use super::*;

#[test]
fn test_simple_fraction() {
    let xml =
        "<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>";
    assert_eq!(omml_to_typst(xml), "frac(a, b)");
}

#[test]
fn test_superscript() {
    let xml =
        "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>";
    assert_eq!(omml_to_typst(xml), "x^2");
}

#[test]
fn test_subscript() {
    let xml =
        "<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>";
    assert_eq!(omml_to_typst(xml), "x_1");
}

#[test]
fn test_sub_superscript() {
    let xml = "<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>";
    assert_eq!(omml_to_typst(xml), "x_i^2");
}

#[test]
fn test_square_root() {
    let xml = r#"<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>"#;
    assert_eq!(omml_to_typst(xml), "sqrt(x)");
}

#[test]
fn test_nth_root() {
    let xml = r#"<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>"#;
    assert_eq!(omml_to_typst(xml), "root(3, x)");
}

#[test]
fn test_parentheses() {
    let xml = r#"<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:d>"#;
    assert_eq!(omml_to_typst(xml), "(x+y)");
}

#[test]
fn test_complex_equation() {
    let xml = "<m:f><m:num><m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:num><m:den><m:d><m:e><m:r><m:t>b</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>c</m:t></m:r></m:e></m:d></m:den></m:f>";
    assert_eq!(omml_to_typst(xml), "frac(a^2, (b+c))");
}

#[test]
fn test_sum_with_limits() {
    let xml = r#"<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>"#;
    assert_eq!(omml_to_typst(xml), "sum_(i=1)^n i");
}

#[test]
fn test_emc2() {
    let xml = "<m:r><m:t>E</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>";
    // Space before sSup separates run "m" from base "c" (both valid in Typst math)
    assert_eq!(omml_to_typst(xml), "E=m c^2");
}

#[test]
fn test_quadratic_formula() {
    let xml = r#"<m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r><m:f><m:num><m:r><m:t>-b±</m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>-4ac</m:t></m:r></m:e></m:rad></m:num><m:den><m:r><m:t>2a</m:t></m:r></m:den></m:f>"#;
    // ± → plus.minus; -4ac → -4a c (letters split for implicit multiplication)
    assert_eq!(
        omml_to_typst(xml),
        "x=frac(-b plus.minus sqrt(b^2-4a c), 2a)"
    );
}

#[test]
fn test_scan_display_math() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <w:body>
                <w:p>
                    <m:oMathPara>
                        <m:oMath><m:r><m:t>E</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>
                    </m:oMathPara>
                </w:p>
            </w:body>
        </w:document>"#;

    let results = scan_math_equations(xml);
    assert_eq!(results.len(), 1);
    assert_eq!(results[0].0, 0);
    assert_eq!(results[0].1, "E=m c^2");
    assert!(results[0].2);
}

#[test]
fn test_scan_inline_math() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <w:body>
                <w:p><w:r><w:t>Text</w:t></w:r></w:p>
                <w:p>
                    <m:oMath><m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>5</m:t></m:r></m:oMath>
                </w:p>
            </w:body>
        </w:document>"#;

    let results = scan_math_equations(xml);
    assert_eq!(results.len(), 1);
    assert_eq!(results[0].0, 1);
    assert_eq!(results[0].1, "x=5");
    assert!(!results[0].2);
}

#[test]
fn test_scan_multiple_equations() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <w:body>
                <w:p><m:oMathPara><m:oMath><m:r><m:t>a=1</m:t></m:r></m:oMath></m:oMathPara></w:p>
                <w:p><w:r><w:t>text</w:t></w:r></w:p>
                <w:p><m:oMath><m:r><m:t>b=2</m:t></m:r></m:oMath></w:p>
            </w:body>
        </w:document>"#;

    let results = scan_math_equations(xml);
    assert_eq!(results.len(), 2);
    assert_eq!(results[0].0, 0);
    assert_eq!(results[0].1, "a=1");
    assert!(results[0].2);
    assert_eq!(results[1].0, 2);
    assert_eq!(results[1].1, "b=2");
    assert!(!results[1].2);
}

// --- map_math_text tests ---

#[test]
fn test_map_math_text_pi() {
    assert_eq!(map_math_text("π"), "pi");
}

#[test]
fn test_map_math_text_pi_r() {
    // π followed by r should insert space for implicit multiplication
    assert_eq!(map_math_text("πr"), "pi r");
}

#[test]
fn test_map_math_text_multiple_greek() {
    assert_eq!(map_math_text("αβ"), "alpha beta");
}

#[test]
fn test_map_math_text_greek_with_operator() {
    // Operators separate naturally, no extra spaces needed
    assert_eq!(map_math_text("α+β"), "alpha+beta");
}

#[test]
fn test_map_math_text_digit_before_greek() {
    assert_eq!(map_math_text("2π"), "2 pi");
}

#[test]
fn test_map_math_text_letter_before_greek() {
    assert_eq!(map_math_text("rπ"), "r pi");
}

#[test]
fn test_map_math_text_ascii_passthrough() {
    // Plain ASCII letters and digits pass through unchanged
    assert_eq!(map_math_text("x+y=5"), "x+y=5");
}

#[test]
fn test_map_math_text_uppercase_greek() {
    assert_eq!(map_math_text("Ω"), "Omega");
    assert_eq!(map_math_text("Δ"), "Delta");
    assert_eq!(map_math_text("Σ"), "Sigma");
}

#[test]
fn test_map_math_text_math_symbols() {
    assert_eq!(map_math_text("∞"), "infinity");
    assert_eq!(map_math_text("∂"), "partial");
    assert_eq!(map_math_text("∇"), "nabla");
    assert_eq!(map_math_text("∅"), "emptyset");
}

#[test]
fn test_map_math_text_operator_symbols() {
    assert_eq!(map_math_text("±"), "plus.minus");
    assert_eq!(map_math_text("×"), "times");
    assert_eq!(map_math_text("÷"), "div");
    assert_eq!(map_math_text("≤"), "lt.eq");
    assert_eq!(map_math_text("≥"), "gt.eq");
    assert_eq!(map_math_text("≠"), "eq.not");
    assert_eq!(map_math_text("≈"), "approx");
}

#[test]
fn test_map_math_text_set_symbols() {
    assert_eq!(map_math_text("∈"), "in");
    assert_eq!(map_math_text("∉"), "in.not");
    assert_eq!(map_math_text("⊂"), "subset");
    assert_eq!(map_math_text("⊃"), "supset");
    assert_eq!(map_math_text("∪"), "union");
    assert_eq!(map_math_text("∩"), "sect");
}

#[test]
fn test_map_math_text_empty() {
    assert_eq!(map_math_text(""), "");
}

#[test]
fn test_map_math_text_unknown_unicode_passthrough() {
    // Unknown Unicode characters pass through unchanged
    assert_eq!(map_math_text("★"), "★");
}

#[test]
fn test_map_math_text_complex_expression() {
    // Mixed Greek and ASCII with operators
    assert_eq!(map_math_text("θ=2πr"), "theta=2 pi r");
}

#[test]
fn test_map_math_text_splits_unknown_multi_letter() {
    // Unknown multi-letter sequences split into individual variables
    assert_eq!(map_math_text("ka"), "k a");
    assert_eq!(map_math_text("abc"), "a b c");
}

#[test]
fn test_map_math_text_preserves_known_function_names() {
    // Known math function names are preserved intact
    assert_eq!(map_math_text("cos"), "cos");
    assert_eq!(map_math_text("sin"), "sin");
    assert_eq!(map_math_text("log"), "log");
    assert_eq!(map_math_text("lim"), "lim");
}

#[test]
fn test_map_math_text_split_with_digits() {
    // Letters after digits are split
    assert_eq!(map_math_text("-4ac"), "-4a c");
    // Single letter after digit is fine
    assert_eq!(map_math_text("2a"), "2a");
}

// --- parse_math_run with Greek letters via omml_to_typst ---

#[test]
fn test_omml_pi_mapped() {
    let xml = "<m:r><m:t>π</m:t></m:r>";
    assert_eq!(omml_to_typst(xml), "pi");
}

#[test]
fn test_omml_pi_r_spaced() {
    let xml = "<m:r><m:t>πr</m:t></m:r>";
    assert_eq!(omml_to_typst(xml), "pi r");
}

#[test]
fn test_omml_alpha_plus_beta() {
    let xml = "<m:r><m:t>α+β</m:t></m:r>";
    assert_eq!(omml_to_typst(xml), "alpha+beta");
}

// --- US-310: groupChr (overbrace/underbrace) tests ---

#[test]
fn test_group_chr_overbrace() {
    let xml = r#"<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/></m:groupChrPr><m:e><m:r><m:t>a+b</m:t></m:r></m:e></m:groupChr>"#;
    assert_eq!(omml_to_typst(xml), "overbrace(a+b)");
}

#[test]
fn test_group_chr_underbrace() {
    let xml = r#"<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/></m:groupChrPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:groupChr>"#;
    assert_eq!(omml_to_typst(xml), "underbrace(x+y)");
}

#[test]
fn test_group_chr_default_underbrace() {
    // Default groupChr without explicit chr attr should use underbrace
    let xml = r#"<m:groupChr><m:groupChrPr><m:pos m:val="bot"/></m:groupChrPr><m:e><m:r><m:t>z</m:t></m:r></m:e></m:groupChr>"#;
    assert_eq!(omml_to_typst(xml), "underbrace(z)");
}

// --- US-311: subscript/superscript parentheses tests ---

#[test]
fn test_superscript_multi_token_parens() {
    let xml = "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>n+1</m:t></m:r></m:sup></m:sSup>";
    assert_eq!(omml_to_typst(xml), "x^(n+1)");
}

#[test]
fn test_subscript_multi_token_parens() {
    let xml = "<m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i+1</m:t></m:r></m:sub></m:sSub>";
    assert_eq!(omml_to_typst(xml), "a_(i+1)");
}

// --- US-312: empty radicand tests ---

#[test]
fn test_radical_empty_radicand() {
    let xml = r#"<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e></m:e></m:rad>"#;
    let result = omml_to_typst(xml);
    assert!(
        result.contains("sqrt(") && result.ends_with(')'),
        "Empty radicand should produce valid sqrt(): got '{result}'"
    );
    // Should not be "sqrt()" — needs a placeholder
    assert_ne!(result, "sqrt()", "Empty radicand should have a placeholder");
}

#[test]
fn test_root_empty_radicand_with_degree() {
    let xml = r#"<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e></m:e></m:rad>"#;
    let result = omml_to_typst(xml);
    assert!(
        result.starts_with("root(3,") && result.ends_with(')'),
        "Empty radicand with degree should produce valid root(): got '{result}'"
    );
}

// --- US-313: delimiter balancing tests ---

#[test]
fn test_delimiter_empty_begin_chr() {
    // When begChr is empty, should not produce unbalanced `)` alone
    let xml = r#"<m:d><m:dPr><m:begChr m:val=""/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"#;
    let result = omml_to_typst(xml);
    // Must not end with bare `)` without matching `(`
    assert!(
        !result.ends_with(')') || result.contains('('),
        "Empty begChr should not produce unmatched ')': got '{result}'"
    );
}

#[test]
fn test_delimiter_empty_end_chr() {
    // When endChr is empty, should not produce unbalanced `(`
    let xml = r#"<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=""/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"#;
    let result = omml_to_typst(xml);
    // Must not have bare `(` without matching `)`
    assert!(
        !result.starts_with('(') || result.contains(')'),
        "Empty endChr should not produce unmatched '(': got '{result}'"
    );
}

#[test]
fn test_delimiter_both_empty() {
    // When both begChr and endChr are empty, should just emit content
    let xml = r#"<m:d><m:dPr><m:begChr m:val=""/><m:endChr m:val=""/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"#;
    let result = omml_to_typst(xml);
    assert_eq!(
        result, "x",
        "Both empty delimiters should emit bare content: got '{result}'"
    );
}

// --- US-314: non-ASCII text in math context ---

#[test]
fn test_non_ascii_cyrillic_in_math() {
    let xml = r#"<m:r><m:t>если</m:t></m:r>"#;
    let result = omml_to_typst(xml);
    // Cyrillic text in math should be wrapped in upright() to avoid "unknown variable"
    assert!(
        result.contains("upright("),
        "Cyrillic text in math should be wrapped in upright(): got '{result}'"
    );
}

#[test]
fn test_non_ascii_single_char_passthrough() {
    // Single non-ASCII char that maps to a Typst symbol should pass through
    let xml = r#"<m:r><m:t>α</m:t></m:r>"#;
    assert_eq!(omml_to_typst(xml), "alpha");
}

// --- US-380: subscript/superscript with empty base ---

#[test]
fn test_subscript_empty_base() {
    // When base is empty (e.g., <m:e/> or <m:e></m:e>), the output must
    // not start with bare `_` which is invalid in Typst math.
    let xml = r#"<m:sSub><m:e></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub>"#;
    let result = omml_to_typst(xml);
    assert!(
        !result.starts_with('_'),
        "Empty base subscript must not start with bare '_': got '{result}'"
    );
    assert!(
        result.contains("_2"),
        "Should still contain subscript: got '{result}'"
    );
}

#[test]
fn test_superscript_empty_base() {
    let xml = r#"<m:sSup><m:e></m:e><m:sup><m:r><m:t>1</m:t></m:r></m:sup></m:sSup>"#;
    let result = omml_to_typst(xml);
    assert!(
        !result.starts_with('^'),
        "Empty base superscript must not start with bare '^': got '{result}'"
    );
    assert!(
        result.contains("^1"),
        "Should still contain superscript: got '{result}'"
    );
}

#[test]
fn test_sub_superscript_empty_base() {
    let xml = r#"<m:sSubSup><m:e></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup></m:sSubSup>"#;
    let result = omml_to_typst(xml);
    assert!(
        !result.starts_with('_'),
        "Empty base sub-superscript must not start with bare '_': got '{result}'"
    );
}

// --- US-381: literal parens in math run text ---

#[test]
fn test_math_text_literal_parens() {
    // Literal ( and ) in <m:t> should not break Typst math function calls
    let xml = r#"<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:deg/><m:e><m:r><m:t>)2(</m:t></m:r></m:e></m:rad>"#;
    let result = omml_to_typst(xml);
    // Must produce valid Typst: sqrt() must have its radicand argument
    // The result must not be "sqrt()2()" which would fail compilation
    assert!(
        !result.contains("sqrt()"),
        "Literal parens in radicand must not produce empty sqrt(): got '{result}'"
    );
}

// --- Blackboard bold letter mappings ---

#[test]
fn test_blackboard_bold_letters() {
    assert_eq!(unicode_to_typst('ℂ'), Some("CC"));
    assert_eq!(unicode_to_typst('ℍ'), Some("HH"));
    assert_eq!(unicode_to_typst('ℕ'), Some("NN"));
    assert_eq!(unicode_to_typst('ℙ'), Some("PP"));
    assert_eq!(unicode_to_typst('ℚ'), Some("QQ"));
    assert_eq!(unicode_to_typst('ℝ'), Some("RR"));
    assert_eq!(unicode_to_typst('ℤ'), Some("ZZ"));
}

#[test]
fn test_blackboard_bold_in_math_text() {
    // Blackboard bold letters should map to Typst symbols, not be wrapped in upright()
    assert_eq!(map_math_text("ℝ"), "RR");
    assert_eq!(map_math_text("x∈ℝ"), "x in RR");
}

#[test]
fn test_blackboard_bold_via_omml() {
    let xml = r#"<m:r><m:t>ℝ</m:t></m:r>"#;
    assert_eq!(omml_to_typst(xml), "RR");
}

// --- Extended relation symbol mappings ---

#[test]
fn test_extended_relations() {
    assert_eq!(unicode_to_typst('≡'), Some("equiv"));
    assert_eq!(unicode_to_typst('∼'), Some("tilde.op"));
    assert_eq!(unicode_to_typst('≅'), Some("tilde.eq"));
    assert_eq!(unicode_to_typst('≪'), Some("lt.double"));
    assert_eq!(unicode_to_typst('≫'), Some("gt.double"));
    assert_eq!(unicode_to_typst('⊆'), Some("subset.eq"));
    assert_eq!(unicode_to_typst('⊇'), Some("supset.eq"));
    assert_eq!(unicode_to_typst('∝'), Some("prop"));
    assert_eq!(unicode_to_typst('∴'), Some("therefore"));
    assert_eq!(unicode_to_typst('∵'), Some("because"));
}

#[test]
fn test_extended_relations_in_math_text() {
    assert_eq!(map_math_text("a≡b"), "a equiv b");
    assert_eq!(map_math_text("A⊆B"), "A subset.eq B");
}

// --- Arrow symbol mappings ---

#[test]
fn test_arrow_symbols() {
    assert_eq!(unicode_to_typst('←'), Some("arrow.l"));
    assert_eq!(unicode_to_typst('↑'), Some("arrow.t"));
    assert_eq!(unicode_to_typst('→'), Some("arrow.r"));
    assert_eq!(unicode_to_typst('↓'), Some("arrow.b"));
    assert_eq!(unicode_to_typst('↔'), Some("arrow.l.r"));
    assert_eq!(unicode_to_typst('↦'), Some("arrow.r.bar"));
    assert_eq!(unicode_to_typst('↪'), Some("arrow.r.hook"));
    assert_eq!(unicode_to_typst('↼'), Some("harpoon.lt"));
    assert_eq!(unicode_to_typst('⇀'), Some("harpoon.rt"));
    assert_eq!(unicode_to_typst('⇐'), Some("arrow.l.double"));
    assert_eq!(unicode_to_typst('⇒'), Some("arrow.r.double"));
    assert_eq!(unicode_to_typst('⇔'), Some("arrow.l.r.double"));
}

#[test]
fn test_arrow_in_math_text() {
    assert_eq!(map_math_text("x→y"), "x arrow.r y");
    assert_eq!(map_math_text("A⇒B"), "A arrow.r.double B");
}

// --- Additional operator mappings ---

#[test]
fn test_additional_operators() {
    assert_eq!(unicode_to_typst('∠'), Some("angle"));
    assert_eq!(unicode_to_typst('∧'), Some("and"));
    assert_eq!(unicode_to_typst('∨'), Some("or"));
    assert_eq!(unicode_to_typst('∘'), Some("compose"));
    assert_eq!(unicode_to_typst('⋅'), Some("dot.op"));
    assert_eq!(unicode_to_typst('∓'), Some("minus.plus"));
    assert_eq!(unicode_to_typst('¬'), Some("not"));
    assert_eq!(unicode_to_typst('⊕'), Some("plus.circle"));
    assert_eq!(unicode_to_typst('⊗'), Some("times.circle"));
    assert_eq!(unicode_to_typst('⊙'), Some("dot.circle"));
    assert_eq!(unicode_to_typst('⊢'), Some("tack.r"));
    assert_eq!(unicode_to_typst('⊣'), Some("tack.l"));
    assert_eq!(unicode_to_typst('⊤'), Some("top"));
    assert_eq!(unicode_to_typst('⊥'), Some("perp"));
}

#[test]
fn test_operators_in_math_text() {
    assert_eq!(map_math_text("a∧b"), "a and b");
    assert_eq!(map_math_text("¬p"), "not p");
    assert_eq!(map_math_text("f∘g"), "f compose g");
}

// --- Floor/ceiling delimiter tests ---

#[test]
fn test_floor_ceiling_delimiters() {
    assert_eq!(map_delimiter("⌈"), "⌈");
    assert_eq!(map_delimiter("⌉"), "⌉");
    assert_eq!(map_delimiter("⌊"), "⌊");
    assert_eq!(map_delimiter("⌋"), "⌋");
}

#[test]
fn test_floor_delimiter_via_omml() {
    let xml = r#"<m:d><m:dPr><m:begChr m:val="⌊"/><m:endChr m:val="⌋"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"#;
    assert_eq!(omml_to_typst(xml), "⌊x⌋");
}

#[test]
fn test_ceiling_delimiter_via_omml() {
    let xml = r#"<m:d><m:dPr><m:begChr m:val="⌈"/><m:endChr m:val="⌉"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"#;
    assert_eq!(omml_to_typst(xml), "⌈x⌉");
}

// --- Extended accent mappings ---

#[test]
fn test_extended_accents() {
    assert_eq!(map_accent("\u{0301}"), "acute");
    assert_eq!(map_accent("\u{0300}"), "grave");
    assert_eq!(map_accent("\u{0305}"), "macron");
    assert_eq!(map_accent("\u{030A}"), "circle");
}

#[test]
fn test_acute_accent_via_omml() {
    let xml =
        r#"<m:acc><m:accPr><m:chr m:val="́"/></m:accPr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:acc>"#;
    assert_eq!(omml_to_typst(xml), "acute(a)");
}

#[test]
fn test_grave_accent_via_omml() {
    let xml =
        r#"<m:acc><m:accPr><m:chr m:val="̀"/></m:accPr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:acc>"#;
    assert_eq!(omml_to_typst(xml), "grave(a)");
}

// --- Additional n-ary operator mappings ---

#[test]
fn test_additional_nary_operators() {
    assert_eq!(map_nary_operator("\u{2210}"), "product.co");
    assert_eq!(map_nary_operator("\u{22C0}"), "and.big");
    assert_eq!(map_nary_operator("\u{22C1}"), "or.big");
}

#[test]
fn test_coproduct_via_omml() {
    let xml = r#"<m:nary><m:naryPr><m:chr m:val="∐"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup/><m:e><m:r><m:t>A</m:t></m:r></m:e></m:nary>"#;
    assert_eq!(omml_to_typst(xml), "product.co_i A");
}

#[test]
fn test_big_and_via_omml() {
    let xml = r#"<m:nary><m:naryPr><m:chr m:val="⋀"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup/><m:e><m:r><m:t>p</m:t></m:r></m:e></m:nary>"#;
    assert_eq!(omml_to_typst(xml), "and.big_i p");
}
