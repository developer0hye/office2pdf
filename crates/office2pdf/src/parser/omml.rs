//! OMML (Office Math Markup Language) to Typst math notation converter.
//!
//! Parses OMML XML elements (m:oMath, m:oMathPara) from DOCX documents
//! and converts them to Typst math notation strings.

use quick_xml::Reader;
use quick_xml::events::Event;

use super::xml_util;

/// Convert an OMML XML fragment to Typst math notation.
///
/// The input should be the inner content of an `<m:oMath>` element.
/// Returns the Typst math notation string (without `$` delimiters).
pub(crate) fn omml_to_typst(xml: &str) -> String {
    let wrapped = format!(
        "<root xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">{xml}</root>"
    );
    let mut reader = Reader::from_str(&wrapped);
    let mut result = String::new();

    // Skip the <root> wrapper start
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"root" => break,
            Ok(Event::Eof) => return String::new(),
            Err(_) => return String::new(),
            _ => {}
        }
    }

    parse_omml_children(&mut reader, &mut result, b"root");
    result.trim().to_string()
}

/// Recursively parse OMML children and append Typst math notation.
fn parse_omml_children(reader: &mut Reader<&[u8]>, out: &mut String, end_tag: &[u8]) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    // Elements that produce multi-char identifiers at the start
                    // need a separator to prevent concatenation (e.g., n+cos → "ncos")
                    b"f" => {
                        ensure_math_separator(out);
                        parse_fraction(reader, out);
                    }
                    b"rad" => {
                        ensure_math_separator(out);
                        parse_radical(reader, out);
                    }
                    b"nary" => {
                        ensure_math_separator(out);
                        parse_nary(reader, out);
                    }
                    b"func" => {
                        ensure_math_separator(out);
                        parse_function(reader, out);
                    }
                    b"acc" => {
                        ensure_math_separator(out);
                        parse_accent(reader, out);
                    }
                    b"bar" => {
                        ensure_math_separator(out);
                        parse_bar(reader, out);
                    }
                    b"m" => {
                        ensure_math_separator(out);
                        parse_matrix(reader, out);
                    }
                    b"eqArr" => {
                        ensure_math_separator(out);
                        parse_eq_array(reader, out);
                    }
                    b"limLow" => {
                        ensure_math_separator(out);
                        parse_lim_low(reader, out);
                    }
                    b"limUpp" => {
                        ensure_math_separator(out);
                        parse_lim_upp(reader, out);
                    }
                    // Script elements can concatenate base chars with preceding output
                    b"sSup" => {
                        ensure_math_separator(out);
                        parse_superscript(reader, out);
                    }
                    b"sSub" => {
                        ensure_math_separator(out);
                        parse_subscript(reader, out);
                    }
                    b"sSubSup" => {
                        ensure_math_separator(out);
                        parse_sub_superscript(reader, out);
                    }
                    b"groupChr" => {
                        ensure_math_separator(out);
                        parse_group_chr(reader, out);
                    }
                    b"d" => parse_delimiter(reader, out),
                    b"r" => parse_math_run(reader, out),
                    b"oMath" => parse_omml_children(reader, out, b"oMath"),
                    b"oMathPara" => parse_omml_children(reader, out, b"oMathPara"),
                    _ => xml_util::skip_element(reader, name),
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            Err(_) => return,
            _ => {}
        }
    }
}

/// Add a space separator if `out` ends with an alphanumeric character,
/// preventing identifier concatenation across OMML elements.
fn ensure_math_separator(out: &mut String) {
    if out.chars().last().is_some_and(|c| c.is_alphanumeric()) {
        out.push(' ');
    }
}

fn parse_sub_element(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> String {
    let mut out = String::new();
    parse_omml_children(reader, &mut out, end_tag);
    out.trim().to_string()
}

fn parse_fraction(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut num = String::new();
    let mut den = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"num" => num = parse_sub_element(reader, b"num"),
                b"den" => den = parse_sub_element(reader, b"den"),
                b"fPr" => xml_util::skip_element(reader, b"fPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"f" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let _ = std::fmt::Write::write_fmt(out, format_args!("frac({num}, {den})"));
}

fn parse_superscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sup = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => base = parse_sub_element(reader, b"e"),
                b"sup" => sup = parse_sub_element(reader, b"sup"),
                b"sSupPr" => xml_util::skip_element(reader, b"sSupPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSup" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Empty base needs a placeholder to avoid bare `^` in Typst math
    if base.is_empty() {
        out.push_str("\"\"");
    } else {
        out.push_str(&base);
    }
    let _ = std::fmt::Write::write_fmt(out, format_args!("^{}", wrap_if_needed(&sup)));
}

fn parse_subscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sub = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => base = parse_sub_element(reader, b"e"),
                b"sub" => sub = parse_sub_element(reader, b"sub"),
                b"sSubPr" => xml_util::skip_element(reader, b"sSubPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSub" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Empty base needs a placeholder to avoid bare `_` in Typst math
    if base.is_empty() {
        out.push_str("\"\"");
    } else {
        out.push_str(&base);
    }
    let _ = std::fmt::Write::write_fmt(out, format_args!("_{}", wrap_if_needed(&sub)));
}

fn parse_sub_superscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sub = String::new();
    let mut sup = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => base = parse_sub_element(reader, b"e"),
                b"sub" => sub = parse_sub_element(reader, b"sub"),
                b"sup" => sup = parse_sub_element(reader, b"sup"),
                b"sSubSupPr" => xml_util::skip_element(reader, b"sSubSupPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSubSup" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Empty base needs a placeholder to avoid bare `_` in Typst math
    if base.is_empty() {
        out.push_str("\"\"");
    } else {
        out.push_str(&base);
    }
    let _ = std::fmt::Write::write_fmt(
        out,
        format_args!("_{}^{}", wrap_if_needed(&sub), wrap_if_needed(&sup)),
    );
}

fn parse_radical(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut deg = String::new();
    let mut content = String::new();
    let mut deg_hide = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"radPr" => deg_hide = parse_radical_props(reader),
                b"deg" => deg = parse_sub_element(reader, b"deg"),
                b"e" => content = parse_sub_element(reader, b"e"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rad" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Use placeholder for empty radicand to avoid Typst "missing argument" error
    let radicand = if content.is_empty() { "\"\"" } else { &content };
    if deg_hide || deg.is_empty() {
        let _ = std::fmt::Write::write_fmt(out, format_args!("sqrt({radicand})"));
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("root({deg}, {radicand})"));
    }
}

fn parse_radical_props(reader: &mut Reader<&[u8]>) -> bool {
    let mut deg_hide = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"degHide" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            deg_hide = v == "1" || v == "true" || v == "on";
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"radPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    deg_hide
}

fn parse_delimiter(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut beg_chr = "(".to_string();
    let mut end_chr = ")".to_string();
    let mut elements: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"dPr" => parse_delimiter_props(reader, &mut beg_chr, &mut end_chr),
                b"e" => elements.push(parse_sub_element(reader, b"e")),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"d" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let beg = map_delimiter(&beg_chr);
    let end = map_delimiter(&end_chr);
    let content = elements.join(", ");
    // If either delimiter is empty, omit both to avoid unbalanced delimiters in Typst
    if beg.is_empty() && end.is_empty() {
        out.push_str(&content);
    } else if beg.is_empty() || end.is_empty() {
        // One-sided invisible delimiter: emit content without delimiters
        out.push_str(&content);
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("{beg}{content}{end}"));
    }
}

fn map_delimiter(chr: &str) -> &str {
    match chr {
        "(" | ")" | "[" | "]" | "{" | "}" | "|" => chr,
        "\u{2016}" | "||" => "\u{2016}",
        "\u{27E8}" | "<" => "\u{27E8}",
        "\u{27E9}" | ">" => "\u{27E9}",
        // Floor/ceiling delimiters — explicit arms for regression safety
        "\u{2308}" | "\u{2309}" | "\u{230A}" | "\u{230B}" => chr,
        _ => chr,
    }
}

fn parse_delimiter_props(reader: &mut Reader<&[u8]>, beg: &mut String, end: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"begChr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *beg = v.to_string();
                        }
                    }
                }
                b"endChr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *end = v.to_string();
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn parse_math_run(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut in_text = false;
    let mut text_buf = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"t" => in_text = true,
                b"rPr" => xml_util::skip_element(reader, b"rPr"),
                _ => {}
            },
            Ok(Event::Text(ref t)) if in_text => {
                if let Ok(s) = t.xml_content() {
                    text_buf.push_str(s.as_ref());
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"t" => in_text = false,
                b"r" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if !text_buf.is_empty() {
        let mapped = map_math_text(&text_buf);
        // Prevent concatenation with preceding content
        if !mapped.is_empty()
            && out.chars().last().is_some_and(|c| c.is_alphanumeric())
            && mapped.chars().next().is_some_and(|c| c.is_alphanumeric())
        {
            out.push(' ');
        }
        out.push_str(&mapped);
        // Add trailing space after multi-char identifiers to prevent
        // cross-element concatenation (e.g., "plus.minus" + "sqrt" → "plus.minussqrt").
        // Single letters don't need this since they're separate tokens in Typst math.
        let chars: Vec<char> = mapped.chars().collect();
        let len = chars.len();
        if len >= 2 && chars[len - 1].is_alphabetic() && chars[len - 2].is_alphabetic() {
            out.push(' ');
        }
    }
}

/// Map Unicode characters in math text to Typst math identifiers.
///
/// Converts Greek letters and common math symbols to their Typst equivalents.
/// Splits consecutive ASCII letters into individual variables (Typst treats
/// multi-letter sequences as single identifiers), but preserves known math
/// function names like `cos`, `sin`, `log`.
fn map_math_text(input: &str) -> String {
    let mut result = String::new();
    let mut word_buf = String::new();
    let mut last_was_name = false;

    let mut non_ascii_buf = String::new();

    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            // Flush non-ASCII buffer first
            if !non_ascii_buf.is_empty() {
                flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
                non_ascii_buf.clear();
            }
            word_buf.push(ch);
            continue;
        }

        // Flush accumulated word before processing this character
        if !word_buf.is_empty() {
            // Flush non-ASCII buffer first
            if !non_ascii_buf.is_empty() {
                flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
                non_ascii_buf.clear();
            }
            flush_math_word(&mut result, &word_buf, &mut last_was_name);
            word_buf.clear();
        }

        if let Some(name) = unicode_to_typst(ch) {
            if !non_ascii_buf.is_empty() {
                flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
                non_ascii_buf.clear();
            }
            if !result.is_empty()
                && (last_was_name || result.chars().last().is_some_and(|c| c.is_alphanumeric()))
            {
                result.push(' ');
            }
            result.push_str(name);
            last_was_name = true;
        } else if ch.is_ascii_digit() {
            if !non_ascii_buf.is_empty() {
                flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
                non_ascii_buf.clear();
            }
            if last_was_name {
                result.push(' ');
            }
            result.push(ch);
            last_was_name = false;
        } else if !ch.is_ascii() && ch.is_alphabetic() {
            // Non-ASCII alphabetic (Cyrillic, CJK, etc.) — accumulate for upright() wrapping
            non_ascii_buf.push(ch);
        } else {
            if !non_ascii_buf.is_empty() {
                flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
                non_ascii_buf.clear();
            }
            // Parentheses from <m:t> are literal characters, not Typst math
            // grouping. Quote them to prevent breaking function call syntax
            // (e.g., `sqrt()` when radicand contains `)` from OMML text).
            if ch == '(' || ch == ')' {
                result.push('"');
                result.push(ch);
                result.push('"');
            } else {
                result.push(ch);
            }
            last_was_name = false;
        }
    }

    // Flush remaining buffers
    if !word_buf.is_empty() {
        flush_math_word(&mut result, &word_buf, &mut last_was_name);
    }
    if !non_ascii_buf.is_empty() {
        flush_non_ascii_text(&mut result, &non_ascii_buf, &mut last_was_name);
    }

    result
}

/// Flush accumulated non-ASCII alphabetic text as `upright("text")` for Typst math mode.
fn flush_non_ascii_text(result: &mut String, text: &str, last_was_name: &mut bool) {
    if !result.is_empty()
        && (*last_was_name || result.chars().last().is_some_and(|c| c.is_alphanumeric()))
    {
        result.push(' ');
    }
    result.push_str("upright(\"");
    result.push_str(text);
    result.push_str("\")");
    *last_was_name = true;
}

/// Flush an accumulated word of ASCII letters to the result.
///
/// Known math function names (cos, sin, etc.) are kept intact.
/// Unknown multi-letter sequences are split into individual characters
/// with spaces to prevent Typst from treating them as single identifiers.
fn flush_math_word(result: &mut String, word: &str, last_was_name: &mut bool) {
    if is_known_math_name(word) {
        // Known math name — emit as a single identifier
        if !result.is_empty()
            && (*last_was_name || result.chars().last().is_some_and(|c| c.is_alphanumeric()))
        {
            result.push(' ');
        }
        result.push_str(word);
        *last_was_name = true;
    } else if word.len() == 1 {
        // Single letter — emit as-is
        if *last_was_name {
            result.push(' ');
        }
        result.push_str(word);
        *last_was_name = false;
    } else {
        // Multiple unknown letters — split into individual characters
        for (i, c) in word.chars().enumerate() {
            if i > 0 || *last_was_name {
                result.push(' ');
            }
            result.push(c);
        }
        *last_was_name = false;
    }
}

/// Check if a word is a known math function name that should not be split.
fn is_known_math_name(text: &str) -> bool {
    matches!(
        text,
        "sin"
            | "cos"
            | "tan"
            | "cot"
            | "sec"
            | "csc"
            | "arcsin"
            | "arccos"
            | "arctan"
            | "sinh"
            | "cosh"
            | "tanh"
            | "coth"
            | "ln"
            | "log"
            | "lg"
            | "exp"
            | "det"
            | "dim"
            | "gcd"
            | "lcm"
            | "max"
            | "min"
            | "sup"
            | "inf"
            | "lim"
            | "arg"
            | "deg"
            | "mod"
    )
}

/// Map a single Unicode character to its Typst math identifier, if applicable.
fn unicode_to_typst(ch: char) -> Option<&'static str> {
    match ch {
        // Greek lowercase
        'α' => Some("alpha"),
        'β' => Some("beta"),
        'γ' => Some("gamma"),
        'δ' => Some("delta"),
        'ε' => Some("epsilon"),
        'ζ' => Some("zeta"),
        'η' => Some("eta"),
        'θ' => Some("theta"),
        'ι' => Some("iota"),
        'κ' => Some("kappa"),
        'λ' => Some("lambda"),
        'μ' => Some("mu"),
        'ν' => Some("nu"),
        'ξ' => Some("xi"),
        'ο' => Some("omicron"),
        'π' => Some("pi"),
        'ρ' => Some("rho"),
        'σ' | 'ς' => Some("sigma"),
        'τ' => Some("tau"),
        'υ' => Some("upsilon"),
        'φ' => Some("phi"),
        'χ' => Some("chi"),
        'ψ' => Some("psi"),
        'ω' => Some("omega"),
        // Greek uppercase
        'Α' => Some("Alpha"),
        'Β' => Some("Beta"),
        'Γ' => Some("Gamma"),
        'Δ' => Some("Delta"),
        'Ε' => Some("Epsilon"),
        'Ζ' => Some("Zeta"),
        'Η' => Some("Eta"),
        'Θ' => Some("Theta"),
        'Ι' => Some("Iota"),
        'Κ' => Some("Kappa"),
        'Λ' => Some("Lambda"),
        'Μ' => Some("Mu"),
        'Ν' => Some("Nu"),
        'Ξ' => Some("Xi"),
        'Ο' => Some("Omicron"),
        'Π' => Some("Pi"),
        'Ρ' => Some("Rho"),
        'Σ' => Some("Sigma"),
        'Τ' => Some("Tau"),
        'Υ' => Some("Upsilon"),
        'Φ' => Some("Phi"),
        'Χ' => Some("Chi"),
        'Ψ' => Some("Psi"),
        'Ω' => Some("Omega"),
        // Math symbols
        '∞' => Some("infinity"),
        '∂' => Some("partial"),
        '∇' => Some("nabla"),
        '∅' => Some("emptyset"),
        '±' => Some("plus.minus"),
        '×' => Some("times"),
        '÷' => Some("div"),
        '≤' => Some("lt.eq"),
        '≥' => Some("gt.eq"),
        '≠' => Some("eq.not"),
        '≈' => Some("approx"),
        '∈' => Some("in"),
        '∉' => Some("in.not"),
        '⊂' => Some("subset"),
        '⊃' => Some("supset"),
        '∪' => Some("union"),
        '∩' => Some("sect"),
        // Additional operators
        '∠' => Some("angle"),
        '∧' => Some("and"),
        '∨' => Some("or"),
        '∘' => Some("compose"),
        '⋅' => Some("dot.op"),
        '∓' => Some("minus.plus"),
        '¬' => Some("not"),
        '⊕' => Some("plus.circle"),
        '⊗' => Some("times.circle"),
        '⊙' => Some("dot.circle"),
        '⊢' => Some("tack.r"),
        '⊣' => Some("tack.l"),
        '⊤' => Some("top"),
        '⊥' => Some("perp"),
        // Arrows
        '←' => Some("arrow.l"),
        '↑' => Some("arrow.t"),
        '→' => Some("arrow.r"),
        '↓' => Some("arrow.b"),
        '↔' => Some("arrow.l.r"),
        '↦' => Some("arrow.r.bar"),
        '↪' => Some("arrow.r.hook"),
        '↼' => Some("harpoon.lt"),
        '⇀' => Some("harpoon.rt"),
        '⇐' => Some("arrow.l.double"),
        '⇒' => Some("arrow.r.double"),
        '⇔' => Some("arrow.l.r.double"),
        // Extended relations
        '≡' => Some("equiv"),
        '∼' => Some("tilde.op"),
        '≅' => Some("tilde.eq"),
        '≪' => Some("lt.double"),
        '≫' => Some("gt.double"),
        '⊆' => Some("subset.eq"),
        '⊇' => Some("supset.eq"),
        '∝' => Some("prop"),
        '∴' => Some("therefore"),
        '∵' => Some("because"),
        // Blackboard bold (double-struck) letters
        'ℂ' => Some("CC"),
        'ℍ' => Some("HH"),
        'ℕ' => Some("NN"),
        'ℙ' => Some("PP"),
        'ℚ' => Some("QQ"),
        'ℝ' => Some("RR"),
        'ℤ' => Some("ZZ"),
        _ => None,
    }
}

fn parse_nary(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut chr = "\u{2211}".to_string();
    let mut sub = String::new();
    let mut sup = String::new();
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"naryPr" => parse_nary_props(reader, &mut chr),
                b"sub" => sub = parse_sub_element(reader, b"sub"),
                b"sup" => sup = parse_sub_element(reader, b"sup"),
                b"e" => content = parse_sub_element(reader, b"e"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nary" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let op = map_nary_operator(&chr);
    out.push_str(op);
    if !sub.is_empty() {
        let _ = std::fmt::Write::write_fmt(out, format_args!("_{}", wrap_if_needed(&sub)));
    }
    if !sup.is_empty() {
        let _ = std::fmt::Write::write_fmt(out, format_args!("^{}", wrap_if_needed(&sup)));
    }
    out.push(' ');
    out.push_str(&content);
}

fn parse_nary_props(reader: &mut Reader<&[u8]>, chr: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chr" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *chr = v.to_string();
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"naryPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn map_nary_operator(chr: &str) -> &str {
    match chr {
        "\u{2211}" => "sum",
        "\u{220F}" => "product",
        "\u{222B}" => "integral",
        "\u{222C}" => "integral.double",
        "\u{222D}" => "integral.triple",
        "\u{222E}" => "integral.cont",
        "\u{22C3}" => "union.big",
        "\u{22C2}" => "sect.big",
        "\u{2210}" => "product.co",
        "\u{22C0}" => "and.big",
        "\u{22C1}" => "or.big",
        _ => "sum",
    }
}

fn parse_function(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut name = String::new();
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"fName" => name = parse_sub_element(reader, b"fName"),
                b"e" => content = parse_sub_element(reader, b"e"),
                b"funcPr" => xml_util::skip_element(reader, b"funcPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"func" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let func_name = name.trim();
    let _ = std::fmt::Write::write_fmt(out, format_args!("{func_name} {content}"));
}

fn parse_lim_low(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut lim = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => base = parse_sub_element(reader, b"e"),
                b"lim" => lim = parse_sub_element(reader, b"lim"),
                b"limLowPr" => xml_util::skip_element(reader, b"limLowPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"limLow" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let _ = std::fmt::Write::write_fmt(out, format_args!("_{}", wrap_if_needed(&lim)));
}

fn parse_lim_upp(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut lim = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => base = parse_sub_element(reader, b"e"),
                b"lim" => lim = parse_sub_element(reader, b"lim"),
                b"limUppPr" => xml_util::skip_element(reader, b"limUppPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"limUpp" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let _ = std::fmt::Write::write_fmt(out, format_args!("^{}", wrap_if_needed(&lim)));
}

fn parse_accent(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut chr = "\u{0302}".to_string();
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"accPr" => parse_accent_props(reader, &mut chr),
                b"e" => content = parse_sub_element(reader, b"e"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"acc" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let accent = map_accent(&chr);
    let _ = std::fmt::Write::write_fmt(out, format_args!("{accent}({content})"));
}

fn parse_accent_props(reader: &mut Reader<&[u8]>, chr: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chr" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *chr = v.to_string();
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"accPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn map_accent(chr: &str) -> &str {
    match chr {
        "\u{0302}" | "^" => "hat",
        "\u{0303}" | "~" => "tilde",
        "\u{0304}" | "\u{00AF}" => "macron",
        "\u{0307}" | "\u{02D9}" => "dot",
        "\u{0308}" | "\u{00A8}" => "dot.double",
        "\u{20D7}" | "\u{2192}" => "arrow",
        "\u{0301}" => "acute",
        "\u{0300}" => "grave",
        "\u{0305}" => "macron",
        "\u{030A}" => "circle",
        "\u{030C}" => "caron",
        "\u{0306}" => "breve",
        _ => "hat",
    }
}

fn parse_bar(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut content = String::new();
    let mut pos = "top".to_string();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"barPr" => parse_bar_props(reader, &mut pos),
                b"e" => content = parse_sub_element(reader, b"e"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bar" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if pos == "bot" {
        let _ = std::fmt::Write::write_fmt(out, format_args!("underline({content})"));
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("overline({content})"));
    }
}

fn parse_bar_props(reader: &mut Reader<&[u8]>, pos: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"pos" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *pos = v.to_string();
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"barPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn parse_group_chr(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut chr = "\u{23DF}".to_string(); // default: underbrace ⏟
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"groupChrPr" => parse_group_chr_props(reader, &mut chr),
                b"e" => content = parse_sub_element(reader, b"e"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"groupChr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let func = match chr.as_str() {
        "\u{23DE}" => "overbrace",  // ⏞
        "\u{23DF}" => "underbrace", // ⏟
        _ => "underbrace",
    };
    let _ = std::fmt::Write::write_fmt(out, format_args!("{func}({content})"));
}

fn parse_group_chr_props(reader: &mut Reader<&[u8]>, chr: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chr" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(v) = attr.unescape_value()
                        {
                            *chr = v.to_string();
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"groupChrPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn parse_matrix(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut rows: Vec<Vec<String>> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"mr" => rows.push(parse_matrix_row(reader)),
                b"mPr" => xml_util::skip_element(reader, b"mPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"m" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str("mat(");
    for (i, row) in rows.iter().enumerate() {
        if i > 0 {
            out.push_str("; ");
        }
        out.push_str(&row.join(", "));
    }
    out.push(')');
}

fn parse_matrix_row(reader: &mut Reader<&[u8]>) -> Vec<String> {
    let mut elements = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"e" {
                    elements.push(parse_sub_element(reader, b"e"));
                } else {
                    xml_util::skip_element(reader, e.local_name().as_ref());
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"mr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    elements
}

fn parse_eq_array(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut equations: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"e" => equations.push(parse_sub_element(reader, b"e")),
                b"eqArrPr" => xml_util::skip_element(reader, b"eqArrPr"),
                other => xml_util::skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"eqArr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    for (i, eq) in equations.iter().enumerate() {
        if i > 0 {
            out.push_str(" \\ ");
        }
        out.push_str(eq);
    }
}

fn wrap_if_needed(s: &str) -> String {
    let trimmed = s.trim();
    if trimmed.chars().count() <= 1 {
        trimmed.to_string()
    } else {
        format!("({trimmed})")
    }
}

/// Scan `word/document.xml` for math equations.
///
/// Returns `(body_child_index, typst_math, is_display)` tuples.
pub(crate) fn scan_math_equations(xml: &str) -> Vec<(usize, String, bool)> {
    let mut results = Vec::new();
    let mut reader = Reader::from_str(xml);

    let mut in_body = false;
    let mut body_child_index: usize = 0;
    let mut depth_in_body: u32 = 0;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();

                if name == b"body" {
                    in_body = true;
                    depth_in_body = 0;
                    body_child_index = 0;
                    continue;
                }

                if in_body {
                    depth_in_body += 1;

                    if name == b"oMathPara" {
                        let inner = capture_element_inner(&mut reader, b"oMathPara");
                        let typst = omml_to_typst(&inner);
                        if !typst.is_empty() {
                            results.push((body_child_index, typst, true));
                        }
                        // capture_element_inner consumed the End event, adjust depth
                        depth_in_body -= 1;
                    } else if name == b"oMath" {
                        let inner = capture_element_inner(&mut reader, b"oMath");
                        let typst = omml_to_typst(&inner);
                        if !typst.is_empty() {
                            results.push((body_child_index, typst, false));
                        }
                        // capture_element_inner consumed the End event, adjust depth
                        depth_in_body -= 1;
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"body" {
                    in_body = false;
                } else if in_body && depth_in_body > 0 {
                    depth_in_body -= 1;
                    if depth_in_body == 0 {
                        body_child_index += 1;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    results
}

fn capture_element_inner(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> String {
    let mut depth = 1u32;
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth += 1;
                }
                content.push('<');
                content.push_str(&String::from_utf8_lossy(e.name().as_ref()));
                for attr in e.attributes().flatten() {
                    content.push(' ');
                    content.push_str(&String::from_utf8_lossy(attr.key.as_ref()));
                    content.push_str("=\"");
                    if let Ok(val) = attr.unescape_value() {
                        content.push_str(&val);
                    }
                    content.push('"');
                }
                content.push('>');
            }
            Ok(Event::Empty(ref e)) => {
                content.push('<');
                content.push_str(&String::from_utf8_lossy(e.name().as_ref()));
                for attr in e.attributes().flatten() {
                    content.push(' ');
                    content.push_str(&String::from_utf8_lossy(attr.key.as_ref()));
                    content.push_str("=\"");
                    if let Ok(val) = attr.unescape_value() {
                        content.push_str(&val);
                    }
                    content.push('"');
                }
                content.push_str("/>");
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth -= 1;
                    if depth == 0 {
                        return content;
                    }
                }
                content.push_str("</");
                content.push_str(&String::from_utf8_lossy(e.name().as_ref()));
                content.push('>');
            }
            Ok(Event::Text(ref t)) => {
                if let Ok(text) = t.xml_content() {
                    content.push_str(text.as_ref());
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    content
}

#[cfg(test)]
#[path = "omml_tests.rs"]
mod tests;
