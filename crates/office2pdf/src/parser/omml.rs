//! OMML (Office Math Markup Language) to Typst math notation converter.
//!
//! Parses OMML XML elements (m:oMath, m:oMathPara) from DOCX documents
//! and converts them to Typst math notation strings.

use quick_xml::Reader;
use quick_xml::events::Event;

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
                    b"d" => parse_delimiter(reader, out),
                    b"r" => parse_math_run(reader, out),
                    b"oMath" => parse_omml_children(reader, out, b"oMath"),
                    b"oMathPara" => parse_omml_children(reader, out, b"oMathPara"),
                    _ => skip_element(reader, name),
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
                b"fPr" => skip_element(reader, b"fPr"),
                other => skip_element(reader, other),
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
                b"sSupPr" => skip_element(reader, b"sSupPr"),
                other => skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSup" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
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
                b"sSubPr" => skip_element(reader, b"sSubPr"),
                other => skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSub" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
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
                b"sSubSupPr" => skip_element(reader, b"sSubSupPr"),
                other => skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSubSup" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
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
                other => skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rad" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if deg_hide || deg.is_empty() {
        let _ = std::fmt::Write::write_fmt(out, format_args!("sqrt({content})"));
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("root({deg}, {content})"));
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
                other => skip_element(reader, other),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"d" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let beg = map_delimiter(&beg_chr);
    let end = map_delimiter(&end_chr);
    let content = elements.join(", ");
    let _ = std::fmt::Write::write_fmt(out, format_args!("{beg}{content}{end}"));
}

fn map_delimiter(chr: &str) -> &str {
    match chr {
        "(" | ")" | "[" | "]" | "{" | "}" | "|" => chr,
        "\u{2016}" | "||" => "\u{2016}",
        "\u{27E8}" | "<" => "\u{27E8}",
        "\u{27E9}" | ">" => "\u{27E9}",
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
                b"rPr" => skip_element(reader, b"rPr"),
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

    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            word_buf.push(ch);
            continue;
        }

        // Flush accumulated word before processing this character
        if !word_buf.is_empty() {
            flush_math_word(&mut result, &word_buf, &mut last_was_name);
            word_buf.clear();
        }

        if let Some(name) = unicode_to_typst(ch) {
            if !result.is_empty()
                && (last_was_name || result.chars().last().is_some_and(|c| c.is_alphanumeric()))
            {
                result.push(' ');
            }
            result.push_str(name);
            last_was_name = true;
        } else if ch.is_ascii_digit() {
            if last_was_name {
                result.push(' ');
            }
            result.push(ch);
            last_was_name = false;
        } else {
            result.push(ch);
            last_was_name = false;
        }
    }

    // Flush remaining word
    if !word_buf.is_empty() {
        flush_math_word(&mut result, &word_buf, &mut last_was_name);
    }

    result
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
                other => skip_element(reader, other),
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
                b"funcPr" => skip_element(reader, b"funcPr"),
                other => skip_element(reader, other),
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
                b"limLowPr" => skip_element(reader, b"limLowPr"),
                other => skip_element(reader, other),
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
                b"limUppPr" => skip_element(reader, b"limUppPr"),
                other => skip_element(reader, other),
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
                other => skip_element(reader, other),
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
                other => skip_element(reader, other),
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

fn parse_matrix(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut rows: Vec<Vec<String>> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"mr" => rows.push(parse_matrix_row(reader)),
                b"mPr" => skip_element(reader, b"mPr"),
                other => skip_element(reader, other),
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
                    skip_element(reader, e.local_name().as_ref());
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
                b"eqArrPr" => skip_element(reader, b"eqArrPr"),
                other => skip_element(reader, other),
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

fn skip_element(reader: &mut Reader<&[u8]>, end_tag: &[u8]) {
    let mut depth = 1u32;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth -= 1;
                    if depth == 0 {
                        return;
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => return,
            _ => {}
        }
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
mod tests {
    use super::*;

    #[test]
    fn test_simple_fraction() {
        let xml = "<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>";
        assert_eq!(omml_to_typst(xml), "frac(a, b)");
    }

    #[test]
    fn test_superscript() {
        let xml = "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>";
        assert_eq!(omml_to_typst(xml), "x^2");
    }

    #[test]
    fn test_subscript() {
        let xml = "<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>";
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
}
