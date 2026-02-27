//! OMML (Office Math Markup Language) to Typst math notation converter.
//!
//! Parses OMML XML elements (m:oMath, m:oMathPara) from DOCX documents
//! and converts them to Typst math notation strings.

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

/// A parsed math equation from the document.
pub(crate) struct MathEquationInfo {
    /// Typst math notation content (without `$` delimiters).
    pub(crate) content: String,
    /// Whether this is display math (centered, block-level).
    pub(crate) display: bool,
}

/// Convert an OMML `<m:oMath>` XML fragment to Typst math notation.
///
/// The input should be the complete `<m:oMath>...</m:oMath>` element.
/// Returns the Typst math notation string (without `$` delimiters).
#[cfg(test)]
fn omml_to_typst(xml: &str) -> String {
    let mut reader = Reader::from_str(xml);
    let mut result = String::new();
    parse_omml_children(&mut reader, &mut result, b"oMath");
    result.trim().to_string()
}

/// Recursively parse OMML children and append Typst math notation.
/// `end_tag` is the local name of the element whose End event should stop parsing.
fn parse_omml_children(reader: &mut Reader<&[u8]>, out: &mut String, end_tag: &[u8]) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"f" => parse_fraction(reader, out),
                    b"sSup" => parse_superscript(reader, out),
                    b"sSub" => parse_subscript(reader, out),
                    b"sSubSup" => parse_sub_superscript(reader, out),
                    b"rad" => parse_radical(reader, out),
                    b"d" => parse_delimiter(reader, out),
                    b"r" => parse_math_run(reader, out),
                    b"nary" => parse_nary(reader, out),
                    b"func" => parse_function(reader, out),
                    b"limLow" => parse_lim_low(reader, out),
                    b"limUpp" => parse_lim_upp(reader, out),
                    b"acc" => parse_accent(reader, out),
                    b"bar" => parse_bar(reader, out),
                    // Note: b"m" would conflict with the module itself,
                    // matrix element local name is just "m" but with namespace prefix
                    b"eqArr" => parse_eq_array(reader, out),
                    b"oMath" => {
                        // Nested oMath — just process its children
                        parse_omml_children(reader, out, b"oMath");
                    }
                    b"oMathPara" => {
                        // Display math wrapper — process children
                        parse_omml_children(reader, out, b"oMathPara");
                    }
                    _ => {
                        // Skip unknown elements
                        skip_element(reader, name);
                    }
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

/// Parse a single sub-element into a string, returning the content.
fn parse_sub_element(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> String {
    let mut out = String::new();
    parse_omml_children(reader, &mut out, end_tag);
    out.trim().to_string()
}

/// Parse `<m:f>` fraction element: `<m:num>...</m:num><m:den>...</m:den>`.
fn parse_fraction(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut num = String::new();
    let mut den = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"num" => num = parse_sub_element(reader, b"num"),
                    b"den" => den = parse_sub_element(reader, b"den"),
                    b"fPr" => skip_element(reader, b"fPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"f" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let _ = std::fmt::Write::write_fmt(out, format_args!("frac({num}, {den})"));
}

/// Parse `<m:sSup>` superscript: `<m:e>base</m:e><m:sup>exp</m:sup>`.
fn parse_superscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sup = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => base = parse_sub_element(reader, b"e"),
                    b"sup" => sup = parse_sub_element(reader, b"sup"),
                    b"sSupPr" => skip_element(reader, b"sSupPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSup" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let wrapped = wrap_if_needed(&sup);
    let _ = std::fmt::Write::write_fmt(out, format_args!("^{wrapped}"));
}

/// Parse `<m:sSub>` subscript: `<m:e>base</m:e><m:sub>sub</m:sub>`.
fn parse_subscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sub = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => base = parse_sub_element(reader, b"e"),
                    b"sub" => sub = parse_sub_element(reader, b"sub"),
                    b"sSubPr" => skip_element(reader, b"sSubPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSub" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let wrapped = wrap_if_needed(&sub);
    let _ = std::fmt::Write::write_fmt(out, format_args!("_{wrapped}"));
}

/// Parse `<m:sSubSup>`: `<m:e>base</m:e><m:sub>sub</m:sub><m:sup>sup</m:sup>`.
fn parse_sub_superscript(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut sub = String::new();
    let mut sup = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => base = parse_sub_element(reader, b"e"),
                    b"sub" => sub = parse_sub_element(reader, b"sub"),
                    b"sup" => sup = parse_sub_element(reader, b"sup"),
                    b"sSubSupPr" => skip_element(reader, b"sSubSupPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sSubSup" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let w_sub = wrap_if_needed(&sub);
    let w_sup = wrap_if_needed(&sup);
    let _ = std::fmt::Write::write_fmt(out, format_args!("_{w_sub}^{w_sup}"));
}

/// Parse `<m:rad>` radical: `<m:deg>degree</m:deg><m:e>content</m:e>`.
fn parse_radical(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut deg = String::new();
    let mut content = String::new();
    let mut deg_hide = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"radPr" => deg_hide = parse_rad_props(reader),
                    b"deg" => deg = parse_sub_element(reader, b"deg"),
                    b"e" => content = parse_sub_element(reader, b"e"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rad" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    if deg_hide || deg.is_empty() {
        let _ = std::fmt::Write::write_fmt(out, format_args!("sqrt({content})"));
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("root({deg}, {content})"));
    }
}

/// Parse radical properties to check if degree is hidden (square root).
fn parse_rad_props(reader: &mut Reader<&[u8]>) -> bool {
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
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    deg_hide
}

/// Parse `<m:d>` delimiter: `<m:dPr>...</m:dPr><m:e>content</m:e>`.
fn parse_delimiter(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut beg_chr = "(".to_string();
    let mut end_chr = ")".to_string();
    let mut elements: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"dPr" => parse_delimiter_props(reader, &mut beg_chr, &mut end_chr),
                    b"e" => elements.push(parse_sub_element(reader, b"e")),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"d" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let beg = map_delimiter(&beg_chr);
    let end = map_delimiter(&end_chr);
    let content = elements.join(", ");
    let _ = std::fmt::Write::write_fmt(out, format_args!("{beg}{content}{end}"));
}

/// Map OMML delimiter characters to Typst math equivalents.
fn map_delimiter(chr: &str) -> &str {
    match chr {
        "(" | ")" | "[" | "]" | "|" => chr,
        "{" => "{",
        "}" => "}",
        _ => chr,
    }
}

/// Parse delimiter properties for begin/end characters.
fn parse_delimiter_props(reader: &mut Reader<&[u8]>, beg: &mut String, end: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"begChr" => {
                        if let Some(v) = get_val_attr(e) {
                            *beg = v;
                        }
                    }
                    b"endChr" => {
                        if let Some(v) = get_val_attr(e) {
                            *end = v;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dPr" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
}

/// Extract the `m:val` or `val` attribute value from an element.
fn get_val_attr(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val"
            && let Ok(v) = attr.unescape_value()
        {
            return Some(v.to_string());
        }
    }
    None
}

/// Parse `<m:r>` math run: extract text from `<m:t>`.
fn parse_math_run(reader: &mut Reader<&[u8]>, out: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"t" => {
                        // Read text content
                        if let Ok(Event::Text(ref text)) = reader.read_event()
                            && let Ok(s) = text.xml_content()
                        {
                            out.push_str(&s);
                        }
                    }
                    b"rPr" => skip_element(reader, b"rPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"r" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
}

/// Parse `<m:nary>` n-ary operator (sum, product, integral, etc.).
fn parse_nary(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut chr = "\u{2211}".to_string(); // default is sum ∑
    let mut sub = String::new();
    let mut sup = String::new();
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"naryPr" => parse_nary_props(reader, &mut chr),
                    b"sub" => sub = parse_sub_element(reader, b"sub"),
                    b"sup" => sup = parse_sub_element(reader, b"sup"),
                    b"e" => content = parse_sub_element(reader, b"e"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nary" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let op = map_nary_operator(&chr);
    out.push_str(op);
    if !sub.is_empty() {
        let w = wrap_if_needed(&sub);
        let _ = std::fmt::Write::write_fmt(out, format_args!("_{w}"));
    }
    if !sup.is_empty() {
        let w = wrap_if_needed(&sup);
        let _ = std::fmt::Write::write_fmt(out, format_args!("^{w}"));
    }
    out.push(' ');
    out.push_str(&content);
}

/// Parse n-ary properties to extract the operator character.
fn parse_nary_props(reader: &mut Reader<&[u8]>, chr: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chr"
                    && let Some(v) = get_val_attr(e)
                {
                    *chr = v;
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"naryPr" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
}

/// Map OMML n-ary operator characters to Typst math operators.
fn map_nary_operator(chr: &str) -> &str {
    match chr {
        "\u{2211}" => "sum",             // ∑
        "\u{220f}" => "product",         // ∏
        "\u{222b}" => "integral",        // ∫
        "\u{222c}" => "integral.double", // ∬
        "\u{222d}" => "integral.triple", // ∭
        "\u{222e}" => "integral.cont",   // ∮
        "\u{22c3}" => "union.big",       // ⋃
        "\u{22c2}" => "sect.big",        // ⋂
        _ => "sum",
    }
}

/// Parse `<m:func>` function application: `<m:fName>name</m:fName><m:e>arg</m:e>`.
fn parse_function(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut name = String::new();
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"fName" => name = parse_sub_element(reader, b"fName"),
                    b"e" => content = parse_sub_element(reader, b"e"),
                    b"funcPr" => skip_element(reader, b"funcPr"),
                    _ => skip_element(reader, local.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"func" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let func_name = name.trim();
    let _ = std::fmt::Write::write_fmt(out, format_args!("{func_name} {content}"));
}

/// Parse `<m:limLow>` lower limit: `<m:e>base</m:e><m:lim>limit</m:lim>`.
fn parse_lim_low(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut lim = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => base = parse_sub_element(reader, b"e"),
                    b"lim" => lim = parse_sub_element(reader, b"lim"),
                    b"limLowPr" => skip_element(reader, b"limLowPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"limLow" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let w = wrap_if_needed(&lim);
    let _ = std::fmt::Write::write_fmt(out, format_args!("_{w}"));
}

/// Parse `<m:limUpp>` upper limit: `<m:e>base</m:e><m:lim>limit</m:lim>`.
fn parse_lim_upp(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut base = String::new();
    let mut lim = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => base = parse_sub_element(reader, b"e"),
                    b"lim" => lim = parse_sub_element(reader, b"lim"),
                    b"limUppPr" => skip_element(reader, b"limUppPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"limUpp" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    out.push_str(&base);
    let w = wrap_if_needed(&lim);
    let _ = std::fmt::Write::write_fmt(out, format_args!("^{w}"));
}

/// Parse `<m:acc>` accent: `<m:accPr>...</m:accPr><m:e>content</m:e>`.
fn parse_accent(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut chr = "\u{0302}".to_string(); // combining circumflex accent (hat)
    let mut content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"accPr" => parse_props_with_chr(reader, b"accPr", &mut chr),
                    b"e" => content = parse_sub_element(reader, b"e"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"acc" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let accent = map_accent(&chr);
    let _ = std::fmt::Write::write_fmt(out, format_args!("{accent}({content})"));
}

/// Parse a properties element looking for a `chr` child with a `val` attribute.
fn parse_props_with_chr(reader: &mut Reader<&[u8]>, end_tag: &[u8], chr: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chr"
                    && let Some(v) = get_val_attr(e)
                {
                    *chr = v;
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
}

/// Map OMML accent characters to Typst accent functions.
fn map_accent(chr: &str) -> &str {
    match chr {
        "\u{0302}" | "^" => "hat",
        "\u{0303}" | "~" => "tilde",
        "\u{0304}" | "\u{00af}" => "macron",
        "\u{0307}" | "\u{02d9}" => "dot",
        "\u{0308}" | "\u{00a8}" => "dot.double",
        "\u{20d7}" | "\u{2192}" => "arrow",
        "\u{030c}" => "caron",
        "\u{0306}" => "breve",
        _ => "hat",
    }
}

/// Parse `<m:bar>` bar/overline: `<m:barPr>...</m:barPr><m:e>content</m:e>`.
fn parse_bar(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut content = String::new();
    let mut pos = "top".to_string();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"barPr" => parse_bar_props(reader, &mut pos),
                    b"e" => content = parse_sub_element(reader, b"e"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bar" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    if pos == "bot" {
        let _ = std::fmt::Write::write_fmt(out, format_args!("underline({content})"));
    } else {
        let _ = std::fmt::Write::write_fmt(out, format_args!("overline({content})"));
    }
}

/// Parse bar properties (position: top or bot).
fn parse_bar_props(reader: &mut Reader<&[u8]>, pos: &mut String) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"pos"
                    && let Some(v) = get_val_attr(e)
                {
                    *pos = v;
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"barPr" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
}

/// Parse `<m:eqArr>` equation array (system of equations).
fn parse_eq_array(reader: &mut Reader<&[u8]>, out: &mut String) {
    let mut equations: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"e" => equations.push(parse_sub_element(reader, b"e")),
                    b"eqArrPr" => skip_element(reader, b"eqArrPr"),
                    _ => skip_element(reader, name.as_ref()),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"eqArr" => break,
            Ok(Event::Eof) => break,
            Err(_) => break,
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

/// Wrap a string in parentheses if it contains multiple characters (for sub/superscripts).
fn wrap_if_needed(s: &str) -> String {
    let trimmed = s.trim();
    if trimmed.chars().count() <= 1 {
        trimmed.to_string()
    } else {
        format!("({trimmed})")
    }
}

/// Skip an element and all its children until its matching end tag.
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
            Ok(Event::Eof) => return,
            Err(_) => return,
            _ => {}
        }
    }
}

/// Scan a DOCX `word/document.xml` for math equations.
///
/// Returns a map from body child index (0-based) to list of math equations.
/// Each top-level `<w:p>`, `<w:tbl>`, or `<w:sdt>` in `<w:body>` is a "body child".
pub(crate) fn scan_math_equations(xml: &str) -> HashMap<usize, Vec<MathEquationInfo>> {
    let mut results: HashMap<usize, Vec<MathEquationInfo>> = HashMap::new();
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
                        // Display math — parse inline
                        let mut math_content = String::new();
                        parse_omml_children(&mut reader, &mut math_content, b"oMathPara");
                        let content = math_content.trim().to_string();
                        if !content.is_empty() {
                            results
                                .entry(body_child_index)
                                .or_default()
                                .push(MathEquationInfo {
                                    content,
                                    display: true,
                                });
                        }
                        // The End event for oMathPara was consumed by parse_omml_children
                        depth_in_body -= 1;
                    } else if name == b"oMath" {
                        // Inline math — parse inline
                        let mut math_content = String::new();
                        parse_omml_children(&mut reader, &mut math_content, b"oMath");
                        let content = math_content.trim().to_string();
                        if !content.is_empty() {
                            results
                                .entry(body_child_index)
                                .or_default()
                                .push(MathEquationInfo {
                                    content,
                                    display: false,
                                });
                        }
                        depth_in_body -= 1;
                    }
                }
            }
            Ok(Event::Empty(_)) => {
                // Self-closing elements at body level (e.g. bookmarkStart/End)
                if in_body && depth_in_body == 0 {
                    body_child_index += 1;
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

#[cfg(test)]
mod tests {
    use super::*;

    // ===== OMML to Typst conversion tests =====

    #[test]
    fn test_simple_fraction() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:num><m:r><m:t>a</m:t></m:r></m:num>
                <m:den><m:r><m:t>b</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "frac(a, b)");
    }

    #[test]
    fn test_superscript() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSup>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSup>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "x^2");
    }

    #[test]
    fn test_subscript() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSub>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
            </m:sSub>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "x_1");
    }

    #[test]
    fn test_sub_superscript() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSubSup>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSubSup>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "x_i^2");
    }

    #[test]
    fn test_square_root() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr><m:degHide m:val="1"/></m:radPr>
                <m:deg/>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "sqrt(x)");
    }

    #[test]
    fn test_nth_root() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr><m:degHide m:val="0"/></m:radPr>
                <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "root(3, x)");
    }

    #[test]
    fn test_parentheses() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:dPr>
                    <m:begChr m:val="("/>
                    <m:endChr m:val=")"/>
                </m:dPr>
                <m:e><m:r><m:t>x+y</m:t></m:r></m:e>
            </m:d>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "(x+y)");
    }

    #[test]
    fn test_complex_equation_fraction_with_superscript() {
        // a^2 / (b + c)
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:num>
                    <m:sSup>
                        <m:e><m:r><m:t>a</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                    </m:sSup>
                </m:num>
                <m:den>
                    <m:d>
                        <m:e>
                            <m:r><m:t>b</m:t></m:r>
                            <m:r><m:t>+</m:t></m:r>
                            <m:r><m:t>c</m:t></m:r>
                        </m:e>
                    </m:d>
                </m:den>
            </m:f>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "frac(a^2, (b+c))");
    }

    #[test]
    fn test_sum_with_limits() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:nary>
                <m:naryPr><m:chr m:val="∑"/></m:naryPr>
                <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
                <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
                <m:e><m:r><m:t>i</m:t></m:r></m:e>
            </m:nary>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "sum_(i=1)^n i");
    }

    #[test]
    fn test_plain_text_e_equals_mc2() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r><m:t>E</m:t></m:r>
            <m:r><m:t>=</m:t></m:r>
            <m:r><m:t>m</m:t></m:r>
            <m:sSup>
                <m:e><m:r><m:t>c</m:t></m:r></m:e>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSup>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "E=mc^2");
    }

    #[test]
    fn test_quadratic_formula() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r><m:t>x</m:t></m:r>
            <m:r><m:t>=</m:t></m:r>
            <m:f>
                <m:num>
                    <m:r><m:t>-b±</m:t></m:r>
                    <m:rad>
                        <m:radPr><m:degHide m:val="1"/></m:radPr>
                        <m:deg/>
                        <m:e>
                            <m:sSup>
                                <m:e><m:r><m:t>b</m:t></m:r></m:e>
                                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                            </m:sSup>
                            <m:r><m:t>-4ac</m:t></m:r>
                        </m:e>
                    </m:rad>
                </m:num>
                <m:den>
                    <m:r><m:t>2a</m:t></m:r>
                </m:den>
            </m:f>
        </m:oMath>"#;
        let result = omml_to_typst(xml);
        assert_eq!(result, "x=frac(-b\u{00b1}sqrt(b^2-4ac), 2a)");
    }

    // ===== scan_math_equations tests =====

    #[test]
    fn test_scan_display_math() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <w:body>
                <w:p>
                    <w:r><w:t>Before math</w:t></w:r>
                </w:p>
                <w:p>
                    <m:oMathPara>
                        <m:oMath>
                            <m:f>
                                <m:num><m:r><m:t>a</m:t></m:r></m:num>
                                <m:den><m:r><m:t>b</m:t></m:r></m:den>
                            </m:f>
                        </m:oMath>
                    </m:oMathPara>
                </w:p>
                <w:p>
                    <w:r><w:t>After math</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"#;

        let math = scan_math_equations(xml);
        assert!(
            math.contains_key(&1),
            "Math should be in body child index 1"
        );
        let eqs = &math[&1];
        assert_eq!(eqs.len(), 1);
        assert!(eqs[0].display);
        assert_eq!(eqs[0].content, "frac(a, b)");
    }

    #[test]
    fn test_scan_inline_math() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <w:body>
                <w:p>
                    <w:r><w:t>The equation </w:t></w:r>
                    <m:oMath>
                        <m:sSup>
                            <m:e><m:r><m:t>x</m:t></m:r></m:e>
                            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                        </m:sSup>
                    </m:oMath>
                    <w:r><w:t> is nice</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"#;

        let math = scan_math_equations(xml);
        assert!(math.contains_key(&0));
        let eqs = &math[&0];
        assert_eq!(eqs.len(), 1);
        assert!(!eqs[0].display);
        assert_eq!(eqs[0].content, "x^2");
    }
}
