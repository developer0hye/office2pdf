//! Math text processing for OMML-to-Typst conversion.
//!
//! Handles mapping of Unicode characters, splitting multi-letter sequences,
//! and recognizing known math function names.

use super::unicode::unicode_to_typst;

/// Map Unicode characters in math text to Typst math identifiers.
///
/// Converts Greek letters and common math symbols to their Typst equivalents.
/// Splits consecutive ASCII letters into individual variables (Typst treats
/// multi-letter sequences as single identifiers), but preserves known math
/// function names like `cos`, `sin`, `log`.
pub(crate) fn map_math_text(input: &str) -> String {
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
