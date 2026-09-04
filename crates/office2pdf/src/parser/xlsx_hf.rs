use super::xlsx_cells::NormalFont;
use crate::error::ConvertWarning;
use crate::ir::{
    Alignment, Color, HFInline, HeaderFooter, HeaderFooterParagraph, ParagraphStyle, Run, TextStyle,
};
use crate::parser::xml_util::parse_hex_color;

/// One run of header/footer text and the face it takes.
///
/// A section's `&"Font,Style"` and `&<n>` codes apply to everything after them,
/// so a section is a sequence of these rather than one string (issue #633).
#[derive(Clone, Default)]
pub(super) struct HfSegment {
    text: String,
    family: Option<String>,
    size_pt: Option<f64>,
    color: Option<Color>,
    bold: bool,
    italic: bool,
}

impl HfSegment {
    /// The style this segment's runs carry. `bold`/`italic` stay `None` when
    /// unset so the renderer's own default survives, rather than being pinned
    /// to `false` by a section that named no style word.
    fn text_style(&self) -> TextStyle {
        TextStyle {
            font_family: self.family.clone(),
            east_asian_font_family: self.family.clone(),
            font_size: self.size_pt,
            color: self.color,
            bold: self.bold.then_some(true),
            italic: self.italic.then_some(true),
            ..TextStyle::default()
        }
    }
}

/// Append a string to the section's open segment.
fn push_str(section: &mut [HfSegment], text: &str) {
    if let Some(last) = section.last_mut() {
        last.text.push_str(text);
    }
}

/// Append a character to the section's open segment.
fn push_char(section: &mut [HfSegment], ch: char) {
    if let Some(last) = section.last_mut() {
        last.text.push(ch);
    }
}

/// Start a new segment carrying `style`, reusing the open one while it is still
/// empty so a run of codes does not leave blanks behind.
///
/// `style` is cloned from the previous segment to inherit the face and size a
/// code did not change, so its text has to be dropped or the new segment
/// repeats what came before it.
fn open_segment(section: &mut Vec<HfSegment>, mut style: HfSegment) {
    style.text.clear();
    match section.last_mut() {
        Some(last) if last.text.is_empty() => *last = style,
        _ => section.push(style),
    }
}

/// Sentinel standing for a decoded line break inside a section, alongside the
/// `\x01`/`\x02` page-number ones. A raw `\r` cannot serve: it would have to
/// survive `build_segment_elements`' catch-all and reach the renderer.
const LINE_BREAK_SENTINEL: char = '\x03';

/// Decode ECMA-376's `_xHHHH_` escapes (§22.9.2.19, `ST_Xstring`), the form a
/// spreadsheet string uses for a character it cannot carry directly.
///
/// `_x005F_` is the escape for a literal underscore, and decoding it here —
/// without rescanning what it produces — is what lets a genuine `_x000D_` in
/// the *text* survive as seven characters rather than turning into a line
/// break (issue #929). The four hex digits are case-insensitive per the schema.
///
/// Anything that is not a well-formed escape is passed through, so a lone
/// underscore costs nothing.
fn decode_xstring_escapes(text: &str) -> String {
    let chars: Vec<char> = text.chars().collect();
    let mut out: String = String::with_capacity(text.len());
    let mut index: usize = 0;
    while index < chars.len() {
        if let Some(decoded) = xstring_escape_at(&chars, index) {
            out.push(decoded);
            index += XSTRING_ESCAPE_LEN;
        } else {
            out.push(chars[index]);
            index += 1;
        }
    }
    out
}

/// `_xHHHH_` is always seven characters.
const XSTRING_ESCAPE_LEN: usize = 7;

/// The character the escape starting at `index` names, if one starts there.
fn xstring_escape_at(chars: &[char], index: usize) -> Option<char> {
    let escape: &[char] = chars.get(index..index + XSTRING_ESCAPE_LEN)?;
    if escape[0] != '_' || !escape[1].eq_ignore_ascii_case(&'x') || escape[6] != '_' {
        return None;
    }
    let mut code: u32 = 0;
    for digit in &escape[2..6] {
        code = code * 16 + digit.to_digit(16)?;
    }
    char::from_u32(code)
}

/// Parse an Excel header/footer format string into IR HeaderFooter.
///
/// Excel format strings use `&L`, `&C`, `&R` to define left/center/right sections,
/// `&P` for current page number, `&N` for total page count, and `&A` for the
/// worksheet name.
///
/// `sheet_name` resolves `&A`. It is a component of Excel's built-in "Sheet
/// name" header, so it turns up in files nobody customised.
///
/// `normal_font` is the workbook's Normal font, which every run before the
/// string's first face, size, or colour code takes — the same style an
/// unstyled cell inherits. Leaving the family unstated sent those runs to a
/// serif that nothing else on the sheet used (issue #951); leaving size and
/// colour unstated made a 10pt themed prefix print at the renderer's ambient
/// 11pt in black (issue #1210). `None` when `xl/styles.xml` could not be read,
/// which leaves the style unstated rather than inventing one.
///
/// Codes naming data this parser does not hold now warn instead of vanishing
/// (issue #690). `&F` and `&Z` want the workbook's file name and path, which
/// never reach `Parser::parse` — it takes bytes. `&D` and `&T` are Excel's
/// *print* date and time, which a deterministic converter has no defensible
/// value for. `&G` is a picture.
///
/// Returns `None` if the format string is empty.
pub(super) fn parse_hf_format_string(
    format_str: &str,
    sheet_name: &str,
    normal_font: Option<&NormalFont>,
    warnings: &mut Vec<ConvertWarning>,
) -> Option<HeaderFooter> {
    // `ST_Xstring` escapes ride the whole string, under the `&` codes, so they
    // come off first — a decoded `_x000D_` then reaches the walk below as the
    // line break Excel means by it rather than as seven characters
    // (issue #929).
    let decoded: String = decode_xstring_escapes(format_str);
    let s = decoded.trim();
    if s.is_empty() {
        return None;
    }

    // Split into left/center/right sections. Each section is a run of
    // segments rather than one string, because `&"Font,Style"` and `&<n>`
    // change the face and size partway through a section and every run after
    // the code takes the new values (issue #633).
    // Seeded with the Normal font so a run ahead of the first `&"Font"` code
    // carries it; `open_segment` clones the open segment's style, so a later
    // code overrides it from that point on and nothing else has to thread it.
    let opening_segment = || HfSegment {
        family: normal_font.map(|font| font.family.clone()),
        size_pt: normal_font.map(|font| font.size_pt),
        color: normal_font.and_then(|font| font.color),
        ..HfSegment::default()
    };
    let mut left: Vec<HfSegment> = vec![opening_segment()];
    let mut center: Vec<HfSegment> = vec![opening_segment()];
    let mut right: Vec<HfSegment> = vec![opening_segment()];
    let mut current = &mut center; // Default section is center if no &L/&C/&R prefix

    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '&' && i + 1 < chars.len() {
            match chars[i + 1] {
                'L' => {
                    current = &mut left;
                    i += 2;
                }
                'C' => {
                    current = &mut center;
                    i += 2;
                }
                'R' => {
                    current = &mut right;
                    i += 2;
                }
                'P' => {
                    push_char(current, '\x01'); // Sentinel for page number
                    i += 2;
                }
                'N' => {
                    push_char(current, '\x02'); // Sentinel for total pages
                    i += 2;
                }
                '&' => {
                    // Escaped ampersand: && → &
                    push_char(current, '&');
                    i += 2;
                }
                '"' => {
                    // `&"Calibri,Bold"` — the face, and optionally a style
                    // word, for every run after it. `-` means "keep the
                    // current face", which Excel writes as `&"-,Bold"`.
                    i += 2; // skip &"
                    let start = i;
                    while i < chars.len() && chars[i] != '"' {
                        i += 1;
                    }
                    let spec: String = chars[start..i].iter().collect();
                    if i < chars.len() {
                        i += 1; // skip closing "
                    }
                    let mut parts = spec.splitn(2, ',');
                    let family = parts.next().unwrap_or("").trim();
                    let style = parts.next().unwrap_or("").trim();
                    let mut next = current.last().cloned().unwrap_or_default();
                    if !family.is_empty() && family != "-" {
                        next.family = Some(family.to_string());
                    }
                    // A style word replaces both flags: Excel writes
                    // "Regular" to turn them off again.
                    if !style.is_empty() {
                        let lower = style.to_ascii_lowercase();
                        next.bold = lower.contains("bold");
                        next.italic = lower.contains("italic");
                    }
                    open_segment(current, next);
                }
                c if c.is_ascii_digit() => {
                    // `&12` — the point size for every run after it.
                    i += 1; // skip &
                    let start = i;
                    while i < chars.len() && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                    let digits: String = chars[start..i].iter().collect();
                    let mut next = current.last().cloned().unwrap_or_default();
                    if let Ok(size) = digits.parse::<f64>()
                        && size > 0.0
                    {
                        next.size_pt = Some(size);
                    }
                    open_segment(current, next);
                }
                'K' => {
                    // Font color: `&KRRGGBB`. Theme-form `&KTTSNN` codes are
                    // still consumed, but cannot be resolved without carrying
                    // the workbook theme into this format-string parser.
                    i += 2;
                    let start = i;
                    let mut consumed = 0;
                    while i < chars.len() && consumed < 6 && chars[i].is_ascii_alphanumeric() {
                        i += 1;
                        consumed += 1;
                    }
                    let code: String = chars[start..i].iter().collect();
                    if consumed == 6
                        && let Some(color) = parse_hex_color(&code)
                    {
                        let mut next = current.last().cloned().unwrap_or_default();
                        next.color = Some(color);
                        open_segment(current, next);
                    }
                }
                'A' => {
                    // Worksheet name. Excel's built-in "Sheet name" header is
                    // `&C&A`, so this turns up in files nobody customised; it
                    // used to fall through the catch-all and, being the whole
                    // section, took the paragraph with it (issue #690).
                    push_str(current, sheet_name);
                    i += 2;
                }
                'F' | 'Z' | 'D' | 'T' | 'G' => {
                    // Codes that name text Excel prints but this parser cannot
                    // produce. Report them rather than dropping them silently,
                    // so a missing header is traceable to its cause.
                    let described = match chars[i + 1] {
                        // `Parser::parse` takes bytes; no source path reaches it.
                        'F' => "&F (file name)",
                        'Z' => "&Z (file path)",
                        // Excel's *print* date/time. A converter that must be
                        // deterministic — byte-identical output for identical
                        // input — has no defensible value to substitute.
                        'D' => "&D (print date)",
                        'T' => "&T (print time)",
                        _ => "&G (picture)",
                    };
                    warnings.push(ConvertWarning::UnsupportedElement {
                        format: "XLSX".to_string(),
                        element: format!("header/footer field code {described}"),
                    });
                    i += 2;
                }
                _ => {
                    // The remaining codes are formatting toggles — &B bold,
                    // &I italic, &U underline, &E double underline, &S strike,
                    // &X superscript, &Y subscript, &O outline, &H shadow —
                    // plus section separators this parser does not model. They
                    // carry no text of their own, so skipping one changes how
                    // the header looks, never whether it appears. That is why
                    // they stay silent while the codes above warn.
                    i += 2;
                }
            }
        } else {
            // A decoded carriage return (or the line feed that may follow it)
            // breaks the section; the pair collapses to one break.
            match chars[i] {
                '\r' | '\n' => {
                    let already_broken = current
                        .last()
                        .is_some_and(|segment| segment.text.ends_with(LINE_BREAK_SENTINEL));
                    if !(chars[i] == '\n' && already_broken) {
                        push_char(current, LINE_BREAK_SENTINEL);
                    }
                }
                // A control character the escape named but that no run can
                // carry would otherwise reach the renderer as an unprintable
                // glyph. The two page-number sentinels are ours and cannot
                // appear here, since the codes that emit them are handled
                // above.
                other if other.is_control() => {}
                other => push_char(current, other),
            }
            i += 1;
        }
    }

    let mut paragraphs = Vec::new();

    // Build paragraph for each non-empty section
    let sections = [
        (&left, Alignment::Left),
        (&center, Alignment::Center),
        (&right, Alignment::Right),
    ];

    for (text, alignment) in &sections {
        if text.iter().all(|segment| segment.text.is_empty()) {
            continue;
        }
        // One paragraph per line of the section. A line that carries no ink is
        // skipped, exactly as an empty section is: Excel seats a footer's text
        // against the band, so materialising a blank leading line would move
        // the visible label without any measurement asking for it.
        for line in split_section_lines(text) {
            let elements = build_hf_elements(&line);
            if !elements.is_empty() {
                paragraphs.push(HeaderFooterParagraph {
                    style: ParagraphStyle {
                        alignment: Some(*alignment),
                        ..ParagraphStyle::default()
                    },
                    elements,
                    border: None,
                    border_space: None,
                    frame: None,
                });
            }
        }
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter {
            paragraphs,
            distance_from_edge: None,
            sheet_print_scale: None,
            shapes: Vec::new(),
        })
    }
}

/// Split a section's segments at every [`LINE_BREAK_SENTINEL`], keeping each
/// piece's style.
///
/// A break can fall inside a segment as easily as between two, since the codes
/// that open a segment and the break that ends a line are independent
/// (issue #929).
fn split_section_lines(section: &[HfSegment]) -> Vec<Vec<HfSegment>> {
    let mut lines: Vec<Vec<HfSegment>> = Vec::new();
    let mut current: Vec<HfSegment> = Vec::new();
    for segment in section {
        let mut pieces = segment.text.split(LINE_BREAK_SENTINEL).peekable();
        while let Some(piece) = pieces.next() {
            current.push(HfSegment {
                text: piece.to_string(),
                ..segment.clone()
            });
            if pieces.peek().is_some() {
                lines.push(std::mem::take(&mut current));
            }
        }
    }
    lines.push(current);
    lines
}

/// Build HFInline elements from a section string, replacing sentinel chars.
pub(super) fn build_hf_elements(section: &[HfSegment]) -> Vec<HFInline> {
    let mut elements = Vec::new();
    for segment in section {
        let style: TextStyle = segment.text_style();
        build_segment_elements(&mut elements, &segment.text, &style);
    }
    elements
}

/// Turn one segment's text into runs, expanding the page-number sentinels.
fn build_segment_elements(elements: &mut Vec<HFInline>, section: &str, style: &TextStyle) {
    let mut current_text = String::new();

    for ch in section.chars() {
        match ch {
            '\x01' => {
                // Page number sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::PageNumber(style.clone()));
            }
            '\x02' => {
                // Total pages sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::TotalPages(style.clone()));
            }
            _ => {
                current_text.push(ch);
            }
        }
    }

    if !current_text.is_empty() {
        elements.push(HFInline::Run(Run {
            text: current_text,
            style: style.clone(),
            href: None,
            footnote: None,
        }));
    }
}
