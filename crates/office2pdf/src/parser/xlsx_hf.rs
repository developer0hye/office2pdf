use crate::ir::{
    Alignment, HFInline, HeaderFooter, HeaderFooterParagraph, ParagraphStyle, Run, TextStyle,
};

/// Parse an Excel header/footer format string into IR HeaderFooter.
///
/// Excel format strings use `&L`, `&C`, `&R` to define left/center/right sections,
/// `&P` for current page number, and `&N` for total page count.
/// Returns `None` if the format string is empty.
pub(super) fn parse_hf_format_string(format_str: &str) -> Option<HeaderFooter> {
    let s = format_str.trim();
    if s.is_empty() {
        return None;
    }

    // Split into left/center/right sections
    let mut left = String::new();
    let mut center = String::new();
    let mut right = String::new();
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
                    current.push('\x01'); // Sentinel for page number
                    i += 2;
                }
                'N' => {
                    current.push('\x02'); // Sentinel for total pages
                    i += 2;
                }
                '&' => {
                    // Escaped ampersand: && → &
                    current.push('&');
                    i += 2;
                }
                '"' => {
                    // Font name: &"FontName" — skip to closing quote
                    i += 2; // skip &"
                    while i < chars.len() && chars[i] != '"' {
                        i += 1;
                    }
                    if i < chars.len() {
                        i += 1; // skip closing "
                    }
                }
                c if c.is_ascii_digit() => {
                    // Font size: &NN — skip digits
                    i += 1; // skip &
                    while i < chars.len() && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                }
                _ => {
                    // Unknown code — skip it
                    i += 2;
                }
            }
        } else {
            current.push(chars[i]);
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
        if text.is_empty() {
            continue;
        }
        let elements = build_hf_elements(text);
        if !elements.is_empty() {
            paragraphs.push(HeaderFooterParagraph {
                style: ParagraphStyle {
                    alignment: Some(*alignment),
                    ..ParagraphStyle::default()
                },
                elements,
            });
        }
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter { paragraphs })
    }
}

/// Build HFInline elements from a section string, replacing sentinel chars.
pub(super) fn build_hf_elements(section: &str) -> Vec<HFInline> {
    let mut elements = Vec::new();
    let mut current_text = String::new();

    for ch in section.chars() {
        match ch {
            '\x01' => {
                // Page number sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::PageNumber);
            }
            '\x02' => {
                // Total pages sentinel
                if !current_text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: std::mem::take(&mut current_text),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }));
                }
                elements.push(HFInline::TotalPages);
            }
            _ => {
                current_text.push(ch);
            }
        }
    }

    if !current_text.is_empty() {
        elements.push(HFInline::Run(Run {
            text: current_text,
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }));
    }

    elements
}
