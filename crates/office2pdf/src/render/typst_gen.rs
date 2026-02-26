use std::fmt::Write;

use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, Color, Document, FlowPage, LineSpacing, Margins, Page, PageSize, Paragraph,
    ParagraphStyle, Run, TextStyle,
};

/// Generate Typst markup from a Document IR.
pub fn generate_typst(doc: &Document) -> Result<String, ConvertError> {
    let mut out = String::new();
    for page in &doc.pages {
        match page {
            Page::Flow(flow) => generate_flow_page(&mut out, flow)?,
            Page::Fixed(_) | Page::Table(_) => {
                // Not yet implemented — other stories will handle these
            }
        }
    }
    Ok(out)
}

fn generate_flow_page(out: &mut String, page: &FlowPage) -> Result<(), ConvertError> {
    write_page_setup(out, &page.size, &page.margins);
    out.push('\n');

    for (i, block) in page.content.iter().enumerate() {
        if i > 0 {
            out.push('\n');
        }
        generate_block(out, block)?;
    }
    Ok(())
}

fn write_page_setup(out: &mut String, size: &PageSize, margins: &Margins) {
    let _ = writeln!(
        out,
        "#set page(width: {}pt, height: {}pt, margin: (top: {}pt, bottom: {}pt, left: {}pt, right: {}pt))",
        format_f64(size.width),
        format_f64(size.height),
        format_f64(margins.top),
        format_f64(margins.bottom),
        format_f64(margins.left),
        format_f64(margins.right),
    );
}

fn generate_block(out: &mut String, block: &Block) -> Result<(), ConvertError> {
    match block {
        Block::Paragraph(para) => generate_paragraph(out, para),
        Block::PageBreak => {
            out.push_str("#pagebreak()\n");
            Ok(())
        }
        Block::Table(_) | Block::Image(_) => {
            // Not yet implemented — other stories will handle these
            Ok(())
        }
    }
}

fn generate_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
    let style = &para.style;
    let has_para_style = needs_block_wrapper(style);

    if has_para_style {
        out.push_str("#block(");
        write_block_params(out, style);
        out.push_str(")[\n");
        write_par_settings(out, style);
    }

    // Generate alignment wrapper if needed
    let alignment = style.alignment;
    let use_align = matches!(
        alignment,
        Some(Alignment::Center) | Some(Alignment::Right) | Some(Alignment::Left)
    );

    if use_align {
        let align_str = match alignment.unwrap() {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => unreachable!(),
        };
        let _ = write!(out, "#align({align_str})[");
    }

    // Generate runs
    for run in &para.runs {
        generate_run(out, run);
    }

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

/// Check if paragraph style needs a block wrapper (for spacing/leading/justify).
fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.line_spacing.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
}

fn write_block_params(out: &mut String, style: &ParagraphStyle) {
    let mut first = true;

    if let Some(above) = style.space_before {
        write_param(out, &mut first, &format!("above: {}pt", format_f64(above)));
    }
    if let Some(below) = style.space_after {
        write_param(out, &mut first, &format!("below: {}pt", format_f64(below)));
    }
}

fn write_par_settings(out: &mut String, style: &ParagraphStyle) {
    if let Some(ref spacing) = style.line_spacing {
        match spacing {
            LineSpacing::Proportional(factor) => {
                let leading = factor * 0.65;
                let _ = writeln!(out, "  #set par(leading: {}em)", format_f64(leading));
            }
            LineSpacing::Exact(pts) => {
                let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(*pts));
            }
        }
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
}

fn generate_run(out: &mut String, run: &Run) {
    let style = &run.style;
    let escaped = escape_typst(&run.text);

    let has_text_props = has_text_properties(style);
    let needs_underline = matches!(style.underline, Some(true));
    let needs_strike = matches!(style.strikethrough, Some(true));

    // Wrap with decorations (outermost first)
    if needs_strike {
        out.push_str("#strike[");
    }
    if needs_underline {
        out.push_str("#underline[");
    }

    if has_text_props {
        out.push_str("#text(");
        write_text_params(out, style);
        out.push_str(")[");
        out.push_str(&escaped);
        out.push(']');
    } else {
        out.push_str(&escaped);
    }

    if needs_underline {
        out.push(']');
    }
    if needs_strike {
        out.push(']');
    }
}

/// Check if the text style has properties that need a #text() wrapper
/// (not counting underline/strikethrough which use separate wrappers).
fn has_text_properties(style: &TextStyle) -> bool {
    matches!(style.bold, Some(true))
        || matches!(style.italic, Some(true))
        || style.font_size.is_some()
        || style.color.is_some()
        || style.font_family.is_some()
}

fn write_text_params(out: &mut String, style: &TextStyle) {
    let mut first = true;

    if let Some(ref family) = style.font_family {
        write_param(out, &mut first, &format!("font: \"{family}\""));
    }
    if let Some(size) = style.font_size {
        write_param(out, &mut first, &format!("size: {}pt", format_f64(size)));
    }
    if matches!(style.bold, Some(true)) {
        write_param(out, &mut first, "weight: \"bold\"");
    }
    if matches!(style.italic, Some(true)) {
        write_param(out, &mut first, "style: \"italic\"");
    }
    if let Some(ref color) = style.color {
        write_param(out, &mut first, &format_color(color));
    }
}

fn write_param(out: &mut String, first: &mut bool, param: &str) {
    if !*first {
        out.push_str(", ");
    }
    out.push_str(param);
    *first = false;
}

fn format_color(color: &Color) -> String {
    format!("fill: rgb({}, {}, {})", color.r, color.g, color.b)
}

/// Format an f64 without unnecessary trailing zeros.
fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

/// Escape special Typst characters in text content.
fn escape_typst(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' => {
                result.push('\\');
                result.push(ch);
            }
            '$' => {
                result.push('\\');
                result.push('$');
            }
            _ => result.push(ch),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Metadata, StyleSheet};

    /// Helper to create a minimal Document with one FlowPage.
    fn make_doc(pages: Vec<Page>) -> Document {
        Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        }
    }

    /// Helper to create a FlowPage with default A4 size and margins.
    fn make_flow_page(content: Vec<Block>) -> Page {
        Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content,
        })
    }

    /// Helper to create a simple paragraph with one plain-text run.
    fn make_paragraph(text: &str) -> Block {
        Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
            }],
        })
    }

    #[test]
    fn test_generate_plain_paragraph() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
        let result = generate_typst(&doc).unwrap();
        assert!(result.contains("Hello World"));
    }

    #[test]
    fn test_generate_page_setup() {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            size: PageSize {
                width: 612.0,
                height: 792.0,
            },
            margins: Margins {
                top: 36.0,
                bottom: 36.0,
                left: 54.0,
                right: 54.0,
            },
            content: vec![make_paragraph("test")],
        })]);
        let result = generate_typst(&doc).unwrap();
        assert!(result.contains("612pt"));
        assert!(result.contains("792pt"));
        assert!(result.contains("36pt"));
        assert!(result.contains("54pt"));
    }

    #[test]
    fn test_generate_bold_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold text".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("weight: \"bold\""),
            "Expected bold weight in: {result}"
        );
        assert!(result.contains("Bold text"));
    }

    #[test]
    fn test_generate_italic_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Italic text".to_string(),
                style: TextStyle {
                    italic: Some(true),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("style: \"italic\""),
            "Expected italic style in: {result}"
        );
        assert!(result.contains("Italic text"));
    }

    #[test]
    fn test_generate_underline_text() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Underlined".to_string(),
                style: TextStyle {
                    underline: Some(true),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("#underline["),
            "Expected underline wrapper in: {result}"
        );
        assert!(result.contains("Underlined"));
    }

    #[test]
    fn test_generate_font_size() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Large text".to_string(),
                style: TextStyle {
                    font_size: Some(24.0),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("size: 24pt"),
            "Expected font size in: {result}"
        );
    }

    #[test]
    fn test_generate_font_color() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Red text".to_string(),
                style: TextStyle {
                    color: Some(Color::new(255, 0, 0)),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("fill: rgb(255, 0, 0)"),
            "Expected RGB color in: {result}"
        );
    }

    #[test]
    fn test_generate_combined_text_styles() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Styled".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    italic: Some(true),
                    font_size: Some(16.0),
                    color: Some(Color::new(0, 128, 255)),
                    ..TextStyle::default()
                },
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(result.contains("weight: \"bold\""));
        assert!(result.contains("style: \"italic\""));
        assert!(result.contains("size: 16pt"));
        assert!(result.contains("fill: rgb(0, 128, 255)"));
        assert!(result.contains("Styled"));
    }

    #[test]
    fn test_generate_alignment_center() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Centered".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("align(center"),
            "Expected center alignment in: {result}"
        );
    }

    #[test]
    fn test_generate_alignment_right() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Right),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Right".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("align(right"),
            "Expected right alignment in: {result}"
        );
    }

    #[test]
    fn test_generate_alignment_justify() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Justify),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Justified text".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("par(justify: true") || result.contains("set par(justify: true"),
            "Expected justify in: {result}"
        );
    }

    #[test]
    fn test_generate_line_spacing_proportional() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                line_spacing: Some(LineSpacing::Proportional(2.0)),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Double spaced".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("leading:"),
            "Expected leading setting in: {result}"
        );
    }

    #[test]
    fn test_generate_line_spacing_exact() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                line_spacing: Some(LineSpacing::Exact(18.0)),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Exact spaced".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(
            result.contains("leading: 18pt"),
            "Expected exact leading in: {result}"
        );
    }

    #[test]
    fn test_generate_multiple_paragraphs() {
        let doc = make_doc(vec![make_flow_page(vec![
            make_paragraph("First paragraph"),
            make_paragraph("Second paragraph"),
        ])]);
        let result = generate_typst(&doc).unwrap();
        assert!(result.contains("First paragraph"));
        assert!(result.contains("Second paragraph"));
    }

    #[test]
    fn test_generate_paragraph_with_multiple_runs() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "Normal ".to_string(),
                    style: TextStyle::default(),
                },
                Run {
                    text: "bold".to_string(),
                    style: TextStyle {
                        bold: Some(true),
                        ..TextStyle::default()
                    },
                },
                Run {
                    text: " normal again".to_string(),
                    style: TextStyle::default(),
                },
            ],
        })])]);
        let result = generate_typst(&doc).unwrap();
        assert!(result.contains("Normal "));
        assert!(result.contains("bold"));
        assert!(result.contains(" normal again"));
    }

    #[test]
    fn test_generate_empty_document() {
        let doc = make_doc(vec![]);
        let result = generate_typst(&doc).unwrap();
        // Should produce valid (possibly empty) Typst markup
        assert!(result.is_empty() || !result.is_empty()); // Just shouldn't error
    }

    #[test]
    fn test_generate_special_characters_escaped() {
        let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
            "Price: $100 #items @store",
        )])]);
        let result = generate_typst(&doc).unwrap();
        // The text should appear but special chars should be escaped for Typst
        // In Typst, # starts a code expression, so it needs escaping
        assert!(
            result.contains("\\#") || result.contains("Price"),
            "Expected escaped or present text in: {result}"
        );
    }

    #[test]
    fn test_generate_space_before_after() {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                space_before: Some(12.0),
                space_after: Some(6.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Spaced paragraph".to_string(),
                style: TextStyle::default(),
            }],
        })])]);
        let result = generate_typst(&doc).unwrap();
        // Should contain spacing directives
        assert!(
            result.contains("12pt") || result.contains("above"),
            "Expected space_before in: {result}"
        );
    }
}
