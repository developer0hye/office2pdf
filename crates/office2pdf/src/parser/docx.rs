use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, Color, Document, FlowPage, LineSpacing, Margins, Metadata, Page, PageSize,
    Paragraph, ParagraphStyle, Run, StyleSheet, TextStyle,
};
use crate::parser::Parser;

pub struct DocxParser;

impl Parser for DocxParser {
    fn parse(&self, data: &[u8]) -> Result<Document, ConvertError> {
        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX: {e}")))?;

        let (size, margins) = extract_page_setup(&docx.document.section_property);

        let mut content: Vec<Block> = Vec::new();
        for child in &docx.document.children {
            if let docx_rs::DocumentChild::Paragraph(para) = child {
                convert_paragraph_blocks(para, &mut content);
            }
        }

        Ok(Document {
            metadata: Metadata::default(),
            pages: vec![Page::Flow(FlowPage {
                size,
                margins,
                content,
            })],
            styles: StyleSheet::default(),
        })
    }
}

/// Extract page size and margins from DOCX section properties.
fn extract_page_setup(section_prop: &docx_rs::SectionProperty) -> (PageSize, Margins) {
    let size = extract_page_size(&section_prop.page_size);
    let margins = extract_margins(&section_prop.page_margin);
    (size, margins)
}

/// Extract page size from docx-rs PageSize (which has private fields).
/// Uses serde serialization to access the private `w` and `h` fields.
/// Values in DOCX are in twips (1/20 of a point).
fn extract_page_size(page_size: &docx_rs::PageSize) -> PageSize {
    if let Ok(json) = serde_json::to_value(page_size) {
        let w = json.get("w").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let h = json.get("h").and_then(|v| v.as_f64()).unwrap_or(0.0);
        if w > 0.0 && h > 0.0 {
            return PageSize {
                width: w / 20.0,  // twips to points
                height: h / 20.0, // twips to points
            };
        }
    }
    PageSize::default()
}

/// Extract margins from docx-rs PageMargin.
/// PageMargin fields are public i32 values in twips.
fn extract_margins(page_margin: &docx_rs::PageMargin) -> Margins {
    Margins {
        top: page_margin.top as f64 / 20.0,
        bottom: page_margin.bottom as f64 / 20.0,
        left: page_margin.left as f64 / 20.0,
        right: page_margin.right as f64 / 20.0,
    }
}

/// Convert a docx-rs Paragraph to IR blocks, handling page breaks.
/// If the paragraph has `page_break_before`, a `Block::PageBreak` is emitted first.
fn convert_paragraph_blocks(para: &docx_rs::Paragraph, out: &mut Vec<Block>) {
    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    let runs: Vec<Run> = para
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::ParagraphChild::Run(run) => {
                let text = extract_run_text(run);
                if text.is_empty() {
                    None
                } else {
                    Some(Run {
                        text,
                        style: extract_run_style(&run.run_property),
                    })
                }
            }
            _ => None,
        })
        .collect();

    out.push(Block::Paragraph(Paragraph {
        style: extract_paragraph_style(&para.property),
        runs,
    }));
}

/// Extract paragraph-level formatting from docx-rs ParagraphProperty.
fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
    let alignment = prop.alignment.as_ref().and_then(|j| match j.val.as_str() {
        "center" => Some(Alignment::Center),
        "right" | "end" => Some(Alignment::Right),
        "left" | "start" => Some(Alignment::Left),
        "both" | "justified" => Some(Alignment::Justify),
        _ => None,
    });

    let (indent_left, indent_right, indent_first_line) = extract_indent(&prop.indent);

    let (line_spacing, space_before, space_after) = extract_line_spacing(&prop.line_spacing);

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        space_before,
        space_after,
    }
}

/// Extract indentation values from docx-rs Indent.
/// All values in docx-rs are in twips; convert to points (÷20).
fn extract_indent(indent: &Option<docx_rs::Indent>) -> (Option<f64>, Option<f64>, Option<f64>) {
    let Some(indent) = indent else {
        return (None, None, None);
    };

    let left = indent.start.map(|v| v as f64 / 20.0);
    let right = indent.end.map(|v| v as f64 / 20.0);
    let first_line = indent.special_indent.map(|si| match si {
        docx_rs::SpecialIndentType::FirstLine(v) => v as f64 / 20.0,
        docx_rs::SpecialIndentType::Hanging(v) => -(v as f64 / 20.0),
    });

    (left, right, first_line)
}

/// Extract line spacing from docx-rs LineSpacing (private fields, use serde).
///
/// OOXML line spacing:
/// - Auto: `line` is in 240ths of a line (240=single, 360=1.5×, 480=double)
/// - Exact/AtLeast: `line` is in twips (÷20 → points)
/// - `before`/`after`: twips (÷20 → points)
fn extract_line_spacing(
    spacing: &Option<docx_rs::LineSpacing>,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
    let Some(spacing) = spacing else {
        return (None, None, None);
    };

    let json = match serde_json::to_value(spacing) {
        Ok(j) => j,
        Err(_) => return (None, None, None),
    };

    let space_before = json
        .get("before")
        .and_then(|v| v.as_f64())
        .map(|v| v / 20.0);
    let space_after = json.get("after").and_then(|v| v.as_f64()).map(|v| v / 20.0);

    let line_spacing = json.get("line").and_then(|line_val| {
        let line = line_val.as_f64()?;
        let rule = json.get("lineRule").and_then(|v| v.as_str());
        match rule {
            Some("exact") | Some("atLeast") => Some(LineSpacing::Exact(line / 20.0)),
            _ => {
                // Auto: 240ths of a line
                Some(LineSpacing::Proportional(line / 240.0))
            }
        }
    });

    (line_spacing, space_before, space_after)
}

/// Extract inline text style from a docx-rs RunProperty.
///
/// docx-rs types with private fields serialize directly as their inner value
/// (e.g. Bold → `true`, Sz → `24`, Color → `"FF0000"`), not as `{"val": ...}`.
/// Strike has a public `val` field and can be accessed directly.
fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    TextStyle {
        bold: extract_bool_prop(&rp.bold),
        italic: extract_bool_prop(&rp.italic),
        underline: rp.underline.as_ref().and_then(|u| {
            let json = serde_json::to_value(u).ok()?;
            let val = json.as_str()?;
            if val == "none" { None } else { Some(true) }
        }),
        strikethrough: rp.strike.as_ref().map(|s| s.val),
        font_size: rp.sz.as_ref().and_then(|sz| {
            let json = serde_json::to_value(sz).ok()?;
            let half_points = json.as_f64()?;
            Some(half_points / 2.0)
        }),
        color: rp.color.as_ref().and_then(|c| {
            let json = serde_json::to_value(c).ok()?;
            let hex = json.as_str()?;
            parse_hex_color(hex)
        }),
        font_family: rp.fonts.as_ref().and_then(|f| {
            let json = serde_json::to_value(f).ok()?;
            // Prefer ascii font name, fall back to hi_ansi, east_asia, cs
            json.get("ascii")
                .or_else(|| json.get("hi_ansi"))
                .or_else(|| json.get("east_asia"))
                .or_else(|| json.get("cs"))
                .and_then(|v| v.as_str())
                .map(String::from)
        }),
    }
}

/// Extract a boolean property (Bold, Italic) via serde. Returns None if absent.
/// docx-rs serializes Bold/Italic directly as a boolean (e.g. `true`).
fn extract_bool_prop<T: serde::Serialize>(prop: &Option<T>) -> Option<bool> {
    prop.as_ref().and_then(|p| {
        let json = serde_json::to_value(p).ok()?;
        json.as_bool()
    })
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Extract text content from a docx-rs Run.
fn extract_run_text(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(_) => text.push('\n'),
            _ => {}
        }
    }
    text
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;
    use std::io::Cursor;

    /// Helper: build a minimal DOCX as bytes using docx-rs builder.
    fn build_docx_bytes(paragraphs: Vec<docx_rs::Paragraph>) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new();
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with custom page size and margins.
    fn build_docx_bytes_with_page_setup(
        paragraphs: Vec<docx_rs::Paragraph>,
        width_twips: u32,
        height_twips: u32,
        margin_top: i32,
        margin_bottom: i32,
        margin_left: i32,
        margin_right: i32,
    ) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new()
            .page_size(width_twips, height_twips)
            .page_margin(
                docx_rs::PageMargin::new()
                    .top(margin_top)
                    .bottom(margin_bottom)
                    .left(margin_left)
                    .right(margin_right),
            );
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    // ----- Basic parsing tests -----

    #[test]
    fn test_parse_empty_docx() {
        let data = build_docx_bytes(vec![]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        // An empty DOCX should produce a document with one FlowPage and no content blocks
        assert_eq!(doc.pages.len(), 1);
        match &doc.pages[0] {
            Page::Flow(page) => {
                assert!(page.content.is_empty());
            }
            _ => panic!("Expected FlowPage"),
        }
    }

    #[test]
    fn test_parse_single_paragraph() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello, world!")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert_eq!(para.runs.len(), 1);
                assert_eq!(para.runs[0].text, "Hello, world!");
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    #[test]
    fn test_parse_multiple_paragraphs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Third paragraph")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 3);

        let texts: Vec<&str> = page
            .content
            .iter()
            .map(|b| match b {
                Block::Paragraph(p) => p.runs[0].text.as_str(),
                _ => panic!("Expected Paragraph"),
            })
            .collect();
        assert_eq!(
            texts,
            vec!["First paragraph", "Second paragraph", "Third paragraph"]
        );
    }

    #[test]
    fn test_parse_paragraph_with_multiple_runs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hello, "))
                .add_run(docx_rs::Run::new().add_text("beautiful "))
                .add_run(docx_rs::Run::new().add_text("world!")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].text, "Hello, ");
        assert_eq!(para.runs[1].text, "beautiful ");
        assert_eq!(para.runs[2].text, "world!");
    }

    #[test]
    fn test_parse_empty_paragraph() {
        let data = build_docx_bytes(vec![docx_rs::Paragraph::new()]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // An empty paragraph should still be present (it may have no runs)
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert!(para.runs.is_empty());
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    // ----- Page setup tests -----

    #[test]
    fn test_default_page_size_is_used() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // docx-rs default: 11906 x 16838 twips (A4)
        // = 595.3 x 841.9 pt
        assert!(page.size.width > 0.0);
        assert!(page.size.height > 0.0);
    }

    #[test]
    fn test_custom_page_size_extracted() {
        // A5 page: 148mm x 210mm
        // In twips: 8392 x 11907 (1 pt = 20 twips)
        let width_twips: u32 = 8392;
        let height_twips: u32 = 11907;
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            width_twips,
            height_twips,
            1440,
            1440,
            1440,
            1440,
        );
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_width = width_twips as f64 / 20.0;
        let expected_height = height_twips as f64 / 20.0;
        assert!(
            (page.size.width - expected_width).abs() < 1.0,
            "Expected width ~{expected_width}, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - expected_height).abs() < 1.0,
            "Expected height ~{expected_height}, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_custom_margins_extracted() {
        // Margins: 0.5 inch = 720 twips = 36pt
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            12240,
            15840,
            720,
            720,
            720,
            720,
        );
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_margin = 720.0 / 20.0; // 36pt
        assert!(
            (page.margins.top - expected_margin).abs() < 1.0,
            "Expected top margin ~{expected_margin}, got {}",
            page.margins.top
        );
        assert!((page.margins.bottom - expected_margin).abs() < 1.0);
        assert!((page.margins.left - expected_margin).abs() < 1.0);
        assert!((page.margins.right - expected_margin).abs() < 1.0);
    }

    // ----- Error handling tests -----

    #[test]
    fn test_parse_invalid_data_returns_error() {
        let parser = DocxParser;
        let result = parser.parse(b"not a valid docx file");
        assert!(result.is_err());
        match result.unwrap_err() {
            ConvertError::Parse(_) => {}
            other => panic!("Expected Parse error, got: {other:?}"),
        }
    }

    // ----- Text style defaults -----

    #[test]
    fn test_parsed_runs_have_default_text_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        let run = &para.runs[0];
        // Plain text should have default style (all None)
        assert!(run.style.bold.is_none() || run.style.bold == Some(false));
        assert!(run.style.italic.is_none() || run.style.italic == Some(false));
        assert!(run.style.underline.is_none() || run.style.underline == Some(false));
    }

    #[test]
    fn test_parsed_paragraphs_have_default_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        // Default paragraph style should have no explicit alignment
        assert!(para.style.alignment.is_none());
    }

    // ----- Inline formatting tests (US-004) -----

    /// Helper: extract the first run from the first paragraph of a parsed document.
    fn first_run(doc: &Document) -> &Run {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        &para.runs[0]
    }

    #[test]
    fn test_bold_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bold text").bold()),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_italic_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Italic text").italic()),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.italic, Some(true));
    }

    #[test]
    fn test_underline_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Underlined text")
                    .underline("single"),
            ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.underline, Some(true));
    }

    #[test]
    fn test_strikethrough_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Struck text").strike()),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.strikethrough, Some(true));
    }

    #[test]
    fn test_font_size_extracted() {
        // docx-rs size is in half-points: 24 half-points = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Sized text").size(24)),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_size, Some(12.0));
    }

    #[test]
    fn test_font_color_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Red text").color("FF0000")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_font_family_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Arial text")
                    .fonts(docx_rs::RunFonts::new().ascii("Arial")),
            ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_family, Some("Arial".to_string()));
    }

    #[test]
    fn test_combined_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Styled text")
                    .bold()
                    .italic()
                    .underline("single")
                    .strike()
                    .size(28) // 14pt
                    .color("0000FF")
                    .fonts(docx_rs::RunFonts::new().ascii("Courier")),
            ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.italic, Some(true));
        assert_eq!(run.style.underline, Some(true));
        assert_eq!(run.style.strikethrough, Some(true));
        assert_eq!(run.style.font_size, Some(14.0));
        assert_eq!(run.style.color, Some(Color::new(0, 0, 255)));
        assert_eq!(run.style.font_family, Some("Courier".to_string()));
    }

    #[test]
    fn test_plain_text_has_no_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let run = first_run(&doc);
        assert!(run.style.bold.is_none());
        assert!(run.style.italic.is_none());
        assert!(run.style.underline.is_none());
        assert!(run.style.strikethrough.is_none());
        assert!(run.style.font_size.is_none());
        assert!(run.style.color.is_none());
        assert!(run.style.font_family.is_none());
    }

    // ----- Paragraph formatting tests (US-005) -----

    /// Helper: extract the first paragraph from a parsed document.
    fn first_paragraph(doc: &Document) -> &Paragraph {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph block"),
        }
    }

    /// Helper: get all blocks from the first page.
    fn all_blocks(doc: &Document) -> &[Block] {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        &page.content
    }

    #[test]
    fn test_paragraph_alignment_center() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Centered"))
                .align(docx_rs::AlignmentType::Center),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_paragraph_alignment_right() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Right"))
                .align(docx_rs::AlignmentType::Right),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Right));
    }

    #[test]
    fn test_paragraph_alignment_left() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Left"))
                .align(docx_rs::AlignmentType::Left),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Left));
    }

    #[test]
    fn test_paragraph_alignment_justify() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Justified"))
                .align(docx_rs::AlignmentType::Both),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Justify));
    }

    #[test]
    fn test_paragraph_indent_left() {
        // 720 twips = 36pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(Some(720), None, None, None),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
    }

    #[test]
    fn test_paragraph_indent_right() {
        // 360 twips = 18pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(None, None, Some(360), None),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_right, Some(18.0));
    }

    #[test]
    fn test_paragraph_indent_first_line() {
        // first line indent: 480 twips = 24pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First line indented"))
                .indent(
                    None,
                    Some(docx_rs::SpecialIndentType::FirstLine(480)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_first_line, Some(24.0));
    }

    #[test]
    fn test_paragraph_indent_hanging() {
        // hanging indent: 360 twips = 18pt (negative first-line indent)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hanging indent"))
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::Hanging(360)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(-18.0));
    }

    #[test]
    fn test_paragraph_line_spacing_auto() {
        // Auto line spacing: line=480 means 480/240 = 2.0 (double spacing)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Double spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(480),
                ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 2.0).abs() < 0.01,
                    "Expected 2.0 (double spacing), got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_line_spacing_exact() {
        // Exact line spacing: line=240 twips = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Exact spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Exact)
                        .line(240),
                ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Exact(pts)) => {
                assert!((pts - 12.0).abs() < 0.01, "Expected 12pt, got {pts}");
            }
            other => panic!("Expected Exact line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_space_before_after() {
        // before=240 twips = 12pt, after=120 twips = 6pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Spaced paragraph"))
                .line_spacing(docx_rs::LineSpacing::new().before(240).after(120)),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.space_before, Some(12.0));
        assert_eq!(para.style.space_after, Some(6.0));
    }

    #[test]
    fn test_paragraph_page_break_before() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before break")),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("After break"))
                .page_break_before(true),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let blocks = all_blocks(&doc);
        // Should have: Paragraph("Before break"), PageBreak, Paragraph("After break")
        assert_eq!(blocks.len(), 3, "Expected 3 blocks, got {}", blocks.len());
        assert!(matches!(&blocks[0], Block::Paragraph(_)));
        assert!(matches!(&blocks[1], Block::PageBreak));
        assert!(matches!(&blocks[2], Block::Paragraph(_)));
    }

    #[test]
    fn test_paragraph_combined_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Styled paragraph"))
                .align(docx_rs::AlignmentType::Center)
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::FirstLine(360)),
                    None,
                    None,
                )
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(360)
                        .before(120)
                        .after(60),
                ),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(18.0));
        assert_eq!(para.style.space_before, Some(6.0));
        assert_eq!(para.style.space_after, Some(3.0));
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 1.5).abs() < 0.01,
                    "Expected 1.5 spacing, got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_multiple_runs_with_different_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bold ").bold())
                .add_run(docx_rs::Run::new().add_text("Italic ").italic())
                .add_run(docx_rs::Run::new().add_text("Plain")),
        ]);
        let parser = DocxParser;
        let doc = parser.parse(&data).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert!(para.runs[0].style.italic.is_none());
        assert!(para.runs[1].style.bold.is_none());
        assert_eq!(para.runs[1].style.italic, Some(true));
        assert!(para.runs[2].style.bold.is_none());
        assert!(para.runs[2].style.italic.is_none());
    }
}
